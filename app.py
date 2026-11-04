import base64
import hmac
import io
import json
import os
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Hệ Thống Kiểm Kê Kho", layout="wide")

DB_FILE = "inventory.db"
GH_PATH = "inventory.db"
META_COLS = ["Người cập nhật", "Cập nhật bởi (IP)", "Thời gian sửa"]
SECTION_NAMES = ["Xuất khẩu", "HBM", "Nội địa", "Vật tư giao thẳng"]


# ───────────────────────── SECRETS ─────────────────────────
def get_secret(key):
    try:
        return str(st.secrets[key])
    except Exception:
        return None


# ───────────────────────── DATABASE ─────────────────────────
@contextmanager
def db():
    conn = sqlite3.connect(DB_FILE, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with db() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS sheets(
                name TEXT PRIMARY KEY, cols TEXT, pos INTEGER);
            CREATE TABLE IF NOT EXISTS items(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet TEXT, data TEXT, slkk REAL,
                updated_by TEXT DEFAULT '', ip TEXT DEFAULT '',
                updated_at TEXT DEFAULT '');
            CREATE INDEX IF NOT EXISTS idx_items_sheet ON items(sheet);
            CREATE TABLE IF NOT EXISTS log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT, user TEXT, ip TEXT, sheet TEXT,
                item_id INTEGER, col TEXT, old TEXT, new TEXT);
            """
        )


# ───────────────────────── GITHUB BACKUP ─────────────────────────
@st.cache_resource
def backup_lock():
    return threading.Lock()


def gh_conf():
    token, repo = get_secret("GITHUB_TOKEN"), get_secret("REPO_NAME")
    if token and repo:
        return token, repo
    return None


def restore_from_github():
    cfg = gh_conf()
    if not cfg:
        return
    token, repo = cfg
    try:
        r = requests.get(
            f"https://api.github.com/repos/{repo}/contents/{GH_PATH}",
            headers={
                "Authorization": f"token {token}",
                "Accept": "application/vnd.github.raw",
            },
            timeout=60,
        )
        if r.status_code == 200 and r.content:
            with open(DB_FILE, "wb") as f:
                f.write(r.content)
    except Exception as e:
        print(f"Restore GitHub lỗi: {e}")


def backup_to_github():
    cfg = gh_conf()
    if not cfg:
        return False, "Chưa cấu hình GITHUB_TOKEN / REPO_NAME"
    token, repo = cfg
    url = f"https://api.github.com/repos/{repo}/contents/{GH_PATH}"
    headers = {"Authorization": f"token {token}"}
    tmp = DB_FILE + ".bak"
    try:
        with backup_lock():
            src, dst = sqlite3.connect(DB_FILE), sqlite3.connect(tmp)
            src.backup(dst)
            dst.close()
            src.close()
            with open(tmp, "rb") as f:
                content = base64.b64encode(f.read()).decode()

            last = None
            for _ in range(2):
                g = requests.get(url, headers=headers, timeout=30)
                payload = {
                    "message": "Auto-backup inventory.db từ Streamlit",
                    "content": content,
                }
                if g.status_code == 200:
                    payload["sha"] = g.json().get("sha")
                p = requests.put(url, headers=headers, json=payload, timeout=60)
                last = p
                if p.status_code in (200, 201):
                    return True, "OK"
            return False, f"GitHub trả về {last.status_code}: {last.text[:150]}"
    except Exception as e:
        return False, str(e)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


@st.cache_resource
def init_app():
    if not os.path.exists(DB_FILE):
        restore_from_github()
    init_db()
    return True


# ───────────────────────── LOGIC ─────────────────────────
def get_section_name(sheet_name):
    s = sheet_name.lower()
    if "hbm" in s:
        return "HBM"
    if "nội địa" in s or "noi dia" in s or "domestic" in s:
        return "Nội địa"
    if any(k in s for k in ["direct", "giao thẳng", "vật tư", "vat tu"]):
        return "Vật tư giao thẳng"
    return "Xuất khẩu"


def get_nhom_name(sheet_name):
    """Lấy tên nhóm là phần trước dấu '-' trên sheet (ví dụ 'dầu mỡ-985Y' -> 'dầu mỡ')"""
    if "-" in sheet_name:
        return sheet_name.split("-")[0].strip()
    return sheet_name.strip()


def norm(v):
    return None if pd.isna(v) else v


def now_vn():
    return datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%Y-%m-%d %H:%M:%S")


def get_client_ip():
    try:
        ip = st.context.headers.get("X-Forwarded-For", "")
        return ip.split(",")[0].strip() or "Cloud User"
    except Exception:
        return "Không xác định"


def import_excel(file_bytes, header_idx):
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=header_idx, dtype=str)
    with db() as c:
        c.execute("DELETE FROM sheets")
        c.execute("DELETE FROM items")
        for pos, (name, df) in enumerate(raw.items()):
            df = df.dropna(how="all").dropna(axis=1, how="all")
            df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
            df.columns = [str(x).strip() for x in df.columns]
            slkk = (
                pd.to_numeric(df["SLKK"], errors="coerce")
                if "SLKK" in df.columns
                else pd.Series([None] * len(df), index=df.index, dtype="float")
            )
            df = df.drop(columns=[x for x in ["SLKK"] + META_COLS if x in df.columns])
            df = df.astype(object).where(df.notna(), None)
            c.execute(
                "INSERT INTO sheets(name, cols, pos) VALUES (?,?,?)",
                (name, json.dumps(list(df.columns), ensure_ascii=False), pos),
            )
            rows = [
                (name, json.dumps(rec, ensure_ascii=False), None if pd.isna(s) else float(s))
                for rec, s in zip(df.to_dict("records"), slkk)
            ]
            c.executemany("INSERT INTO items(sheet, data, slkk) VALUES (?,?,?)", rows)


def list_sheets():
    with db() as c:
        return [r["name"] for r in c.execute("SELECT name FROM sheets ORDER BY pos")]


def load_df(sheet):
    with db() as c:
        cols = json.loads(c.execute("SELECT cols FROM sheets WHERE name=?", (sheet,)).fetchone()["cols"])
        rows = c.execute(
            "SELECT id, data, slkk, updated_by, ip, updated_at FROM items WHERE sheet=? ORDER BY id",
            (sheet,),
        ).fetchall()
    df = pd.DataFrame([json.loads(r["data"]) for r in rows], columns=cols)
    df["SLKK"] = pd.Series([r["slkk"] for r in rows], dtype="float")
    df[META_COLS[0]] = [r["updated_by"] for r in rows]
    df[META_COLS[1]] = [r["ip"] for r in rows]
    df[META_COLS[2]] = [r["updated_at"] for r in rows]
    df.index = pd.Index([r["id"] for r in rows], name="id")
    return df, cols


def save_changes(sheet, original, edited, data_cols, user, ip, is_admin):
    ts, changed, conflicts = now_vn(), 0, 0
    with db() as c:
        for rid in edited.index:
            if rid not in original.index:
                continue
            old, new = original.loc[rid], edited.loc[rid]
            cur = c.execute("SELECT data, slkk FROM items WHERE id=?", (int(rid),)).fetchone()
            if cur is None:
                continue
            data, slkk_val, dirty = json.loads(cur["data"]), cur["slkk"], False

            new_s, old_s = norm(new["SLKK"]), norm(old["SLKK"])
            if new_s != old_s:
                if cur["slkk"] != old_s:
                    conflicts += 1
                c.execute(
                    "INSERT INTO log(ts,user,ip,sheet,item_id,col,old,new) VALUES (?,?,?,?,?,?,?,?)",
                    (ts, user, ip, sheet, int(rid), "SLKK", str(cur["slkk"]), str(new_s)),
                )
                slkk_val, dirty = (None if new_s is None else float(new_s)), True

            if is_admin:
                for col in data_cols:
                    if norm(new[col]) != norm(old[col]):
                        c.execute(
                            "INSERT INTO log(ts,user,ip,sheet,item_id,col,old,new) VALUES (?,?,?,?,?,?,?,?)",
                            (ts, user, ip, sheet, int(rid), col, str(old[col]), str(new[col])),
                        )
                        data[col] = norm(new[col])
                        dirty = True

            if dirty:
                c.execute(
                    "UPDATE items SET data=?, slkk=?, updated_by=?, ip=?, updated_at=? WHERE id=?",
                    (json.dumps(data, ensure_ascii=False), slkk_val, user, ip, ts, int(rid)),
                )
                changed += 1
    return changed, conflicts


def build_xlsx():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for s in list_sheets():
            df, _ = load_df(s)
            df.reset_index(drop=True).to_excel(w, index=False, sheet_name=s[:31])
    return out.getvalue()


def build_log_csv():
    with db() as c:
        df = pd.read_sql_query("SELECT * FROM log ORDER BY id DESC", c)
    return df.to_csv(index=False).encode("utf-8-sig")


# ───────────────────────── UI ─────────────────────────
init_app()
st.session_state.setdefault("ver", 0)

st.title("📋 HỆ THỐNG KIỂM KÊ KHO")
st.markdown("---")

with st.sidebar:
    st.header("👤 THÔNG TIN NGƯỜI DÙNG")
    user_name = st.text_input("Nhập họ tên của bạn:", placeholder="VD: Nguyễn Thị Trang").strip()

    st.markdown("---")
    st.header("🔐 VAI TRÒ")
    role = st.selectbox("Chọn vai trò:", ["Nhân viên kho (Chỉ nhập SLKK)", "Quản trị viên (Admin)"])
    is_admin = False
    if role.startswith("Quản trị"):
        expected = get_secret("ADMIN_PASSWORD")
        pw = st.text_input("Mật khẩu Admin:", type="password")
        if not expected:
            st.error("Chưa cấu hình ADMIN_PASSWORD trong secrets.")
        elif pw:
            if hmac.compare_digest(pw.encode(), expected.encode()):
                is_admin = True
            else:
                st.error("Sai mật khẩu.")

    if is_admin:
        st.markdown("---")
        st.header("⚙️ QUẢN LÝ DỮ LIỆU")
        header_line = st.number_input("Dòng chứa tiêu đề (Header row):", min_value=1, value=4, step=1)
        up = st.file_uploader("📂 Tải list kiểm kê mới (Excel):", type=["xlsx", "xls"])
        confirm = st.checkbox("Tôi hiểu: áp dụng file mới sẽ thay toàn bộ dữ liệu hiện tại")
        if st.button("✅ Áp dụng file này", disabled=not (up and confirm)):
            try:
                import_excel(up.getvalue(), header_line - 1)
                ok, msg = backup_to_github()
                st.session_state["ver"] += 1
                st.success("Đã nạp list mới." + ("" if ok else f" (Chưa sao lưu GitHub: {msg})"))
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}")

        if st.button("☁️ Sao lưu lên GitHub ngay"):
            ok, msg = backup_to_github()
            st.success("Đã sao lưu.") if ok else st.error(msg)

        st.download_button(
            "📜 Tải nhật ký chỉnh sửa (CSV)", build_log_csv(), "nhat_ky_chinh_sua.csv", "text/csv"
        )

sheets = list_sheets()
if not sheets:
    st.warning("⚠️ Chưa có dữ liệu. Quản trị viên hãy đăng nhập Admin và tải file Excel lên ở thanh bên trái.")
    st.stop()

if is_admin:
    st.success("👑 Chế độ **Quản trị viên**: được sửa tất cả các cột dữ liệu.")
else:
    st.info("💡 Nhập họ tên (thanh bên trái) → chọn Section Name → Nhóm → nhập **SLKK** → bấm **LƯU**. "
            "Hãy lưu trước khi đổi bộ lọc, vì đổi bộ lọc sẽ làm mới bảng.")

# Xây dựng cấu trúc phân cấp: Section Name -> Nhóm (tên trước dấu '-') -> Danh sách sheet chi tiết
section_map = {sec: {} for sec in SECTION_NAMES}
for s in sheets:
    sec = get_section_name(s)
    nhom = get_nhom_name(s)
    if nhom not in section_map[sec]:
        section_map[sec][nhom] = []
    section_map[sec][nhom].append(s)

available_sections = [sec for sec, nhoms in section_map.items() if nhoms]

c1, c2, c3 = st.columns(3)
with c1:
    section_name = st.selectbox("📂 Section Name:", available_sections)
with c2:
    available_nhoms = list(section_map[section_name].keys())
    nhom_choice = st.selectbox("📋 Nhóm:", available_nhoms)
with c3:
    sheet_choices = section_map[section_name][nhom_choice]
    # Nếu nhóm có nhiều sheet chi tiết (khác tên), cho chọn sheet cụ thể; nếu chỉ có 1 sheet thì tự động lấy luôn
    if len(sheet_choices) > 1:
        sheet = st.selectbox("📌 Sheet chi tiết:", sheet_choices)
    else:
        sheet = sheet_choices[0]
        st.text_input("📌 Sheet chi tiết:", value=sheet, disabled=True)

full_df, data_cols = load_df(sheet)
done = int(full_df["SLKK"].notna().sum())
st.caption(f"Đã nhập SLKK: **{done} / {len(full_df)}** dòng trong sheet này.")

view = full_df
f1, f2 = st.columns(2)
with f1:
    s_loc = st.text_input("🔍 Lọc theo Location:", key="filter_loc")
    if s_loc and "location" in view.columns:
        view = view[view["location"].astype(str).str.contains(s_loc, case=False, na=False)]
with f2:
    s_mat = st.text_input("🔍 Lọc theo Mã vật tư:", key="filter_mat")
    if s_mat and "material_no" in view.columns:
        view = view[view["material_no"].astype(str).str.contains(s_mat, case=False, na=False)]

locked = META_COLS if is_admin else [c for c in view.columns if c != "SLKK"]
original = view.copy()

with st.form(key=f"form_{sheet}"):
    edited = st.data_editor(
        view,
        disabled=locked,
        hide_index=True,
        use_container_width=True,
        height=500,
        column_config={"SLKK": st.column_config.NumberColumn("SLKK", min_value=0)},
        key=f"editor_{sheet}_{s_loc}_{s_mat}_{st.session_state['ver']}",
    )
    submitted = st.form_submit_button(f"💾 LƯU THAY ĐỔI ({sheet})", type="primary", use_container_width=True)

if submitted:
    if not user_name:
        st.warning("⚠️ Vui lòng nhập **Họ tên** ở thanh bên trái trước khi lưu!")
    else:
        changed, conflicts = save_changes(
            sheet, original, edited, data_cols, user_name, get_client_ip(), is_admin
        )
        if changed:
            ok, msg = backup_to_github()
            st.session_state["ver"] += 1
            st.success(f"🎉 Đã lưu {changed} dòng của **{user_name}**.")
            if conflicts:
                st.warning(f"⚠️ {conflicts} dòng đã được người khác sửa trong lúc bạn nhập; giá trị của bạn đã ghi đè.")
            if not ok:
                st.warning(f"Đã lưu trên server nhưng chưa sao lưu được lên GitHub: {msg}")
        else:
            st.info("ℹ️ Không có thay đổi mới nào.")

st.markdown("---")
st.subheader("💾 TẢI XUỐNG KẾT QUẢ")
if st.button("Chuẩn bị file Excel"):
    st.session_state["xlsx"] = build_xlsx()
if "xlsx" in st.session_state:
    st.download_button(
        "📥 Tải file Excel (kèm Tên, IP, Thời gian)",
        st.session_state["xlsx"],
        "KetQua_KiemKe_Audit.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
