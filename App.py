import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, date, timedelta
import os
import io

# Cấu hình trang Streamlit
st.set_page_config(page_title="Mecake - Quản lý Kho & Order", layout="wide", page_icon="🧁")

DB_FILE = "mecake_management.db"
EXCEL_FILE = "Chốt tồn và oder Mecake.xlsx"
RECOVERY_CODE = "MECAKE-ADMIN-999" # Mã khôi phục hệ thống khi quên mật khẩu

# ---------------------------------------------------------
# DATABASE SETUP
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS products (
            code TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            unit_price REAL DEFAULT 0,
            shelf_life_days INTEGER DEFAULT 7
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS inventory_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT,
            import_date DATE,
            expiry_date DATE,
            quantity INTEGER,
            FOREIGN KEY (code) REFERENCES products (code) ON DELETE CASCADE
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS order_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_date DATE,
            code TEXT,
            order_qty INTEGER,
            note TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS discarded_goods (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER,
            code TEXT,
            discard_date DATE,
            quantity INTEGER,
            unit_price REAL,
            reason TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('admin_password', '123456')")
    conn.commit()
    conn.close()

def get_admin_password():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key = 'admin_password'")
    row = c.fetchone()
    conn.close()
    return row[0] if row else "123456"

def set_admin_password(new_pass):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('admin_password', ?)", (new_pass,))
    conn.commit()
    conn.close()

def load_data_from_excel_if_empty():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM products")
    count = c.fetchone()[0]
    
    if count == 0 and os.path.exists(EXCEL_FILE):
        try:
            df_price = pd.read_excel(EXCEL_FILE, sheet_name='Tiền nhập bánh ')
            df_price = df_price.dropna(subset=['Mã Bánh'])
            
            for _, row in df_price.iterrows():
                code = str(row['Mã Bánh']).strip()
                name = str(row['Tên Bánh']).strip()
                price = float(row['Giá Sỉ (Có VAT)']) if pd.notnull(row['Giá Sỉ (Có VAT)']) else 0.0
                shelf_life = 30 if ('Pate' in name or 'gối' in name) else 7
                
                c.execute('''
                    INSERT OR REPLACE INTO products (code, name, unit_price, shelf_life_days)
                    VALUES (?, ?, ?, ?)
                ''', (code, name, price, shelf_life))
            conn.commit()
            st.success("Đã nạp tự động Danh mục bánh & Giá sỉ từ file Excel Mecake!")
        except Exception as e:
            st.error(f"Lỗi khi đọc file Excel: {e}")
    conn.close()

init_db()
load_data_from_excel_if_empty()

# ---------------------------------------------------------
# POPUP DIALOG BẢO MẬT XÓA KHO
# ---------------------------------------------------------
@st.dialog("🔐 Xác Nhận Mật Khẩu Để Xóa Dữ Liệu")
def confirm_delete_dialog(target_type):
    st.warning(f"⚠️ Bạn đang yêu cầu: **XÓA SẠCH {target_type.upper()}**.")
    st.write("Vui lòng nhập Mật Khẩu Admin để hoàn tất thao tác này:")
    
    pwd_input = st.text_input("Nhập Mật khẩu:", type="password", key="dlg_pwd_input")
    current_pass = get_admin_password()
    
    col_act1, col_act2 = st.columns(2)
    
    if col_act1.button("✅ Xác Nhận Xóa", type="primary", use_container_width=True):
        if pwd_input == current_pass:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            if target_type == "Tồn Kho":
                c.execute("DELETE FROM inventory_batches")
                st.success("🎉 Đã xóa sạch toàn bộ dữ liệu Tồn Kho!")
            elif target_type == "Kho Hàng Hủy":
                c.execute("DELETE FROM discarded_goods")
                st.success("🎉 Đã xóa sạch toàn bộ dữ liệu Kho Hàng Hủy!")
            conn.commit()
            conn.close()
            st.rerun()
        else:
            st.error("❌ Mật khẩu không chính xác!")

    with st.expander("❓ Quên Mật Khẩu?"):
        st.info("Nhập Mã Khôi Phục Hệ Thống (`MECAKE-ADMIN-999`) để reset mật khẩu về `123456`.")
        rec_code = st.text_input("Mã Khôi Phục:", key="dlg_rec_inp")
        if st.button("Reset Mật Khẩu Về 123456"):
            if rec_code.strip() == RECOVERY_CODE:
                set_admin_password("123456")
                st.success("✅ Đã đặt lại Mật Khẩu Admin về mặc định: 123456")
                st.rerun()
            else:
                st.error("Mã khôi phục không đúng!")

# ---------------------------------------------------------
# GIAO DIỆN CHÍNH (SIDEBAR MENU - ĐÃ RÚT GỌN NÚT TÍNH NĂNG 3)
# ---------------------------------------------------------
st.sidebar.title("🧁 Mecake Manager")
menu = st.sidebar.radio(
    "Chọn chức năng", 
    ["1. Soạn thảo & Nhập Order", "2. Báo cáo Tồn Kho & Hạn Sử Dụng", "3. Quản lý Danh mục Bánh"]
)

# ---------------------------------------------------------
# CHỨC NĂNG 1: SOẠN THẢO & NHẬP ORDER
# ---------------------------------------------------------
if menu == "1. Soạn thảo & Nhập Order":
    tab_create, tab_history = st.tabs(["📝 Soạn Thảo Order Mới", "📦 Lịch Sử Đơn Order Đã Chốt"])
    
    conn = sqlite3.connect(DB_FILE)
    
    # TAB 1: SOẠN THẢO ORDER
    with tab_create:
        st.header("📋 Soạn Thảo & Nhập Số Lượng Order Bánh")
        
        df_prod = pd.read_sql_query("SELECT code AS 'Mã Bánh', name AS 'Tên Bánh', unit_price AS 'Giá Sỉ (có VAT)' FROM products", conn)
        df_inv = pd.read_sql_query("SELECT code, SUM(quantity) as current_stock FROM inventory_batches WHERE quantity > 0 GROUP BY code", conn)
        df_order_view = pd.merge(df_prod, df_inv, left_on='Mã Bánh', right_on='code', how='left').fillna({'current_stock': 0})
        df_order_view['Tồn Kho'] = df_order_view['current_stock'].astype(int)
        
        col_date, col_search = st.columns([1, 2])
        order_date = col_date.date_input("Ngày Order", date.today(), key="order_date_input")
        search_keyword = col_search.text_input("🔍 Tìm kiếm theo Mã Bánh hoặc Tên Bánh:", "", key="search_order")
        
        if search_keyword:
            mask = df_order_view['Mã Bánh'].str.contains(search_keyword, case=False, na=False) | \
                   df_order_view['Tên Bánh'].str.contains(search_keyword, case=False, na=False)
            df_order_view = df_order_view[mask]
        
        df_order_view['Số Lượng Order'] = 0
        display_cols = ['Mã Bánh', 'Tên Bánh', 'Giá Sỉ (có VAT)', 'Tồn Kho', 'Số Lượng Order']
        
        edited_df = st.data_editor(
            df_order_view[display_cols],
            column_config={
                "Số Lượng Order": st.column_config.NumberColumn("Số Lượng Order", min_value=0, step=1, default=0),
                "Giá Sỉ (có VAT)": st.column_config.NumberColumn("Giá Sỉ (VNĐ)", format="%d ₫"),
                "Tồn Kho": st.column_config.NumberColumn("Tồn Kho", disabled=True)
            },
            disabled=["Mã Bánh", "Tên Bánh", "Giá Sỉ (có VAT)", "Tồn Kho"],
            hide_index=True,
            use_container_width=True,
            key="order_editor"
        )
        
        edited_df['Thành Tiền'] = edited_df['Số Lượng Order'] * edited_df['Giá Sỉ (có VAT)']
        total_qty = edited_df['Số Lượng Order'].sum()
        total_val = edited_df['Thành Tiền'].sum()
        
        c_res1, c_res2 = st.columns(2)
        c_res1.metric("Tổng Số Lượng Order", f"{total_qty} cái")
        c_res2.metric("Tổng Giá Trị Đơn Hàng", f"{total_val:,.0f} VNĐ")
        
        if st.button("💾 Lưu & Chốt Đơn Order này", type="primary"):
            c = conn.cursor()
            orders_to_save = edited_df[edited_df['Số Lượng Order'] > 0]
            if not orders_to_save.empty:
                c.execute("DELETE FROM order_records WHERE order_date = ?", (order_date,))
                for _, row in orders_to_save.iterrows():
                    c.execute('''
                        INSERT INTO order_records (order_date, code, order_qty, note)
                        VALUES (?, ?, ?, ?)
                    ''', (order_date, row['Mã Bánh'], int(row['Số Lượng Order']), 'Order hàng'))
                conn.commit()
                st.success(f"✅ Đã lưu thành công đơn Order ngày {order_date}! Chuyển sang Tab 'Lịch Sử Đơn Order Đã Chốt' để xem/xuất dữ liệu.")
            else:
                st.warning("Vui lòng nhập số lượng Order (> 0) cho ít nhất 1 loại bánh!")

    # TAB 2: LỊCH SỬ, SỬA ĐƠN ORDER TRƯỚC KHIN HẬP KHO, XUẤT FILE VÀ CHUYỂN KHO
    with tab_history:
        st.header("📦 Quản Lý, Chỉnh Sửa & Xuất Đơn Order")
        
        df_all_dates = pd.read_sql_query("SELECT DISTINCT order_date FROM order_records ORDER BY order_date DESC", conn)
        
        if df_all_dates.empty:
            st.info("Chưa có đơn order nào được chốt trong hệ thống.")
        else:
            selected_date_str = st.selectbox("Chọn ngày đã chốt Order để xem/sửa:", df_all_dates['order_date'].tolist())
            
            df_all_prod = pd.read_sql_query("SELECT code, name, unit_price, shelf_life_days FROM products", conn)
            prod_codes_list = df_all_prod['code'].tolist()
            
            query_detail = '''
                SELECT 
                    o.id AS 'ID_Record',
                    o.code AS 'Mã Bánh',
                    p.name AS 'Tên Bánh',
                    p.unit_price AS 'Giá Sỉ',
                    p.shelf_life_days AS 'HSD Chuẩn',
                    o.order_qty AS 'Số Lượng Order'
                FROM order_records o
                JOIN products p ON o.code = p.code
                WHERE o.order_date = ? AND o.order_qty > 0
            '''
            df_detail = pd.read_sql_query(query_detail, conn, params=(selected_date_str,))
            
            st.subheader(f"Danh sách Bánh Order cho ngày: {selected_date_str}")
            st.info("💡 Bạn có thể **SỬA TRỰC TIẾP Mã Bánh hoặc Số Lượng Order** trong bảng bên dưới rồi bấm **'💾 Lưu Thay Đổi Đơn Order'** trước khi Nhập Kho!")
            
            edited_order_df = st.data_editor(
                df_detail,
                column_config={
                    "ID_Record": None,
                    "Mã Bánh": st.column_config.SelectboxColumn("Mã Bánh", options=prod_codes_list, required=True),
                    "Tên Bánh": st.column_config.TextColumn("Tên Bánh", disabled=True),
                    "Giá Sỉ": st.column_config.NumberColumn("Giá Sỉ (VNĐ)", format="%d ₫", disabled=True),
                    "HSD Chuẩn": st.column_config.NumberColumn("HSD Chuẩn", disabled=True),
                    "Số Lượng Order": st.column_config.NumberColumn("Số Lượng Order", min_value=0, step=1, required=True)
                },
                hide_index=True,
                use_container_width=True,
                key="edit_history_order_editor"
            )
            
            col_save_ord, col_space = st.columns([1, 2])
            if col_save_ord.button("💾 Lưu Thay Đổi Đơn Order", type="primary"):
                c = conn.cursor()
                for _, row_ord in edited_order_df.iterrows():
                    c.execute('''
                        UPDATE order_records
                        SET code = ?, order_qty = ?
                        WHERE id = ?
                    ''', (str(row_ord['Mã Bánh']), int(row_ord['Số Lượng Order']), int(row_ord['ID_Record'])))
                conn.commit()
                st.success("✅ Đã lưu cập nhật chỉnh sửa đơn order!")
                st.rerun()

            edited_order_df['Thành Tiền'] = edited_order_df['Số Lượng Order'] * edited_order_df['Giá Sỉ']
            sum_qty = edited_order_df['Số Lượng Order'].sum()
            sum_money = edited_order_df['Thành Tiền'].sum()
            st.markdown(f"**Tổng cộng:** **{sum_qty}** cái | **Tổng tiền:** **{sum_money:,.0f} VNĐ**")
            
            st.markdown("---")
            col_dl, col_del = st.columns(2)
            
            # 1. Tải file Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                edited_order_df[['Mã Bánh', 'Tên Bánh', 'Giá Sỉ', 'Số Lượng Order', 'Thành Tiền']].to_excel(writer, index=False, sheet_name='Don_Order')
            
            col_dl.download_button(
                label="📥 Tải Excel Đơn Order",
                data=buffer.getvalue(),
                file_name=f"Don_Order_Mecake_{selected_date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            # 2. Xóa Đơn Order
            if col_del.button("🗑️ XÓA HOÀN TOÀN ĐƠN ORDER NÀY", type="secondary", use_container_width=True):
                c = conn.cursor()
                c.execute("DELETE FROM order_records WHERE order_date = ?", (selected_date_str,))
                conn.commit()
                st.success(f"🗑️ Đã xóa hoàn toàn dữ liệu đơn Order ngày {selected_date_str}!")
                st.rerun()

            st.markdown("---")
            with st.expander("🚚 CHUYỂN ĐƠN ORDER NÀY THÀNH LÔ NHẬP KHO (CHỌN NGÀY NHẬP THỰC TẾ)", expanded=True):
                col_inp_date, col_inp_btn = st.columns([1, 1])
                user_selected_import_date = col_inp_date.date_input("Chọn Ngày Thực Tế Nhập Kho:", date.today())
                
                if col_inp_btn.button("🚀 XÁC NHẬN NHẬP KHO", type="primary", use_container_width=True):
                    c = conn.cursor()
                    count_imported = 0
                    for _, row in edited_order_df.iterrows():
                        code = str(row['Mã Bánh'])
                        qty = int(row['Số Lượng Order'])
                        if qty > 0:
                            c.execute("SELECT shelf_life_days FROM products WHERE code = ?", (code,))
                            p_row = c.fetchone()
                            shelf_days = p_row[0] if p_row else 7
                            
                            expiry_date = user_selected_import_date + timedelta(days=shelf_days)
                            
                            c.execute('''
                                INSERT INTO inventory_batches (code, import_date, expiry_date, quantity)
                                VALUES (?, ?, ?, ?)
                            ''', (code, user_selected_import_date, expiry_date, qty))
                            count_imported += 1
                    conn.commit()
                    st.success(f"🎉 Đã chuyển {count_imported} mã bánh vào kho với Ngày Nhập Kho chính thức là: **{user_selected_import_date}**!")
                    st.info("💡 Bạn có thể sang mục '2. Báo cáo Tồn Kho & Hạn Sử Dụng' để xem chi tiết hạn sử dụng.")

    conn.close()

# ---------------------------------------------------------
# CHỨC NĂNG 2: BÁO CÁO TỒN KHO & CẢNH BÁO CẬN DATE / HỦY HÀNG
# ---------------------------------------------------------
elif menu == "2. Báo cáo Tồn Kho & Hạn Sử Dụng":
    tab_inv, tab_discard = st.tabs(["⏳ Tồn Kho & Cảnh Báo HSD", "🗑️ Kho Hàng Hủy & Báo Cáo Lãng Phí"])
    
    conn = sqlite3.connect(DB_FILE)
    
    with tab_inv:
        st.header("⏳ Kiểm Soát Tồn Kho & Hạn Sử Dụng Theo Thời Gian")
        
        with st.expander("⚙️ CẤU HÌNH NGƯỠNG CẢNH BÁO & KHUYẾN CÁO BÁN HÀNG", expanded=False):
            c_cfg1, c_cfg2 = st.columns([1, 2])
            warn_days = c_cfg1.number_input("Cảnh báo cận Date khi số ngày còn lại ≤ (ngày):", min_value=1, value=2, step=1)
            discount_policy = c_cfg2.text_input("Nội dung Khuyến cáo bán hàng (Chiết khấu):", "Giảm giá 30% - 50% hoặc Tặng kèm Combo")
        
        st.subheader("🔍 Lọc & Tìm Kiếm Tồn Kho:")
        col_f1, col_f2 = st.columns([2, 1])
        filter_text = col_f1.text_input("Mã Bánh / Tên Bánh:", "", placeholder="Nhập mã hoặc tên bánh...")
        
        df_dates_imp = pd.read_sql_query("SELECT DISTINCT import_date FROM inventory_batches WHERE quantity > 0 ORDER BY import_date DESC", conn)
        imp_dates_list = ["Tất cả ngày nhập"] + df_dates_imp['import_date'].tolist()
        filter_date = col_f2.selectbox("Ngày Nhập Kho:", imp_dates_list)
        
        query = '''
            SELECT 
                b.id AS 'ID Lô',
                b.code AS 'Mã Bánh',
                p.name AS 'Tên Bánh',
                p.unit_price AS 'Giá Sỉ',
                b.import_date AS 'Ngày Nhập',
                b.expiry_date AS 'Hạn Sử Dụng',
                b.quantity AS 'Số Lượng Tồn'
            FROM inventory_batches b
            JOIN products p ON b.code = p.code
            WHERE b.quantity > 0
            ORDER BY b.expiry_date ASC
        '''
        df_batches = pd.read_sql_query(query, conn)
        
        if filter_text:
            mask_txt = df_batches['Mã Bánh'].str.contains(filter_text, case=False, na=False) | \
                       df_batches['Tên Bánh'].str.contains(filter_text, case=False, na=False)
            df_batches = df_batches[mask_txt]
            
        if filter_date != "Tất cả ngày nhập":
            df_batches = df_batches[df_batches['Ngày Nhập'] == filter_date]
            
        if df_batches.empty:
            st.info("Không tìm thấy lô hàng nào phù hợp với điều kiện tìm kiếm!")
        else:
            today = date.today()
            df_batches['Hạn Sử Dụng'] = pd.to_datetime(df_batches['Hạn Sử Dụng']).dt.date
            df_batches['Số Ngày Còn Lại'] = (pd.to_datetime(df_batches['Hạn Sử Dụng']) - pd.to_datetime(today)).dt.days
            
            def set_status(days):
                if days < 0:
                    return "🔴 Đã Hết Hạn"
                elif days <= warn_days:
                    return "🟡 Cận Date"
                else:
                    return "🟢 An Toàn"
                    
            def set_recommendation(days):
                if days < 0:
                    return "🔴 Cần bấm Hủy lô bánh"
                elif days <= warn_days:
                    return f"⚠️ {discount_policy}"
                else:
                    return "Bán bình thường"

            df_batches['Trạng Thái'] = df_batches['Số Ngày Còn Lại'].apply(set_status)
            df_batches['Khuyến Cáo Bán Hàng'] = df_batches['Số Ngày Còn Lại'].apply(set_recommendation)
            
            expired_df = df_batches[df_batches['Số Ngày Còn Lại'] < 0]
            warning_df = df_batches[(df_batches['Số Ngày Còn Lại'] >= 0) & (df_batches['Số Ngày Còn Lại'] <= warn_days)]
            safe_df = df_batches[df_batches['Số Ngày Còn Lại'] > warn_days]
            
            c1, c2, c3 = st.columns(3)
            c1.metric("🟢 Lô An Toàn", f"{len(safe_df)} lô")
            c2.metric(f"🟡 Lô Cận Date (≤{warn_days} ngày)", f"{len(warning_df)} lô")
            c3.metric("🔴 Lô Hết Hạn (Cần Hủy)", f"{len(expired_df)} lô")
            
            if not warning_df.empty:
                st.warning(f"⚠️ **CẢNH BÁO BÁN HÀNG CẬN DATE:** Có {len(warning_df)} lô hàng cận HSD! 👉 **Khuyến cáo:** {discount_policy}")
            
            st.subheader("Chi tiết Tồn kho từng Lô bánh (Có thể chỉnh sửa cột 'Số Lượng Tồn' trực tiếp):")
            st.info("💡 Bạn có thể bấm đúp vào cột **'Số Lượng Tồn'** để thay đổi số lượng tồn kho của từng lô, sau đó bấm **'💾 Lưu Cập Nhật Tồn Kho'**.")
            
            edited_batches_df = st.data_editor(
                df_batches[['ID Lô', 'Mã Bánh', 'Tên Bánh', 'Ngày Nhập', 'Hạn Sử Dụng', 'Số Ngày Còn Lại', 'Số Lượng Tồn', 'Trạng Thái', 'Khuyến Cáo Bán Hàng']],
                column_config={
                    "ID Lô": st.column_config.NumberColumn("ID Lô", disabled=True),
                    "Mã Bánh": st.column_config.TextColumn("Mã Bánh", disabled=True),
                    "Tên Bánh": st.column_config.TextColumn("Tên Bánh", disabled=True),
                    "Ngày Nhập": st.column_config.DateColumn("Ngày Nhập", disabled=True),
                    "Hạn Sử Dụng": st.column_config.DateColumn("Hạn Sử Dụng", disabled=True),
                    "Số Ngày Còn Lại": st.column_config.NumberColumn("Số Ngày Còn Lại", disabled=True),
                    "Số Lượng Tồn": st.column_config.NumberColumn("Số Lượng Tồn", min_value=0, step=1, required=True),
                    "Trạng Thái": st.column_config.TextColumn("Trạng Thái", disabled=True),
                    "Khuyến Cáo Bán Hàng": st.column_config.TextColumn("Khuyến Cáo Bán Hàng", disabled=True)
                },
                hide_index=True,
                use_container_width=True,
                key="inventory_qty_editor"
            )
            
            if st.button("💾 Lưu Cập Nhật Tồn Kho", type="primary"):
                c = conn.cursor()
                for _, row_b in edited_batches_df.iterrows():
                    c.execute("UPDATE inventory_batches SET quantity = ? WHERE id = ?", (int(row_b['Số Lượng Tồn']), int(row_b['ID Lô'])))
                conn.commit()
                st.success("✅ Đã cập nhật lại số lượng tồn kho thành công!")
                st.rerun()

            if not expired_df.empty:
                st.markdown("---")
                st.error("🚨 **KHỐI THAO TÁC HỦY BÁNH HẾT HẠN (SỐ NGÀY CÒN LẠI < 0):**")
                for _, exp_row in expired_df.iterrows():
                    b_id = exp_row['ID Lô']
                    b_code = exp_row['Mã Bánh']
                    b_name = exp_row['Tên Bánh']
                    b_qty = exp_row['Số Lượng Tồn']
                    b_price = exp_row['Giá Sỉ']
                    
                    c_del1, c_del2, c_del3 = st.columns([2, 2, 1])
                    c_del1.write(f"🔴 **Lô #{b_id}** - [{b_code}] {b_name} (Tồn: {b_qty} cái)")
                    reason_val = c_del2.text_input("Lý do hủy:", "Quá hạn sử dụng", key=f"rec_reason_{b_id}")
                    if c_del3.button(f"🗑️ Hủy Lô #{b_id}", type="primary", key=f"btn_del_tab_{b_id}"):
                        c = conn.cursor()
                        c.execute("UPDATE inventory_batches SET quantity = 0 WHERE id = ?", (b_id,))
                        c.execute('''
                            INSERT INTO discarded_goods (batch_id, code, discard_date, quantity, unit_price, reason)
                            VALUES (?, ?, ?, ?, ?, ?)
                        ''', (b_id, b_code, date.today(), b_qty, b_price, reason_val))
                        conn.commit()
                        st.success(f"✅ Đã hủy {b_qty} bánh lô #{b_id} và chuyển vào Kho Hàng Hủy!")
                        st.rerun()

        st.markdown("---")
        with st.expander("🔒 QUẢN LÝ MẬT KHẨU & XÓA SẠCH DỮ LIỆU KHO (BẢO MẬT)", expanded=False):
            current_pass = get_admin_password()
            
            with st.expander("⚙️ Đổi Mật Khẩu Admin"):
                c_p1, c_p2 = st.columns(2)
                old_p = c_p1.text_input("Mật khẩu cũ:", type="password", key="old_p_inp")
                new_p = c_p2.text_input("Mật khẩu mới:", type="password", key="new_p_inp")
                if st.button("Lưu Mật Khẩu Mới"):
                    if old_p == current_pass:
                        if new_p.strip():
                            set_admin_password(new_p.strip())
                            st.success("✅ Đã đổi Mật khẩu Admin thành công!")
                        else:
                            st.warning("Mật khẩu mới không được để trống!")
                    else:
                        st.error("Mật khẩu cũ không chính xác!")

            col_clr_inv, col_clr_dis = st.columns(2)
            
            if col_clr_inv.button("🔐 XÓA SẠCH TỒN KHO (RESET KHO VỀ 0)", type="primary", use_container_width=True):
                confirm_delete_dialog("Tồn Kho")

            if col_clr_dis.button("🔐 XÓA SẠCH DỮ LIỆU KHO HÀNG HỦY", type="secondary", use_container_width=True):
                confirm_delete_dialog("Kho Hàng Hủy")

    # TAB KHO HÀNG HỦY
    with tab_discard:
        st.header("🗑️ Báo Cáo Kho Hàng Hủy & Thiệt Hại Lãng Phí")
        
        query_discard = '''
            SELECT 
                d.id AS 'Mã Bản Ghi',
                d.discard_date AS 'Ngày Hủy',
                d.code AS 'Mã Bánh',
                p.name AS 'Tên Bánh',
                d.quantity AS 'Số Lượng Hủy',
                d.unit_price AS 'Đơn Giá Sỉ',
                (d.quantity * d.unit_price) AS 'Tổng Thiệt Hại (VNĐ)',
                d.reason AS 'Lý Do Hủy'
            FROM discarded_goods d
            JOIN products p ON d.code = p.code
            ORDER BY d.discard_date DESC
        '''
        df_dis = pd.read_sql_query(query_discard, conn)
        
        if df_dis.empty:
            st.info("Chưa có ghi nhận hủy bánh nào trong hệ thống. Cửa hàng vận hành rất tốt!")
        else:
            total_dis_qty = df_dis['Số Lượng Hủy'].sum()
            total_dis_money = df_dis['Tổng Thiệt Hại (VNĐ)'].sum()
            
            cm1, cm2 = st.columns(2)
            cm1.metric("Tổng Số Lượng Bánh Hủy", f"{total_dis_qty} cái")
            cm2.metric("Tổng Giá Trị Thiệt Hại (Giá sỉ)", f"{total_dis_money:,.0f} VNĐ")
            
            st.subheader("Chi tiết Danh sách Bánh Đã Hủy:")
            st.dataframe(df_dis, use_container_width=True, hide_index=True)
            
            buffer_dis = io.BytesIO()
            with pd.ExcelWriter(buffer_dis, engine='openpyxl') as writer:
                df_dis.to_excel(writer, index=False, sheet_name='Hang_Huy')
            
            st.download_button(
                label="📥 Tải Báo Cáo Kho Hàng Hủy (Excel)",
                data=buffer_dis.getvalue(),
                file_name=f"Bao_Cao_Hang_Huy_Mecake_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    conn.close()

# ---------------------------------------------------------
# CHỨC NĂNG 3: QUẢN LÝ DANH MỤC BÁNH
# ---------------------------------------------------------
elif menu == "3. Quản lý Danh mục Bánh":
    st.header("⚙️ Danh Mục Bánh & Cấu Hình Hạn Sử Dụng Chuẩn")
    st.info("💡 Bạn có thể bấm đúp vào từng ô để **SỬA TRỰC TIẾP** Tên Bánh, Giá Sỉ hoặc HSD Chuẩn, sau đó nhấn **'💾 Lưu Thay Đổi'** ở bên dưới.")
    
    conn = sqlite3.connect(DB_FILE)
    df_p = pd.read_sql_query("SELECT code AS 'Mã Bánh', name AS 'Tên Bánh', unit_price AS 'Giá Sỉ', shelf_life_days AS 'HSD Chuẩn (Ngày)' FROM products", conn)
    
    edited_products_df = st.data_editor(
        df_p,
        column_config={
            "Mã Bánh": st.column_config.TextColumn("Mã Bánh", disabled=True),
            "Tên Bánh": st.column_config.TextColumn("Tên Bánh", required=True),
            "Giá Sỉ": st.column_config.NumberColumn("Giá Sỉ (VNĐ)", format="%d ₫", min_value=0),
            "HSD Chuẩn (Ngày)": st.column_config.NumberColumn("HSD Chuẩn (Ngày)", min_value=1, step=1)
        },
        hide_index=True,
        use_container_width=True,
        key="product_catalog_editor"
    )
    
    c_btn1, c_btn2 = st.columns([1, 4])
    if c_btn1.button("💾 Lưu Thay Đổi", type="primary"):
        c = conn.cursor()
        for _, row in edited_products_df.iterrows():
            code_val = str(row['Mã Bánh'])
            name_val = str(row['Tên Bánh'])
            price_val = float(row['Giá Sỉ'])
            shelf_days_val = int(row['HSD Chuẩn (Ngày)'])
            
            c.execute('''
                UPDATE products 
                SET name = ?, unit_price = ?, shelf_life_days = ?
                WHERE code = ?
            ''', (name_val, price_val, shelf_days_val, code_val))
            
            c.execute("SELECT id, import_date FROM inventory_batches WHERE code = ? AND quantity > 0", (code_val,))
            batches = c.fetchall()
            for batch_id, import_date_str in batches:
                imp_date = pd.to_datetime(import_date_str).date()
                new_expiry_date = imp_date + timedelta(days=shelf_days_val)
                c.execute("UPDATE inventory_batches SET expiry_date = ? WHERE id = ?", (new_expiry_date, batch_id))
                
        conn.commit()
        st.success("✅ Đã cập nhật Danh mục bánh và TỰ ĐỘNG ĐỒNG BỘ lại Hạn Sử Dụng của các lô trong kho!")
        st.rerun()

    st.markdown("---")
    
    col_add, col_del = st.columns(2)
    
    with col_add:
        with st.expander("➕ Thêm Mã Bánh Mới"):
            new_code = st.text_input("Mã Bánh mới (ví dụ: BB012)")
            new_name = st.text_input("Tên Bánh mới")
            new_price = st.number_input("Giá Sỉ (VNĐ)", min_value=0.0, step=1000.0, value=10000.0)
            new_life = st.number_input("Hạn sử dụng chuẩn (Số ngày)", min_value=1, value=7)
            
            if st.button("➕ Xác nhận Thêm Bánh"):
                if new_code and new_name:
                    c = conn.cursor()
                    c.execute("INSERT OR REPLACE INTO products VALUES (?, ?, ?, ?)", (new_code.strip(), new_name.strip(), new_price, new_life))
                    conn.commit()
                    st.success(f"Đã thêm bánh {new_code} vào hệ thống!")
                    st.rerun()
                else:
                    st.warning("Vui lòng điền đầy đủ Mã Bánh và Tên Bánh!")
                    
    with col_del:
        with st.expander("🗑 Xóa Bỏ Mã Bánh Không Còn Tồn Tại"):
            list_codes = df_p['Mã Bánh'].tolist()
            if list_codes:
                code_to_del = st.selectbox("Chọn Mã bánh cần xóa:", list_codes)
                st.warning(f"⚠️ Cảnh báo: Xóa mã '{code_to_del}' sẽ xóa cả dữ liệu lịch sử liên quan đến mã này!")
                if st.button("🗑️ Xác nhận Xóa Mã Bánh", type="primary"):
                    c = conn.cursor()
                    c.execute("DELETE FROM products WHERE code = ?", (code_to_del,))
                    c.execute("DELETE FROM inventory_batches WHERE code = ?", (code_to_del,))
                    conn.commit()
                    st.success(f"Đã xóa hoàn toàn mã bánh {code_to_del} khỏi hệ thống!")
                    st.rerun()
            else:
                st.write("Chưa có mã bánh nào để xóa.")
                
    conn.close()
