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
# GIAO DIỆN CHÍNH (SIDEBAR MENU)
# ---------------------------------------------------------
st.sidebar.title("🧁 Mecake Manager")
menu = st.sidebar.radio(
    "Chọn chức năng", 
    ["1. Soạn thảo & Nhập Order", "2. Báo cáo Tồn Kho & Hạn Sử Dụng", "3. Nhập Kho Lô Mới", "4. Quản lý Danh mục Bánh"]
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
                # Xóa đơn trùng ngày nếu có để cập nhật đè mới
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

    # TAB 2: LỊCH SỬ, XUẤT FILE VÀ XÓA ĐƠN ORDER
    with tab_history:
        st.header("📦 Quản Lý, Xuất & Xóa Đơn Order")
        
        df_all_dates = pd.read_sql_query("SELECT DISTINCT order_date FROM order_records ORDER BY order_date DESC", conn)
        
        if df_all_dates.empty:
            st.info("Chưa có đơn order nào được chốt trong hệ thống.")
        else:
            selected_date = st.selectbox("Chọn ngày đã chốt Order để xem:", df_all_dates['order_date'].tolist())
            
            query_detail = '''
                SELECT 
                    o.code AS 'Mã Bánh',
                    p.name AS 'Tên Bánh',
                    p.unit_price AS 'Giá Sỉ',
                    p.shelf_life_days AS 'HSD Chuẩn',
                    o.order_qty AS 'Số Lượng Order',
                    (p.unit_price * o.order_qty) AS 'Thành Tiền'
                FROM order_records o
                JOIN products p ON o.code = p.code
                WHERE o.order_date = ? AND o.order_qty > 0
            '''
            df_detail = pd.read_sql_query(query_detail, conn, params=(selected_date,))
            
            st.subheader(f"Danh sách Bánh Order cho ngày: {selected_date}")
            st.dataframe(df_detail[['Mã Bánh', 'Tên Bánh', 'Giá Sỉ', 'Số Lượng Order', 'Thành Tiền']], use_container_width=True, hide_index=True)
            
            sum_qty = df_detail['Số Lượng Order'].sum()
            sum_money = df_detail['Thành Tiền'].sum()
            st.markdown(f"**Tổng cộng:** **{sum_qty}** cái | **Tổng tiền:** **{sum_money:,.0f} VNĐ**")
            
            st.markdown("---")
            col_dl, col_import, col_delete = st.columns(3)
            
            # 1. Tải file Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_detail[['Mã Bánh', 'Tên Bánh', 'Giá Sỉ', 'Số Lượng Order', 'Thành Tiền']].to_excel(writer, index=False, sheet_name='Don_Order')
            
            col_dl.download_button(
                label="📥 Tải Excel Đơn Order",
                data=buffer.getvalue(),
                file_name=f"Don_Order_Mecake_{selected_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            # 2. Chuyển thành Lô Nhập Kho
            if col_import.button("🚚 CHUYỂN THÀNH LÔ NHẬP KHO", type="secondary", use_container_width=True):
                c = conn.cursor()
                import_today = date.today()
                count_imported = 0
                for _, row in df_detail.iterrows():
                    code = str(row['Mã Bánh'])
                    qty = int(row['Số Lượng Order'])
                    shelf_days = int(row['HSD Chuẩn']) if pd.notnull(row['HSD Chuẩn']) else 7
                    expiry_date = import_today + timedelta(days=shelf_days)
                    
                    c.execute('''
                        INSERT INTO inventory_batches (code, import_date, expiry_date, quantity)
                        VALUES (?, ?, ?, ?)
                    ''', (code, import_today, expiry_date, qty))
                    count_imported += 1
                conn.commit()
                st.success(f"🎉 Đã chuyển toàn bộ {count_imported} mã bánh từ Đơn Order ngày {selected_date} thành Lô Nhập Kho ngày hôm nay!")
                st.info("💡 Bạn có thể kiểm tra tồn kho & HSD ở mục '2. Báo cáo Tồn Kho & Hạn Sử Dụng'.")

            # 3. XÓA ĐƠN ORDER BỊ SAI
            if col_delete.button("🗑️ XÓA HOÀN TOÀN ĐƠN ORDER NÀY", type="primary", use_container_width=True):
                c = conn.cursor()
                c.execute("DELETE FROM order_records WHERE order_date = ?", (selected_date,))
                conn.commit()
                st.success(f"🗑️ Đã xóa hoàn toàn dữ liệu đơn Order ngày {selected_date}!")
                st.rerun()

    conn.close()

# ---------------------------------------------------------
# CHỨC NĂNG 2: BÁO CÁO TỒN KHO THEO THỜI GIAN & CẢNH BÁO HSD
# ---------------------------------------------------------
elif menu == "2. Báo cáo Tồn Kho & Hạn Sử Dụng":
    st.header("⏳ Kiểm Soát Tồn Kho & Hạn Sử Dụng Theo Thời Gian")
    
    conn = sqlite3.connect(DB_FILE)
    query = '''
        SELECT 
            b.id AS 'ID Lô',
            b.code AS 'Mã Bánh',
            p.name AS 'Tên Bánh',
            b.import_date AS 'Ngày Nhập',
            b.expiry_date AS 'Hạn Sử Dụng',
            b.quantity AS 'Số Lượng Tồn'
        FROM inventory_batches b
        JOIN products p ON b.code = p.code
        WHERE b.quantity > 0
        ORDER BY b.expiry_date ASC
    '''
    df_batches = pd.read_sql_query(query, conn)
    
    if df_batches.empty:
        st.info("Hiện chưa có lô hàng nào trong kho. Vui lòng sang tab 'Nhập Kho Lô Mới' hoặc 'Chuyển đơn order thành lô nhập kho' để thêm hàng!")
    else:
        today = date.today()
        df_batches['Hạn Sử Dụng'] = pd.to_datetime(df_batches['Hạn Sử Dụng']).dt.date
        df_batches['Số Ngày Còn Lại'] = (pd.to_datetime(df_batches['Hạn Sử Dụng']) - pd.to_datetime(today)).dt.days
        
        def set_status(days):
            if days < 0:
                return "🔴 Đã Hết Hạn"
            elif days <= 2:
                return "🟡 Cận Date (Cần Bán Gấp)"
            else:
                return "🟢 An Toàn"
                
        df_batches['Trạng Thái'] = df_batches['Số Ngày Còn Lại'].apply(set_status)
        
        expired_count = len(df_batches[df_batches['Số Ngày Còn Lại'] < 0])
        warning_count = len(df_batches[(df_batches['Số Ngày Còn Lại'] >= 0) & (df_batches['Số Ngày Còn Lại'] <= 2)])
        safe_count = len(df_batches[df_batches['Số Ngày Còn Lại'] > 2])
        
        c1, c2, c3 = st.columns(3)
        c1.metric("🟢 Lô An Toàn", f"{safe_count} lô")
        c2.metric("🟡 Lô Cận Date (≤2 ngày)", f"{warning_count} lô")
        c3.metric("🔴 Lô Hết Hạn", f"{expired_count} lô")
        
        st.subheader("Chi tiết Tồn kho từng Lô bánh (Ưu tiên Lô gần hết hạn lên đầu):")
        st.dataframe(
            df_batches[['ID Lô', 'Mã Bánh', 'Tên Bánh', 'Ngày Nhập', 'Hạn Sử Dụng', 'Số Ngày Còn Lại', 'Số Lượng Tồn', 'Trạng Thái']],
            use_container_width=True,
            hide_index=True
        )
    conn.close()

# ---------------------------------------------------------
# CHỨC NĂNG 3: NHẬP KHO LÔ MỚI
# ---------------------------------------------------------
elif menu == "3. Nhập Kho Lô Mới":
    st.header("📥 Nhập Kho Lô Bánh Mới (Tự động tính HSD)")
    
    conn = sqlite3.connect(DB_FILE)
    df_prod = pd.read_sql_query("SELECT code, name, shelf_life_days FROM products", conn)
    
    if df_prod.empty:
        st.warning("Chưa có danh mục bánh. Vui lòng thêm bánh ở menu 'Quản lý Danh mục Bánh'!")
    else:
        prod_dict = {f"{row['code']} - {row['name']}": (row['code'], row['shelf_life_days']) for _, row in df_prod.iterrows()}
        selected_prod = st.selectbox("Chọn Mã bánh nhập kho:", list(prod_dict.keys()))
        code, default_shelf_life = prod_dict[selected_prod]
        
        col1, col2 = st.columns(2)
        import_date = col1.date_input("Ngày Nhập Kho", date.today())
        quantity = col2.number_input("Số Lượng Nhập", min_value=1, value=15, step=1)
        
        calculated_expiry = import_date + timedelta(days=default_shelf_life)
        expiry_date = st.date_input("Hạn Sử Dụng (Tự động đề xuất theo cấu hình bánh)", calculated_expiry)
        
        if st.button("➕ Xác Nhận Nhập Kho Lô Hàng"):
            c = conn.cursor()
            c.execute('''
                INSERT INTO inventory_batches (code, import_date, expiry_date, quantity)
                VALUES (?, ?, ?, ?)
            ''', (code, import_date, expiry_date, int(quantity)))
            conn.commit()
            st.success(f"Đã nhập thành công {quantity} bánh mã {code} (HSD: {expiry_date}) vào hệ thống kho!")
    conn.close()

# ---------------------------------------------------------
# CHỨC NĂNG 4: QUẢN LÝ DANH MỤC BÁNH
# ---------------------------------------------------------
elif menu == "4. Quản lý Danh mục Bánh":
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
            c.execute('''
                UPDATE products 
                SET name = ?, unit_price = ?, shelf_life_days = ?
                WHERE code = ?
            ''', (str(row['Tên Bánh']), float(row['Giá Sỉ']), int(row['HSD Chuẩn (Ngày)']), str(row['Mã Bánh'])))
        conn.commit()
        st.success("Đã cập nhật thành công các thông tin bánh!")
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
                st.warning(f"⚠️️ Cảnh báo: Xóa mã '{code_to_del}' sẽ xóa cả dữ liệu lịch sử liên quan đến mã này!")
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
