import streamlit as st
import pandas as pd
import io
import calendar

st.set_page_config(page_title="ระบบสรุปยอดทำงานลูกจ้าง (แบบ 51)", layout="wide")

# ==========================================
# 1. ข้อมูลตั้งต้น
# ==========================================
months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

# ตัวเลือกในหน้า UI (กรอกแบบไม่มีเครื่องหมายทับ)
shift_codes = ["", "/", "พ", "ป", "น", "ย", "2น", "2ย"]

st.title("📝 ระบบสรุปยอดทำงานและวันหยุดลูกจ้าง (Export แบบ 51)")

# ==========================================
# 2. ตั้งค่าเดือนและสร้างตาราง Input
# ==========================================
with st.container(border=True):
    st.subheader("⚙️ 1. เลือกเดือนปัจจุบันที่ต้องการเบิกเงิน")
    c1, c2 = st.columns(2)
    with c1:
        target_month_name = st.selectbox("เดือนปัจจุบัน", months_list, index=7) # ค่าเริ่มต้น สิงหาคม
    with c2:
        target_year_be = st.number_input("ปี พ.ศ.", value=2569)
        
    target_month_idx = months_list.index(target_month_name) + 1
    
    if target_month_idx == 1:
        prev_month_idx = 12
        prev_year_be = target_year_be - 1
    else:
        prev_month_idx = target_month_idx - 1
        prev_year_be = target_year_be
        
    prev_month_name = months_list[prev_month_idx - 1]
    _, num_days_prev = calendar.monthrange(prev_year_be - 543, prev_month_idx)

st.markdown("### ✍️ 2. ตารางกรอกข้อมูลลงเวลา (46 วัน)")
st.info(f"**คำแนะนำ:** ให้กรอกข้อมูลของเดือนก่อนหน้า ({prev_month_name}) ให้ครบเดือน และกรอกของเดือนปัจจุบัน ({target_month_name}) วันที่ 1-15")

config = {"ชื่อ-สกุล": st.column_config.TextColumn("ชื่อ-สกุล", width="medium")}
columns_list = ["ชื่อ-สกุล"]

# สร้างคอลัมน์เดือนก่อน (1 ถึง สิ้นเดือน)
prev_cols = [f"P{d}" for d in range(1, num_days_prev + 1)]
for d, col in enumerate(prev_cols, 1):
    columns_list.append(col)
    config[col] = st.column_config.SelectboxColumn(f"{d} {prev_month_name[:3]}.", options=shift_codes, width="small")

# สร้างคอลัมน์เดือนปัจจุบัน (1 ถึง 15)
curr_cols = [f"C{d}" for d in range(1, 16)]
for d, col in enumerate(curr_cols, 1):
    columns_list.append(col)
    config[col] = st.column_config.SelectboxColumn(f"{d} {target_month_name[:3]}.", options=shift_codes, width="small")

if 'form51_data' not in st.session_state:
    st.session_state.form51_data = pd.DataFrame(columns=columns_list)
else:
    current_df = st.session_state.form51_data
    new_df = pd.DataFrame(columns=columns_list)
    if not current_df.empty and "ชื่อ-สกุล" in current_df.columns:
        new_df["ชื่อ-สกุล"] = current_df["ชื่อ-สกุล"]
    st.session_state.form51_data = new_df

edited_df = st.data_editor(
    st.session_state.form51_data,
    num_rows="dynamic",
    column_config=config,
    use_container_width=True,
    height=400
)

# ==========================================
# 3. ระบบคำนวณและ Export ไปยัง Excel
# ==========================================
st.markdown("---")
if st.button("📊 คำนวณและส่งออกไฟล์ Excel (แบบ 51)", type="primary", use_container_width=True):
    if edited_df.empty or edited_df["ชื่อ-สกุล"].isnull().all():
        st.warning("กรุณากรอกชื่อพนักงานอย่างน้อย 1 คน")
    else:
        work_data = []    
        holiday_data = [] 

        # 🧠 กฎการแยกตัวหนังสือที่ถูกต้อง 100% ตามที่คุณสั่ง
        def map_work_sheet(v):
            v = str(v).strip()
            if v == "2น": return "น"
            if v == "2ย": return "ย"
            if v in ["/", "พ", "ป", "น", "ย"]: return v
            return ""

        def map_holiday_sheet(v):
            v = str(v).strip()
            if v == "2น": return "2/น"  
            if v == "2ย": return "2/ย"  
            return "" 
        
        for idx, row in edited_df.iterrows():
            name = str(row.get("ชื่อ-สกุล", "")).strip()
            if not name or name == "nan": continue
            
            # --- ดึงข้อมูลดิบจากตาราง 46 วัน ---
            curr_1_15_raw = [str(row.get(c, "")).strip() for c in curr_cols]
            prev_1_15_raw = [str(row.get(prev_cols[i], "")).strip() for i in range(0, 15)]
            prev_16_end_raw = [str(row.get(prev_cols[i], "")).strip() for i in range(15, num_days_prev)]
            
            # ========================================================
            # 📁 สรุปข้อมูล ชีท "ค่าทำงาน" (รอบ 16 ด.ก่อน - 15 ด.นี้)
            # ========================================================
            # แปลงรหัส
            work_curr_1_15 = [map_work_sheet(v) for v in curr_1_15_raw]
            work_prev_16_end = [map_work_sheet(v) for v in prev_16_end_raw]
            
            working_range = work_curr_1_15 + work_prev_16_end
            
            count_normal = working_range.count("/")
            count_n = working_range.count("น")
            count_vac = working_range.count("พ")
            count_sick = working_range.count("ป")
            
            total_prb = count_n + count_vac + count_sick
            
            row_work = {
                "ที่": len(work_data) + 1,
                "ชื่อ-นามสกุล": name,
            }
            # เรียงคอลัมน์: 1-15 (เดือนนี้)
            for i in range(1, 16): row_work[str(i)] = work_curr_1_15[i-1]
            # เรียงคอลัมน์: 16-สิ้นเดือน (เดือนก่อน)
            for i in range(16, num_days_prev + 1): row_work[str(i)] = work_prev_16_end[i-16]
            
            row_work["ปกติ"] = count_normal
            row_work["พรบ."] = total_prb
            row_work["373"] = count_normal
            row_work[".1 (ป)"] = count_sick
            row_work[".2 (พ)"] = count_vac
            row_work[".6 (น)"] = count_n
            
            work_data.append(row_work)
            
            # ========================================================
            # 🏖️ สรุปข้อมูล ชีท "วันหยุด" (รอบ 1 - สิ้นเดือน ด.ก่อน)
            # ========================================================
            # แปลงรหัส
            holiday_prev_1_15 = [map_holiday_sheet(v) for v in prev_1_15_raw]
            holiday_prev_16_end = [map_holiday_sheet(v) for v in prev_16_end_raw]
            
            holiday_range = holiday_prev_1_15 + holiday_prev_16_end
            
            count_2n = holiday_range.count("2/น")
            count_2y = holiday_range.count("2/ย")
            
            val_6 = count_2n * 1
            val_7 = count_2y * 2
            total_holiday_prb = val_6 + val_7
            
            row_holiday = {
                "ที่": len(holiday_data) + 1,
                "ชื่อ-นามสกุล": name,
            }
            # เรียงคอลัมน์: 1-15 (เดือนก่อน)
            for i in range(1, 16): row_holiday[str(i)] = holiday_prev_1_15[i-1]
            # เรียงคอลัมน์: 16-สิ้นเดือน (เดือนก่อน)
            for i in range(16, num_days_prev + 1): row_holiday[str(i)] = holiday_prev_16_end[i-16]
            
            row_holiday["พรบ."] = total_holiday_prb
            row_holiday[".6 (2/น)"] = val_6
            row_holiday[".7 (2/ย)"] = val_7
            
            holiday_data.append(row_holiday)

        # --- สร้างไฟล์ Excel ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_work = pd.DataFrame(work_data)
            df_holiday = pd.DataFrame(holiday_data)
            
            df_work.to_excel(writer, index=False, sheet_name='ค่าทำงาน')
            df_holiday.to_excel(writer, index=False, sheet_name='วันหยุด')
            
            workbook = writer.book
            worksheet_w = writer.sheets['ค่าทำงาน']
            worksheet_h = writer.sheets['วันหยุด']
            
            format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
            worksheet_w.set_column('A:A', 5, format_center)
            worksheet_w.set_column('B:B', 25)
            worksheet_w.set_column('C:BZ', 5, format_center)  
            
            worksheet_h.set_column('A:A', 5, format_center)
            worksheet_h.set_column('B:B', 25)
            worksheet_h.set_column('C:BZ', 5, format_center)

        output.seek(0)
        
        st.success("✅ คำนวณเสร็จสมบูรณ์! ข้อมูลจากตาราง 46 วัน ถูกดึงไปสร้าง 2 ชีทตามกฎเรียบร้อยแล้ว")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (สรุปยอดแบบ 51)",
            data=output,
            file_name=f"สรุปยอด_{target_month_name}_{target_year_be}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
