"""
Northern Solar - Web Sistem Pengekstrak DWG & Penjana Laporan PVsyst PDF
=======================================================================
Aliran Kerja 2 Langkah (Diasingkan):
- Langkah 1: Muat naik DWG PDF ➡️ Jana & Muat Turun Fail Excel Template ([NamaPelanggan]_Template.xlsx).
  Drafter buka fail Excel untuk menyemak & mengesahkan data.
- Langkah 2: Muat naik Fail Excel yang telah disemak ➡️ Jana & Muat Turun Laporan PVsyst PDF Bersih.
  Auto-pilih template (1Battery vs NoBattery), 100% bebas highlight kuning, graf & carta dikemaskini.
"""

import os
import io
import re
import openpyxl
import pandas as pd
import streamlit as st
import pymupdf
import dwg_excel_extractor as extractor

# ==============================================================================
# KONFIGURASI HALAMAN STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Northern Solar - PVSyst & Excel Automation System",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Theme Adaptive)
st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 800;
        color: #0284C7;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 14px;
        color: #64748B;
        margin-bottom: 20px;
    }
    .step-card {
        border-radius: 10px;
        padding: 18px;
        border: 1px solid #BAE6FD;
        background-color: #F0F9FF;
        margin-bottom: 20px;
    }
    .download-box {
        border-radius: 10px;
        padding: 20px;
        border: 2px solid #10B981;
        background-color: #ECFDF5;
        margin-top: 15px;
        margin-bottom: 20px;
    }
    .info-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        margin-right: 6px;
    }
    .stDownloadButton button {
        font-size: 16px !important;
        font-weight: 700 !important;
        padding: 12px 28px !important;
        border-radius: 8px !important;
        width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)


def main():
    current_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() and __file__ else os.getcwd()
    default_excel_template = os.path.join(current_dir, "Template.xlsx")
    
    templates_status = {
        "🔋 Battery (2 Ori)": os.path.join(current_dir, "BATTERY (2 Orientation).pdf"),
        "🔋 Battery (3 Ori)": os.path.join(current_dir, "BATTERY (3 Orientation).pdf"),
        "🔋 Battery (4 Ori)": os.path.join(current_dir, "BATTERY (4 Orientation).pdf"),
        "⚡ No-Battery (2 Ori)": os.path.join(current_dir, "NO BATTERY (2 Orientation).pdf"),
        "⚡ No-Battery (3 Ori)": os.path.join(current_dir, "NO BATTERY (3 Orientation).pdf"),
        "⚡ No-Battery (4 Ori)": os.path.join(current_dir, "NO BATTERY (4 Orientation).pdf"),
    }

    # SIDEBAR
    with st.sidebar:
        st.image("https://northernsolar.com.my/wp-content/uploads/2023/06/northern-solar-logo.png", width=220)
        st.markdown("### ☀️ Modul Automasi Drafter")
        st.info("""
        **Aliran Kerja Profesional:**
        
        **Langkah 1 (Tab 1):**
        1. Muat naik fail lukisan DWG PDF.
        2. Sistem isi semua sel kuning dalam fail Excel.
        3. Muat turun & semak fail Excel.
        
        **Langkah 2 (Tab 2):**
        1. Muat naik fail Excel yang telah disahkan.
        2. Sistem auto-pilih daripada **6 Template PVsyst PDF** (2/3/4 Orientasi x Battery/No-Battery).
        3. Laporan PVsyst PDF dijana 100% bersih tanpa highlight kuning dan sifar pertindihan teks.
        """)
        
        st.markdown("---")
        st.markdown("#### ⚙️ Status 6 Template Sistem")
        st.write(f"📊 **Excel Template:** `Template.xlsx` {'✅' if os.path.exists(default_excel_template) else '❌'}")
        for t_label, t_path in templates_status.items():
            t_fname = os.path.basename(t_path)
            st.write(f"{t_label}: `{t_fname}` {'✅' if os.path.exists(t_path) else '❌'}")
        
        st.markdown("---")
        st.caption("🔒 Formula, PVsyst Graph Data & GlobHor_DiffHor 100% dilindungi.")

    # HEADER
    st.markdown('<div class="main-title">☀️ NORTHERN SOLAR - SISTEM PENJANAAN PVSYST & EXCEL</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Automasi pemprosesan lukisan teknikal DWG AutoCAD kepada fail Excel Template dan penjanaan semula Laporan Simulasi PVsyst PDF (Matriks 6 Template).</div>', unsafe_allow_html=True)

    if not os.path.exists(default_excel_template):
        st.error("Ralat: Fail Template.xlsx tidak dijumpai dalam folder projek.")
        return

    # TABS UTAMA (LANGKAH 1 VS LANGKAH 2)
    tab1, tab2 = st.tabs([
        "📄 LANGKAH 1: Muat Naik DWG PDF ➡️ Jana Excel Template",
        "📊 LANGKAH 2: Muat Naik Excel Disemak ➡️ Jana Laporan PVsyst PDF"
    ])

    # ==========================================================================
    # TAB 1: DWG PDF ➡️ EXCEL TEMPLATE
    # ==========================================================================
    with tab1:
        st.markdown("""
        <div class="step-card">
            <h4>📄 Langkah 1: Pengekstrakan Lukisan DWG PDF ke Template Excel</h4>
            Muat naik fail lukisan <b>DWG PDF AutoCAD</b> untuk mengisi semua sel kuning dalam <code>Template.xlsx</code> secara automatik (Nama Pelanggan, Alamat Ringkas, Orientasi Roof Fall 1-4, Inverter, Modul PV, dan Bateri).
        </div>
        """, unsafe_allow_html=True)

        uploaded_dwg = st.file_uploader(
            "Pilih atau Tarik (Drag & Drop) Fail Lukisan DWG PDF (.pdf)",
            type=["pdf"],
            key="tab1_dwg_uploader"
        )

        if uploaded_dwg:
            with st.spinner("🔍 Sedang mengekstrak maklumat lukisan DWG PDF..."):
                try:
                    dwg_bytes = uploaded_dwg.getvalue()
                    dwg_data = extractor.extract_dwg_info(dwg_bytes)
                    project_data = extractor.combine_project_data(dwg_data=dwg_data)
                except Exception as e:
                    st.error(f"Ralat mengekstrak fail DWG: {str(e)}")
                    return

            client_name = project_data.get("client_name") or "Pelanggan"
            safe_name = re.sub(r'[\\/*?:"<>|]', '', client_name).strip() or "Projek"
            kwp_val = project_data.get("system_size_kwp", 0.0)

            st.success(f"✅ Berjaya mengekstrak maklumat projek: **{client_name}** ({kwp_val} kWp)")

            # Ringkasan Metrik
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Nama Pelanggan", client_name)
            m2.metric("Kapasiti Sistem", f"{kwp_val} kWp", f"{project_data.get('system_size_kwac', 0.0)} kWac")
            m3.metric("Jumlah Panel", f"{project_data['pv_module'].get('total_panels', 0)} PCS", f"{project_data['pv_module'].get('unit_power_w', 630)}W")
            m4.metric("Model Inverter", project_data['inverter'].get('model') or "N/A", f"{project_data['inverter'].get('unit_power_kw', 0)} kW")
            m5.metric("Bateri", f"{project_data['battery'].get('units', 0)} Unit", f"{project_data['battery'].get('nominal_energy_kwh', 0)} kWh")

            # Isi Sel Kuning Excel
            with open(default_excel_template, "rb") as f:
                excel_tmpl_bytes = f.read()

            tmpl_wb = openpyxl.load_workbook(io.BytesIO(excel_tmpl_bytes))
            extractor.fill_yellow_cells(tmpl_wb, project_data)
            out_stream = io.BytesIO()
            tmpl_wb.save(out_stream)
            filled_excel_bytes = out_stream.getvalue()

            excel_filename = f"{safe_name}.xlsx"

            # Download Box
            st.markdown('<div class="download-box">', unsafe_allow_html=True)
            st.markdown(f"### 📥 Muat Turun Fail Excel: `{excel_filename}`")
            st.write("Semua sel kuning telah diisi mengikut peraturan teknikal lukisan. Sila muat turun fail ini untuk disemak oleh drafter:")

            st.download_button(
                label=f"⬇️ MUAT TURUN FAIL EXCEL ({excel_filename})",
                data=filled_excel_bytes,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_excel_tab1"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            # Semakan Sel Kuning
            with st.expander("🔍 Klik untuk melihat butiran sel kuning yang diisi dalam fail Excel", expanded=False):
                mapped_dict = extractor.get_mapped_values_for_template(project_data)
                rows = []
                for (s_name, coord), val in mapped_dict.items():
                    rows.append({
                        "Sheet": s_name,
                        "Sel": coord,
                        "Nilai Dimasukkan": str(val)
                    })
                df_map = pd.DataFrame(rows)
                st.dataframe(df_map, use_container_width=True, hide_index=True)

            st.info("👉 **Langkah Seterusnya:** Selepas drafter menyemak dan menyimpan fail Excel ini, sila klik pada **LANGKAH 2** di atas untuk menjana Laporan PVsyst PDF.")

        else:
            st.info("👈 Sila muat naik fail lukisan DWG PDF di atas untuk memulakan.")


    # ==========================================================================
    # TAB 2: EXCEL DISEDIAKAN ➡️ PVSYST PDF REPORT
    # ==========================================================================
    with tab2:
        st.markdown("""
        <div class="step-card">
            <h4>📊 Langkah 2: Penjanaan Laporan PVsyst PDF daripada Fail Excel Disemak</h4>
            Muat naik fail <b>Excel yang telah disemak oleh drafter</b>. Sistem akan mengesan konfigurasi sistem secara automatik mengikut matriks 6 template:
            <ul>
                <li>🔋 <b>Auto-Pilih Template:</b>
                    <ul>
                        <li><b>2 Orientasi:</b> <code>BATTERY (2 Orientation).pdf</code> (jika ada bateri) / <code>NO BATTERY (2 Orientation).pdf</code> (jika tiada bateri)</li>
                        <li><b>3 Orientasi:</b> <code>BATTERY (3 Orientation).pdf</code> (jika ada bateri) / <code>NO BATTERY (3 Orientation).pdf</code> (jika tiada bateri)</li>
                        <li><b>4 Orientasi:</b> <code>BATTERY (4 Orientation).pdf</code> (jika ada bateri) / <code>NO BATTERY (4 Orientation).pdf</code> (jika tiada bateri)</li>
                    </ul>
                </li>
                <li>✨ <b>100% Bebas Highlight Kuning:</b> Semua kotak highlight kuning dan anotasi dipadam sepenuhnya.</li>
                <li>📈 <b>Format Kemas & Sifar Pertindihan:</b> Teks dan formula diganti dengan tepat tanpa sebarang overlap.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        uploaded_checked_excel = st.file_uploader(
            "Pilih atau Tarik (Drag & Drop) Fail Excel Yang Telah Disemak (.xlsx)",
            type=["xlsx", "xls"],
            key="tab2_excel_uploader"
        )

        if uploaded_checked_excel:
            with st.spinner("🔍 Sedang membaca maklumat daripada fail Excel..."):
                try:
                    excel_bytes = uploaded_checked_excel.getvalue()
                    checked_excel_data = extractor.extract_excel_info(excel_bytes)
                    checked_project_data = extractor.combine_project_data(excel_data=checked_excel_data)
                except Exception as e:
                    st.error(f"Ralat membaca fail Excel: {str(e)}")
                    return

            client_name_ex = checked_project_data.get("client_name") or "Pelanggan"
            safe_name_ex = re.sub(r'[\\/*?:"<>|]', '', client_name_ex).strip() or "Projek"
            kwp_ex = checked_project_data.get("system_size_kwp", 0.0)
            bat_units_ex = checked_project_data.get("battery", {}).get("units", 0)
            has_bat = bat_units_ex > 0

            # Pemilihan Template Automatik (Matriks 6 Template)
            target_tmpl_path = extractor.select_pvsyst_template(checked_project_data)
            target_tmpl_name = os.path.basename(target_tmpl_path)

            st.success(f"✅ Data projek dimuatkan: **{client_name_ex}** ({kwp_ex} kWp)")

            # Ringkasan Maklumat Disemak
            col_inf1, col_inf2, col_inf3, col_inf4 = st.columns(4)
            col_inf1.metric("Pelanggan", client_name_ex)
            col_inf1.caption(f"📍 {checked_project_data.get('short_address', '-')}")
            
            col_inf2.metric("Kapasiti PV / AC", f"{kwp_ex} kWp", f"{checked_project_data.get('system_size_kwac', 0.0)} kWac")
            col_inf2.caption(f"☀️ Orientasi: {len(checked_project_data.get('orientations', []))} falls")
            
            pr_val = checked_project_data.get('perf_ratio_pr', 0.0)
            col_inf3.metric("Performance Ratio (PR)", f"{pr_val:.2f} %", f"Ratio: {round(pr_val/100.0, 3):.3f}")
            col_inf3.caption(f"⚡ Tenaga: {checked_project_data.get('produced_energy_kwh', 0):,.0f} kWh/thn")
            
            col_inf4.metric("Konfigurasi Bateri", f"{bat_units_ex} Unit" if has_bat else "Tiada Bateri", f"Template: {target_tmpl_name}")
            col_inf4.caption(f"🔋 {f'{bat_units_ex}-Battery System' if has_bat else 'No-Battery System'}")

            # Penjanaan PDF Bersih
            with st.spinner("📑 Sedang menjana semula Laporan PVsyst PDF bersih..."):
                try:
                    clean_pdf_bytes = extractor.generate_pvsyst_pdf(target_tmpl_path, checked_project_data)
                except Exception as e:
                    st.error(f"Ralat semasa menjana PDF PVsyst: {str(e)}")
                    return

            pdf_out_name = f"{safe_name_ex}.pdf"

            # Download Box
            st.markdown('<div class="download-box">', unsafe_allow_html=True)
            st.markdown(f"### 📑 Muat Turun Laporan PVsyst PDF: `{pdf_out_name}`")
            st.write(f"Laporan PDF rasmi telah dijana berdasarkan template **{target_tmpl_name}**. Semua highlight kuning telah dipadam dan maklumat projek baharu telah dimasukkan secara kemas:")

            st.download_button(
                label=f"⬇️ MUAT TURUN LAPORAN PVSYST PDF ({pdf_out_name})",
                data=clean_pdf_bytes,
                file_name=pdf_out_name,
                mime="application/pdf",
                key="btn_download_pdf_tab2"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            # Preview Pages
            st.markdown("#### 👁️ Praterap Halaman Laporan PVsyst Yang Dijana")
            doc_preview = pymupdf.open(stream=clean_pdf_bytes, filetype="pdf")
            p_cols = st.columns(3)
            
            # Cover Page
            with p_cols[0]:
                st.caption("Muka Surat 1: Cover Projek")
                img1 = doc_preview[0].get_pixmap(dpi=120).tobytes("png")
                st.image(img1, use_container_width=True)

            # Summary Page
            with p_cols[1]:
                st.caption("Muka Surat 2: System Summary")
                if len(doc_preview) >= 2:
                    img2 = doc_preview[1].get_pixmap(dpi=120).tobytes("png")
                    st.image(img2, use_container_width=True)

            # P50 Page
            with p_cols[2]:
                st.caption(f"Muka Surat {len(doc_preview)-1}: P50-P90 Evaluation")
                img_p50 = doc_preview[-2 if len(doc_preview)>=10 else -1].get_pixmap(dpi=120).tobytes("png")
                st.image(img_p50, use_container_width=True)

        else:
            st.info("👈 Sila muat naik fail Excel projek yang telah disemak di atas untuk menjana Laporan PVsyst PDF.")


if __name__ == "__main__":
    main()
