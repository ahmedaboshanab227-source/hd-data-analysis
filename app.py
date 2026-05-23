import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io

# 1. Page Configuration & Professional Fonts
st.set_page_config(page_title="EXTRA Logistics Analyzer", layout="wide")

# Advanced Custom CSS to match the 'extra' brand image
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background: #f8fafc;
    }

    /* Extra Branding Header - Matching the Image */
    .extra-header {
        background-color: #005da9; /* Extra Blue */
        padding: 30px;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 10px 25px rgba(0,93,169,0.2);
    }
    
    .extra-logo-box {
        display: inline-block;
    }
    
    .arabic-brand {
        color: white;
        font-size: 50px;
        font-weight: 800;
        margin-right: -10px;
        font-family: 'Arial'; /* Standard Arabic Bold */
    }
    
    .yellow-x {
        color: #ffc20e; /* Extra Yellow */
        font-size: 75px;
        font-style: italic;
        font-weight: 900;
        margin: 0 5px;
        display: inline-block;
        transform: skewX(-10deg);
    }
    
    .english-brand {
        color: white;
        font-size: 45px;
        font-weight: 700;
        letter-spacing: -2px;
    }

    /* 3D Neumorphic Cards */
    div[data-testid="stMetricBlock"] {
        background: white !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1) !important;
        border-left: 5px solid #28a745 !important;
    }

    .stDataFrame {
        background: white;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    </style>
    """, unsafe_allow_html=True)

# 2. Top Navigation Bar (Logo Simulation)
st.markdown("""
    <div class="extra-header">
        <div class="extra-logo-box">
            <span class="arabic-brand">اكسترا</span>
            <span class="yellow-x">X</span>
            <span class="english-brand">extra</span>
        </div>
        <div style="color: #dae8f5; font-size: 1.1em; margin-top: 10px; font-weight: 400;">
            LOGISTICS DATA INTELLIGENCE UNIT
        </div>
    </div>
    """, unsafe_allow_html=True)

# 3. File Upload Component
uploaded_file = st.file_uploader("📂 Upload the Logistics Excel Masterfile (.xlsx)", type=["xlsx"])

# Specialized function for Professional Green 3D Charts
def create_green_chart(df, x_col, y_col, title):
    fig = px.bar(df, x=x_col, y=y_col, title=title, text_auto='.0f',
                 color_discrete_sequence=['#28a745']) # Professional Green
    
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        xaxis_title=x_col,
        yaxis_title="Order Volume",
        font=dict(color="#1e293b", family="Inter"),
        title_font=dict(size=20, color="#0f172a", font_family="Inter"),
        hovermode="x unified"
    )
    
    fig.update_traces(
        marker_line_color='#14532d', 
        marker_line_width=1,
        textposition='outside',
        opacity=0.9,
        hovertemplate="<b>%{x}</b><br>Total Unique Orders: %{y}<extra></extra>"
    )
    return fig

# 4. Processing Core
if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        
        required = ['HD', 'Confirmation', 'Return']
        if not all(s in sheets for s in required):
            st.error(f"⚠️ Error: Missing sheets. Ensure file has: {required}")
        else:
            ppt_export_data = {}

            tab_hd, tab_conf, tab_ret = st.tabs(["🚛 HD Analysis", "✔️ Confirmations", "🔄 Returns"])

            # ---------------------------------------------------------
            # HD ANALYSIS
            # ---------------------------------------------------------
            with tab_hd:
                df_hd = pd.read_excel(xls, sheet_name='HD').apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                
                # Math Check: Overall Summary Metrics
                total_orders = df_hd['BOOK ID'].nunique()
                st.metric("Total Unique Orders (HD)", f"{total_orders:,}")
                
                if 'Actual Delivery Date' in df_hd.columns:
                    df_hd['Actual Delivery Date'] = pd.to_datetime(df_hd['Actual Delivery Date'], errors='coerce')
                    df_hd['Month'] = df_hd['Actual Delivery Date'].dt.strftime('%B %Y')

                cols = ['Group Name', 'Region Name', 'Technician', 'Driver', 'Month']
                ppt_export_data['HD'] = {}

                for col in cols:
                    if col in df_hd.columns:
                        summary = df_hd.groupby(col)['BOOK ID'].nunique().reset_index()
                        summary.columns = [col, 'Unique Orders']
                        summary = summary.sort_values('Unique Orders', ascending=False)
                        
                        ppt_export_data['HD'][col] = summary
                        
                        c1, c2 = st.columns([2, 1])
                        with c1:
                            st.plotly_chart(create_green_chart(summary, col, 'Unique Orders', f"HD Performance by {col}"), use_container_width=True)
                        with c2:
                            st.markdown(f"**{col} Stats**")
                            st.dataframe(summary, hide_index=True, use_container_width=True)
                        st.divider()

            # ---------------------------------------------------------
            # CONFIRMATION ANALYSIS
            # ---------------------------------------------------------
            with tab_conf:
                df_conf = pd.read_excel(xls, sheet_name='Confirmation').apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                st.metric("Total Confirmations", f"{df_conf['BOOK ID'].nunique():,}")
                
                ppt_export_data['Confirmation'] = {}
                for col in ['Technician', 'Driver', 'Truck No']:
                    if col in df_conf.columns:
                        summary = df_conf.groupby(col)['BOOK ID'].nunique().reset_index()
                        summary.columns = [col, 'Unique Orders']
                        summary = summary.sort_values('Unique Orders', ascending=False)
                        
                        ppt_export_data['Confirmation'][col] = summary
                        st.plotly_chart(create_green_chart(summary, col, 'Unique Orders', f"Confirmation Count by {col}"), use_container_width=True)
                        st.divider()

            # ---------------------------------------------------------
            # RETURN ANALYSIS
            # ---------------------------------------------------------
            with tab_ret:
                df_ret = pd.read_excel(xls, sheet_name='Return').apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                st.metric("Total Returned Items", f"{df_ret['BOOK ID'].nunique():,}")
                
                if 'Date' in df_ret.columns:
                    df_ret['Month'] = pd.to_datetime(df_ret['Date'], errors='coerce').dt.strftime('%B %Y')
                    summary = df_ret.groupby('Month')['BOOK ID'].nunique().reset_index()
                    summary.columns = ['Month', 'Returns']
                    ppt_export_data['Return'] = summary
                    st.plotly_chart(create_green_chart(summary, 'Month', 'Returns', "Monthly Return Volume"), use_container_width=True)

            # ---------------------------------------------------------
            # PROFESSIONAL POWERPOINT EXPORT
            # ---------------------------------------------------------
            st.markdown("### 📄 Export to Corporate Presentation")
            if st.button("Generate Professional extra Presentation 🚀", use_container_width=True):
                prs = Presentation()
                BLUE_EXTRA = RGBColor(0, 93, 169)
                GREEN_EXTRA = RGBColor(40, 167, 69)
                WHITE = RGBColor(255, 255, 255)

                # Slide 1: Professional Cover
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                # Add Blue background header to Cover
                shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(2))
                shape.fill.solid()
                shape.fill.fore_color.rgb = BLUE_EXTRA
                shape.line.width = 0
                
                title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
                tf = title.text_frame
                p = tf.add_paragraph()
                p.text = "اكسترا X extra"
                p.font.size = Pt(44)
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER

                subtitle = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1))
                p2 = subtitle.text_frame.add_paragraph()
                p2.text = "Logistics Performance Analysis Report"
                p2.font.size = Pt(28)
                p2.font.color.rgb = BLUE_EXTRA
                p2.alignment = PP_ALIGN.CENTER

                # Content Slides
                for section, cats in ppt_export_data.items():
                    if isinstance(cats, dict):
                        for cat_name, df_s in cats.items():
                            slide = prs.slides.add_slide(prs.slide_layouts[6])
                            # Header Bar
                            header_bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.8))
                            header_bar.fill.solid()
                            header_bar.fill.fore_color.rgb = BLUE_EXTRA
                            header_bar.line.width = 0

                            # Slide Title
                            title_shape = slide.shapes.add_textbox(Inches(0.2), Inches(0.1), Inches(8), Inches(0.5))
                            title_shape.text_frame.text = f"extra LOGISTICS | {section} - {cat_name}"
                            title_shape.text_frame.paragraphs[0].font.color.rgb = WHITE
                            title_shape.text_frame.paragraphs[0].font.size = Pt(20)
                            title_shape.text_frame.paragraphs[0].font.bold = True

                            # Data Table
                            df_p = df_s.head(12)
                            rows, cols = df_p.shape[0] + 1, 2
                            table = slide.shapes.add_table(rows, cols, Inches(1), Inches(1.5), Inches(7.5), Inches(4)).table
                            
                            # Table Headers
                            table.cell(0, 0).text = str(df_p.columns[0])
                            table.cell(0, 1).text = "Orders"
                            for c in range(2):
                                table.cell(0, c).fill.solid()
                                table.cell(0, c).fill.fore_color.rgb = GREEN_EXTRA
                                table.cell(0, c).text_frame.paragraphs[0].font.color.rgb = WHITE

                            # Table Content
                            for i, row in df_p.iterrows():
                                table.cell(i+1, 0).text = str(row[0])
                                table.cell(i+1, 1).text = str(row[1])
                    
                buffer = io.BytesIO()
                prs.save(buffer)
                buffer.seek(0)
                st.download_button("📥 Download Official extra Presentation", buffer, "extra_Logistics_Report.pptx")

    except Exception as e:
        st.error(f"❌ Error during analysis: {e}")
