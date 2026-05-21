import streamlit as st
import pandas as pd
import plotly.express as px

# إعدادات الصفحة لتكون عريضة ومريحة للعين
st.set_page_config(page_title="نظام تحليل بيانات HD", layout="wide")

st.title("📊 لوحة تحكم وتحليل بيانات شيت HD")
st.write("قم برفع ملف الإكسيل ليتولى النظام تحليل الأعمدة المطلوبة بناءً على إجمالي الكميات (Qty).")

# أداة رفع الملف
uploaded_file = st.file_uploader("اختر ملف الإكسيل الخاص بك", type=["xlsx"])

if uploaded_file is not None:
    try:
        # قراءة شيت 'HD' تحديداً من ملف الإكسيل
        df = pd.read_excel(uploaded_file, sheet_name='HD')
        
        st.success("✅ تم تحميل شيت 'HD' بنجاح!")
        
        # التأكد من وجود عمود الكمية Qty
        if 'Qty' in df.columns:
            # تحويل عمود Qty إلى أرقام وتعبئة الفراغات بـ 0 لضمان دقة الحسابات بنسبة 100%
            df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
            
            # قائمة الأعمدة المراد تحليلها بناءً على طلبك
            target_columns = ['Group Name', 'Region Name', 'Store Code', 'Technician', 'Truck No']
            
            # التأكد من وجود الأعمدة في الملف
            available_columns = [col for col in target_columns if col in df.columns]
            
            if available_columns:
                # إنشاء قائمة منسدلة ليختار المستخدم أي عمود يريد رؤية تحليله الآن
                st.markdown("---")
                st.subheader("🎯 اختر البُعد الفني الذي تريد تحليله:")
                selected_col = st.selectbox("اضغط هنا لاختيار العمود:", available_columns)
                
                # تجميع البيانات الحسابية: حساب إجمالي الـ Qty والنسبة المئوية بدقة
                summary_df = df.groupby(selected_col)['Qty'].sum().reset_index()
                # ترتيب تنازلي من الأكثر للأقل
                summary_df = summary_df.sort_values(by='Qty', ascending=False)
                
                # حساب النسبة المئوية الدقيقة (ظهور الفروع والأعشار بنسبة 100%)
                total_qty = summary_df['Qty'].sum()
                if total_qty > 0:
                    summary_df['Percentage (%)'] = (summary_df['Qty'] / total_qty) * 100
                else:
                    summary_df['Percentage (%)'] = 0

                # عرض النتائج في قسمين بجانب بعضهما
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.markdown(f"### 📋 جدول البيانات لـ {selected_col}")
                    # استعراض الجدول مع إظهار الفواصل العشرية دقيقة
                    st.dataframe(summary_df.style.format({'Qty': '{:,.2f}', 'Percentage (%)': '{:.2f}%'}), use_container_width=True)
                
                with col2:
                    st.markdown(f"### 📊 الرسم البياني بالأعمدة لـ {selected_col}")
                    fig = px.bar(
                        summary_df, 
                        x=selected_col, 
                        y='Qty', 
                        title=f"إجمالي الكميات (Qty) حسب {selected_col}",
                        text_auto='.2f', # ظهور الكسور فوق الأعمدة بدقة تامة
                        color='Qty',
                        color_continuous_scale=px.colors.sequential.Viridis
                    )
                    fig.update_layout(xaxis_title=selected_col, yaxis_title="إجمالي الكمية (Qty)")
                    st.plotly_chart(fig, use_container_width=True)
                
                # إضافة رسمة دائرية إضافية للنسب المئوية تحتهم
                st.markdown("---")
                st.markdown(f"### 🍕 توزيع النسب المئوية لـ {selected_col}")
                fig_pie = px.pie(
                    summary_df, 
                    names=selected_col, 
                    values='Qty', 
                    title=f"نسبة مساهمة كل {selected_col} من إجمالي الـ Qty"
                )
                st.plotly_chart(fig_pie, use_container_width=True)
                
            else:
                st.error("⚠️ لم نجد أي من الأعمدة المطلوبة (Group Name, Region Name, Store Code, Technician, Truck No) داخل شيت HD. يرجى مراجعة وتدقيق أسماء الأعمدة.")
        else:
            st.error("⚠️ عمود الكمية 'Qty' غير موجود في شيت HD، وهو أساسي لإجراء التحليلات.")
            
    except ValueError as ve:
        st.error("⚠️ لم يتم العثور على شيت باسم 'HD' داخل هذا الملف. تأكد من أن اسم الشيت مكتوب كابيتال وبدون مسافات زائدة.")
    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
else:
    st.info("💡 في انتظار رفع ملف الإكسيل لقراءة شيت HD وبدء التحليل الفوري...")