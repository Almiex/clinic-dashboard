import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta

# ==============================================================================
# НАСТРОЙКА СТРАНИЦЫ
# ==============================================================================
st.set_page_config(page_title="Аналитический отчет клиники: загрузка специализаций", layout="wide")

st.markdown("""
    <style>
    .clinic-header { 
        background-color: #f8f9fa; 
        border-left: 5px solid #B5838D; 
        padding: 20px; 
        margin-bottom: 25px; 
        border-radius: 4px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    .clinic-title { font-size: 22px; font-weight: bold; color: #2B2D42; text-transform: uppercase; letter-spacing: 0.5px;}
    .clinic-subtitle { font-size: 14px; color: #6C757D; margin-top: 6px; font-weight: 500;}
    
    .custom-dash-table {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        border-collapse: collapse;
        width: 100%;
        background-color: #FFFFFF;
        color: #4A4A4A;
        margin-top: 15px;
        margin-bottom: 30px;
        font-size: 13px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        border-radius: 4px;
        overflow: hidden;
    }
    .custom-dash-table th {
        background-color: #B5838D;
        color: #FFFFFF !important;
        font-weight: bold;
        text-align: center !important;
        padding: 12px 10px;
        border: 1px solid #B5838D;
        vertical-align: middle;
        line-height: 1.2;
    }
    .custom-dash-table td {
        padding: 10px;
        border-bottom: 1px solid #E8D5D5;
        text-align: center !important;
        vertical-align: middle;
    }
    .custom-dash-table td:first-child {
        text-align: center !important;
        font-weight: bold;
        background-color: #FAF6F2;
        border-right: 1px solid #E8D5D5;
    }
    .custom-dash-table tr:nth-child(even) {
        background-color: #F5F0EB;
    }
    
    .analytics-block {
        margin-bottom: 25px; 
        background-color: #FFFFFF; 
        padding: 20px; 
        border-radius: 6px; 
        border-left: 5px solid #6C9D9D; 
        font-family: sans-serif; 
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    .analytics-block h4 { color: #4A4A4A; margin-top:0; font-size: 16px; }
    .analytics-block p { margin: 8px 0; font-size: 14px; color: #4A4A4A; }
    </style>
""", unsafe_allow_html=True)

st.title("🏥 Аналитический отчет клиники: загрузка специализаций")
st.write("Загрузите выгрузку из МИС в формате Excel для построения интерактивного отчета.")

uploaded_file = st.file_uploader("Выберите Excel файл (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # =========================================================================
        # 1. ЧТЕНИЕ МЕТАДАННЫХ И ОСНОВНЫХ ДАННЫХ
        # =========================================================================
        df_raw_meta = pd.read_excel(uploaded_file, sheet_name=1, header=None, nrows=3)
        
        start_date_str = str(df_raw_meta.iloc[2, 0]).replace("С:", "").strip() if df_raw_meta.shape[0] > 2 else ""
        end_date_str = str(df_raw_meta.iloc[2, 1]).replace("ПО:", "").strip() if df_raw_meta.shape[1] > 1 else ""
        clinic_name = str(df_raw_meta.iloc[2, 2]).replace("Клиника:", "").strip() if df_raw_meta.shape[1] > 2 else "ООО КЛИНИКА"
        period_str = f"с {start_date_str} по {end_date_str}"

        df_clean = pd.read_excel(uploaded_file, sheet_name=1, skiprows=3)
        df_clean.columns = [str(col).split('~')[0].strip() for col in df_clean.columns]
        
        required_cols = ['Специализация', 'Дата', 'Табель', 'Занято записями', 'Дошло пациентов']
        missing_cols = [c for c in required_cols if c not in df_clean.columns]
        
        if missing_cols:
            st.error(f"❌ В таблице не найдены необходимые колонки: {', '.join(missing_cols)}")
            st.warning(f"Доступные колонки на Листе 2 после очистки: {', '.join(df_clean.columns)}")
            st.stop()
            
        for col in ['Табель', 'Занято записями', 'Дошло пациентов']:
            df_clean[col] = pd.to_numeric(df_clean[col].astype(str).str.replace(",", "."), errors="coerce").fillna(0)

        # =========================================================================
        # 2. АГРЕГАЦИЯ И РАСЧЕТ МЕТРИК
        # =========================================================================
        sp_report = df_clean.groupby('Специализация', as_index=False).agg({
            'Табель': 'sum', 
            'Занято записями': 'sum', 
            'Дошло пациентов': 'sum'
        })
        
        sp_report['Свободно'] = sp_report['Табель'] - sp_report['Занято записями']
        sp_report['Потери'] = sp_report['Занято записями'] - sp_report['Дошло пациентов']
        sp_report['Загрузка %'] = np.where(sp_report['Табель'] > 0, (sp_report['Занято записями'] / sp_report['Табель']) * 100, 0)
        sp_report['Явка %'] = np.where(sp_report['Занято записями'] > 0, (sp_report['Дошло пациентов'] / sp_report['Занято записями']) * 100, 0)
        sp_report['Время без записи в %'] = np.where(sp_report['Табель'] > 0, (sp_report['Свободно'] / sp_report['Табель']) * 100, 0)
        sp_report['Неявки %'] = np.where(sp_report['Занято записями'] > 0, (sp_report['Потери'] / sp_report['Занято записями']) * 100, 0)
        sp_report = sp_report.round(1)

        # =========================================================================
        # 3. ШАПКА И KPI-КАРТОЧКИ
        # =========================================================================
        st.markdown(f"""
            <div class="clinic-header">
                <div class="clinic-title">🏥 КЛИНИКА: {clinic_name}</div>
                <div class="clinic-subtitle">📊 Аналитический отчет: Загруженность медицинских специализаций ({period_str})</div>
            </div>
        """, unsafe_allow_html=True)

        total_tabel = sp_report['Табель'].sum()
        total_active = sp_report['Дошло пациентов'].sum()
        total_free = sp_report['Свободно'].sum()
        total_lost = sp_report['Потери'].sum()
        avg_load = (sp_report['Занято записями'].sum() / total_tabel * 100) if total_tabel > 0 else 0
        avg_show = (total_active / sp_report['Занято записями'].sum() * 100) if sp_report['Занято записями'].sum() > 0 else 0

        kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
        kpi1.metric("📅 Выделено часов", f"{total_tabel:,.1f}")
        kpi2.metric("✅ Фактически занято", f"{total_active:,.1f}")
        kpi3.metric("📉 Свободно часов", f"{total_free:,.1f}")
        kpi4.metric("⚠️ Потери от неявок (часов)", f"{total_lost:,.1f}")
        kpi5.metric("📊 Средняя загрузка", f"{avg_load:.1f}%")
        kpi6.metric("🚶 Средняя явка", f"{avg_show:.1f}%")

        # =========================================================================
        # 4. СВОДНАЯ ТАБЛИЦА (СВОРАЧИВАЕМАЯ)
        # =========================================================================
        table_df = pd.DataFrame()
        table_df['Специализация'] = sp_report['Специализация']
        table_df['Выделено<br>часов'] = sp_report['Табель'].map('{:,.1f}'.format)
        table_df['Записано<br>пациентов<br>(часов)'] = sp_report['Занято записями'].map('{:,.1f}'.format)
        table_df['Фактически<br>занято<br>пациентами<br>(часов)'] = sp_report['Дошло пациентов'].map('{:,.1f}'.format)
        table_df['Время<br>без записи<br>(часов)'] = sp_report['Свободно'].map('{:,.1f}'.format)
        table_df['Время<br>без записи<br>в %'] = sp_report['Время без записи в %'].map('{:,.1f}%'.format)
        table_df['Неявки<br>пациентов<br>(часов)'] = sp_report['Потери'].map('{:,.1f}'.format)
        table_df['Загрузка %'] = sp_report['Загрузка %'].map('{:,.1f}%'.format)
        table_df['Явка %'] = sp_report['Явка %'].map('{:,.1f}%'.format)

        table_html = table_df.to_html(index=False, classes='custom-dash-table', escape=False)

        st.markdown(f"""
        <style>
        .details-table {{
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            margin-bottom: 25px;
            background-color: #FFFFFF;
            border-radius: 6px;
            border-left: 5px solid #B5838D;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            overflow: hidden;
        }}
        .details-table summary {{
            padding: 15px 20px;
            font-size: 16px;
            font-weight: bold;
            color: #4A4A4A;
            cursor: pointer;
            list-style: none;
            display: flex;
            align-items: center;
            gap: 8px;
            user-select: none;
        }}
        .details-table summary::-webkit-details-marker {{ display: none; }}
        .details-table summary::before {{
            content: '▶';
            font-size: 12px;
            color: #B5838D;
            transition: transform 0.2s;
            display: inline-block;
            width: 16px;
        }}
        .details-table[open] summary::before {{
            content: '▼';
        }}
        .details-table .table-wrapper {{
        padding: 0 20px 20px 20px;
        }}
</style>

<details class="details-table" open>
    <summary>📋 Сводная таблица эффективности по специализациям</summary>
    <div class="table-wrapper">
        {table_html}
    </div>
</details>
""", unsafe_allow_html=True)

        # =========================================================================
        # 5. АНАЛИТИЧЕСКИЕ БЛОКИ
        # =========================================================================
        #t10_sp = sp_report.sort_values('Табель', ascending=False)['Специализация'].head(10).tolist()
        #anti_load = sp_report.sort_values('Загрузка %', ascending=True).head(3)[['Специализация', 'Загрузка %']].values.tolist()
        #anti_show = sp_report.sort_values('Явка %', ascending=True).head(3)[['Специализация', 'Явка %']].values.tolist()
        #m_loss = sp_report.sort_values('Потери', ascending=False).head(3)[['Специализация', 'Потери']].values.tolist()
        #m_free = sp_report.sort_values('Свободно', ascending=False).head(3)[['Специализация', 'Свободно']].values.tolist()

        #st.markdown(f"""
        #<div class='analytics-block'>
        #    <h4>📊 АНАЛИТИЧЕСКИЕ БЛОКИ И РЕЙТИНГИ КЛИНИКИ</h4>
        #    <p><b>🩺 ТОП-10 специализаций по Выделено часов:</b> {", ".join(t10_sp)}</p>
        #    <p><b>📉 Anti-load (Низкая загрузка):</b> {", ".join([f"{name} ({val:.1f}%)" for name, val in anti_load])}</p>
        #    <p><b>🚶‍♂️ Anti-show (Низкая явка пациентов):</b> {", ".join([f"{name} ({val:.1f}%)" for name, val in anti_show])}</p>
        #    <p><b>⚠️ Максимальные Потери пациентов (часов):</b> {", ".join([f"{name} ({val:.1f}ч)" for name, val in m_loss])}</p>
        #    <p><b>📅 Максимальное Свободное время (часов):</b> {", ".join([f"{name} ({val:.1f}ч)" for name, val in m_free])}</p>
        #</div>
        #""", unsafe_allow_html=True)

        # =========================================================================
        # 6. ГРАФИК 1: ЛИНЕЙНЫЙ
        # =========================================================================
        st.subheader("1. Линейный график: Динамика использования рабочего времени врачей по дням")
        
        df_clean["Parsed_Date_All"] = pd.to_datetime(df_clean["Дата"], dayfirst=True, errors="coerce")
        df_daily = df_clean.dropna(subset=["Parsed_Date_All"]).groupby('Parsed_Date_All', as_index=False)[['Табель', 'Занято записями', 'Дошло пациентов']].sum()
        df_daily = df_daily.sort_values("Parsed_Date_All")
        p1 = go.Figure()
        p1.add_trace(go.Scatter(
            x=df_daily['Parsed_Date_All'], y=df_daily['Табель'], 
            name='Выделено часов', line=dict(color='#005F73', width=3.5)
        ))
        p1.add_trace(go.Scatter(
            x=df_daily['Parsed_Date_All'], y=df_daily['Занято записями'], 
            name='Записано пациентов (часов)', line=dict(color='#D6A4BB', width=3)
        ))
        p1.add_trace(go.Scatter(
            x=df_daily['Parsed_Date_All'], y=df_daily['Дошло пациентов'], 
            name='Фактически занято (часов)', line=dict(color='#6C9D9D', width=3)
        ))
        p1.update_layout(
            template="plotly_white", 
            xaxis=dict(
                tickmode='array',
                tickvals=df_daily['Parsed_Date_All'],   # ← засечка на КАЖДУЮ дату
                tickformat="%d.%m.%Y", 
                tickangle=-45,
                automargin=True
            ),
            yaxis_title="Часы",
            legend_title="Показатели",
            height=500
        )
        st.plotly_chart(p1, use_container_width=True)

        # =========================================================================
        # 7. ГРАФИК 2: СОСТАВНАЯ ДИАГРАММА
        # =========================================================================
        st.subheader("2. Анализ нагрузки и невостребованного времени по специализациям")
        
        df_p2 = sp_report.sort_values('Табель', ascending=False).copy()
        df_p2['Неявки пациентов'] = df_p2['Потери']
        df_p2['Свободные часы'] = df_p2['Табель'] - df_p2['Занято записями']
        
        p2 = go.Figure()
        p2.add_trace(go.Bar(x=df_p2['Специализация'], y=df_p2['Дошло пациентов'], name='Отработано (Дошли)', marker_color='#6C9D9D', offsetgroup=0,
                           hovertemplate="<b>Отработано (Дошли)</b><br>Специализация: %{x}<br>Время: %{y:,.1f} ч<extra></extra>"))
        p2.add_trace(go.Bar(x=df_p2['Специализация'], y=df_p2['Неявки пациентов'], name='Потери (Неявки)', marker_color='#6A323A', offsetgroup=0,
                           base=df_p2['Дошло пациентов'], customdata=np.stack([df_p2['Неявки пациентов']], axis=-1),
                           hovertemplate="<b>Потери (Неявки)</b><br>Специализация: %{x}<br>Время: %{customdata[0]:,.1f} ч<extra></extra>"))
        p2.add_trace(go.Bar(x=df_p2['Специализация'], y=df_p2['Свободные часы'], name='Незанятое время', marker_color='#e8d3dd', offsetgroup=0,
                           base=df_p2['Дошло пациентов'] + df_p2['Неявки пациентов'], customdata=np.stack([df_p2['Свободные часы']], axis=-1),
                           hovertemplate="<b>Незанятое время</b><br>Специализация: %{x}<br>Время: %{customdata[0]:,.1f} ч<extra></extra>"))
        p2.add_trace(go.Bar(x=df_p2['Специализация'], y=df_p2['Табель'], name='Всего выделено часов', marker_color='#005F73', offsetgroup=1,
                           hovertemplate="<b>Всего выделено часов</b><br>Специализация: %{x}<br>Всего часов: %{y:,.1f} ч<extra></extra>"))
        p2.update_layout(barmode='group', template="plotly_white", xaxis={'categoryorder':'total descending'}, yaxis_title="Часы", height=600)
        st.plotly_chart(p2, use_container_width=True)

        # =========================================================================
        # 8. ГРАФИК 3: ТОП ПО ВЫДЕЛЕННЫМ ЧАСАМ
        # =========================================================================
        st.subheader("3. ТОП специализаций по выделенному времени")
        
        df_p3 = sp_report.sort_values('Табель', ascending=True).copy()
        p3 = px.bar(df_p3, x='Табель', y='Специализация', orientation='h', title="Выделено рабочих часов по табелю",
                    color='Табель', color_continuous_scale=[[0.0, '#e6fcfb'], [1.0, '#005F73']])
        p3.update_layout(xaxis_title="Всего часов (ч.)", height=600)
        p3.update_traces(hovertemplate="<b>%{y}</b><br>Выделено: %{x:,.1f} ч.<extra></extra>")
        st.plotly_chart(p3, use_container_width=True)

        # =========================================================================
        # 9. ГРАФИКИ 4.1 и 4.2 — РЯДОМ В ДВУХ КОЛОНКАХ (как 5.1 и 5.2)
        # =========================================================================
        st.subheader("4. Анализ загрузки и недозагрузки специализаций")
        col1, col2 = st.columns(2)

        with col1:
            df_p4 = sp_report.sort_values('Загрузка %', ascending=True).copy()
            p4 = px.bar(
                df_p4, x='Загрузка %', y='Специализация', orientation='h',
                title="4.1. ТОП специализаций по Загрузке %",
                color='Загрузка %',
                color_continuous_scale=[[0.0, '#fce3ef'], [1.0, '#6A323A']],
                custom_data=['Табель', 'Занято записями']
            )
            p4.update_layout(
                xaxis_title="Загрузка расписания (%)",
                coloraxis_colorbar=dict(title="Загрузка %")
            )
            p4.update_traces(
                hovertemplate="<b>%{y}</b><br>Загрузка: %{x:.1f}%<br>Выделено часов: %{customdata[0]:,.1f} ч.<br>Занято записью: %{customdata[1]:,.1f} ч.<extra></extra>"
            )
            st.plotly_chart(p4, use_container_width=True)

        with col2:
            df_p42 = sp_report.copy()
            df_p42['Незанято часов всего'] = df_p42['Свободно']
            df_p42['Свободное время %'] = np.where(df_p42['Табель'] > 0, (df_p42['Незанято часов всего'] / df_p42['Табель']) * 100, 0)
            df_p42 = df_p42.sort_values('Свободное время %', ascending=True)

            p42 = px.bar(
                df_p42, x='Свободное время %', y='Специализация', orientation='h',
                title="4.2. ТОП специализаций по НЕДОзагрузке %",
                color='Свободное время %',
                color_continuous_scale=[[0.0, '#e6fcfb'], [1.0, '#005F73']],
                custom_data=['Табель', 'Незанято часов всего']
            )
            p42.update_layout(
                xaxis_title="Процент незагруженного времени (%)",
                coloraxis_colorbar=dict(title="% недозагрузки")
            )
            p42.update_traces(
                hovertemplate="<b>%{y}</b><br>Свободное время: %{x:.1f}%<br>Выделено часов всего: %{customdata[0]:,.1f} ч.<br>Незанято часов всего: %{customdata[1]:,.1f} ч.<extra></extra>"
            )
            st.plotly_chart(p42, use_container_width=True)

        # =========================================================================
        # 10. ГРАФИКИ 5.1 и 5.2 — РЯДОМ В ДВУХ КОЛОНКАХ (исправлены цвета)
        # =========================================================================
        st.subheader("5. Анализ неявок пациентов")
        col1, col2 = st.columns(2)
        
        with col1:
            df_p51 = sp_report.sort_values('Потери', ascending=True).copy()
            p51 = px.bar(
                df_p51, x='Потери', y='Специализация', orientation='h',
                title="5.1. ТОП по Неявкам (часов)",
                color='Потери',
                color_continuous_scale=[[0.0, '#fce3ef'], [1.0, '#6A323A']]
            )
            p51.update_layout(
                xaxis_title="Неявки пациентов (часов)",
                coloraxis_colorbar=dict(title="Потери")
            )
            p51.update_traces(hovertemplate="<b>%{y}</b><br>Неявки: %{x:,.1f} ч.<extra></extra>")
            st.plotly_chart(p51, use_container_width=True)
            
        with col2:
            df_p52 = sp_report.sort_values('Неявки %', ascending=True).copy()
            p52 = px.bar(
                df_p52, x='Неявки %', y='Специализация', orientation='h',
                title="5.2. ТОП по Неявкам (%)",
                color='Неявки %',
                color_continuous_scale=[[0.0, '#e6fcfb'], [1.0, '#005F73']]
            )
            p52.update_layout(
                xaxis_title="Неявки пациентов (%)",
                coloraxis_colorbar=dict(title="Неявки %")
            )
            p52.update_traces(hovertemplate="<b>%{y}</b><br>Процент неявок: %{x:.1f}%<extra></extra>")
            st.plotly_chart(p52, use_container_width=True)

        # =========================================================================
        # 11. ГРАФИК 6: TREEMAP
        # =========================================================================
        st.subheader("6. Плиточная диаграмма: Объем выделенного времени и процент явки пациентов")
        
        p6 = px.treemap(
            sp_report, path=['Специализация'], values='Табель', color='Явка %',
            color_continuous_scale=[[0.0, '#e6fcfb'], [1.0, '#005F73']],
            custom_data=['Загрузка %', 'Явка %']
        )
        p6.update_traces(
            marker=dict(line=dict(width=2, color='#FFFFFF')),
            texttemplate="<b>%{label}</b><br>Выделено часов: %{value:,.1f} ч.<br>Загрузка: %{customdata[0]:.1f}%<br>Явка: %{customdata[1]:.1f}%",
            textfont=dict(size=11, color='#2B2D42'),
            hovertemplate="<b>%{label}</b><br>Выделено часов: %{value:,.1f} ч.<br>Загрузка: %{customdata[0]:.1f}%<br>Явка: %{customdata[1]:.1f}%<extra></extra>"
        )
        p6.update_layout(coloraxis=dict(cmin=sp_report['Явка %'].min(), cmax=100), margin=dict(t=60, b=20, l=20, r=20), height=600)
        st.plotly_chart(p6, use_container_width=True)

        # =========================================================================
        # 12. ГРАФИК 7: МАТРИЦА ЭФФЕКТИВНОСТИ
        # =========================================================================
        st.subheader("7. Матрица эффективности: Анализ загрузки и неявок по направлениям")
        
        p7 = px.scatter(
            sp_report, x='Табель', y='Загрузка %', size='Табель', color='Неявки %',
            hover_name='Специализация', text='Специализация',
            title="Анализ загрузки и неявок по направлениям",
            color_continuous_scale=[[0.0, '#00A896'], [0.5, '#F4A261'], [1.0, '#D62828']],
            size_max=65
        )
        p7.update_traces(
            textposition='top center', mode='markers+text',
            marker=dict(sizeref=12, sizemode='diameter', opacity=0.6, line=dict(width=1, color='#2B2D42')),
            hovertemplate="<b>%{hovertext}</b><br>Выделено часов: %{x:,.1f} ч.<br>Загрузка расписания: %{y:.1f}%<br>Неявки пациентов: %{marker.color:.1f}%<extra></extra>"
        )
        p7.update_layout(template="plotly_white", xaxis_title="Выделено часов (ч.)", yaxis_title="Загрузка расписания (%)", coloraxis_colorbar=dict(title="Неявки %"), height=650)
        st.plotly_chart(p7, use_container_width=True)

        # =========================================================================
        # 13. ГРАФИК 8: КАСКАДНАЯ ДИАГРАММА
        # =========================================================================
        st.subheader("8. Каскадная диаграмма: Баланс рабочего времени и структура операционных потерь")
        
        t_h = float(sp_report['Табель'].sum())
        free_h = float(sp_report['Свободно'].sum())
        lost_h = float(sp_report['Потери'].sum())
        active_h = float(sp_report['Дошло пациентов'].sum())
        
        pct_free = (free_h / t_h * 100)
        pct_lost = (lost_h / t_h * 100)
        pct_active = (active_h / t_h * 100)
        
        x_labels = ["1. Выделено часов<br>(План по табелю)", "Время без записи<br>(Свободные окна)", "Неявки пациентов<br>(Сорванные приемы)", "Фактически занято<br>пациентами (Факт)"]
        base = [0, t_h - free_h, t_h - free_h - lost_h, 0]
        y_values = [t_h, free_h, lost_h, active_h]
        text_labels = [f"{t_h:,.1f} ч. (100%)", f"-{free_h:,.1f} ч. (-{pct_free:.1f}%)", f"-{lost_h:,.1f} ч. (-{pct_lost:.1f}%)", f"{active_h:,.1f} ч. ({pct_active:.1f}%)"]
        custom_colors = ['#005F73', '#B5838D', '#B5838D', '#6C9D9D']
        
        p8 = go.Figure(go.Bar(x=x_labels, y=y_values, base=base, text=text_labels, textposition='outside',
                               marker_color=custom_colors, textfont=dict(size=12, color='#2B2D42'),
                               hovertemplate="<b>%{x}</b><br>Баланс: %{text}<extra></extra>"))
        p8.update_layout(title="Баланс рабочего времени и структура операционных потерь", yaxis_title="Количество часов",
                         template="plotly_white", yaxis=dict(range=[0, t_h * 1.15]), height=550)
        st.plotly_chart(p8, use_container_width=True)

        # =========================================================================
        # 14. ГРАФИК 9: КОЛЬЦЕВАЯ ДИАГРАММА
        # =========================================================================
        st.subheader("9. Кольцевая диаграмма: Структура использования рабочего времени докторов")
        
        labels_list = ['Фактически занято', 'Время без записи', 'Неявки пациентов']
        values_list = [active_h, free_h, lost_h]
        full_text_labels = [f"Фактически занято<br>пациентами ({pct_active:.1f}%)", f"Время без записи<br>(Свободные окна) ({pct_free:.1f}%)", f"Неявки пациентов<br>(Простои) ({pct_lost:.1f}%)"]
        donut_colors = ['#6C9D9D', '#d1fff4', '#B5838D']
        
        p9 = go.Figure(data=[go.Pie(labels=labels_list, values=values_list, hole=.4,
                                     marker=dict(colors=donut_colors, line=dict(color='#FFFFFF', width=2)),
                                     text=full_text_labels, textinfo='text', textposition='outside',
                                     textfont=dict(size=11, color='#2B2D42'))])
        p9.update_layout(template="plotly_white", height=500, showlegend=True,
                         legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05))
        p9.update_traces(hovertemplate="<b>%{label}</b><br>Объем: %{value:,.1f} ч.<br>Доля: %{percent}<extra></extra>")
        st.plotly_chart(p9, use_container_width=True)

        # =========================================================================
        # 15. ГРАФИКИ 10 И 11: ТЕПЛОВЫЕ КАРТЫ (14 ДНЕЙ)
        # =========================================================================
        st.subheader("📅 Тепловые карты расписания (последние 14 дней)")
        
        df_local = df_clean.copy()
        df_local["Parsed_Date"] = pd.to_datetime(df_local["Дата"], dayfirst=True, errors="coerce")
        df_local = df_local.dropna(subset=["Parsed_Date"])
        
        if not df_local.empty:
            last_date = df_local["Parsed_Date"].max()
            start_date = last_date - timedelta(days=13)
            df_local = df_local[(df_local["Parsed_Date"] >= start_date) & (df_local["Parsed_Date"] <= last_date)].copy()
            df_local["Дата"] = df_local["Parsed_Date"].dt.strftime("%d.%m.%Y")
            ordered_dates = [d.strftime("%d.%m.%Y") for d in pd.date_range(start=start_date, end=last_date, freq="D")]
            
            agg = df_local.groupby(["Дата", "Специализация"], as_index=False).agg({
                "Табель": "sum", "Занято записями": "sum", "Дошло пациентов": "sum"
            })
            
            # --- 10. ЗАПОЛНЕННОСТЬ ---
            agg["Загрузка %"] = np.where(agg["Табель"] > 0, (agg["Занято записями"] / agg["Табель"]) * 100, np.nan)
            h10 = agg.pivot(index="Дата", columns="Специализация", values="Загрузка %").reindex(ordered_dates)
            t10 = agg.pivot(index="Дата", columns="Специализация", values="Табель").reindex(ordered_dates)
            z10 = agg.pivot(index="Дата", columns="Специализация", values="Занято записями").reindex(ordered_dates)
            
            hover_text_10, text_matrix_10 = [], []
            for date in h10.index:
                row_hover, row_text = [], []
                for spec in h10.columns:
                    val, t_val, z_val = h10.loc[date, spec], t10.loc[date, spec], z10.loc[date, spec]
                    if pd.isna(t_val) or t_val == 0:
                        row_hover.append(f"Дата: {date}<br>Специализация: {spec}<br><b>Нет приема</b>")
                        row_text.append("Нет приема")
                    else:
                        row_hover.append(f"Дата: {date}<br>Специализация: {spec}<br>Загрузка: {val:.1f}%<br>Табель: {t_val:.1f} ч.<br>Записано: {z_val:.1f} ч.")
                        row_text.append("")
                hover_text_10.append(row_hover)
                text_matrix_10.append(row_text)
            
            p10 = go.Figure(data=go.Heatmap(
                z=h10.fillna(-1).values, x=h10.columns, y=h10.index,
                text=text_matrix_10, hovertext=hover_text_10, hoverinfo="text", texttemplate="%{text}",
                colorscale=[[0.0, '#E0E0E0'], [0.009, '#E0E0E0'], [0.01, '#F0F8FF'], [0.4, '#BDE0FE'], [0.8, '#4EA8DE'], [1.0, '#003049']],
                zmin=-1, zmax=100, xgap=1, ygap=1
            ))
            p10.update_layout(title="10. Заполненность расписания (Записано / Табель)", height=650, template="plotly_white",
                              xaxis=dict(tickangle=-45), yaxis=dict(type="category"))
            st.plotly_chart(p10, use_container_width=True)
            
            # --- 11. ЯВКА ---
            agg["Явка %"] = np.where(agg["Занято записями"] > 0, (agg["Дошло пациентов"] / agg["Занято записями"]) * 100, np.nan)
            h11 = agg.pivot(index="Дата", columns="Специализация", values="Явка %").reindex(ordered_dates)
            z11 = agg.pivot(index="Дата", columns="Специализация", values="Занято записями").reindex(ordered_dates)
            d11 = agg.pivot(index="Дата", columns="Специализация", values="Дошло пациентов").reindex(ordered_dates)
            
            hover_text_11, text_matrix_11 = [], []
            for date in h11.index:
                row_hover, row_text = [], []
                for spec in h11.columns:
                    val, z_val, d_val = h11.loc[date, spec], z11.loc[date, spec], d11.loc[date, spec]
                    if pd.isna(z_val) or z_val == 0:
                        row_hover.append(f"Дата: {date}<br>Специализация: {spec}<br><b>Нет приема</b>")
                        row_text.append("Нет приема")
                    else:
                        row_hover.append(f"Дата: {date}<br>Специализация: {spec}<br>Явка: {val:.1f}%<br>Записано: {z_val:.1f} ч.<br>Дошло: {d_val:.1f} ч.")
                        row_text.append("")
                hover_text_11.append(row_hover)
                text_matrix_11.append(row_text)
            
            p11 = go.Figure(data=go.Heatmap(
                z=h11.fillna(-1).values, x=h11.columns, y=h11.index,
                text=text_matrix_11, hovertext=hover_text_11, hoverinfo="text", texttemplate="%{text}",
                colorscale=[[0.0, '#E0E0E0'], [0.009, '#E0E0E0'], [0.01, '#FFF0F3'], [0.4, '#FFB3C1'], [0.8, '#C9184A'], [1.0, '#4F000B']],
                zmin=-1, zmax=100, xgap=1, ygap=1
            ))
            p11.update_layout(title="11. Процент явки пациентов (Дошло / Записано)", height=650, template="plotly_white",
                              xaxis=dict(tickangle=-45), yaxis=dict(type="category"))
            st.plotly_chart(p11, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Ошибка при обработке файла: {e}")
        st.exception(e)
