import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta

# Настройка страницы дашборда
st.set_page_config(page_title="Аналитический отчет клиники", layout="wide")

# Кастомные стили
st.markdown("""
    <style>
    .clinic-header { background-color: #f8f9fa; border-left: 5px solid #005F73; padding: 15px; margin-bottom: 25px; }
    .clinic-title { font-size: 24px; font-weight: bold; color: #2B2D42; }
    .clinic-subtitle { font-size: 14px; color: #6C9D9D; }
    </style>
""", unsafe_allow_html=True)

st.title("🏥 Аналитическая панель клиники")
st.write("Загрузите выгрузку из МИС в формате Excel для построения интерактивного отчета.")

uploaded_file = st.file_uploader("Выберите Excel файл (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 1. Читаем метаданные со второго листа (sheet_name=1), первые 3 строки
        df_raw_meta = pd.read_excel(uploaded_file, sheet_name=1, header=None, nrows=3)
        
        # Извлекаем текст из третьей строки
        start_date_str = str(df_raw_meta.iloc[2, 0]).replace("С:", "").strip()
        end_date_str = str(df_raw_meta.iloc[2, 1]).replace("ПО:", "").strip()
        clinic_name = str(df_raw_meta.iloc[2, 2]).replace("Клиника:", "").strip()
        period_str = f"с {start_date_str} по {end_date_str}"

        # 2. Читаем основные данные со второго листа, пропуская первые 3 строки-шапки
        df_clean = pd.read_excel(uploaded_file, sheet_name=1, skiprows=3)
        
        # Очищаем названия колонок от технического хвоста "~000"
        df_clean.columns = [str(col).split('~')[0].strip() for col in df_clean.columns]
        
        # Проверяем наличие всех ключевых колонок
        required_cols = ['Специализация', 'Дата', 'Табель', 'Занято записями', 'Дошло пациентов']
        missing_cols = [c for c in required_cols if c not in df_clean.columns]
        
        if missing_cols:
            st.error(f"❌ В таблице не найдены необходимые колонки: {', '.join(missing_cols)}")
            st.warning(f"Доступные колонки на Листе 2 после очистки: {', '.join(df_clean.columns)}")
            st.stop()
            
        # Приводим числовые колонки к правильному типу данных
        for col in ['Табель', 'Занято записями', 'Дошло пациентов']:
            df_clean[col] = pd.to_numeric(df_clean[col].astype(str).str.replace(",", "."), errors="coerce").fillna(0)

        # 3. Агрегация данных по специализациям
        sp_report = df_clean.groupby('Специализация', as_index=False).agg({
            'Табель': 'sum', 'Занято записями': 'sum', 'Дошло пациентов': 'sum'
        })
        
        # Расчет метрик эффективности
        sp_report['Свободно'] = sp_report['Табель'] - sp_report['Занято записями']
        sp_report['Потери'] = sp_report['Занято записями'] - sp_report['Дошло пациентов']
        sp_report['Загрузка %'] = np.where(sp_report['Табель'] > 0, (sp_report['Занято записями'] / sp_report['Табель']) * 100, 0)
        sp_report['Явка %'] = np.where(sp_report['Занято записями'] > 0, (sp_report['Дошло пациентов'] / sp_report['Занято записями']) * 100, 0)
        sp_report = sp_report.round(1)

        # Вывод красивой шапки клиники
        st.markdown(f"""
            <div class="clinic-header">
                <div class="clinic-title">🏥 КЛИНИКА: {clinic_name}</div>
                <div class="clinic-subtitle">📊 Аналитический отчет: Загруженность медицинских специализаций ({period_str})</div>
            </div>
        """, unsafe_allow_html=True)

        st.subheader("📋 Сводная таблица эффективности")
        st.dataframe(sp_report, use_container_width=True)

        # График 1: Динамика по дням
        st.subheader("1. Линейный график: Динамика использования времени")
        df_clean["Parsed_Date_All"] = pd.to_datetime(df_clean["Дата"], dayfirst=True, errors="coerce")
        df_daily = df_clean.dropna(subset=["Parsed_Date_All"]).groupby('Дата').sum().reset_index()
        p1 = px.line(df_daily, x='Дата', y='Занято записями', title="Динамика записей по дням")
        st.plotly_chart(p1, use_container_width=True)

        # График 4: ТОП по загрузке
        st.subheader("4. Горизонтальный Bar Chart (ТОП по Загрузке)")
        df_p4 = sp_report.sort_values('Загрузка %', ascending=True).copy()
        p4 = px.bar(
            df_p4, x='Загрузка %', y='Специализация', orientation='h',
            title="ТОП специализаций по Загрузке %", color='Загрузка %',
            color_continuous_scale=[[0.0, '#fce3ef'], [1.0, '#6A323A']],
            custom_data=['Табель', 'Занято записями']
        )
        p4.update_traces(hover_template="<b>%{y}</b><br>Загрузка: %{x:.1f}%<br>Выделено часов: %{customdata[0]:.1f} ч.<br>Занято записью: %{customdata[1]:.1f} ч.<extra></extra>")
        st.plotly_chart(p4, use_container_width=True)

        # График 5: Анализ неявок
        st.subheader("5. Анализ неявок пациентов")
        col1, col2 = st.columns(2)
        with col1:
            df_p51 = sp_report.sort_values('Потери', ascending=True).copy()
            p51 = px.bar(df_p51, x='Потери', y='Специализация', orientation='h', title="5.1. ТОП по Неявкам (часов)", color='Потери', color_continuous_scale=[[0.0, '#e6fcfb'], [1.0, '#005F73']])
            st.plotly_chart(p51, use_container_width=True)
        with col2:
            sp_report['Неявки %'] = (sp_report['Потери'] / sp_report['Занято записями'].clip(lower=1) * 100).round(1)
            df_p52 = sp_report.sort_values('Неявки %', ascending=True).copy()
            p52 = px.bar(df_p52, x='Неявки %', y='Специализация', orientation='h', title="5.2. ТОП по Неявкам (%)", color='Неявки %', color_continuous_scale=[[0.0, '#fce3ef'], [1.0, '#6A323A']])
            st.plotly_chart(p52, use_container_width=True)

        # График 6: Плиточная диаграмма (Treemap)
        st.subheader("6. Плиточная диаграмма: Объемы и явка")
        p6 = px.treemap(
            sp_report, path=['Специализация'], values='Табель', color='Явка %',
            color_continuous_scale=[[0.0, '#005F73'], [1.0, '#E0FFFF']],
            title="Объем выделенного времени и процент явки пациентов",
            custom_data=['Загрузка %', 'Явка %']
        )
        p6.update_traces(texttemplate="<b>%{label}</b><br>Выделено часов: %{value:.1f} ч.<br>Загрузка: %{customdata[0]:.1f}%<br>Явка: %{customdata[1]:.1f}%")
        st.plotly_chart(p6, use_container_width=True)

        # График 7: Матрица эффективности
        st.subheader("7. Матрица эффективности: Загрузка и неявки")
        p7 = px.scatter(
            sp_report, x='Табель', y='Загрузка %', size='Табель', color='Неявки %',
            hover_name='Специализация', text='Специализация', title="Анализ загрузки и неявок по направлениям",
            color_continuous_scale=[[0.0, '#00A896'], [0.5, '#F4A261'], [1.0, '#D62828']]
        )
        st.plotly_chart(p7, use_container_width=True)

        # График 8: Каскадная диаграмма (Waterfall Balance)
        st.subheader("8. Каскадная диаграмма: Баланс рабочего времени")
        t_h, f_h, l_h, a_h = sp_report['Табель'].sum(), sp_report['Свободно'].sum(), sp_report['Потери'].sum(), sp_report['Дошло пациентов'].sum()
        p8 = go.Figure(go.Bar(
            x=["1. План по табелю", "Время без записи", "Неявки пациентов", "Фактически занято"], y=[t_h, f_h, l_h, a_h],
            base=[0, t_h - f_h, t_h - f_h - l_h, 0], marker_color=['#005F73', '#B5838D', '#B5838D', '#6C9D9D']
        ))
        st.plotly_chart(p8, use_container_width=True)

        # График 9: Кольцевая структура (Donut)
        st.subheader("9. Структура использования времени")
        p9 = go.Figure(data=[go.Pie(
            labels=['Фактически занято', 'Время без записи', 'Неявки пациентов'], values=[a_h, f_h, l_h],
            hole=.4, marker=dict(colors=['#6C9D9D', '#d1fff4', '#B5838D'])
        )])
        st.plotly_chart(p9, use_container_width=True)

        # --- ТЕПЛОВЫЕ КАРТЫ за последние 14 дней (Графики 10 и 11) ---
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
            
            # Тепловая карта 10
            agg["Загрузка %"] = np.where(agg["Табель"] > 0, (agg["Занято записями"] / agg["Табель"]) * 100, np.nan)
            h10 = agg.pivot(index="Дата", columns="Специализация", values="Загрузка %").reindex(ordered_dates)
            p10 = go.Figure(data=go.Heatmap(
                z=h10.fillna(-1).values, x=h10.columns, y=h10.index,
                colorscale=[[0.0, '#E0E0E0'], [0.009, '#E0E0E0'], [0.01, '#005F73'], [1.0, '#d1fff4']], zmin=-1, zmax=100
            ))
            p10.update_layout(title="10. Заполненность расписания (Записано / Табель)", height=650)
            st.plotly_chart(p10, use_container_width=True)
            
            # Тепловая карта 11
            agg["Явка %"] = np.where(agg["Занято записями"] > 0, (agg["Дошло пациентов"] / agg["Занято записями"]) * 100, np.nan)
            h11 = agg.pivot(index="Дата", columns="Специализация", values="Явка %").reindex(ordered_dates)
            p11 = go.Figure(data=go.Heatmap(
                z=h11.fillna(-1).values, x=h11.columns, y=h11.index,
                colorscale=[[0.0, '#E0E0E0'], [0.009, '#E0E0E0'], [0.01, '#F3E6E8'], [1.0, '#5A1A2A']], zmin=-1, zmax=100
            ))
            p11.update_layout(title="11. Процент явки пациентов (Дошло / Записано)", height=650)
            st.plotly_chart(p11, use_container_width=True)

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")

