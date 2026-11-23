import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
import io
import re

# Настройка страницы
st.set_page_config(
    page_title="Поурочный план КР",
    page_icon="📚",
    layout="wide"
)

API_KEY = st.secrets.get("GEMINI_API_KEY", "")

if API_KEY:
    genai.configure(api_key=API_KEY)

if "last_result" not in st.session_state:
    st.session_state.last_result = None
if "last_title" not in st.session_state:
    st.session_state.last_title = "Поурочный_план"

# Функция превращения текста и Markdown-таблиц в настоящий файл Word (.docx)
def create_docx(markdown_text):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    lines = markdown_text.split('\n')
    table_buffer = []

    def flush_table(buf):
        if not buf:
            return
        clean_rows = [r for r in buf if not re.match(r'^\|[\s\-:|]+\|$', r.strip())]
        if not clean_rows:
            return
        
        parsed_rows = []
        for r in clean_rows:
            cells = [c.strip() for c in r.strip().strip('|').split('|')]
            parsed_rows.append(cells)
            
        cols_count = max(len(r) for r in parsed_rows)
        table = doc.add_table(rows=len(parsed_rows), cols=cols_count)
        table.style = 'Table Grid'
        
        for row_idx, row_data in enumerate(parsed_rows):
            for col_idx, val in enumerate(row_data):
                if col_idx < cols_count:
                    cell = table.cell(row_idx, col_idx)
                    cell.text = val.replace('**', '')
                    if row_idx == 0:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.bold = True

    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            table_buffer.append(stripped)
        else:
            if table_buffer:
                flush_table(table_buffer)
                table_buffer = []
            
            if not stripped:
                continue
            elif stripped.startswith('### '):
                doc.add_heading(stripped.replace('### ', '').replace('**', ''), level=2)
            elif stripped.startswith('## '):
                doc.add_heading(stripped.replace('## ', '').replace('**', ''), level=1)
            else:
                p = doc.add_paragraph()
                parts = re.split(r'(\*\*.*?\*\*)', stripped)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        run = p.add_run(part[2:-2])
                        run.bold = True
                    else:
                        p.add_run(part)

    if table_buffer:
        flush_table(table_buffer)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Заголовок приложения
st.title("📚 Генератор поурочных планов по стандартам КР")
st.markdown("Введите данные ниже, чтобы получить качественный поурочный план и скачать его готовым файлом для Word.")

st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    subject = st.text_input("Предмет:", placeholder="Например: История Кыргызстана / Русский язык")
    grade = st.selectbox("Класс:", ["5 класс", "6 класс", "7 класс", "8 класс", "9 класс", "10 класс", "11 класс"])
with col2:
    language = st.selectbox("Язык обучения:", ["Русский", "Кыргызский"])
    duration = st.selectbox("Длительность урока:", ["45 минут", "90 минут (пара)"])

lesson_topic = st.text_input("Тема урока:", placeholder="Например: Древний Кыргызстан. Первые стоянки человека")

st.markdown("---")

def generate_lesson_plan(prompt_text):
    if not API_KEY:
        return "⚠️ Ошибка: API ключ не найден в конфигурации secrets Streamlit."
        
    # Прописываем линейку моделей от 2.5 до 3.8
    models_to_try = [
        "gemini-3.8-flash", "gemini-3.8-pro",
        "gemini-3.5-flash", "gemini-3.5-pro",
        "gemini-3.0-flash", "gemini-3.0-pro",
        "gemini-2.8-flash", "gemini-2.8-pro",
        "gemini-2.5-flash", "gemini-2.5-pro",
        "gemini-2.0-flash", "gemini-2.0-pro"
    ]
    
    # Автоматически подтягиваем всё, что есть на ключе
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                short_name = m.name.replace("models/", "")
                if "gemini" in short_name and short_name not in models_to_try:
                    models_to_try.append(short_name)
    except Exception:
        pass
    
    last_err = ""
    for m_name in models_to_try:
        try:
            model = genai.GenerativeModel(m_name)
            response = model.generate_content(prompt_text)
            if response and response.text:
                return response.text
        except Exception as e:
            last_err = str(e)
            continue
            
    return f"⚠️ Ошибка генерации: {last_err}"

if st.button("Сгенерировать поурочный план", type="primary", use_container_width=True):
    if not subject or not lesson_topic:
        st.warning("⚠️ Пожалуйста, заполните предмет и тему урока.")
    else:
        with st.spinner("Создаю профессиональный поурочный план..."):
            prompt = (
                f"Ты — опытный методист-педагог в Кыргызской Республике. Составь максимально подробный, профессиональный, развернутый поурочный план "
                f"по предмету '{subject}' для {grade} (обучение на {language} языке). Длительность: {duration}. Тема: '{lesson_topic}'. "
                f"Строго учти государственные стандарты образования Кыргызской Республики.\n\n"
                f"СТРОГАЯ СТРУКТУРА ДОКУМЕНТА, КОТОРУЮ ТЫ ОБЯЗАН СОБЛЮСТЬ:\n"
                f"1. Заголовок: ### ПОУРОЧНЫЙ ПЛАН: {lesson_topic}\n"
                f"2. Блок «1. Цели урока:» (детально распиши: а) Обучающая, б) Развивающая, в) Воспитательная).\n"
                f"3. Блок «2. Ключевые и предметные компетенции (по стандарту КР):» (обязательно укажи ключевые компетенции КК-1, КК-2, КК-3 и предметные ПК-1, ПК-2 с их подробной расшифровкой под тему урока).\n"
                f"4. Блок «3. Оборудование и наглядные пособия:» (учебник, раздаточный материал, интерактивная доска, карты и т.д.).\n"
                f"5. Блок «4. Ход урока:» — оформи ИСКЛЮЧИТЕЛЬНО в виде большой, детализированной Markdown-таблицы со следующими колонками:\n"
                f"| Этап урока и время | Деятельность учителя | Деятельность ученика | Оценивание / Методические указания |\n\n"
                f"Пиши ОЧЕНЬ развернуто, глубоко и профессионально на {language} языке. Не сокращай этапы урока!"
            )
            
            res = generate_lesson_plan(prompt)
            if res and not res.startswith("⚠️"):
                st.session_state.last_result = res
                clean_name = re.sub(r'[^\w\s-]', '', f"{subject}_{grade}_{lesson_topic}").strip().replace(' ', '_')
                st.session_state.last_title = clean_name
                st.success("План успешно создан!")
            else:
                st.error(res)

if st.session_state.last_result:
    st.markdown("---")
    st.subheader("📄 Готовый поурочный план:")
    
    docx_bytes = create_docx(st.session_state.last_result)
    
    st.download_button(
        label="📥 Скачать готовый документ Word (.docx)",
        data=docx_bytes,
        file_name=f"{st.session_state.last_title}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )
    st.caption("💡 Файл автоматически сохранится в папку «Загрузки» (Downloads) на вашем устройстве с ровными таблицами и шрифтом Times New Roman.")
            
    st.markdown("---")
    st.markdown(st.session_state.last_result)
