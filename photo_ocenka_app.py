import streamlit as st
from docx import Document
from docx.shared import Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# Настройка страницы (идеально для мобилок и планшетов)
st.set_page_config(page_title="Эксперт - Фотоотчет", layout="wide", initial_sidebar_state="collapsed")

st.title("📸 Быстрый Фотоотчет")
st.markdown("Загрузите фото с места осмотра. Приложение автоматически соберет таблицу для отчета.")

# Поле для названия файла
reg_num = st.text_input("Гос. номер автомобиля (для названия файла):", placeholder="Например: 04KG021AJF")

st.header("Фотографии")

# Шаблонные фразы
CAPTION_OPTIONS = [
    "Вид спереди", "Вид сзади", "Вид слева", "Вид справа",
    "VIN-код", "Показания одометра", "Обзорный снимок салона",
    "вмятина", "вытяжение", "гофра", "деформация", "изгиб", 
    "повреждение ЛКП", "потертость", "разрушение", "разрыв", 
    "раскол", "складка", "царапина", "скол", "перекос",
    "без образования видимых складок", "с глубокой вытяжкой металла", 
    "с нарушением геометрии кромок", "с нарушением геометрии ребер жесткости", 
    "с незначительной вытяжкой металла", "с образованием незначительных складок", 
    "с образованием острых складок", "с образованием плавных складок", "с образованием трещин",
    "на площади менее 10%", "на площади от 10 до 20%", 
    "на площади от 20 до 30%", "на площади от 30 до 40%", "на площади более 40%"
]

# Инициализация состояния
if "photo_order" not in st.session_state:
    st.session_state.photo_order = []
if "last_files" not in st.session_state:
    st.session_state.last_files = []
if "deleted_files" not in st.session_state:
    st.session_state.deleted_files = set()

# --- КОЛБЭКИ ---
def move_photo(file_name, current_index):
    new_val = st.session_state.get(f"pos_{file_name}")
    if new_val is None:
        return
    new_index = new_val - 1
    if new_index != current_index:
        item = st.session_state.photo_order.pop(current_index)
        st.session_state.photo_order.insert(new_index, item)
        for idx, fname in enumerate(st.session_state.photo_order):
            st.session_state[f"pos_{fname}"] = idx + 1

def delete_photo(filename):
    if filename in st.session_state.photo_order:
        st.session_state.photo_order.remove(filename)
        st.session_state.deleted_files.add(filename)
        for idx, fname in enumerate(st.session_state.photo_order):
            st.session_state[f"pos_{fname}"] = idx + 1

# --- ЗАГРУЗКА ---
uploaded_photos = st.file_uploader("Загрузите фотографии", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
photo_data_list = []

if uploaded_photos:
    current_filenames = [f.name for f in uploaded_photos]
    
    if set(current_filenames) != set(st.session_state.last_files):
        new_files = [f for f in current_filenames if f not in st.session_state.last_files]
        removed_files = [f for f in st.session_state.last_files if f not in current_filenames]
        
        for f in new_files:
            if f in st.session_state.deleted_files:
                st.session_state.deleted_files.remove(f)
            if f not in st.session_state.photo_order and f not in st.session_state.deleted_files:
                st.session_state.photo_order.append(f)
                
        for f in removed_files:
            if f in st.session_state.photo_order:
                st.session_state.photo_order.remove(f)
            if f in st.session_state.deleted_files:
                st.session_state.deleted_files.remove(f)
                
        st.session_state.last_files = current_filenames
        
        for idx, fname in enumerate(st.session_state.photo_order):
            key = f"pos_{fname}"
            if key not in st.session_state:
                st.session_state[key] = idx + 1

    file_dict = {f.name: f for f in uploaded_photos}

    st.write("Настройте подписи и порядок фотографий:")
    
    for i, filename in enumerate(list(st.session_state.photo_order)):
        if filename not in file_dict:
            continue
            
        photo = file_dict[filename]
        
        # Для мобильных экранов делаем вертикальную компоновку более удобной
        with st.container():
            c_img, c_controls = st.columns([1, 2])
            
            with c_img:
                st.image(photo, use_container_width=True)
                st.selectbox(
                    "📍 Позиция:",
                    options=list(range(1, len(st.session_state.photo_order) + 1)),
                    key=f"pos_{filename}",
                    on_change=move_photo,
                    args=(filename, i)
                )
                st.button("❌ Исключить", key=f"del_{filename}", on_click=delete_photo, args=(filename,))

            with c_controls:
                selected_tags = st.multiselect("Шаблонные фразы:", CAPTION_OPTIONS, key=f"tags_{filename}")
                custom_text = st.text_input("Свой текст:", key=f"custom_{filename}")
                
                final_caption_parts = []
                if selected_tags:
                    final_caption_parts.append(", ".join(selected_tags))
                if custom_text:
                    final_caption_parts.append(custom_text)
                    
                final_caption = ", ".join(final_caption_parts)
                
                if final_caption:
                    st.caption(f"📝 Подпись: **{final_caption}**")
                    
            st.divider() 
            
            photo_data_list.append({"file": photo, "caption": final_caption})

# --- ГЕНЕРАЦИЯ ---
if photo_data_list:
    if st.button("СГЕНЕРИРОВАТЬ ФОТООТЧЕТ", type="primary", use_container_width=True):
        try:
            # Создаем Word документ с нуля (без сторонних шаблонов)
            doc = Document()
            
            # Добавляем заголовок документа
            heading = doc.add_heading(f"Фотоотчет осмотра: {reg_num if reg_num else 'Автомобиль'}", level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Создаем таблицу
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            
            for i in range(0, len(photo_data_list), 2):
                cells = table.add_row().cells
                
                # Левая ячейка
                img1_file = photo_data_list[i]["file"]
                img1_file.seek(0)
                
                p1 = cells[0].paragraphs[0]
                p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run1 = p1.add_run()
                run1.add_picture(img1_file, width=Mm(80))
                run1.add_break()
                run1.add_text(photo_data_list[i]["caption"])
                
                # Правая ячейка
                p2 = cells[1].paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if i + 1 < len(photo_data_list):
                    img2_file = photo_data_list[i+1]["file"]
                    img2_file.seek(0)
                    
                    run2 = p2.add_run()
                    run2.add_picture(img2_file, width=Mm(80))
                    run2.add_break()
                    run2.add_text(photo_data_list[i+1]["caption"])
            
            # Сохранение в буфер
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ Отчет успешно сгенерирован!")
            
            safe_reg_num = reg_num.strip() if reg_num.strip() else "Без_номера"
            file_name = f"Фотоотчет_{safe_reg_num}.docx"
            
            st.download_button(
                label=f"📥 СКАЧАТЬ {file_name}",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при генерации отчета: {e}")
