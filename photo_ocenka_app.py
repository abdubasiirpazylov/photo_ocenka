import streamlit as st
from docx import Document
from docx.shared import Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from PIL import Image, ImageOps # Новые библиотеки для выравнивания фото

st.set_page_config(page_title="Эксперт - Фотоотчет", layout="wide", initial_sidebar_state="collapsed")

st.title("📸 Быстрый Фотоотчет")
st.markdown("Загрузите фото с места осмотра. Приложение автоматически соберет таблицу для отчета.")

reg_num = st.text_input("Гос. номер автомобиля (для названия файла):", placeholder="Например: 04KG021AJF")

st.header("Фотографии")

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

# Железобетонная инициализация памяти
if "photo_order" not in st.session_state:
    st.session_state.photo_order = []  # Хранит ID файлов по порядку
if "processed_ids" not in st.session_state:
    st.session_state.processed_ids = set() # Хранит ID уже загруженных файлов
if "photo_data" not in st.session_state:
    st.session_state.photo_data = {} # Хранит сами картинки в памяти

# --- ФУНКЦИЯ ДЛЯ ВЫРАВНИВАНИЯ ФОТО (EXIF FIX) ---
def fix_image_rotation(uploaded_file):
    image = Image.open(uploaded_file)
    # Читаем EXIF телефона и поворачиваем фото как надо
    image = ImageOps.exif_transpose(image)
    
    # Если формат PNG/RGBA, переводим в обычный RGB
    if image.mode in ("RGBA", "P"):
        image = image.convert("RGB")
        
    # Сохраняем готовую ровную картинку в буфер
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='JPEG', quality=85)
    img_byte_arr.seek(0)
    return img_byte_arr

# --- КОЛБЭКИ ---
def move_photo(file_id, current_index):
    new_val = st.session_state.get(f"pos_{file_id}")
    if new_val is None:
        return
    new_index = new_val - 1
    if new_index != current_index:
        item = st.session_state.photo_order.pop(current_index)
        st.session_state.photo_order.insert(new_index, item)
        # Пересчитываем позиции виджетов
        for idx, fid in enumerate(st.session_state.photo_order):
            st.session_state[f"pos_{fid}"] = idx + 1

def delete_photo(file_id):
    # Полностью удаляем фото из памяти приложения
    if file_id in st.session_state.photo_order:
        st.session_state.photo_order.remove(file_id)
    if file_id in st.session_state.processed_ids:
        st.session_state.processed_ids.remove(file_id)
    if file_id in st.session_state.photo_data:
        del st.session_state.photo_data[file_id]
        
    # Пересчитываем оставшиеся позиции
    for idx, fid in enumerate(st.session_state.photo_order):
        st.session_state[f"pos_{fid}"] = idx + 1

# --- ЗАГРУЗКА ФОТОГРАФИЙ ---
uploaded_photos = st.file_uploader("Сделать снимок или выбрать из галереи", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_photos:
    for uf in uploaded_photos:
        # Если этого фото еще нет в нашей памяти (проверяем по ID)
        if uf.file_id not in st.session_state.processed_ids:
            # 1. Выравниваем и сжимаем фото
            fixed_bytes = fix_image_rotation(uf)
            
            # 2. Сохраняем в "копилку" (теперь оно не исчезнет)
            st.session_state.processed_ids.add(uf.file_id)
            st.session_state.photo_data[uf.file_id] = fixed_bytes
            st.session_state.photo_order.append(uf.file_id)
            
            # 3. Присваиваем позицию для нового фото
            st.session_state[f"pos_{uf.file_id}"] = len(st.session_state.photo_order)

photo_data_list = []

if st.session_state.photo_order:
    st.write("Настройте подписи и порядок фотографий:")
    
    # Отрисовываем фото из нашей "копилки", а не из загрузчика
    for i, file_id in enumerate(list(st.session_state.photo_order)):
        img_bytes = st.session_state.photo_data[file_id]
        
        with st.container():
            c_img, c_controls = st.columns([1, 2])
            
            with c_img:
                st.image(img_bytes, use_container_width=True)
                st.selectbox(
                    "📍 Позиция:",
                    options=list(range(1, len(st.session_state.photo_order) + 1)),
                    key=f"pos_{file_id}",
                    on_change=move_photo,
                    args=(file_id, i)
                )
                st.button("❌ Исключить", key=f"del_{file_id}", on_click=delete_photo, args=(file_id,))

            with c_controls:
                selected_tags = st.multiselect("Шаблонные фразы:", CAPTION_OPTIONS, key=f"tags_{file_id}")
                custom_text = st.text_input("Свой текст:", key=f"custom_{file_id}")
                
                final_caption_parts = []
                if selected_tags:
                    final_caption_parts.append(", ".join(selected_tags))
                if custom_text:
                    final_caption_parts.append(custom_text)
                    
                final_caption = ", ".join(final_caption_parts)
                
                if final_caption:
                    st.caption(f"📝 Подпись: **{final_caption}**")
                    
            st.divider() 
            
            # Собираем данные для генерации документа
            photo_data_list.append({"bytes": img_bytes, "caption": final_caption})

# --- ГЕНЕРАЦИЯ ОТЧЕТА ---
if photo_data_list:
    if st.button("СГЕНЕРИРОВАТЬ ФОТООТЧЕТ", type="primary", use_container_width=True):
        try:
            doc = Document()
            
            heading = doc.add_heading(f"Фотоотчет осмотра: {reg_num if reg_num else 'Автомобиль'}", level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            
            for i in range(0, len(photo_data_list), 2):
                cells = table.add_row().cells
                
                # Левая ячейка
                img1_file = photo_data_list[i]["bytes"]
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
                    img2_file = photo_data_list[i+1]["bytes"]
                    img2_file.seek(0)
                    
                    run2 = p2.add_run()
                    run2.add_picture(img2_file, width=Mm(80))
                    run2.add_break()
                    run2.add_text(photo_data_list[i+1]["caption"])
            
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
