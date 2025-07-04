import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

def set_paragraph_style(paragraph, alignment=None, font_size=None):
    """Налаштування стилів параграфа: вирівнювання та розмір шрифта."""
    if alignment is not None:
        paragraph.alignment = alignment
    if font_size is not None:
        for run in paragraph.runs:
            run.font.size = Pt(font_size)

def process_document(file_path):
    """Обробка одного документа для форматування пункту 8."""
    try:
        doc = Document(file_path)
        in_section_8 = False

        for paragraph in doc.paragraphs:
            # Перевіряємо, чи параграф є початком пункту 8
            if paragraph.text.strip().startswith("8."):
                in_section_8 = True
            # Якщо починається наступний пункт, виходимо з пункту 8
            elif in_section_8 and paragraph.text.strip().startswith(("9.", "10.", "11.")):
                in_section_8 = False

            # Застосовуємо стилі до параграфів пункту 8
            if in_section_8:
                set_paragraph_style(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, font_size=9)

        # Зберігаємо оновлений документ
        doc.save(file_path)
        print(f"✅ Оброблено: {os.path.basename(file_path)}")
    except Exception as e:
        print(f"❌ Помилка при обробці {os.path.basename(file_path)}: {str(e)}")

def process_folder(folder_path):
    """Обробка всіх .docx файлів у папці."""
    if not os.path.exists(folder_path):
        print(f"❌ Папка {folder_path} не існує")
        return

    for filename in os.listdir(folder_path):
        if filename.endswith('.docx') and not filename.startswith('~$'):
            file_path = os.path.join(folder_path, filename)
            print(f"\n📄 Обробка файлу: {filename}")
            process_document(file_path)
    print("\n🎉 Обробка всіх файлів завершена!")

# Запитуємо шлях до папки та запускаємо обробк
folder_path = input(r"C:\Users\1\Desktop\ЗВ-41")
process_folder(folder_path)
