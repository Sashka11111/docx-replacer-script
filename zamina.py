import os
import re
import shutil
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Inches
from docxcompose.composer import Composer

# 🔷 Шляхи
folder_path = r"C:\Users\1\Desktop\ЗВ-41"
dyploma_path = r"C:\Users\1\Desktop\dyploma.docx"

# 🔄 Заміни
replacements = {
    'Назва кваліфікації та присвоєний ступінь': 'Назва освітньої кваліфікації та присвоєний освітньо-професійний ступінь (мовою оригіналу)',
    'Name of qualification and (if applicable) title conferred': 'Name of educational qualification and educational-professional degree conferred (in original language)',
    'Ступінь вищої освіти': 'Освітньо-професійний ступінь фахової передвищої освіти',
    'Degree': 'Professional pre-higher education educational-professional degree',
    'Професійна кваліфікація (у разі присвоєння)': 'Освітньо-професійна програма',
    'Professional Qualification (if awarded)': 'Educational-professional programme',
    'Фахівець з геодезії та землеустрою': '',
    'Specialist in geodesy and land management': '',
    'Основна (основні) галузь (галузі) знань за кваліфікацією': 'Професійна кваліфікація (у разі присвоєння)',
    'Main field(s) of study for the qualification': 'Professional qualification (if awarded)',
    '19 Архітектура та будівництво': 'Фахівець з геодезії та землеустрою',
    '19 Architecture and Construction': 'Specialist in geodesy and land management',
    'Найменування і статус закладу (якщо відмінні від п. 2.3), який реалізує освітню програму': '',
    'Name and status of institution (if different from 2.3)': '',
    'administering studies': '',
    'Зазначено у пункті 2.3': '',
    'Specified in 2.3': '',
    '2.5': '2.4',
    'ІНФОРМАЦІЯ ПРО РІВЕНЬ КВАЛІФІКАЦІЇ І ТРИВАЛІСТЬ ЇЇ ЗДОБУТТЯ': 'ІНФОРМАЦІЯ ПРО КВАЛІФІКАЦІЮ І ТРИВАЛІСТЬ ЇЇ ЗДОБУТТЯ',
    'Тривалість освітньої програми в кредитах та/або роках': 'Офіційна тривалість освітньо-професійної програми в кредитах та/або роках',
    'Official duration of programme in credits and/or years': 'Official length of educational-professional programme in credits and/or years',
    'ІНФОРМАЦІЯ ПРО ЗАВЕРШЕНУ ОСВІТНЮ ПРОГРАМУ ТА ЗДОБУТІ РЕЗУЛЬТАТИ НАВЧАННЯ': 'ІНФОРМАЦІЯ ПРО ЗАВЕРШЕНУ ОСВІТНЬО-ПРОФЕСІЙНУ ПРОГРАМУ ТА ЗДОБУТІ РЕЗУЛЬТАТИ НАВЧАННЯ',
    'INFORMATION ON THE PROGRAMME COMPLETED AND THE RESULTS OBTAINED': 'INFORMATION ON THE COMPLETED EDUCATIONAL-PROFESSIONAL PROGRAMME AND LEARNING OUTCOMES',
    'Найменування всіх закладів вищої освіти (наукових установ) (відокремлених структурних підрозділів...)': 'Найменування всіх закладів фахової передвищої освіти (структурних підрозділів або філій...)',
    'Name of all higher education (research) institutions (separate structural units of higher education institutions)...': 'Names of all professional pre-higher education institutions... (academic mobility programme)',
    'Контактна інформація закладу вищої освіти (наукової установи)': 'Контактна інформація закладу фахової передвищої освіти (іншого суб’єкта освітньої діяльності)',
    'Contact information of the higher education (research) institution': 'Contact information of the professional pre-higher education institution (other educational entity)',
    'Керівник або уповноважена особа закладу вищої освіти': 'Керівник або уповноважена особа закладу фахової передвищої освіти',
    'Capacity': 'Head or other authorized person of professional pre-higher education institution',
    'Посада керівника або іншої уповноваженої особи закладу вищої освіти (наукової установи)': 'Посада керівника або іншої уповноваженої особи закладу фахової передвищої освіти',
    'Position of the Head or another authorized person of the Higher Education (Research) Institution': 'Position of the professional pre-higher education institution head or other authorized person',
    'Печатка': 'Офіційна печатка',
    'Official stamp or seal': 'Official Seal',
    '8. ІНФОРМАЦІЯ ПРО НАЦІОНАЛЬНУ СИСТЕМУ ВИЩОЇ ОСВІТИ': '',
    'ІНФОРМАЦІЯ ПРО СИСТЕМУ ФАХОВОЇ ПЕРЕДВИЩОЇ ОСВІТИ': '',
}

keys_to_format = [
    'Офіційна печатка',
    'Official Seal',
]
no_line_break_keys = keys_to_format

def replace_text_in_paragraph(paragraph, doc):
    # Отримуємо повний текст абзацу та список runs для збереження форматування
    runs = paragraph.runs
    if not runs:
        return True  # Пропускаємо порожні абзаци

    full_text = ''.join(run.text for run in runs).strip()
    if not full_text:
        return True  # Пропускаємо порожні абзаци

    changed = False
    needs_formatting = False
    for old, new in replacements.items():
        if re.search(r'\b' + re.escape(old) + r'\b', full_text):
            if new == '':
                # Не видаляємо абзаци, які починаються з "4.3" або "4.4"
                if full_text.startswith("4.3") or full_text.startswith("4.4"):
                    continue
                # Видаляємо абзац, якщо замінюємо на порожній текст
                paragraph._element.getparent().remove(paragraph._element)
                return False

            full_text = full_text.replace(old, new)
            changed = True
            if old in keys_to_format or new in keys_to_format:
                needs_formatting = True

    if changed:
        # Зберігаємо форматування першого run (шрифт, розмір, колір тощо)
        first_run = runs[0]
        font_name = first_run.font.name
        font_size = first_run.font.size
        font_bold = first_run.font.bold
        font_italic = first_run.font.italic
        font_color = first_run.font.color.rgb if first_run.font.color and first_run.font.color.rgb else None

        # Очищаємо всі runs
        for run in runs:
            run.text = ''

        # Додаємо новий текст із збереженням форматування
        new_run = paragraph.add_run(full_text)
        new_run.font.name = font_name
        new_run.font.size = font_size
        new_run.font.bold = font_bold
        new_run.font.italic = font_italic
        if font_color:
            new_run.font.color.rgb = font_color

        # Застосовуємо вирівнювання, якщо потрібно
        if needs_formatting and full_text not in no_line_break_keys:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    return True

def replace_text_in_cell(cell, doc):
    for paragraph in cell.paragraphs:
        replace_text_in_paragraph(paragraph, doc)

def delete_after_section(doc):
    pattern = re.compile(r'ІНФОРМАЦІЯ\s*ПРО\s*НАЦІОНАЛЬНУ\s*СИСТЕМУ\s*ВИЩОЇ', re.IGNORECASE | re.UNICODE)
    start_index = None
    for i, paragraph in enumerate(doc.paragraphs):
        full_text = ''.join(run.text for run in paragraph.runs).strip()
        if pattern.search(full_text):
            start_index = i
            break

    if start_index is not None:
        # Видаляємо абзаци після знайденого розділу
        for i in reversed(range(start_index + 1, len(doc.paragraphs))):
            doc.paragraphs[i]._element.getparent().remove(doc.paragraphs[i]._element)

        # Видаляємо таблиці, які розташовані після start_index
        for table in reversed(doc.tables):
            if table._element.getparent().index(table._element) > start_index:
                table._element.getparent().remove(table._element)

def append_dyploma_with_formatting(original_path, dyploma_path):
    try:
        temp_path = original_path.replace(".docx", "_temp.docx")
        shutil.copyfile(original_path, temp_path)
        base_doc = Document(temp_path)
        composer = Composer(base_doc)
        dyploma_doc = Document(dyploma_path)
        composer.append(dyploma_doc)
        composer.save(original_path)
        os.remove(temp_path)
        print(f"📎 Вставлено dyploma.docx у {os.path.basename(original_path)}")
    except Exception as e:
        print(f"❌ Помилка вставки dyploma.docx у {os.path.basename(original_path)}: {str(e)}")

def process_docx(file_path, dyploma_path):
    try:
        doc = Document(file_path)
    except Exception as e:
        print(f"❌ Помилка при відкритті {file_path}: {str(e)}")
        return

    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, doc)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_in_cell(cell, doc)

    delete_after_section(doc)

    try:
        doc.save(file_path)
    except Exception as e:
        print(f"❌ Помилка при збереженні {file_path}: {str(e)}")
        return

    if os.path.exists(dyploma_path):
        append_dyploma_with_formatting(file_path, dyploma_path)
    else:
        print(f"❌ Не знайдено файл dyploma.docx за шляхом {dyploma_path}")

    print(f"✅ Оброблено: {os.path.basename(file_path)}")

# 🔁 Обробка всіх файлів
for filename in os.listdir(folder_path):
    if filename.endswith('.docx') and not filename.startswith('~$'):
        full_path = os.path.join(folder_path, filename)
        print(f"\n📄 Починаємо: {filename}")
        process_docx(full_path, dyploma_path)
