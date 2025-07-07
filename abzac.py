import os
from docx import Document

# 📁 Шлях до папки з .docx файлами
folder_path = r"C:\Users\1\Desktop\ЗВ-41 – копія"

def remove_specific_paragraphs(doc):
    removed = False
    # Список цільових фраз для видалення
    target_phrases = [
        "Contact information of the professional pre-higher education institution (other educational entity)",
        "Head or other authorized person of professional pre-higher education institution"
    ]

    # Збираємо параграфи для видалення
    paragraphs_to_remove = []
    for paragraph in doc.paragraphs:
        if any(phrase in paragraph.text for phrase in target_phrases):
            paragraphs_to_remove.append(paragraph)

    # Видаляємо знайдені параграфи
    for paragraph in paragraphs_to_remove:
        p = paragraph._element
        p.getparent().remove(p)
        removed = True

    return removed

# 🔁 Проходимо всі файли в папці
for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        file_path = os.path.join(folder_path, filename)
        print(f"📂 Обробка: {filename}")
        try:
            doc = Document(file_path)
            if remove_specific_paragraphs(doc):
                doc.save(file_path)  # 💾 Зберігаємо зміни
                print(f"✅ Оновлено: {filename}")
            else:
                print(f"ℹ️ Не знайдено абзаців для видалення в {filename}")
        except Exception as e:
            print(f"❌ Помилка в {filename}: {e}")
