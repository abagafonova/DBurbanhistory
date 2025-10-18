from docx import Document  
import pandas as pd  
  
# Загрузка справочника "Губерния → Регион"  
region_df = pd.read_excel("region_table.xlsx")  
gubernia_to_region = {}  
  
for _, row in region_df.iterrows():  
    region = row["Регион"]  
    gubernias = str(row["Губернии"]).split(",")  
    for gub in gubernias:  
        gubernia_to_region[gub.strip()] = region.strip()  
  
# Загрузка Word-документа  
doc = Document("postanovleniya.docx")  
  
# Переменные текущего контекста  
current_gubernia = ""  
current_city = ""  
current_section = ""  
rows = []  
  
# Обход абзацев документа  
for para in doc.paragraphs:  
    style = para.style.name  
    text = para.text.strip()  
    if not text:  
        continue  
  
    if style == "Heading 1":  
        current_gubernia = text  
    elif style == "Heading 2":  
        current_city = text  
    elif style == "Heading 3":  
        current_section = text  
    else:  
        if text.lower().startswith("примечани"):  
            if rows:  
                rows[-1]["Примечание"] += " " + text  
        else:  
            region = gubernia_to_region.get(current_gubernia.strip(), "")  
            rows.append({  
                "Регион": region,  
                "Губерния/Область": current_gubernia,  
                "Город/Село": current_city,  
                "Раздел постановления": current_section,  
                "Формулировка": text,  
                "Примечание": "",  
                "Источник (страница)": "",  
                "Совпадение с Врем. правилами (да/нет)": ""  
            })  
  
# Сохранение в Excel  
df = pd.DataFrame(rows)  
df.to_excel("result_postanovleniya.xlsx", index=False)  
print("✅ Готово! Сохранено в result_postanovleniya.xlsx")
