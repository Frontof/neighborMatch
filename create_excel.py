import pandas as pd

# --- внешний каталог ---
data_external = {
    "Производитель": ["Производитель A", "Производитель B", "Производитель C"],
    "Наименование": ["Двигатель X100", "Двигатель Y200", "Двигатель Z300"],
    "Характеристика_1": ["Мощность 100W", "Мощность 200W", "Мощность 300W"],
    "Характеристика_2": ["Напряжение 220V", "Напряжение 220V", "Напряжение 380V"]
}
df_external = pd.DataFrame(data_external)
df_external.to_excel("external_products.xlsx", index=False)

# --- система System Electric ---
data_systeme = {
    "Наименование": ["Двигатель X", "Двигатель Y", "Двигатель Z"],
    "Характеристика_1": ["Мощность 100W", "Мощность 210W", "Мощность 300W"],
    "Характеристика_2": ["Напряжение 220V", "Напряжение 220V", "Напряжение 380V"]
}
df_systeme = pd.DataFrame(data_systeme)
df_systeme.to_excel("systeme_products.xlsx", index=False)

print("✅ Файлы external_products.xlsx и systeme_products.xlsx успешно созданы!")
