import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.neighbors import NearestNeighbors
import matplotlib.pyplot as plt
import os


# 1. Загрузка данных
df_external = pd.read_excel("external_products.xlsx")
df_systeme = pd.read_excel("systeme_products.xlsx")

# 2. Объединяем текстовые характеристики
cols_to_merge = ['Наименование', 'Характеристика_1', 'Характеристика_2']
df_external['text_features'] = df_external[cols_to_merge].astype(str).agg(' '.join, axis=1)
df_systeme['text_features'] = df_systeme[cols_to_merge].astype(str).agg(' '.join, axis=1)

# 3. Векторизация текста
vectorizer = TfidfVectorizer()
X_external = vectorizer.fit_transform(df_external['text_features'])
X_systeme = vectorizer.transform(df_systeme['text_features'])

# 4. Поиск ближайшего аналога
nn_model = NearestNeighbors(n_neighbors=1, metric='cosine')
nn_model.fit(X_external)

distances, indices = nn_model.kneighbors(X_systeme)

# 5. Добавляем результаты в таблицу
df_systeme['closest_match_index'] = indices.flatten()
df_systeme['closest_match_name'] = df_external.loc[indices.flatten(), 'Наименование'].values
df_systeme['similarity_score'] = 1 - distances.flatten()  # чем ближе к 1, тем больше сходство

# 6. Сохраняем результат
df_systeme.to_excel("matches.xlsx", index=False)

print(" Поиск завершен. Результат сохранен в matches.xlsx")

plt.figure(figsize=(8, 5))
plt.bar(df_systeme['Наименование'], df_systeme['similarity_score'], color='skyblue')
plt.title('Сходство продукции Systeme Electric с аналогами', fontsize=14)
plt.xlabel('Товары Systeme Electric')
plt.ylabel('Коэффициент сходства (0–1)')
plt.ylim(0, 1)
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()

plt.savefig("similarity_chart.png", dpi=200)
plt.show()

print("📊 График сохранён как similarity_chart.png")