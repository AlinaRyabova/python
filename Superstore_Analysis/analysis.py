import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# 1. Завантаження даних
file_path = "./Sample - Superstore.xlsx"  # або повний шлях
data = pd.read_excel(file_path, sheet_name='Orders')

# 2. Аналіз сезонності
data['Order Date'] = pd.to_datetime(data['Order Date'])
data['Month'] = data['Order Date'].dt.month
monthly_sales = data.groupby('Month')['Sales'].sum()

plt.figure(figsize=(10, 5))
monthly_sales.plot(kind='bar', color='skyblue')
plt.title('Сезонність продажів за місяцями')
plt.xlabel('Місяць')
plt.ylabel('Сума продажів')
plt.grid(axis='y')
plt.savefig('seasonality.png')  # Збереження графіка
plt.show()

# 3. Визначення аномалій
Q1 = data['Sales'].quantile(0.25)
Q3 = data['Sales'].quantile(0.75)
IQR = Q3 - Q1
anomaly_threshold = 1.5 * IQR
anomalies = data[(data['Sales'] < (Q1 - anomaly_threshold)) | 
                (data['Sales'] > (Q3 + anomaly_threshold))]
print(f"Знайдено {len(anomalies)} аномалійних транзакцій")

# 4. Когортний аналіз (наприклад, за сегментами)
segment_stats = data.groupby('Segment').agg({
    'Sales': ['mean', 'median', 'std'],
    'Profit': 'sum'
})
print("\nСтатистика за сегментами:")
print(segment_stats)

# 5. Візуалізація
data.groupby('Category')['Sales'].sum().plot.pie(
    autopct='%1.1f%%',
    explode=(0.1, 0, 0),
    shadow=True
)
plt.title('Розподіл продажів за категоріями')
plt.savefig('categories.png')
plt.show()

# 6. Збереження результатів
with pd.ExcelWriter('analysis_results.xlsx') as writer:
    data.describe().to_excel(writer, sheet_name='Описова статистика')
    anomalies.to_excel(writer, sheet_name='Аномалії', index=False)
    segment_stats.to_excel(writer, sheet_name='Статистика за сегментами')

print("Аналіз завершено! Результати збережено у файлі analysis_results.xlsx")