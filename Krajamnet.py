import pandas as pd

# Загрузка данных из CSV файла
file_path = r'D:\pythoncodes\krajamnet\Отчет_Сургут\08.23.csv'
df = pd.read_csv(file_path, header=None, sep=';')

# Добавление названий столбцов
df.columns = [
    'Дата и время поездки', 'Дата и время авторизации', 'ID-терминала',
    'Номер маршрута', 'ID-владельца маршрута', 'Владелец маршрута', 'Номер ТС',
    'ID-перевозчика', 'Владелец ТС', 'Категория платежа', 'Тип транзакции',
    'Вид льготы', 'Пересадка', 'Сумма', 'Идентификатор БК или ТК'
]

# Ввод пользовательских данных
count = int(input("Введите количество прикладываний: "))

# Фильтрация данных по количеству прикладываний
filtered_df = df.groupby(['Тип транзакции', 'Идентификатор БК или ТК']).filter(lambda x: len(x) >= count)

# Создание Excel файла с несколькими листами
with pd.ExcelWriter('результат.xlsx', engine='openpyxl') as writer:
    # Лист для школьных карт
    filtered_df[filtered_df['Вид льготы'] == 'Карта школьника'].to_excel(writer, sheet_name='Школьные карты', index=False)

    # Лист для социальных карт
    filtered_df[filtered_df['Вид льготы'] == 'СТК'].to_excel(writer, sheet_name='Социальные карты', index=False)

    # Лист для Карт горожанина с проездным на месяц
    filtered_df[(filtered_df['Вид льготы'] == 'Карта горожанина') & (filtered_df['Категория платежа'] == 'Проездной на месяц')].to_excel(writer, sheet_name='Проездной на месяц', index=False)
