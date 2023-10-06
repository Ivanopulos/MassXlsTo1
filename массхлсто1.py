import os
import pandas as pd

path = 'C:\\Users\\IMatveev\\Desktop\\выкачка\\Новая папка (7)'  # укажите путь к вашей папке
all_files = [f for f in os.listdir(path) if f.endswith('.xlsx') or f.endswith('.xls')]

result = []

for file in all_files:
    filepath = os.path.join(path, file)
    xl = pd.ExcelFile(filepath)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None)

        # Получение значения из A1
        value_A1 = df.iloc[0, 0]

        # Получение значений начиная с A3 по D
        subset = df.iloc[2:, 0:4]

        # Преобразование этого подмножества в одну строку
        combined_string = '\n'.join([' // '.join(map(str, row)) for _, row in subset.iterrows()])

        result.append([value_A1, combined_string])

# Создание итогового DataFrame и запись в файл
result_df = pd.DataFrame(result)
result_df.to_excel(os.path.join(path, 'combined.xlsx'), index=False, header=None)