import pandas as pd

# Загрузка данных из Excel-файла
file_path = 'Book1.xlsx'  # Укажите путь к вашему файлу Excel
df = pd.read_excel(file_path)

while True:
    print("\nМеню:")
    print("1. Вывести весь набор данных")
    print("2. Вывести конкретную строку по индексу")
    print("3. Вывести столбец по индексу")
    print("4. Вывести ячейку по индексу")
    print("5. Выход")

    choice = input("Выберите опцию: ")

    if choice == '1':
        print(df)
    elif choice == '2':
        index = int(input("Введите индекс строки: "))
        print(df.iloc[index])
    elif choice == '3':
        column_index = int(input("Введите индекс столбца: "))
        print(df.iloc[:, column_index])
    elif choice == '4':
        row_index = int(input("Введите индекс строки: "))
        col_index = int(input("Введите индекс столбца: "))
        print(df.iat[row_index, col_index])

        cell_choice = input("Что вы хотите сделать с ячейкой? (Удалить, изменить, создать): ")
        if cell_choice == 'Удалить':
            df.iat[row_index, col_index] = None
        elif cell_choice == 'Изменить':
            new_value = input("Введите новое значение: ")
            df.iat[row_index, col_index] = new_value
        elif cell_choice == 'Создать':
            new_value = input("Введите новое значение: ")
            df.iat[row_index, col_index] = new_value

        # Сохранение изменений в Excel-файл
        df.to_excel(file_path, index=False)
    elif choice == '5':
        break
    else:
        print("Неверный выбор. Пожалуйста, выберите снова.")