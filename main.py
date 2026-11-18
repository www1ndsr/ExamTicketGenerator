import os
import random
import msvcrt
from datetime import datetime

from openpyxl import Workbook, load_workbook


def save_student_to_excel(last_name, first_name, ticket_number):
    """Добавляет данные студента в журнал и сохраняет файл."""
    file_name = "journal.xlsx"
    record_date = datetime.now()

    while True:
        if os.path.exists(file_name):
            workbook = load_workbook(file_name)
            worksheet = workbook.active
        else:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["Last name", "First name", "Номер билета", "Дата и время"])

        worksheet.append([last_name, first_name, ticket_number, record_date])

        try:
            workbook.save(file_name)
            return
        except PermissionError:
            print(
                "Не удалось сохранить данные. Возможно, файл journal.xlsx открыт "
                "в Excel. Закройте файл и нажмите Enter для повторной попытки"
            )
            input()


def get_input_with_escape(prompt):
    """Запрашивает непустую строку; возвращает None при нажатии ESC."""
    while True:
        print(prompt, end="", flush=True)
        first_character = msvcrt.getwch()

        if first_character == "\x1b":
            print()
            return None

        # Enter подтверждает пустую строку; остальные символы становятся ее началом.
        if first_character == "\r":
            print()
            value = ""
        else:
            print(first_character, end="", flush=True)
            value = first_character + input()

        value = value.strip()
        if value:
            return value

        print("Поле не может быть пустым. Пожалуйста, введите значение.")


def main():
    # Сообщаем пользователю, как завершить работу приложения.
    print("Для выхода нажмите ESC")

    # Повторяем ввод данных и выдачу билета, пока пользователь не выйдет.
    while True:
        last_name = get_input_with_escape("Last name: ")
        if last_name is None:
            print("Программа завершена.")
            break
        if last_name.upper() == "ESC":
            print("Программа завершена.")
            break

        first_name = get_input_with_escape("First name: ")
        if first_name is None:
            print("Программа завершена.")
            break
        if first_name.upper() == "ESC":
            print("Программа завершена.")
            break

        # Выбираем случайный номер экзаменационного билета.
        ticket_number = random.randint(1, 20)
        print(f"Билет № {ticket_number}")
        save_student_to_excel(last_name, first_name, ticket_number)


if __name__ == "__main__":
    main()