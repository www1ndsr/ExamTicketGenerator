import os
import random
import msvcrt
from abc import ABC, abstractmethod
from datetime import datetime

from openpyxl import Workbook, load_workbook


class TextReader(ABC):
    @abstractmethod
    def read_text(self, prompt):
        pass


class TextWriter(ABC):
    @abstractmethod
    def write(self, text):
        pass


class TextValidator(ABC):
    @abstractmethod
    def validate(self, value):
        pass


class TicketGenerator(ABC):
    @abstractmethod
    def generate(self):
        pass


class Journal(ABC):
    @abstractmethod
    def save(self, last_name, first_name, ticket_number):
        pass


class LettersOnlyValidator(TextValidator):
    """Проверяет, состоит ли значение только из букв."""

    def validate(self, value):
        return bool(value) and all(character.isalpha() for character in value)


class WindowsConsoleReader(TextReader):
    """Консольный ввод для Windows с поддержкой клавиши Esc."""

    def read_text(self, prompt):
        while True:
            print(prompt, end="", flush=True)
            first_character = msvcrt.getwch()

            if first_character == "\x1b":
                print()
                return None

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


class ConsoleWriter(TextWriter):
    def write(self, text):
        print(text)


class RandomTicketGenerator(TicketGenerator):
    def __init__(self, minimum=1, maximum=20):
        self.minimum = minimum
        self.maximum = maximum

    def generate(self):
        return random.randint(self.minimum, self.maximum)


class ExcelJournal(Journal):
    def __init__(self, file_name="journal.xlsx"):
        self.file_name = file_name

    def save(self, last_name, first_name, ticket_number):
        record_date = datetime.now()

        while True:
            if os.path.exists(self.file_name):
                workbook = load_workbook(self.file_name)
                worksheet = workbook.active
            else:
                workbook = Workbook()
                worksheet = workbook.active
                worksheet.append(["Last name", "First name", "Номер билета", "Дата и время"])

            worksheet.append([last_name, first_name, ticket_number, record_date])

            try:
                workbook.save(self.file_name)
                return
            except PermissionError:
                print(
                    "Не удалось сохранить данные. Возможно, файл journal.xlsx открыт "
                    "в Excel. Закройте файл и нажмите Enter для повторной попытки"
                )
                input()


class ExamApplication:
    """Координирует сценарий работы через абстракции зависимостей."""

    def __init__(self, reader, writer, validator, ticket_generator, journal):
        self.reader = reader
        self.writer = writer
        self.validator = validator
        self.ticket_generator = ticket_generator
        self.journal = journal

    def read_name(self, prompt):
        while True:
            value = self.reader.read_text(prompt)
            if value is None:
                return None

            if self.validator.validate(value):
                return value

            self.writer.write("Разрешены только буквы. Введите значение без цифр и знаков.")

    def run(self):
        self.writer.write("Для выхода нажмите ESC")

        while True:
            last_name = self.read_name("Last name: ")
            if last_name is None:
                break

            first_name = self.read_name("First name: ")
            if first_name is None:
                break

            ticket_number = self.ticket_generator.generate()
            self.writer.write(f"Билет № {ticket_number}")
            self.journal.save(last_name, first_name, ticket_number)

        self.writer.write("Программа завершена.")


# Сохраняем прежние функции для совместимости с существующими вызовами.
def get_input_with_escape(prompt):
    application = ExamApplication(
        reader=WindowsConsoleReader(),
        writer=ConsoleWriter(),
        validator=LettersOnlyValidator(),
        ticket_generator=RandomTicketGenerator(),
        journal=ExcelJournal(),
    )
    return application.read_name(prompt)


def save_student_to_excel(last_name, first_name, ticket_number):
    ExcelJournal().save(last_name, first_name, ticket_number)


def main():
    application = ExamApplication(
        reader=WindowsConsoleReader(),
        writer=ConsoleWriter(),
        validator=LettersOnlyValidator(),
        ticket_generator=RandomTicketGenerator(),
        journal=ExcelJournal(),
    )
    application.run()


if __name__ == "__main__":
    main()