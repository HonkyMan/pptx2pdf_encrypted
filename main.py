import os
import subprocess
from pathlib import Path
import pikepdf
import yaml
from typing import List, Optional
from pathlib import Path

class Config:
    """Класс для загрузки конфигурации из YAML файла.

    Attributes:
        config (dict): Словарь с конфигурацией, загруженной из файла.
    """

    def __init__(self, config_path: str) -> None:
        """Инициализирует класс, загружая конфигурацию из указанного файла.

        Args:
            config_path (str): Путь к файлу конфигурации.
        """
        with open(config_path, 'r') as file:
            self.config = yaml.safe_load(file)

    def get(self, key: str) -> Optional[str]:
        """Возвращает значение по ключу из конфигурационного файла.

        Args:
            key (str): Ключ, для которого нужно получить значение.

        Returns:
            Optional[str]: Значение по указанному ключу или None, если ключ не найден.
        """
        return self.config.get(key)

class PDFConverter:
    def __init__(self, source_dir: str, dist_dir: str, convert_command_prefix: str) -> None:
        """Инициализирует конвертер с настройками для конвертации.

        Args:
            source_dir (str): Исходная директория с файлами PPTX.
            dist_dir (str): Целевая директория для сохранения PDF.
            convert_command_prefix (str): Команда для запуска конвертера.
        """
        self.source_dir = source_dir
        self.dist_dir = dist_dir
        self.convert_command_prefix = convert_command_prefix

    def convert(self) -> List[str]:
        """Проходит по всем файлам в source_dir, конвертирует PPTX в PDF и возвращает список путей к PDF.

        Returns:
            List[str]: Список путей к сконвертированным PDF файлам.
        """
        converted_files = []
        for root, _, files in os.walk(self.source_dir):
            relative_path = os.path.relpath(root, self.source_dir)
            target_dir = os.path.join(self.dist_dir, relative_path)
            os.makedirs(target_dir, exist_ok=True)

            for file in files:
                if file.endswith(".pptx"):
                    pdf_path = self._convert_file(root, file, target_dir)
                    if pdf_path:
                        converted_files.append(pdf_path)
        return converted_files

    def _convert_file(self, root: str, file: str, target_dir: str) -> Optional[str]:
        """Конвертирует отдельный файл PPTX в PDF и возвращает путь к PDF.

        Args:
            root (str): Корневая директория, в которой находится файл.
            file (str): Имя файла для конвертации.
            target_dir (str): Целевая директория для сохранения результата.

        Returns:
            Optional[str]: Путь к сконвертированному PDF файлу или None в случае ошибки.
        """
        source_file = os.path.join(root, file)
        target_file_pdf = os.path.splitext(file)[0] + '.pdf'
        target_file = os.path.join(target_dir, target_file_pdf)
        
        convert_command = f'{self.convert_command_prefix} --convert-to pdf --outdir "{target_dir}" "{source_file}"'
        try:
            subprocess.run(convert_command, shell=True, check=True)
            # После выполнения команды проверяем, существует ли сконвертированный файл
            if Path(target_file).is_file():
                print(f"Файл {file} успешно сконвертирован в PDF.")
                return target_file
            else:
                # Если файл не существует, сообщаем об ошибке
                print(f"Ошибка: Файл {file} не был сконвертирован в PDF. Результирующий файл не найден.")
                return None
        except subprocess.CalledProcessError:
            print(f"Ошибка при конвертации файла {file}.")
            return None

class PDFProtector:
    """Класс для наложения защиты на PDF файлы."""

    def __init__(self, owner_password: str) -> None:
        """Инициализирует защитник PDF с паролем владельца.

        Args:
            owner_password (str): Пароль владельца для наложения защиты.
        """
        self.owner_password = owner_password

    def protect(self, pdf_path: str) -> None:
        """Накладывает защиту на указанный PDF файл.

        Args:
            pdf_path (str): Путь к PDF файлу для наложения защиты.
        """
        with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
            pdf.save(pdf_path,
                     encryption=pikepdf.Encryption(
                         owner=self.owner_password,
                         allow=pikepdf.Permissions(
                             extract=False,
                             modify_annotation=False,
                             modify_assembly=False,
                             modify_form=False,
                             modify_other=False,
                             print_highres=False,
                             print_lowres=False,
                         )
                     ))
            print(f"Защита наложена на {pdf_path}")

class App:
    """Основной класс приложения для управления процессом конвертации и защиты."""

    def __init__(self, config_path: str) -> None:
        """Инициализирует приложение с конфигурацией из файла.

        Args:
            config_path (str): Путь к файлу конфигурации.
        """
        self.config = Config(config_path)
        self.converter = PDFConverter(
            source_dir=self.config.get('source_dir'),
            dist_dir=self.config.get('dist_dir'),
            convert_command_prefix='/opt/homebrew/bin/soffice'
        )
        self.protector = PDFProtector(
            owner_password=self.config.get('owner_password')
        )

    def run(self) -> None:
        """Запускает процесс конвертации и защиты."""
        Path(self.config.get('dist_dir')).mkdir(parents=True, exist_ok=True)
        converted_files = self.converter.convert()
        for pdf_path in converted_files:
            self.protector.protect(pdf_path)

if __name__ == "__main__":
    app = App('config.yaml')
    app.run()
