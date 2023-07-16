import wmi
import win32print
import win32api
import tempfile

def convert_bytes_to_gb(bytes_value):
    # Конвертирование байтов в гигабайты
    gb_value = bytes_value / (1024 ** 3)
    return round(gb_value, 2)

def get_disk_type(media_type):
    # Определение типа диска по MediaType
    if media_type == 3:
        return "HDD"
    elif media_type == 4:
        return "SSD"
    else:
        return "Неизвестно"

def print_computer_info():
    # Получение имени принтера по умолчанию
    printer_name = win32print.GetDefaultPrinter()

    # Подключение к WMI-серверу
    computer_name = input('Введи имя ПК: ')
    c = wmi.WMI(computer=computer_name)

    # Создание временного файла для сохранения информации
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
    temp_file_path = temp_file.name

    try:
        # Получение информации о компьютере
        computer_system = c.Win32_ComputerSystem(Name=computer_name)[0]
        base_board = c.Win32_BaseBoard()[0]
        os = c.Win32_OperatingSystem()[0]
        processor = c.Win32_Processor()[0]
        physical_memory = c.Win32_PhysicalMemory()
        disks = c.Win32_DiskDrive()
        physical_media = c.Win32_PhysicalMedia()

        # Запись информации о компьютере во временный файл
        with open(temp_file_path, "w") as file:
            file.write("Информация о компьютере:\n")
            file.write(f"Имя компьютера: {computer_system.Name}\n")
            file.write(f"Модель материнской платы: {base_board.Product}\n\n")

            file.write("Процессор:\n")
            file.write(f"Имя: {processor.Name}\n\n")

            file.write("Оперативная память:\n")
            for memory in physical_memory:
                capacity_gb = convert_bytes_to_gb(int(memory.Capacity))
                file.write(f"Модуль {capacity_gb} GB\n")
            file.write("\n")

            file.write("Накопители:\n")
            for disk in disks:
                file.write(f"Модель: {disk.Model}\n")
                file.write(f"Емкость: {convert_bytes_to_gb(int(disk.Size))} GB\n")
                # Поиск соответствующего физического носителя для диска
                for media in physical_media:
                    if media.Tag.strip() == disk.DeviceID.strip():
                        disk_type = get_disk_type(media.MediaType)
                        file.write(f"Тип: {disk_type}\n")
                        break
                file.write("\n")

        # Отправка временного файла на принтер
        win32api.ShellExecute(0, "print", temp_file_path, f'"{printer_name}"', ".", 0)

        print("Информация отправлена на печать.")
    finally:
        temp_file.close()

# Пример использования
print_computer_info()