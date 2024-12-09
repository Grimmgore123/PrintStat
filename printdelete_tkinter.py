import tkinter as tk
from tkinter import ttk
import wmi
import os
import subprocess

# Получение списка принтеров с помощью WMI
def get_printers():
    c = wmi.WMI()
    printers = c.Win32_Printer()
    # print(printers)
    return [(printer.Name, "Включен" if not printer.WorkOffline else "Выключен", printer.PortName) for printer in printers]

# Обновление списка в Treeview
def update_printer_list():
    if var1.get() == 0:
        for item in tree.get_children():
            tree.delete(item)
        printers = get_printers()
        print(printers)
        for printer in printers:
            for i in printer:
                if "ES4192" and "USB" in i:
                    tree.insert("", tk.END, values=printer)
                    # print(printer[1])
    else:
        for item in tree.get_children():
            tree.delete(item)
        printers = get_printers()
        for printer in printers:
            print(printer[1])
            tree.insert("", tk.END, values=printer)

# Выделение принтера и вывод информации о нем
def select_printer(event):
    selected_item = tree.selection()
    if selected_item:
        item_values = tree.item(selected_item)['values']
        print(f"Выбран принтер: {item_values[0]}")
        print(f"Статус: {item_values[1]}")
        print(f"Состояние: {item_values[2]}")

# Функция для удаления выбранного принтера
def delete_printer():
    selected_item = tree.selection()
    if selected_item:
        item_values = tree.item(selected_item)['values']
        printer_name = item_values[0]
        # Удаление принтера с помощью WMI
        c = wmi.WMI()
        for printer in c.Win32_Printer(Name=printer_name):
            try:
                printer.Delete_()
                print(f"Принтер {printer_name} был удален.")
            except wmi.x_wmi as e:
                print(f"Ошибка при удалении принтера {printer_name}: {e}")
        update_printer_list()  # Обновление списка после удаления

    # Выполняем команду PowerShell и получаем вывод
    result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)

    # Проверяем, что команда выполнена успешно
    if result.returncode == 0:
        print("Перезагрузка USB-устройств выполнена успешно.")
    else:
        print("Ошибка при выполнении команды PowerShell.")
        print(result.stderr)

# Создание главного окна
root = tk.Tk()
root.title("Список принтеров")

# Создание Treeview для отображения списка принтеров
# tree = ttk.Treeview(root, selectmode="extended", columns=("Printer", "Status", "testState"), show="headings")
tree = ttk.Treeview(root, columns=("Printer", "Status", "testState"), show="headings")
tree.column("Printer", width=200)
tree.column("Status", width=100)
tree.column("testState", width=150)
tree.heading("Printer", text="Принтер")
tree.heading("Status", text="Статус")
tree.heading("testState", text="Порт")
tree.pack(fill=tk.BOTH, expand=True)

var1 = tk.IntVar()

def updconf():
    command = 'PnPUtil.exe /restart-device /class "USB"'
    os.system(command)

# Настройка выделения и удаления принтера
tree.bind('<<TreeviewSelect>>', select_printer)
delete_button = tk.Button(root, text="Удалить выбранный принтер", command=delete_printer)
delete_button.pack(fill=tk.BOTH)
updt_button = tk.Button(root, text="Обновить информацию о принтерах", command=updconf)
updt_button.pack(fill=tk.BOTH)
all_print = tk.Checkbutton(root, text="Показать все принтеры", variable=var1, onvalue=1, offvalue=0, command=update_printer_list)
all_print.pack(anchor="w")

# Обновление списка принтеров
update_printer_list()

# Запуск главного цикла обработки событий
root.mainloop()