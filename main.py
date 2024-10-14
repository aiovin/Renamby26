import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
from mutagen.id3 import ID3, TRCK
import traceback

class Renamby:
    def __init__(self, root):
        self.root = root
        self.root.title('Renamby26')

        self.set_fixed_size(350, 200)

        # Grid для размещения элементов
        self.label = tk.Label(root, text="Select files or a folder to process")
        self.label.grid(row=0, column=0, columnspan=3, pady=10, sticky='n')

        # Контейнер для кнопок выбора файлов и папки
        button_frame = tk.Frame(root)
        button_frame.grid(row=1, column=0, columnspan=3, pady=5)
        
        self.select_files_button = tk.Button(button_frame, text="Select Files", command=self.select_files)
        self.select_files_button.pack(side=tk.LEFT, padx=10)
        
        self.select_folder_button = tk.Button(button_frame, text="Select Folder", command=self.select_folder)
        self.select_folder_button.pack(side=tk.LEFT, padx=10)

        self.rename_button = tk.Button(root, text="Add/Update prefix according to its №", command=self.rename_files, state=tk.DISABLED)
        self.rename_button.grid(row=2, column=0, columnspan=3, pady=5, sticky='n')

        self.exit_button = tk.Button(root, text="Exit", command=root.quit)
        self.exit_button.grid(row=3, column=0, columnspan=3, pady=5, sticky='n')

        # Веса строк и столбцов для выравнивания
        root.grid_rowconfigure(0, weight=1)
        root.grid_rowconfigure(1, weight=1)
        root.grid_rowconfigure(2, weight=1)
        root.grid_rowconfigure(3, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)
        root.grid_columnconfigure(2, weight=1)

        self.selected_files = []

        # Кнопка для ручного добавления префиксов
        self.manual_prefix_button = tk.Button(root, text="Manual Sorting", command=self.open_manual_prefix_window)
        self.manual_prefix_button.grid(row=4, column=0, columnspan=3, pady=5, sticky='n')

    def set_fixed_size(self, width, height):
        self.root.geometry(f"{width}x{height}")  # Размер окна
        self.root.resizable(False, False)  # Отключить изменение размера окна
        
        # Центрирование окна с учетом смещения
        window_width = width
        window_height = height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Расчет центра экрана
        center_x = (screen_width // 2) - (window_width // 2)
        center_y = (screen_height // 2) - (window_height // 2)
        
        # Смещение на 4/5 вверх от центра
        offset_y = screen_height * 3 // 7 - (window_height // 2)
        
        # Новое значение y с учетом смещения
        x = center_x
        y = offset_y
        
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def select_files(self):
        file_paths = filedialog.askopenfilenames(title="Select Files")
        if file_paths:
            self.selected_files = list(file_paths)
            file_count = len(self.selected_files)
            self.label.config(text=f"Files selected: {file_count}")
            self.rename_button.config(state=tk.NORMAL)

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="Select Folder")
        if folder_path:
            self.selected_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
            file_count = len(self.selected_files)
            self.label.config(text=f"Files selected: {file_count}")
            self.rename_button.config(state=tk.NORMAL)

    def get_file_number(self, file_path):
        shell = win32com.client.Dispatch("Shell.Application")
        folder_path = os.path.dirname(file_path)
        folder_path = os.path.normpath(folder_path)  # Преобразовать путь в формат Windows
        file_name = os.path.basename(file_path)

        # Проверка на существование файла
        if not os.path.isfile(file_path):
            print(f"Error: file does not exist: {file_name}")
            return "00"

        folder = shell.NameSpace(folder_path)
        
        if folder is None:
            print(f"Error: could not open folder: {folder_path}")
            return "00"

        item = folder.ParseName(file_name)
        
        if item is None:
            print(f"Error: could not find file: {file_name} в {folder_path}")
            return "00"

        file_number = folder.GetDetailsOf(item, 26)
        if not file_number:
            print(f"Error: property 26 not found for file {file_name}")
            return "00"
        
        print(f"File: {file_name} has been assigned prefix {file_number}")
        return file_number

    def rename_files(self):
        for file_path in self.selected_files:
            file_number = self.get_file_number(file_path)
            directory, old_name = os.path.split(file_path)

            # Проверка, есть ли уже числовой префикс в имени файла
            prefix_pattern = r'^\d{3}\s+'
            match = re.match(prefix_pattern, old_name)

            if match:
                # Заменить старый префикс на новый
                new_name = re.sub(prefix_pattern, f"{file_number.zfill(3)} ", old_name)
            else:
                # Добавить новый префикс, если его не было
                new_name = f"{file_number.zfill(3)} {old_name}"

            new_path = os.path.join(directory, new_name)
            os.rename(file_path, new_path)

        self.label.config(text="Done! Select other files or a folder.")
        self.rename_button.config(state=tk.DISABLED)

    def open_manual_prefix_window(self):
        manual_prefix_window = tk.Toplevel(self.root)
        manual_prefix_window.title("Manual order")

        # Создать контейнер для кнопок и списка
        button_frame = tk.Frame(manual_prefix_window)
        button_frame.pack(pady=10, padx=10, fill=tk.X)

        self.select_file_button = tk.Button(button_frame, text="Select Files", command=self.open_files_for_prefix)
        self.select_file_button.pack(side=tk.LEFT, padx=5)

        self.add_prefix_button = tk.Button(button_frame, text="Add/Update Prefix", command=self.add_prefixes)
        self.add_prefix_button.pack(side=tk.LEFT, padx=5)

        self.change_metadata_button = tk.Button(button_frame, text="Set №", command=self.change_metadata)
        self.change_metadata_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = tk.Button(button_frame, text="Clear List", command=self.clear_list)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        self.file_listbox = tk.Listbox(manual_prefix_window)
        self.file_listbox.pack(pady=10, fill=tk.BOTH, expand=True, padx=10)

        self.file_paths = []

    def open_files_for_prefix(self):
        while True:
            file_paths = filedialog.askopenfilenames(title="Select Files")
            
            if not file_paths:
                break

            # Проверка на дублирование файлов
            for file_path in file_paths:
                if file_path not in self.file_paths:
                    self.file_paths.append(file_path)
                    self.file_listbox.insert(tk.END, os.path.basename(file_path))
                else:
                    messagebox.showwarning("Warning", f"File '{os.path.basename(file_path)}' is already in the list.")
        
    def add_prefixes(self):
        if not self.file_paths:
            messagebox.showwarning("Warning", "File list is empty.")
            return

        for index, file_path in enumerate(self.file_paths):
            directory, old_name = os.path.split(file_path)
            # Добавить новый префикс
            prefix = f"{str(index + 1).zfill(3)} "
            
            # Проверить, есть ли уже числовой префикс в имени файла
            prefix_pattern = r'^\d{3}\s+'
            match = re.match(prefix_pattern, old_name)
            
            if match:
                new_name = re.sub(prefix_pattern, prefix, old_name)
            else:
                new_name = f"{prefix}{old_name}"
            
            new_path = os.path.join(directory, new_name)
            os.rename(file_path, new_path)
            print(f"File: {old_name} has been assigned prefix {prefix}")

    def clear_list(self):
        self.file_paths.clear()
        self.file_listbox.delete(0, tk.END)

    def change_metadata(self):
        for file_path in self.file_paths:
            try:
                new_number = self.file_paths.index(file_path) + 1
                new_number_str = str(new_number).zfill(3)

                audio = ID3(file_path)
                if 'TRCK' in audio:
                    audio['TRCK'] = TRCK(encoding=3, text=new_number_str)
                else:
                    audio.add(TRCK(encoding=3, text=new_number_str))
                
                audio.save()
                print(f"Metadata successfully changed for {file_path}. New track number: {new_number}")
            except Exception as e:
                print(f"Error changing metadata for {file_path}: {str(e)}")
                traceback.print_exc()

if __name__ == '__main__':
    root = tk.Tk()
    app = Renamby(root)
    root.mainloop()
