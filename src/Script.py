# Импортируем библиотеки
import pandas as pd
import os
import datetime

class Script():
    def __init__(self, path_list, folder_name):
        # Список путей с файлами для обработки
        self.file_paths = path_list
        # Список фреймов, которые надо будет соединить между собой
        self.files = []
        # Финальный Фрейм, в котором соединяются фреймы файлов, и который надо будет сохранить
        self.df = pd.DataFrame()
        # Имя папки ,где лежат файлы, и в которую надо сохранить итоговый файл
        self.folder_name = folder_name

    def run(self):
        """Запускает скрипт. Обрабатывает все файлы, которые лежат в self.files, соединяет их и сохраняет"""
        # Перебираем все файлы в списке и обрабатываем их
        for file in self.file_paths:
            self._processFile(file)
        # Соединяем обработанные файлы
        self._appendFrames()
        # Сохраняем итоговый файл и возвращаем его имя
        filename = self._saveFrame()
        return filename

    def _processFile(self, excel_file_path):
        """Обрабатывает файл, чтобы он был в нужном формате"""
        path_to_file = os.path.normpath(excel_file_path)
        df = pd.read_excel(path_to_file, dtype={'ШК': str})
        # Если первая строчка содержит сумму объемов, как это обычно бывает, то удаляем ее
        if len(df.columns) - df.iloc[0].isnull().sum() == 1:
            df.drop(0, inplace=True)
        # Сортируем в нужном порядке
        df.sort_values(
            by=[
            'SAP код склада', 
            'ШК', 
            'Неделя планируемого заказа', 
            'Неделя составления прогноза', 
            'Запланированная дата поставки'], 
            inplace=True
        )
        # Удаляем старые прогнозы
        df.drop_duplicates(
            subset=[
                'SAP код склада', 
                'ШК', 
                'Неделя планируемого заказа', 
                'Запланированная дата поставки'
                ], 
                keep='last', 
                inplace=True
        )
        # Добавляем фрейм в список фреймов, которые потом надо будет соединить
        self.files.append(df)

    def _appendFrames(self):
        """Соединяет  фреймы из self.files"""
        df = pd.DataFrame()
        for frame in self.files:
            df = df.append(frame, ignore_index=True)

        self.df = df

    def _saveFrame(self):
        """Сохраняет итоговый фрейм self.df"""
        week = str(datetime.datetime.now().isocalendar()[1])
        year = str(datetime.datetime.now().year)[2:]
        # Сохраняем файл
        fileName = f"X5_ALL_{week + year}.xlsx"
        whereToPut = os.path.join(self.folder_name, fileName)
        self.df.to_excel(whereToPut, index=False)

        return fileName


