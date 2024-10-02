import configparser
import os
from datetime import date

import pandas as pd

from Config import CreateConfig
from PUI import Ui_Proto

class XLSX():
    def createXLSX(self, columns: list, filename: str, sheet_name: str = 'Лист тест'):
        self.df = pd.DataFrame(columns=columns)
        # если нет директории создать директорию  
        if not os.path.exists('Protokol/output'):
            os.makedirs('Protokol/output')

        self.faile_path = os.path.join('Protokol/output', filename)
        
        self.Workbook = pd.ExcelWriter(self.faile_path, engine='xlsxwriter') 
        self.df.to_excel(self.Workbook, index=False, sheet_name=sheet_name) 
        # алеас на self.Workbook с заходом в property book класа ExcelWriter
        self.wb = self.Workbook.book
        # выбираем активным лист
        self.sheet = self.Workbook.sheets['Лист тест']
        # нвстраеваем размер клолонок 
        self.sheet.set_column(0, 0, 4.7)
        self.sheet.set_column(1, 1, 30.29)
        self.sheet.set_column(2, 2, 49.29)
        self.sheet.set_column(3, 3, 8.43)
        self.sheet.set_column(4, 4, 43.14)
        self.sheet.set_column(5, 5, 29.86)
        self.sheet.set_column(6, 6, 13)

        # через условное фрматирование каидаем 
        # границу на ячейки
        self.border_format = self.wb.add_format({'border': 1}) 
        self.no_border_format = self.wb.add_format({'border': 0}) 
        self.sheet.conditional_format('A1:C9999', {'type': 'cell', 'criteria': '!=', 'value': '"aaaaaaaaaaaaaaaaaaaaaaaa"', 'format': self.border_format})
        self.sheet.conditional_format('E1:G9999', {'type': 'cell', 'criteria': '!=', 'value': '"aaaaaaaaaaaaaaaaaaaaaaaa"', 'format': self.border_format})
        
        # настойка текста заголовка
        self.had_form = self.wb.add_format({'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'font': 'Times New Roman'})
        for col_num, val in enumerate(self.df.columns.values):
            self.sheet.write(0, col_num, val, self.had_form)
        
        # подпись заголовков
        self.index = [1, 2, 3, '', 4, 5, 6]
        self.col = 0
        self.index_form = self.wb.add_format({'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'font': 'Times New Roman', 'size': '12'})
        for col_num in (self.index):
            self.sheet.write(1, self.col, self.index[self.col], self.index_form)
            self.col += 1
        return self.faile_path
    
    # заводим temp для временного хранения LocID
    # в массиве 
    temp = []
    def WriteXLSX(self, G, row: int):
        # LocID/GlobID
        # GlobID из config файла
        # При добавлении человека увеличиваеися а при закрытии программы
        # сбрасывается 
        # условя для привоения LocId
        # Если temp пусто единица
        # иначе длинна temp + 1 
        if  self.temp == []:
            self.LocId = 1
            self.temp.append(self.LocId)
        elif self.temp != []:
            self.LocId = len( self.temp) + 1
            self.temp.append(self.LocId)
        
        self.text_form = self.wb.add_format({'bold': False, 'align': 'center', 'font': 'Times New Roman', 'size': '12'})
        self.sheet.write(row, 0, f'{self.LocId}/{G}', self.text_form)
        return self.LocId
    
    def SaveXLSX(self):
        self.Workbook._save()
        #self.temp.clear()

    def CurentDate(self):
        # берём текющую дату
        self.curdate = date.today()
        # переводим <datetime.date> в str
        self.Date_Str = date.strftime(self.curdate, '%Y-%m-%d')
        # сплитаем str по разделителю
        self.Strdate = str.split(self.Date_Str, '-')
        self.Date = []
        i = 2
        # кидаем в масиив Date с конца
        # так как по усолчанию маска datetime.date 
        # 'Y-m-d
        while i >= 0:
            if i != 0:
                self.Date.append(self.Strdate[i])
            i -= 1
        return self.Date

    def Read_Config(self):
        # меняем директорию на %AppData%/Proto
        os.chdir('c:/Users/AppData/Roaming/Proto')
        # читаем config
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        return self.config

    def Get_Value(self, read):
        # с полдученного в Read_Config  смотрим разделы
        # ['section 0', 'section 1', 'section 2']  
        read.sections()
        # получаем номер протокола
        self.ProtoNumber = read[read.sections()[2]]['number']
        # получаем Id
        self.GlobId = read[read.sections()[0]]['globid']
        print('def Get_Value')
        print(read.sections())
        print(f'Номер протокола --> {self.ProtoNumber}')
        os.chdir('c:/Users/Desktop/Trust/')
        return self.ProtoNumber, self.GlobId

    # Все равно создаётся 2 протокола
    # должен обнавляься первый а не создаваться второй
    #   

    def Update(self):
        # обновляем и перезаписываем config
        # если LocId = 1 ProtoNumber + 1
        if self.LocId == 1:
            CreateConfig.UpdatePN(CreateConfig)
        CreateConfig.UpdateID(CreateConfig)
        CreateConfig.Rewrite(CreateConfig)

    def Print_Setting(self):
        # размер области печати
        self.sheet.print_area('A1:G9999')
        # порядок страниц
        self.sheet.print_across()
        # размер страницы на листе
        self.sheet.set_print_scale(94)
