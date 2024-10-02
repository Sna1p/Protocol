# При каждом изменении LocID кроме сброса GlobID инкрементируется
# расположение конфика "C:\Users\\AppData\Roaming\Proto\config.ini"
import configparser
import os

class CreateConfig():
    # смотрим какие папки есть в appdata
    appdata = os.listdir(os.getenv('appdata'))       
    
    if 'Proto' in appdata:
        print('Proto  IN  AppData')
    # Если нет Proto  создаем ее
    else:                                           
        print('NO')
        print('ja sozdam')
        way =  os.getenv('appdata')
        new_dir = "Proto"
        join = os.path.join(way, new_dir)
        os.mkdir(join)
        print('Done')
    def Create(self, path):
        self.config = configparser.ConfigParser()

        self.GlobID = 1  # всего = LocID
        
        self.path = path  # путь до папки в которой хрататься
        #             спрашивается только в 1-й раз дпльше бертся с config-a

        self.ProtoNumber = 0 # всего протоколов = LocNum

        # file.save(f"Протокол по ___ №{LocNum} от {CurDate}.xlsx")
        self.config['ID'] = {'GlobID': self.GlobID}
        self.config['Path'] = {'Path': self.path}
        self.config['Protocol'] = {'Number': self.ProtoNumber}

        with open("c:/Users/AppData/Roaming/Proto/config.ini", 'w') as configfile:
            self.config.write(configfile)

        
    def UpdateID(self):
        os.chdir('c:/Users/AppData/Roaming/Proto')
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        self.config.sections()
        print(self.config.sections())

        self.GlobID = int(self.config[self.config.sections()[0]]['globid'])

        self.GlobID += 1
        print('===Update===')
        print('Данные в конфиге были обновленны')

    def UpdatePN(self):
        os.chdir('c:/Users/AppData/Roaming/Proto')
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        self.config.sections()
        self.ProtoNumber = int(self.config[self.config.sections()[2]]['number'])
        print(self.config.sections())
        print(f'Номер протокола --> {self.ProtoNumber}')

        self.ProtoNumber += 1
        print('===Update===')
        print('Данные в конфиге были обновленны')    
    
    def Rewrite(self):
        self.config['ID'] = {'GlobID': self.GlobID}
        self.config['Protocol'] = {'Number': self.ProtoNumber}
        
        with open("c:/Users/AppData/Roaming/Proto/config.ini", 'w') as configfile:
            self.config.write(configfile)
        print('===Rewrite===')
        print('Данные в конфиге были перезаписаны')

#CreateConfig.Create(CreateConfig, 'c:/User/master.BASE/Desktop/Протоколы по ВУЗ')
#CreateConfig.Update(CreateConfig)
#CreateConfig.Rewrite(CreateConfig)

