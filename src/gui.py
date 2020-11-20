from tkinter import *
from Dialog import *
from Lables import *
from Script import Script

# Создаем окно с необходимыми параметрами
window = Tk()
window.title('Преобразователь недельных стоков X5')
window.geometry("500x300")
window.resizable(width=False, height=False)
window.configure(background='#2e8b57')

# Создаем экемпляр класса диалогового окна
dialog = Dialog()
# Создаем экземпляр лейбла
lb = Lables()


# Функция, определяющая адреса файлов, пароль и запускающая процесс блокировки файлов
def runConverting():
    tup = dialog.getPaths()
    folder = dialog.getFolderName()

    # Удаляем сообщения если есть
    lb.clearMessages()
    lb.destroyLables()

    # Проверяем выбран ли файлы
    if len(tup) == 0:
        return messagebox.showerror("Ошибка", "Необходимо выбрать файлы")

    script = Script(tup, folder)
    filename = script.run()

    # Удаляем сообщения об успешном выполнении и выводим сообщение об успешном завершении процесса
    lb.destroyLables()
    lb.addMessageToList(f'{filename} готов.\n Если в консоле нет ошибок, то все ок.', 6)
    lb.createLables(background='#2e8b57')
        

# Инициализация кнопки выбора файлов
chooseFileBtn = Button(window, text="Выбрать файлы для обработки", command=dialog.callDialog, width=25, height=1)
chooseFileBtn.pack(side=TOP, padx=5, pady=5)

# Инициализация кнопки запуска
showFileBtn = Button(window, text="Запустить обработку файлов", command=runConverting, width=25, height=1)
showFileBtn.pack(padx=5, pady=5)

window.mainloop()
