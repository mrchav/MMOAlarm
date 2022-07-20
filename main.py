from time import sleep
import numpy as np
import datetime

import cv2
import win32gui
import win32ui
import win32con
import pyttsx3
#import pytesseract


engine = pyttsx3.init()
#pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

#
#Данный клас отвечает за информацию о чарах их локации, текущем состоянии
#
class EveChars():
    def __init__(self,windowhwnd,WindowText,winname):
        self.charname = WindowText.split(' - ')[1] #из названия окна понимаем что это за чар
        self.winname = winname #название самого окна чара
        self.windowhwnd = windowhwnd #hwnd окна чара
        self.getWindow_W_H() #размеры окна чара
        self.nextcheck = datetime.datetime.now() #время следующего обновления врагов чара
        #self.system = self.getCharSystem() # в какой системе чар

    #получаем основные размеры окна клиента
    def getWindow_W_H(self):
        self.left, self.top, self.right, self.bot = win32gui.GetWindowRect(self.windowhwnd)
        #width = right - left - 15
        # height = bot - top - 11
        self.width = self.right - self.left
        self.height = self.bot - self.top
        return True

    #скриншот игрового экрана чара
    def getScreenData(self):
        # Замените hwnd на WindowLong
        s = win32gui.GetWindowLong(self.windowhwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(self.windowhwnd, win32con.GWL_EXSTYLE, s | win32con.WS_EX_LAYERED)

        # Определите, свернуто ли окно
        show = win32gui.IsIconic(self.windowhwnd)

        # Измените атрибут слоя окна на прозрачный
        # Восстановите окно и увеличьте масштаб вперед
        # Отменить максимальную анимацию минимизации
        if show == 1:
            win32gui.SystemParametersInfo(win32con.SPI_SETANIMATION, 0)
            win32gui.SetLayeredWindowAttributes(self.windowhwnd, 0, 0, win32con.LWA_ALPHA)
            win32gui.ShowWindow(self.windowhwnd, win32con.SW_RESTORE)

            # Создать выходной слой
        try:
            hwindc = win32gui.GetWindowDC(self.windowhwnd)
        except:
            print(f'Задача {self.windowhwnd} уже не выполняется !!')
            return []

        srcdc = win32ui.CreateDCFromHandle(hwindc)
        memdc = srcdc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()

        # Скопируйте целевой слой и вставьте его в bmp.
        bmp.CreateCompatibleBitmap(srcdc, self.width, self.height)
        memdc.SelectObject(bmp)
        memdc.BitBlt((0, 0), (self.width, self.height), srcdc, (8, 3), win32con.SRCCOPY)

        # Преобразовать растровое изображение в np
        signedIntsArray = bmp.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape = (self.height, self.width, 4)

        # Освободить содержимое устройства
        srcdc.DeleteDC()
        memdc.DeleteDC()
        win32gui.ReleaseDC(self.windowhwnd, hwindc)
        win32gui.DeleteObject(bmp.GetHandle())

        # Восстановить целевые атрибуты
        if show == 1:
            win32gui.SetLayeredWindowAttributes(self.windowhwnd, 0, 255, win32con.LWA_ALPHA)
            win32gui.SystemParametersInfo(win32con.SPI_SETANIMATION, 1)
        # Вернуться изображение
        return img

    #ищем врагов в локальном чате
    def searchEnemy(self):
        method = cv2.TM_SQDIFF_NORMED
        ressult_search = []
        #
        #Нам надо найти 4 варианта иконки врагов, которые нарезаем с одной картинки
        #
        for jj in range(0,4):
            xx = jj * 13
            yy = xx + 13
            target_img = cv2.cvtColor(cv2.imread('alarmpicture2.bmp'), cv2.COLOR_BGR2RGB)[0:13, xx:yy]
            # cv2.imshow("Image", target_img)
            # cv2.waitKey(0)

            #скрин экрана клиента переводим в RBG формат и обрезаем
            large_image = cv2.cvtColor(self.getScreenData(), cv2.COLOR_BGR2RGB)
            large_image = large_image[100:self.height,
                          0:0 + int(self.width-200)]

            #Результат поиска таргет картинки на скрине экрана, записываем все результаты в массив
            result = cv2.matchTemplate(target_img, large_image, method)
            mn = cv2.minMaxLoc(result)
            ressult_search.append(round(mn[0], 4))
        print(ressult_search)

        #Минимальное значение будет указывать на лучшее совпадение поиска картинки, выбираем его.
        most_similar_enemy = min(ressult_search)

        # опытным путем определили значение 0.11, при нем практически нет ложных срабатываний и всегда видим врага
        if most_similar_enemy < 0.11:
            print(f'{self.charname} есть вары так как  {most_similar_enemy} ')
            return True
        else:
            print(f'{self.charname} __варов нет так как  совпадение {most_similar_enemy}')
            return False

    #Проверяем в космосе ли пилот или в доке
    def getLocation(self):
        method = cv2.TM_SQDIFF_NORMED

        arm_img = cv2.imread('shiel_armor.bmp')
        small_image = cv2.cvtColor(arm_img, cv2.COLOR_BGR2RGB)

        #Eсли у корабля есть панель защиты, значит он в космосе, если же нет, то в доке.
        #Переводим картинку в RGB и определяем в какой области искать.
        large_image = cv2.cvtColor(self.getScreenData(), cv2.COLOR_BGR2RGB)
        screen = large_image[int(self.height * 0.6):self.height,
                      int(self.width * 0.2): int(self.width * 0.8)]
        #cv2.imshow("Image", screen)
        #cv2.waitKey(0)

        #Сравниваем картинки
        result = cv2.matchTemplate(small_image, screen, method)

        mn = cv2.minMaxLoc(result)

        #Определяем на сколько совпали картинки и присваиваем соотвествующий статус.
        if mn[0] < 0.42:
             self.location = 'space'
             print(f'{self.charname} расположение {self.location} так как {mn[0]} <0.42')
        else:
             self.location = 'station'
             print(f'{self.charname} расположение {self.location} так как {mn[0]} >0.42')
        return self.location

    # Определяем в какой системе находится чар
    # def getCharSystem(self):
    #     # Делаем скрин экрана и переводим его в RBG.
    #     # Вырезаем кусок скрина, в верхней части экрана, где пишится система.
    #     image = cv2.cvtColor(self.getScreenData(), cv2.COLOR_BGR2RGB)
    #     image = image[90:130,
    #                   100:160]
    #     # Распознаем надпись с картинки в текст.
    #     self.system = pytesseract.image_to_string(image)
    #     return self.system

    # Устанавливаем дате следующей проверки локала чара.
    def setCharNextCheck(self, min = 1):
        self.nextcheck = datetime.datetime.now() + datetime.timedelta(minutes=min)
    # Обновляем дату о чаре.
    def updateData(self):
        self.getLocation()
        #self.getCharSystem()

# Ники всех активных чаров переводим в список.
def allActiveChars(chars):
    char_list = list()
    for ch in chars:
        char_list.append(ch.charname)
    return char_list

# Функция ищет все окна с запущенной игрой.
def winEnumHandler(hwnd, ctx ):
   if win32gui.IsWindowVisible(hwnd):
        # Ищим все поля с названием игры
        if win32gui.GetWindowText(hwnd).find('EVE') >= 0:
            left, top, right, bottom = win32gui.GetClientRect(hwnd)
            w = right - left
            # Оставляем только окна больших размеров и в которых есть через дефис имя чара.
            # Также берем только окна, которых еще нет в списке
            if (w > 600) and len(win32gui.GetWindowText(hwnd).split(' - ')) > 1:
                if win32gui.GetWindowText(hwnd).split(' - ')[1] not in allActiveChars(chars):
                    chars.append(EveChars(hwnd, win32gui.GetWindowText(hwnd), win32gui.GetWindowText(hwnd)))

# Задаем время через которое обновлять открытые окна
def timeToCheckActiveWindows(min=1):
    return datetime.datetime.now() + datetime.timedelta(minutes=min)

if __name__ == '__main__':
    chars = []
    nexttimecheckwin = datetime.datetime.now()

    # Проверяем все открытые окна на комьютере и ищем нужные нам.
    # из каждого окна достаем название и его hwnd.
    # формируем из полученных данных объекты активных чаров
    win32gui.EnumWindows(winEnumHandler, None)
    while True:

        # Если время пришло, проверям есть ли новые активные окна с чарами.
        if nexttimecheckwin < datetime.datetime.now():
            print(f'после проверки {nexttimecheckwin} < {datetime.datetime.now()}')
            nexttimecheckwin = timeToCheckActiveWindows()
            win32gui.EnumWindows(winEnumHandler, None)

        for char in chars:                                  # Для каждого активного чара
            if char.nextcheck < datetime.datetime.now():    # Если пришло время для обновления
                char.updateData()                           # Обновляем данные чара
                if char.location == 'space':                # Если чар сейчас в космосе
                    if char.searchEnemy():                  # Проверяем есть ли враги в локале
                        char.setCharNextCheck()             # Обновляем дату следующей проверки данно чара,
                                                            # что слишком часть не мучал
                        print(f'\n{char.charname} =  ALARM!!!\n')                 # Выводим в консоль чара
                        engine.say(f'Осторожно {char.charname} в опасности, повторяю - '   # Записываем глосовую фразу 
                                   f'осторожно {char.charname} в опасности!')
                    engine.runAndWait()                                                     # проигрываем голос
        sleep(1)      # Задержка основного цикла. Можно без задержки, но может подтормаживать



