import math


# создается два класса. Класс окон/дверей и клас обоев.
# оба класса возвращают площадь этих объектов
class WinDoor:
    def __init__(self, len_windoor, height_windoor):
        self.sq_windoor = len_windoor * height_windoor


class Wallpaper:
    def __init__(self, width_wallpaper, len_wallpaper):
        self.sq_wallpapper = width_wallpaper * len_wallpaper


# класс Room — это класс-контейнер,который вызывает два ранее созданных классах
class Room:
    def __init__(self, width_room, len_room, height_room):
        self.width = width_room
        self.length = len_room
        self.height = height_room
        self.windoor = []
        self.sq = 0
        self.new_sq = 0
        self.res = 0

    # метод расчета площади всех стен комнта
    def square(self):
        self.sq = round(2 * self.height * (self.width + self.length), 2)
        return print(f'Площадь оклеиваемой поверхности без учета '
                     f'окон и дверей равна {self.sq}')

    # метод добавляющий площадь дверей окон в список
    def add_windoor(self, width_wallpaper, len_wallpaper):
        self.windoor.append(WinDoor(width_wallpaper, len_wallpaper))

    # метод корректирующий площадь комнаты на площади окон дверей
    def work_square(self):
        self.new_sq = self.sq
        for i in self.windoor:
            self.new_sq -= i.sq_windoor
        return self.new_sq, print(f'Площадь оклеиваемой поверхности c учетом '
                                  f'окон и дверей равна {self.new_sq}')

    # метод расчета необходимого количества рулонов обоев
    def count_pepper(self):
        self.res = math.ceil(self.new_sq / c.sq_wallpapper)
        return print(f'Необходимо {self.res} рулона(ов)')


# интерфейс для ввода данных о комнате
x = float(input('Введите ширину '))
y = float(input('Введите длину '))
h = float(input('Введите высоту '))

a = Room(x, y, h)
Room.square(a)

# интерфейс ввода и рассчета окон и дверей
count_windoor = int(input('Введите суммарное количетво окон и дверей '))

for num_windoor in range(count_windoor):
    z = float(input(f'Введите ширину окна/двери №{num_windoor+1} '))
    h = float(input(f'Введите высоту окна/двери №{num_windoor+1} '))
    Room.add_windoor(a, z, h)

Room.work_square(a)

# интерфейс ввода информации о рулоне обоев
t = float(input('Введите ширину рулона '))
j = float(input('Введите длину рулона '))

c = Wallpaper(t, j)
Room.count_pepper(a)
