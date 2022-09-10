from win32gui import GetWindowText, GetForegroundWindow
import ctypes, time

# получаем значение скорости прокрутки по умолчанию
default_wheel_speed = ctypes.c_int()
ctypes.windll.user32.SystemParametersInfoA(104, 0, ctypes.byref(default_wheel_speed), 0)
default_wheel_speed = default_wheel_speed.value

# устанавливаем текущее значение скорости прокрутки
cur_wheel_speed = default_wheel_speed

# проверяем что окно которое сейчас в фокусе это Excel
# ставим скорость по умолчанию или равной 1 в зависимости от текушей скорости и наличия окна экселя в фокусе
while True:
    in_excel = str(GetWindowText(GetForegroundWindow()))[-5:] == 'Excel'

    if cur_wheel_speed == default_wheel_speed and in_excel:
        ctypes.windll.user32.SystemParametersInfoA(105, 1, 0, 3)
        cur_wheel_speed = 1

    elif cur_wheel_speed == 1 and not in_excel:
        ctypes.windll.user32.SystemParametersInfoA(105, default_wheel_speed, 0, 3)
        cur_wheel_speed = default_wheel_speed

    time.sleep(1 / 30)
