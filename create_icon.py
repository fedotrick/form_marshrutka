from PIL import Image, ImageDraw, ImageFont
import os

# Создаем директорию, если ее нет
if not os.path.exists('images'):
    os.makedirs('images')

# Создаем новое изображение с синим фоном
img = Image.new('RGBA', (256, 256), color=(1, 118, 211, 255))
draw = ImageDraw.Draw(img)

# Пытаемся загрузить шрифт для текста
try:
    font = ImageFont.truetype("arial.ttf", 100)
except IOError:
    font = ImageFont.load_default()

# Добавляем текст
draw.text((128, 128), "МК", font=font, fill=(255, 255, 255, 255), anchor="mm")

# Сохраняем в формате ICO
img.save('images/app.ico', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("Иконка создана в файле images/app.ico") 