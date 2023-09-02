import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO

from openpyxl.reader.excel import load_workbook


def image_to_excel(pictures, name):
    w = load_workbook(f'./docs/{name}.xlsx')
    sheet = workbook.active

    # Получаем содержимое картинки
    response = requests.get(pictures[0])
    image_content = response.content

    # Создаем объект изображения PIL
    pil_image = PILImage.open(BytesIO(image_content))

    # Устанавливаем размер ячейки для изображения
    ws.column_dimensions['A'].width = 20

    # Создаем объект изображения openpyxl
    img = Image(BytesIO(image_content))

    # Масштабируем изображение, чтобы оно поместилось в ячейку
    img.width = 100
    img.height = 100

    # Вставляем изображение в ячейку A1
    ws.add_image(img, 'A1')

    # Сохраняем рабочую книгу
    wb.save('example.xlsx')
