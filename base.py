from PIL import Image, ImageDraw, ImageFont
import xlrd
import os

path = "Book1.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)
rows = inputWorksheet.nrows

user = []
objects = {}
for i in range(rows):
    objects['name'] = inputWorksheet.cell_value(i, 1)
    objects['team'] = inputWorksheet.cell_value(i, 0)
    user.append(objects)
    objects = {}


for person in user:
    image = Image.open('certificate.jpg')
    name = person['name']
    team = person['team']
    font = ImageFont.truetype(f'{os.getcwd()}/Roboto-Bold.ttf', size=40)

    draw = ImageDraw.Draw(image)
    draw.text(xy=(1280, 608), text=name, fill=(0, 0, 0), font=font)
    draw.text(xy=(1310, 770), text=team, fill=(0,0,0), font=font)
    image.save(f'{team}_{name}.jpg')
    print(f'Generated for {team}_{name}')