import json
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import numpy as np
import cv2
import os
import openpyxl
import pandas as pd

filename = input("파일명:")
fdname = input("폴더명: ")
book = openpyxl.load_workbook(filename+'.xlsx')

if not os.path.isdir(os.getcwd()+'/'+fdname):
    os.mkdir(os.getcwd()+'/'+fdname)


fontsFolder = 'C:/Windows/Fonts'
selectedFont = ImageFont.truetype(os.path.join(fontsFolder,'Arial.ttf'),35)

sheet = book.worksheets[0]

#image = cv2.imread(os.getcwd()+'/'+imgfn)



data = []
for row in sheet.rows:
    data.append([row[0].value, row[1].value])
del data[0]


datadots = []
for k in data:
    if str(k[0]) != '\\N':
        input=str(k[0])
        datadots.append(json.loads(input))
    else:
        input='None'
        datadots.append(input)


lenn = len(datadots)

i=0
j=0


max_all = len(datadots)


while j < max_all:
    bg = Image.open(os.getcwd()+'/'+str(data[j][1]))
    if datadots[j] != 'None':
        while i < (len(datadots[j]['data'])):
            if int(datadots[j]['data'][i]['dots'][0]['x']) < int(datadots[j]['data'][i]['dots'][2]['x']):
                w = int(datadots[j]['data'][i]['dots'][0]['x'])
            else:
                w = int(datadots[j]['data'][i]['dots'][2]['x'])

            if int(datadots[j]['data'][i]['dots'][1]['y']) < int(datadots[j]['data'][i]['dots'][3]['y']):
                h = int(datadots[j]['data'][i]['dots'][1]['y'])
            else:
                h = int(datadots[j]['data'][i]['dots'][3]['y'])

            exam = Image.open(os.getcwd()+'/'+str(datadots[j]['data'][i]['data'][1]['value'])+'.png')
            bg.paste(exam,(w,h))
            bg.save(os.getcwd()+'/'+fdname+'/bounding_'+str(data[j][1]))
#            print(str(i)+' over '+str(len(datadots[j]['data'])))
            i+=1

    else:
        continue

    i=0
    j+=1

i=0
j=0

while j < max_all:
    image = cv2.imread(os.getcwd()+'/'+fdname+'/bounding_'+str(data[j][1]))
    if datadots[j] != 'None':
        while i < (len(datadots[j]['data'])):
            cv2.line(image, (int(datadots[j]['data'][i]['dots'][0]['x']) , int(datadots[j]['data'][i]['dots'][0]['y'])), (int(datadots[j]['data'][i]['dots'][1]['x']),int(datadots[j]['data'][i]['dots'][1]['y'])), (0,255,0), 2)
            cv2.line(image, (int(datadots[j]['data'][i]['dots'][1]['x']) , int(datadots[j]['data'][i]['dots'][1]['y'])), (int(datadots[j]['data'][i]['dots'][2]['x']),int(datadots[j]['data'][i]['dots'][2]['y'])), (0,255,0), 2)
            cv2.line(image, (int(datadots[j]['data'][i]['dots'][2]['x']) , int(datadots[j]['data'][i]['dots'][2]['y'])), (int(datadots[j]['data'][i]['dots'][3]['x']),int(datadots[j]['data'][i]['dots'][3]['y'])), (0,255,0), 2)
            cv2.line(image, (int(datadots[j]['data'][i]['dots'][3]['x']) , int(datadots[j]['data'][i]['dots'][3]['y'])), (int(datadots[j]['data'][i]['dots'][0]['x']),int(datadots[j]['data'][i]['dots'][0]['y'])), (0,255,0), 2)

            if int(datadots[j]['data'][i]['dots'][0]['x']) < int(datadots[j]['data'][i]['dots'][2]['x']):
                w = int(datadots[j]['data'][i]['dots'][0]['x'])
            else:
                w = int(datadots[j]['data'][i]['dots'][2]['x'])

            if int(datadots[j]['data'][i]['dots'][1]['y']) < int(datadots[j]['data'][i]['dots'][3]['y']):
                h = int(datadots[j]['data'][i]['dots'][1]['y'])
            else:
                h = int(datadots[j]['data'][i]['dots'][3]['y'])


            if datadots[j]['data'][i]['data'][0]['value'] == 1 :
                cv2.putText(image, 'back', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            elif datadots[j]['data'][i]['data'][0]['value'] == 2 :
                cv2.putText(image, 'left', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            elif datadots[j]['data'][i]['data'][0]['value'] == 3 :
                cv2.putText(image, 'right', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            elif datadots[j]['data'][i]['data'][0]['value'] == 4 :
                cv2.putText(image, 'top', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            elif datadots[j]['data'][i]['data'][0]['value'] == 5 :
                cv2.putText(image, 'bottom', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            elif datadots[j]['data'][i]['data'][0]['value'] == 6 :
                cv2.putText(image, 'laydown', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            elif datadots[j]['data'][i]['data'][0]['value'] == 7 :
                cv2.putText(image, 'standup', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            elif datadots[j]['data'][i]['data'][0]['value'] == 8 :
                cv2.putText(image, 'front', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            else :
                cv2.putText(image, 'None', (w,h), cv2.FONT_HERSHEY_DUPLEX, 0.5, (0,0,0), 1)
            cv2.imwrite(os.getcwd()+'/'+fdname+'/bounding_'+str(data[j][1]),image)
#            print(str(i)+' over '+str(max_all))
            i+=1

    else:
        cv2.imwrite(os.getcwd()+'/'+fdname+'/bounding_'+str(data[j][1]),image)

    i=0
    j+=1


print('complited!!')
os.system("pause")
