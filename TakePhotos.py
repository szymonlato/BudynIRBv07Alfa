import os
import numpy as np
import pandas as pd
import cv2 as cv
from datetime import datetime

if os.path.exists('id_names.csv'):
    id_names = pd.read_csv('id_names.csv')
    id_names = id_names[['id', 'name']]
else:
    id_names = pd.DataFrame(columns=['id', 'name'])
    id_names.to_csv('id_names.csv')

if not os.path.exists('faces'):
    os.makedirs('faces')


id = int(input('Podaj swoje id: '))
name = ''

if id in id_names['id'].values:
    name = id_names[id_names['id'] == id]['name'].item()
    print(f'Dzien dobry {name}!!')
else:
    name = input('Wpisz swoje imie: ')
    os.makedirs(f'faces/{id}')
    id_names = id_names.append({'id': id, 'name': name}, ignore_index=True)
    id_names.to_csv('id_names.csv')


camera = cv.VideoCapture(0)
face_classifier = cv.CascadeClassifier('Classifiers/haarface.xml')

photos_taken = 0

while(cv.waitKey(1) & 0xFF != ord('q')):
    _, img = camera.read()
    grey = cv.cvtColor(img, cv.COLOR_BGR2GRAY)
    faces = face_classifier.detectMultiScale(grey, scaleFactor=1.1, minNeighbors=5, minSize=(50, 50))
    for (x, y, w, h) in faces:
        cv.rectangle(img, (x, y), (x + w, y + h), (0, 0, 255), 2)

        face_region = grey[y:y + h, x:x + w]
        if cv.waitKey(1) & 0xFF == ord('s') and np.average(face_region) > 50:
            face_img = cv.resize(face_region, (220, 220))
            img_name = f'face.{id}.{datetime.now().microsecond}.jpeg'
            cv.imwrite(f'faces/{id}/{img_name}', face_img)
            photos_taken += 1
            print(f'Zdjecie nr: {photos_taken}')

    cv.imshow('Face', img)

camera.release()
cv.destroyAllWindows()
