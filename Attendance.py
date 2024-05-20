from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch
from sklearn.neighbors import KNeighborsClassifier





def speak(str1):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str1)


video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier(r"C:\Users\Vandhana\pythonProject\haarcascade_frontalface_default .xml")

with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

imgBackground = cv2.imread("background.png")

COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    for (x, y, w, h) in faces:
        crop_img = frame[y:y + h, x:x + w, :]
        # Resize the cropped image to 50x75 pixels
        resized_gray_img = cv2.resize(crop_img, (75, 50))
        # Convert the resized image to grayscale
        resized_gray_img = cv2.cvtColor(resized_gray_img, cv2.COLOR_BGR2GRAY)
        # Flatten the grayscale image to have 3750 features
        resized_gray_img = resized_gray_img.flatten().reshape(1, -1)
        output = knn.predict(resized_gray_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        attendance = [str(output[0]), str(timestamp)]

        if imgBackground is not None:
            imgBackground[162:162 + 480, 55:55 + 640] = frame
            cv2.imshow("Frame", imgBackground)
        else:
            cv2.imshow("Frame", frame)

        k = cv2.waitKey(1)
        if k == ord('o'):
            speak("Attendance Taken..")
            time.sleep(5)
            attendance_file = f"Attendance/Attendance_{date}.csv"
            if exist:
                with open(attendance_file, "a", newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(attendance)
            else:
                with open(attendance_file, "w", newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    writer.writerow(attendance)
        if k == ord('q'):
            break

video.release()
cv2.destroyAllWindows()


