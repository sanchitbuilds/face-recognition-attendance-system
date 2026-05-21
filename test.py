from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch


def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(text)


# ================= CAMERA & FACE DETECTOR =================
video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')


# ================= LOAD DATA =================
with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)

with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print("Shape of Faces matrix -->", FACES.shape)
print("FACES shape:", FACES.shape)
print("LABELS length:", len(LABELS))


# ================= FIX MISMATCH =================
# Ensure one label per face
min_len = min(FACES.shape[0], len(LABELS))

print(f"⚠ Aligning data to {min_len} samples")

FACES = FACES[:min_len]
LABELS = LABELS[:min_len]



# ================= TRAIN MODEL =================
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)


# ================= UI BACKGROUND =================
imgBackground = cv2.imread("Background Image.png")
COL_NAMES = ['NAME', 'TIME']


# ================= MAIN LOOP =================
while True:
    ret, frame = video.read()
    if not ret:
        break

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    attendance = None

    for (x, y, w, h) in faces:
        crop_img = frame[y:y + h, x:x + w]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)

        distances, indices = knn.kneighbors(resized_img, n_neighbors=1)
        distance = distances[0][0]

        THRESHOLD = 4000  # you can tune this

        if distance > THRESHOLD:
            output = "Unknown"
        else:
            output = knn.predict(resized_img)[0]


        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")

        attendance = [str(output), str(timestamp)]

        # Draw UI
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, output, (x, y - 15),
                    cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)

    cv2.imshow("Frame", frame)


    k = cv2.waitKey(1)

    if k == ord('o') and attendance is not None:
        speak("Attendance Taken")
        time.sleep(1)

        filename = f"Attendance/Attendance_{date}.csv"
        file_exists = os.path.isfile(filename)

        with open(filename, "a", newline="") as csvfile:
            writer = csv.writer(csvfile)
            if not file_exists:
                writer.writerow(COL_NAMES)
            writer.writerow(attendance)

    if k == ord('q'):
        break


video.release()
cv2.destroyAllWindows()
