import cv2
import numpy as np
from PIL import Image
import os
import win32com.client
import time

# ---------------- VOICE ----------------
speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    speaker.Speak(text, 1)

# ---------------- FACE RECOGNITION ----------------
greeted = set()          # tracks known names already greeted
unknown_zones = []       # tracks positions of unknown faces already announced
ZONE_THRESHOLD = 80      # pixel distance — faces closer than this = same person

def is_new_unknown(x, y):
    """Check if this unknown face is at a new position not seen before."""
    for (ux, uy) in unknown_zones:
        distance = ((x - ux) ** 2 + (y - uy) ** 2) ** 0.5
        if distance < ZONE_THRESHOLD:
            return False  # same unknown face position — already announced
    return True  # new position — new unknown face

def draw_boundary(img, classifier, scaleFactor, minNeighbors, color, text, clf):
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    features = classifier.detectMultiScale(gray_img, scaleFactor, minNeighbors)
    coords = []

    for (x, y, w, h) in features:
        cv2.rectangle(img, (x, y), (x + w, y + h), color, 2)
        id, pred = clf.predict(gray_img[y:y+h, x:x+w])
        confidence = int(100 * (1 - pred / 300))

        if confidence > 72:
            if id == 1:
                name = "Person-1"
            elif id == 2:
                name = "Person-2"
            elif id == 3:
                name = "Person-3"
            else:
                name = "Unknown"

            cv2.putText(img, name, (x, y - 5),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 1, cv2.LINE_AA)

            # Greet known person only once per session
            if name not in greeted:
                speak(f"Hi {name}")
                greeted.add(name)

        else:
            cv2.putText(img, "Unknown", (x, y - 5),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 1, cv2.LINE_AA)

            # Speak for every different unknown face position
            if is_new_unknown(x, y):
                speak("Unknown person detected")
                unknown_zones.append((x, y))  # remember this position

        coords = [x, y, w, h]
    return coords

def recognize(img, clf, faceCascade):
    color = (255, 255, 255)
    coords = draw_boundary(img, faceCascade, 1.1, 10, color, "Face", clf)
    return img

# ---------------- LOAD MODELS ----------------
faceCascade = cv2.CascadeClassifier("haarcascade_frontalface_default.xml")
clf = cv2.face.LBPHFaceRecognizer_create()
clf.read("classifier.xml")

# ---------------- CAMERA ----------------
video_capture = cv2.VideoCapture(0)
speak("Face detection started")

while True:
    ret, img = video_capture.read()
    if not ret:
        break

    img = recognize(img, clf, faceCascade)
    cv2.imshow("Face Detection", img)

    key = cv2.waitKey(1)
    if key == 13 or key == 27:
        break

video_capture.release()
cv2.destroyAllWindows()
speak("Face detection stopped")
