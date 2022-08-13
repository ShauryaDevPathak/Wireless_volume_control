import cv2
import mediapipe as mp
import math
import numpy as np

from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

print("Master Volume level", volume.GetMasterVolumeLevel())
print('Volume Range', volume.GetVolumeRange())


mpDraw = mp.solutions.drawing_utils
mpHands = mp.solutions.hands
hands = mpHands.Hands()
cap = cv2.VideoCapture(0)

while True:
    success, img = cap.read()
    imgRGB = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    results = hands.process(imgRGB)

    if results.multi_hand_landmarks:
        for handlms in results.multi_hand_landmarks:
            lmList = [] #landmark List
            for id, lm in enumerate(handlms.landmark):
                h, w, c = img.shape
                cx, cy = int(lm.x*w), int(lm.y*h)
                lmList.append([id, cx, cy])
            if lmList:
                x1, y1 = lmList[4][1], lmList[4][2]
                x2, y2 = lmList[8][1], lmList[8][2]

                cv2.circle(img, (x1, y1), 10, (2, 6, 233), cv2.FILLED)
                cv2.circle(img, (x1, y1), 10, (2, 6, 233), cv2.FILLED)
                cv2.line(img, (x1, y1), (x2, y2), (34, 16, 123), 4)

                length = math.hypot((x2-x1), (y2-y1))
                
                volRange=volume.GetVolumeRange()
                minVol=volRange[0]
                maxVol=volRange[1]

                vol = np.interp(length, [50,300], [minVol, maxVol])
                volume.SetMasterVolumeLevel(vol, None)
                volBar = np.interp(length, [50, 300], [400, 150])
                volPer = np.interp(length, [50, 300], [0, 100])
                cv2.rectangle(img, (50, 150), (85, 400), (0, 0, 0), 3)
                cv2.rectangle(img, (50, int(volBar)), (85, 400), (0, 0, 0), cv2.FILLED)
                cv2.putText(img, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 0), 3)
    cv2.imshow("Image", img)
    cv2.waitKey(1)