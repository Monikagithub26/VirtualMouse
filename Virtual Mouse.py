import cv2
from PIL import ImageGrab
import time
import subprocess
import platform
import wmi
import numpy as np
import HandTracking as ht
from cvzone.HandTrackingModule import HandDetector
import mediapipe as mp
import pyautogui
import math
from enum import IntEnum
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from google.protobuf.json_format import MessageToDict # Used to convert protobuf message to a dictionary.
import screen_brightness_control as sbcontrol
import pyttsx3  #text to speech conversion library in python
import datetime
import speech_recognition as sr
import pyaudio
import wikipedia
import webbrowser
import os


pyautogui.FAILSAFE = False
mp_drawing = mp.solutions.drawing_utils
mp_hands = mp.solutions.hands


# Gesture Encodings
class Gest(IntEnum):
    # Binary Encoded
    FIST = 0
    PINKY = 1
    RING = 2
    MID = 4
    LAST3 = 7
    INDEX = 8
    FIRST2 = 12
    LAST4 = 15
    THUMB = 16
    PALM = 31

    # Extra Mappings
    V_GEST = 33
    TWO_FINGER_CLOSED = 34
    PINCH_MAJOR = 35
    PINCH_MINOR = 36
    #FINGER_CROSSED = 37

    def ring():
        touchpad_x = 500
        touchpad_y = 500
        ring_finger_x = 600
        ring_finger_y = 600

        # Continuously check the position of the ring finger and simulate a left-click when tapped
        while True:
            x, y = pyautogui.position()
            if x == ring_finger_x and y == ring_finger_y:
                pyautogui.click(touchpad_x, touchpad_y)




# Multi-handedness Labels
class HLabel(IntEnum):
    MINOR = 0
    MAJOR = 1


# Convert Mediapipe Landmarks to recognizable Gestures
class HandRecog:
    def __init__(self, hand_label):

        self.finger = 0
        self.ori_gesture = Gest.PALM
        self.prev_gesture = Gest.PALM
        self.frame_count = 0
        self.hand_result = None
        self.hand_label = hand_label

    def update_hand_result(self, hand_result):
        self.hand_result = hand_result

    def get_signed_dist(self, point):
        sign = -1
        if self.hand_result.landmark[point[0]].y < self.hand_result.landmark[point[1]].y:
            sign = 1
        dist = (self.hand_result.landmark[point[0]].x - self.hand_result.landmark[point[1]].x) ** 2
        dist += (self.hand_result.landmark[point[0]].y - self.hand_result.landmark[point[1]].y) ** 2
        dist = math.sqrt(dist)
        return dist * sign

    def get_dist(self, point):
        dist = (self.hand_result.landmark[point[0]].x - self.hand_result.landmark[point[1]].x) ** 2
        dist += (self.hand_result.landmark[point[0]].y - self.hand_result.landmark[point[1]].y) ** 2
        dist = math.sqrt(dist)
        return dist

    def get_dz(self, point):
        return abs(self.hand_result.landmark[point[0]].z - self.hand_result.landmark[point[1]].z)

    # Function to find Gesture Encoding using current finger_state.
    # Finger_state: 1 if finger is open, else 0
    def set_finger_state(self):
        if self.hand_result == None:
            return

        points = [[8, 5, 0], [12, 9, 0], [16, 13, 0], [20, 17, 0]]
        self.finger = 0
        self.finger = self.finger | 0  # thumb
        for idx, point in enumerate(points):

            dist = self.get_signed_dist(point[:2])
            dist2 = self.get_signed_dist(point[1:])

            try:
                ratio = round(dist / dist2, 1)
            except:
                ratio = round(dist1 / 0.01, 1)

            self.finger = self.finger << 1
            if ratio > 0.5:
                self.finger = self.finger | 1

    # Handling Fluctations due to noise
    def get_gesture(self):
        if self.hand_result == None:
            return Gest.PALM

        current_gesture = Gest.PALM
        if self.finger in [Gest.LAST3, Gest.LAST4] and self.get_dist([8, 4]) < 0.05:
            if self.hand_label == HLabel.MINOR:
                current_gesture = Gest.PINCH_MINOR
            else:
                current_gesture = Gest.PINCH_MAJOR

        elif Gest.FIRST2 == self.finger:
            point = [[8, 12], [5, 9]]
            dist1 = self.get_dist(point[0])
            dist2 = self.get_dist(point[1])
            ratio = dist1 / dist2
            if ratio > 1.7:
                current_gesture = Gest.V_GEST
            else:
                if self.get_dz([8, 12]) < 0.1:
                    current_gesture = Gest.TWO_FINGER_CLOSED
                else:
                    current_gesture = Gest.MID

        else:
            current_gesture = self.finger

        if current_gesture == self.prev_gesture:
            self.frame_count += 1
        else:
            self.frame_count = 0

        self.prev_gesture = current_gesture

        if self.frame_count > 4:
            self.ori_gesture = current_gesture
        return self.ori_gesture


# Executes commands according to detected gestures
class Controller:

    tx_old = 0
    ty_old = 0
    trial = True
    flag = False
    grabflag = False
    pinchmajorflag = False
    pinchminorflag = False
    pinchstartxcoord = None
    pinchstartycoord = None
    pinchdirectionflag = None
    prevpinchlv = 0
    pinchlv = 0
    framecount = 0
    prev_hand = None
    pinch_threshold = 0.3
    #fingers = detector.fingersUp()


    def getpinchylv(hand_result):
        """returns distance beween starting pinch y coord and current hand position y coord."""
        dist = round((Controller.pinchstartycoord - hand_result.landmark[8].y) * 10, 1)
        return dist

    def getpinchxlv(hand_result):
        """returns distance beween starting pinch x coord and current hand position x coord."""
        dist = round((hand_result.landmark[8].x - Controller.pinchstartxcoord) * 10, 1)
        return dist


    def changesystembrightness():
        """sets system brightness based on 'Controller.pinchlv'."""
        l=sbcontrol.get_brightness(display=0)

        currentBrightnessLv = l[0]/ 100.0
        currentBrightnessLv += Controller.pinchlv / 50.0
        if currentBrightnessLv > 1.0:
            currentBrightnessLv = 1.0
        elif currentBrightnessLv < 0.0:
            currentBrightnessLv = 0.0
        sbcontrol.fade_brightness(int(100 * currentBrightnessLv), start=l[0])

    def changesystemvolume():
        """sets system volume based on 'Controller.pinchlv'."""
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        currentVolumeLv = volume.GetMasterVolumeLevelScalar()
        currentVolumeLv += Controller.pinchlv / 50.0
        if currentVolumeLv > 1.0:
            currentVolumeLv = 1.0
        elif currentVolumeLv < 0.0:
            currentVolumeLv = 0.0
        volume.SetMasterVolumeLevelScalar(currentVolumeLv, None)


    # def changeClock():
    #     # Open the clock application
    #     pyautogui.press('win')
    #     pyautogui.typewrite('clock')
    #     pyautogui.press('enter')
    #     # Move the mouse to the time field
    #     pyautogui.moveTo(500, 300)
    #     # Click and hold the left mouse button
    #     pyautogui.mouseDown()
    #     # Move the mouse to increase the hours by 1
    #     pyautogui.moveRel(50, 0)
    #     # Release the left mouse button
    #     pyautogui.mouseUp()

    def scrollVertical():
        """scrolls on screen vertically."""
        pyautogui.scroll(120 if Controller.pinchlv > 0.0 else -120)

    def scrollHorizontal():
        """scrolls on screen horizontally."""
        pyautogui.keyDown('shift')
        pyautogui.keyDown('ctrl')
        pyautogui.scroll(-120 if Controller.pinchlv > 0.0 else 120)
        pyautogui.keyUp('ctrl')
        pyautogui.keyUp('shift')

    # def crossed_fingers():
    #     pyautogui.hotkey('ctrl', 'alt', 'f')  # Execute the "crossed fingers" command
    #
    #     # Register the crossed fingers gesture with PyAutoGUI
    #     pyautogui.PAUSE = 1  # Set the pause time to 1 second for demonstration purposes
    #     pyautogui.keyDown('ctrl')  # Press the "ctrl" key to start recording the gesture
    #     pyautogui.keyDown('alt')  # Press the "alt" key to continue recording the gesture
    #     pyautogui.press('f')  # Press the "f" key to complete the gesture
    #     pyautogui.keyUp('ctrl')  # Release the "ctrl" key to stop recording the gesture
    #     pyautogui.keyUp('alt')  # Release the "alt" key to complete the gesture registration
    #
    #     # Test the crossed fingers gesture
    #     time.sleep(1)  # Wait for 1 second
    #     #crossed_fingers()  # Perform the crossed fingers gesture

    # Locate Hand to get Cursor Position
    # Stabilize cursor by Dampening
    def get_position(hand_result):
        point = 9
        position = [hand_result.landmark[point].x, hand_result.landmark[point].y]
        sx, sy = pyautogui.size()
        x_old, y_old = pyautogui.position()
        x = int(position[0] * sx)
        y = int(position[1] * sy)
        if Controller.prev_hand is None:
            Controller.prev_hand = x, y
        delta_x = x - Controller.prev_hand[0]
        delta_y = y - Controller.prev_hand[1]

        distsq = delta_x ** 2 + delta_y ** 2
        ratio = 1
        Controller.prev_hand = [x, y]

        if distsq <= 25:
            ratio = 0
        elif distsq <= 900:
            ratio = 0.07 * (distsq ** (1 / 2))
        else:
            ratio = 2.1
        x, y = x_old + delta_x * ratio, y_old + delta_y * ratio
        return (x, y)

    def pinch_control_init(hand_result):
        """Initializes attributes for pinch gesture."""
        Controller.pinchstartxcoord = hand_result.landmark[8].x
        Controller.pinchstartycoord = hand_result.landmark[8].y
        Controller.pinchlv = 0
        Controller.prevpinchlv = 0
        Controller.framecount = 0

    # Hold final position for 5 frames to change status
    def pinch_control(hand_result, controlHorizontal, controlVertical):
        if Controller.framecount == 5:
            Controller.framecount = 0
            Controller.pinchlv = Controller.prevpinchlv

            if Controller.pinchdirectionflag == True:
                controlHorizontal()  # x

            elif Controller.pinchdirectionflag == False:
                controlVertical()  # y

        lvx = Controller.getpinchxlv(hand_result)
        lvy = Controller.getpinchylv(hand_result)

        if abs(lvy) > abs(lvx) and abs(lvy) > Controller.pinch_threshold:
            Controller.pinchdirectionflag = False
            if abs(Controller.prevpinchlv - lvy) < Controller.pinch_threshold:
                Controller.framecount += 1
            else:
                Controller.prevpinchlv = lvy
                Controller.framecount = 0

        elif abs(lvx) > Controller.pinch_threshold:
            Controller.pinchdirectionflag = True
            if abs(Controller.prevpinchlv - lvx) < Controller.pinch_threshold:
                Controller.framecount += 1
            else:
                Controller.prevpinchlv = lvx
                Controller.framecount = 0

    def handle_controls(gesture, hand_result):
        """Impliments all gesture functionality."""
        x, y = None, None
        # if fingers[1] == 1 and fingers[2] == 0:
        #     Controller.changeClock()

        if gesture != Gest.PALM:
            x, y = Controller.get_position(hand_result)


        # flag reset
        if gesture != Gest.FIST and Controller.grabflag:
            Controller.grabflag = False
            pyautogui.mouseUp(button="left")

        if gesture != Gest.PINCH_MAJOR and Controller.pinchmajorflag:
            Controller.pinchmajorflag = False

        if gesture != Gest.PINCH_MINOR and Controller.pinchminorflag:
            Controller.pinchminorflag = False

        if gesture == Gest.INDEX and Controller.flag:
            def screenshot():
                screenshot = pyautogui.screenshot()
                screenshot.save('screenshot1.png')

            screenshot()
            # subprocess.Popen(['control', 'sysdm.cpl'])

        if gesture == Gest.MID and Controller.flag:
            # Move the mouse to the Windows Start button
            #pyautogui.moveTo(50, 1050)
            screenWidth, screenHeight = pyautogui.size()

            # Move the mouse to the start button
            startButtonX = screenWidth // 3
            startButtonY = screenHeight - 10
            pyautogui.moveTo(startButtonX, startButtonY)

            # Click on the Start button
            pyautogui.click()

            # Type "Settings" in the search bar and hit enter
            pyautogui.typewrite('Settings')
            time.sleep(1)
            pyautogui.press('enter')

            time.sleep(1)

            pyautogui.typewrite('Bluetooth and other devices settings')
            time.sleep(1)
            pyautogui.press('enter')
            #pyautogui.click()


            # Click on the "Devices" option
            blueButtonY = screenHeight // 4
            blueButtonX = screenWidth // 3
            pyautogui.moveTo(blueButtonX, blueButtonY)
            pyautogui.click()
            #pyautogui.click(400, 300)

            #time.sleep(1)

            # Click on "Bluetooth & other devices"
            mainButtonY = screenHeight // 6
            mainButtonX = screenWidth // 1.12
            pyautogui.moveTo(mainButtonX, mainButtonY)
            pyautogui.click()

            Controller.flag = False

        if gesture == (Gest.PINKY and Gest.INDEX) and Controller.flag:
            #pyautogui.click()
            subprocess.Popen(['control', 'sysdm.cpl'])
            Controller.flag = False




        # implementation
        if gesture == Gest.V_GEST:
            Controller.flag = True
            pyautogui.moveTo(x, y, duration=0.1)

        elif gesture == Gest.FIST:
            if not Controller.grabflag:
                Controller.grabflag = True
                pyautogui.mouseDown(button="left")
            pyautogui.moveTo(x, y, duration=0.1)

        elif gesture == Gest.MID and Controller.flag:
            pyautogui.click()
            Controller.flag = False

        elif gesture == Gest.INDEX and Controller.flag:
            pyautogui.click(button='right')
            Controller.flag = False

        # elif not gesture != HandRecog.rotated_points and Controller.flag:
        #     subprocess.Popen(['control', 'sysdm.cpl'])
        #     Controller.flag = False



        elif gesture == Gest.PINKY and Controller.flag:

            #4.To open the clock of the system

            pyautogui.press('win')
            pyautogui.typewrite('clock')
            pyautogui.press('enter')
            # Move the mouse to the time field
            pyautogui.moveTo(500, 300)
            # Click and hold the left mouse button
            pyautogui.mouseDown()
            # Move the mouse to increase the hours by 1
            pyautogui.moveRel(50, 0)
            # Release the left mouse button
            pyautogui.mouseUp()
            Controller.flag = False

        elif gesture == Gest.RING and Controller.flag:
            engine = pyttsx3.init('sapi5')  # sapi5--> TTS engine on windows
            voices = engine.getProperty('voices')
            engine.setProperty('voice', voices[1].id)

            def speak(audio):
                engine.say(audio)
                engine.runAndWait()

            def wishMe():
                hour = int(datetime.datetime.now().hour)
                if hour >= 0 and hour < 12:
                    speak("Good Morning!")
                elif hour >= 12 and hour < 18:
                    speak("Good afternoon!")
                else:
                    speak("Good evening!")
                speak("I am Jarvis Mam. Please tell me how may I help you?")

            def takeCommand():
                # It takes microphone input from the user and returns string output
                r = sr.Recognizer()
                with sr.Microphone() as source:
                    print("Listening...")
                    r.pause_threshold = 1
                    audio = r.listen(source)

                try:
                    print("Recognizing...")
                    query = r.recognize_google(audio, language='en-in')
                    print(f"User said: {query}\n")

                except Exception as e:
                    # print(e)
                    print("Say that again please...")
                    return "None"
                return query
            #Controller.voiceAssistant()
            wishMe()
            while True:
                query = takeCommand().lower()
                # Logic for executing tasks based on query
                if 'wikipedia' in query:
                    speak("Searching wikipedia...")
                    query = query.replace("wikipedia", "")
                    results = wikipedia.summary(query, sentences=2)
                    speak("According to Wikipedia")
                    print(results)
                    speak(results)

                elif 'open youtube' in query:
                    webbrowser.open("youtube.com")

                elif 'open google' in query:
                    webbrowser.open("google.com")

                elif 'open stackoverflow' in query:
                    webbrowser.open("stackoverflow.com")

                elif 'the time' in query:
                    strTime = datetime.datetime.now().strftime("%H:%M:%S")
                    speak(f"Mam, the time is {strTime}")

                elif 'open code' in query:
                    codePath = "C:\\Users\\thepa\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
                    os.startfile(codePath)

                elif 'close yourself' in query:
                    exit()


        elif gesture == Gest.TWO_FINGER_CLOSED and Controller.flag:
            pyautogui.doubleClick()
            Controller.flag = False

        elif gesture == Gest.PINCH_MINOR:
            if Controller.pinchminorflag == False:
                Controller.pinch_control_init(hand_result)
                Controller.pinchminorflag = True
            Controller.pinch_control(hand_result, Controller.scrollHorizontal, Controller.scrollVertical)

        elif gesture == Gest.PINCH_MAJOR:
            if Controller.pinchmajorflag == False:
                Controller.pinch_control_init(hand_result)
                Controller.pinchmajorflag = True
            Controller.pinch_control(hand_result, Controller.changesystembrightness, Controller.changesystemvolume)


'''
----------------------------------------  Main Class  ----------------------------------------
    Entry point of Gesture Controller
'''


class GestureController:
    gc_mode = 0
    cap = None
    CAM_HEIGHT = None
    CAM_WIDTH = None
    hr_major = None  # Right Hand by default
    hr_minor = None  # Left hand by default
    dom_hand = True
    def __init__(self):
        """Initilaizes attributes."""
        GestureController.gc_mode = 1
        GestureController.cap = cv2.VideoCapture(0)
        GestureController.CAM_HEIGHT = GestureController.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
        GestureController.CAM_WIDTH = GestureController.cap.get(cv2.CAP_PROP_FRAME_WIDTH)


    def classify_hands(results):
        left, right = None, None
        try:
            handedness_dict = MessageToDict(results.multi_handedness[0])
            if handedness_dict['classification'][0]['label'] == 'Right':
                right = results.multi_hand_landmarks[0]
            else:
                left = results.multi_hand_landmarks[0]
        except:
            pass

        try:
            handedness_dict = MessageToDict(results.multi_handedness[1])
            if handedness_dict['classification'][0]['label'] == 'Right':
                right = results.multi_hand_landmarks[1]
            else:
                left = results.multi_hand_landmarks[1]
        except:
            pass

        if GestureController.dom_hand == True:
            GestureController.hr_major = right
            GestureController.hr_minor = left
        else:
            GestureController.hr_major = left
            GestureController.hr_minor = right

    def start(self):
        handmajor = HandRecog(HLabel.MAJOR)
        handminor = HandRecog(HLabel.MINOR)

        with mp_hands.Hands(max_num_hands=2, min_detection_confidence=0.5, min_tracking_confidence=0.5) as hands:
            while GestureController.cap.isOpened() and GestureController.gc_mode:

                success, image = GestureController.cap.read()


                if not success:
                    print("Ignoring empty camera frame.")
                    continue

                image = cv2.cvtColor(cv2.flip(image, 1), cv2.COLOR_BGR2RGB)
                image.flags.writeable = False
                results = hands.process(image)


                image.flags.writeable = True
                image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

                if results.multi_hand_landmarks:
                    GestureController.classify_hands(results)
                    handmajor.update_hand_result(GestureController.hr_major)
                    handminor.update_hand_result(GestureController.hr_minor)

                    handmajor.set_finger_state()
                    handminor.set_finger_state()
                    gest_name = handminor.get_gesture()

                    if gest_name == Gest.PINCH_MINOR:
                        Controller.handle_controls(gest_name, handminor.hand_result)
                    else:
                        gest_name = handmajor.get_gesture()
                        Controller.handle_controls(gest_name, handmajor.hand_result)

                    for hand_landmarks in results.multi_hand_landmarks:
                        mp_drawing.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)


                else:
                    Controller.prev_hand = None
                cv2.imshow('Gesture Controller', image)
                if cv2.waitKey(5) & 0xFF == ord('q'):
                    break

# uncommqent to run directly
gc1 = GestureController()
gc1.start()