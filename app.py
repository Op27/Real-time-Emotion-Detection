"""
Real-time Emotion Detection Application
---------------------------------------
This script captures screen content in real-time, detects emotions from faces displayed,
and overlays the detected emotions with text labels and colors. It's designed for enhancing
online meetings by providing live emotional feedback.
"""

from rmn import RMN  # Import the Residual Masking Network (RMN) model for emotion detection
import cv2  # Import OpenCV for image processing
import mss  # Import MSS for screen capture functionality
import numpy  # Import NumPy for numerical operations on arrays
import win32gui  # Import win32gui for window management
import win32con  # Import win32con for constants used with win32gui

# Initialize the RMN model for emotion detection
m = RMN()

# Use MSS to capture the screen
with mss.mss() as sct:
    monitor = {"top": 0, "left": 0, "width": 1000, "height": 1000}
    
    # Main loop to continuously capture the screen and detect emotions
    while True:
        img = numpy.array(sct.grab(monitor))
        
        # Convert the image from RGBA to RGB (if necessary)
        if img.shape[2] == 4:
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        
        # Use the RMN model to detect emotion from the captured frame
        results = m.detect_emotion_for_single_frame(img)
        
        # Draw the results on the frame
        img = m.draw(img, results)
        
        # Overlay the instruction text on the frame
        cv2.putText(img, "Press 'q' to close window", (10, 30), 
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)

        # Display the frame with detected emotions in a window titled "Emotion Detection"
        cv2.imshow("Emotion Detection", img)
        
        # Configure the "Emotion Detection" window to be normal sized and always on top
        cv2.namedWindow("Emotion Detection", cv2.WINDOW_NORMAL)
        cv2.resizeWindow("Emotion Detection", 1000, 960)


        hwnd = win32gui.FindWindow(None, "Emotion Detection")
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)

        # Break the loop and close the window if the 'q' key is pressed
        if cv2.waitKey(25) & 0xFF == ord("q"):
            cv2.destroyAllWindows()
            break
