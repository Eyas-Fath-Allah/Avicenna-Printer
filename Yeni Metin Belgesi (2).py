import pyautogui
import time
import cv2
import numpy as np

def mse(img1, img2):
   h, w = img1.shape
   diff = cv2.subtract(img1, img2)
   err = np.sum(diff**2)
   mse = err/(float(h*w))
   return mse


# # Click to Avicenna
# pyautogui.click(30, 10)
# time.sleep(0.5)
# # Write Patient No.
# pyautogui.moveTo(100, 85)
# pyautogui.click(100, 85)
# pyautogui.click(100, 85)
# pyautogui.press('del')
# time.sleep(0.5)
# pyautogui.typewrite('1738856')
# pyautogui.press('enter')
# time.sleep(4)
# # Click to 'HASTA BAZINDA'
# pyautogui.moveTo(40, 300)
# pyautogui.click(40, 300)
# time.sleep(2)
# # Click to 'Laboratuvar'
# pyautogui.moveTo(400, 185)
# pyautogui.click(400, 185)
# time.sleep(2)
# # Search for the test.
# pyautogui.moveTo(90, 275)
# pyautogui.click(90, 275)
# pyautogui.typewrite('MENEN')
# pyautogui.press('enter')

screenshot = pyautogui.screenshot(region=(60, 310, 200, 400))
screenshot.save('C:\\Users\\user\\Desktop\\Test\\screenshot.jpg')
img1 = cv2.imread('C:\\Users\\user\\Desktop\\Test\\screenshot.jpg')
img2 = cv2.imread('C:\\Users\\user\\Desktop\\Test\\1test.jpg')

img1_gray = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
img2_gray = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)

error = mse(img1_gray, img2_gray)
print("Image matching Error between the two images:", error)
cv2.waitKey(0)
cv2.destroyAllWindows()