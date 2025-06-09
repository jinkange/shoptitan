try:
    import win32gui
    import win32con
    import time
    import pyautogui
    import cv2
    import numpy as np
    import random
    from PIL import ImageGrab
    import os
    import datetime
    import time
    import re
    import pyautogui
    # import pytesseract
    import cv2
    import numpy as np
    

except Exception as e:
    print(e)
    
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# def find_number_in_region(target_number):
    
#     # region: (x, y, width, height)
#     screenshot = pyautogui.screenshot(region=(24, 229,866, 171))
#     img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

#     # OCR 텍스트 추출
#     text = pytesseract.image_to_string(img, config='--psm 6 digits')
#     print(f"OCR 결과:\n{text}")

#     # 숫자 포함 여부 확인
#     return str(target_number) in text

# 예시 사용
region = (12, 256, 1200, 800)
number_to_find = 123456

# if find_number_in_region( number_to_find):
#     print("숫자를 찾았습니다!")
# else:
#     print("숫자를 찾을 수 없습니다.")
    
def enum_windows_by_title(title):
    """특정 창 제목과 일치하는 핸들을 반환"""
    hwnds = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and title in win32gui.GetWindowText(hwnd):
            hwnds.append(hwnd)
    win32gui.EnumWindows(callback, None)
    return hwnds
def enum_windows_by_class_and_title( title):
    hwnds = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            window_class = win32gui.GetClassName(hwnd)
            window_title = win32gui.GetWindowText(hwnd)
            if  window_title == title:
                hwnds.append(hwnd)

    win32gui.EnumWindows(callback, None)
    return hwnds
    
def move_resize_window(hwnd, x, y, width, height):
    """창 위치와 크기 조절"""
    win32gui.MoveWindow(hwnd, x, y, width, height, True)
def get_window_rect(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    return rect  # (left, top, right, bottom)
def get_sorted_odin_windows():
    odin_windows = enum_windows_by_class_and_title( "Shop Titans")
    odin_windows = sorted(odin_windows, key=lambda hwnd: get_window_rect(hwnd)[0])  # 좌측 기준 정렬
    return odin_windows
def image_exists_at_region(template_path, region, threshold=0.95):
    """
    template_path: 찾을 이미지 파일 경로
    region: (x, y, width, height)1281 631
    threshold: 일치 정도 (0.0 ~ 1.0)
    """
    screenshot = pyautogui.screenshot(region=region)
    # screenshot = screenshot_all_monitors()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

    template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)

    result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    max_val = np.max(result)
    print(f"{template_path}찾기 :{max_val}")
    return max_val >= threshold

def click_image_if_exists(template_path, region, threshold=0.95):
    """
    template_path: 찾을 이미지 파일 경로
    region: (x, y, width, height)
    threshold: 일치 정도 (0.0 ~ 1.0)
    """
    # 스크린샷을 찍고 흑백 변환
    screenshot = pyautogui.screenshot(region=region)
    screenshot_gray = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

    # 템플릿 이미지 불러오기
    template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
    h, w = template.shape[:2]


    # 템플릿 매칭
    result = cv2.matchTemplate(screenshot_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    print(f"{template_path} 찾기 : {max_val}")

    if max_val >= threshold:
        # region은 (x, y, width, height)이므로 클릭 좌표는 실제 화면 기준으로 변환해야 함
        region_x, region_y = region[0], region[1]
        match_x, match_y = max_loc
        click_x = region_x + match_x + w // 2
        click_y = region_y + match_y + h // 2

        # 클릭 실행
        pyautogui.click(click_x, click_y)
        return True

    return False


def main():
    odin_windows = get_sorted_odin_windows()
    console_windows = enum_windows_by_title("shoptitan")
    move_resize_window(odin_windows[0], 0, 0, 2060, 1240)# 왼쪽
    move_resize_window(console_windows[0], 0, 1240, 2060,200)
    # while True:
    #     # click_image("./images/talk.png")
    #     click_image_if_exists("./images/talk.png",region)
    #     time.sleep(1)

def click_image_with_alpha(template_path, region, threshold=0.95):
    import pyautogui
    import cv2
    import numpy as np

    # 스크린샷: RGB로 얻기
    screenshot = pyautogui.screenshot(region=region)
    screenshot_rgb = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # 템플릿 이미지 불러오기 (RGBA → BGR)
    template = cv2.imread(template_path, cv2.IMREAD_UNCHANGED)
    if template.shape[2] == 4:  # 알파 채널이 있는 경우
        # 알파값이 0인 부분을 제외한 부분만 사용
        alpha_mask = template[:, :, 3] > 0
        template_rgb = template[:, :, :3].copy()
        template_rgb[~alpha_mask] = 0
    else:
        template_rgb = template

    # 스크린샷에서도 알파 없는 부분만 비교하도록 마스크를 만들어야 가장 정확
    result = cv2.matchTemplate(screenshot_rgb, template_rgb, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    print(f"{template_path} 찾기 (알파 고려): {max_val}")

    if max_val >= threshold:
        region_x, region_y = region[:2]
        h, w = template_rgb.shape[:2]
        click_x = region_x + max_loc[0] + w // 2
        click_y = region_y + max_loc[1] + h // 2
        pyautogui.click(click_x, click_y)
        return True

    return False

def click_image(image_path, confidence=0.8):
    location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
    x,y = location
    if location:
        pyautogui.click(x-2,y-5)
        print(f"이미지 클릭 성공: {location}")
    else:
        print("이미지 찾지 못함.")
        
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("매크로가 종료되었습니다. (Ctrl+C)")
    finally:
        True
        