import win32gui
import win32api
import win32ui
import win32con
from ctypes import windll
from PIL import Image
import numpy as np
from paddleocr import PaddleOCR
import time
import re
from collections import Counter

def capture_window(hwnd,crop_coords,width, height):

    # 创建设备上下文（DC）
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    # 创建位图对象
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)

    # 复制窗口内容到位图
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

    # 将位图转换为PIL图像
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    # 清理资源
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    if result == 1:
        # 裁剪图像
        im = im.crop((crop_coords[0][0], crop_coords[0][1], crop_coords[1][0], crop_coords[1][1]))
        return im
    else:
        print("【截图失败】")
        return None

def ocr_image(ocr, image):
    # 执行OCR
    result = ocr.ocr(np.array(image), cls=True)
    
    # 提取文本和坐标
    texts_with_coords = []
    if result and isinstance(result[0], list):
        for line in result[0]:
            # print(line)
            if isinstance(line, list) and len(line) > 1:
                texts_with_coords.append((line[0][0],line[1][0]))  # 提取每行的文本内容
    
    return texts_with_coords

def set_window_topmost(hwnd):
    if hwnd:
        
        flagTrue = win32con.HWND_TOPMOST
        flagFlase = win32con.HWND_NOTOPMOST
      
        win32gui.SetWindowPos(hwnd, flagTrue, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(hwnd, flagFlase, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        print(f"【窗口已置顶】")
    else:
        print("【未找到窗口】")

def move_mouse_to(right, bottom):
    # 获取屏幕尺寸
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    # 计算绝对坐标
    # 注意这里的计算顺序：先除以屏幕宽度/高度，再乘以 19
    x_abs = int(65536 * (right / screen_width) * 19/20)
    y_abs = int(65536 * (bottom / screen_height) * 55/100)
    y_abs_s = int(65536 * (bottom / screen_height) * 56/100)

    # 使用win32api.mouse_event移动鼠标
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, x_abs, y_abs, 0, 0)
    # 按下鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.1)
    # 计算偏移量
    delta_y = y_abs_s - y_abs

    # 移动到终点
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 0, delta_y, 0, 0)
    time.sleep(0.1)

    # 释放鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def process_texts(texts_with_coords):
    name = ""
    message = ""
    coord = []
    processed_texts = []
    for texts_with_coord in texts_with_coords:
        if "互联"  in texts_with_coord[1] or "帮派"  in texts_with_coord[1]:
            processed_texts.append((coord,name, message))
            name = re.sub(r'互联|\[|\]|\s', '', texts_with_coord[1])
            coord = texts_with_coord[0]
            message = ""
        else:
            message += re.sub(r'\s+', '', texts_with_coord[1])
    # 最后一条信息会被忽略
    return processed_texts[1:]

def jaccard_similarity(ele1, ele2):
    # 使用 Counter 来统计每个字符的出现次数
    # print(ele1[0]+ele1[1])
    # print((ele2[0]+ele2[1]))
    counter1 = Counter(ele1[0]+ele1[1])
    counter2 = Counter(ele2[0]+ele2[1])
    
    # 计算交集和并集
    intersection = sum((counter1 & counter2).values())
    union = sum((counter1 | counter2).values())
    
    # 计算 Jaccard 相似度
    if union == 0:
        return 0.0
    else:
        print(intersection / union)
        return intersection / union

def collect_message(texts_record,last_text_record):
    """ 用于记录世界消息，待完善
    """
    repeat_index = len(texts_record)
    print(repeat_index)
    # 遍历 texts_record 列表
    for index, element in enumerate(texts_record):
        # 如果元素也在 last_text_record 中，则记录重复的位置
        for element_l in last_text_record:
            # 由于每次识别有误差，所以通过相似度计算来判断重复消息和新消息从哪里开始
            if jaccard_similarity(element,element_l) < 0.6:
                repeat_index = index
                break
        break

    # print(texts_record)
    # print(last_text_record)
    print(repeat_index)
    new_elements = texts_record[repeat_index:]
    if new_elements != []:
        with open("message.txt", "a", encoding='utf-8') as f:
            for element in new_elements:
                f.write(element[0] + "：" + element[1] + "\n")


def click_head(relative_coords,crop_coords,left, top):
    # 获取屏幕尺寸
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    right = relative_coords[0]+crop_coords[0][0]+left
    bottom = relative_coords[1]+crop_coords[0][1]+top
    # print(right,bottom)
    # 计算绝对坐标
    # 注意这里的计算顺序：先除以屏幕宽度/高度，再乘以 19, 头像位置跟文字位置稍微有偏移
    x_abs = int(65536 * (right / screen_width) )
    y_abs = int(65536 * (bottom / screen_height) * 58/50)

    # 使用win32api.mouse_event移动鼠标,先激活屏幕
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE,int(0.5* x_abs), int(y_abs), 0, 0)
    # 按下鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.1)
    # 释放鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.5)
    # 使用win32api.mouse_event移动鼠标
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, x_abs, y_abs, 0, 0)
    # 按下鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.1)
    # 释放鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def click_once(x,y):
    # 获取屏幕尺寸
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    # 计算绝对坐标
    # 注意这里的计算顺序：先除以屏幕宽度/高度，再乘以 19
    x_abs = int(65536 * (x / screen_width) * 29/30)
    y_abs = int(65536 * (y / screen_height)* 29/30)

    # 使用win32api.mouse_event移动鼠标
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, x_abs, y_abs, 0, 0)
    # 按下鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.1)
    # 释放鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def click_text(hwnd, ocr,left,top,width, height):
    screenshot = capture_window(hwnd,[(0, 0), (width, height)],width, height)
    texts_with_coords = ocr_image(ocr, screenshot)
    for c,t in texts_with_coords:
        if "申请入队" in t:
            click_once(left+c[0],top+c[1])


def main():
    window_title = "一梦江湖"  # 替换为目标窗口的实际标题
    
    target_words = ["随便来"]  #"临本","团临","临条","五人临","奖励号"
    
    # 查找窗口
    hwnd = win32gui.FindWindow(None, window_title)
    # 获取窗口大小
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    print(left, top, right, bottom)
    # 默认窗口大小
    windows_size = [1350,789]
    crop_coords = [(1137, 85), (1330, 738)]  # 裁剪坐标
    width = right - left
    height = bottom - top
    # 根据当前窗口大小放缩
    crop_coords = [
        (int(width * crop_coords[0][0] / windows_size[0]),
            int(height * crop_coords[0][1] / windows_size[1])),
        (int(width * crop_coords[1][0] / windows_size[0]),
            int(height * crop_coords[1][1] / windows_size[1]))
    ]
    print(crop_coords)
    if not hwnd:
        print(f"【未找到标题为 '{window_title}' 的窗口】")
        return

    # 初始化PaddleOCR
    ocr = PaddleOCR(use_angle_cls=True, lang="ch")  # 使用中文模型
    last_text_record = []

    flag = True
    while flag:
        screenshot = capture_window(hwnd,crop_coords,width, height)
        if screenshot:
            # 确定保存截图没问题后可以省略保存步骤
            screenshot.save("cropped_game_screenshot.png")
            # print("【裁剪后的截图已保存为 cropped_game_screenshot.png】")
            
            # 执行OCR
            texts_with_coords = ocr_image(ocr, screenshot)
            texts = [text[1:] for text in texts_with_coords]
            print(texts_with_coords)
            # 收集消息并保存
            texts_record = process_texts(texts_with_coords)
            # collect_message(texts_record,last_text_record)
            # last_text_record = texts_record
            print("【识别到的文本】：")
            print(texts_record)

            for i, text in enumerate(texts_record):
                if any(word in text[2] for word in target_words):
                    print(f"【检测到】：{text}")
                    set_window_topmost(hwnd)
                    move_mouse_to(right, bottom)
                    flag = False
                    break
        else:
            print("【无法获取截图】")
        
        time.sleep(1)

    # 最终确定位置
    screenshot = capture_window(hwnd,crop_coords,width, height)
    texts_with_coords = ocr_image(ocr, screenshot)
    for i, text in enumerate(texts_record):
        if any(word in text[2] for word in target_words):
            print(f"【检测到】：{text}")
            click_head(text[0],crop_coords,left, top)
            time.sleep(1)
            click_text(hwnd,ocr,left,top,width, height)
            break

    

if __name__ == "__main__":
    main()