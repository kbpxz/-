import cv2
import numpy as np
import pyautogui
import os
import pyperclip
import time
import json

def load_steps():
    """加载步骤配置"""
    current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_path = os.path.join(current_dir, "config", "steps.json")
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def find_image_center(template, screenshot, threshold=0.8):
    """查找图片并返回中心坐标"""
    try:
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            center_x = max_loc[0] + template.shape[1] // 2
            center_y = max_loc[1] + template.shape[0] // 2
            return center_x, center_y, max_val
        return None
    except Exception as e:
        print(f"图片匹配错误：{e}")
        return None

def execute_step(step, screenshot, root_dir):
    """执行单个步骤"""
    template_path = os.path.join(root_dir, step["template_path"])
    if not os.path.exists(template_path):
        return None, f"模板文件不存在：{template_path}"
    
    template = cv2.imread(template_path)
    if template is None:
        return None, f"无法读取模板图片：{template_path}"

    threshold = step.get("threshold", 0.8)
    result = find_image_center(template, screenshot, threshold)
    
    if result is None:
        return None, f"未找到匹配图像：{step['name']}"
    
    x, y, confidence = result
    
    if step["action"] == "copy":
        offset = step.get("offset", {"x": 0, "y": 0})
        pyautogui.moveTo(x + offset["x"], y + offset["y"], duration=0.2)
        for _ in range(3):
            pyautogui.click()
            time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
    else:  # click
        pyautogui.moveTo(x, y, duration=0.2)
        pyautogui.click()
    
    time.sleep(0.2)
    content = pyperclip.paste()
    return content, None

def start_capture(app_instance):
    """开始捕获和处理"""
    try:
        # 获取根目录
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        # 加载步骤配置
        steps = load_steps()
        
        # 截取屏幕
        screen_width, screen_height = pyautogui.size()
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        
        # 存储步骤结果
        results = []
        
        # 执行每个步骤
        for step in steps:
            content, error = execute_step(step, screenshot, root_dir)
            if error:
                print(f"步骤 {step['name']} 执行失败：{error}")
                continue
            if content:
                results.append(content)
        
        # 如果获取到了所有需要的信息，更新UI
        if len(results) >= 3:
            app_instance.add_new_data(
                客户昵称=results[0],
                商家=results[1],
                订单号=results[2]
            )
            return True
        
        return False
    
    except Exception as e:
        print(f"捕获过程发生错误：{e}")
        return False