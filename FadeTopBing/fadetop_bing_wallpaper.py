# -*- coding: utf-8 -*-
import datetime
import re
import os
import time
import ssl
import shutil
import subprocess
from pathlib import Path

ROOT_DIR = os.path.dirname(os.path.abspath(os.path.dirname(__file__)))

# 创建必要目录
WALLPAPER_SAVE_DIR = os.path.join(ROOT_DIR, 'bing_wallpaper')
os.makedirs(os.path.join(ROOT_DIR, 'FadeTopBing'), exist_ok=True)
os.makedirs(os.path.join(ROOT_DIR, 'FadeTop'), exist_ok=True)
os.makedirs(WALLPAPER_SAVE_DIR, exist_ok=True)

try:
    from urllib.request import Request
except ImportError:
    from urllib2 import Request


def urlopen(req):
    try:
        from urllib.request import urlopen as f
        return f(req, context=ssl._create_unverified_context())
    except:
        from urllib2 import urlopen as f
        return f(req)


def urlretrieve(url, path):
    try:
        from urllib.request import urlretrieve as f
        return f(url, path)
    except:
        try:
            with open(path, 'wb') as f:
                f.write(urlopen(url).read())
            return True
        except Exception as e:
            print(f"下载图片失败: {e}")
            return False


headers = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'referer': 'https://cn.bing.com/',
}


def get_config():
    config_file_path = os.path.join(ROOT_DIR, 'config.txt')
    motto = '请在config.txt文件中配置'
    if os.path.exists(config_file_path):
        try:
            with open(config_file_path, 'r', encoding='utf8') as m:
                motto = m.read().strip() or motto
        except Exception as e:
            print(f"读取配置文件失败: {e}")
    else:
        try:
            with open(config_file_path, 'w', encoding='utf8') as m:
                m.write(motto)
        except Exception as e:
            print(f"创建配置文件失败: {e}")
    return motto


def change_wallpaper(motto, image_path):
    if not image_path or not os.path.exists(image_path):
        print("壁纸文件不存在，跳过配置修改")
        return
        
    setting_xml_path = os.path.join(ROOT_DIR, 'FadeTop', 'Settings.xml')
    setting_xml_str = r'''
<?xml version="1.0" encoding="UTF-8" ?>
<FadeTopSettings version="3.0" xmlns="http://www.fadetop.com/fadetop/xmlns/settings">
    <States>
        <Application debut="0" auto_fade_enabled="1" />
        <User break_trail="0:0" />
    </States>
    <Options>
        <General activity_timeout="15" idle_timeout="5" fade_again_delay="1" auto_block_enabled="1" />
        <Fader max_opacity="100">
            <Foreground fg_color="#FFFFFF" fg_position="center" fg_offset_x="0" fg_offset_y="0" fg_time_format="auto" fg_message="请在config.txt文件中配置" />
            <Background bg_color="#008040" bg_image_enabled="1" bg_random_image_enabled="0" bg_image_file="" bg_image_position="fill" />
        </Fader>
        <Sound sound_enabled="0" sound_file="" sound_fadein_enabled="0" sound_fadeout_enabled="1" sound_volume_default="50" />
    </Options>
</FadeTopSettings>
    '''
    if not os.path.exists(setting_xml_path):
        with open(setting_xml_path, 'wb') as f:
            f.write(setting_xml_str.encode('utf8'))    

    try:
        with open(setting_xml_path, 'r', encoding='utf8') as f:
            setting_xml_str = f.read()
            if not setting_xml_str:
                return
            
            # 替换壁纸路径
            bg_file_match = re.search(r'(bg_image_file=".*?")', setting_xml_str)
            if bg_file_match:
                setting_xml_str = setting_xml_str.replace(bg_file_match.group(),
                                                          f'bg_image_file="{image_path}"')
            
            # 启用壁纸
            bg_enabled_match = re.search(r'(bg_image_enabled=".*?")', setting_xml_str)
            if bg_enabled_match:
                setting_xml_str = setting_xml_str.replace(bg_enabled_match.group(),
                                                          'bg_image_enabled="1"')
            
            # 替换座右铭
            fg_match = re.search(r'(<Foreground.*?fg_message.*?/>)', setting_xml_str, re.S)
            if fg_match:
                new_fg_tag = f'<Foreground fg_color="#FFFFFF" fg_position="center" fg_offset_x="0" fg_offset_y="0" fg_time_format="auto" fg_message="{motto}" />'
                setting_xml_str = setting_xml_str.replace(fg_match.group(), new_fg_tag)
        
        with open(setting_xml_path, 'wb') as f:
            f.write(setting_xml_str.encode('utf8'))
        print("配置文件修改成功")
    except Exception as e:
        print(f"修改配置文件失败: {e}")


def kill_process_by_name(process_name):
    try:
        # 使用tasklist /FI过滤，更稳定
        tasklist_output = subprocess.check_output(
            ['tasklist', '/FI', f'IMAGENAME eq {process_name}', '/FO', 'CSV'],
            encoding='gbk',
            errors='ignore'
        )
        lines = [line.strip() for line in tasklist_output.split('\n') if line.strip()]
        # 跳过标题行，处理进程行
        for line in lines[1:]:
            if process_name in line:
                parts = line.replace('"', '').split(',')
                if len(parts) >= 2:
                    pid = parts[1].strip()
                    os.system(f'taskkill /F /PID {pid}')
                    print(f"已终止进程 {process_name} (PID: {pid})")
                    time.sleep(1)
                    break
    except Exception as e:
        # 降级方案
        print(f"⚠️ 进程查杀失败，尝试降级方案: {e}")
        for line in os.popen('tasklist'):
            if process_name in line:
                try:
                    pid = line.strip().split()[1]
                    os.system(f'taskkill /F /PID {pid}')
                    print(f"已终止进程 {process_name} (PID: {pid})")
                    time.sleep(1)
                    break
                except:
                    continue

def kill_FadeTop():
    kill_process_by_name('FadeTop.exe')

def start_FadeTop():
    try:
        exe_path = os.path.join(ROOT_DIR, 'FadeTop','FadeTop.exe')
        if os.path.exists(exe_path):
            # 使用subprocess启动，避免cmd窗口
            subprocess.Popen(
                [exe_path],
                shell=False,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            print("FadeTop 已启动")
        else:
            print(f"FadeTop.exe 不存在: {exe_path}")
    except Exception as e:
        print(f"启动FadeTop失败: {e}")


def get_bing_image():
    """
    获取必应壁纸，保留每日壁纸到 bing_wallpaper 文件夹（按日期命名）
    """
    # 策略1: 先尝试从DynamicTheme缓存获取
    image_path = get_dynamic_bing_image()
    if image_path and os.path.exists(image_path):
        print(f"从DynamicTheme缓存获取壁纸: {image_path}")
        
        # 生成按日期命名的保存路径
        today = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        history_image_path = os.path.join(WALLPAPER_SAVE_DIR, f'bing_wallpaper_{today}.jpg')
        
        # 复制缓存中的壁纸到历史文件夹
        shutil.copy2(image_path, history_image_path)
        print(f"历史壁纸已保存: {history_image_path}")
        
        # 同时复制到FadeTopBing目录供程序使用
        local_path = os.path.join(ROOT_DIR, 'FadeTopBing', 'fadetop_wallpaper.jpg')
        shutil.copy2(image_path, local_path)
        return local_path
    
    # 策略2: 直接从必应API获取（更稳定）
    print("尝试从必应API获取壁纸...")
    bing_api_url = "https://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=zh-CN"
    try:
        req = Request(bing_api_url, headers=headers)
        response = urlopen(req)
        content = response.read().decode('utf-8', errors='ignore')
        
        # 解析JSON获取壁纸URL
        import json
        data = json.loads(content)
        if data.get('images') and len(data['images']) > 0:
            image_info = data['images'][0]
            image_url = image_info['url']
            # 拼接完整URL
            full_image_url = f"https://cn.bing.com{image_url}"
            
            # 1. 先下载到历史文件夹（按日期命名）
            today = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            history_image_path = os.path.join(WALLPAPER_SAVE_DIR, f'bing_wallpaper_{today}.jpg')
            
            # 下载壁纸到历史文件夹
            if urlretrieve(full_image_url, history_image_path):
                print(f"历史壁纸已保存: {history_image_path}")
                
                # 2. 复制到FadeTopBing目录供程序使用
                local_path = os.path.join(ROOT_DIR, 'FadeTopBing', 'fadetop_wallpaper.jpg')
                shutil.copy2(history_image_path, local_path)
                print(f"当前使用壁纸: {local_path}")
                return local_path
    except Exception as e:
        print(f"必应API获取失败: {e}")
    
    # 策略3: 备用爬取方案（兼容旧版正则）
    print("尝试从必应首页爬取壁纸...")
    try:
        url = 'https://cn.bing.com'
        req = Request(url)
        for k, v in headers.items():
            req.add_header(k, v)
        
        content = urlopen(req).read().decode('utf-8', errors='ignore')
        
        # 多种正则匹配模式，提高成功率
        patterns = [
            r'<link id="bgLink".*?href="([^"]+th\?id=[^"]+\.jpg)"',
            r'"backgroundImage":"([^"]+th\?id=[^"]+\.jpg)"',
            r'"url":"([^"]+th\?id=[^"]+\.jpg)"',
            r'/th\?id=([^&]+)\.jpg'
        ]
        
        image_url = None
        for pattern in patterns:
            match = re.search(pattern, content, re.S | re.I)
            if match:
                image_url = match.group(1)
                break
        
        if image_url:
            # 处理URL格式
            if not image_url.startswith('http'):
                if image_url.startswith('/'):
                    image_url = f"https://cn.bing.com{image_url}"
                else:
                    image_url = f"https://cn.bing.com/th?id={image_url}"
            
            # 下载到历史文件夹
            today = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            history_image_path = os.path.join(WALLPAPER_SAVE_DIR, f'bing_wallpaper_{today}.jpg')
            
            # 下载到历史文件夹
            if urlretrieve(image_url, history_image_path):
                print(f"历史壁纸已保存: {history_image_path}")
                
                # 复制到FadeTopBing目录
                local_path = os.path.join(ROOT_DIR, 'FadeTopBing', 'fadetop_wallpaper.jpg')
                shutil.copy2(history_image_path, local_path)
                print(f"当前使用壁纸: {local_path}")
                return local_path
        else:
            print("未能匹配到壁纸URL")
            
    except Exception as e:
        print(f"爬取必应首页失败: {e}")
    
    # 如果所有方法都失败，检查是否有历史壁纸
    fallback_path = os.path.join(ROOT_DIR, 'FadeTopBing', 'fadetop_wallpaper.jpg')
    if os.path.exists(fallback_path):
        print(f"使用历史壁纸: {fallback_path}")
        return fallback_path
    
    print("获取壁纸失败")
    return None


def get_dynamic_bing_image():
    try:
        tmp_path = Path(os.environ.get('USERPROFILE', '')) / 'AppData' / 'Local' / 'Packages'
        if not tmp_path.exists():
            return None
        
        dy_folders = [f for f in tmp_path.iterdir() if 'DynamicTheme' in f.name and f.is_dir()]
        if not dy_folders:
            return None
        
        dynamic_theme_path = dy_folders[0] / 'LocalState' / 'Bing'
        if not dynamic_theme_path.exists():
            return None
        
        # 获取最新的图片文件
        image_files = []
        for ext in ['.jpg', '.png', '.jpeg']:
            image_files.extend(dynamic_theme_path.glob(f'*{ext}'))
        
        if not image_files:
            return None
        
        # 按修改时间排序，取最新的
        latest_file = max(image_files, key=lambda x: x.stat().st_mtime)
        return str(latest_file)
    except Exception as e:
        print(f"读取DynamicTheme缓存失败: {e}")
        return None


def run():
    print("="*50)
    print("开始执行FadeTop壁纸更新程序")
    print("="*50)
    
    try:
        motto = get_config()
        print(f"读取到座右铭: {motto}")
        
        # 获取壁纸
        image_path = get_bing_image()
        
        # 终止FadeTop
        kill_FadeTop()
        time.sleep(1)  # 等待进程完全退出
        
        # 修改壁纸配置
        if image_path:
            change_wallpaper(motto, image_path)
        
        # 启动FadeTop
        start_FadeTop()
        time.sleep(2)  # 等待启动完成
        
        # 修复：使用兼容所有Windows版本的进程检查方式
        fade_top_running = False
        try:
            # 方案1：使用tasklist /FI过滤
            tasklist_output = subprocess.check_output(
                ['tasklist', '/FI', 'IMAGENAME eq FadeTop.exe', '/FO', 'CSV'],
                encoding='gbk',
                errors='ignore'
            )
            lines = [line.strip() for line in tasklist_output.split('\n') if line.strip()]
            fade_top_running = len(lines) > 1 and 'FadeTop.exe' in tasklist_output
        except Exception as e:
            # 方案2：降级使用os.popen
            print(f"⚠️ tasklist检查进程失败，使用降级方案: {e}")
            for line in os.popen('tasklist'):
                if 'FadeTop.exe' in line:
                    fade_top_running = True
                    break
        
        # 输出检查结果
        if fade_top_running:
            print("✅ FadeTop 已成功启动并应用新配置")
        else:
            print("⚠️ 未检测到FadeTop进程，请手动确认是否启动成功")
            
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n程序执行完成！")

if __name__ == "__main__":
    # 设置控制台编码
    os.system('chcp 65001 > nul')
    run()
    # 防止窗口立即关闭
    input("\n按回车键退出...")
