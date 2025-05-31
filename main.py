import win32gui, os, json, time, re

version = "1.0.0"

def get_netease_music_info():
    # 网易云音乐主窗口类名
    MAIN_WINDOW_CLASS = "OrpheusBrowserHost"

    try:
        # 尝试查找主窗口句柄
        hwnd = win32gui.FindWindow(MAIN_WINDOW_CLASS, None)
        if not hwnd:
            return {"error": "未找到网易云音乐窗口"}

        # 获取窗口标题
        title = win32gui.GetWindowText(hwnd)

        # 跳过空标题或默认标题
        if not title or title == "网易云音乐":
            return {"error": "窗口标题无播放信息"}

        # 格式1: 歌曲名 - 歌手
        if " - " in title:
            parts = title.split(" - ", 1)
            return {"song": parts[0].strip()}

        # 无法解析的标题
        return {"song": title, "artist": "未知歌手"}

    except Exception as e:
        return {"error": f"系统错误: {str(e)}"}

def write_file(song_info):
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    num = 1

    if os.path.exists(config["save_path"]):
        with open(config["save_path"], 'r', encoding='utf-8') as f:
            info = f.readlines()
        num = len(info) + 1

    with open(config["save_path"], 'a+', encoding='utf-8') as file:
        file.write(f'{num}{config["order_number"]}{song_info["song"]}\n')

def switch_save():
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    config["save_to_file"] = not config.get("save_to_file", False)

    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

    return config["save_to_file"]

def set_path():
    path = input("请输入保存路径(后缀为.txt，例：D:\music.txt)：")
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    if not path:  # 处理空路径
        main()
    else:
        if re.match(r'^[a-zA-Z]:', path):
            normalized_path = os.path.normpath(path) # 标准化路径

            if os.path.isdir(normalized_path):
                path = os.path.join(normalized_path, "songs.txt")

            elif normalized_path.endswith(os.sep) or normalized_path.endswith('/'):
                path = os.path.join(normalized_path, "songs.txt")

            else:
                dir_path = os.path.dirname(normalized_path)

                if dir_path and not os.path.exists(dir_path):
                    os.makedirs(dir_path, exist_ok=True)

                path = normalized_path

        else:
            path = os.path.join(os.getcwd(), path)

        config["save_path"] = path

        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)

    os.system("cls")
    main()

def get_info():
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    if config.get("save_path", None) == None and config.get("save_to_file", True):
        print("请先设置保存路径或关闭保存至文件...")
        time.sleep(3)
        os.system("cls")
        main()
    else:
        song_info = get_netease_music_info()

        if "error" in song_info:
            print("错误:", song_info["error"])
            input("按任意键继续...")
            main()
        else:
            print(f"当前歌曲: {song_info['song']}")
            if config.get("save_to_file", True):
                write_file(song_info)

def get_info_loop():
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    if config.get("save_path", None) == None and config.get("save_to_file", True):
        print("请先设置保存路径或关闭保存至文件...")
        time.sleep(3)
        os.system("cls")
    else:
        last_song = None
        while True:
            song_info = get_netease_music_info()
            song = song_info['song']

            if "error" in song_info:
                print("错误:", song_info["error"])
                input("按任意键继续...")
                break
            else:
                if song != last_song:
                    print(f"当前歌曲: {song}")
                    if config.get("save_to_file", True):
                        write_file(song_info)

            last_song = song
            time.sleep(int(config["rate"]))  # 每5秒检查一次

    main()

def main():
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    print(f'''1. 获取歌名
2. 自动获取歌名
3. 设置保存路径({config.get("save_path", None)})
4. 设置序号({config["order_number"]})
4. 开关保存至文件({"开启" if config.get("save_to_file", True) else "关闭"})
''')

    choice = input("选项：")
    if choice == "1":
        get_info()
        time.sleep(2)
        os.system("cls")
        main()
    elif choice == "2":
        print(f'正在自动获取歌曲信息(获取频率：{config["rate"]}s)...')
        print("按 Ctrl+C 退出")
        get_info_loop()
    elif choice == "3":
        set_path()
        time.sleep(2)
        os.system("cls")
        main()
    elif choice == "4":
        config["order_number"] = input("请设置序号：")
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        os.system("cls")
        main()
    elif choice == "5":
        status = switch_save()
        print(f'保存至文件状态: {"开启" if status else "关闭"}')
        time.sleep(2)
        os.system("cls")
        main()
    else:
        print("无效选项")
        time.sleep(2)
        os.system("cls")
        main()

if __name__ == "__main__":
    if not os.path.exists("config.json"):
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump({"save_to_file": True,
                       "order_number": ".",
                       "rate": 5
                       }, f, ensure_ascii=False, indent=4)
    else:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)

    song_info = get_netease_music_info()

    if "error" in song_info:
        print("错误:", song_info["error"])
        input("按任意键退出...")
    else:
        print(f"version: v{version}\n")
        main()