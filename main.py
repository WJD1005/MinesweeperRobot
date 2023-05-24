import pyautogui
import random
import win32con
import win32gui
import time

# 系统显示缩放为100%时的方格大小
block_size = 16

# 棋盘位置
global left, top, right, bottom

# 格子数
global blocks_x, blocks_y

# 棋盘数组
global map_num
# 棋盘未定格子序号
global map_unsure
# 格子点击坐标
global map_click

# 重开标志
global reset_flag
# 成功标志
win_flag = 0

# 雷数
global mine_num


def init_game():
    global left, top, right, bottom
    global blocks_x, blocks_y
    global map_num, map_unsure
    global reset_flag, win_flag
    global mine_num
    map_num = []
    map_unsure = []
    # 主窗体类别
    class_name = 'TMain'
    # 主窗体名
    title_name = 'Minesweeper Arbiter '
    # 找窗口句柄
    hwnd = win32gui.FindWindow(class_name, title_name)
    if hwnd:
        # 获取位置
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        # 取地图位置
        left += 15
        top += 101
        right -= 15  # 实际上偏移右边1px
        bottom -= 43  # 实际上偏移下边1px
        # 计算格子数
        blocks_x = int((right - left) / block_size)
        blocks_y = int((bottom - top) / block_size)
        # 信息数组初始化
        for i in range(blocks_x * blocks_y):
            # 棋盘赋初值
            map_num.append(-1)
            # 全部未定
            map_unsure.append(i)
        # 重开标志位清零
        reset_flag = 0
        # 雷数
        if blocks_x == 8:
            mine_num = 10
        elif blocks_x == 16:
            mine_num = 40
        elif blocks_x == 30:
            mine_num = 99
        else:
            print('检测到为自定义模式')
            mine_num = int(input('请输入雷数：'))

        # 解决窗口被最小化
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # 窗口前置
        win32gui.SetForegroundWindow(hwnd)
        # 等待一下防止未被前置
        time.sleep(1)
        # 点一个
        click_random()


def run_game():
    global win_flag
    # 脚本循环开始
    while True:
        pyautogui.moveTo(100, 100)  # 防挡
        get_map()  # 读取扫雷窗口当前所有信息
        # 查看是否需要重开
        if reset_flag:
            break
        a = get_mine_num()
        solve_a()
        solve_b()
        b = get_mine_num()
        if a == b:
            click_random()  # 若a方法和b方法均不能判断，即出现需要蒙的情况，则随机点一个（无奈之举）
        if get_mine_num() == mine_num:
            for i in map_unsure:
                if map_num[i] == -1:
                    click(i)
            win_flag = 1
            break


def reset_game():
    global reset_flag
    reset_flag = 0
    time.sleep(0.1)  # 缓一缓
    pyautogui.press('f3')
    # 信息数组初始化
    for i in range(blocks_x * blocks_y):
        # 棋盘赋初值
        map_num.append(-1)
        # 全部未定
        map_unsure.append(i)
    # 重开标志位清零
    reset_flag = 0
    # 点一个
    click_random()


def get_map():
    global map_num
    global reset_flag
    # 截图棋盘
    map_img = pyautogui.screenshot('./img/map.png', region=(left, top, right - left, bottom - top))
    # 通过特征像素对所有格子进行判断，记录在数组中
    # -1：未开
    # 0：已开空白
    # 1~8：数字
    # 9：旗子
    # 10：雷
    # 11：炸雷
    # 遍历所有未定格子
    for i in map_unsure:
        # 格子索引坐标
        x = i % blocks_x
        y = i // blocks_x
        # 取特征像素
        pix = map_img.getpixel((x * 16 + 10, y * 16 + 12))
        # 对特征像素进行判断
        if pix == (192, 192, 192):
            # 第二特征像素
            pix_0 = map_img.getpixel((x * 16, y * 16))
            # 未开
            if pix_0 == (255, 255, 255):
                map_num[i] = -1
                continue
            # 已开空白
            elif pix_0 == (128, 128, 128):
                map_num[i] = 0
                continue
        # 1
        elif pix == (0, 0, 255):
            map_num[i] = 1
            continue
        # 2
        elif pix == (0, 128, 0):
            map_num[i] = 2
            continue
        # 3
        elif pix == (255, 0, 0):
            map_num[i] = 3
            continue
        # 4
        elif pix == (0, 0, 128):
            map_num[i] = 4
            continue
        # 5
        elif pix == (128, 0, 0):
            map_num[i] = 5
            continue
        # 6
        elif pix == (0, 128, 0):
            map_num[i] = 6
            continue
        elif pix == (0, 0, 0):
            # 第二特征像素判断
            pix_7 = map_img.getpixel((x * 16 + 7, y * 16 + 7))
            # 旗子
            if pix_7 == (255, 0, 0):
                map_num[i] = 9
                continue
            elif pix_7 == (255, 255, 255):
                pix_10 = map_img.getpixel((x * 16 + 2, y * 16 + 2))
                # 雷
                if pix_10 == (192, 192, 192):
                    map_num[i] = 10
                    # 重开标志
                    reset_flag = 1
                    continue
                # 炸雷
                elif pix_10 == (255, 0, 0):
                    map_num[i] = 11
                    # 重开标志
                    reset_flag = 1
                    continue
            # 7
            elif pix_7 == (192, 192, 192):
                map_num[i] = 7
                continue
        # 8
        elif pix == (128, 128, 128):
            map_num[i] = 8
            continue

    # 打印棋盘检查
    # for i in range(blocks_x * blocks_y):
    #     print('%d,' % map_num[i])


def click(index):
    x = (index % blocks_x * block_size) + left + (block_size // 2)
    y = (index // blocks_x * block_size) + top + (block_size // 2)
    pyautogui.click(x, y, _pause=False)


def click_right(index):
    x = (index % blocks_x * block_size) + left + (block_size // 2)
    y = (index // blocks_x * block_size) + top + (block_size // 2)
    pyautogui.rightClick(x, y, _pause=False)
    if index in map_unsure:
        map_unsure.remove(index)
    map_num[index] = 9


def click_mid(index):
    x = (index % blocks_x * block_size) + left + (block_size // 2)
    y = (index // blocks_x * block_size) + top + (block_size // 2)
    pyautogui.middleClick(x, y, _pause=False)
    map_unsure.remove(index)


def click_random():
    while True:
        i = random.randint(0, blocks_x * blocks_y - 1)
        if i in map_unsure:
            if map_num[i] == -1:
                click(i)
                break


def get_mine_num():
    # 获取当前窗口有多少颗雷
    mine_num = 0
    for num in map_num:
        if num == 9:
            mine_num += 1
    return mine_num


def check_8(index):
    # 返回第index个格子，周围8个格子中，空白格子+雷格子序号的列表
    # 若第index个格子周围只有标记出的雷，没有空格子，说明该格子无须再判断，返回None
    if index not in map_unsure:
        return
    i, j = index // blocks_x, index % blocks_x
    up = [1, 0][i == 0]
    down = [2, 1][i == blocks_y - 1]
    left_ = [1, 0][j == 0]
    right_ = [2, 1][j == blocks_y - 1]
    gezi_8 = []
    for x in range(i - up, i + down):
        for y in range(j - left_, j + right_):
            if x != i or y != j:
                gezi_8.append(x * blocks_x + y)
    temp = []
    booms = 0
    for i in gezi_8:
        if map_num[i] >= 0 and map_num[i] != 9:
            temp.append(i)
    for tem in temp:
        gezi_8.remove(tem)

    for i in gezi_8:
        if map_num[i] == 9:
            booms += 1
    if booms == map_num[index] and len(gezi_8) == map_num[index]:
        map_unsure.remove(index)
        return
    elif booms == map_num[index] and len(gezi_8) != map_num[index]:
        click_mid(index)
        return
    else:
        return gezi_8


def check_4(index):
    # 返回第index个格子，上下左右8个格子中，空白格子+雷格子序号的列表
    # 若第index个格子周围只有标记出的雷，没有空格子，说明该格子无须再判断，返回None
    if index not in map_unsure:
        return
    gezi_4 = [index - blocks_x, index - 1, index + 1, index + blocks_x]
    temp = []
    if index % blocks_x == 0:
        gezi_4.remove(index - 1)
    if (index + 1) % blocks_x == 0:
        gezi_4.remove(index + 1)
    if 0 <= index <= blocks_x - 1:
        gezi_4.remove(index - blocks_x)
    if (blocks_x * blocks_y) - blocks_x <= index <= (blocks_x * blocks_y) - 1:
        gezi_4.remove(index + blocks_x)
    for m in gezi_4:
        if map_num[m] == 0:
            temp.append(m)
        if map_num[m] == -1:
            temp.append(m)
        if map_num[m] == 9:
            temp.append(m)
    for m in temp:
        gezi_4.remove(m)
    return gezi_4


def get_blank_boom(index):
    # 得到index格子的周围8个格子中，雷格子、空白格子列表，不算已经点开的格子
    booms, blank = [], []
    arr_8 = check_8(index)
    if arr_8:
        for i in arr_8:
            if map_num[i] == -1:
                blank.append(i)
            if map_num[i] == 9:
                booms.append(i)
        if blank:
            return [booms, blank]
        else:
            return


def solve_a():
    # 中心找空和找雷
    # 遍历所有未定格子
    for i in map_unsure:
        num = map_num[i]
        if 0 < num < 9:
            booms, blank = 0, 0
            arr = check_8(i)
            if arr:
                for j in arr:
                    if map_num[j] == -1:
                        blank += 1
                    if map_num[j] == 9:
                        booms += 1
                if booms + blank == num:
                    for j in arr:
                        if map_num[j] == -1:
                            click_right(j)  # 找到是雷
                elif booms == num and blank != 0:
                    click_mid(i)  # 中心的空格是空


def solve_b():
    # 重叠找空和找雷，需拆分为 找偏移九宫格、对比偏移九宫格空白格+雷格
    for i in map_unsure:
        num = map_num[i]
        if 0 < num < 9:
            booms_blank = get_blank_boom(i)  # 存储雷格子、空白格子的列表
            if not booms_blank:
                continue
            arr_4 = check_4(i)  # arr为index的四个方向偏移序号列表
            for j in arr_4:
                booms_blank_4 = get_blank_boom(j)  # 存储偏移格子的 雷格子、空白格子的列表
                if not booms_blank_4:
                    continue
                result = all(elem in booms_blank_4[1] for elem in booms_blank[1])  # 偏移空格列表包含中心空格列表
                if result:
                    m = map_num[j] - len(booms_blank_4[0])  # 偏移剩余雷数
                    n = map_num[i] - len(booms_blank[0])  # 中心剩余雷数
                    if m == n:
                        # 如果 中心剩余雷数 == 偏移剩余雷数，则中心不重叠的地方都是空格
                        for elem in booms_blank_4[1]:
                            if elem not in booms_blank[1]:
                                # print('重叠找空', elem)
                                click(elem)
                    if (m - n) >= (len(booms_blank_4[1]) - len(booms_blank[1])) and m != n:
                        # 如果 偏移剩余雷数 - 中心剩余雷数 == 偏移空格数 - 中心空格数，则中心不重叠的地方都是雷
                        for elem in booms_blank_4[1]:
                            if elem not in booms_blank[1]:
                                # print('重叠找雷', elem)
                                click_right(elem)


if __name__ == '__main__':
    # print('请将系统显示缩放更改为100%，按回车键开始运行。')
    # input()
    init_game()
    while win_flag == 0:
        run_game()
        if reset_flag and win_flag == 0:
            reset_game()
    print('成功')
