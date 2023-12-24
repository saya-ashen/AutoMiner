import time
from enum import Enum

import mouse
import win32api
import win32con
import win32gui  # pip install pywin32
import win32ui
from pymem import *


class CellType(Enum):
    UNKNOWN = 0
    EMPTY = 0x0F
    BOMB = 0x8F

    @staticmethod
    def get(value):
        if value == 0x0F:
            return CellType.EMPTY
        elif value == 0x8F:
            return CellType.BOMB
        else:
            return CellType.UNKNOWN


class Cell:
    def __init__(self, x, y, width, height, type: CellType, addr):
        color_map = {
            CellType.EMPTY: (3, 192, 60),
            CellType.BOMB: (255, 0, 0),
        }
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.type: CellType = type
        self.color = color_map.get(type, None)
        self.addr = addr


class WinmineCracker:
    def __init__(self):
        self.pymem = Pymem("winmine.exe")
        self.addr_base = self.pymem.base_address
        self.bomb_counts_offset = 22180
        self.addr_bomb_counts = self.addr_base + self.bomb_counts_offset
        self.timer_offset = 22428
        self.board_offset = 21344
        self.width_offset = 0x5334
        self.height_offset = 0x5338
        self.time_adder_offset = 0x2FF5
        self.addr_timer = self.addr_base + self.timer_offset
        self.fixed_with = 32
        self.addr_board_fist = self.addr_base + self.board_offset

    # 读取格子
    def get_board_types(self):
        board_ = self.pymem.read_bytes(self.addr_board_fist, self.board_size())
        board = []
        for i in range(0, self.height_count()):
            line = []
            for j in range(1, self.width_count() + 1):
                line.append(
                    (
                        CellType.get(board_[i * self.fixed_with + j]),
                        self.addr_board_fist + i * self.fixed_with + j,
                    )
                )
                # line.append(board_[i * self.fixed_with + j])
            board.append(line)
        return board

    def width_count(self):
        return self.pymem.read_int(self.addr_base + self.width_offset)

    def height_count(self):
        return self.pymem.read_int(self.addr_base + self.height_offset)

    def board_size(self):
        return self.fixed_with * self.height_count()

    # 读取地雷数量
    def get_counts(self):
        counts = self.pymem.read_int(self.addr_bomb_counts)
        return counts

    # 读取计时器
    def get_timer(self):
        timer = self.pymem.read_int(self.addr_timer)
        return timer

    # 用Nop填充增加时间的代码
    def fill_nop(self):
        replace_bytes = bytes([0x90, 0x90, 0x90, 0x90, 0x90, 0x90])
        self.pymem.write_bytes(
            self.addr_base + self.time_adder_offset, replace_bytes, len(replace_bytes)
        )

    # 修改游戏时间
    def set_timer(self, time):
        self.pymem.write_int(self.addr_timer, time)

    # 读取游戏状态
    def get_game_status(self):
        return self.pymem.read_int(self.addr_base + 0x5160)


class Board:
    def __init__(self, board_types):
        self.x = 15
        self.y = 103
        cell_width = 16
        self.width = len(board_types[0]) * cell_width
        self.height = len(board_types) * cell_width
        self.lines = []
        for i, line in enumerate(board_types):
            line_ = []
            for j, cell_info in enumerate(line):
                line_.append(
                    Cell(
                        self.x + j * cell_width,
                        self.y + i * cell_width,
                        cell_width,
                        cell_width,
                        cell_info[0],
                        cell_info[1],
                    )
                )
            self.lines.append(line_)

    def is_in_board(self, x, y):
        if (
            x > self.x
            and x < self.x + self.width
            and y > self.y
            and y < self.y + self.height
        ):
            return True
        return False


class Drawer:
    def __init__(
        self,
        hwnd,
    ):
        self.hwnd = hwnd
        self.hdc = win32gui.GetWindowDC(self.hwnd)
        self.dc = win32ui.CreateDCFromHandle(self.hdc)
        self.pens = {}
        self.brushes = {}
        red_pen = win32ui.CreatePen(win32con.PS_SOLID, 1, win32api.RGB(255, 0, 0))
        self.pens[(255, 0, 0)] = red_pen
        red_brush = win32ui.CreateBrush(win32con.BS_SOLID, win32api.RGB(255, 0, 0), 0)
        self.brushes[(255, 0, 0)] = red_brush
        green_pen = win32ui.CreatePen(win32con.PS_SOLID, 1, win32api.RGB(3, 192, 60))
        self.pens[(3, 192, 60)] = green_pen
        green_brush = win32ui.CreateBrush(
            win32con.BS_SOLID, win32api.RGB(3, 192, 60), 0
        )
        self.brushes[(3, 192, 60)] = green_brush

    def draw_rect(self, x, y, width, color):
        # 创建一个画笔对象

        pen = self.pens[color]
        self.dc.SelectObject(pen)
        brush = self.brushes[color]
        self.dc.SelectObject(brush)

        # 绘制矩形
        self.dc.Rectangle((x, y, x + width, y + width))

    # 清理dc，释放窗口句柄
    def __del__(self):
        self.dc.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, self.hdc)

    def draw_board(self, board: Board):
        # 刷新窗口
        win32gui.InvalidateRgn(self.hwnd, None, True)
        win32gui.UpdateWindow(self.hwnd)
        time.sleep(0.01)
        # 绘制
        for line in board.lines:
            for cell in line:
                if cell.type != CellType.UNKNOWN:
                    self.draw_cell(cell)

    def draw_cell(self, cell: Cell):
        self.draw_rect(cell.x + 3, cell.y + 3, cell.width - 6, cell.color)


def on_click(hwnd, cracker, drawer):
    time.sleep(0.1)
    try:
        board_types = cracker.get_board_types()
    except:
        return
    board = Board(board_types)
    x, y = mouse.get_position()
    winmine_position = win32gui.GetWindowRect(hwnd)
    relative_x = x - winmine_position[0]
    relative_y = y - winmine_position[1]
    if board.is_in_board(relative_x, relative_y):
        drawer.draw_board(board)


# 一键自动扫雷
def auto_crack(hwnd, cracker, drawer: Drawer):
    time.sleep(0.1)
    board_types = cracker.get_board_types()
    board = Board(board_types)
    for line in board.lines:
        for cell in line:
            if cell.type == CellType.EMPTY:
                x = cell.x + 5
                y = cell.y + 5 - 40
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, y << 16 | x)
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, y << 16 | x)
            elif cell.type == CellType.BOMB:
                x = cell.x + 5
                y = cell.y + 5 - 40
                win32gui.SendMessage(hwnd, win32con.WM_RBUTTONDOWN, 0, y << 16 | x)
                win32gui.SendMessage(hwnd, win32con.WM_RBUTTONUP, 0, y << 16 | x)

            if cracker.get_game_status() == 3:
                return
            if cell.type == CellType.EMPTY:
                drawer.draw_cell(cell)

            time.sleep(0.01)


if __name__ == "__main__":
    while True:
        try:
            hwnd = win32gui.FindWindow(None, "扫雷")
            if hwnd == 0:
                print("未检测到扫雷窗口，请打开扫雷！")
                time.sleep(5)
                continue
            print("检测到扫雷窗口，开始监听鼠标事件...")
        except:
            print("未知错误！")
            time.sleep(1)
            continue

        try:
            cracker = WinmineCracker()
            drawer = Drawer(hwnd)
            mouse.on_button(
                on_click, (hwnd, cracker, drawer), buttons=("left",), types=("up",)
            )
            while True:
                key = input("输入q退出，输入a自动扫雷，输入t停止计时并设置计时器为0，点击任意格子显示地雷分布：")
                if key == "q":
                    break
                elif key == "a":
                    print("开始扫雷, 鼠标请不要放到扫雷棋盘内...")
                    auto_crack(hwnd, cracker, drawer)
                    print("扫雷完成！")
                elif key == "t":
                    cracker.fill_nop()
                    cracker.set_timer(0)
                    print("设置成功！")
                else:
                    print("输入错误！")
        except:
            print("未知错误！")
            time.sleep(1)
