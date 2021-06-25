"""

统一一下, 以下所有横纵坐标都是先行再列.
"""

import copy
import time
import random
from typing import Any, Dict, List, Optional
import win32api, win32con, win32gui, win32ui, pywintypes


class Window:
    """ 连连看窗口类 """

    def __init__(self) -> None:
        self.get_window()
        self.get_screenshot()

    def get_window(self) -> int:
        """ 找到窗口并设置焦点 """
        classname = "ThunderRT6FormDC"
        titlename = "连连看V3.0"
        self.hwnd = win32gui.FindWindow(classname, titlename)
        win32gui.SetForegroundWindow(self.hwnd)

    def get_screenshot(self) -> None:
        """ 截图并获得 RGBA 矩阵. """
        # 截图
        self.left, self.top, self.right, self.bottom = win32gui.GetWindowRect(self.hwnd)
        self.width, self.height = self.right - self.left, self.bottom - self.top
        
        hwnd_dc = win32gui.GetWindowDC(self.hwnd)                       # 返回句柄窗口的设备环境 (覆盖整个窗口, 包括标题栏菜单等)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)                    # 创建设备描述表
        save_dc = mfc_dc.CreateCompatibleDC()                           # 创建内存设备描述表
        bitmap = win32ui.CreateBitmap()                                 # 创建位图对象准备保存图片
        bitmap.CreateCompatibleBitmap(mfc_dc, self.width, self.height)  # 为 bitmap 开辟存储空间
        save_dc.SelectObject(bitmap)                                    # 将截图保存到 bitmap 中
        save_dc.BitBlt((0,0), (self.width, self.height), mfc_dc, (0, 0), win32con.SRCCOPY)    # 保存 bitmap 到内存设备描述表
        bitmap.SaveBitmapFile(save_dc, "screenshot.bmp")                # 保存截图

        # 获得 RGBA 矩阵
        bits: tuple = bitmap.GetBitmapBits(False)
        assert len(bits) == self.width * self.height * 4
        self.screen = []
        for r in range(self.height):
            sr = 4 * r * self.width
            line = []
            self.screen.append(line)
            for c in range(self.width):
                sc = sr + 4 * c
                rgba = [b if b >= 0 else b + 256 for b in bits[sc: sc + 4]]
                rgba[0], rgba[2] = rgba[2], rgba[0] # GetBitmapBits 返回的是 BGRA, 所以要转一下
                line.append(rgba)
        # self.sc_info = bitmap.GetInfo()

        # 善后
        win32gui.DeleteObject(bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwnd_dc)

    def mouse_teleport(self, top: int, left: int) -> None:
        """ 把鼠标挪到指定位置 (坐标原点为窗口左上角). """
        win32api.SetCursorPos([self.left + left, self.top + top])

    def mouse_click_left(self) -> None:
        """ 鼠标左击. """
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    def mouse_click_right(self) -> None:
        """ 鼠标右击. """
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)

    def keybord_click(self, code: int) -> None:
        """ 键盘单击. """
        win32api.keybd_event(code, 0, 0, 0)
        win32api.keybd_event(code, 0, win32con.KEYEVENTF_KEYUP, 0)


class Game:
    """ 连连看游戏类 """

    CELL_TOP, CELL_LEFT = 141, 64
    CELL_HEIGHT, CELL_WIDTH = 50, 40
    CELL_MAX_ROW, CELL_MAX_COL = 9, 16
    
    @staticmethod
    def __gray_of_rgba(rgba: List[int]) -> int:
        """ 快速计算 RGBA 的灰度. """
        r, g, b, a = rgba
        return (r * 38 + g * 75 + b * 15) >> 7
    
    @classmethod
    def __init_id_points(cls):
        """ 初始化定位格子 ID 的像素点. """
        res = []
        for dx in range(11, cls.CELL_HEIGHT - 10, 4):
            for dy in range(6, cls.CELL_WIDTH - 5, 4):
                res.append([dx, dy])
        return res

    @staticmethod
    def __intersection_of_range(range1: List[int], range2: List[int]) -> List[int]:
        """ 取两个范围的交集. 范围都是左闭右开 (方便遍历). """
        if range1[0] > range2[0]: range1, range2 = range2, range1
        [a1, b1], [a2, b2] = range1, range2
        if b1 <= a2: return [-1, -1]
        return [a2, min(b1, b2)]

    def __init__(self, win: Window) -> None:
        self.__id_points = self.__init_id_points()
        self.win = win
        self.refresh()
        self.print()

    def __moveable_v_h(self, p: List[int]) -> List[int]:
        """ 点在垂直和水平方向的可移动范围. """
        top, bottom = p[0], p[0]
        while top >= 1 and self.board[top - 1][p[1]] == 0: top -= 1
        while bottom <= self.CELL_MAX_ROW and self.board[bottom + 1][p[1]] == 0: bottom += 1
        left, right = p[1], p[1]
        while left >= 1 and self.board[p[0]][left - 1] == 0: left -= 1
        while right <= self.CELL_MAX_COL and self.board[p[0]][right + 1] == 0: right += 1
        return [top, bottom + 1, left, right + 1]

    def __path(self, from_p: List[int], to_p: List[int]) -> Optional[List[List[int]]]:
        """ 计算两点之间是否有可行路径. """
        if from_p[0] == to_p[0] and abs(from_p[1] - to_p[1]) == 1: return [from_p, to_p]
        if from_p[1] == to_p[1] and abs(from_p[0] - to_p[0]) == 1: return [from_p, to_p]
        t1, b1, l1, r1 = self.__moveable_v_h(from_p)
        t2, b2, l2, r2 = self.__moveable_v_h(to_p)
        for r in range(*self.__intersection_of_range([t1, b1], [t2, b2])):
            for c in range(*sorted([from_p[1], to_p[1]])):
                if self.board[r][c] != 0: break
            else: return [from_p, [r, from_p[1]], [r, to_p[1]], to_p]
        for c in range(*self.__intersection_of_range([l1, r1], [l2, r2])):
            for r in range(*sorted([from_p[0], to_p[0]])):
                if self.board[r][c] != 0: break
            else: return [from_p, [from_p[0], c], [to_p[0], c], to_p]
        return None
    
    def __cell_point(self, cell_x: int, cell_y: int) -> List[int]:
        """ 格子的左上坐标点. 用 List 而不用 Tuple 返回是为了允许直接修改元素 (Tuple 内部元素不可变). """
        return [self.CELL_TOP + self.CELL_HEIGHT * cell_x, self.CELL_LEFT + self.CELL_WIDTH * cell_y]
    
    def __cell_id(self, cell_x: int, cell_y: int) -> int:
        """ 格子识别码 (根据几个像素确定). """
        def is_px_black():
            rgba = self.win.screen[px + dx][py + dy]
            gray = self.__gray_of_rgba(rgba)
            return gray < 125
        hash = 0
        black_cnt = 0
        px, py = self.__cell_point(cell_x, cell_y)
        _dx1, _dx2, _dy1, _dy2 = 3, self.CELL_HEIGHT - 4, 3, self.CELL_WIDTH - 4
        for dx, dy in [[_dx1, _dy1], [_dx1, _dy2], [_dx2, _dy1], [_dx2, _dy2]]:
            if is_px_black(): return 0
        for dx, dy in [[0, 0], [self.CELL_HEIGHT - 1, 0], [0, self.CELL_WIDTH - 1]]:
            if not is_px_black(): return 0
        for dx, dy in self.__id_points:
            is_black = is_px_black()
            if is_black: black_cnt += 1
            hash = (hash << 1) | is_black
        if black_cnt == 0 or black_cnt == len(self.__id_points): return 0
        return hash

    def refresh(self):
        """ 计算出当前棋盘局面. """
        self.win.get_screenshot()
        board, dic = [], {}
        board.append([0] * (self.CELL_MAX_COL + 2))
        for i in range(self.CELL_MAX_ROW):
            line = [0]
            board.append(line)
            for j in range(self.CELL_MAX_COL):
                cid = self.__cell_id(i, j)
                line.append(cid)
                dic[cid] = dic.get(cid, 0) + 1
            line.append(0)
        board.append([0] * (self.CELL_MAX_COL + 2))
        vis = {}
        vis_cnt = 0
        for i, line in enumerate(board):
            for j in range(len(line)):
                h = line[j]
                if h == 0: continue
                # if dic[h] != 4: print('[%d, %d]' % (i, j), dic[h])
                if dic[h] & 1 == 1 or dic[h] > 4: line[j] = 0
                elif h in vis: line[j] = vis[h]
                else:
                    vis_cnt += 1
                    line[j] = vis_cnt
                    vis[h] = vis_cnt
        self.type_cnt = len(vis)
        self.board = board
    
    def hint(self):
        """ 提示一次. """
        card_pos: Dict[int, List[List[int]]] = {}
        for r in range(len(self.board)):
            for c in range(len(self.board[0])):
                cid = self.board[r][c]
                if cid == 0: continue
                card_pos.setdefault(cid, []).append([r, c])
        assert self.type_cnt == len(card_pos)
        same = [v for v in card_pos.values()]
        random.shuffle(same)
        for points in same:
            size = len(points)
            for i in range(size):
                for j in range(i + 1, size):
                    path = self.__path(points[i], points[j])
                    if path is not None:
                        self.print(hint=path)
                        return
        print('Hint not found.')


    def print(self, hint: Optional[List[List[int]]] = None):
        """ 打印棋盘. """
        to_print = self.board
        if hint is not None:
            to_print = copy.deepcopy(self.board)
            for px, py in hint[1:-1]: to_print[px][py] = '+'
            to_print[hint[0][0]][hint[0][1]] = '*'
            to_print[hint[-1][0]][hint[-1][1]] = '*'

        print('牌面类型数量:', self.type_cnt)
        header_footer = '+' + ('-' * (3 * self.CELL_MAX_COL + 6 + 1)) + '+'
        print(header_footer)
        for line in to_print:
            print(end='|')
            for n in line:
                if type(n) == int: print(('%3d' % n) if n > 0 else '   ', end='')
                else: print('  %s' % n, end='')
            print(' |')
        print(header_footer)


def main():
    win = Window()
    game = Game(win)
    while True:
        game.refresh()
        game.hint()
        time.sleep(4)
    # win.mouse_teleport(game.CELL_TOP + 41, game.CELL_LEFT + 12 + game.CELL_WIDTH)


if __name__ == '__main__':
    try: main()
    except pywintypes.error as e:
        if e.args[2] == '无效的窗口句柄。': print('错误: 未检测到正在运行的 "连连看V3.0".')
        else: raise e
