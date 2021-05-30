"""

统一一下, 以下所有横纵坐标都是先行再列.
"""

import time
from typing import Any, List
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

    def mouse_click_left() -> None:
        """ 鼠标左击. """
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    def mouse_click_right() -> None:
        """ 鼠标右击. """
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)

    def keybord_click(code: int) -> None:
        """ 键盘单击. """
        win32api.keybd_event(code, 0, 0, 0)
        win32api.keybd_event(code, 0, win32con.KEYEVENTF_KEYUP, 0)


class Game:
    """ 连连看游戏类 """

    CELL_TOP, CELL_LEFT = 141, 64
    CELL_HEIGHT, CELL_WIDTH = 50, 40
    CELL_MAX_ROW, CELL_MAX_COL = 16, 9

    _ID_UP, _ID_DOWN = CELL_HEIGHT / 2 - 2, CELL_HEIGHT / 2 + 2
    _ID_LEFT, _ID_RIGHT = CELL_WIDTH / 2 - 2, CELL_WIDTH / 2 + 2

    def __init__(self, win: Window) -> None:
        self.win = win
    
    def cell_point(self, cell_x: int, cell_y: int) -> List[int]:
        """ 格子的左上坐标点. 用 List 而不用 Tuple 返回是为了允许直接修改元素 (Tuple 内部元素不可变). """
        return [self.CELL_TOP + self.CELL_HEIGHT * cell_x, self.CELL_LEFT + self.CELL_WIDTH * cell_y]
    
    def cell_id(self, cell_x: int, cell_y: int) -> int:
        """ 格子识别码 (根据几个像素确定). """
        hash = 0
        px, py = self.cell_point(cell_x, cell_y)
        for dx in [self._ID_UP, self._ID_DOWN]:
            for dy in [self._ID_LEFT, self._ID_RIGHT]:
                rgba = self.win.screen[px + dx][py + dy]
                v = 1
                for i in range(3): v = (v << 8) + rgba[i]
                hash = (hash << 8) + (v % 233)  # 随便取模个小于 256 的素数 233
        return hash

    def cal_board(self):
        board, dic = [], {}
        for i in range(self.CELL_MAX_ROW):
            line = []
            board.append(line)
            for j in range(self.CELL_MAX_COL):
                cid = self.cell_id(i, j)
                line.append(cid)
                dic[cid] = dic.get(cid, 0) + 1


def main():
    win = Window()


if __name__ == '__main__':
    try: main()
    except pywintypes.error as e:
        if e.args[2] == '无效的窗口句柄。': print('错误: 未检测到正在运行的 "连连看 V3.0".')
        else: raise e