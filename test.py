import time


from PIL import ImageGrab
import pygetwindow as gw


def get_window_pos_size_by_title(title):
    try:
        window = gw.getWindowsWithTitle(title)[0]
        window.restore()
        time.sleep(0.5)
        x = window.left
        y = window.top
        width = window.width
        height = window.height
        print("找到了窗口", title)

        return x, y, width, height
    except:
        print("找不到 窗口", title)


x, y, width, height = get_window_pos_size_by_title("社保云缴费")
print("窗口位置：({}, {})".format(x, y))
print("窗口大小：{}x{}".format(width, height))

if __name__ == '__main__':
    x, y, width, height = get_window_pos_size_by_title("社保云缴费")
    im = ImageGrab.grab(bbox=(x, y, x + width, y + height))
    path = './result/{a}_{b}_{c}.jpg'.format(a='xiaoliang', b='1231564', c='未知')
    im.save(path)
    time.sleep(1)

    # 读取文件 并显示
    #
    # im = ImageGrab.grab()
    # im.save('./res/screen.png', 'png')
    #
    # img_rgb = cv2.imread('./res/screen.png')
    #
    # # 所有操作在灰度版中进行
    # img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    # template = cv2.imread('./res/xiaochengxu.png', 0)
    #
    # res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    # threshold = 0.7
    # loc = np.where(res >= threshold)
    #
    # for pt in zip(*loc[::-1]):
    #     print(pt[0], pt[1])
    #     time.sleep(0.5)
    #     pyautogui.moveTo(pt[0] + template.shape[0] / 2, pt[1] + template.shape[1] / 2)
    #     # pyautogui.moveTo(pt[0], pt[1])
    #     # pyautogui.doubleClick(pt[0] + template.shape[0] / 2, pt[1] + template.shape[1] / 2)
    #     pass
    # print('找到 小程序')
    # pass
