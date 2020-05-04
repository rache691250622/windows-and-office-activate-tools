import difflib
import json
import os
import re
import string
import subprocess
import traceback

import win32api
import win32con
import requests
from PyQt5.QtCore import *
from Init import *


def re_search(text, txt):
    value = ''
    a = re.compile(text)
    b = a.search(txt)
    if b is not None and len(b.groups()) != 0:
        value += (b.groups()[0] + " ")
    return value


def command(command):
    cmd = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
    result = cmd.wait()
    return result


def echo_command(cmd):  # 执行cmd并实时回显
    cmd = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
    for line in iter(cmd.stdout.readline, b''):
        echo = line.decode('gbk')
    # output = os.popen(cmd, "r")
    # echo = output.read().encode('gbk').decode('gbk')
    cmd = cmd.wait()
    return echo


class activate(QThread):
    show_text = pyqtSignal(str)

    def __init__(self, win):
        super(activate, self).__init__()
        self.win = win
        # 初始化参数
        self.n = 0
        # 服务器列表
        self.server_ip = ['54.223.212.31', 'kms.guowaifuli.com', 'mhd.kmdns.net', 'xykz.f3322.org', 'kms.xspace.in',
                          'zh.us.to', '106.186.25.239', '110.noip.me', '3rss.vicp.net:20439', '45.78.3.223',
                          'kms.chinancce.com', 'kms.didichuxing.com', 'skms.ddns.net', 'franklv.ddns.net',
                          'k.zpale.com', 'm.zpale.com', 'mvg.zpale.com', '122.226.152.230', '222.76.251.188',
                          'annychen.pw', 'heu168.6655.la', 'kms.aglc.cc', 'kms.landiannews.com', 'kms.shuax.com',
                          'kms.xspace.in', 'winkms.tk', 'wrlong.com', 'kms7.MSGuides.com', 'kms8.MSGuides.com',
                          'kms9.MSGuides.com']
        # 管理员权限
        self.admin = 'mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","",' \
                     '"runas",1)(window.close)&&exit '
        # windows key
        self.win_key = '{"win10_key":["TX9XD-98N7V-6WMQ6-BX7FG-H8Q99", "PVMJN-6DFY6-9CCP6-7BKTT-D3WVR", ' \
                       '"3KHY7-WNT83-DGQKR-F7HPR-844BM", "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH", ' \
                       '"YNXW3-HV3VB-Y83VG-KPBXM-6VH3Q", "8G9XJ-GN6PJ-GW787-MVV7G-GMR99", ' \
                       '"W269N-WFGWX-YVC9B-4J6C9-T83GX", "MH37W-N47XK-V7XM9-C7227-GCQG9", ' \
                       '"8Q36Y-N2F39-HRMHT-4XW33-TCQR4", "NKPM6-TCVPT-3HRFX-Q4H9B-QJ34Y", ' \
                       '"NPPR9-FWDCX-D2C8J-H872K-2YT43", "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4", ' \
                       '"NW6C2-QMPVW-D7KKK-3GKT6-VCFB2", "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ", ' \
                       '"WNMTR-4C88C-JK8YV-HQ7T2-76DF9", "2F77B-TNFGY-69QQF-B8YKP-D69TJ"], ' \
                       '"win8_key":["W269N-WFGWX-YVC9B-4J6C9-T83GX", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2", ' \
                       '"MH37W-N47XK-V7XM9-C7227-GCQG9", "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4", ' \
                       '"2WH4N-8QGBV-H22JP-CT43Q-MDWWJ", "NPPR9-FWDCX-D2C8J-H872K-2YT43", ' \
                       '"TX9XD-98N7V-6WMQ6-BX7FG-H8Q99", "WNMTR-4C88C-JK8YV-HQ7T2-76DF9", ' \
                       '"2F77B-TNFGY-69QQF-B8YKP-D69TJ", "KW3QN-KFYWP-3KTGY-W678X-RM2KV", ' \
                       '"9DDD3-84PXF-QNPXF-3PV8Q-G8XWD", "HV8NK-MCCM2-J7CDF-B46H7-46V3H", ' \
                       '"NQC2J-HCVC6-Y6DM4-DWJX3-2WD2Y", "NB78R-2KTRP-MY37P-WTJK2-J426M", ' \
                       '"27CKQ-6JNDK-FKCJV-97FCG-GMQG3", "2TPRN-B4BYT-JWGC8-YR4C7-46V3H"], ' \
                       '"win7_key":["KH2J9-PC326-T44D4-39H6V-TVPBY", "TFP9Y-VCY3P-VVH3T-8XXCC-MF4YK", ' \
                       '"236TW-X778T-8MV9F-937GT-QVKBB", "MT39G-9HYXX-J3V3Q-RPXJB-RQ6D7", ' \
                       '"W3KYM-BWXG4-M2PDC-V3K36-8QPJK", "PRP24-C66WQ-6KH43-83P9M-KRQK8", ' \
                       '"C2236-JBPWG-TGWVG-GC2WV-D6V6Q", "YH6QF-4R473-TDKKR-KD9CB-MQ6VQ", ' \
                       '"6F4BB-YCB3T-WK763-3P6YJ-BVH24", "6K2KY-BFH24-PJW6W-9GK29-TMPWP", ' \
                       '"TC48D-Y44RV-34R62-VQRK8-64VYG", "J6QGR-6CFJQ-C4HKH-RJPVP-7V83X", ' \
                       '"FJGCP-4DFJD-GJY49-VJBQ7-HYRR2", "PPH4P-JYK6Q-J46QP-3TRCD-QGKTG", ' \
                       '"Q8JXJ-8HDJR-X4PXM-PW99R-KTJ3H", "37X8Q-CJ46F-RB8XP-GJ6RK-RHYT7"]}'
        self.win_key = json.loads(self.win_key)
        self.Office_key = '{"Office 专业增强版 2019":"NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP", ' \
                          '"Office 标准版 2019":"6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK", ' \
                          '"Visio 专业版 2019":"9BGNQ-K37YR-RQHF2-38RQ3-7VCBB", ' \
                          '"Visio 标准版 2019":"7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2", ' \
                          '"Office 专业增强版 2016":"XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99", ' \
                          '"Office 标准版 2016":"JNRGM-WHDWX-FJJG3-K47QV-DRTFM", ' \
                          '"Visio 专业版 2016":"PD3PC-RHNGV-FXJ29-8JK7D-RJRJK", ' \
                          '"Visio 标准版 2019":"7WHWN-4T7MP-G96JF-G33KR-W8GF4", ' \
                          '"Office 专业增强版 2013":"YC7DK-G2NP3-2QQC3-J6H88-GVGXT", ' \
                          '"Office 标准版 2013":"KBKQT-2NMXY-JJWGP-M62JB-92CD4", ' \
                          '"Visio 专业版 2013":"C2FG9-N6J68-H8BTJ-BW3QX-RM3B3", ' \
                          '"Visio 标准版 2013":"J484Y-4NKBF-W2HMG-DBMJC-PGWR7"}'
        self.Office_key = json.loads(self.Office_key)

    def __del__(self):
        self.wait()

    def run(self):
        if self.win.windows_ver.currentText() != '':
            self.activate_windows()
        elif self.win.office_ver.currentText() != '':
            self.activate_office()
        else:
            win32api.MessageBox(0, '未选择版本', '错误', win32con.MB_ICONWARNING)
            return

    def activate_windows(self):
        self.show_text.emit('lock')
        # 检查版本
        ver_find = echo_command('ver')
        rever = r'(?:版本 )([^.]+)'
        ver = re_search(rever, ver_find)
        ver = 'Windows ' + ver
        ver_choose = self.win.windows_ver.currentText()
        ver_choose += ' '
        if ver != ver_choose:
            win32api.MessageBox(0, '选择的版本与系统不符，请重新选择', '错误', win32con.MB_ICONWARNING)
            self.show_text.emit('block')
            return
        # 卸载key
        uninstall_key = command(self.admin + '&&\ slmgr /upk')
        if uninstall_key != 0:
            win32api.MessageBox(0, '卸载key失败，请重新尝试', '错误', win32con.MB_ICONWARNING)
            self.show_text.emit('block')
            return
        key_input = self.win.windows_key_input.text()
        if key_input == '':
            # 安装key
            if ver == 'Windows 10 ':
                win_key = self.win_key['win10_key']
            elif ver == 'Windows 8 ':
                win_key = self.win_key['win8_key']
            elif ver == 'Windows 7 ':
                win_key = self.win_key['win7_key']
            for key in win_key:
                self.n = self.n + 1
                key_install = command(self.admin + '&&\ slmgr /ipk ' + key)
                if key_install == 0:
                    self.n = 0
                    break
                if self.n == 16:
                    self.n = 0
                    win32api.MessageBox(0, '预置key都无法使用，请自行输入key后重新激活', '错误', win32con.MB_ICONWARNING)
                    self.show_text.emit('block')
                    return
        else:
            key_install = command(self.admin + '&&\ slmgr /ipk ' + key_input)
            if key_install != 0:
                win32api.MessageBox(0, '输入的key不可用，请重新输入', '错误', win32con.MB_ICONWARNING)
                self.show_text.emit('block')
                return
        # 设置KMS服务器
        server_input = self.win.windows_server_input.text()
        if server_input == '':
            for server in self.server_ip:
                self.n = self.n + 1
                set_server = command(self.admin + '&&\ slmgr /skms ' + server)
                if set_server == 0:
                    self.n = 0
                    break
                if self.n == 30:
                    self.n = 0
                    win32api.MessageBox(0, '设置KMS服务器失败，请重新尝试', '错误', win32con.MB_ICONWARNING)
                    self.show_text.emit('block')
                    return
        else:
            set_server = command(self.admin + '&&\ slmgr /skms ' + server_input)
            if set_server != 0:
                win32api.MessageBox(0, '设置KMS服务器失败，请重新尝试', '错误', win32con.MB_ICONWARNING)
                self.show_text.emit('block')
                return
        # 激活
        activate = command(self.admin + '&&\ slmgr /ato')
        if activate == 0:
            win32api.MessageBox(0, '已成功激活系统', '成功', win32con.MB_OK)
        else:
            win32api.MessageBox(0, '激活失败，请重新尝试', '错误', win32con.MB_ICONWARNING)
        self.show_text.emit('block')
        return

    def activate_office(self):
        self.show_text.emit('lock')
        # 调用office cmd
        # 获取本地磁盘
        disk_list = []
        for c in string.ascii_uppercase:
            disk = c + ':\\'
            if os.path.isdir(disk):
                disk_list.append(disk)
        # 获取office路径
        for disk_list in disk_list:
            for root, dirs, files in os.walk(disk_list):
                if 'Microsoft Office' in dirs:
                    path1 = os.path.join(root, 'Microsoft Office')
                    path2 = os.listdir(path1)
                    a = difflib.get_close_matches('Office', path2, 5, cutoff=0.7)
                    path = path1 + '\\' + a[0]
                    break
            break
        cmd = self.admin + '"&&\ cd "' + path
        # 安装key
        ver = self.win.office_ver.currentText()
        key_input = self.win.office_key_input.text()
        if key_input == '':
            key_install = command(cmd + '&&\ cscript ospp.vbs /inpkey: ' + self.Office_key[ver])
            if key_install == 0:
                self.show_text.emit('已安装key为：' + key_input)
            else:
                self.show_text.emit('预置key都无法使用，请自行输入key后重新激活')
                win32api.MessageBox(0, '预置key都无法使用，请自行输入key后重新激活', '错误', win32con.MB_ICONWARNING)
                self.show_text.emit('block')
                return False
        else:
            key_install = command(cmd + '&&\ cscript ospp.vbs /inpkey:' + self.office_key_input)
            if key_install == 0:
                self.show_text.emit('已安装key为：' + self.Office_key[ver])
            else:
                self.show_text.emit('输入的key不可用，请重新输入')
                win32api.MessageBox(0, '输入的key不可用，请重新输入', '错误', win32con.MB_ICONWARNING)
                self.show_text.emit('block')
                return False
        # 设置kms服务器
        server_input = self.win.windows_server_input.text()
        if server_input == '':
            for server in self.server_ip:
                self.n = self.n + 1
                set_server = command(cmd + '&&\ cscript ospp.vbs /sethst:' + server)
                if set_server == 0:
                    self.show_text.emit('已将KMS服务器设置为：' + server)
                    self.n = 0
                    break
                if self.n == 30:
                    self.n = 0
                    self.show_text.emit('设置KMS服务器失败')
                    win32api.MessageBox(0, '设置KMS服务器失败，请重新尝试', '错误', win32con.MB_ICONWARNING)
                    self.show_text.emit('block')
                    return False
        else:
            set_server = command(cmd + '&&\ cscript ospp.vbs /sethst:' + server_input)
            if set_server == 0:
                self.show_text.emit('已将KMS服务器设置为：' + server_input)
                win32api.MessageBox(0, '设置KMS服务器失败，请重新尝试', '错误', win32con.MB_ICONWARNING)
                self.show_text.emit('block')
                return False
        activate = command(cmd + '&&\ cscript ospp.vbs /act:')
        if activate == 0:
            self.show_text.emit('已成功激活系统')
            win32api.MessageBox(0, '已成功激活Office', '成功', win32con.MB_OK)
        else:
            self.win.offce_echo.appendPlainText('激活失败，请重新尝试')
            win32api.MessageBox(0, '激活失败，请重新尝试', '错误', win32con.MB_ICONWARNING)
            self.show_text.emit('block')
            return False
        self.show_text.emit('block')
        return
