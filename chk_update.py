from ftplib import FTP
import py7zr
from py7zr import exceptions as py7z_exceptions
import subprocess
import requests
import threading
import time
from pathlib import Path
import re
import traceback
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
import pickle
import shutil
import os
import time
import socket
import urllib3
from config import *

class HeaderError(Exception):
    pass

class chk_update(threading.Thread):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.parent = parent
        self.__flag = threading.Event()
        self.start()
    def pause(self):
        self.__flag.clear()
    def resume(self):
        self.__flag.set()
    def run(self):
        self.parent.take_text.insert(tk.END, '檢查更新中...\n')
        # #查詢本機版本號
        self.chk_robot_version()
        self.parent.take_text.insert(tk.END, f'本機PostMan版本號:{self.local_ver}\n')
        # #查詢最新版本號
        self.chk_ftp_robot_version()
        self.parent.take_text.insert(tk.END, f'最新PostMan版本號:{self.new_ver}\n')
        if not self.local_ver or (self.new_ver and self.local_ver and self.local_ver != self.new_ver):
            #取出原設定檔
            self.take_config()
            #下載新版機器人
            download_result = self.download_robot()
            if not download_result:
                self.parent.take_text.insert(tk.END, f'下載失敗, 將啟動目前版本:{self.local_ver}機器人\n')
                self.start_local_robot()
                #線程結束  
                self.parent.quit()
                self.pause
                return
            #解壓縮機器人
            zip = self.unzip()
            if not zip:
                self.parent.take_text.insert(tk.END, f'壓縮檔異常, 將啟動目前版本:{self.local_ver}機器人\n')
                self.start_local_robot()
                #線程結束  
                self.parent.quit()
                self.pause
                return
            self.start_robot_path = self.dir_path + '\\' + str(self.path).split('.7z')[0].rsplit('\\',1)[1] + f'\\PR6_PostMan.exe'
            self.start_path = self.dir_path + '\\' + str(self.path).split('.7z')[0].rsplit('\\',1)[1]
            #複蓋設定檔
            self.update_config()
            #啟動新版機器人
            self.start_robot(self.start_robot_path, self.start_path)
            self.parent.take_text.insert(tk.END, f'新版本{self.new_ver}機器人啟動完成\n')
            #刪除新版機器人壓縮檔
            self.del_7z()
            #線程結束        
            self.pause
            self.parent.quit()
            return
        else:
            self.start_local_robot()
            #線程結束        
            self.pause
            self.parent.quit()

    #啟動目前機器人
    def start_local_robot(self):
        self.start_robot_path = str(Path('.').absolute() / f'PR6_PostMan.exe')
        self.start_path = str(Path('.').absolute())
        chk_update.start_robot(self.start_robot_path, self.start_path)

    #檢查本地機器人版本
    @classmethod
    def chk_robot_version(cls):
        cls.local_ver = ''
        path = Path('.env')
        if not path.exists():
            return
        with path.open('r') as r:
            cls.version = r.read()
        cls.local_ver = re.search('VERSION = (?P<version>.+)', cls.version).groupdict().get('version', '')
        cls.app_name = re.search('APP_NAME = (?P<app_name>.+)', cls.version).groupdict().get('app_name', '')
        return

    #確認遠端主機最新機器人版本
    def chk_ftp_robot_version(self):
        ver_list = []
        self.new_ver = ''
        while True:
            try:
                ftp = FTP()
                ftp.connect('18.163.192.24', timeout=5)
                ftp.login('taiwinner','taiwinner999')
                ftp.encoding = 'utf-8'
                ftp.cwd('pr6/public_html/botdownload/小工具/history')
                data_list = ftp.nlst()
                for data in data_list:
                    ver_list.append(int(re.search('PR6_PostMan_V(?P<ver>\d+.\d+.\d+).+', data).groupdict().get('ver', '').replace('.','')))
                for i, d in enumerate((ver_list)):
                    if d == max(ver_list):
                        self.new_ver = re.search('PR6_PostMan_(?P<ver>\w+.\d+.\d+).+', data_list[i]).groupdict().get('ver', '')
                return
            except (requests.exceptions.Timeout, 
                        requests.exceptions.ConnectionError, 
                        socket.timeout, 
                        urllib3.exceptions.ReadTimeoutError):
                    self.parent.take_text.insert(tk.END, '確認最新機器人版本時發生Timeout\n將進行重試..\n')
                    continue
            except Exception:
                print('\n' + traceback.format_exc())
                return
            
    #下載最新版本機器人
    def download_robot(self):
        chunk_size = 4096
        if not hasattr(self, 'app_name'):
            self.parent.take_text.insert(tk.END, f'無法確認機器人類型, 【ENV】NOT APP_NAME\n')
            return False
        self.path = Path('.').absolute().parent / f'{self.app_name}.7z'
        if not self.path.exists():
            self.path.touch()
        url = f'https://testcdn.test998.com/botdownload/小工具/{self.app_name}.7z'
        #url = 'http://118.163.18.126:62341/botdownload/小工具/PR6_PostMan.7z'
        self.parent.take_text.insert(tk.END, f'新版本{self.new_ver}機器人下載中...\n')
        while True:
            try:
                with self.path.open('ab+') as f:
                    f.seek(0)
                    file_size = len(self.path.read_bytes())
                    i = file_size // chunk_size
                    headers = {'Range': f'bytes={len(f.read())}-'}
                    resp = requests.get(url, stream=True, verify=False, timeout=5, headers=headers)
                    if resp.status_code == 416:
                        resp.close()
                        return True
                    if resp.status_code not in [200, 206]:
                        raise HeaderError((
                            f"機器人壓縮檔下載失敗({url})\n"
                            f"取得的狀態碼為：{resp.status_code}"
                        ))
                    if resp.headers['Content-Type'] not in ['application/x-msdownload', 'application/x-7z-compressed']:
                        raise HeaderError((
                            f"機器人壓縮檔下載失敗({url})\n"
                            f"取得的檔案類型為：{resp.headers['Content-Type']}"
                        ))
                    # with self.path.open('wb') as w:
                    #     w.write(resp.content)
                    #total = len(resp.content) // 4096
                    total = (file_size + float(resp.headers['Content-Length'])) // chunk_size
                    for chunk in resp.iter_content(chunk_size=chunk_size):
                        f.write(chunk)
                        i += 1
                        self.parent.progressbar['value'] = percentage = 100 * i / total
                        #print(f'percentage:{percentage}, total:{total}, i:{i}')
                        if percentage < 25:
                            self.parent.style.configure('text.Horizontal.TProgressbar', text=f'{int(percentage)} %', background='brown')
                        elif percentage < 50:
                            self.parent.style.configure('text.Horizontal.TProgressbar', text=f'{int(percentage)} %', background='orange')
                        elif percentage < 75:
                            self.parent.style.configure('text.Horizontal.TProgressbar', text=f'{int(percentage)} %', background='gold')
                        else:
                            self.parent.style.configure('text.Horizontal.TProgressbar', text=f'{int(percentage)} %', background='mediumseagreen')
                        self.parent.root.update_idletasks()

            except HeaderError as e:
                self.parent.take_text.insert(tk.END, f'下載機器人發生異常:{e}\n將進行重試..\n')
                continue
            except (requests.exceptions.Timeout, 
                    requests.exceptions.ConnectionError, 
                    socket.timeout, 
                    urllib3.exceptions.ReadTimeoutError):
                self.parent.take_text.insert(tk.END, '下載機器人時發生Timeout\n將進行重試..\n')
                continue
            except Exception as e:
                print('\n' + traceback.format_exc())
                self.parent.take_text.insert(tk.END, f'下載機器人發生未知異常:{e}\n')
                return False


    #解壓縮
    def unzip(self):
        try:
            self.dir_path = str(Path('.').absolute().parent)
            self.parent.take_text.insert(tk.END, f'新版本{self.new_ver}機器人解壓縮中...\n')
            with py7zr.SevenZipFile(str(self.path), mode='r') as sevenZ_f:
                # sevenZ_f_file = sevenZ_f.getnames()
                # selective_files = [f for f in sevenZ_f_file if '啟動機器人.exe' not in f]
                # sevenZ_f.extract(self.dir_path, targets=selective_files)
                sevenZ_f.extractall(self.dir_path)
            return True
        except PermissionError:
            print('\n' + traceback.format_exc())
            return True
        except py7z_exceptions.Bad7zFile:
            return False
        except Exception:
            print('\n' + traceback.format_exc())
            return False

    #啓動機器人
    @classmethod
    def start_robot(cls, start_robot_path, start_path):
        print(f'啟動機器人位置:{start_robot_path}, 資料夾位置{start_path}')
        try:
            subprocess.Popen(
                start_robot_path,
                shell=True, 
                cwd=start_path,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE)
            return
        except (FileNotFoundError, NotADirectoryError) as e:
            message = (
                '机器人启动失败(档案不存在)。\n'
                '1) 请检查是否将资料夹设为防毒白名单。\n'
                '2) 重启本更新机器人。'
            )
            messagebox.showerror(title='错误', message=message)
            return

    #刪除壓縮資料夾
    def del_7z(self):
        path = Path('.').absolute().parent / f'PR6_PostMan.7z'
        try:
            path.unlink()
            return
        except FileNotFoundError as e:
            return
        except Exception:
            print('\n' + traceback.format_exc())
            return

    #刪除舊版機器人
    @classmethod
    def del_old_robot(cls, local_ver):
        path = Path('.').absolute().parent / f'PR6_PostMan_{local_ver}'
        try:
            #shutil.rmtree(path, ignore_errors=True) ignore=忽略
            #shutil.rmtree(path)
            os.system(f"rd/s/q {str(path)}") #強制刪除
            #os.system(f"rm -f {str(path)}")
            return
        except FileNotFoundError as e:
            print(f'查無此檔案:{path}')
            return
        except Exception:
            print('\n' + traceback.format_exc())
            return

    #更新版本號
    def update_env(self):
        path = Path('.env')
        with path.open('r') as r:
            str_env = r.read()
        str_env = str_env.replace(self.local_ver, self.new_ver)
        with path.open('w') as w:
            w.write(str_env)
        return
    
    #取出原設定檔
    def take_config(self):
        #讀取local端config.dll
        self.str_config = ''
        path = Path('config.dll')
        if not path.exists():
            self.parent.take_text.insert(tk.END, f'查無當前版本設定檔\n')
            return
        with path.open('rb') as r:
            self.str_config = pickle.load(r)
        self.parent.take_text.insert(tk.END, f'儲存當前版本設定檔\n')
        return

    #複蓋設定檔
    def update_config(self):
        #寫入最新版機器人的config.dll
        try:
            if not self.str_config:
                return 
            new_path = Path(self.start_path + '\config.dll')
            with new_path.open('wb') as w:
                pickle.dump(self.str_config, w)
            self.parent.take_text.insert(tk.END, f'新版本{self.new_ver}設定檔更新完成\n')
            return
        except Exception:
            print('\n' + traceback.format_exc())
            return

    #確認新機器人版本號
    def check_ver(self):
        path = Path(self.start_path + '\.env')
        with path.open('r') as r:
            data = r.read()
        result = re.search('VERSION = (?P<version>.+)', data).groupdict().get('version', '')
        return result
        

class MainPage(tk.Frame):
    def __init__(self, root):
        super().__init__()
        self.root = root
        self.title_font = ('Tahoma', 10, 'bold')
        self.root.title(f'檢查更新程式 V{robot_version}')
         #設定畫面
        self.setup_base_setting()
        self.msg_windows_frame.pack(expand=1, fill='both', padx=2, pady=1)
        self.take_text.pack(expand=1, side='top',fill='both')
        self.middle.pack(fill='both', padx=2, pady=1)
        self.progressbar.pack(fill='x')
        chk_update(self)

    def setup_base_setting(self):
        # log訊息框
        self.msg_windows_frame = tk.Frame(self.root)
        self.take_text = tk.Text(self.msg_windows_frame, width=40, height=15)
        # 進度條
        self.middle = tk.Frame(self.root)
        self.style = ttk.Style(self.middle)
        self.style.theme_use('default') #不確定功能, 但使用後configure, 可以使用參數background
        self.style.layout('text.Horizontal.TProgressbar',
                        [
                            ('Horizontal.Progressbar.trough',
                                {'children': [
                                                (
                                                    'Horizontal.Progressbar.pbar',
                                                    {'side': 'left', 'sticky': 'ns'}
                                                )
                                            ],
                                'sticky': 'nswe'}),
                                ('Horizontal.Progressbar.label', {'sticky': ''}
                            )
                        ]
                    )
        self.style.configure('text.Horizontal.TProgressbar', text='0 %', foreground='black', background='red')
        self.progressbar = ttk.Progressbar(self.middle, orient='horizontal', style='text.Horizontal.TProgressbar',)

    def quit(self):
        self.take_text.insert(tk.END, '檢查更新程式關閉中..')
        self.root.quit()

def main():
    root = tk.Tk()
    MainPage(root)
    root.mainloop()
    
if __name__ == '__main__':
    main()