# -*- coding=utf-8 -*-
# author: 刘东明 liu_dongming@qq.com

import tkinter
from tkinter import *
from tkinter import ttk
import tkinter as tk
import tkinter.messagebox
import os
import pandas as pd
from pandas import Series, DataFrame
import numpy as np
from pandas.core.arrays import PandasArray
import array
from threading import Thread
from queue import Queue
from tkinter.filedialog import askdirectory
from tkinter import messagebox
import logging
import time


def cur_file_dir():
    path = sys.argv[0]

    #判断是exe还是py
    if os.path.isdir(path):
        return path
    elif os.path.isfile(path):
        return os.path.dirname(path)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logfile = time.strftime('%Y%m%d', time.localtime(time.time()))
logpath = cur_file_dir() + '/log/'
if not os.path.isdir(logpath):
    os.makedirs(logpath)

logfullpath = logpath + logfile + '.log'
fh = logging.FileHandler(logfullpath, mode='a')
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
fh.setFormatter(formatter)
logger.addHandler(fh)


class Application(tk.Tk):
    '''程序主界面'''

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.indir = StringVar()
        self.outdir = StringVar()
        self.queue = Queue()
        self.status = StringVar(self, value='准备')
        self.worker = WorkThread(self.queue, '', '')

        self.title('Excel表合并工具')
        self.resizable(width=True, height=True)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=40)
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=4)
        self.columnconfigure(2, weight=4)
        self.columnconfigure(3, weight=4)
        self.columnconfigure(4, weight=1)

        self.labelindir = ttk.Label(self, text='Excel源目录', font=('微软雅黑', 16))
        self.labelindir.grid(row=0, column=0, stick=W)
        self.entryindir = ttk.Entry(self, textvariable=self.indir)
        self.entryindir.grid(row=0, column=1, columnspan=4, stick=(W, E))
        self.buttonindir = ttk.Button(self, text='...', command=self.select_indir)
        self.buttonindir.grid(row=0, column=5)

        self.labeloutdir = ttk.Label(self, text='Excel输出目录', font=('微软雅黑', 16))
        self.labeloutdir.grid(row=1, column=0, stick=W)
        self.entryoutdir = ttk.Entry(self, textvariable=self.outdir)
        self.entryoutdir.grid(row=1, column=1, columnspan=4, stick=(W, E))
        self.buttonoutdir = ttk.Button(self, text='...', command=self.select_outdir)
        self.buttonoutdir.grid(row=1, column=5)

        self.label_status = ttk.Label(self, textvariable=self.status, font=('微软雅黑', 10))
        self.label_status.grid(row=2, column=0, columnspan=5, stick=(W, E))

        self.button_start = ttk.Button(self, text='开始合并', command=self.btn_start)
        self.button_start.grid(row=2, column=5)

        self.protocol('WM_DELETE_WINDOW', self.on_closing)

    def select_indir(self):
        path = askdirectory()
        self.indir.set(path)

    def select_outdir(self):
        path = askdirectory()
        indir = self.indir.get()
        if indir in path:
            messagebox.showwarning('错误', '输出目录请不要放在输入目录低下')
        else:
            self.outdir.set(path)

    def check_queue(self):
        if not self.queue.empty():
            self.status.set(self.queue.get())

        if self.status.get() != 'done':
            self.after(100, self.check_queue)
            self.button_start['state'] = DISABLED
        else:
            self.button_start['state'] = NORMAL

    def btn_start(self):
        if self.indir.get() == '' or self.outdir.get() == '':
            messagebox.showwarning('操作失败', '请选择Excel路径')
            return

        self.worker = WorkThread(self.queue, self.indir.get(), self.outdir.get())
        # self.worker.setDaemon(True)
        self.worker.start()
        self.after(100, self.check_queue)

    def on_closing(self):
        print('closing')
        if not self.worker.exit:
            if messagebox.askokcancel('退出', '正在合并，是否退出？'.format('Excel表合并工具')):
                self.worker.stop()
                self.quit()
        else:
            self.worker.stop()
            self.quit()


class WorkThread(Thread):
    def __init__(self, queue, indir, outdir, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.queue = queue
        self.indir = indir
        self.outdir = outdir
        self.exit = True

        if not os.path.exists(self.indir):
            print(self.indir, '不存在')

        if not os.path.exists(self.outdir):
            print(self.outdir, '不存在')

        if not self.outdir.endswith('\\'):
            self.outdir += '\\'

    def stop(self):
        self.exit = True

    def run(self):
        logger.info('开始')

        dictExcel = {}
        index = 0
        self.exit = False
        for root, dirs, files in os.walk(self.indir):
            if self.exit:
                print('exit 1')
                return

            for file in files:

                if self.exit:
                    print('exit 2')
                    return

                print(os.path.join(root, file))
                if os.path.splitext(file)[-1] != '.xlsx' and os.path.splitext(file)[-1] != '.xls' and os.path.splitext(file)[-1] != '.xlsm':
                    print('不支持的文件:', os.path.join(root, file))
                    logger.info('不支持的文件:' + os.path.join(root, file))
                    continue

                # excel转换成DataFrame
                self.queue.put('读取:' + file)
                df = pd.read_excel(os.path.join(root, file))
                new = np.array(df.columns)
                equal = False
                if dictExcel:
                    for i in dictExcel:
                        if isinstance(dictExcel[i].get_nparray() == new, bool):
                            equal = (dictExcel[i].get_nparray() == new)
                        else:
                            equal = all(dictExcel[i].get_nparray() == new)

                        if equal:
                            index = i
                            break
                else:
                    equal = False

                if equal:
                    print('same cloumns, index=', index)
                    dictExcel[index].addfile(os.path.join(root, file))
                else:
                    print('new cloumns')
                    print(new)
                    dictlen = len(dictExcel)
                    dictExcel[dictlen] = ExcelList(new)
                    dictExcel[dictlen].addfile(os.path.join(root, file))
                    index = dictlen

                logger.info('读取文件:' + os.path.join(root, file) + ' ---> ' + dictExcel[index].get_newfilename())
                dictExcel[index].frames.append(df)

        if dictExcel:
            for i in dictExcel:

                if self.exit:
                    print('exit 3')
                    return

                if False:
                    self.queue.put('合并:' + dictExcel[i].get_newfilename())
                    result = pd.concat([pd.read_excel(f) for f in dictExcel[i].get_filelist()], sort=False)
                    result.to_excel(self.outdir + str(i) + dictExcel[i].get_newfilename(), index=False,
                                    sheet_name=dictExcel[i].get_newfilename(''))
                else:
                    logger.info('合并以下文件到 --> ' + dictExcel[i].get_newfilename())
                    for filename in dictExcel[i].get_filelist():
                        logger.info(filename)
                    self.queue.put('合并:' + dictExcel[i].get_newfilename())
                    result = pd.concat(dictExcel[i].frames, sort=False)
                    result.to_excel(self.outdir + str(i) + dictExcel[i].get_newfilename(), index=False,
                                    sheet_name=dictExcel[i].get_newfilename(''))
            self.queue.put('done')

        logger.info('完成')
        self.exit = True


class ExcelList(object):
    def __init__(self, nparray):
        self.nparray = nparray
        self.filelist = []
        self.equal = False
        self.newexcelname = ''
        self.frames = []

    def __eq__(self, other):
        self.equal = False
        if len(self.nparray) > 0:
            if isinstance(self.nparray == other, bool):
                self.equal = (old == new)
            else:
                self.equal = all(self.nparray == other)
        else:
            self.equal = False

        if self.equal:
            print('same cloumns')
        else:
            print('new cloumns')
        return self.equal

    def __str__(self):
        return str(self.nparray)

    def addfile(self, filepath):
        self.filelist.append(filepath)
        filepath, tempfilename = os.path.split(filepath)
        newexcelname, extension = os.path.splitext(tempfilename)

        if self.newexcelname == '':
            self.newexcelname = newexcelname
            print('get:', self.newexcelname)
        else:
            print(self.newexcelname, ',', newexcelname)
            newlen, self.newexcelname = self.get_max_common_substr(self.newexcelname, newexcelname)
            print('new:', self.newexcelname, ',', newlen)

    def get_nparray(self):
        return self.nparray

    def get_filelist(self):
        return self.filelist

    def get_newfilename(self, ext='.xlsx'):
        return self.newexcelname + ext

    def get_max_common_substr(self, s1, s2):
        # 求两个字符串的最长公共子串
        # 思想：建立一个二维数组，保存连续位相同与否的状态

        len_s1 = len(s1)
        len_s2 = len(s2)

        # 生成0矩阵，为方便后续计算，多加了1行1列
        # 行: (len_s1+1)
        # 列: (len_s2+1)
        record = [[0 for i in range(len_s2 + 1)] for j in range(len_s1 + 1)]

        maxNum = 0  # 最长匹配长度
        p = 0  # 字符串匹配的终止下标

        for i in range(len_s1):
            for j in range(len_s2):
                if s1[i] == s2[j]:
                    # 相同则累加
                    record[i + 1][j + 1] = record[i][j] + 1

                    if record[i + 1][j + 1] > maxNum:
                        maxNum = record[i + 1][j + 1]
                        p = i  # 匹配到下标i

        # 返回 子串长度，子串
        return maxNum, s1[p + 1 - maxNum: p + 1]


if __name__ == '__test__':
    indir = ''
    oudir = ''
    while True:
        indir = input('请输入Excel表所在目录>')
        if not os.path.exists(indir):
            print(indir, '不存在')
            continue
        else:
            break

    while True:
        outdir = input('请输入输出目录>')
        if not os.path.exists(outdir):
            print(outdir, '不存在')
            continue
        else:
            if not outdir.endswith('\\'):
                outdir += '\\'
            break

    # 新建列表，存放文件名（可以忽略，但是为了做的过程能心里有数，先放上）
    filename_excel = []

    # 新建列表，存放每个文件数据框（每一个excel读取后存放在数据框）
    frames = []

    oldArray = np.array([])
    dictExcel = {}
    index = 0

    # pd.set_option('display.max_columns', 48)
    # pd.set_option('display.max_rows', 100)

    for root, dirs, files in os.walk(indir):
        for file in files:
            print(os.path.join(root, file))
            if os.path.splitext(file)[-1] != '.xlsx' and os.path.splitext(file)[-1] != '.xls':
                print('不支持的文件:', os.path.join(root, file))
                logger.warning('不支持的文件:' + os.path.join(root, file))
                continue
            filename_excel.append(os.path.join(root, file))
            # excel转换成DataFrame
            df = pd.read_excel(os.path.join(root, file))
            new = np.array(df.columns)
            equal = False
            if dictExcel:
                for i in dictExcel:
                    if isinstance(dictExcel[i].get_nparray() == new, bool):
                        equal = (dictExcel[i].get_nparray() == new)
                    else:
                        equal = all(dictExcel[i].get_nparray() == new)

                    if equal:
                        index = i
                        break
            else:
                equal = False

            if equal:
                print('same cloumns, index=', index)
                dictExcel[index].addfile(os.path.join(root, file))
            else:
                print('new cloumns')
                print(new)
                dictlen = len(dictExcel)
                dictExcel[dictlen] = ExcelList(new)
                dictExcel[dictlen].addfile(os.path.join(root, file))

    if dictExcel:
        for i in dictExcel:
            result = pd.concat([pd.read_excel(f) for f in dictExcel[i].get_filelist()], sort=False)
            result.to_excel(outdir + str(i) + dictExcel[i].get_newfilename(), index=False,
                            sheet_name=dictExcel[i].get_newfilename(''))

if __name__ == '__main__':
    app = Application()
    app.mainloop()
