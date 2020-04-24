#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/16 14:04:24
# @Author  : HouWk
# @Site    : 
# @File    : fun_chart.py
# @Software: PyCharm
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d


def ax_plt(figsize=(18, 7)):
    f, ax = plt.subplots(figsize=figsize)
    return ax

def smooth_chart(x, y,figsize=(18, 7)):
    f,ax = plt.subplots(figsize=figsize)
    f = interp1d(x, y)
    # f2 = interp1d(x, y, kind='zero')
    # f3 = interp1d(x, y, kind='nearest')
    # f4 = interp1d(x, y, kind='slinear')
    # f5 = interp1d(x, y, kind='linear')
    f6 = interp1d(x, y, kind='quadratic')
    # f7 = interp1d(x, y, kind='cubic')

    xnew = np.linspace(x.min(), x.max(), num=500, endpoint=True)
    plt.plot(x, y, 'o')
    # plt.plot( xnew, f2(xnew), '--')
    # plt.plot( xnew, f3(xnew), '--')
    # plt.plot( xnew, f4(xnew), '--')
    # plt.plot( xnew, f5(xnew), '--')
    return plt.plot(xnew, f6(xnew), '--')
    # plt.plot( xnew, f7(xnew), '--')
    # plt.legend(['原始','zero','nearest','slinear','linear','quadratic','cubic'], loc='best')



def chart_parameters(ax,figsize=(18, 7)):
    ax.rcParams['font.sans-serif'] = ['SimHei']  # 用黑体显示中文
    ax.rcParams['axes.unicode_minus'] = False  # 正常显示负号
    ax.figure(figsize=(18, 7))
    ax.show()


def component_chart():
    plt.plot(x,y)

if __name__ == '__main__':
    x= np.arange(1,31)
    y = np.random.randint(1,30,30)
    chart_parameters(smooth_chart(x,y))
