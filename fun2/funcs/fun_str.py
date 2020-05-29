#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/11 10:16:17
# @Author  : HouWk
# @Site    : 
# @File    : fun_str.py
# @Software: PyCharm
import re
def replace_0(str,arg='/.-'):
    pattern = '([%s])0' % arg
    new = re.sub(pattern,lambda x:x.group(1),str)
    pattern1 = '^0'
    new = re.sub(pattern1,'',new)
    return new

if __name__ == '__main__':
    print(replace_0('04-05.06'))