#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019-06-06 13:09
# @Author  : LiYahui
# @Description :  set_index demo
import pandas as pd

data = {'A': ['A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11'],
        'B': ['B0', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11'],
        'C': ['C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11'],
        'D': ['D0', 'D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10', 'D11']}
df = pd.DataFrame(data)
print(df)
'''
      A    B    C    D
0    A0   B0   C0   D0
1    A1   B1   C1   D1
2    A2   B2   C2   D2
3    A3   B3   C3   D3
4    A4   B4   C4   D4
5    A5   B5   C5   D5
6    A6   B6   C6   D6
7    A7   B7   C7   D7
8    A8   B8   C8   D8
9    A9   B9   C9   D9
10  A10  B10  C10  D10
11  A11  B11  C11  D11
'''
# drop=True
df1 = df.set_index("A", drop=True, append=False, inplace=False, verify_integrity=False)
print(df1)
'''
       B    C    D
A                 
A0    B0   C0   D0
A1    B1   C1   D1
A2    B2   C2   D2
A3    B3   C3   D3
A4    B4   C4   D4
A5    B5   C5   D5
A6    B6   C6   D6
A7    B7   C7   D7
A8    B8   C8   D8
A9    B9   C9   D9
A10  B10  C10  D10
A11  B11  C11  D11
'''
# drop=False
df2 = df.set_index("A", drop=False, append=False, inplace=False, verify_integrity=False)
print(df2)
'''
       A    B    C    D
A                      
A0    A0   B0   C0   D0
A1    A1   B1   C1   D1
A2    A2   B2   C2   D2
A3    A3   B3   C3   D3
A4    A4   B4   C4   D4
A5    A5   B5   C5   D5
A6    A6   B6   C6   D6
A7    A7   B7   C7   D7
A8    A8   B8   C8   D8
A9    A9   B9   C9   D9
A10  A10  B10  C10  D10
A11  A11  B11  C11  D11
'''
# append=True
df3 = df.set_index("A", drop=False, append=True, inplace=False, verify_integrity=False)
print(df3)
'''
          A    B    C    D
   A                      
0  A0    A0   B0   C0   D0
1  A1    A1   B1   C1   D1
2  A2    A2   B2   C2   D2
3  A3    A3   B3   C3   D3
4  A4    A4   B4   C4   D4
5  A5    A5   B5   C5   D5
6  A6    A6   B6   C6   D6
7  A7    A7   B7   C7   D7
8  A8    A8   B8   C8   D8
9  A9    A9   B9   C9   D9
10 A10  A10  B10  C10  D10
11 A11  A11  B11  C11  D11
'''

# inplance=True
df4 = df.set_index("A", drop=False, append=True, inplace=True, verify_integrity=False)
print(df4)
# 不知道为什么
'''
None
'''