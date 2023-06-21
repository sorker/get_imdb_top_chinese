# encoding: utf-8
"""
@author: sorke
@contact: sorker0129@hotmail.com
@version: 1.0.0
@time: 2023/6/21 15:28
@file: translate.py
@desc: this is a translate file
"""
import time

import pandas as pd
# from googletrans import Translator
#
# data = pd.read_excel('data.xlsx')[250:]
#
# translator = Translator()
# df = []
# for index, name in enumerate(data['name']):
#     translations = translator.translate(name, dest='zh-cn')
#     df.append(translations.text)
#
# df1 = pd.Series(df, index=data.index, name='chinese name')
#
# data.insert(len(data.columns), 'chinese name', df1)
#
# data.to_excel('data3.xlsx')

data = pd.read_excel('data2.xlsx')

data.drop_duplicates(subset=['name'])

data.to_excel('data4.xlsx')
