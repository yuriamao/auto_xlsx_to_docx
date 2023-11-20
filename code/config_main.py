# -*- coding: utf-8 -*-
"""
Created on Tue Aug 30 08:21:10 2022

@author: lxl
"""
import sys
import os

from config_build import *
import pandas as pd

# 读取配置文件
config_df = pd.read_excel(r'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/pmi_config.xlsx',header=0)