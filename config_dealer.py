#——————————————————处理json文件———————————————————#


#—————————————————载入并处理配置文件————————————————#

import os
import json

#————————————————————————————————————————————————#

def get_config()->dict:
    '''获取配置文件'''
    try:
        fp = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.json'), 'r')
    except:
        print("配置文件出错")
        return {
            'empty_filler': ' ',
            'formular': True
        }
    config = json.load(fp)
    return config

def save_config(config: dict):
    '''保存配置文件'''
    fp = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.json'), 'w')
    json.dump(config, fp=fp)

#————————————————————————————————————————————————#

# save_config({'empty_filler': ' ', 'formular': True, 'toprule': True})