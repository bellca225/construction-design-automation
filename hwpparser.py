import olefile
import os
import pandas as pd
from loguru import logger

def read_hwp(path): # path = "파일위치/파일.hwp"
    f = olefile.OleFileIO(path)
    # 신청서가 짧은 분량이라 미리보기 뷰로도 충분히 내용을 가져올 수 있음
    encoded_txt = f.openstream("PrvText").read()
    text = encoded_txt.decode("utf-16",errors="ignore")
    return text

def text2table(text):
    # 표의 내용만 가져오기 위해 다음과 같이 "><" 기호로 셀을 분리하고,
    # "<" 기호가 있는 경우에만 text데이터를 리스트로 담음
    text = text.replace(" ","").replace("><","\t").replace(">","") 
    text = text.split("\r\n")
    table = [t.replace("<","").split("\t") for t in text if len(t) > 0 and t[0] == "<"]
    return table

def getData():
    logger.info('Parsing hwp file..')
    data = text2table(read_hwp("./datas/wages.hwp"))

    published_wage = {}
    isFind = False
    for d in data:
        if isFind :
            print(d)
            published_wage[d[1]] = int(str.replace(d[3], ',',''))
        elif "공표임금" in d:
            isFind = True

    logger.info('Parsing hwp file done!')
    return published_wage