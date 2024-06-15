import pandas as pd
import json

def getData():
    # 파일명
    file_name = './datas/costs.xlsx'

    # Daraframe형식으로 엑셀 파일 읽기
    data = pd.read_excel(file_name, sheet_name='건축제비율', header=None, engine='openpyxl')

    # 데이터 프레임 출력
    dict = data.to_dict()
    final = {
        '간접노무비' : 0,
        '기타경비' : 0,
        '일반관리비' : 0,
        '이윤' : 0,
        '고용보험료' : 0,
        '건강보험료' : 0,
        '노인장기요양보험료' : 0,
        '연금보험료' : 0,
    }
    for d in dict.values():
        vList = list(d.values())

        for i in range(len(vList)):
            dd = vList[i]
            if type(dd) == str :
                if dd.find('[간접노무비]') != -1:
                    final['간접노무비'] = float(d[14])
                if dd.find('[기타경비]') != -1:
                    final['기타경비'] = float(d[14])
                elif dd.find('전문·\n전기·통신·\n소방·기타') != -1:
                    final['일반관리비'] = float(d[14])
                elif dd.find('[이윤]') != -1:
                    final['이윤'] = float(d[14])
                elif dd.find('산업\n설비\n(건축)') != -1 and vList.count('요율') > 0:
                    index = vList.index('요율')
                    final['고용보험료'] = float(vList[index+2])
                elif dd.find('(직노) x') != -1 :
                    if vList.count('[건강보험료]') > 0:
                        final['건강보험료'] = float(d[96].split('x')[1])
                    elif vList.count('[연금보험료]') > 0:
                        final['연금보험료'] = float(d[96].split('x')[1])    
                elif dd.find('(건강보험료) x') != -1 and vList.count('[노인장기요양보험료]') > 0:
                    final['노인장기요양보험료'] = float(d[96].split('x')[1])
    return final;