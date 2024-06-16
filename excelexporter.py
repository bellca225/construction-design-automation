import openpyxl
file_path = "./excel_result";
wb = openpyxl.Workbook()

#시트
wb.active.title = "간이공사산출명세서";
#a2 - w2 merge
ws = wb.active
ws.merge_cells('A2:W2')
# insert data to a2
# bold, fontsize 24, text align center
ws['A2'].font = openpyxl.styles.Font(size=24, bold=True)
ws['A2'].alignment = openpyxl.styles.Alignment(horizontal='center')
ws['A2'] = '간이공사비 산출명세서'
# insert data to a3 
# bold, fontsize 16
ws['A3'].font = openpyxl.styles.Font(size=16, bold=True)
ws['A3'] = 'o 공사명 : 보호계전기 데이터 취득시스템(PDAS) 포인트 증설공사'
defaultData = [
    {'data' : '구      분', 's' : 'a4', 'e' : 'c4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '재료비', 's' : 'd4', 'e' : 'e4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '노무비', 's' : 'f4', 'e' : 'g4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '경 비', 's' : 'h4', 'e' : 'i4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '일반관리비', 's' : 'j4', 'e' : 'k4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '이 윤', 's' : 'l4', 'e' : 'm4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '소  계', 's' : 'n4', 'e' : 'o4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '부가세', 's' : 'p4', 'e' : 'r4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '합계', 's' : 's4', 'e' : 't4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '작성자', 's' : 'u4', 'e' : 'v4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '검토자', 's' : 'w4', 'e' : 'w4', 'fontsize' : 11, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '설\n계\n공\n사\n비', 's' : 'a5', 'e' : 'a7', 'fontsize' : 9, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '회사분', 's' : 'b5', 'e' : 'c5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'd5', 'e' : 'e5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'f5', 'e' : 'g5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'h5', 'e' : 'i5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'j5', 'e' : 'k5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'l5', 'e' : 'm5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'n5', 'e' : 'o5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'p5', 'e' : 'r5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 's5', 'e' : 't5', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    # 넣어야됨
    {'data' : '도급분', 's' : 'b6', 'e' : 'c6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'd6', 'e' : 'e6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'f6', 'e' : 'g6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'h6', 'e' : 'i6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'j6', 'e' : 'k6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'l6', 'e' : 'm6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'n6', 'e' : 'o6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'p6', 'e' : 'r6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 's6', 'e' : 't6', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    #넣어야됨
    {'data' : '계', 's' : 'b7', 'e' : 'c7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'd7', 'e' : 'e7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'f7', 'e' : 'g7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'h7', 'e' : 'i7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'j7', 'e' : 'k7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'l7', 'e' : 'm7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'n7', 'e' : 'o7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 'p7', 'e' : 'r7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '-', 's' : 's7', 'e' : 't7', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '팀 원 \n이현영\n\n(인)', 's' : 'u5', 'e' : 'v7', 'fontsize' : 10, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '차 장 \n이현영\n\n(인)', 's' : 'w5', 'e' : 'w7', 'fontsize' : 10, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '예정가격(결정권자)', 's' : 'a8', 'e' : 'd9', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '                                         (원)  - 부가세 포함', 's' : 'e8', 'e' : 'o9', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '부      장      이 영 성             (인)', 's' : 'p8', 'e' : 'w9', 'fontsize' : 11, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '공   사   비   예   산   서', 's' : 'a10', 'e' : 'o10', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '공 사 비 정 산 서', 's' : 'p10', 'e' : 'w10', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},    
    {'data' : 'No.', 's' : 'a11', 'e' : 'b11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '품 명', 's' : 'c11', 'e' : 'd11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '규격', 's' : 'e11', 'e' : 'g11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '단 위', 's' : 'h11', 'e' : 'h11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '수 량', 's' : 'i11', 'e' : 'j11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '단 가	', 's' : 'k11', 'e' : 'l11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '금   액', 's' : 'm11', 'e' : 'n11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '비  고', 's' : 'o11', 'e' : 'o11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '수량', 's' : 'p11', 'e' : 'q11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '단 가', 's' : 'r11', 'e' : 's11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '금  액	', 's' : 't11', 'e' : 'v11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},
    {'data' : '증감', 's' : 'w11', 'e' : 'w11', 'fontsize' : 12, 'fontstyle' : 'bold', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFC0'},

    # 재료비 종류별로 여러개
    {'data' : '1. 재료비', 's' : 'a12', 'e' : 'd12', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
	{'data' : '도급분', 's' : 'c13', 'e' : 'd13', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
	# 1. 재료비 금액
    {'data' : '579,080', 's' : 'm12', 'e' : 'n12', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},
    # 1. 재료비 도급분 금액
    {'data' : '579,080', 's' : 'm13', 'e' : 'n13', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},
    # 여러개종류
    {'data' : ' '+'UTP케이블', 's' : 'c14', 'e' : 'd14', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    {'data' : 'UTP CAT.5E 4PR', 's' : 'e14', 'e' : 'g14', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : 'M', 's' : 'h14', 'e' : 'h14', 'fontsize' : 15, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '324', 's' : 'i14', 'e' : 'j14', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '620', 's' : 'k14', 'e' : 'l14', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '200,880 	', 's' : 'm14', 'e' : 'n14', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '견적', 's' : 'o14', 'e' : 'o14', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    # 넣어야됨 노무비 종류별로 여러개
    {'data' : '2. 노무비', 's' : 'a20', 'e' : 'd20', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    # 금액
    {'data' : '2,933,640', 's' : 'm20', 'e' : 'n20', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},
    # 직접노무비
    {'data' : '가. 직접노무비', 's' : 'c21', 'e' : 'd21', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    # 금액
    {'data' : '2,596,142', 's' : 'm21', 'e' : 'n21', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},
    #간접노무비
    {'data' : '나. 간접노무비', 's' : 'c30', 'e' : 'd30', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    # 금액
    {'data' : '337,498', 's' : 'm30', 'e' : 'n30', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},

    # 직접노무비
    {'data' : ' '+' UTP케이블 포설', 's' : 'c22', 'e' : 'd22', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    {'data' : '4P(1열)', 's' : 'e22', 'e' : 'g22', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : 'M', 's' : 'h22', 'e' : 'h22', 'fontsize' : 15, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '6', 's' : 'i22', 'e' : 'j22', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '5,473', 's' : 'k22', 'e' : 'l22', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '32,838', 's' : 'm22', 'e' : 'n22', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '품1', 's' : 'o22', 'e' : 'o22', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    # 넣어야됨 간접노무비
    {'data' : '직접노무비 * 13.0%', 's' : 'c31', 'e' : 'd31', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'center', 'background' : 'FFFFFF'},
    {'data' : '0.13', 's' : 'i31', 'e' : 'j31', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '2,596,142', 's' : 'k31', 'e' : 'l31', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    {'data' : '337,498', 's' : 'm31', 'e' : 'n31', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},
    # 넣어야됨 경비
	{'data' : '3. 경비', 's' : 'a33', 'e' : 'd33', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
	# 금액
    {'data' : '379,492', 's' : 'm33', 'e' : 'n33', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},
    # 도급분
    {'data' : '가. 도급분', 's' : 'c34', 'e' : 'd34', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    # 금액
    {'data' : '379,492', 's' : 'm34', 'e' : 'n34', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},
    # 넣어야됨 일반관리비
    {'data' : '4. 일반관리비', 's' : 'a44', 'e' : 'd44', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    {'data' : '(도급재료비 + 노무비 + 도급경비) * 6%', 's' : 'e44', 'e' : 'l44', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    # 금액
    {'data' : '233,533', 's' : 'm44', 'e' : 'n44', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},

    # 넣어야됨 이윤
    {'data' : '5. 이윤', 's' : 'a46', 'e' : 'd46', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    {'data' : '(노무비 + 도급경비 + 일반관리비) * 15%', 's' : 'e46', 'e' : 'l46', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'left', 'background' : 'FFFFFF'},
    # 금액
    {'data' : '532,000', 's' : 'm46', 'e' : 'n46', 'fontsize' : 12, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFF00'},
    {'data' : '일반공사', 's' : 'o46', 'e' : 'o46', 'fontsize' : 10, 'fontstyle' : '', 'borderline' : True, 'textalign' : 'right', 'background' : 'FFFFFF'},

]
# 개귀찮네.
lenData = {
    0: 5,
    1: 7, 
}

for d in defaultData:
    ws.merge_cells(d['s'] + ':' + d['e'])
    ws[d['s']].font = openpyxl.styles.Font(size=d['fontsize'], bold=d['fontstyle'] == 'bold')
    ws[d['s']].alignment = openpyxl.styles.Alignment(horizontal=d['textalign'])
    ws[d['s']].fill = openpyxl.styles.PatternFill(start_color=d['background'], end_color=d['background'], fill_type="solid")
    #border s to e all cells
    if d['borderline']:
        for i in range(ord(d['s'][0]), ord(d['e'][0]) + 1):
            for j in range(int(d['s'][1:]), int(d['e'][1:]) + 1):
                ws[chr(i) + str(j)].border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'), top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin'))

    # ws[d['s']].border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'), top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin'))
    ws[d['s']] = d['data']

# #a4 - c4 merge
# # insert data to a4
# # bold, fontsize 11, textalign center, backgroundcolor #ffffc0, border line all
# ws.merge_cells('A4:C4')
# ws['A4'].font = openpyxl.styles.Font(size=11, bold=True)
# ws['A4'].alignment = openpyxl.styles.Alignment(horizontal='center')
#alignment right
# ws['A4'].fill = openpyxl.styles.PatternFill(start_color="FFFFC0", end_color="FFFFC0", fill_type="solid")
# ws['A4'].border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'), top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin'))
# ws['A4'] = '구      분'

new_filename = 'new_costs.xlsx'

wb.save(file_path + "/" + new_filename)