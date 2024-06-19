from filedownloader import downloadFiles;
from excelparser import getData as excelData;
from pdfparser import getData as pdfData;
from hwpparser import getData as hwpData;

# juho : 파일다운로드, 엑셀, pdf, hwp 파일 파싱을 한번에 진행시키는 함수
def getData():
    downloadFiles();
    data = {}
    data.update(excelData())
    data.update(pdfData())
    data.update(hwpData())
    # print(data)
    return data

# getData()