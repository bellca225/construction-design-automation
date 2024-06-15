from filedownloader import downloadFiles;
from excelparser import getData as excelData;
from pdfparser import getData as pdfData;
from hwpparser import getData as hwpData;

def getData():
    downloadFiles();
    data = {}
    data.update(excelData())
    data.update(pdfData())
    data.update(hwpData())
    print(data)
    return data

getData()