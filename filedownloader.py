import requests
import ssl
from loguru import logger

# juho : SSL 오류를 해결하기 위한 클래스
class TLSAdapter(requests.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.set_ciphers("AES128-SHA256")
        kwargs["ssl_context"] = ctx
        return super(TLSAdapter, self).init_poolmanager(*args, **kwargs)

# juho : 파일 다운로드할 origin 과 url 리스트
origin = 'https://ictis.kica.or.kr/file/download'
urls = [
    {'url' : 'c621f4a1-eb93-4b07-bdba-3b344d9d0c0f', 'name' : 'costs.xlsx'},
    {'url' : 'b7d6a621-439f-471d-afd1-a16f663a4697', 'name' : 'standard_production.pdf'},
    {'url' : 'e1a2241e-c424-4808-8e27-fad033b5c1d1', 'name' : 'wages.hwp'}
]
# juho : 파일 다운로드 함수
def downloadFiles():
    for d in urls:
        logger.info(f'Downloading {d["name"]} ..')
        # juho : User-Agent를 추가하지않으면 400 에러 발생
        headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'}

        with requests.session() as s:
            s.mount("https://", TLSAdapter())
            response = s.get(f"{origin}/{d['url']}", headers=headers)
        # response = requests.get(, verify=certifi.where())
        with open(f"./datas/{d['name']}", 'wb') as file:
            file.write(response.content)
        logger.info(f'{d["name"]} downloaded!')