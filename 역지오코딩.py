import openpyxl
import requests
import pandas as pd
import time

# excel 파일 열기
excel = pd.read_excel("/content/drive/MyDrive/data/불법좌회전.xlsx") # 불법유턴, 불법좌회전, 신호위반, 중앙선침범, 진로변경방법위반

# 위도와 경도 읽어오기 및 결과 저장 리스트 생성
latitude = []
longitude = []
addresses = []

for row in excel.iter_rows(min_row=2):
    if row[0].value and row[1].value and row[2].value:
        latitude.append(row[1].value)
        longitude.append(row[2].value)

# 주소 요청 함수 정의 및 각각의 위도와 경도에 대해 주소 요청하기 
def get_address(latitude, longitude):
    url = f"https://naveropenapi.apigw.ntruss.com/map-reversegeocode/v2/gc?request=coordsToaddr&coords={longitude},{latitude}&sourcecrs=epsg:4326&output=json&orders=addr,admcode,roadaddr"
    response = requests.get(url, headers={
        "X-NCP-APIGW-API-KEY-ID": "",
        "X-NCP-APIGW-API-KEY": "",
    })
    return response.json()

for lat, lon in zip(latitude, longitude):
    try:
        result = get_address(lat, lon)
        
        # 결과가 OK인 경우에만 처리하기 
        if result['status']['name'] == 'ok':
            road_address_name = result['results'][0]['region']['area1']['name'] + ' ' + \
                result['results'][0]['region']['area2']['name'] + ' ' + \
                result['results'][0]['region']['area3']['name'] + ' ' + \
                result['results'][0]['region']['area4']['name']

    # 도로명 주소가 있는 경우 추가하기 
            if 'road_address' in result['results'][0]:
                road_address_name += ' ' + result['results'][0]['road_address']['name']

            # 지번주소 사용하기 
            else:
                jibun_address_name = road_address_name
                if 'land' in result['results'][0]:
                    jibun_address_name += f" {result['results'][0]['land']['number1']}번길"
                    if result['results'][0]['land'].get('number2'):
                        jibun_address_name += f"-{result['results'][0]['land'].get('number2')}"
                
                addresses.append(jibun_address_name)

        else:
            addresses.append(road_address_name)

    except Exception as e:
        print(f"An error occurred: {e}")
        addresses.append('')
    
    time.sleep(0.1)  # API 호출 간에 약간의 지연 시간 추가

# 받아온 주소를 excel 파일에 저장하기
for i in range(len(addresses)):
    sheet.cell(row=i+2, column=4).value = addresses[i]

# excel 파일 저장하기
wb.save('data/불법좌회전_result.xlsx')

df = pd.read_excel('data/불법좌회전_result.xlsx')
df['Address'] = addresses  # Address라는 새 열을 만들고 그곳에 addresses 리스트의 값을 넣습니다.
df.to_excel('data/불법좌회전_result.xlsx', index=False)