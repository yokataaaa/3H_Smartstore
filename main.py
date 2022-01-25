from all4lab_smartstore import convert_all4lab_to_excel
from lklab_smartstore import convert_lklab_to_excel
from selenium import webdriver
from openpyxl import Workbook
import openpyxl
import time

url_file = 'url_to_excel.xlsx'
wb = openpyxl.load_workbook(url_file)
ws = wb['lklab']

url_list = list()

# 엑셀 파일 마지막 행 찾기
current_row = 0
for cell in ws['E']:
    current_row += 1
    url_list.append(cell.value)
    if cell.value == None:
        break

print(current_row)
wb.close()

# 올포랩
# file_name = 'test.xlsx'
# dst_url = 'https://www.allforlab.com/pdt/SL21CAT00002439?keywords='
# category = 50003304    # DEFAULT = 50003439
# mode = 'w'      # w 새로 쓰기, r 이어 쓰기
# convert_all4lab_to_excel(dst_url, file_name, category, mode)

# 엘케이랩
# url = 'http://www.lklab.com/product/product_info.asp?g_no=8743&t_no=745'
# url = 'http://www.lklab.com/product/product_info.asp?g_no=5444&t_no=780'
# url = 'http://www.lklab.com/product/product_info.asp?g_no=5451&t_no=780'

category = 50003439
new_file = 'form_{}.xlsx'
mode = 'r'          # s 초기화  w 덮어쓰기(새파일 생성후)  r 이어쓰기

flag = 0
if mode == 's':
    convert_lklab_to_excel(url_list[0], new_file.format(flag), category, mode)
else:
    for url in url_list[1:]:
        if convert_lklab_to_excel(url, new_file.format(flag), category, mode):
            print(url+' 스마트 스토어 엑셀로 전환 성공!')
        else:
            flag += 1
            convert_lklab_to_excel(url, new_file.format(flag), category, 's')
            convert_lklab_to_excel(url, new_file.format(flag), category, mode)


# 웹 호스팅 https://www.pythonanywhere.com/user/principe84/files/home/principe84/
# file_name 스마트스토어 양식 엑셀파일 이름 ex) test.xlsx
# dst_url 스크래핑 대상 url ( 올포랩 : https://www.allforlab.com/ )
# 스마트스토어 카테고리ID
# 기초실험장비 => 50003439	생활/건강	공구	측정공구	기타측정기  50003306	가구/인테리어	서재/사무용가구	사무/교구용가구	기타사무/교구용가구
# 분석,여과,측정 => 50003439	생활/건강	공구	측정공구	기타측정기
# 유리기구 => 50004538   생활/건강	주방용품	잔/컵	유리컵   50005257	생활/건강	주방용품	보관/밀폐용기	기타보관용기
# 플라스틱기구 => 50004576	생활/건강	주방용품	보관/밀폐용기	플라스틱용기 50004541	생활/건강	주방용품	잔/컵	플라스틱컵
# 실험소모품 => 50003304 가구/인테리어	서재/사무용가구	사무/교구용가구	교구용가구/소품
# 시약,화학 => 50003539	생활/건강	공구	페인트용품	기타페인트용품
# 실험 안전 보호기구 => 50003454	생활/건강	공구	안전용품	기타안전용품
# 사무,일반,생활 => 50003304	가구/인테리어	서재/사무용가구	사무/교구용가구	교구용가구/소품
