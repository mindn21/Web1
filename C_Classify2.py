from tabnanny import check
from numpy import record
import pandas as pd
import A_GetBrandCode
import re


#--------------------------------------------
#코드별로 금액이 아닌 수량 기준으로 분류하는 코드
#--------------------------------------------


df = pd.read_excel('TARGET_ORDER.xlsx') #<- 추후엔 B 에서 df_del 불러오는 방식으로
df = df.fillna(0)


#세트코드들 추출
df_SetList = pd.read_excel('BRAND_LIST.xlsx',sheet_name="Set_LIST").fillna(0)
setBrand_list = [] #세트 코드 저장용 리스트
setBrand_list = df_SetList["세트코드"].tolist()
# print(setBrand_list)


onlyCode3_list = [] #복수도 아니고 세트도 아닌 찐 제품 코드

onlyCode2_list = A_GetBrandCode.onlyCode_list
for code in onlyCode2_list:
    #print(code)
    if type(code) == int:
        onlyCode3_list.append(code)

    #세트 코드이면 저장하지 않음
    elif code in setBrand_list:
        continue

    else:
        checkPop = code.find('-')
        if checkPop == -1:
            onlyCode3_list.append(code)


# print(onlyCode3_list)
n = len(onlyCode3_list)

recordQuantity_list = [0 for i in range(n)] 
#날짜별 코드의 수량을 저장하기 위한 list. onlyCode3_list와 똑같은 순서로 구성


dayRecordQuantity_dict = {} #날짜별 수량을 날짜에 매칭해서 저장하는 dict

firstDay = df["주문일자"][0] #첫번째 날짜를 가져오기
dayRecordQuantity_dict[firstDay] = {}

k = 0
unknownCode = []

#print(df_SetList)
#print(A_GetBrandCode.BrandCode_dict.values())
valueList = []

for val in A_GetBrandCode.BrandCode_dict.values():
    for vval in val:
        valueList.append(vval)

#print(valueList)
print(onlyCode3_list)

for ind in df.index:
    read_date = df["주문일자"][ind]
    read_code = df["판매자상품코드"][ind]
    read_mall = df["쇼핑몰"][ind]
    read_numb = df["수량"][ind]
    read_pric = df["실결제금액"][ind]
    #print(read_code)
    if read_code == 0:
        continue

    elif read_date in dayRecordQuantity_dict:
        
        if read_code in valueList: #행여나 새로운 코드가 등장해도 문제가 생기지 않도록
            #print('yes' + str(ind))
        
            #복수 코드 구분하기----------------------------------
            checkContain = read_code.find('-')


            #세트코드 구분하기-------------------------
            if read_code in setBrand_list:
                #print("세트코드")
                part1 = ""
                part1_num = 0
                part2 = ""
                part2_num = 0
                part3 = ""
                part3_num = 0

                #해당 코드의 인덱스 가져오기
                idd = df_SetList.index[df_SetList['세트코드']==read_code].tolist()

                #각각 구성 저장하기
                part1 = df_SetList['구성1'][idd].item()
                part1_num = int(df_SetList['수량1'][idd])
                kk1 = onlyCode3_list.index(part1)
                recordQuantity_list[kk1] += read_numb * part1_num
                
                part2 = df_SetList['구성2'][idd].item()
                if part2 != 0:
                    part2_num = int(df_SetList['수량2'][idd])
                    kk2 = onlyCode3_list.index(part2)
                    recordQuantity_list[kk2] += read_numb * part2_num

                part3 = df_SetList['구성3'][idd].item()
                if part3 != 0:
                    
                    part3_num = int(df_SetList['수량3'][idd])
                    kk3 = onlyCode3_list.index(part3)
                    recordQuantity_list[kk3] += read_numb * part3_num



            elif checkContain == -1:
                k = onlyCode3_list.index(read_code)
                recordQuantity_list[k] += read_numb


            else: 
                #print(read_code)
                #print("contain =" + str(checkContain))

                #띄어쓰기 된 부분 없애기. '-'이 없는데도 띄어쓰기 되 있는 코드도 있으니 여기 안에 넣어야 함
                checkSpace = read_code.find(' ')
                if checkSpace != -1:
                    read_code=re.sub(' ', '',read_code)
                
                #띄어쓰기가 지워졌으니 '-'의 문자열 인덱스도 변함
                checkContain = read_code.find('-')

                #추가된 수량 가져오기
                changedQuantity = int(read_code[-1])

                #원래 코드만 추출하기
                read_code = read_code[0:checkContain]
                #print('수정된 read_code ' + read_code)
                k2 = onlyCode3_list.index(read_code)
                recordQuantity_list[k2] += read_numb * changedQuantity


        else: 
            #Brand_List에 없는 새로운 코드 저장
            #print('no' + str(ind))
            if read_code in unknownCode:
                continue
            else: 
                unknownCode.append(read_code)
            continue

        dayRecordQuantity_dict[read_date] = recordQuantity_list








    else: 
        dayRecordQuantity_dict[read_date] = {}
        recordQuantity_list = [0 for i in range(n)] 
        
        if read_code in valueList: #행여나 새로운 코드가 등장해도 문제가 생기지 않도록
            

            #복수 코드 구분하기----------------------------------
            checkContain = read_code.find('-')


            #세트코드 구분하기-------------------------
            if read_code in setBrand_list:
                
                part1 = ""
                part1_num = 0
                part2 = ""
                part2_num = 0
                part3 = ""
                part3_num = 0

                #해당 코드의 인덱스 가져오기
                
                idd = df_SetList.index[df_SetList['세트코드']==read_code].tolist()

                #각각 구성 저장하기
                part1 = df_SetList['구성1'][idd].item()
                part1_num = int(df_SetList['수량1'][idd])
                kk1 = onlyCode3_list.index(part1)
                recordQuantity_list[kk1] += (read_numb * part1_num)
                
                part2 = df_SetList['구성2'][idd].item()
                if part2 != 0:
                    
                    part2_num = int(df_SetList['수량2'][idd])
                    kk2 = onlyCode3_list.index(part2)
                    recordQuantity_list[kk1] += (read_numb * part2_num)

                part3 = df_SetList['구성3'][idd].item()
                if part3 != 0:
                    
                    part3_num = int(df_SetList['수량3'][idd])
                    kk3 = onlyCode3_list.index(part3)
                    recordQuantity_list[kk1] += (read_numb * part3_num)



            elif checkContain == -1:
                k = onlyCode3_list.index(read_code)
                recordQuantity_list[k] += read_numb


            else: 
                #print(read_code)
                #print("contain =" + str(checkContain))

                #띄어쓰기 된 부분 없애기. '-'이 없는데도 띄어쓰기 되 있는 코드도 있으니 여기 안에 넣어야 함
                checkSpace = read_code.find(' ')
                if checkSpace != -1:
                    read_code=re.sub(' ', '',read_code)
                
                #띄어쓰기가 지워졌으니 '-'의 문자열 인덱스도 변함
                checkContain = read_code.find('-')

                #추가된 수량 가져오기
                changedQuantity = int(read_code[-1])

                #원래 코드만 추출하기
                read_code = read_code[0:checkContain]
                #print('수정된 read_code ' + read_code)
                k2 = onlyCode3_list.index(read_code)
                recordQuantity_list[k2] += read_numb * changedQuantity

        else: 
            #print(read_code)
            #Brand_List에 없는 새로운 코드 저장
            unknownCode.append(read_code)
            continue

        dayRecordQuantity_dict[read_date] = recordQuantity_list

#print(dayRecordQuantity_dict)
print("--날짜별 브랜드 수량 매칭 완료--")   

#print(A_GetBrandCode.BrandCode_dict.items())
#print(unknownCode)