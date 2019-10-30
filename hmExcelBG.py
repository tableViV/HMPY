#!/bin/python
import re,os,sys,ast,xlrd,hashlib,os.path,datetime,importlib,configparser
from pymysql import connect
from datetime import datetime
from xlrd import xldate_as_tuple


#背钩生产总表插入
hmInsertBgProduct_Str= "INSERT IGNORE INTO hm_workshop_datas(`pd_uuid`,`pd_type`,`wp_units`,`pd_batch`, `pd_create_date`, `seal_num`, `pd_num`, `pd_count`, `pd_unit`, `pd_guest`, `pd_code`, `pd_guest_code`,`pd_ramk`,`pd_cn_ramk`, `pd_out_date`,`code_num`,`reay_pcs`,`virtual_e_pcs`,`virtual_h_pcs`,`virtual_e_str`,`virtual_h_str`,`created_at`,`hm_product_state`) \
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW(),0)"

#背钩生产子表插入
hmInsertBgProduct_childStr= "INSERT IGNORE INTO hm_products_alls(`pd_uuid`,`pai_count`,`b_count`,`eye_size`, `cloth_type`,`hm_color`, `eye_fabric`,`hook_fabric`, `eye_method`, `hook_method`, `code_num`, `created_at`) \
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW())"

#背钩备料表创建
hmInsertBgSocks = "INSERT IGNORE INTO hm_stocks(`pd_uuid`,`created_at`) VALUES (%s,NOW())"

#INSERT IGNORE INTO 如果已经存在数据就跳过继续插入下一

proDir = os.path.split(os.path.realpath(__file__))[0]
configPath = os.path.join(proDir, "HM_config.ini")
config = configparser.ConfigParser()
config.read(configPath,encoding='UTF-8')
global upload_path      #文件上传路径
global download_path    #文件下载路径
global code_Url_Str     #二维码接口路径
upload_path = config.get('File_Path', 'upload_path')
download_path = config.get('File_Path', 'download_path')
code_Url_Str = config.get('QrCode_Url', 'code_Url_Str')
global dictEye          #眼类型
global dictHook         #钩类型
dictEye = ast.literal_eval(config.get('Hm_Regular', 'dictEye'))
dictHook = ast.literal_eval(config.get('Hm_Regular', 'dictHook'))
global specialGuests    #需要加数的客人
specialGuests = ast.literal_eval(config.get('Special_Guests', 'specialGuests'))


class HmSqlDatas:
    def __init__(self):
        self.masterList = []        #主表数据
        self.childList = []         #从表数据
        self.socks = []             #初始备料表
        self.sealNumList = []       #销售单号
        self.uuidList = []          #UUID列表

class HmProduct:
    '产品基类'
    pdCount = 0
    def __init__(self):
        HmProduct.pdCount += 1
        self.hm_pd_uuid = ''
        self.is_MA = False#码装
        self.ma_D = 0   #码装分母
        self.ma_U = 0   #码装分子
        self.ma_N = 0   #码装N
        self.is_E_and_H = False#钩车眼
        self.eIsSpecial = False#眼单特殊（BGLY/LY/F1/布筒）
        self.hIsSpecial = False#钩单特殊（BGLY/LY/F1/布筒）
        self.isNeedAdd = False#是否需要加数
        self.additions = 0#加数
        self.productB = 0
        self.productP = 0
        self.clothType = "-"
        self.clothTubeEye = 0
        self.clothTubeHook = 0
        self.productSize = 0
        self.productYardage = 0#码数
        self.productUnit = ''
        self.colorRamk = ''
        self.productColorNum = ''#颜色号
        self.productRamk = ''
        self.productGuest = ''#客人号
        self.productECutType = ''
        self.productHCutType = ''
        self.productReadyPcs = 0 #实际粒数
        self.productEyePcs = 0 #眼愈大粒数
        self.productHookPcs = 0 #钩愈大粒数


class hmDataHandle(object):
    def __init__(self):
        self.conn = connect(
            host='127.0.0.1', 
            port=3306, 
            db='hmpd_erp', 
            user='root', 
            password='root', 
            charset='utf8')

    def hmInsertBgProduct_sql(self,temp,data):
        cur = self.conn.cursor()
        try:
            cur.executemany(temp,data)
            self.conn.commit()
        except Exception:
            self.conn.rollback()
        finally:
            cur.close()

    def hmInsertBgProduct_sqlOne(self,temp):
        cur = self.conn.cursor()
        try:
            cur.execute(temp)
            self.conn.commit()
        except Exception:
            self.conn.rollback()
        finally:
            cur.close()


# 文档md5 用于不可描述的生产单防止重复匹配
# 中文描述
# fe65976af809170084acebb8d6af1fdd
# textValue：生产单字符串集合
# ##
def bgGetPageMd5(textValue):
    pdfTxt = textValue.replace(' ', '')
    try:
        md5 = hashlib.md5()
        enc = pdfTxt.encode('utf-8','strict')
        md5.update(enc)
        result = md5.hexdigest()
    except :
        result = ''
    finally:
        return result



# 获取产品编号/产品颜色号
# EX:J-HE-16621-EH02160
# EX:EH02160
# textValue：生产单字符串集合
# ##
def bgGetProductNumber(textValue,hmPd):
    result = ''
    try:
        productText = str(textValue)
        productText = productText.replace(' ', '')
        colorArr = productText.split('-')
        if (colorArr):
            result = colorArr[-1]
        else:
            result = ''
    except :
        result = '颜色号错误'
    finally:
        hmPd.productColorNum = result
        return result



# 获取产品类型/排数/B数/尺寸
# EX：IS/HHEE-3P-2B 3/4" 51X57MM
# EX：3
# EX：2
# EX：57
# textValue：生产单字符串集合
# ##
def bgGetProductSpecification(textValue,hmPd):
    result = ''
    productB = 0
    productP = 0
    productSize = 0
    clothType = '-'
    try:
        productText = str(textValue)
        
        productTextALL = productText

        productText = productText.replace(' ', '')

        productText = productText.replace('BG', '')
        productText = productText.replace('CB', '')
        productText = productText.replace('AB', '')

        matchObj = re.search( r'\d+[Bb]', productText, re.M|re.I)
        if matchObj:
            bCount = matchObj.group()
            bCount = bCount.replace('B', '')
            bCount = bCount.replace('b', '')
            productB= int(bCount)

        matchObj = re.search( r'\d+[Pp]', productText, re.M|re.I)
        if matchObj:
            pCount = matchObj.group()
            pCount = pCount.replace('P', '')
            pCount = pCount.replace('p', '')
            productP = int(pCount)

        matchObj = re.search( r'[Xx]\d+', productText, re.M|re.I)
        if matchObj:
            sizeCount = matchObj.group()
            sizeCount = sizeCount.replace('X', '')
            sizeCount = sizeCount.replace('x', '')
            productSize = int(sizeCount)

        clothObj1 = re.search( r'BGLY-HF', productTextALL, re.M|re.I)
        clothObj2 = re.search( r'BGLY-MCB', productTextALL, re.M|re.I)
        clothObj3 = re.search( r'MICF', productTextALL, re.M|re.I)
        clothObj4 = re.search( r'MCBF1', productTextALL, re.M|re.I)
        clothObj5 = re.search( r'MCBLY', productTextALL, re.M|re.I)
        clothObj6 = re.search( r'MCBEE', productTextALL, re.M|re.I)
        clothObj7 = re.search( r'MICFLY', productTextALL, re.M|re.I)
        clothObj8 = re.search( r'SATIN', productTextALL, re.M|re.I)
        clothObj9 = re.search( r'SATINF1', productTextALL, re.M|re.I)
        clothObj10 = re.search( r'TIT', productTextALL, re.M|re.I)
        clothObj11 = re.search( r'HFLY', productTextALL, re.M|re.I)
        clothObj12 = re.search( r'MCB', productTextALL, re.M|re.I)
        clothObj13 = re.search( r'BGLY-FABRIC', productTextALL, re.M|re.I)
        clothObj14 = re.search( r'LY', productTextALL, re.M|re.I)
        clothObj15 = re.search( r'HE', productTextALL, re.M|re.I)
        clothObj16 = re.search( r'HF-', productTextALL, re.M|re.I)
        clothObj17 = re.search( r'HFF1', productTextALL, re.M|re.I)

        if clothObj1:
            clothType = "尼龙布+拉架布"
        elif clothObj2:
            clothType = "MCB布+拉架布"
        elif clothObj3:
            clothType = "MICF布"
        elif clothObj4:
            clothType = "MCB布+F1毛布"
        elif clothObj5:
            clothType = "MCB布+拉架布"
        elif clothObj6:
            clothType = "MCB布+尼龙布"
        elif clothObj7:
            clothType = "MICF布+拉架布"
        elif clothObj8:
            clothType = "色丁布"
        elif clothObj9:
            clothType = "色丁布+F1毛布"
        elif clothObj10:
            clothType = "塔夫绸"
        elif clothObj11:
            clothType = "尼龙布+拉架布"
        elif clothObj12:
            clothType = "MCB布"
        elif clothObj13:
            clothType = "MCB布+拉架布"
        elif clothObj14:
            clothType = "拉架布"
        elif clothObj15:
            clothType = "尼龙布+毛布"
        elif clothObj16:
            clothType = "尼龙布"
        elif clothObj17:
            clothType = "尼龙布+F1毛布"

    except :
        productB = 0
        productP = 0
        productSize = 0
        clothType = "-"
    finally:
        hmPd.productB = productB
        hmPd.productP = productP
        hmPd.productSize = productSize
        hmPd.clothType = clothType
        return result


##
# 计算布料码数
# 【M】= S订单数 / （914 / 尺寸）
# hookEyePillCount   粒数
# productSize        尺寸
# productB           B数
# 
# ##
def countHmYard(hmPd):
    result = 0
    productSum = hmPd.productCount
    productSize = hmPd.productSize

    try:
        if ( hmPd.productUnit == 'YDS'or hmPd.productUnit == 'YARD'):
            result =  productSum
        else: 
            intValue = ( 914 / productSize )
            intValueM = intValue
            yardCount = productSum / intValueM
            result = yardCount

    except :
        result = 0
    finally:
        return round(result,2)


    


# 获取product产品钩眼布筒字段类型,clothTubeHook=0是钩车眼
# 布筒:55MM
# 布筒:22MM
# textValue：生产单字符串集合
# productObj：对象
#  ##
def bgGetProductClothTube(textValue,productObj):
    hmClothTubeEye   =  0
    hmClothTubeHook  =  0
    try:
        product_matchObj = re.finditer( r'布筒(?::|：)\s?(\d*)MM', textValue)
        if product_matchObj:

            tubeArr = []
            for match in product_matchObj:
                item = match.group()
                item_matchObj = re.search( r'(\d\d\d|\d\d|\d)', item, re.M|re.I)

                if (item_matchObj):  
                    tubeItem = item_matchObj.group()
                    tubeArr.append(tubeItem)

            #判断钩车眼数量，如果是2的话有钩有眼；如果是1的话就有可能是钩车眼
            if(len(tubeArr) == 2 ):
                hmClothTubeEye   =  tubeArr[0]
                hmClothTubeHook  =  tubeArr[1]

            elif(len(tubeArr) == 1 ):
                hmClothTubeEye   =  tubeArr[0]

            else:
                hmClothTubeEye   =  0
                hmClothTubeHook  =  0
                #眼
                #钩
        else :
            hmClothTubeEye   =  0
            hmClothTubeHook  =  0

    except :
        hmClothTubeEye   =  0
        hmClothTubeHook  =  0
    finally:
        productObj.clothTubeEye   =  int(hmClothTubeEye)
        productObj.clothTubeHook  =  int(hmClothTubeHook)



# 获取product产品钩眼切类型字段类型,clothTubeHook=0是钩车眼
# 圆角热切
# 直角热切
# textValue：生产单字符串集合
# productObj：对象
#  ##
def bgGetProductCutType(textValue,productObj):
    productECutType = ''
    productHCutType = ''
    try:
        product_matchObj = re.finditer( r'[反对四有圆圓A直N][^,，。.:：\*\(（]{2,15}切', textValue)
        if product_matchObj:

            cutArr = []
            for match in product_matchObj:

                if (match):  
                    item = match.group()
                    item = item.replace('\n', '')
                    cutArr.append(item)

            #判断钩车眼数量，如果是2的话有钩有眼；如果是1的话就是钩车眼或单眼
            if(len(cutArr) == 2 ):
                productECutType  =  cutArr[0]
                productHCutType  =  cutArr[1]

            elif(len(cutArr) == 1 and productObj.clothTubeHook == 0):
                productECutType  =  cutArr[0]
                productHCutType  =  ''
            else:
                productECutType = ''
                productHCutType = ''
                #眼
                #钩
        else :
            productObj.productECutType = ''
            productObj.productHCutType = ''

    except :
        productObj.productECutType = ''
        productObj.productHCutType = ''
    else:
        productObj.productECutType = productECutType
        productObj.productHCutType = productHCutType


# 获取product中文详细备注
# textValue：生产单字符串集合
# ##销售订单
def getGuestAdditions(hmPd):
    result = 0
    try:
        product_matchObj = re.search( r'\+\d+', hmPd.productRamkadd, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = productSN.replace('+', '')
            productSN = productSN.replace(' ', '')
            result = int(productSN)
        else:
            result = 50
    except :
        result = 50
    finally:
        hmPd.additions = result
        return result

# 获取product眼类型
# textValue：生产单字符串集合
# ##销售订单
def bgGetEye(textValue,hmPd):
    result = ''
    resultType = False
    try:
        product_matchObj = re.search( r'[比练不黑红金玫尼普铜无哑][^,，++。.:：\*\(（）))]{1,15}眼', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            result = productSN
            resultType = bgCheckSpHookEye(productSN)
        else:
            result = ''
            resultType = True
    except :
        result = ''
        resultType = True
    finally:
        hmPd.productEyeStr = result
        hmPd.is_EyeSp = resultType
        return result

# 获取product钩类型
# textValue：生产单字符串集合
# ##销售订单
def bgGetHook(textValue,hmPd):
    result = ''
    resultType = False
    try:
        product_matchObj = re.search( r'[比练不黑红金玫尼普铜无哑][^,，++。.:：\*\(（））)]{1,15}[钩勾]', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            result = productSN
            resultType = bgCheckSpHookEye(productSN)
        else:
            result = ''
            resultType = True
    except :
        result = ''
        resultType = True
    finally:
        hmPd.productHookStr = result
        hmPd.is_HookSp = resultType
        return result

# 判断product特殊钩眼类型
# hookOrEye:判断特殊钩眼
# ##销售订单
def bgCheckSpHookEye(hookOrEye):
    result = True
    try:
        sp_matchObj = re.search( r'(尼龙)|(尼龍)', hookOrEye, re.M|re.I)
        if sp_matchObj:
            result = False
        else:
            result = True
    except :
        result = True
    finally:
        return result

def clearNullStr(textValue):
    result = textValue.replace('\n', '')
    result = result.replace(' ', '')
    return result

##
# 特殊情况判断
# /钩车眼/压字/印字/BGLY/LY/F1/码装/需要愈多的客人加数
# 有是否有钩眼单
# 1,2,2,2,2,2,2
#
# ##
def getHmHookEyeIsSpecial(hmPdObject):

    eIsSpecial = False
    hIsSpecial = False

    try:
        #以下客人需要多加数量付
        try:
            specialGuestsNum = specialGuests[hmPdObject.productGuest]
            hmPdObject.isNeedAdd = True
        except :
            specialGuestsNum = 0
        specialGuestsNum += 0

        if hmPdObject.isNeedAdd :
            getGuestAdditions(hmPdObject)

        bgly_matchObj = re.search( r'(码装)|(YARD)|(YDS)', hmPdObject.productSf, re.M|re.I)

        if (bgly_matchObj or hmPdObject.productUnit == 'YDS'or hmPdObject.productUnit == 'YARD'):
            hmPdObject.is_MA = True
            
            bgly_matchObj = re.search( r'\d*/*\d+"', hmPdObject.productSf, re.M|re.I)
            if (bgly_matchObj):
                productSN = bgly_matchObj.group()
                productSN = productSN.replace('"', '')
                hmArr = productSN.split('/')

            if (len(hmArr) == 2 ):
                hmPdObject.ma_D = int(hmArr[1])  #码装分母
                hmPdObject.ma_U = int(hmArr[0])  #码装分子
            else:
                hmPdObject.ma_D = 1
                hmPdObject.ma_U = 1
            
            if(hmPdObject.ma_D == 2 and hmPdObject.ma_U == 1):
                hmPdObject.ma_N = 57
            elif(hmPdObject.ma_D == 4 and hmPdObject.ma_U == 3):
                hmPdObject.ma_N = 48
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 11):
                hmPdObject.ma_N = 52
            elif(hmPdObject.ma_D == 1 and hmPdObject.ma_U == 1):
                hmPdObject.ma_N = 36
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 15):
                hmPdObject.ma_N = 38
            elif(hmPdObject.ma_D == 32 and hmPdObject.ma_U == 19):
                hmPdObject.ma_N = 60
            elif(hmPdObject.ma_D == 38 and hmPdObject.ma_U == 15):
                hmPdObject.ma_N = 38
            elif(hmPdObject.ma_D == 8 and hmPdObject.ma_U == 5):
                hmPdObject.ma_N = 58
            elif(hmPdObject.ma_D == 8 and hmPdObject.ma_U == 7):
                hmPdObject.ma_N = 41
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 9):
                hmPdObject.ma_N = 64
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 13):
                hmPdObject.ma_N = 44
            elif(hmPdObject.ma_D == 64 and hmPdObject.ma_U == 25):
                hmPdObject.ma_N = 72
            elif(hmPdObject.ma_D == 32 and hmPdObject.ma_U == 21):
                hmPdObject.ma_N = 55
            else:
                hmPdObject.ma_N = 0

            hmPdObject.eIsSpecial = False
            hmPdObject.is_E_and_H = False
            return True

        bgly_matchObj = re.search( r'(钩眼车)|(钩加眼车)|(钩与眼车)', hmPdObject.productRamk, re.M|re.I)
        if (bgly_matchObj): 
            hmPdObject.eIsSpecial = True
            hmPdObject.is_E_and_H = True
            return True

        #B数大于等于5
        if hmPdObject.productB >= 5:
            eIsSpecial = True
            hIsSpecial = True
        
        #排数大于等于4
        if hmPdObject.productP >= 4:
            eIsSpecial = True
            hIsSpecial = True


        #BGLY|LY|F1 和其它特殊情况
        bgly_matchObj = re.search( r'(BGLY)|(LY)|(F1)', hmPdObject.productSf, re.M|re.I)
        if (bgly_matchObj): 
            eIsSpecial = True
            hIsSpecial = True

        #眼是否散口切
        bgly_matchObj = re.search( r'(钩位压字)|(勾位压字)|(钩压)', hmPdObject.productRamk, re.M|re.I)
        if (bgly_matchObj):
            hmPdObject.productIsHookPressureWord = True 
            hIsSpecial = True

        bgly_matchObj = re.search( r'(眼位压字)|(眼压)', hmPdObject.productRamk, re.M|re.I)
        if (bgly_matchObj): 
            hmPdObject.productIsEyePressureWord = True 
            eIsSpecial = True

        if (hmPdObject.clothTubeEye >= 66): 
            eIsSpecial = True

        if (hmPdObject.clothTubeHook >= 66): 
            hIsSpecial = True

        bgly_matchObj = re.search( r'散口', hmPdObject.productECutType, re.M|re.I)
        if (bgly_matchObj):
            hmPdObject.productIsEyeLooseCut = True  
            eIsSpecial = True

        bgly_matchObj = re.search( r'散口', hmPdObject.productHCutType, re.M|re.I)
        if (bgly_matchObj): 
            hIsSpecial = True

        bgly_matchObj = re.search( r'四角圆角', hmPdObject.productECutType, re.M|re.I)
        if (bgly_matchObj):
            eIsSpecial = True

        bgly_matchObj = re.search( r'四角圆角', hmPdObject.productHCutType, re.M|re.I)
        if (bgly_matchObj): 
            hIsSpecial = True

    except :
        hmPdObject.eIsSpecial = False
        hmPdObject.hIsSpecial = False

    else:
        hmPdObject.eIsSpecial = eIsSpecial
        hmPdObject.hIsSpecial = hIsSpecial
        return True


# 计算钩眼粒数
# （粒数）【T】= 订单数量 * 预多百分数 * B数
# 
# btType       钩眼类型  H/E
# productSum   订单总数
# productB    B数
# productP    P数
# ##
def countHmPillx(hookEyeType,hmPdObject):
    productSum = hmPdObject.productCount
    productB = hmPdObject.productB
    isSpecial = (hmPdObject.eIsSpecial if (hookEyeType == 'E') else hmPdObject.hIsSpecial) 
    isSpHookEye = (hmPdObject.is_EyeSp if (hookEyeType == 'E') else hmPdObject.is_HookSp) 
    if(hmPdObject.is_MA):
        productSum *= hmPdObject.ma_N

    #以下客人需要多加50付
    if hmPdObject.isNeedAdd:
        productSum += hmPdObject.additions

    try:
        result = 0
        #预多增量
        beforeCount = 1
        #初始订单数
        beginSum = productSum

        if isSpecial:
            beforeCount += 0.06
                          
        if productSum <= 500:
            if isSpHookEye:
                beforeCount += (0.12 if (hookEyeType == 'E') else 0.06)
            else:
                beforeCount += (0.096 if (hookEyeType == 'E') else 0.048)
            beginSum += (50 if (hmPdObject.is_MA==False) else 0)

        elif 501 <= productSum and productSum <= 2000:
            if isSpHookEye:
                beforeCount += (0.1 if (hookEyeType == 'E') else 0.06)
            else:
                beforeCount += (0.08 if (hookEyeType == 'E') else 0.048)  
            beginSum += (50 if (hmPdObject.is_MA==False) else 0)
        
        elif 2001 <= productSum and productSum <= 5000 :
            if isSpHookEye:
                beforeCount += (0.08 if (hookEyeType == 'E') else 0.06)
            else:
                beforeCount += (0.064 if (hookEyeType == 'E') else 0.048)
            beginSum += (50 if (hmPdObject.is_MA==False) else 0)
        
        elif 5001 <= productSum and productSum <= 10000 :
            if isSpHookEye:
                beforeCount += (0.06 if (hookEyeType == 'E') else 0.05)
            else:
                beforeCount += (0.048 if (hookEyeType == 'E') else 0.04)
        else:
            if isSpHookEye:
                beforeCount += (0.05 if (hookEyeType == 'E') else 0.04)
            else:
                beforeCount += (0.04 if (hookEyeType == 'E') else 0.032)

        if(hmPdObject.is_MA):
            result = beginSum * beforeCount 
        else:
            result = beginSum * beforeCount * productB
            
    except :
        return 0
    else:
        return round(result)


# 计算钩眼粒数
# （粒数）【T】= 订单数量 * 预多百分数 * B数
# 
# btType       钩眼类型  H/E
# productSum   订单总数
# productB    B数
# productP    P数
# ##
def countHmPillr(hmPdObject):
    productSum = hmPdObject.productCount
    productB = hmPdObject.productB
    result = 0
    if hmPdObject.is_MA:
        result = productSum * hmPdObject.ma_N
        return result
    else :
        result = productSum * productB
        return result


def bgGetHookEyeKg(hookEyeType,hmPd):
    result = 0
    hookEyeKg = 0
    productHEPill = 0
    hookEyeStr = ''
    productEPillSum = 0
    try:
        if hookEyeType =='E':
            hookEyeStr = hmPd.productEyeStr
            productHEPill = hmPd.productEPill
        else :
            hookEyeStr = hmPd.productHookStr
            productHEPill = hmPd.productHPill
        #如果钩眼种类找不到，粒数为0，返回空。
        if hookEyeStr == '':
            return 0
        if productHEPill == 0:
            return 0

        if hookEyeType =='E':
            hookEyeKg = dictEye[hookEyeStr]
            productEPillSum = hmPd.productEPill *  hmPd.productP
        else :
            hookEyeKg = dictHook[hookEyeStr]
            productEPillSum = hmPd.productHPill 

        result = productEPillSum / hookEyeKg 
    except :
        return 0
    else:
        return round(result,3)

##
# 这里是转换Excel中的Data格式
# ##
def get_excel_data(fileName,hmSqlDatas):
    text_path = upload_path + fileName + ".xls"
    workbook = xlrd.open_workbook(text_path)
    sheet = workbook.sheets()[0]  # 读取第一个sheet
    nrows = sheet.nrows  # 行数
    first_row_values = sheet.row_values(0)  # 第一行数据
    num = 1
    for row_num in range(1, nrows):
        row_values = sheet.row_values(row_num)
        
        #init
        hmPd = HmProduct()
        pd_uuid = ''
        pd_seal = ''

        masterList_obj = []
        childList_obj = []
        socks_obj = []

        #开始组装UUID,用于构造两个表第一列的uuid
        if row_values:
            #还要做部门验证
            pd_seal = sheet.cell_value(num, 4)
            pd_num = sheet.cell_value(num, 5)
            pd_code = sheet.cell_value(num, 9)
            pd_batch = sheet.cell_value(num, 2)
            #生产单号+产品编号+批次
            pd_uuid = str(pd_num) + str(pd_code) + str(pd_batch)
            pd_uuid = bgGetPageMd5(pd_uuid)

            masterList_obj.append(pd_uuid)
            childList_obj.append(pd_uuid)
            socks_obj.append(pd_uuid)

            hmPd.productUnit = sheet.cell_value(num, 7)
            hmPd.productGuest = sheet.cell_value(num, 8)
            hmPd.productSf = sheet.cell_value(num, 0)
            hmPd.productRamk = sheet.cell_value(num, 12)
        #开始组装两个表的其它字段
        for i in range(len(first_row_values)):
            ctype = sheet.cell(num, i).ctype
            cell = sheet.cell_value(num, i)
            if ctype == 2 and cell % 1 == 0.0:  # ctype为2且为浮点
                cell = int(cell)  # 浮点转成整型
                #cell = str(cell)  # 转成整型后再转成字符串，如果想要整型就去掉该行
            elif ctype == 3:
                date = datetime(*xldate_as_tuple(cell, 0))
                cell = date.strftime('%Y/%m/%d %H:%M:%S')
            elif ctype == 4:
                cell = True if cell == 1 else False

            masterList_obj.append(cell)
            #排数、b数、眼尺寸、布类型
            if i == 0:
                bgGetProductSpecification(cell,hmPd)
                childList_obj.append(hmPd.productP)
                childList_obj.append(hmPd.productB)
                childList_obj.append(hmPd.productSize)
                childList_obj.append(hmPd.clothType)
            if i == 6:
                hmPd.productCount = int(cell)
            #颜色号
            elif i == 9:
                bgGetProductNumber(cell,hmPd)
                childList_obj.append(hmPd.productColorNum)
            #布筒、切法、
            elif i == 12:
                bgGetProductClothTube(cell,hmPd)
                bgGetProductCutType(cell,hmPd)
                bgGetEye(cell,hmPd)
                bgGetHook(cell,hmPd)

                hmPd.productYardage = countHmYard(hmPd)
                getHmHookEyeIsSpecial(hmPd)
                hmPd.productEyePcs = countHmPillx('E',hmPd)
                hmPd.productHookPcs = countHmPillx('H',hmPd)
                hmPd.productReadyPcs = countHmPillr(hmPd)
                #产品钩眼公斤数
                hmPd.productEyeKG = bgGetHookEyeKg('E',hmPd)
                hmPd.productHookKG = bgGetHookEyeKg('H',hmPd)

                childList_obj.append(hmPd.clothTubeEye)
                childList_obj.append(hmPd.clothTubeHook)
                childList_obj.append(hmPd.productECutType)
                childList_obj.append(hmPd.productHCutType)
                childList_obj.append(hmPd.productYardage)

            elif i == 13:
                masterList_obj.append(hmPd.productYardage)
                masterList_obj.append(hmPd.productReadyPcs)
                masterList_obj.append(hmPd.productEyePcs)
                masterList_obj.append(hmPd.productHookPcs)
                masterList_obj.append(hmPd.productEyeStr)
                masterList_obj.append(hmPd.productHookStr)

        hmSqlDatas.masterList.append(masterList_obj)
        hmSqlDatas.childList.append(childList_obj)
        hmSqlDatas.socks.append(socks_obj)
        
        if(pd_seal not in hmSqlDatas.sealNumList):
            hmSqlDatas.sealNumList.append(pd_seal)
        hmSqlDatas.uuidList.append(pd_uuid)
        num += 1

    print(text_path)
    return hmSqlDatas

def parse(fileName):
    hmDatas = HmSqlDatas()
    
    hmSql = hmDataHandle()
    get_excel_data(fileName,hmDatas)
    if hmDatas.masterList:
        hmSql.hmInsertBgProduct_sql(hmInsertBgProduct_Str,hmDatas.masterList)
        
    hmSql2 = hmDataHandle()
    if hmDatas.childList:
       hmSql2.hmInsertBgProduct_sql(hmInsertBgProduct_childStr,hmDatas.childList)
       
    hmSql3 = hmDataHandle()
    if hmDatas.socks:
       hmSql3.hmInsertBgProduct_sql(hmInsertBgSocks,hmDatas.socks)

    #更新数据库中，包含Excel的数据全部设置为失效产品。
    if hmDatas.sealNumList:
        for item in hmDatas.sealNumList :
            hmSql6 = hmDataHandle()
            sql = "UPDATE hm_workshop_datas SET `is_change`=1 WHERE `seal_num` ='"+item+"'"
            hmSql6.hmInsertBgProduct_sqlOne(sql)

    #重新设置为有效产品。
    hmSql4 = hmDataHandle()
    if hmDatas.uuidList:
       sss = "INSERT INTO hm_workshop_datas (`pd_uuid`,`is_change`)VALUES(%s,0) ON DUPLICATE KEY UPDATE `is_change`=0"
       hmSql4.hmInsertBgProduct_sql(sss,hmDatas.uuidList)

if __name__ == '__main__':
    shellValues = sys.argv[1]
    shellArr =  shellValues.split(",")
    flag = int(shellArr[1])
    fileName = shellArr[0]
    res = parse(fileName)
