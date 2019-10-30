import re,gc,sys,ast,uuid,time,pyocr,string,PyPDF2,qrcode,hashlib,os.path,datetime,importlib,configparser
from ctypes import cdll
from pymysql import connect
from decimal import Decimal
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import portrait
from reportlab.pdfbase.ttfonts import TTFont
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfparser import PDFDocument
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
from pdfminer.layout import LTTextBoxHorizontal,LAParams,LTTextBox
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
#----------------1.Code run time----------------#
importlib.reload(sys)
timeStart = time.time()
#----------------2.Configuration file initialization----------------#
proDir              = os.path.split(os.path.realpath(__file__))[0]
configPath          = os.path.join(proDir, "HM_config_DEV.ini")
config              = configparser.ConfigParser()
config.read(configPath,encoding='UTF-8')
global upload_path,download_path,code_Url_Str       #文件上传路径,文件下载路径,二维码接口路径
upload_path         = config.get('File_Path', 'upload_path')
download_path       = config.get('File_Path', 'download_path')
code_Url_Str        = config.get('QrCode_Url', 'code_Url_Str')
global dictEye,dictHook                             #眼类型,钩类型
dictEye             = ast.literal_eval(config.get('Hm_Regular', 'dictEye'))
dictHook            = ast.literal_eval(config.get('Hm_Regular', 'dictHook'))
global specialGuests                                #需要加数的客人
specialGuests       = ast.literal_eval(config.get('Special_Guests', 'specialGuests'))
global BGLYEye,BGLYHook,BGLYOther                   #专利背钩工序配置      
BGLYEye             = ast.literal_eval(config.get('OptionsBGLY', 'optionsCarBGLY'))
BGLYHook            = ast.literal_eval(config.get('OptionsBGLY', 'optionsCatBGLY'))
BGLYOther           = ast.literal_eval(config.get('OptionsBGLY', 'optionsOtherBGLY'))
global FBEye,FBHook,FBOther                         #反包布工序配置       
FBEye               = ast.literal_eval(config.get('OptionsFB', 'optionsCarFB'))
FBHook              = ast.literal_eval(config.get('OptionsFB', 'optionsCatFB'))
FBOther             = ast.literal_eval(config.get('OptionsFB', 'optionsOtherFB'))
global HEEye,HEHook,HEOther                         #钩车眼工序配置
HEEye               = ast.literal_eval(config.get('OptionsHE', 'optionsCarHE'))
HEHook              = ast.literal_eval(config.get('OptionsHE', 'optionsCatHE'))
HEOther             = ast.literal_eval(config.get('OptionsHE', 'optionsOtherHE'))
global PUBLICEye,PUBLICHook,PUBLICOther             #普通背钩工序配置
PUBLICEye           = ast.literal_eval(config.get('OptionsPUBLIC', 'optionsCarPUBLIC'))
PUBLICHook          = ast.literal_eval(config.get('OptionsPUBLIC', 'optionsCatPUBLIC'))
PUBLICOther         = ast.literal_eval(config.get('OptionsPUBLIC', 'optionsOtherPUBLIC'))
global OTHERPeople,SpecialGuestsBag,SpecialGuestsFbb#间接人员单价列表,单独含有包钩工序的客人,含反包布特殊工序的客人
OTHERPeople         = ast.literal_eval(config.get('OtherPeople', 'peopleList'))
SpecialGuestsBag    = ast.literal_eval(config.get('Special_Guests_Bag', 'specialGuestsBag'))
SpecialGuestsFbb    = ast.literal_eval(config.get('Special_Guests_Bag', 'specialGuestsFbb'))
#----------------3.PDF font initialization----------------#
# In Windows
hmFont = '微软雅黑'
pdfmetrics.registerFont(TTFont(hmFont, 'msyh.ttf'))
# In MacOS
#hmFont = 'Songti'
#pdfmetrics.registerFont(TTFont('Songti', 'Songti.ttc'))
#----------------4.Custom class----------------#
class HmProduct:
    'ProductBaseClass'
    pdCount = 0
    def __init__(self):
        HmProduct.pdCount += 1
        self.hm_pd_uuid = ''                    #UUID
        self.is_MA = False                      #码装标识
        self.ma_D = 0                           #码装分母
        self.ma_U = 0                           #码装分子
        self.ma_N = 0                           #码装N
        self.is_E_and_H = False                 #钩车眼标识
        self.eIsSpecial = False                 #眼单特殊（BGLY/LY/F1/布筒）
        self.hIsSpecial = False                 #钩单特殊（BGLY/LY/F1/布筒）
        self.isNeedAdd = False                  #加数标识
        self.additions = 0                      #加数
        self.productB = 0                       #B数
        self.productP = 0                       #排数
        self.clothTubeEye = 0                   #眼布筒尺寸MM
        self.clothTubeHook = 0                  #眼布筒尺寸MM
        self.productSize = 0                    #产品尺寸
        self.productUnit = ''                   #产品单位
        self.colorRamk = ''                     #颜色备注
        self.productRamk = ''                   #生产工艺备注
        self.productECutType = ''               #眼切发
        self.productHCutType = ''               #钩切发
        self.productHookStr = ''                #钩类型
        self.productEyeStr = ''                 #眼类型
        self.is_HookSp = False                  #钩特殊（不锈钢或者电镀的）
        self.is_EyeSp = False                   #眼特殊（不锈钢电镀的）
        self.productHookKG = ''                 #钩KG
        self.productEyeKG = ''                  #眼KG
        self.productIsHookPressureWord = False  #钩压字标识
        self.productIsEyePressureWord = False   #眼压字标识
        self.productIsEyeLooseCut = False       #眼散口切标识
        self.arrEye = ''                        #眼工序
        self.arrHook = ''                       #钩工序
        self.areOther = ''                      #其它工序
        self.html_model_type = 0                #背钩类型（专利背钩、反包布、钩车眼、普通背钩）
#----------------5.HM Custom method----------------#
def clearNullStr(textValue):
    result = textValue.replace('\n', '')
    result = result.replace(' ', '')
    return result

# 获取product生产单位字段类型
# 背钩部 CHE
# textValue：生产单字符串集合
# ##销售订单
def bgGetProductDp(textValue):
    result = ''
    try:
        product_matchObj = re.search(r'生产单位:.*', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            productSN = productSN.replace('生产单位:', '')
            result = productSN
    except :
        return ''
    else:
        return result

# 获取生产单编号
# EX:CN-1905-0222
# textValue：生产单字符串集合
# ##
def bgGetProductInvoicesNum(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'[^主]生产单编号:.*', textValue, re.M|re.I)
        if product_matchObj:
            productIN = product_matchObj.group()
            productIN = clearNullStr(productIN)
            productIN = productIN.replace('生产单编号:', '')
            result = productIN
    except :
        return ''
    else:
        return result

# 获取产品编号/产品颜色号
# EX:J-HE-16621-EH02160
# EX:EH02160
# textValue：生产单字符串集合
# ##
def bgGetProductNumber(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'产品编号:.*', textValue, re.M|re.I)
        if product_matchObj:
            productNum = product_matchObj.group()
            productNum = clearNullStr(productNum)
            productNum = productNum.replace('产品编号:', '')
            hmColorArr = productNum.split('-')
            if (hmColorArr):
                hmPd.productColorNum = hmColorArr[-1]
            result = productNum
    except :
        return ''
    else:
        return result

# 获取产品类型/排数/B数/尺寸
# EX：IS/HHEE-3P-2B 3/4" 51X57MM
# EX：3
# EX：2
# EX：57
# textValue：生产单字符串集合
#修复换行问题python3 -m py_compile hmCreatePdf0701.py
# ##
def bgGetProductSpecification(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'产品名称:([\s\S]*)批次:', textValue, re.M|re.I)
        if product_matchObj:
            productSF = product_matchObj.group()
            productSF = clearNullStr(productSF)
            productSF = productSF.replace('产品名称', '')
            productSF = productSF.replace('批次', '')
            productSF = productSF.replace(':', '')
            result = productSF

        bStr = productSF
        bStr = bStr.replace('BG', '')
        bStr = bStr.replace('CB', '')
        bStr = bStr.replace('AB', '')

        matchObj = re.search( r'\d+[Bb]', bStr, re.M|re.I)
        if matchObj:
            bCount = matchObj.group()
            bCount = bCount.replace('B', '')
            bCount = bCount.replace('b', '')
            hmPd.productB = float(bCount)

        matchObj = re.search( r'\d+[Pp]', bStr, re.M|re.I)
        if matchObj:
            pCount = matchObj.group()
            pCount = pCount.replace('P', '')
            pCount = pCount.replace('p', '')
            hmPd.productP = float(pCount)

        matchObj = re.search( r'[Xx]\d+', bStr, re.M|re.I)
        if matchObj:
            sizeCount = matchObj.group()
            sizeCount = sizeCount.replace('X', '')
            sizeCount = sizeCount.replace('x', '')
            hmPd.productSize = float(sizeCount)

    except :
        return ''
    else:
        return result

# 获取销售单编号
# EX：SO-1905-0149
# textValue：生产单字符串集合
# ##销售订单
def bgGetProductSealNum(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'来源单据:.*', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            productSN = productSN.replace('来源单据', '')
            productSN = productSN.replace(':', '')
            result = productSN
    except :
        return ''
    else:
        return result

# 获取订单数量
# EX：SO-1905-0149
# textValue：825 SET
# ##销售订单
def getHmProductCount(textValue,hmPd):
    try:
        result = 0
        matchObj = re.search( r'订单数量:.*', textValue, re.M|re.I)
        if matchObj:
            productCount = matchObj.group()
            productCount = productCount.replace('\n', '')
            productCount = productCount.replace(',', '')
            hmArr = productCount.split(' ')
            if (len(hmArr) == 3):
                hmPd.productUnit = hmArr[-1]
                productNum = hmArr[-2]
                result = float(productNum)
        else:
            return 0
    except :
        return 0
    else:
        return result

# 获取客人号
# HMC1444
# textValue：生产单字符串集合
# ##
def bgGetProductGuest(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'客户编号: .*', textValue, re.M|re.I)
        if product_matchObj:
            productGuest = product_matchObj.group()
            productGuest = clearNullStr(productGuest)
            productGuest = productGuest.replace('客户编号', '')
            productGuest = productGuest.replace(':', '')
            result = productGuest
    except :
        return ''
    else:
        return result

# 获取product批次字段类型
# A、B
# textValue：生产单字符串集合
# ##销售订单
def bgGetProductBatch(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'批次:.*', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            productSN = productSN.replace('批次:', '')
            result = productSN
    except :
        return ''
    else:
        return result

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
        product_matchObj = re.finditer( r'布筒((:)|(：))?\s?(\d*)MM', textValue)
        if product_matchObj:

            tubeArr = []
            for match in product_matchObj:
                item = match.group()
                item_matchObj = re.search( r'(\d\d\d|\d\d|\d)', item, re.M|re.I)

                if (item_matchObj):  
                    tubeItem = item_matchObj.group()
                    tubeArr.append(tubeItem)

            #判断钩车眼数量，如果是2的话有钩有眼；
            if(len(tubeArr) == 2 ):
                hmClothTubeEye   =  tubeArr[0]
                hmClothTubeHook  =  tubeArr[1]
            #如果是1的话就有可能是钩车眼，或者单独钩，单独眼
            elif(len(tubeArr) == 1 ):
                he_matchObj = re.search( r'(钩眼车)|(钩加眼车)|(钩与眼车)', textValue, re.M|re.I)
                h_matchObj = re.search( r'(-H-)', productObj.productNum, re.M|re.I)
                #e_matchObj = re.search( r'(-E-)', productObj.productNum, re.M|re.I)
                if (he_matchObj):  
                    hmClothTubeEye   =  tubeArr[0]
                elif (h_matchObj):  
                    hmClothTubeHook   =  tubeArr[0]
                else:  
                    hmClothTubeEye   =  tubeArr[0]
            else:
                hmClothTubeEye   =  0
                hmClothTubeHook  =  0

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

            elif(len(cutArr) == 1):
                if(productObj.clothTubeHook == 0):
                    productECutType  =  cutArr[0]
                    productHCutType  =  ''
                else:
                    productECutType  =  ''
                    productHCutType  =  cutArr[0]
            else:
                productECutType = ''
                productHCutType = ''

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
def bgGetProductDetilRamk(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'详细说明：([\s\S]*).。\n', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            hmArr = productSN.split('.。')
            productSN = hmArr[0]
            productSN = productSN.replace('详细说明：', '')
            productSN = productSN.replace(' ', '')
            bgGetProductClothTube(productSN,hmPd)
            bgGetProductCutType(productSN,hmPd)
            bgGetEye(productSN,hmPd)
            bgGetHook(productSN,hmPd)
            colorRamk = hmArr[1]
            colorRamk = colorRamk.replace('颜色备注:', '')
            colorRamk = clearNullStr(colorRamk)
            hmPd.colorRamk = colorRamk
            productRamk = hmArr[2]
            productRamk = productRamk.replace('产品补充说明:', '')
            productRamk = clearNullStr(productRamk)
            hmPd.productRamkadd = productRamk

            result = productSN
    except :
        return ''
    else:
        return result

# 文档md5 用于不可描述的生产单防止重复匹配
# 中文描述
# fe65976af809170084acebb8d6af1fdd
# textValue：生产单字符串集合
#'PI-1906-0148J-HE-15294-KS02066A'
# ##
def bgGetPageMd5(hmPd):
    textValue = str(hmPd.productCasNum)+str(hmPd.productNum)+str(hmPd.productBatch)
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

# 获取product钩公斤数
# textValue：生产单字符串集合
# ##销售订单
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
# 特殊情况判断
# /钩车眼/压字/印字/BGLY/LY/F1/码装/需要愈多的客人加数
# 有是否有钩眼单
# 1,2,2,2,2,2,2
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

        #单价类型等
        bgly_matchObj = re.search( r'(BGLY)', hmPdObject.productRamk, re.M|re.I)
        he_matchObj = re.search( r'(钩眼车)|(钩加眼车)|(钩与眼车)', hmPdObject.productRamk, re.M|re.I)
        fbb_matchObj = re.search( r'(切货后)', hmPdObject.productRamk, re.M|re.I)

        if (bgly_matchObj):
            hmPdObject.html_model_type = 1
            hmPdObject.arrEye = BGLYEye
            hmPdObject.arrHook = BGLYHook
            hmPdObject.areOther = BGLYOther

        elif(he_matchObj):
            hmPdObject.html_model_type = 2
            hmPdObject.arrEye = HEEye
            hmPdObject.arrHook = HEHook
            hmPdObject.areOther = HEOther

        elif(fbb_matchObj):
            hmPdObject.html_model_type = 3
            hmPdObject.arrEye = FBEye
            hmPdObject.arrHook = FBHook
            hmPdObject.areOther = FBOther

        else:
            hmPdObject.html_model_type = 4
            hmPdObject.arrEye = PUBLICEye
            hmPdObject.arrHook = PUBLICHook
            hmPdObject.areOther = PUBLICOther

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

#创建二维码图片并保存
#codeStr        用于生成二维码的字符串
def hmCreateQRImage(codeStr):
    result = ''
    try:
        qr = qrcode.QRCode(
            version = 1,
            error_correction = qrcode.constants.ERROR_CORRECT_L,
            box_size = 2.5,
            border = 1,
        )
        hmQrCodeText = code_Url_Str + codeStr
        qr.add_data(hmQrCodeText)
        qr.make(fit = True)
        imgPath = upload_path + codeStr + '.png'
        img = qr.make_image()
        img.save(imgPath)
        img.close()
        qr.clear()
        result = imgPath
        del img
        gc.collect()
    except IOError:
        return ''
    else:
        return result

def hmCreateQRCode(hmPdObject,hookEyeType,payStr):
    codeStr = hmPdObject.hm_pd_uuid
    matchPill = 0
    matchYard = 0
    imgPath = ''

    if (hookEyeType == 'E'):
        matchPill = hmPdObject.productEPill
        matchYard = hmPdObject.productEYard
        imgPath = hmCreateQRImage(codeStr)
    else:
        matchPill = hmPdObject.productHPill
        matchYard = hmPdObject.productHYard
        imgPath = upload_path + codeStr + '.png'
        if(os.path.exists(imgPath)==False):
            imgPath = hmCreateQRImage(codeStr)

    #把二维码图片、HE标题文字写入PDF
    titleX = 310
    titleY = 702
    titleXX = 360
    titleYY = 692
    titleAX = 150
    titleAY = 533
    titleFontSize = 18
    pdfPath = upload_path + str(uuid.uuid1()) + '.pdf'
    c = canvas.Canvas(pdfPath)
    c.drawImage(imgPath, 490, 30)
    c.setFillColorRGB(0,0,0)
    c.setFont(psfontname=hmFont,size=titleFontSize)
    if (hookEyeType == 'E'):
        if hmPdObject.is_E_and_H:
            c.drawString(titleX,titleY,"（E+H）")
        else:
            c.drawString(titleX,titleY,"（E）")
    else:
        c.drawString(titleX,titleY,"（H）")

    c.setFont(psfontname=hmFont,size=50)
    if (hookEyeType == 'E'):
        if hmPdObject.is_E_and_H:
            c.drawString(titleX+80,titleYY,"眼钩")
        else:
            c.drawString(titleXX,titleYY,"眼")
        #眼压字
        c.setFont(psfontname=hmFont,size=80)
        if (hmPdObject.productIsEyePressureWord):
            c.drawString(200,100,"压字")
        #眼散口切
        if (hmPdObject.productIsEyeLooseCut):
            c.drawString(200,100,"散口")
    else:
        c.drawString(titleXX,titleYY,"钩")
        #钩压字
        c.setFont(psfontname=hmFont,size=80)
        if (hmPdObject.productIsHookPressureWord):
            c.drawString(200,100,"压字")

    #粒数码数
    c.setStrokeColorRGB(0, 1, 0)
    c.setFillAlpha(0.4)
    c.rect(25,705,90,13,0,1)
    c.rect(25,688,90,13,0,1)
    c.setFillAlpha(1)
    
    c.setFont(psfontname=hmFont,size=18)
    if(hmPdObject.is_MA == False):
        c.drawString(30,705,str(matchYard) + " Y")
    else:
        c.drawString(30,705,str(int(hmPdObject.productCount)) + "+"+str(matchYard) + " Y")
    c.drawString(30,688,str(matchPill) + " PCS")
    if hmPdObject.is_MA == False:
        c.setFillColorRGB(0.5,0.5,0.5)
        c.setFont(psfontname=hmFont,size=10)
        c.drawString(30,676,str(int(matchPill / hmPdObject.productB)) + " 付")
    
    #以下客人需要多加数量付
    if hmPdObject.isNeedAdd:
        csum = hmPdObject.additions + hmPdObject.productCount
        cText = "+"+str(hmPdObject.additions)+"="+str(int(csum))
        c.setFillColorRGB(0,0,0)
        c.setFont(psfontname=hmFont,size=20)
        c.drawString(titleAX,titleAY,cText)

    #钩眼重量
    c.setFillColorRGB(0.5,0.5,0.5)
    c.setFont(psfontname=hmFont,size=8)
    c.drawString(490,6,hmPdObject.productEyeStr+": "+str(hmPdObject.productEyeKG) + " kg")
    c.drawString(490,18,hmPdObject.productHookStr+": "+str(hmPdObject.productHookKG) + " kg")
    
    #工资
    c.setFont(psfontname=hmFont,size=10)
    c.drawString(25,17,payStr )
    
    #间接人员(差码装的打印)
    c.setFont(psfontname=hmFont,size=8)
    XRoo = 830
    tTO = 8
    ppsumPrice = 0
    for item in OTHERPeople:
        XRoo = XRoo - tTO
        cP = OTHERPeople[item]
        aP = Decimal(str(cP))
        bgly_matchObj = re.search( r'(备料)|(衣车)|(跟单)|(文员)', item, re.M|re.I)
        pP = Decimal(str(int(hmPdObject.productCount * hmPdObject.productB))) if (bgly_matchObj) else Decimal(str(int(hmPdObject.productCount)))
        tP =  aP * pP * 2
        c.drawString(5+430,XRoo,item)
        c.drawString(80+430,XRoo,str(cP))
        c.drawString(120+430,XRoo,str(tP))
        ppsumPrice += tP

    XRoo = XRoo - tTO
    c.drawString(120+430,XRoo,str(ppsumPrice))

    if(SpecialGuestsFbb.__contains__(hmPdObject.productGuest) == False):
        if(hmPdObject.arrHook.__contains__(53)):
            hmPdObject.arrHook.pop(53)
        if(hmPdObject.arrHook.__contains__(55)):
            hmPdObject.arrHook.pop(55)
        if(hmPdObject.arrHook.__contains__(51)):
            hmPdObject.arrHook.pop(51)
            
    if(hmPdObject.arrHook.__contains__(6) and hmPdObject.html_model_type == 3 and SpecialGuestsFbb.__contains__(hmPdObject.productGuest)):
            hmPdObject.arrHook.pop(6)

    if(SpecialGuestsBag.__contains__(hmPdObject.productGuest) == False):
        if(hmPdObject.areOther.__contains__(81)):
            hmPdObject.areOther.pop(81)
        if(hmPdObject.areOther.__contains__(82)):
            hmPdObject.areOther.pop(82)

    if(hmPdObject.arrEye.__contains__(4)):
        hmPdObject.arrEye.pop(4)
    if(hmPdObject.arrHook.__contains__(7)):
        hmPdObject.arrHook.pop(7)
    
    aE = hmPdObject.arrEye
    aH = hmPdObject.arrHook
    aO = hmPdObject.areOther

    aE.update(aH)
    aE.update(aO)
    XAoo = 320
    tAO = 8
    allsumPrice = 0
    allsumPrice  += ppsumPrice
    for item in aE:
        XAoo = XAoo - tAO
        singlePrice =  hmGetPrice(hmPdObject,aE,item)
        sumPrice = Decimal(str(int(hmPdObject.productCount)))* Decimal(str(singlePrice))
        c.drawString(5+430,XAoo,aE[item])
        c.drawString(80+430,XAoo,str(singlePrice))
        c.drawString(120+430,XAoo,str(sumPrice))
        #if(item != 4 and item != 7):
        allsumPrice += sumPrice
    XAoo = XAoo - tAO
    c.drawString(80+430,XAoo,'合计：')
    c.drawString(120+430,XAoo,str(allsumPrice))
    c.save()
    del c
    gc.collect()
    
    #创建成功PDF成功移除图片
    if (hookEyeType == 'H'):
        os.remove(imgPath)
    if (hmPdObject.clothTubeHook == 0):
        os.remove(imgPath)
    return pdfPath

# 计算钩眼粒数
# （粒数）【T】= 订单数量 * 预多百分数 * B数
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

##
# 计算布料码数
# 【M】= T粒数 / （914 / 尺寸 * B数）
# hookEyePillCount   粒数
# productSize        尺寸
# productB           B数
# ##
def countHmYard(hookEyePillCount,hmPd):
    productSize = hmPd.productSize
    productB = hmPd.productB
    try:
        result = 0
        if(hmPd.is_MA == False):
            intValue = ( 914 / productSize * productB )
            intValueM = int(intValue)
            yardCount = hookEyePillCount / intValueM
        else:
            yardCount = (hookEyePillCount / hmPd.ma_N - hmPd.productCount) if(hmPd.ma_N != 0) else 0
        result = yardCount
    except :
        return 0
    else:
        return round(result)


def updateOutDate(hmPd,pdfTxt):
    result = ''
    try:
        product_matchObj = re.search( r'交货日期:.*', pdfTxt, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = productSN.replace('交货日期: ', '')
            hmPd.outDate = productSN
            result = productSN

            my_con = connect(
                host='127.0.0.1', 
                port=3306, 
                db='hmpd_erp', 
                user='root', 
                password='root', 
                charset='utf8')

            my_cousor = my_con.cursor()
            sql_update = "UPDATE hm_workshop_datas SET pd_nb_out_date = '%s' WHERE pd_uuid = '%s'" % (hmPd.outDate,hmPd.hm_pd_uuid)
            my_cousor.execute(sql_update)
            my_con.commit() 
            my_cousor.close()
            my_con.close() 
    except :
        return ''
    else:
        return result
    

def parse(fileName):
    text_path = upload_path + fileName + ".pdf"
    hmPdfSaveName = ""
    fileOpen = open(text_path,'rb')
    doc = PDFDocument()
    parser = PDFParser(fileOpen)
    parser.set_document(doc)
    doc.set_parser(parser)
    doc.initialize()
    #原文件
    hmPdfReaderEye = PyPDF2.PdfFileReader(fileOpen)
    hmPdfReaderHook = PyPDF2.PdfFileReader(fileOpen)
    #待写入数据文件
    hmPdfWriter = PyPDF2.PdfFileWriter()
    #检测文档是否提供txt转换，不提供就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        #创建PDF，资源管理器，来共享资源
        rsrcmgr = PDFResourceManager()
        #创建一个PDF设备对象
        device = PDFPageAggregator(rsrcmgr,laparams=LAParams())
        #创建一个PDF解释其对象
        interpreter = PDFPageInterpreter(rsrcmgr,device)
        openFileArr = []
        allPages = doc.get_pages()
        for page in allPages:
            interpreter.process_page(page)
            layout = device.get_result()

            textValueArr = []
            for x in layout:
                if(isinstance(x,LTTextBoxHorizontal)):
                    textValueArr.append(x.get_text())
            pdfTxt = ''.join(textValueArr)
            textValueArr.clear()
            
            dpText = bgGetProductDp(pdfTxt)
            if dpText != 'CHE':
                #print('不是背钩部生产单')
                continue

            #init
            hmPd = HmProduct()
            #生产车间
            hmPd.productDp = dpText
            #生产单编号
            hmPd.productCasNum = bgGetProductInvoicesNum(pdfTxt)
            #产品编号
            hmPd.productNum = bgGetProductNumber(pdfTxt,hmPd)
            #钩眼规格/排数/B数/尺寸
            hmPd.productSf = bgGetProductSpecification(pdfTxt,hmPd)
            #销售单号
            hmPd.productSealNum = bgGetProductSealNum(pdfTxt)
            #订单数量/单位
            hmPd.productCount = getHmProductCount(pdfTxt,hmPd)
            #客人号
            hmPd.productGuest  = bgGetProductGuest(pdfTxt)
            #产品批次
            hmPd.productBatch = bgGetProductBatch(pdfTxt)
            #产品中文描述/颜色备注/产品补充说明/钩眼布筒/钩眼切法
            hmPd.productRamk = bgGetProductDetilRamk(pdfTxt,hmPd)
            #生成生产单uuid
            hmPd.hm_pd_uuid = bgGetPageMd5(hmPd)
            #处理特殊订单
            getHmHookEyeIsSpecial(hmPd)
            #产品排数粒数
            hmPd.productEPill = countHmPillx('E',hmPd)
            hmPd.productHPill = countHmPillx('H',hmPd)
            hmPd.productEYard = countHmYard(hmPd.productEPill,hmPd)
            hmPd.productHYard = countHmYard(hmPd.productHPill,hmPd)
            #产品钩眼公斤数
            hmPd.productEyeKG = bgGetHookEyeKg('E',hmPd)
            hmPd.productHookKG = bgGetHookEyeKg('H',hmPd)
            #根据总表更新生产单日期
            updateOutDate(hmPd,pdfTxt)
            
            #----------------1、生成文件_STAR----------------#
            payStr = ''
            layoutPageId = layout.pageid - 1 

            #生成【眼】单
            if hmPd.clothTubeEye > 0:
                eyeNewPage = hmPdfReaderEye.getPage(layoutPageId) 
                ePatch = hmCreateQRCode(hmPd,'E',payStr)
                eMarkFile = open(ePatch,'rb')
                pdfECodePage = PyPDF2.PdfFileReader(eMarkFile)
                eyeNewPage.mergePage(pdfECodePage.getPage(0))
                hmPdfWriter.addPage(eyeNewPage)
                openFileArr.append(eMarkFile)
                del eyeNewPage
                del pdfECodePage
                gc.collect()
            #生成【钩】单
            if hmPd.clothTubeHook > 0:
                hookNewPage = hmPdfReaderHook.getPage(layoutPageId) 
                hPatch = hmCreateQRCode(hmPd,'H',payStr)
                hMarkFile = open(hPatch,'rb')
                pdfHCodePage = PyPDF2.PdfFileReader(hMarkFile)
                hookNewPage.mergePage(pdfHCodePage.getPage(0))
                hmPdfWriter.addPage(hookNewPage)
                openFileArr.append(hMarkFile)
                del hookNewPage
                del pdfHCodePage
                gc.collect()
            #用销售单号做文件名
            if hmPdfSaveName == "":
                hmPdfSaveName = hmPd.productSealNum
            #----------------1、生成文件_END----------------#

        #完结时关闭文件和保存文件
        #----------------生成文件时关闭----------------#
        nowTime = datetime.datetime.now()
        nowTimeStr = nowTime.strftime("%Y%m%d%H%M%S_s")
        hmPdfSaveName = nowTimeStr +"_"+ hmPdfSaveName+ ".pdf"
        hmPdfSavePath = download_path + hmPdfSaveName
        resultPdfFile = open(hmPdfSavePath,'wb')
        hmPdfWriter.write(resultPdfFile)
        for closeItem in openFileArr :
            closeItem.close()
            os.remove(closeItem.name)
        openFileArr.clear()
        resultPdfFile.close()

        fileOpen.close()
        return hmPdfSaveName

def hmGetPrice(hmPd,arrData,workType):
    productHMCount = 0
    #某几个工序需要用粒数来
    yc_matchObj = re.search( r'(衣车)', arrData[workType], re.M|re.I)
    productHMCount = (hmPd.productCount * hmPd.productB) if yc_matchObj else hmPd.productCount
    pKey = (str(int(hmPd.productP))+'U') if (productHMCount > 5000) else (str(int(hmPd.productP))+'D')

    productPrice = 0.00
    if (workType == 2)or(workType == 3)or(workType == 9)or(workType == 5)or(workType == 6)or(workType == 8)or(workType == 81)or(workType == 82)or(workType == 10)or(workType == 61)or(workType == 62)or(workType == 63)or(workType == 64)or(workType == 13)or(workType == 14)or(workType == 11)or(workType == 12)or(workType == 15)or(workType == 16) or(workType == 17)or(workType == 18)or(workType == 19)or(workType == 20)or(workType == 23)or(workType == 24)or(workType == 21)or(workType == 22)or(workType == 25)or(workType == 26)or(workType == 51)or(workType == 52)or(workType == 53)or(workType == 54)or(workType == 55)or(workType == 56):
        
        bgly_matchObj = re.search( r'(BGLY)', hmPd.productRamk, re.M|re.I)
        fbb_matchObj = re.search( r'(切货后)', hmPd.productRamk, re.M|re.I)
        if (bgly_matchObj and (workType == 2 or workType == 3)):
            pCarUnitPrice = ast.literal_eval(config.get('OptionsPRICE', str(workType)+'BGLY'))
        elif (fbb_matchObj and (workType == 2 or workType == 3)):
            pCarUnitPrice = ast.literal_eval(config.get('OptionsPRICE', str(workType)+'FBB'))
        else :
            pCarUnitPrice = ast.literal_eval(config.get('OptionsPRICE', str(workType)))
        
        if(pCarUnitPrice.__contains__(pKey)):
            if(pCarUnitPrice[pKey].__contains__(int(hmPd.productB))):
                productPrice = pCarUnitPrice[pKey][int(hmPd.productB)]
            else:
                productPrice = Decimal(str(int(hmPd.productB))) * Decimal(str(pCarUnitPrice[pKey][1]))

        else:
            pKey2 = (str(int(hmPd.productP))+'D') if (productHMCount > 5000) else (str(int(hmPd.productP))+'U')
            if(pCarUnitPrice.__contains__(pKey2)):
                if(pCarUnitPrice[pKey2].__contains__(int(hmPd.productB))) :
                    productPrice = Decimal(str(pCarUnitPrice[pKey2][int(hmPd.productB)]))
                else:
                    productPrice = Decimal(str(int(hmPd.productB))) * Decimal(str(pCarUnitPrice[pKey2][1]))
            else:
                return 0
    
    elif (workType == 4):
        pCarUnitPriceH =ast.literal_eval(config.get('OptionsPRICE', str(2)))
        pCarUnitPriceE =ast.literal_eval(config.get('OptionsPRICE', str(3)))
                
        if((pCarUnitPriceH.__contains__(pKey)) and (pCarUnitPriceE.__contains__(pKey))):
            productPrice1=0.00
            productPrice2=0.00

            if(pCarUnitPriceH[pKey].__contains__(int(hmPd.productB))):
                productPrice1= Decimal(str(pCarUnitPriceH[pKey][int(hmPd.productB)]))
            else:
                productPrice1 = Decimal(str(int(hmPd.productB))) * Decimal(str(pCarUnitPriceH[pKey][1]))
            
            if(pCarUnitPriceE[pKey].__contains__(int(hmPd.productB))):
                productPrice2= pCarUnitPriceE[pKey][int(hmPd.productB)]
            else:
                productPrice2 = Decimal(str(int(hmPd.productB))) * Decimal(str(pCarUnitPriceE[pKey][1]))
            
            productPrice = productPrice1 + productPrice2

        else :
            return 0

    elif (workType == 7):
        pCarUnitPriceH =ast.literal_eval(config.get('OptionsPRICE', str(6)))
        pCarUnitPriceE =ast.literal_eval(config.get('OptionsPRICE', str(5)))
        productPrice1=0.00
        productPrice2=0.00

        if((pCarUnitPriceH.__contains__(pKey)) and (pCarUnitPriceE.__contains__(pKey))):
            productPrice1 = Decimal(str(pCarUnitPriceH[pKey][int(hmPd.productB)]))
            productPrice2 = Decimal(str(pCarUnitPriceE[pKey][int(hmPd.productB)]))
            productPrice = productPrice1 + productPrice2

        else :
            pKey2 = (str(int(hmPd.productP))+'D') if (productHMCount > 5000) else (str(int(hmPd.productP))+'U')
            if(pCarUnitPriceH.__contains__(pKey2) and pCarUnitPriceE.__contains__(pKey2)):
                productPrice1 = Decimal(str(pCarUnitPriceH[pKey2][int(hmPd.productB)]))
                productPrice2 = Decimal(str(pCarUnitPriceE[pKey2][int(hmPd.productB)]))
                productPrice = productPrice1 + productPrice2
            else:
                return 0
    else:
        return 0
    return productPrice

if __name__ == '__main__':
    shellValues = sys.argv[1]
    shellArr =  shellValues.split(":")
    isSaveInData = int(shellArr[1])
    fileName = shellArr[0]
    autoResult = parse(fileName)
    print(autoResult)
