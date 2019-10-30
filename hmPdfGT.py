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
code_Url_Str        = config.get('QrCode_Url', 'code_Url_Str_gt')

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
    textValue = str(hmPd.productCasNum)+str(hmPd.productNum)+str(hmPd.productBatch)+str(int(hmPd.productCount))
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

    imgPath = hmCreateQRImage(codeStr)
    titleFontSize = 18
    pdfPath = upload_path + str(uuid.uuid1()) + '.pdf'
    c = canvas.Canvas(pdfPath)
    c.drawImage(imgPath, 490, 30)
    c.setFillColorRGB(0,0,0)
    #c.setFont(psfontname=hmFont,size=titleFontSize)

    c.save()
    del c

    os.remove(imgPath)
    return pdfPath




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
    hmPdfReaderPDF = PyPDF2.PdfFileReader(fileOpen)

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

            gtList = ['CWF','CWFS','CWFH','CWFN','CSB']
            
            if (gtList.__contains__(dpText) == False):
                continue



            #init
            hmPd = HmProduct()
            #生产车间
            hmPd.productDp = dpText
            #生产单编号
            hmPd.productCasNum = bgGetProductInvoicesNum(pdfTxt)
            #产品编号
            hmPd.productNum = bgGetProductNumber(pdfTxt,hmPd)
            #规格
            hmPd.productSf = bgGetProductSpecification(pdfTxt,hmPd)
            #销售单号
            hmPd.productSealNum = bgGetProductSealNum(pdfTxt)
            #订单数量/单位
            hmPd.productCount = getHmProductCount(pdfTxt,hmPd)
            #客人号
            hmPd.productGuest  = bgGetProductGuest(pdfTxt)
            #产品批次
            hmPd.productBatch = bgGetProductBatch(pdfTxt)

            #产品中文描述
            hmPd.productRamk = bgGetProductDetilRamk(pdfTxt,hmPd)
            #生成生产单uuid
            hmPd.hm_pd_uuid = bgGetPageMd5(hmPd)

            #根据总表更新生产单日期
            #updateOutDate(hmPd,pdfTxt)
            
            #----------------1、生成文件_STAR----------------#
            payStr = ''
            layoutPageId = layout.pageid - 1 

            #生成【眼】单
            if hmPd.hm_pd_uuid :
                newPage = hmPdfReaderPDF.getPage(layoutPageId) 
                ePatch = hmCreateQRCode(hmPd,'E',payStr)
                eMarkFile = open(ePatch,'rb')
                pdfECodePage = PyPDF2.PdfFileReader(eMarkFile)
                newPage.mergePage(pdfECodePage.getPage(0))
                hmPdfWriter.addPage(newPage)
                openFileArr.append(eMarkFile)
                del newPage
                del pdfECodePage
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



if __name__ == '__main__':
    shellValues = sys.argv[1]
    shellArr =  shellValues.split(",")
    isSaveInData = int(shellArr[1])
    fileName = shellArr[0]
    autoResult = parse(fileName)
    print(autoResult)
