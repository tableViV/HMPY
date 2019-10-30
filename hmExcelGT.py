#!/bin/python
import re,os,sys,ast,xlrd,hashlib,os.path,datetime,importlib,configparser
from pymysql import connect
from datetime import datetime
from xlrd import xldate_as_tuple

#钢托生产总表插入
hmInsertBgProduct_Str= "INSERT IGNORE INTO hm_workshopgt_datas(`pd_uuid`,`pd_type`,`wp_units`,`pd_batch`, `pd_create_date`, `seal_num`, `pd_num`, `pd_count`, `pd_unit`, `pd_guest`, `pd_code`, `pd_guest_code`,`pd_ramk`,`pd_cn_ramk`, `pd_out_date`,`code_num`,`reay_pcs`,`virtual_e_pcs`,`virtual_h_pcs`,`virtual_e_str`,`virtual_h_str`,`created_at`,`hm_product_state`) \
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW(),0)"

proDir = os.path.split(os.path.realpath(__file__))[0]
configPath = os.path.join(proDir, "HM_config_DEV.ini")
config = configparser.ConfigParser()
config.read(configPath,encoding='UTF-8')
global upload_path      #文件上传路径
upload_path = config.get('File_Path', 'upload_path')


class HmSqlDatas:
    def __init__(self):
        self.masterList = []        #主表数据
        self.sealNumList = []       #销售单号列表
        self.uuidList = []          #UUID列表

class HmProduct:
    '产品基类'
    pdCount = 0
    def __init__(self):
        HmProduct.pdCount += 1
        self.hm_pd_uuid = ''    

        self.is_MA = False      #码装
        self.productHMWP = ''   #JG 胶骨 GTDF 钢托点粉 GG 钢骨
        self.productHMType = '产品分类' 
        #胶骨:POM不过针胶骨，PPB普通胶骨，FNB对折胶骨 
        #钢骨:SSB不锈钢,SMB 不锈钢
        #钢托:(SS不锈钢)|(NC白线)|(HC高碳钢)
        self.productHMSize = ''  #尺寸
        self.productHMPdType = ''#型号代码
        self.productUnit = ''
        self.productColorNum = ''#颜色号
        self.productRamk = ''
        self.productGuest = ''#客人号


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

def clearNullStr(textValue):
    result = textValue.replace('\n', '')
    result = result.replace(' ', '')
    return result

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


#获取胶骨产品的型号分类和尺寸数字
def getProductJGSizeCount(hmPd):
    productHMSize = ''
    productHMPdType  = ''
    try:
        #型号数字分类
        count_matchObj = re.search( r'(\d\d\d\d)-', hmPd.productHMcode, re.M|re.I)
        if count_matchObj:
            productHMPdType = count_matchObj.group()
            productHMPdType = productHMPdType.replace('-', '')
        
        #尺寸数字
        size_matchObj = re.search( r'-(\d*)MM', hmPd.productHMcode, re.M|re.I)
        if size_matchObj:
            productHMSize = size_matchObj.group()
            productHMSize = productHMSize.replace('-', '')

    except :
        productHMSize = "null"
        productHMPdType = "null"
    finally:
        hmPd.productHMSize = productHMSize
        hmPd.productHMPdType = productHMPdType


#获取胶骨产品的型号分类和尺寸数字
def getProductGGSizeCount(hmPd):
    productHMSize = ''
    productHMPdType  = ''
    try:
        #型号数字分类
        count_matchObj = re.search( r'(\d+)-|\d+-', hmPd.productHMcode, re.M|re.I)
        if count_matchObj:
            productHMPdType = count_matchObj.group()
            productHMPdType = productHMPdType.replace('-', '')
        else:
            productHMPdType = '-'

        all_matchObj = re.findall( r'[^()]+', hmPd.productHMcode, re.M|re.I)
        if len(all_matchObj) > 2:
            productHMPdType = str(productHMPdType) + str('-') + str(all_matchObj[1])
        else:
            productHMPdType = str(productHMPdType) + str('-06')#默认06

        #尺寸数字
        size_matchObj = re.search( r'-(\d*)MM', hmPd.productHMcode, re.M|re.I)
        if size_matchObj:
            productHMSize = size_matchObj.group()
            productHMSize = productHMSize.replace('-', '')

    except :
        productHMSize = "null"
        productHMPdType = "null"
    finally:
        hmPd.productHMSize = productHMSize
        hmPd.productHMPdType = productHMPdType

#获取钢托点粉产品的型号分类和尺寸数字
def getProductGTDFSizeCount(hmPd):
    productHMPdType  = ''
    try:
        #型号数字分类

        hmPd.productHMcnName = clearNullStr(hmPd.productHMcnName)
        count_matchObj = re.search( r'((\d+((\.)*)\d*)MM)|((\d+((\.)*)\d*)X(\d+((\.)*)\d*)MM)|((\d+((\.)*)\d*)X(\d+((\.)*)\d*))|((\d+((\.)*)\d*)x(\d+((\.)*)\d*)MM)|((\d+((\.)*)\d*)x(\d+((\.)*)\d*))|((\d+((\.)*)\d*)×(\d+((\.)*)\d*)MM)|((\d+((\.)*)\d*)×(\d+((\.)*)\d*))|((\d+((\.)*)\d*)X(\d+((\.)*)\d*)X(\d+((\.)*)\d*)MM)', hmPd.productHMcnName, re.M|re.I)
        if count_matchObj:
            productHMPdType = count_matchObj.group()
            productHMPdType = productHMPdType.replace('x', 'X')
            productHMPdType = productHMPdType.replace('×', 'X')
            productHMPdType = productHMPdType.replace('mm', 'MM')


    except :
        productHMPdType = "null"
    finally:
        hmPd.productHMPdType = productHMPdType
        



# 判断产品是否胶骨
def getProductCnNameIsJG(hmPd):
    result = False
    try:
        pd_matchObj = re.search( r'(POM)|(PPB)|(FNB)', hmPd.productHMcode, re.M|re.I)
        pd_flag = True if pd_matchObj else False
        if pd_flag:
            hmPd.productHMType = pd_matchObj.group()

        cn_matchObj = re.search( r'(切骨)|(胶骨)', hmPd.productHMcnName, re.M|re.I)
        cn_flag = True if cn_matchObj else False
        
        result = True if (pd_flag and cn_flag) else False
  
    except :
        result = False
    finally:
        if result :
            hmPd.productHMWP = 'JG'
        return result

# 判断产品是否钢骨
def getProductCnNameIsGG(hmPd):
    result = False
    try:
        pd_matchObj = re.search( r'(SMB)|(SSB)', hmPd.productHMcode, re.M|re.I)
        pd_flag = True if pd_matchObj else False
        if pd_flag:
            hmPd.productHMType = pd_matchObj.group()

        cn_matchObj = re.search( r'(魚鱗)|(鱼鳞)|(魚鳞)|(鱼鱗)', hmPd.productHMcnName, re.M|re.I)
        cn_flag = True if cn_matchObj else False
        
        result = True if (pd_flag and cn_flag) else False
  
    except :
        result = False
    finally:
        if result :
            hmPd.productHMWP = 'GG'
        return result

# 判断产品是否钢托点粉
def getProductCnNameIsGTDF(hmPd):
    result = False
    try:
        pd_matchObj = re.search( r'(SS)|(NC)|(HC)', hmPd.productHMcode, re.M|re.I)
        pd_flag = True if pd_matchObj else False
        if pd_flag:
            hmPd.productHMType = pd_matchObj.group()

        cn_matchObj = re.search( r'(点粉)|(點粉)', hmPd.productHMcnName, re.M|re.I)
        cn_flag = True if cn_matchObj else False
        
        result = True if (pd_flag and cn_flag) else False
  
    except :
        result = False
    finally:
        if result :
            hmPd.productHMWP = 'GTDF'
        return result







##
# 这里是转换Excel中的Data格式
# ##
def get_excel_data(fileName,hmSqlDatas):
    text_path = upload_path + fileName + ".xls"
    workbook = xlrd.open_workbook(text_path)
    sheet = workbook.sheets()[0]  # 读取第一个sheet
    nrows = sheet.nrows  # 行数
    first_row_values = sheet.row_values(0)  # 第一行数据

    for row_num in range(1, nrows):
        row_values = sheet.row_values(row_num)
        

        
        if row_values:
            #init
            hmPd = HmProduct()
            pd_uuid = ''
            pd_seal = ''
            masterList_obj = []

            #还要做部门验证
            hmPd.productHMcode = sheet.cell_value(row_num, 9)
            hmPd.productHMcnName = sheet.cell_value(row_num, 0)

            #胶骨产品
            if (getProductCnNameIsJG(hmPd)):
                getProductJGSizeCount(hmPd)
            #钢骨产品
            elif (getProductCnNameIsGG(hmPd)):
                getProductGGSizeCount(hmPd)
            elif (getProductCnNameIsGTDF(hmPd)):
                getProductGTDFSizeCount(hmPd)
            else :
                continue

            
            pd_seal = sheet.cell_value(row_num, 4)

            #开始组装UUID,用于构造两个表第一列的uuid
            pd_num = sheet.cell_value(row_num, 5)
            pd_code = sheet.cell_value(row_num, 9)
            pd_batch = sheet.cell_value(row_num, 2)
            pd_count = sheet.cell_value(row_num, 6)
            #生产单号+产品编号+批次+产品数量
            pd_uuid = str(pd_num) + str(pd_code) + str(pd_batch) + str(int(pd_count))
            pd_uuid = bgGetPageMd5(pd_uuid)

            masterList_obj.append(pd_uuid)

            hmPd.productUnit = sheet.cell_value(row_num, 7)
            hmPd.productGuest = sheet.cell_value(row_num, 8)
            hmPd.productSf = sheet.cell_value(row_num, 0)
            hmPd.productRamk = sheet.cell_value(row_num, 12)

            #开始组装两个表的其它字段
            for i in range(len(first_row_values)):
                ctype = sheet.cell(row_num, i).ctype
                cell = sheet.cell_value(row_num, i)
                if ctype == 2 and cell % 1 == 0.0:  # ctype为2且为浮点
                    cell = int(cell)  # 浮点转成整型
                    #cell = str(cell)  # 转成整型后再转成字符串，如果想要整型就去掉该行
                elif ctype == 3:
                    date = datetime(*xldate_as_tuple(cell, 0))
                    cell = date.strftime('%Y/%m/%d %H:%M:%S')
                elif ctype == 4:
                    cell = True if cell == 1 else False

                masterList_obj.append(cell)

                if i == 13:
                    masterList_obj.append(hmPd.productHMWP)
                    masterList_obj.append(hmPd.productHMType)
                    masterList_obj.append(hmPd.productHMSize)
                    masterList_obj.append(hmPd.productHMPdType)
                    masterList_obj.append('5')
                    masterList_obj.append('6')

            hmSqlDatas.masterList.append(masterList_obj)

            
            if(pd_seal not in hmSqlDatas.sealNumList):
                hmSqlDatas.sealNumList.append(pd_seal)
            hmSqlDatas.uuidList.append(pd_uuid)

    print(text_path)
    return hmSqlDatas

def parse(fileName):
    hmDatas = HmSqlDatas()
    
    hmSql = hmDataHandle()
    get_excel_data(fileName,hmDatas)
    if hmDatas.masterList:
        hmSql.hmInsertBgProduct_sql(hmInsertBgProduct_Str,hmDatas.masterList)
        

    #更新数据库中，包含Excel的数据全部设置为失效产品。
    if hmDatas.sealNumList:
        for item in hmDatas.sealNumList :
            hmSql6 = hmDataHandle()
            sql = "UPDATE hm_workshopgt_datas SET `is_change`=1 WHERE `seal_num` ='"+item+"'"
            hmSql6.hmInsertBgProduct_sqlOne(sql)

    #重新设置为有效产品。
    hmSql4 = hmDataHandle()
    if hmDatas.uuidList:
       sss = "INSERT INTO hm_workshopgt_datas (`pd_uuid`,`is_change`)VALUES(%s,0) ON DUPLICATE KEY UPDATE `is_change`=0"
       hmSql4.hmInsertBgProduct_sql(sss,hmDatas.uuidList)

if __name__ == '__main__':
    shellValues = sys.argv[1]
    shellArr =  shellValues.split(",")
    flag = int(shellArr[1])
    fileName = shellArr[0]
    res = parse(fileName)
