#!/usr/bin/env python
# -*- coding:utf8 -*-

import xlrd
from xlrd import cellname, cellnameabs, colname
import xlwt
import xlutils
import xlutils.copy

import os.path
from string import strip,lstrip,rstrip,join,split
import re
import logging  
import logging.handlers
import datetime
import ConfigParser

import sys
reload(sys) 
sys.setdefaultencoding('utf-8')

                
def setLogger(name='rebate', rootdir='.',level = logging.INFO):
    LOG_FILE = 'rebate.log'
    handler = logging.handlers.RotatingFileHandler(os.path.join( os.path.realpath(rootdir),LOG_FILE), maxBytes = 1024*1024, backupCount = 5)
    fmt = '%(asctime)s - %(filename)s:%(lineno)s - %(name)s - %(message)s'  
    formatter = logging.Formatter(fmt)
    handler.setFormatter(formatter)
    logger = logging.getLogger()
    logger.addHandler(handler)
    logger.setLevel(level)    
    return logger



quarters = (1,2,3,4)
seasons1  = {1:('JAN','FEB','MAR'),2:('APR','MAY','JUN'),3:('JUL','AUG','SEP'),4:('OCT','NOV','DEC')}
seasons2  = {1:(u'1月',u'2月',u'3月'),2:(u'4月',u'5月',u'6月'),3:(u'7月',u'8月',u'9月'),4:(u'10月',u'11月',u'12月')}
months0  = (u'JAN',u'FEB',u'MAR',u'APR',u'MAY',u'JUN',u'JUL',u'AUG',u'SEP',u'OCT',u'NOV',u'DEC')
months1   = (u'1月',u'2月',u'3月',u'4月',u'5月',u'6月',u'7月',u'8月',u'9月',u'10月',u'11月',u'12月')
months2   = dict(zip(months1,months0))
months1idx   = {u'1月':0,u'2月':1,u'3月':2,u'4月':0,u'5月':1,u'6月':2,u'7月':0,u'8月':1,u'9月':2,u'10月':0,u'11月':1,u'12月':2}

#orgfileext = 'CSV'
orgfileext = 'xls' #edited 20140403 
destfileext = 'xls'

#支持的币种
CCY1 = ('HKD','USD','JPY','MYR','CNY') # L:Total Qty Q:Net Income
#rebate文件中的填写位置
CCY2 = {'HKD':('D','E'),'USD':('I','J'),'JPY':('N','O'),'MYR':('U','V'),'CNY':('AB','AC')}

TimYear = '2014'
TimQuarter = 1


HKDm1 = {'A':('HHIF','HHIO','HSIF','HSIO'),
         'B':('MCHF','MHSIF','MHSIO')}
CCYm1 = {'HKD':(0,0.7,0.3),'USD':(1,0.7,0.3)}

season_Titles=(u'Sales',
            u'Client',
            u'Internet',
            u'Client Name',
            u'Ccy',
            u'Market',
            u'Product',
            u'Underlying',
            u'Day Qty',
            u'Day Brokerage',
            u'Overnight Qty',
            u'Overnight Brokerage',
            u'Total Qty',
            u'Total Brokerage',
            u'Fee Received',
            u'Gross Income',
            u'Fee Paid',
            u'Net Income')

season_c_Titles=(
        u'Day Qty',
        u'Day Brokerage',
        u'Overnight Qty',
        u'Overnight Brokerage',
        u'Total Qty',
        u'Total Brokerage',
        u'Fee Received',
        u'Gross Income',
        u'Fee Paid',
        u'Net Income')

TimStyles = {
    u'datetime': xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm:ss'),
    u'date': xlwt.easyxf(num_format_str='yyyy-mm-dd'),
    u'time': xlwt.easyxf(num_format_str='hh:mm:ss'),
    u'headerBIG0': xlwt.easyxf('alignment: horiz centre; font: name Times New Roman, color-index blue, bold on, height 400', num_format_str='#,##0.00'),
    u'header': xlwt.easyxf('borders: top medium, bottom medium, right medium;alignment: horiz centre; font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00'),
    u'headerBIG': xlwt.easyxf('borders: top medium, bottom medium, right medium;alignment: vert centre, horiz centre; font: name Times New Roman, color-index blue, bold on, height 400', num_format_str='#,##0.00'),
    u'id' : xlwt.easyxf('borders: top medium, bottom medium, right medium;alignment: vert centre, horiz centre; font: name Times New Roman, color-index green, bold on', num_format_str='#,##0.00'),
    u'sum' : xlwt.easyxf('font: bold on; borders: top medium', num_format_str='#,##0.00'),
    u'sum2' : xlwt.easyxf('font: bold on; borders: bottom medium', num_format_str='#,##0.00'),
    u'ccy' : xlwt.easyxf('font: bold on; borders: top medium'),
    u'ccy2' : xlwt.easyxf('font: bold on; borders: bottom medium'),
    u'summary': xlwt.easyxf('pattern: pattern solid, fore_colour light_green,back_colour gray25;borders: right medium;alignment: horiz right; font: name Times New Roman, color-index green, bold on', num_format_str='#,##0.00'),
    u'currency' : xlwt.easyxf('font: bold on; borders: top medium, left medium, bottom medium, right medium'), 
}

CCYRATE={}
SUMMARY={u'HKD':{
            u'市场总张数':[],
            u'总收入':[],
            u'港币净收入(减SP费用后)':[],
            u'港币公司净留存':[],
            u'港币返佣':[]         
        },
        u'USD':{
            u'市场总张数':[],
            u'折美元公司净收入':[],
            u'折美元公司净留存':[],
            u'折美元返佣':[]
        }}
        
TAB_SUMMARY_ADD = {}
TAB_SUMMARY = {}
#TAB_AE_ADD =  {}

IBRate = [0.7,0.3]
IBRateSpecial = {}
IBClientRateSpecial = {}

class ANCurrencySUMBASE(object):
    __c_Titles = season_c_Titles
    __ccy1 = CCY1
    __IDIDX = 0
    CCYRATE_ADDR = {}
    TOTAL_LEN = 0
    IBRate = IBRate
    IBRateSpecial = IBRateSpecial
    IBClientRateSpecial = IBClientRateSpecial
    def __init__(self,currency,begincol):
        self.MyID = ANCurrencySUMBASE.__IDIDX
        ANCurrencySUMBASE.__IDIDX += 1
        self._ccyRate = 1.0
        self.leftskip = 3
        self.__currency = currency
        if (type(begincol) is str):
            self.__colBegin = xlwt.Utils.col_by_name(begincol)
        else:
            self.__colBegin = begincol
        self.__colCHRs = [chr(i) for i in range(self.__colBegin,self.__colBegin+7)] 
        self.__currencyCols = dict(zip(self.colTitles,self.__colCHRs))
        key = u'USD/'+currency
        if (CCYRATE.has_key(key)):
            self.set_ccyRate(CCYRATE[key])
        ANCurrencySUMBASE.TOTAL_LEN += self.__len__() 
        
    def __len__(self):
        return len(self.colTitles)
    
    def set_ccyRate(self,rate):
        self._ccyRate = rate
    
    def set_IBRate(self,rate0,rate1):
        ANCurrencySUMBASE.IBRate = (rate0,rate1)

    def set_IBRateSpecial(self,ae,rate0,rate1):
        ANCurrencySUMBASE.IBRateSpecial[ae] = (rate0,rate1)

    def set_IBClientRateSpecial(self,client,rate0,rate1):
        ANCurrencySUMBASE.IBClientRateSpecial[client] = (rate0,rate1)
        
    def get_headcc(self):
        return len(self.colTitles)
        
    def writeTAB(self,wsheet,row,ae,month,client,outDicts,firstRow):
        real_jump = u"worksheet!"
        add_qty2 = u"%s$%s" % (real_jump,xlwt.Utils.rowcol_to_cell(row+firstRow-6,self.__colBegin,True))
        add_rebate2 = u"%s$%s" % (real_jump,xlwt.Utils.rowcol_to_cell(row+firstRow-6,self.__colBegin+self.__len__()-1,True))
        wsheet.write(row,self.MyID*2+self.leftskip,xlwt.Formula(add_qty2))
        wsheet.write(row,self.MyID*2+self.leftskip+1,xlwt.Formula(add_rebate2))
        
        add_qty = u"$%s" % xlwt.Utils.rowcol_to_cell(row,self.MyID*2+self.leftskip,True)
        add_rebate = u"$%s" % xlwt.Utils.rowcol_to_cell(row,self.MyID*2+self.leftskip+1,True)
        if not TAB_SUMMARY_ADD.has_key(ae):
            TAB_SUMMARY_ADD[ae] = {}
        if not TAB_SUMMARY_ADD[ae].has_key(client):
            TAB_SUMMARY_ADD[ae][client] = {u'HKD':{u'qty':[],u'rebate':[]},u'USD':{u'qty':[],u'rebate':[]}}
        if (self.__currency == u'HKD'):
                TAB_SUMMARY_ADD[ae][client][u'HKD'][u'qty'].append(add_qty)
                TAB_SUMMARY_ADD[ae][client][u'HKD'][u'rebate'].append(add_rebate)
        else:
                TAB_SUMMARY_ADD[ae][client][u'USD'][u'qty'].append(add_qty)
                TAB_SUMMARY_ADD[ae][client][u'USD'][u'rebate'].append(add_rebate)
          
                
    def write(self,wsheet,row,ae,month,client,outDicts):
        if not outDicts.has_key(self.__currency):
            return
        real_jump = u"worksheet!"
        outDict = outDicts[self.__currency]
        i=0
        wsheet.write(row,self.__colBegin+i,outDict[u'Total Qty']) #u'市场总张数'
        add_qty = u"%s$%s" % (real_jump,xlwt.Utils.rowcol_to_cell(row,self.__colBegin,True))
        i += 1
        wsheet.write(row,self.__colBegin+i,outDict[u'Net Income']) #u'总收入'
        i += 1
        if (u'对美元汇率' in self.colTitles):
            wsheet.write(row,self.__colBegin+i,xlwt.Formula(ANCurrencySUMBASE.CCYRATE_ADDR[self.__currency])) #u'对美元汇率'
            i += 1
        if (self.__currency in (u'HKD',u'USD') ):
            m_TotalIncome = u'%f' % outDict[u'Net Income']
        else:
            m_TotalIncome = u"%f * %s" % (outDict[u'Net Income'],ANCurrencySUMBASE.CCYRATE_ADDR[self.__currency])
        if (u'折净美元总收入' in self.colTitles):
            wsheet.write(row,self.__colBegin+i,xlwt.Formula(m_TotalIncome))
            i += 1
        wsheet.write(row,self.__colBegin+i,xlwt.Formula(u"%s - %f" % (m_TotalIncome,outDict[u'SP Fee']))) #u'折美元公司净收入'
        i += 1
        if (ANCurrencySUMBASE.IBClientRateSpecial.has_key(client)):
            ibrate = ANCurrencySUMBASE.IBClientRateSpecial[client]
        else:
            if ANCurrencySUMBASE.IBRateSpecial.has_key(ae):
                ibrate = ANCurrencySUMBASE.IBRateSpecial[ae]
            else:
                ibrate = ANCurrencySUMBASE.IBRate
        wsheet.write(row,self.__colBegin+i,xlwt.Formula(u"$%s*%f" % (xlwt.Utils.rowcol_to_cell(row,self.__colBegin+i-1,True),ibrate[0]))) #u'折美元公司净留存'
        i += 1
        wsheet.write(row,self.__colBegin+i,xlwt.Formula(u"$%s*%f" % (xlwt.Utils.rowcol_to_cell(row,self.__colBegin+i-2,True),ibrate[1]))) #u'折美元返佣'
        add_rebate = u"%s$%s" % (real_jump,xlwt.Utils.rowcol_to_cell(row,self.__colBegin+self.__len__()-1,True))
        '''
        if not TAB_SUMMARY_ADD.has_key(ae):
            TAB_SUMMARY_ADD[ae] = {}
        if not TAB_SUMMARY_ADD[ae].has_key(client):
            TAB_SUMMARY_ADD[ae][client] = {u'HKD':{u'qty':[],u'rebate':[]},u'USD':{u'qty':[],u'rebate':[]}}
        if (self.__currency == u'HKD'):
                TAB_SUMMARY_ADD[ae][client][u'HKD'][u'qty'].append(add_qty)
                TAB_SUMMARY_ADD[ae][client][u'HKD'][u'rebate'].append(add_rebate)
        else:
                TAB_SUMMARY_ADD[ae][client][u'USD'][u'qty'].append(add_qty)
                TAB_SUMMARY_ADD[ae][client][u'USD'][u'rebate'].append(add_rebate)
        '''
                
        
    
    def writeHead(self,wsheet):
        wsheet.write_merge(4,4,self.__colBegin,self.__colBegin+len(self.colTitles)-1,self.__currency,TimStyles[u'headerBIG'])
        for i,v in enumerate(self.colTitles):
            wsheet.write(5,self.__colBegin+i,v,TimStyles[u'header'])

    def writeHeadTAB(self,wsheet):
        wsheet.write_merge(4,4,self.MyID*2+self.leftskip,self.MyID*2+self.leftskip+1,self.__currency,TimStyles[u'header'])
        wsheet.write(5,self.MyID*2+self.leftskip,u'市场总张数',TimStyles[u'header'])
        if self.__currency == u'HKD':
            wsheet.write(5,self.MyID*2+self.leftskip+1,u'港币返佣',TimStyles[u'header'])
        else:
            wsheet.write(5,self.MyID*2+self.leftskip+1,u'美元返佣',TimStyles[u'header'])
                        
    def writeSumTAB(self,sheet,row,row1,row2):
            for i,v in enumerate([u'市场总张数',u'返佣']):  
                    c = self.MyID*2+self.leftskip+i
                    if (row2>row1):
                        sheet.write(row,c,xlwt.Formula("SUM($%s:$%s)" % (xlwt.Utils.rowcol_to_cell(row1,c,True),xlwt.Utils.rowcol_to_cell(row2,c,True))),TimStyles[u'sum']) 
                    else:
                        sheet.write(row,c,xlwt.Formula("$%s" % xlwt.Utils.rowcol_to_cell(row1,c,True)),TimStyles[u'sum']) 

    def writeSum(self,sheet,row,row1,row2):
            for i,v in enumerate(self.colTitles):        
                c = self.__colBegin + i
                if (u'对美元汇率' == v):
                    sheet.write(row,c,xlwt.Formula(ANCurrencySUMBASE.CCYRATE_ADDR[self.__currency]),TimStyles[u'ccy']) 
                else:
                    if (row2>row1):
                        sheet.write(row,c,xlwt.Formula("SUM($%s:$%s)" % (xlwt.Utils.rowcol_to_cell(row1,c,True),xlwt.Utils.rowcol_to_cell(row2,c,True))),TimStyles[u'sum']) 
                    else:
                        sheet.write(row,c,xlwt.Formula("$%s" % xlwt.Utils.rowcol_to_cell(row1,c,True)),TimStyles[u'sum']) 

                cell = "$%s" % xlwt.Utils.rowcol_to_cell(row,c,True)
                if (self.__currency == u'HKD'):
                    SUMMARY[u'HKD'][v].append(cell)
                else:
                    if (v in (u'市场总张数',u'折美元公司净收入',u'折美元公司净留存',u'折美元返佣')):                    
                        SUMMARY[u'USD'][v].append(cell)
 
 
 
    def writeSumHKD(self,sheet,row,row0): 
        for i,v in enumerate(self.colTitles):        
            c = self.__colBegin + i           
            if (v in (u'折净美元总收入',u'折美元公司净收入',u'折美元公司净留存',u'折美元返佣')):
                sheet.write(row,c,xlwt.Formula("$%s/%s" % (xlwt.Utils.rowcol_to_cell(row0,c,True),ANCurrencySUMBASE.CCYRATE_ADDR[u'HKD'])),TimStyles[u'sum2']) 
            else:
                sheet.write(row,c,'',TimStyles[u'sum2']) 

    def writeSumHKDTAB(self,sheet,row,row0,flag=False): 
        caddr = u'$D$3'
        for i,v in enumerate([u'市场总张数',u'返佣']):  
            c = self.MyID*2+self.leftskip+i        
            if (self.__currency == u'HKD'):
                sheet.write(row,c,xlwt.Formula("$%s" % xlwt.Utils.rowcol_to_cell(row0,c,True)),TimStyles[u'sum2']) 
            else:
                sheet.write(row,c,xlwt.Formula("$%s/%s" % (xlwt.Utils.rowcol_to_cell(row0,c,True),caddr)),TimStyles[u'sum2']) 
                

class ANCurrencySUM(ANCurrencySUMBASE):
    colTitles = (u'市场总张数',u'总收入',u'对美元汇率',u'折净美元总收入',u'折美元公司净收入',u'折美元公司净留存',u'折美元返佣')
    def __init__(self,currency,begincol):
        super(ANCurrencySUM, self).__init__(currency,begincol)        
        
class ANCurrencySUM_USD(ANCurrencySUMBASE):
    colTitles = (u'市场总张数',u'总收入',u'折净美元总收入',u'折美元公司净收入',u'折美元公司净留存',u'折美元返佣')
    def __init__(self,currency,begincol):
        super(ANCurrencySUM_USD, self).__init__(currency,begincol)        


class ANCurrencySUM_HK(ANCurrencySUMBASE):
    colTitles = (u'市场总张数',u'总收入',u'港币净收入(减SP费用后)',u'港币公司净留存',u'港币返佣')
    def __init__(self,currency,begincol):
        super(ANCurrencySUM_HK, self).__init__(currency,begincol)  
      
    
class ReadSeasonSUM():
    global logger
    Titles = season_Titles
    __c_Titles = season_c_Titles
    __AE = {}
    __Sales_idx = xlwt.Utils.col_by_name('A')
    __Client_idx = xlwt.Utils.col_by_name('B')
    __Ccy_idx = xlwt.Utils.col_by_name('E')
    __SP_FEE_MultiplierLocal = {u'HHIF':3.0,u'HHIO':3.0,u'HSIF':3.0,u'HSIO':3.0,
                            u'MCHF':1.0,u'MHSIF':1.0,u'MHSIO':1.0}
    __SP_FEE_MultiplierGlobal = 1.0
       
    def run(self,sheets):
        for m in range(3):        
            sheet = sheets[m]
            for r in xrange(1,sheet.nrows):
                rv = dict(zip(self.Titles,sheet.row_values(r)))
                self.dealSPFEE(rv)
                if not self.__AE.has_key(rv[u'Sales']):
                    self.__AE[rv[u'Sales']] = {}
                if not self.__AE[rv[u'Sales']].has_key(m):
                    self.__AE[rv[u'Sales']][m] = {}
                if not self.__AE[rv[u'Sales']][m].has_key(rv[u'Client']):
                    self.__AE[rv[u'Sales']][m][rv[u'Client']] = {}
                if not self.__AE[rv[u'Sales']][m][rv[u'Client']].has_key(rv[u'Ccy']):
                    self.__AE[rv[u'Sales']][m][rv[u'Client']][rv[u'Ccy']] = rv
                else:
                    m_rsd = self.__AE[rv[u'Sales']][m][rv[u'Client']][rv[u'Ccy']]
                    for c in self.__c_Titles:
                        m_rsd[c] = m_rsd[c] + rv[c]
                    m_rsd[u'SP Fee'] = m_rsd[u'SP Fee'] + rv[u'SP Fee']
                    self.__AE[rv[u'Sales']][m][rv[u'Client']][rv[u'Ccy']] = m_rsd    
                        
    def setSP_FEE_MultiplierLocal(self,key,value):
        ReadSeasonSUM.__SP_FEE_MultiplierLocal[key] = value

    def getSP_FEE_MultiplierLocal(self,p_key):
        key = strip(p_key)
        if (ReadSeasonSUM.__SP_FEE_MultiplierLocal.has_key(key)):
            return ReadSeasonSUM.__SP_FEE_MultiplierLocal[key]
        else:
            logger.error("Product:%s not find in conf file section HKEX_SP_FEE!" % key)                
            return 1.0
    
    def setSP_FEE_MultiplierGlable(self,value):
        ReadSeasonSUM.__SP_FEE_MultiplierGlobal = value

    def getSP_FEE_MultiplierGlable(self):
        return ReadSeasonSUM.__SP_FEE_MultiplierGlobal
        
        
    def dealSPFEE(self,rv): # 处理净收入减掉SP的费用特殊方法
        if (rv[u'Ccy']=='HKD'):
            rv[u'SP Fee'] = rv[u'Total Qty'] * self.getSP_FEE_MultiplierLocal(rv[u'Product'])
        else:
            rv[u'SP Fee'] = rv[u'Total Qty'] * self.getSP_FEE_MultiplierGlable()

    def get_client(self,AE,month,client):
        rs = {}
        am = self.get_month(AE,month)
        if (not am): return 
        if am.has_key(client):
            return am[client] 
        
    def get_month(self,AE,month):
        rs = {}
        if (type(month) == int):
            m = month
        else:
            m = months1idx[month]
        ae = self.get_AE(AE)
        if (ae): 
            if ae.has_key(m):
                return ae[m]

    def getKeys(self):
        return self.__AE.keys()

    def get_AE(self,AE):
        if self.__AE.has_key(AE):
            return self.__AE[AE]
        else:
            return 
    
    def __repr__(self):
        lines = ''
        for k,v in self.__AE.items():
            lines += "\n*************%s*************\n" % k
            for kk,vv in v.items():
                for kkk,vvv in vv.items():
                    for kkkk,vvvv in vvv.items():
                        lines += "%s : %s\n" %(kkkk,vvvv)
        return lines

    def __len__(self):
        return len(self.__AE)
        
    def __getitem__(self,key):  
        return self.__AE[key]

class TimRebate:
    __CCYRATE = CCYRATE 
    def __init__(self,rootDir = None, mylogger = None):
        global logger
        self.initFlag = False
        self.succeed = False
        if rootDir:
            self.m_rootDir = rootDir
        else:
            self.m_rootDir = os.path.curdir      
        self.m_rootDir = os.path.realpath(self.m_rootDir)
        if mylogger:
            logger = mylogger
        else:
            logger = setLogger(os.path.curdir,level = logging.DEBUG)     
        self.seasonSUM = ReadSeasonSUM()    
        self.readConf()
        self.initFlag = True
        
    def __call__(self, p_y = None , p_s = None): 
        self.TimYear = strip(p_y)
        self.TimQuarter = p_s
        logger.info( u"开始运行返佣辅助模块" )
        logger.info( u"年:%s  季:%s" % (self.TimYear,self.TimQuarter))
        self.succeed = True
        bookmfn = []
        self.dirname = os.path.join(self.m_rootDir,u'rebate')
        bookfilename = os.path.join(self.dirname,u'rebate %s%1iq.%s' % (self.TimYear,self.TimQuarter,destfileext))
        for i in range(3):
            bookmfn.append(os.path.join(self.dirname,u'%s.%s' % (seasons1[self.TimQuarter][i],orgfileext)))
        try:
            self.xlsBook = xlwt.Workbook()
            logger.info( u"开始生成新xls文件...")
            self.xlsSheet = self.xlsBook.add_sheet(u'worksheet')
        except Exception,e:
            logger.error("生成新xls文件过程出错！%s" % e)   
            self.succeed = False 
        if (self.succeed):
            xlsBookm = []   
            for i in range(3):        
                try:
                    xlsBookm.append(xlrd.open_workbook(bookmfn[i]))
                except Exception,e:
                    logger.error("打开季度数据文件过程出错，error=%s！",e)
                    self.succeed = False
            if (self.succeed):
                self.xlsSheetm = []  
                self.xlsSheetmLastRow = []
                for i in range(3):        
                    try:
                        self.xlsSheetm.append(xlsBookm[i].sheet_by_index(0)) 
                        self.xlsSheetmLastRow.append(self.xlsSheetm[i].nrows) 
                        logger.info( u"已打开文件：%s,共%s行数据" % (bookmfn[i],self.xlsSheetmLastRow[i]-1)) #去掉第一行title，所以 -1
                    except  Exception,e:
                        logger.error("初始化季度数据文件过程出错，error=%s！" % e)
                        self.succeed = False
                self.seasonSUM.run(self.xlsSheetm)
                self.dealSheets()
            else:
                logger.error("初始化季度数据文件过程出错. succeed = False！")
                
            
    def _initrebate(self):
        if (not self.succeed):
            logger.error("初始化过程出错，请检查初始化过程需要的原始数据文件！")
            return         
        logger.info( u"初始化rebate表，把要填写的部分先置空!")
        self.predealsheet01(True)
  
    def writeAAE(self,wSheet,firstRow,ae,mkt):
        row = firstRow
        for m in xrange(3):
            rs = self.seasonSUM.get_month(ae,m)
            m_row = row
            if rs:
                keys = rs.keys()
                keys.sort()
                for client in keys:
                    wSheet.write(row,2,client)
                    for mm in mkt:
                        mm.write(wSheet,row,ae,m,client,rs[client])
                    row += 1
                wSheet.write_merge(m_row,row-1,1,1,seasons1[self.TimQuarter][m],TimStyles[u'id'])
        row += self.writeSum(wSheet,row,firstRow,row-1,mkt)
        wSheet.write_merge(firstRow,row-1,0,0,ae,TimStyles[u'id'])
        return row -  firstRow

    def writeAAE_TAB(self,sheet,firstRow,ae,mkt):
        book = sheet.parent
        wSheet = book.add_sheet(ae)
        
        keys = CCYRATE.keys()
        wSheet.write(1,2,u'汇率:',TimStyles[u'header'])
        for c,v  in  enumerate(keys):
            wSheet.write(1,3+2*c,xlwt.Formula(u"worksheet!$%s" % xlwt.Utils.rowcol_to_cell(2,3+2*c,True)),TimStyles[u'header'])
            wSheet.write(1,4+2*c,xlwt.Formula(u"worksheet!$%s" % xlwt.Utils.rowcol_to_cell(2,4+2*c,True)),TimStyles[u'currency']) 
            if (v == u'USD/HKD'):
                wSheet.write(2,3,xlwt.Formula(u"$%s" % xlwt.Utils.rowcol_to_cell(1,4+2*c,True))) 
                
        wSheet.write(2,2,u'港币汇率:')        
        
        wSheet.write_merge(4,5,1,1,u'结算月份',TimStyles[u'header'])             
        wSheet.write_merge(4,5,2,2,u'客户号',TimStyles[u'header'])        
        for mm in mkt:
            mm.writeHeadTAB(wSheet)    
        row = 6
        for m in xrange(3):
            rs = self.seasonSUM.get_month(ae,m)
            m_row = row
            if rs:
                keys = rs.keys()
                keys.sort()
                for client in keys:
                    wSheet.write(row,2,client)
                    for mm in mkt:
                        mm.writeTAB(wSheet,row,ae,m,client,rs[client],firstRow)
                    row += 1
                wSheet.write_merge(m_row,row-1,1,1,seasons1[self.TimQuarter][m],TimStyles[u'id'])

        self.writeSumTAB(wSheet,row,6,row-1,mkt)
        row += 2
        self.writeAAE_TAB_SUM(wSheet,row,ae)


    def writeAAE_TAB_SUM(self,sheet,firstRow,ae):
        row = firstRow+5
        sheet.write(row,2,u'客户帐号',TimStyles[u'summary'])
        sheet.write(row,3,u'',TimStyles[u'summary'])
        sheet.write(row,4,u'香港市场总张数',TimStyles[u'summary'])
        sheet.write(row,5,u'',TimStyles[u'summary'])
        sheet.write(row,6,u'香港市场总返佣',TimStyles[u'summary'])
        sheet.write(row,7,u'',TimStyles[u'summary'])
        sheet.write(row,8,u'非香港市场总张数',TimStyles[u'summary'])
        sheet.write(row,9,u'',TimStyles[u'summary'])
        sheet.write(row,10,u'非香港市场总返佣',TimStyles[u'summary'])
        row += 1
        s_r1 = row
        clients = TAB_SUMMARY_ADD[sheet.name]
        TAB_SUMMARY[sheet.name] = {}
        for k,v in clients.items():
            TAB_SUMMARY[sheet.name][k] = {}
            sheet.write(row,2,k)
            if v[u'HKD'][u'qty']:
                sheet.write(row,4,xlwt.Formula(join(v[u'HKD'][u'qty'],'+')))
                TAB_SUMMARY[sheet.name][k][u'HKDQTY'] = u"$%s" % xlwt.Utils.rowcol_to_cell(row,4,True)
            if v[u'HKD'][u'rebate']:
                sheet.write(row,6,xlwt.Formula(join(v[u'HKD'][u'rebate'],'+')))
                TAB_SUMMARY[sheet.name][k][u'HKDREBATE'] = u"$%s" % xlwt.Utils.rowcol_to_cell(row,6,True)
            if v[u'USD'][u'qty']:
                #print join(v[u'USD'][u'qty'],'+')
                sheet.write(row,8,xlwt.Formula(join(v[u'USD'][u'qty'],'+')))
                TAB_SUMMARY[sheet.name][k][u'USDQTY'] = u"$%s" % xlwt.Utils.rowcol_to_cell(row,8,True)
            if v[u'USD'][u'rebate']:
                sheet.write(row,10,xlwt.Formula(join(v[u'USD'][u'rebate'],'+')))
                TAB_SUMMARY[sheet.name][k][u'USDREBATE'] = u"$%s" % xlwt.Utils.rowcol_to_cell(row,10,True)
            row += 1
        s_r2 = row -1
        sheet.write(row,2,u'小计',TimStyles[u'sum'])
        for c in [3,5,7,9]:
            sheet.write(row,c,u'',TimStyles[u'sum'])
        for c in [4,6,8,10]:
            sheet.write(row,c,xlwt.Formula(u"SUM($%s:$%s)" % (xlwt.Utils.rowcol_to_cell(s_r1,c,True),xlwt.Utils.rowcol_to_cell(s_r2,c,True))),TimStyles[u'sum'])
        row += 1
        caddr = u'$D$3'
        sheet.write(row,2,u'港币小计',TimStyles[u'sum2'])
        for c in [3,4,5,7,8,9]:
            sheet.write(row,c,u'',TimStyles[u'sum2'])
        sheet.write(row,6,xlwt.Formula(u"$%s" % xlwt.Utils.rowcol_to_cell(row-1,6,True)),TimStyles[u'sum2'])
        sheet.write(row,10,xlwt.Formula(u"$%s/%s" % (xlwt.Utils.rowcol_to_cell(row-1,10,True),caddr)),TimStyles[u'sum2'])
        row += 2
        sheet.write(row,2,u'港币总计：',TimStyles[u'sum2'])
        sheet.write(row,3,xlwt.Formula(u"$%s+$%s" % (xlwt.Utils.rowcol_to_cell(row-2,6,True),xlwt.Utils.rowcol_to_cell(row-2,10,True))),TimStyles[u'sum2'])
        TAB_SUMMARY[sheet.name][u'total']={u'HKDREBATE':u"$%s" % xlwt.Utils.rowcol_to_cell(row-3,6,True),
            u'USDREBATE':u"$%s" % xlwt.Utils.rowcol_to_cell(row-3,10,True)}
        TAB_SUMMARY[sheet.name][u'totalHKD']={    
            u'subsumHKD1':u"$%s" % xlwt.Utils.rowcol_to_cell(row-2,6,True),
            u'subsumHKD2':u"$%s" % xlwt.Utils.rowcol_to_cell(row-2,10,True),
            u'HKDREBATE':u"$%s" % xlwt.Utils.rowcol_to_cell(row,3,True)}

        self.dealRebate(sheet,row,ae)
        
        
    def dealRebate(self,wSheet,firstRow,ae): 
        caddr = u'$D$3'    
        __rebate = {}
        row = firstRow + 5
        if ae in self.cf.sections():
            name = self.cf.get(ae,u'name')
            wSheet.write(row,1,u'%s' % name,TimStyles[u'sum2'])
            row += 1
            wSheet.write(row,1,u'返佣分配',TimStyles[u'sum2'])
            wSheet.write(row,2,u'姓名',TimStyles[u'sum2'])
            wSheet.write(row,4,u'香港市场返佣(港币)',TimStyles[u'sum2'])
            wSheet.write(row,6,u'海外市场返佣（美元）',TimStyles[u'sum2'])
            wSheet.write(row,8,u'总返佣（港币）',TimStyles[u'sum2'])
            for c in [3,5,7,9,10]:
                wSheet.write(row,c,u'',TimStyles[u'sum2'])

            cv = self.cf.get(ae,u'METHODS')
            __rebate[u'methodsStr'] = cv
            __rebate[u'resultList'] = []
            __rebate[u'rebate'] = TAB_SUMMARY[wSheet.name]
            self.dealRebateMethod(__rebate,ae,level=0)
              
            row += 1
            __out_d = {}
            for (c,v1,v2) in __rebate[u'resultList']:
                #print c,v1,v2
                if __out_d.has_key(c):
                   __out_d[c][0] = u'%s + %s' %  (__out_d[c][0],v1)
                   __out_d[c][1] = u'%s + %s' %  (__out_d[c][1],v2)
                else:
                   __out_d[c] = [v1,v2]
            begin_rebate_row = row        
            for k,v in __out_d.items():
                wSheet.write(row,2,k)              
                wSheet.write(row,4,xlwt.Formula(v[0]),TimStyles[u'summary'])              
                wSheet.write(row,6,xlwt.Formula(v[1]),TimStyles[u'summary'])              
                wSheet.write(row,8,xlwt.Formula(u'%s + (%s)/%s' % (xlwt.Utils.rowcol_to_cell(row,4,True),xlwt.Utils.rowcol_to_cell(row,6,True),caddr)),TimStyles[u'summary'])              
                row += 1
            wSheet.write(row,1,u'返佣总计',TimStyles[u'sum'])
            for c in [2,3,5,7,9]:
                wSheet.write(row,c,u'',TimStyles[u'sum'])

            wSheet.write(row,4,xlwt.Formula(u"SUM($%s:$%s)" % (xlwt.Utils.rowcol_to_cell(begin_rebate_row,4,True),xlwt.Utils.rowcol_to_cell(row-1,4,True))),TimStyles[u'sum'])
            wSheet.write(row,6,xlwt.Formula(u"SUM($%s:$%s)" % (xlwt.Utils.rowcol_to_cell(begin_rebate_row,6,True),xlwt.Utils.rowcol_to_cell(row-1,6,True))),TimStyles[u'sum'])
            wSheet.write(row,8,xlwt.Formula(u"SUM($%s:$%s)" % (xlwt.Utils.rowcol_to_cell(begin_rebate_row,8,True),xlwt.Utils.rowcol_to_cell(row-1,8,True))),TimStyles[u'sum'])
            
            
    def dealRebateMethod(self,p_rebate,ae,level=0):
        caddr = u'$D$3'
        rebate = p_rebate[u'rebate']
        p_rebate[u'methodsList'] = self.parseMethodsStr(p_rebate[u'methodsStr'])
        p_rebate[u'childsList'] = []
        p_rebate[u'level'] = level

        if (p_rebate[u'methodsList'][u'confirm']):
            for m in  p_rebate[u'methodsList'][u'confirm']:
                if len(m)==3 or len(m)==4 :
                    v = m[2]
                    if (level):
                        if len(rebate) > 1:
                            logger.error(u'methods 格式定义错误，confirm项收到的待分配项不唯一，只有第一层分配才允许不唯一，AE:%s' % ae)
                        else:
                            rebate = u"%s - %f" % (rabate.items()[0].value, v)
                    else:
                        pass
                    if len(m)==4:
                            try:
                                cv = self.cf.get(ae,m[3])
                                __rebate = {}
                                __rebate[u'methodsStr'] = cv
                                __rebate[u'resultList'] = p_rebate[u'resultList']
                                __rebate[u'rebate'] = rebate 
                                p_rebate[u'childsList'].append(__rebate)
                            except Exception,e:
                                logger.error(u'get conf AE:%s,option:%s, error! %s' % (ae,m[3],e))
                                raise e
                    else:
                        pass
                else:
                        logger.error(u'methods client必须是由4部分组成！ AE:%s, %s' % (ae,p_rebate[u'methods']))
                        
        if (p_rebate[u'methodsList'][u'client']):
            for m in  p_rebate[u'methodsList'][u'client']:
                if len(m)==4:             
                    cvs = split(m[3].lstrip(u'(').rstrip(u')'),u',')
                    for c in cvs:
                        if rebate.has_key(c+u'-000'):
                            try:
                                cv = self.cf.get(ae,c)
                                __rebate = {}
                                __rebate[u'methodsStr'] = cv
                                __rebate[u'resultList'] = p_rebate[u'resultList']
                                __rebate[u'rebate'] = {c+u'-000':rebate[c+u'-000']}
                                p_rebate[u'childsList'].append(__rebate)
                            except Exception,e:
                                logger.error('get conf AE:%s,option:%s, error! %s' % (ae,c,e))
                                raise e
                else:
                        logger.error(u'methods client必须是由4部分组成！ AE:%s, %s' % (ae,p_rebate[u'methods']))

        if (p_rebate[u'methodsList'][u'rate']):
            #if p_rebate[u'methodsList'][u'confirm']:
            #    pass #减掉confirm数
            #if p_rebate[u'methodsList'][u'client']:
            #    for c in p_rebate[u'methodsList'][u'client']:
            #       del rebate[c[3]+u'-000']
            #    pass #减掉client数  
            if level == 0:
               rebate = {u'total':rebate[u'total']} 
            for m in  p_rebate[u'methodsList'][u'rate']:
                rate = m[2]*1.0000
                #print m[0],m[1],m[2]
                c_key0 = rebate.keys()[0]
                c_v = rebate[c_key0]
                c_v_hkd = c_v[u'HKDREBATE']
                c_v_usd = c_v[u'USDREBATE']
                if len(m)==3:
                    p_rebate[u'resultList'].append((u'%s' % m[0],u'(%s) * %f' % (c_v_hkd,rate/100),u'(%s) * %f' % (c_v_usd,rate/100)))
                elif len(m)==4:
                        try:
                            cv = self.cf.get(ae,m[3])
                            __rebate = {}
                            __rebate[u'methodsStr'] = cv
                            __rebate[u'resultList'] = p_rebate[u'resultList']
                            __rebate[u'rebate'] = {c_key0:{u'HKDREBATE':u'(%s) * %f' % (c_v_hkd,rate/100),u'USDREBATE':u'(%s) * %f' % (c_v_usd,rate/100)}} 
                            p_rebate[u'childsList'].append(__rebate)
                        except Exception,e:
                            logger.error('get conf AE:%s,option:%s, error! %s' % (ae,m[3],e))
                            raise e
     
                else:
                        logger.error(u'methods client必须是由4部分组成！ AE:%s, %s' % (ae,p_rebate[u'methods']))

                
        for m in p_rebate[u'childsList']:
            self.dealRebateMethod(m,ae,level+1)
   
   

    def parseMethodsStr(self,methodsStr):
        L0 = [split(i,u':') for i in split(methodsStr,u';')]
        L1 = filter(lambda x: len(x) in (3,4),L0)
        ml = [[],[],[],[]]
        for i,v in enumerate(L1):
            if v[1] ==  u'confirm':
                v[2] = float(v[2])*1.0000
                ml[0].append(v)
            if v[1] ==  u'client':
                ml[1].append(v)
            if v[1] ==  u'rate':
                v[2] = float(v[2])*1.0000
                ml[2].append(v)
            if v[1] ==  u'average':
                ml[3].append(v)
        ttleft = 100.0000
        if (ml[1]):
            mll=[]
            for m in ml[1]:
                cvs = split(m[3].lstrip(u'(').rstrip(u')'),u',')  
                for c in cvs:
                    mll.append([m[0],m[1],m[2],c])
            ml[1]=mll
        if (ml[2]):
            total = sum([i[2] for i in ml[2]])
            
            if (total < 100):
                if (ml[3]):
                    ttleft = 100.0000 - total 
                else:
                    t3 =  100.0000/total
                    for i in ml[2]:
                        i[2] =i[2]*t3
            else :
                ttleft = 0
                if (total > 100):
                    t3 =  100.0000/total
                    for i in ml[2]:
                        i[2] =i[2]*t3 
            
        if (ml[3]):
            if (ttleft > 0):
                rr = ttleft/len(ml[3])
                for i in ml[3]:
                    i[1] = u'rate'
                    i[2] = rr
                    ml[2].append(i)
            else:
                logger.error('methods 定义有问题，有average项，但rate项合计超100:%s' % p_rebate[u'methods'])
        return  {u'confirm':ml[0],u'client':ml[1],u'rate':ml[2]}
                
            
    
          
    def writeHEAD(self,sheet,mkt):
        sheet.write_merge(0,1,0,ANCurrencySUMBASE.TOTAL_LEN+2,u'广发期货（香港）有限公司返佣计算总表',TimStyles[u'headerBIG0'])
        keys = self.__CCYRATE.keys()
        sheet.write(2,2,u'汇率:',TimStyles[u'header'])
        for c,v  in  enumerate(keys):
            sheet.write(2,3+2*c,v,TimStyles[u'header'])
            sheet.write(2,4+2*c,self.__CCYRATE[v],TimStyles[u'currency'])   
            ANCurrencySUMBASE.CCYRATE_ADDR[split(v,u'/')[1]] = "$%s" % xlwt.Utils.rowcol_to_cell(2,4+2*c,True)         
        sheet.write_merge(3,3,3,2+9,u'注意：改变上行的汇率会影响所有的计算结果！')

        sheet.write_merge(4,4,0,2,u'ID',TimStyles[u'headerBIG'])
        sheet.write(5,0,u'用户AE',TimStyles[u'header'])
        sheet.write(5,1,u'结算月份',TimStyles[u'header'])             
        sheet.write(5,2,u'客户号',TimStyles[u'header'])
        sheet.row(4).height = 400
        for m in mkt:
            m.writeHead(sheet)
            
    def writeSummary(self,sheet,firstRow):
        '''
        SUMMARY={u'HKD':{
                    u'市场总张数':[],
                    u'总收入':[],
                    u'港币净收入(减SP费用后)':[],
                    u'港币公司净留存':[],
                    u'港币返佣':[]         
                },
                u'USD':{
                    u'市场总张数':[],
                    u'折美元公司净收入':[],
                    u'折美元公司净留存':[],
                    u'折美元返佣':[]
                }}
        '''
        row = firstRow+5
        sheet.write(row,2,u'返佣数据总结',TimStyles[u'summary'])
        for c in xrange(1,10):
            sheet.write(row+1,c,u'',TimStyles[u'sum2'])
        sheet.write(row+2,2,u'香港市场总张数',TimStyles[u'summary'])
        sheet.write(row+2,3,xlwt.Formula(join(SUMMARY[u'HKD'][u'市场总张数'],'+')))
        sheet.write(row+3,2,u'港元总收入',TimStyles[u'summary'])
        sheet.write(row+3,3,xlwt.Formula(join(SUMMARY[u'HKD'][u'总收入'],'+')))
        sheet.write(row+4,2,u'港元公司留存',TimStyles[u'summary'])
        sheet.write(row+4,3,xlwt.Formula(join(SUMMARY[u'HKD'][u'港币公司净留存'],'+')))
        sheet.write(row+5,2,u'港元返佣总数',TimStyles[u'summary'])
        sheet.write(row+5,3,xlwt.Formula(join(SUMMARY[u'HKD'][u'港币返佣'],'+')))
        
        sheet.write(row+2,8,u'非香港品种总张数',TimStyles[u'summary'])
        sheet.write(row+2,9,xlwt.Formula(join(SUMMARY[u'USD'][u'市场总张数'],'+')))
        sheet.write(row+3,8,u'美元净收入',TimStyles[u'summary'])
        sheet.write(row+3,9,xlwt.Formula(join(SUMMARY[u'USD'][u'折美元公司净收入'],'+')))
        sheet.write(row+4,8,u'美元公司留存',TimStyles[u'summary'])
        sheet.write(row+4,9,xlwt.Formula(join(SUMMARY[u'USD'][u'折美元公司净留存'],'+')))
        sheet.write(row+5,8,u'美元返佣总数',TimStyles[u'summary'])
        sheet.write(row+5,9,xlwt.Formula(join(SUMMARY[u'USD'][u'折美元返佣'],'+')))
        cell1 = "$%s" % xlwt.Utils.rowcol_to_cell(row+5,9,True)
        sheet.write(row+6,8,u'港币/美金汇率',TimStyles[u'summary'])
        sheet.write(row+6,9,xlwt.Formula(ANCurrencySUMBASE.CCYRATE_ADDR[u'HKD']))
        cell2 = "$%s" % xlwt.Utils.rowcol_to_cell(row+6,9,True)
        sheet.write(row+7,8,u'返佣港元等值',TimStyles[u'summary'])
        sheet.write(row+7,9,xlwt.Formula("%s/%s" % (cell1,cell2)))

                
    def writeSum(self,sheet,row,row1,row2,mkt):
        sheet.write(row,2,u'小计',TimStyles[u'sum'])
        sheet.write(row+1,2,u'港币小计',TimStyles[u'sum2'])
        for m in mkt:
            m.writeSum(sheet,row,row1,row2) 
            m.writeSumHKD(sheet,row+1,row)  
        return 2

    def writeSumTAB(self,sheet,row,row1,row2,mkt):
        sheet.write(row,2,u'小计',TimStyles[u'sum'])
        sheet.write(row+1,2,u'港币小计',TimStyles[u'sum2'])
        for m in mkt:
            m.writeSumTAB(sheet,row,row1,row2) 
            m.writeSumHKDTAB(sheet,row+1,row)  
        return 2
            
    def initMKTList(self):
        MKT = list()
        begin_col = 3
        MKT.append(ANCurrencySUM_HK(u'HKD',begin_col))
        begin_col += len(MKT[0])
        #begin_col += MKT[0].get_headcc()
        MKT.append(ANCurrencySUM_USD(u'USD',begin_col))
        #begin_col += MKT[1].get_headcc()
        begin_col += len(MKT[1])
        MKT.append(ANCurrencySUM(u'JPY',begin_col))
        #begin_col += MKT[2].get_headcc()
        begin_col += len(MKT[2])
        MKT.append(ANCurrencySUM(u'MYR',begin_col))
        #begin_col += MKT[3].get_headcc()
        begin_col += len(MKT[3])
        MKT.append(ANCurrencySUM(u'CNY',begin_col))
        #begin_col += MKT[4].get_headcc()
        begin_col += len(MKT[4])
        return  MKT 
                      
    def dealSheets(self):
        wBook = xlwt.Workbook() 
        wSheet = wBook.add_sheet(u"worksheet")
        MKT=self.initMKTList()   
        self.writeHEAD(wSheet,MKT)
        
        AES = self.seasonSUM.getKeys()
        AES.sort()
        row = 6
        for ae in AES:
            row_add = self.writeAAE(wSheet,row,ae,MKT)
            self.writeAAE_TAB(wSheet,row,ae,MKT)
            row += row_add
        self.writeSummary(wSheet,row)
        try:
            wBook.save(os.path.join(self.dirname,u'rebate(%s).xls' % datetime.datetime.now().strftime(u'%Y%m%d')))
            logger.info( u"成功！！！ 已经关闭并保存文件！")
        except  Exception,e:
            logger.error(u"保存结果数据文件过程出错，error=%s！" % e)
            self.succeed = False                
        logger.info( u"predealsheet过程运行结束!")    
            
    def getAECList(self):
        Titles = ReadSeasonSUM.Titles
        AE = {}
        for i in range(3):        
            sheet = self.xlsSheetm[i]
            c_A = sheet.col(0,1)
            for r,v in enumerate(c_A):
                if not AE.has_key(v.value):
                    AE[v.value]=set()
                AE[v.value].add(sheet.cell(r+1,1).value)
        for k,v in AE.items():
            t = list(v)
            t.sort()
            AE[k]=t
        return AE
                
    def getSUMData(self):
        return 
        
        
    def parseRebateDefine(self,cf):
        self.cf_complexAE = cf.get(u"RULES", u"complexAE")
        self.cf_METHODS = cf.get(u"RULES", u"METHODS")
    
    def readConf(self):
        fn = os.path.realpath(os.path.join(self.m_rootDir,u'rebate.conf'))
        self.cf = ConfigParser.ConfigParser()    
        try:
            self.cf.read(fn)   
            for key in (u"USD/HKD",u"USD/USD",u"USD/JPY",u"USD/MYR",u"USD/CNY"):
                CCYRATE[key] = self.cf.getfloat(u"CCYRATE", key)
            
            HKEX_SP_FEE_HKD_A = self.cf.get(u"HKEX_SP_FEE", u"HKD_A")
            Multiplier_A  = self.cf.getfloat(u"HKEX_SP_FEE", u"Multiplier_A")
            HKEX_SP_FEE_HKD_B = self.cf.get(u"HKEX_SP_FEE", u"HKD_B")
            Multiplier_B  = self.cf.getfloat(u"HKEX_SP_FEE", u"Multiplier_B")
            for i in split(HKEX_SP_FEE_HKD_A,','):
                self.seasonSUM.setSP_FEE_MultiplierLocal(i,Multiplier_A)
            for i in split(HKEX_SP_FEE_HKD_B,','):
                self.seasonSUM.setSP_FEE_MultiplierLocal(i,Multiplier_B)
            self.__SP_FEE_MultiplierGlobal =  self.cf.getfloat(u"GLOBAL_SP_FEE", u"Multiplier")
            self.seasonSUM.setSP_FEE_MultiplierGlable(self.__SP_FEE_MultiplierGlobal)

        except Exception, e:
            logger.error(u'Read conf file error! %s' % e)    
        
        self.parseIBRate(self.cf)    
        self.parseRebateDefine(self.cf)

    def parseIBRate(self,cf):
        cv = cf.get(u"IBRATE", u"default")
        cvv = split(cv,u':')
        IBRate[0] = float(cvv[0])*1.0000
        IBRate[1] = float(cvv[1])*1.0000
        cv = cf.get(u"IBRATE", u"special")
        cvs = split(cv.lstrip(u'(').rstrip(u')'),u',')
        for r in cvs:  
            rv = split(r,u':')
            cv = cf.get(u"IBRATE", rv[1])
            cvv = split(cv,u':')                
            IBRateSpecial[rv[0]] = (float(cvv[0])*1.0000,float(cvv[1])*1.0000) 
        self.parseIBClientRate(cf)
        
    def parseIBClientRate(self,cf):
        cv = cf.get(u"IBRATE", u"accountspecial")
        cvs = split(cv.lstrip(u'(').rstrip(u')'),u',')
        for r in cvs:  
            rv = split(r,u':')
            cv = cf.get(u"IBRATE", rv[1])
            cvv = split(cv,u':')                
            IBClientRateSpecial[rv[0]] = (float(cvv[0])*1.0000,float(cvv[1])*1.0000) 


def test():
    global logger,m_rootdir
    xls = TimRebate(rootDir=m_rootdir)
    xls('2014',3)    
    

def main():
    global logger,m_rootdir
    import argparse
    __author__ = 'TianJun'
    parser = argparse.ArgumentParser(description='This is a rebate script by TianJun.')
    parser.add_argument('-d','--rootdir', help='Input log file dir,default is ".".',required=False)
    parser.add_argument('-l','--level', help='Input logger level, default is INFO.',required=False)
    args = parser.parse_args()    
    if (args.rootdir):
        m_rootdir = args.rootdir
    else:
        m_rootdir = os.path.join(os.path.curdir,u'..' )
    if (args.level):
        m_level = args.level  
    else:
        m_level = logging.DEBUG    
    test()
    
    return 0

if __name__ == '__main__':

    main()
        
        
