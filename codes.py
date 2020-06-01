from tkinter import *
from threading import Thread
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showinfo
from os.path import isdir
from os.path import isfile
import pandas as pd
import numpy as np
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
# import os
from os import listdir
import docx
# import re
from re import match
import xlsxwriter
from xlrd import XLRDError
import datetime
import copy
# from copy import 

class ViewController():
    
    def __init__(self, app, title, needReturn=False):
        self.app = app
        self.app.title(title)
        self.frame = Frame(app)
        self.frame.pack()
        self.setupUI(self.frame)
        if needReturn:
            self.setupReturn(self.frame)

    def setupReturn(self, frame):
        Button(frame, text=' < ', bg='CornflowerBlue', fg='GhostWhite', command=self.back).grid(row=0, column=0, sticky= W + N)

    def back(self):
        self.frame.destroy()
        FuncViewController(self.app, 'Welcome')

    def setupUI(self, frame):
        pass

class FuncViewController(ViewController):
    
    def setupUI(self, frame):
        super().setupUI(frame)
        Button(frame, text='Split', width=10, height=2, bg='CornflowerBlue', fg='GhostWhite', font=("Arial, 12"), command=lambda index=0: self.navTo(index)).grid(row=1, column=0, padx=5, pady=10)
        Button(frame, text='Mail', width=10, height=2, bg='CornflowerBlue', fg='GhostWhite', font=("Arial, 12"), command=lambda index=1: self.navTo(index)).grid(row=1, column=1, padx=5)
        
    def navTo(self, index):
        self.frame.destroy()
        if index == 0:
            SplitViewController(self.app, 'Split master report', True)
        else:
            MailViewController(self.app, 'Send Emails to cost center owners', True)

class SplitViewController(ViewController):
    
    def setupUI(self, frame):
        super().setupUI(frame)
        rawDataPath = StringVar()
        savePath = StringVar()
        #raw data
        Label(frame, text='Raw data').grid(row=1, column=0, sticky=W)
        Entry(frame, width=25, textvariable=rawDataPath).grid(row=1, column=1, pady=5, padx=5)
        Button(frame, text='Choose', command=lambda arg=rawDataPath: self.excelSelection(arg)).grid(row=1, column=2)
        #save path
        Label(frame, text='Save path').grid(row=4, column=0, sticky=W)
        Entry(frame, width=25, textvariable=savePath).grid(row=4, column=1, pady=5, padx=5)
        Button(frame, text='Choose', command=lambda arg=savePath: self.dirSelection(arg)).grid(row=4, column=2)
        #Split
        self.splitBtn = Button(frame, width=6, text='Split', bg='lime green', fg='GhostWhite', command=lambda rawData=rawDataPath, save=savePath: self.threadIt(self.split, rawData, save))
        self.splitBtn.grid(row=5, column=2)

    #select file
    def excelSelection(self, path):
        tempStr = askopenfilename(filetypes=[('Excel', '*.xlsx'), ('All Files', '*')])
        path.set(tempStr)

    # select dir
    def dirSelection(self, path):
        tempStr = askdirectory()
        path.set(tempStr)

    #split
    def split(self, rawData, save):
        if len(rawData.get()) * len(save.get()) == 0:
            showinfo(title='Oops', message='Something is missing')
            return
        # self.app.title('Spliting...')
        # self.splitBtn['state'] = 'disabled'
        # self.splitBtn['text'] = 'Spliting...'
        self.isrunning(True)
        try:
            rawdf = pd.DataFrame(pd.read_excel(rawData.get(), sheet_name='Sheet1'))
        except BaseException as e:
            # print(type(e))
            # print(e.values)
            showinfo(title=f'{e.__class__.__name__}', message=f'{e}')
            self.isrunning(False)
            return
        values = ['Asset Number', 
                'Sub-number',
                'Historical Asset Number', 
                'Asset description',
                'Cost Center',
                'Capitalization Date',
                'Original Value',
                'Accumulated Depreciation',
                'Net value',
                ]
        try:
            tempdf = rawdf[values].copy()
        except BaseException as e:
            showinfo(title=f'{e.__class__.__name__}', message=f'{e}')
            self.isrunning(False)
            return
        # tempdf['Asset Number'] = tempdf['Asset Number'].astype('str')
        tempdf['Asset Number'] = tempdf['Asset Number'].map(lambda x: '{:.0f}'.format(x))
        tempdf['Remark/Comment'] = ''
        tempdf['Capitalization Date'] = tempdf['Capitalization Date'].dt.strftime('%m/%d/%Y')
        ccdfs = tempdf.groupby(['Cost Center'])
        # mapdf = tempdf['Cost Center'].drop_duplicates()
        for cc, ccdf in ccdfs:
            # subdf = tempdf.loc[rawdf['Cost Center']==cc]
            # importSub = subdf.copy()
            tempdf = ccdf.append({'Asset Number': 'Total', 'Net value': ccdf['Net value'].sum()}, ignore_index=True)
            # print(tempdf)
            self.saveSub(tempdf, cc, save.get())
            # subdf.to_excel(save.get() + '\\' + str(cc) + '.xlsx', index=False)
        self.isrunning(False)
        # self.splitBtn['state'] = 'normal'
        # self.splitBtn['text'] = 'Split'
        # self.app.title('Spliting ended')
    
    def isrunning(self, state: bool):
        self.app.title('Spliting...' if state else 'Spliting ended')
        self.splitBtn['state'] = 'disabled' if state else 'normal'
        self.splitBtn['text'] = 'Spliting...' if state else 'Split'

    
    def saveSub(self, df, cc, save):
        writer = pd.ExcelWriter(save + '\\' + 'Monthly FA report {:.0f}'.format(cc) + '.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)#, startrow=1, header=False
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1,
        })
        # print(df.head())
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)     
        overall_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
        })
        # pd.io.formats.excel.header_style = None
        worksheet.set_column(0, 0, 15, overall_format)
        worksheet.set_column(1, 1, 11, overall_format)
        worksheet.set_column(2, 2, 11, overall_format)
        worksheet.set_column(3, 3, 50, overall_format)
        worksheet.set_column(4, 8, 12, overall_format)
        worksheet.set_column(9, 9, 16, overall_format)
        # worksheet.set_row(0, None, header_format)
        try:
            writer.save()
        except BaseException as e:
            showinfo(title=f'{e.__class__.__name__}', message=f'{e}')
            self.isrunning(False)
            return

    def threadIt(self, func, *args):
        t = Thread(target=func, args=args)
        t.setDaemon(True)
        t.start()

class MailViewController(ViewController):
    
    def setupUI(self, frame):
        super().setupUI(frame)
        senderStr = StringVar()
        subPath = StringVar()
        toListPath = StringVar()
        copyListPath = StringVar()
        bodyPath = StringVar()
        #sender
        Label(frame, text='Sender:').grid(row=1, column=0, sticky=W)
        Entry(frame, width=25, textvariable=senderStr).grid(row=1, column=1, pady=5, padx=5)
        #cc list
        Label(frame, text='Email to:').grid(row=2, column=0, sticky=W)
        Entry(frame, width=25, textvariable=toListPath).grid(row=2, column=1, pady=5, padx=5)
        Button(frame, text='Choose', command=lambda arg=toListPath: self.excelSelection(arg)).grid(row=2, column=2)  
        #copy list
        Label(frame, text='Copy to:').grid(row=3, column=0, sticky=W)
        Entry(frame, width=25, textvariable=copyListPath).grid(row=3, column=1, pady=5, padx=5)
        Button(frame, text='Choose', command=lambda arg=copyListPath: self.excelSelection(arg)).grid(row=3, column=2)  
        #body
        Label(frame, text='Email body:').grid(row=4, column=0, sticky=W)
        Entry(frame, width=25, textvariable=bodyPath).grid(row=4, column=1, pady=5, padx=5)
        Button(frame, text='Choose', command=lambda arg=bodyPath: self.docSelection(arg)).grid(row=4, column=2)  
         #sub path
        Label(frame, text='Sub path:').grid(row=5, column=0, sticky=W)
        Entry(frame, width=25, textvariable=subPath).grid(row=5, column=1, pady=5, padx=5)
        Button(frame, text='Choose', command=lambda arg=subPath: self.dirSelection(arg)).grid(row=5, column=2)
        #Check
        Button(frame, width=6, text='Check', bg='lime green', fg='GhostWhite', command=lambda cc=toListPath, sub=subPath: self.threadIt(self.checkFiles, cc, sub)).grid(row=6, column=1, sticky='E')
        #Mail
        self.mailBtn = Button(frame, width=6, text='Mail', bg='lime green', fg='GhostWhite', command=lambda sender=senderStr, to=toListPath, copy=copyListPath, body=bodyPath, sub=subPath: self.threadIt(self.mail, sender, to, copy, body, sub))
        self.mailBtn.grid(row=6, column=2)

    #select file
    def docSelection(self, path):
        tempStr = askopenfilename(filetypes=[('Web Page', '*.htm'), ('All Files', '*')], title = "Select Email body")
        path.set(tempStr)

    def excelSelection(self, path):
        tempStr = askopenfilename(filetypes=[('Excel', '*.xlsx'), ('All Files', '*')], title = "Select FA coordinator list")
        path.set(tempStr)

    # select dir
    def dirSelection(self, path):
        tempStr = askdirectory(title = "Select sub files directory")
        path.set(tempStr)

    #check
    def checkFiles(self, cc, sub):
        if len(cc.get()) * len(sub.get()) == 0:
            showinfo(title='Oops', message='Something is missing')
            return
        self.app.title('Checking...')
        try: 
            ccdf = pd.DataFrame(pd.read_excel(cc.get(), sheet_name='Sheet1'))
            tempdf = ccdf['Cost Center'].copy().astype('str')
        except BaseException as e:
            showinfo(title=f'{e.__class__.__name__}', message=f'{e}')
            self.app.title('Check Finished...')
            return
        ccList = tempdf.drop_duplicates().values.tolist()
        tempList = listdir(sub.get())
        # fileFullList = list(map(self.getFiles, tempList)) 
        # showinfo(title='Oops', message='No sub files') if len(fileFullList) ==
        try:
            fileList = list(map(self.getFiles, tempList))
        except BaseException as e:
            showinfo(title=f'{e.__class__.__name__}', message=f'{e}')
            self.app.title('Check Finished...')
            return
        fileDifccSet = set(fileList) - set(ccList)
        # difcc = set(ccList) - set(fileList)
        fileDifccStr = ','.join(fileDifccSet)
        # difStrcc = ','.join(difcc)
        if len(fileDifccStr) != 0:
            strFile = 'Sub file(s) of Cost Center(s): {} cannot find responding info in CC&FA mapping list.'.format(fileDifccStr)
            showinfo(title='Oops', message='{0}'.format(strFile))
        else:
            showinfo(title='Oops', message='Sub file(s) of Cost Center(s) totally matches CC&FA mapping list.')
        self.app.title('Check Finished...')
        # else:
        #     strFile = ''
        # # if len(difStrFile) != 0:
        # #     strcc = 'Cost Center(s): {} does not have matching sub files'.format(difStrcc)
        # # else:
        # #     strcc = ''
        # if len(strFile) != 0 or len(strcc) != 0:
        #     showinfo(title='Oops', message='{0}'.format(strFile))
        # self.app.title('Check Finished...')

    def isrunning(self, state: bool):
        self.app.title('Sending...' if state else 'Send Finished.')
        self.mailBtn['state'] = 'disabled' if state else 'normal'
        self.mailBtn['text'] = 'Sending...' if state else 'Split'

    #mail
    def mail(self, sender, to, copy, body, sub):
        if len(sender.get()) * len(to.get()) * len(body.get()) * len(sub.get()) == 0:
            showinfo(title='Oops', message='Something is missing')
            return
        # self.app.title('Sending...')
        # self.mailBtn['state'] = 'disabled'
        # self.mailBtn['text'] = 'Sending...'
        self.isrunning(True)
        try:
            todf = pd.DataFrame(pd.read_excel(to.get(), sheet_name='Sheet2'))
            # if copy.get() != '':
            copydf = pd.DataFrame(pd.read_excel(copy.get(), sheet_name='Sheet2')) if copy.get() != '' else None
        except BaseException as e:
            showinfo(title=f'{e.__class__.__name__}', message=f'{e}')
            return
        self.sendMail(sender, todf, copydf, body, sub)
        self.isrunning(False)
        # self.app.title('Send Finished...')
        # self.mailBtn['state'] = 'normal'
        # self.mailBtn['text'] = 'Mail'

    def sendMail(self, sender, todf, copydf, body, sub):
        smtp_server = 'rb-smtp-int.bosch.com'
        port = 25
        senderStr = sender.get()
        if not self.checkEmail(senderStr):
            showinfo(title='Format Error', message='Please check Sender email.')
            return
        # if not re.match(r'^[a-zA-Z0-9_.-]+@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z0-9]{2,6}$',senderStr):
        #     showinfo(title='Format Error', message='Please check Sender email.')
        #     return
        #email body
        # paras = ''
        # doc = docx.Document(body.get())
        # for para in doc.paragraphs:
        #     temp = '<p>{}</p>'.format(para.text)
        #     paras = paras + temp
        with open(body.get()) as f:
            paras = f.read()
        # print(paras)
        errorList = []
        succList = []
        for index, row in todf.iterrows():
            # if not(re.match(r'^[a-zA-Z0-9_.-]+@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z0-9]{2,6}$',row[1])):
            #     errorList.append(str(index + 1) + '. ' + row[1])
            #     continue
            if not self.checkEmail(row[1]):
                errorList.append(str(index + 1) + '. ' + row[1])
                continue
            
            receiver = f'{row[1]}'
            msg = MIMEMultipart()
            msg['Subject'] = '{} FA Report'.format(row[0])
            # msg['From'] = 'ye.xiang@cn.bosch.com'
            msg['To'] = receiver
            if row[0] in copydf['Cost Center']:
                msg['Copy'] = copydf[]
            # head = '<b>Dear {},</b>'.format(row[2])
            bodyStr = copy.copy(paras)
            bodyStr = bodyStr.replace('{0}', f'{row[2]}')
            # print(row[2])
            # return
            # body = paras.format(receiver)
            msg.attach(MIMEText(bodyStr, 'html', 'utf-8'))
            filePath = r'{0}\Monthly FA report {1}.xlsx'.format(sub.get(), row[0])
            if not isfile(filePath):
                continue
            attach = MIMEText(open(filePath, 'rb').read(), 'base64', 'utf-8')
            attach['Content-Disposition'] = 'attachment; filename="Monthly FA report {}.xlsx"'.format(row[0])
            msg.attach(attach)
            try:
                server = smtplib.SMTP(smtp_server, port)
                server.sendmail(senderStr, receiver, msg.as_string())
                server.quit()
                succList.append(str(index + 1) + '. ' + row[1])
            except Exception as e:
                errorList.append(str(index + 1) + '. ' + row[1])
                continue
        self.saveLog(succList, errorList)
        
    def checkEmail(self, emailStr):
        return match(r'^[a-zA-Z0-9_.-]+@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z0-9]{2,6}$',emailStr)


    def saveLog(self, succList, errorList):
        now = datetime.datetime.now()
        fileName = 'Log_file_{}'.format(now.strftime('%Y%m%d'))
        logPath = asksaveasfilename(initialfile=fileName, defaultextension=".xlsx", title = "Save log file to...",filetypes = (('Excel', '*.xlsx'),("all files","*.*")))
        if not logPath: return
        writer = pd.ExcelWriter(logPath)
        succ_df = pd.DataFrame({'Send': succList})
        succ_df.to_excel(writer, 'Send', index=False)
        err_df = pd.DataFrame({'Not Send': errorList})
        err_df.to_excel(writer, 'Not Send', index=False)
        try:
            writer.save()
        except BaseException as e:
            showinfo(title=f'{e.__class__.__name__}', message=f'{e}')
            return

    def getFiles(self, file):
        name = os.path.splitext(file)[0][-6:]
        return name


    def threadIt(self, func, *args):
        t = Thread(target=func, args=args)
        t.setDaemon(True)
        t.start()

app = Tk()
FuncViewController(app, 'Please select function')
app.mainloop()
