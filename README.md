import sys
import tkinter.messagebox
from tkinter import filedialog
import pandas as pd
import xlwings as xl
import xlrd
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import (QWidget, QComboBox)
from EDIModel import *
qtCreatorFile = "EDIModel.ui"  # Enter file here.
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)
class MyApp(QtWidgets.QMainWindow,Ui_MainWindow,QWidget):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.pushButton_StartModel.clicked.connect(self.dataModelEDI)
        self.pushButton_importALLDATA.clicked.connect(self.import_ALLDATA)
        self.pushButton_importCatA.clicked.connect(self.import_CatA)
        self.pushButton_importValidation.clicked.connect(self.import_Validation)

    def import_ALLDATA(self):
        filename1 = filedialog.askopenfilename()
        self.textEdit_BossAllPath.setText(filename1)
    def import_CatA(self):
        filename1 = filedialog.askopenfilename()
        self.textEdit_BossApath.setText(filename1)
    def import_Validation(self):
        filename1 = filedialog.askopenfilename()
        self.textEdit_ValidationPath.setText(filename1)
    def dataModelEDI(self):
        try:


            if self.textEdit_BossAllPath.toPlainText() == ""or self.textEdit_BossApath.toPlainText() == ""or self.textEdit_ValidationPath.toPlainText() == "":
                tkinter.messagebox.showinfo("AZ提示", "please select Datasource firstly")
            if self.comboBox_Site.currentText() == "" or self.comboBox_year_2.currentText() == "" or self.comboBox_month_2.currentText() == "" or self.comboBox_date_2.currentText() == "":
                tkinter.messagebox.showinfo("AZ提示", "please select Site, Year, Month, Date carefully")
            else:
                year = self.comboBox_year_2.currentText()
                month = self.comboBox_month_2.currentText()
                date = self.comboBox_date_2.currentText()
                data = year + "/" + month + "/" + date

                if self.comboBox_Site.currentText()=="SZX":
                    self.progressBar.setValue(0)
                    detail = pd.read_excel(self.textEdit_BossAllPath.toPlainText())
                    #print(detail.columns)

                    epusia = detail.loc[
                        (detail["Stts Cd"] == "EP") | (detail["Stts Cd"] == "US") | (detail["Stts Cd"] == "IA"), [
                            "Hawb", "Stts Cd", "Input User"]]
                    epusia = epusia.drop_duplicates(subset=["Hawb"])
                    epusiaVolume = epusia["Input User"].value_counts()
                    indexA = list(epusiaVolume.index)
                    indexB = list(epusiaVolume)
                    epusiaDf = pd.DataFrame()
                    epusiaDf["Name"] = indexA
                    epusiaDf["TTL Shpts"] = indexB

                    pg = detail.loc[detail["Stts Cd"] == "PG", ["Hawb", "Stts Cd", "Input User"]]
                    pg = pg.drop_duplicates(subset=["Hawb"])
                    pgVolume = pg["Input User"].value_counts()
                    indexA = list(pgVolume.index)
                    indexB = list(pgVolume)
                    pgDf = pd.DataFrame()
                    pgDf["Name"] = indexA
                    pgDf["Personal Goods"] = indexB

                    A = pd.read_excel(self.textEdit_BossApath.toPlainText())
                    dataMerge = pd.merge(detail, A, left_on="Hawb", right_on="Hawb")
                    #dataMerge = dataMerge.sort_values(by='SN', axis=0, ascending=True)  # avoid input user get together
                    # dataMerge.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx",index=True)
                    CatA = dataMerge.loc[(dataMerge["Stts Cd"] == "EP") | (dataMerge["Stts Cd"] == "US") | (
                                dataMerge["Stts Cd"] == "IA"), ["Hawb", "Stts Cd", "Input User"]]
                    data2 = CatA.drop_duplicates(subset=["Hawb"])
                    CatAVolume = data2["Input User"].value_counts()
                    indexA = list(CatAVolume.index)
                    indexB = list(CatAVolume)
                    catADf = pd.DataFrame()
                    catADf["Name"] = indexA
                    catADf["Cat. A"] = indexB
                    merge1 = pd.merge(epusiaDf, pgDf, how="outer", left_on="Name", right_on="Name", )
                    merge2 = pd.merge(merge1, catADf, how="outer", left_on="Name", right_on="Name")

                    merge2["InputDate"] = data
                    merge2["ImportSite"] = "SZX"
                    merge2 = merge2.loc[:, ["InputDate", "ImportSite", "Name", "TTL Shpts", "Personal Goods", "Cat. A"]]
                    xtotal = merge2.values
                    #merge2.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx", index=True)
                    ####################
                    self.progressBar.setValue(50)
                    fz = detail.loc[detail["Stts Cd"] == "FZ", ["Hawb", "Stts Cd", "Input User"]]
                    fz = fz.drop_duplicates(subset=["Hawb"], keep="last")  # drop duplicates of Hawb
                    fzVolume = fz["Input User"].value_counts()
                    index1 = list(fzVolume.index)
                    index2 = list(fzVolume)
                    translateDf = pd.DataFrame()  # creat a new df
                    translateDf["Input User"] = index1
                    translateDf["HawbOfCount"] = index2
                    translateDf["Type"] = "Translation"
                    translateDf["InputDate"] = data
                    translateDf["ImportSite"] = "SZX"
                    translateDf = translateDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]
                    # print(translateDf)

                    fm = detail.loc[detail["Stts Cd"] == "FM", ["Hawb", "Stts Cd", "Input User"]]
                    fm = fm.drop_duplicates(subset=["Hawb"], keep="last")
                    fmVolume = fm["Input User"].value_counts()
                    index3 = list(fmVolume.index)
                    index4 = list(fmVolume)
                    fmDf = pd.DataFrame()
                    fmDf["Input User"] = index3
                    fmDf["HawbOfCount"] = index4
                    fmDf["Type"] = "AssignFM"
                    fmDf["InputDate"] = data
                    fmDf["ImportSite"] = "SZX"
                    fmDf = fmDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]
                    final_BossDataDf = translateDf.append(fmDf)
                    x = final_BossDataDf.values
                    self.progressBar.setValue(99)

                    app = xl.App(visible=True, add_book=False)
                    app.display_alerts = False
                    app.screen_updating = True
                    wb1 = app.books.open(self.textEdit_ValidationPath.toPlainText())
                    wb = xlrd.open_workbook(self.textEdit_ValidationPath.toPlainText())
                    table1 = wb.sheet_by_name("TTLShpt")
                    sht1 = wb1.sheets["TTLShpt"]
                    sht1.activate()
                    n = 0
                    for t in range(5, 91):
                        hhh = table1.cell_value(t, 4)
                        if hhh != "":
                            n = n + 1
                    xx = 6 + n
                    wb1.sheets["TTLShpt"].range("E%d" % xx).value = xtotal

                    table2 = wb.sheet_by_name("BossData")
                    sht2 = wb1.sheets["BossData"]
                    sht2.activate()
                    n = 0
                    for t in range(5, 91):
                        hhh = table2.cell_value(t, 5)
                        if hhh != "":
                            n = n + 1

                    xx = 6 + n
                    wb1.sheets["BossData"].range("f%d" % xx).value = x
                    self.progressBar.setValue(100)
                    wb1.save()
                    wb1.close()
                    app.quit()
                    tkinter.messagebox._show("AZ 提示：","Program for SZX completed successfully")

                if self.comboBox_Site.currentText()=="PVG":
                    self.progressBar.setValue(50)
                    namedetail = pd.DataFrame({"Name": ['WANG LEI', 'WANGWEI', 'XU DAHONG', 'ZHANG LIANG',
                                                        'ZHAOZHILIANG', 'Zheng Yaojie', '徐凤巧', '张徽岭', '张勇']},
                                              columns=["Name"])

                    detail = pd.read_excel(self.textEdit_BossAllPath.toPlainText())
                    hfcsdetail=detail
                    detail = pd.merge(detail, namedetail, left_on="Input User", right_on="Name")
                    detail=detail.sort_values(by = 'Input Dtm',axis = 0,ascending = True)#avoid input user get together
                    #detail = detail.drop(labels="Name", axis=1)
                    #####
                    epusia = detail.loc[
                        (detail["Stts Cd"] == "EP") | (detail["Stts Cd"] == "US") | (detail["Stts Cd"] == "IA"), [
                            "Hawb", "Stts Cd", "Input User"]]
                    epusia = epusia.drop_duplicates(subset=["Hawb"])
                    epusiaVolume = epusia["Input User"].value_counts()
                    indexA = list(epusiaVolume.index)
                    indexB = list(epusiaVolume)
                    epusiaDf = pd.DataFrame()
                    epusiaDf["Name"] = indexA
                    epusiaDf["TTL Shpts"] = indexB

                    pg = detail.loc[detail["Stts Cd"] == "PG", ["Hawb", "Stts Cd", "Input User"]]
                    pg = pg.drop_duplicates(subset=["Hawb"])
                    pgVolume = pg["Input User"].value_counts()
                    indexA = list(pgVolume.index)
                    indexB = list(pgVolume)
                    pgDf = pd.DataFrame()
                    pgDf["Name"] = indexA
                    pgDf["Personal Goods"] = indexB

                    A = pd.read_excel(self.textEdit_BossApath.toPlainText())
                    dataMerge = pd.merge(detail, A, left_on="Hawb", right_on="Hawb")
                    #dataMerge = dataMerge.sort_values(by='SN', axis=0, ascending=True)
                    CatA = dataMerge.loc[(dataMerge["Stts Cd"] == "EP") | (dataMerge["Stts Cd"] == "US") | (
                                dataMerge["Stts Cd"] == "IA"), ["Hawb", "Stts Cd", "Input User"]]
                    data2 = CatA.drop_duplicates(subset=["Hawb"])
                    CatAVolume = data2["Input User"].value_counts()
                    indexA = list(CatAVolume.index)
                    indexB = list(CatAVolume)
                    catADf = pd.DataFrame()
                    catADf["Name"] = indexA
                    catADf["Cat. A"] = indexB
                    merge1 = pd.merge(epusiaDf, pgDf, how="outer", left_on="Name", right_on="Name", )
                    merge2 = pd.merge(merge1, catADf, how="outer", left_on="Name", right_on="Name")

                    merge2["InputDate"] = data
                    merge2["ImportSite"] = "PVG"
                    merge2 = merge2.loc[:, ["InputDate", "ImportSite", "Name", "TTL Shpts", "Personal Goods", "Cat. A"]]
                    xtotal = merge2.values
                    # merge2.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx",index=True)
                    ####

                    fz = detail.loc[detail["Stts Cd"] == "FZ", ["Hawb", "Stts Cd", "Input User"]]
                    fz = fz.drop_duplicates(subset=["Hawb"], keep="last")  # drop duplicates of Hawb
                    fzVolume = fz["Input User"].value_counts()
                    index1 = list(fzVolume.index)
                    index2 = list(fzVolume)
                    translateDf = pd.DataFrame()  # creat a new df
                    translateDf["Input User"] = index1
                    translateDf["HawbOfCount"] = index2
                    translateDf["Type"] = "Translation"
                    translateDf["InputDate"] = data
                    translateDf["ImportSite"] = "PVG"
                    translateDf = translateDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fm_kj3d = detail.loc[
                        (detail["Stts Cd"] == "FM") & (detail["REM"] == "kj3d") | (detail["REM"] == "KJ3D"), ["Hawb",
                                                                                                              "Stts Cd",
                                                                                                              "Input User",
                                                                                                              "REM"]]
                    fm_kj3d = fm_kj3d.drop_duplicates(subset=["Hawb"])
                    assignKJ3Volume = fm_kj3d["Input User"].value_counts()
                    index1 = list(assignKJ3Volume.index)
                    index2 = list(assignKJ3Volume)
                    assignKJ3DF = pd.DataFrame()
                    assignKJ3DF["Input User"] = index1
                    assignKJ3DF["HawbOfCount"] = index2
                    assignKJ3DF["Type"] = "AssignKJ3"
                    assignKJ3DF["InputDate"] = data
                    assignKJ3DF["ImportSite"] = "PVG"
                    assignKJ3DF = assignKJ3DF.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fm_kj3z = detail.loc[
                        (detail["Stts Cd"] == "FM") & (detail["REM"] == "KJ3Z") | (detail["REM"] == "kj3z"), ["Hawb",
                                                                                                              "Stts Cd",
                                                                                                              "Input User",
                                                                                                              "REM"]]
                    fm_kj3z = fm_kj3z.drop_duplicates(subset=["Hawb"])
                    assignKJ3ZVolume = fm_kj3z["Input User"].value_counts()
                    index1 = list(assignKJ3ZVolume.index)
                    index2 = list(assignKJ3ZVolume)
                    AssignKJ3ZDF = pd.DataFrame()
                    AssignKJ3ZDF["Input User"] = index1
                    AssignKJ3ZDF["HawbOfCount"] = index2
                    AssignKJ3ZDF["Type"] = "AssignKJ3Z"
                    AssignKJ3ZDF["InputDate"] = data
                    AssignKJ3ZDF["ImportSite"] = "PVG"
                    AssignKJ3ZDF = AssignKJ3ZDF.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fm = detail.loc[detail["Stts Cd"] == "FM", ["Hawb", "Stts Cd", "Input User"]]
                    fm = fm.drop_duplicates(subset=["Hawb"], keep="first")
                    fmVolume = fm["Input User"].value_counts()
                    index3 = list(fmVolume.index)
                    index4 = list(fmVolume)
                    fmDf = pd.DataFrame()
                    fmDf["Input User"] = index3
                    fmDf["HawbOfCount"] = index4
                    fmDf["Type"] = "AssignFM"
                    fmDf["InputDate"] = data
                    fmDf["ImportSite"] = "PVG"
                    fmDf = fmDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fmHFEDI = detail.loc[detail["Stts Cd"] == "FM", ["Hawb", "Stts Cd", "Input User"]]
                    fmHFEDI = fmHFEDI.drop_duplicates(subset="Hawb")
                    csNameDetail = pd.DataFrame({"Name": ['曹宁宁', '曹萍萍', '陈凯杰', 'CHEN MIN', 'CHEN NA', '陈田', '丁冬阳',
                                                          '高玲', '郜琼', 'GE PANPAN', '龚珺', '黄德成', '蒋成理', '李洁',
                                                          'Li Qianqian', '刘云', '陆文娟', '马具玲',
                                                          '钱丽娟', 'QIAN PINGPING', 'QIAO LEILEI', '秦敏', 'SU HUI',
                                                          'SUN MIN', 'TANG HULAN', 'WANG LIMING',
                                                          '王尚', '王桐', 'WANG YAWEN', 'WU LUAN QING', 'WU YIFAN', '谢梦娟',
                                                          'XIONG XIAOMIN', 'XU CAIQING',
                                                          'YANG LI', 'YIN SHANSHAN', '袁兰兰', 'ZHANG LING', 'ZHANGYI',
                                                          '张亚芹', 'ZHONG LIANG YUN',
                                                          'ZHU TINGTING', '朱晓雯']}, columns=["Name"])

                    hfcsdetail = pd.merge(hfcsdetail, csNameDetail, left_on="Input User", right_on="Name")
                    hfcsdetail = hfcsdetail.loc[:, ["Hawb", "Stts Cd", "Input User"]]
                    hfcsdetail = hfcsdetail.drop_duplicates(subset="Hawb")
                    assignHfcs = pd.merge(fmHFEDI, hfcsdetail, left_on="Hawb", right_on="Hawb")
                    # assignHfcs=assignHfcs.drop_duplicates(subset="Hawb")
                    assignHfcsVolume = assignHfcs["Input User_x"].value_counts()
                    indexA = list(assignHfcsVolume.index)
                    indexB = list(assignHfcsVolume)
                    assignHfcsDf = pd.DataFrame()
                    assignHfcsDf["Input User"] = indexA
                    assignHfcsDf["HawbOfCount"] = indexB
                    assignHfcsDf["Type"] = "AssignHFECS"
                    assignHfcsDf["InputDate"] = data
                    assignHfcsDf["ImportSite"] = "PVG"
                    assignHfcsDf = assignHfcsDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    tz = detail.loc[detail["Stts Cd"] == "TZ", ["Hawb", "Stts Cd", "Input User"]]
                    tz = tz.drop_duplicates(subset=["Hawb"], keep="first")
                    tzVolume = tz["Input User"].value_counts()
                    index3 = list(tzVolume.index)
                    index4 = list(tzVolume)
                    tzDf = pd.DataFrame()
                    tzDf["Input User"] = index3
                    tzDf["HawbOfCount"] = index4
                    tzDf["Type"] = "SendArrivalNotice"
                    tzDf["InputDate"] = data
                    tzDf["ImportSite"] = "PVG"
                    tzDf = tzDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    spt33 = detail.loc[detail["Stts Cd"] == "33", ["Hawb", "Stts Cd", "Input User"]]
                    spt33 = spt33.drop_duplicates(subset=["Hawb"])
                    spt33Volume = spt33["Input User"].value_counts()
                    indexA = list(spt33Volume.index)
                    indexB = list(spt33Volume)
                    kj2Df = pd.DataFrame()
                    kj2Df["Input User"] = indexA
                    kj2Df["HawbOfCount"] = indexB
                    kj2Df["Type"] = "AssignKJ2"
                    kj2Df["InputDate"] = data
                    kj2Df["ImportSite"] = "PVG"
                    kj2Df = kj2Df.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    f0 = detail.loc[detail["Stts Cd"] == "F0", ["Hawb", "Stts Cd", "Input User"]]
                    f0 = f0.drop_duplicates(subset=["Hawb"])
                    #f0.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx", index=True)
                    f0Volume = f0["Input User"].value_counts()
                    indexA = list(f0Volume.index)
                    indexB = list(f0Volume)
                    f0Df = pd.DataFrame()
                    f0Df["Input User"] = indexA
                    f0Df["HawbOfCount"] = indexB
                    f0Df["Type"] = "ReturnShptReassign"
                    f0Df["InputDate"] = data
                    f0Df["ImportSite"] = "PVG"
                    f0Df = f0Df.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    x = translateDf.append([fmDf, assignHfcsDf, assignKJ3DF, AssignKJ3ZDF, kj2Df, tzDf, f0Df])
                    x = x.values  # only by .values
                    # print(x)
                    # print(type(x))
                    # print(x[0][1])
                    self.progressBar.setValue(99)
                    app = xl.App(visible=True, add_book=False)
                    app.display_alerts = False
                    app.screen_updating = True
                    wb1 = app.books.open(self.textEdit_ValidationPath.toPlainText())
                    wb = xlrd.open_workbook(self.textEdit_ValidationPath.toPlainText())
                    table1 = wb.sheet_by_name("TTLShpt")
                    sht1 = wb1.sheets["TTLShpt"]
                    sht1.activate()
                    n = 0
                    for t in range(5, 91):
                        hhh = table1.cell_value(t, 4)
                        if hhh != "":
                            n = n + 1
                    xx = 6 + n
                    wb1.sheets["TTLShpt"].range("E%d" % xx).value = xtotal

                    table2 = wb.sheet_by_name("BossData")
                    sht = wb1.sheets["BossData"]
                    sht.activate()
                    n = 0
                    for t in range(5, 91):
                        hhh = table2.cell_value(t, 5)
                        if hhh != "":
                            n = n + 1
                    # print(n)
                    # print(final_BossDataDf)
                    xx = 6 + n
                    wb1.sheets["BossData"].range("f%d" % xx).value = x
                    self.progressBar.setValue(100)
                    # spt33.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx",index=True)
                    wb1.save()
                    wb1.close()
                    app.quit()
                    tkinter.messagebox._show("AZ 提示：", "Program for PVG completed successfully")

                if self.comboBox_Site.currentText()=="BJS":
                    self.progressBar.setValue(50)
                    namedetail = pd.DataFrame({"Name": ['陆勤康', '王蕾', '张徽岭', '赵红梅','周召']},
                                              columns=["Name"])

                    detail = pd.read_excel(self.textEdit_BossAllPath.toPlainText(),header=6)
                    hfcsdetail=detail
                    detail = pd.merge(detail, namedetail, left_on="Input User", right_on="Name")
                    detail=detail.sort_values(by = 'Input Dtm',axis = 0,ascending = True)#avoid input user get together
                    #detail.to_excel("C:\\Users\\alanz\\Desktop\\bjs\\hh.xlsx")
                    #detail = detail.drop(labels="Name", axis=1)
                    #####
                    DOD1D1D3 = detail.loc[
                        (detail["Stts Cd"] == "D0") | (detail["Stts Cd"] == "D1") | (detail["Stts Cd"] == "D2")| (detail["Stts Cd"] == "D3"), [
                            "Hawb", "Stts Cd", "Input User"]]
                    DOD1D1D3 = DOD1D1D3.drop_duplicates(subset=["Hawb"])
                    DOD1D1D3Volume = DOD1D1D3["Input User"].value_counts()
                    indexA = list(DOD1D1D3Volume.index)
                    indexB = list(DOD1D1D3Volume)
                    DOD1D1D3Df = pd.DataFrame()
                    DOD1D1D3Df["Name"] = indexA
                    DOD1D1D3Df["TTL Shpts"] = indexB

                    pg = detail.loc[detail["Stts Cd"] == "PG", ["Hawb", "Stts Cd", "Input User"]]
                    pg = pg.drop_duplicates(subset=["Hawb"])
                    pgVolume = pg["Input User"].value_counts()
                    indexA = list(pgVolume.index)
                    indexB = list(pgVolume)
                    pgDf = pd.DataFrame()
                    pgDf["Name"] = indexA
                    pgDf["Personal Goods"] = indexB

                    A = pd.read_excel(self.textEdit_BossApath.toPlainText())
                    dataMerge = pd.merge(detail, A, left_on="Hawb", right_on="Hawb")
                    #dataMerge = dataMerge.sort_values(by='SN', axis=0, ascending=True)
                    CatA = dataMerge.loc[(dataMerge["Stts Cd"] == "D0") | (dataMerge["Stts Cd"] == "D1") | (
                                dataMerge["Stts Cd"] == "D2")| (dataMerge["Stts Cd"] == "D3"), ["Hawb", "Stts Cd", "Input User"]]
                    data2 = CatA.drop_duplicates(subset=["Hawb"])
                    CatAVolume = data2["Input User"].value_counts()
                    indexA = list(CatAVolume.index)
                    indexB = list(CatAVolume)
                    catADf = pd.DataFrame()
                    catADf["Name"] = indexA
                    catADf["Cat. A"] = indexB
                    merge1 = pd.merge(DOD1D1D3Df, pgDf, how="outer", left_on="Name", right_on="Name", )
                    merge2 = pd.merge(merge1, catADf, how="outer", left_on="Name", right_on="Name")

                    merge2["InputDate"] = data
                    merge2["ImportSite"] = "BJS"
                    merge2 = merge2.loc[:, ["InputDate", "ImportSite", "Name", "TTL Shpts", "Personal Goods", "Cat. A"]]
                    xtotal = merge2.values
                    # merge2.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx",index=True)
                    ####

                    fz = detail.loc[detail["Stts Cd"] == "FZ", ["Hawb", "Stts Cd", "Input User"]]
                    fz = fz.drop_duplicates(subset=["Hawb"], keep="last")  # drop duplicates of Hawb
                    fzVolume = fz["Input User"].value_counts()
                    index1 = list(fzVolume.index)
                    index2 = list(fzVolume)
                    translateDf = pd.DataFrame()  # creat a new df
                    translateDf["Input User"] = index1
                    translateDf["HawbOfCount"] = index2
                    translateDf["Type"] = "Translation"
                    translateDf["InputDate"] = data
                    translateDf["ImportSite"] = "BJS"
                    translateDf = translateDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fm_kj3d = detail.loc[
                        (detail["Stts Cd"] == "FM") & (detail["REM"] == "kj3d") | (detail["REM"] == "KJ3D"), ["Hawb",
                                                                                                              "Stts Cd",
                                                                                                              "Input User",
                                                                                                              "REM"]]
                    fm_kj3d = fm_kj3d.drop_duplicates(subset=["Hawb"])
                    assignKJ3Volume = fm_kj3d["Input User"].value_counts()
                    index1 = list(assignKJ3Volume.index)
                    index2 = list(assignKJ3Volume)
                    assignKJ3DF = pd.DataFrame()
                    assignKJ3DF["Input User"] = index1
                    assignKJ3DF["HawbOfCount"] = index2
                    assignKJ3DF["Type"] = "AssignKJ3"
                    assignKJ3DF["InputDate"] = data
                    assignKJ3DF["ImportSite"] = "BJS"
                    assignKJ3DF = assignKJ3DF.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fm_kj3z = detail.loc[
                        (detail["Stts Cd"] == "FM") & (detail["REM"] == "KJ3Z") | (detail["REM"] == "kj3z"), ["Hawb",
                                                                                                              "Stts Cd",
                                                                                                              "Input User",
                                                                                                              "REM"]]
                    fm_kj3z = fm_kj3z.drop_duplicates(subset=["Hawb"])
                    assignKJ3ZVolume = fm_kj3z["Input User"].value_counts()
                    index1 = list(assignKJ3ZVolume.index)
                    index2 = list(assignKJ3ZVolume)
                    AssignKJ3ZDF = pd.DataFrame()
                    AssignKJ3ZDF["Input User"] = index1
                    AssignKJ3ZDF["HawbOfCount"] = index2
                    AssignKJ3ZDF["Type"] = "AssignKJ3Z"
                    AssignKJ3ZDF["InputDate"] = data
                    AssignKJ3ZDF["ImportSite"] = "BJS"
                    AssignKJ3ZDF = AssignKJ3ZDF.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fm = detail.loc[detail["Stts Cd"] == "FM", ["Hawb", "Stts Cd", "Input User"]]
                    fm = fm.drop_duplicates(subset=["Hawb"], keep="first")
                    fmVolume = fm["Input User"].value_counts()
                    index3 = list(fmVolume.index)
                    index4 = list(fmVolume)
                    fmDf = pd.DataFrame()
                    fmDf["Input User"] = index3
                    fmDf["HawbOfCount"] = index4
                    fmDf["Type"] = "AssignFM"
                    fmDf["InputDate"] = data
                    fmDf["ImportSite"] = "BJS"
                    fmDf = fmDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    fmHFEDI = detail.loc[detail["Stts Cd"] == "FM", ["Hawb", "Stts Cd", "Input User"]]
                    fmHFEDI = fmHFEDI.drop_duplicates(subset="Hawb")
                    csNameDetail = pd.DataFrame({"Name": ['何亚君', '李梅', '黄德成', '黄晴晴', '江嫚', '李洁', '刘云',
                                                          '乔蕾蕾', '尚瑜', '谢素文', '徐盈盈', '袁兰兰']}, columns=["Name"])

                    hfcsdetail = pd.merge(hfcsdetail, csNameDetail, left_on="Input User", right_on="Name")
                    hfcsdetail = hfcsdetail.loc[:, ["Hawb", "Stts Cd", "Input User"]]
                    hfcsdetail = hfcsdetail.drop_duplicates(subset="Hawb")
                    assignHfcs = pd.merge(fmHFEDI, hfcsdetail, left_on="Hawb", right_on="Hawb")
                    # assignHfcs=assignHfcs.drop_duplicates(subset="Hawb")
                    assignHfcsVolume = assignHfcs["Input User_x"].value_counts()
                    indexA = list(assignHfcsVolume.index)
                    indexB = list(assignHfcsVolume)
                    assignHfcsDf = pd.DataFrame()
                    assignHfcsDf["Input User"] = indexA
                    assignHfcsDf["HawbOfCount"] = indexB
                    assignHfcsDf["Type"] = "AssignHFECS"
                    assignHfcsDf["InputDate"] = data
                    assignHfcsDf["ImportSite"] = "BJS"
                    assignHfcsDf = assignHfcsDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    tz = detail.loc[detail["Stts Cd"] == "TZ", ["Hawb", "Stts Cd", "Input User"]]
                    tz = tz.drop_duplicates(subset=["Hawb"], keep="first")
                    tzVolume = tz["Input User"].value_counts()
                    index3 = list(tzVolume.index)
                    index4 = list(tzVolume)
                    tzDf = pd.DataFrame()
                    tzDf["Input User"] = index3
                    tzDf["HawbOfCount"] = index4
                    tzDf["Type"] = "SendArrivalNotice"
                    tzDf["InputDate"] = data
                    tzDf["ImportSite"] = "BJS"
                    tzDf = tzDf.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    spt33 = detail.loc[detail["Stts Cd"] == "33", ["Hawb", "Stts Cd", "Input User"]]
                    spt33 = spt33.drop_duplicates(subset=["Hawb"])
                    spt33Volume = spt33["Input User"].value_counts()
                    indexA = list(spt33Volume.index)
                    indexB = list(spt33Volume)
                    kj2Df = pd.DataFrame()
                    kj2Df["Input User"] = indexA
                    kj2Df["HawbOfCount"] = indexB
                    kj2Df["Type"] = "AssignKJ2"
                    kj2Df["InputDate"] = data
                    kj2Df["ImportSite"] = "BJS"
                    kj2Df = kj2Df.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    f0 = detail.loc[detail["Stts Cd"] == "F0", ["Hawb", "Stts Cd", "Input User"]]
                    f0 = f0.drop_duplicates(subset=["Hawb"])
                    #f0.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx", index=True)
                    f0Volume = f0["Input User"].value_counts()
                    indexA = list(f0Volume.index)
                    indexB = list(f0Volume)
                    f0Df = pd.DataFrame()
                    f0Df["Input User"] = indexA
                    f0Df["HawbOfCount"] = indexB
                    f0Df["Type"] = "ReturnShptReassign"
                    f0Df["InputDate"] = data
                    f0Df["ImportSite"] = "BJS"
                    f0Df = f0Df.loc[:, ["InputDate", "ImportSite", "Input User", "HawbOfCount", "Type"]]

                    x = translateDf.append([fmDf, assignHfcsDf, assignKJ3DF, AssignKJ3ZDF, kj2Df, tzDf, f0Df])
                    x = x.values  # only by .values
                    # print(x)
                    # print(type(x))
                    # print(x[0][1])
                    self.progressBar.setValue(99)
                    app = xl.App(visible=True, add_book=False)
                    app.display_alerts = False
                    app.screen_updating = True
                    wb1 = app.books.open(self.textEdit_ValidationPath.toPlainText())
                    wb = xlrd.open_workbook(self.textEdit_ValidationPath.toPlainText())
                    table1 = wb.sheet_by_name("TTLShpt")
                    sht1 = wb1.sheets["TTLShpt"]
                    sht1.activate()
                    n = 0
                    for t in range(5, 91):
                        hhh = table1.cell_value(t, 4)
                        if hhh != "":
                            n = n + 1
                    xx = 6 + n
                    wb1.sheets["TTLShpt"].range("E%d" % xx).value = xtotal

                    table2 = wb.sheet_by_name("BossData")
                    sht = wb1.sheets["BossData"]
                    sht.activate()
                    n = 0
                    for t in range(5, 91):
                        hhh = table2.cell_value(t, 5)
                        if hhh != "":
                            n = n + 1
                    # print(n)
                    # print(final_BossDataDf)
                    xx = 6 + n
                    wb1.sheets["BossData"].range("f%d" % xx).value = x
                    self.progressBar.setValue(100)
                    # spt33.to_excel("C:\\Users\\alanz\Desktop\\test access\\fz.xlsx",index=True)
                    wb1.save()
                    wb1.close()
                    app.quit()
                    tkinter.messagebox._show("AZ 提示：", "Program for BJS completed successfully")

        except Exception as reason:
            tkinter.messagebox.showerror("AZ 提示", str(reason))
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
