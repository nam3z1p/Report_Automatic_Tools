# -*- coding: utf-8 -*-
import sys
import string
import glob
import os
import argparse
import re
import win32com.client
from PIL import ImageGrab, Image
from datetime import datetime

F_EXCEL_LIST = []
F_EXCEL_TARGET = []
F_TARGET_COMP = []


def Main_title():
    print("######################################################")
    print("##          Report_Automatic_Tools 0.4              ##")
    print("##                                                  ##")
    print("##                            Developed by nam3z1p  ##")
    print("##                                         2020.01  ##")
    print("######################################################")


def Main():

    Main_title()

    fpath = ".\\Input_Files\\*.xlsx"

    for fpath in glob.glob(fpath):
        fpath_s = fpath.split(']')[1].split('_')[0]
        try:
            if fpath.find('~$') < 0:
                F_EXCEL_LIST.append(
                    os.getcwd()+"\\Input_Files\\"+fpath.split("\\")[2])
                F_EXCEL_TARGET.append(fpath_s)
                F_TARGET_COMP.append(fpath.split('[')[1].split(']')[0])
        except:
            print("[-] Error :", sys.exc_info()[1])
            continue

    xl = win32com.client.gencache.EnsureDispatch('Excel.Application')

    for i in range(0, len(F_EXCEL_LIST)):

        wb1 = xl.Workbooks.Open(F_EXCEL_LIST[i])
        xl.Worksheets('점검결과 상세').Activate()
        ws1 = xl.ActiveSheet
        xl.Worksheets('대응방안').Activate()
        ws2 = xl.ActiveSheet

        print("[+] Input - "+F_EXCEL_LIST[i].split('\\')[-1]+"")

        if ws1.Range('B2').Value != '대상':
            print('[!] %s B2열 대상 확인 필요')
            return
        elif ws1.Range('N2').Value != 'FileName':
            print('[!] %s N열 FileName 확인 필요')
            return
        elif ws1.Range('O2').Value != '이미지':
            print('[!] %s O열 이미지 확인 필요')
            return

        ws1_rowcount = ws1.UsedRange.Rows.Count+2
        ws2_rowcount = ws2.UsedRange.Rows.Count+2

        target_num = 3
        report_range = 0
        for a in range(0, ws1_rowcount):
            if ws1.Range('B%s' % (target_num+a)).Value != None:
                report_range = report_range+1

        F_VULN_REPORT = [[] for a in range(report_range)]

        report_index = 0
        for a in range(0, ws1_rowcount):
            if ws1.Range('B%s' % (target_num+a)).Value != None:
                F_VULN_REPORT[report_index].append(target_num+a)
                F_VULN_REPORT[report_index].append(
                    ws1.Range('B%s' % (target_num+a)).Value.split(" - [")[0])
                F_VULN_REPORT[report_index].append(
                    ws1.Range('B%s' % (target_num+a)).Value.split(" - [")[1].split(']')[0])
                report_index = report_index+1

# Array List
#       F_VULN_REPORT[0][0]     F_VULN_REPORT[0][1]     F_VULN_REPORT[0][2]
#       F_TARGET_INDEX,         F_TARGET_NAME,          F_TARGET_URL
#       F_VULN_REPORT[0][3][]   F_VULN_REPORT[0][4][]   F_VULN_REPORT[0][5][]
#       F_VULN_NAME,            F_VULN_DETAIL,          F_VULN_URL,
#       F_VULN_REPORT[0][6][]   F_VULN_REPORT[0][7][]
#       F_VULN_FILENAME         F_VULN_CODE
######

        if len(F_VULN_REPORT) != 1:
            for a in range(0, len(F_VULN_REPORT)):
                F_VULN_NAME, F_VULN_DETAIL, F_VULN_FILENAME, F_VULN_URL, F_VULN_REASON, F_VULN_DEFENSE = [], [], [], [], [], []

                if a == (len(F_VULN_REPORT)-1):
                    for b in range(F_VULN_REPORT[a][0], ws1_rowcount):
                        if ws1.Range('N%s' % (b)).Value != None:
                            F_VULN_FILENAME.append(
                                ws1.Range('N%s' % (b)).Value)
                            if ws1.Range('J%s' % (b)).Value == None:
                                for c in range(b, F_VULN_REPORT[a][0], -1):
                                    if ws1.Range('J%s' % (c-1)).Value != None:
                                        F_VULN_NAME.append(
                                            ws1.Range('J%s' % (c-1)).Value)
                                        break
                            else:
                                F_VULN_NAME.append(
                                    ws1.Range('J%s' % (b)).Value)
                            F_VULN_DETAIL.append(ws1.Range('K%s' % (b)).Value)
                            F_VULN_URL.append(ws1.Range('L%s' % (b)).Value)

                        for c in range(target_num, ws2_rowcount):
                            if ws2.Range('C%s' % (c)).Value != None:
                                if ws2.Range('C%s' % (c)).Value == ws1.Range('M%s' % (b)).Value:
                                    F_VULN_REASON.append(
                                        ws2.Range('D%s' % (c)).Value.replace("\n", "\r"))
                                    F_VULN_DEFENSE.append(
                                        ws2.Range('E%s' % (c)).Value.replace("\n", "\r"))

                else:
                    for b in range(F_VULN_REPORT[a][0], F_VULN_REPORT[a+1][0]):
                        F_VULN_FILENAME.append(ws1.Range('N%s' % (b)).Value)
                        if ws1.Range('J%s' % (b)).Value == None:
                            for c in range(b, F_VULN_REPORT[a][0], -1):
                                if ws1.Range('J%s' % (c-1)).Value != None:
                                    F_VULN_NAME.append(
                                        ws1.Range('J%s' % (c-1)).Value)
                                    break
                        else:
                            F_VULN_NAME.append(ws1.Range('J%s' % (b)).Value)
                        F_VULN_DETAIL.append(ws1.Range('K%s' % (b)).Value)
                        F_VULN_URL.append(ws1.Range('L%s' % (b)).Value)

                        for c in range(target_num, ws2_rowcount):
                            if ws2.Range('C%s' % (c)).Value != None:
                                if ws2.Range('C%s' % (c)).Value == ws1.Range('M%s' % (b)).Value:
                                    F_VULN_REASON.append(
                                        ws2.Range('D%s' % (c)).Value.replace("\n", "\r"))
                                    F_VULN_DEFENSE.append(
                                        ws2.Range('E%s' % (c)).Value.replace("\n", "\r"))

                F_VULN_REPORT[a].append(F_VULN_NAME)
                F_VULN_REPORT[a].append(F_VULN_DETAIL)
                F_VULN_REPORT[a].append(F_VULN_URL)
                F_VULN_REPORT[a].append(F_VULN_FILENAME)
                F_VULN_REPORT[a].append(F_VULN_REASON)
                F_VULN_REPORT[a].append(F_VULN_DEFENSE)

        else:
            F_VULN_NAME, F_VULN_DETAIL, F_VULN_FILENAME, F_VULN_URL, F_VULN_REASON, F_VULN_DEFENSE = [], [], [], [], [], []

            for a in range(0, ws1_rowcount):
                if ws1.Range('N%s' % (target_num+a)).Value != None:
                    F_VULN_FILENAME.append(
                        ws1.Range('N%s' % (target_num+a)).Value)
                    if ws1.Range('J%s' % (target_num+a)).Value == None:
                        for b in range(target_num+a, 0, -1):
                            if ws1.Range('J%s' % (b-1)).Value != None:
                                F_VULN_NAME.append(
                                    ws1.Range('J%s' % (b-1)).Value)
                                break
                    else:
                        F_VULN_NAME.append(
                            ws1.Range('J%s' % (target_num+a)).Value)

                    F_VULN_DETAIL.append(
                        ws1.Range('K%s' % (target_num+a)).Value)
                    F_VULN_URL.append(ws1.Range('L%s' % (target_num+a)).Value)

                    for b in range(target_num, ws2_rowcount):
                        if ws2.Range('C%s' % (b)).Value != None:
                            if ws2.Range('C%s' % (b)).Value == ws1.Range('M%s' % (target_num+a)).Value:
                                F_VULN_REASON.append(
                                    ws2.Range('D%s' % (b)).Value.replace("\n", "\r"))
                                F_VULN_DEFENSE.append(
                                    ws2.Range('E%s' % (b)).Value.replace("\n", "\r"))

            F_VULN_REPORT[0].append(F_VULN_NAME)
            F_VULN_REPORT[0].append(F_VULN_DETAIL)
            F_VULN_REPORT[0].append(F_VULN_URL)
            F_VULN_REPORT[0].append(F_VULN_FILENAME)
            F_VULN_REPORT[0].append(F_VULN_REASON)
            F_VULN_REPORT[0].append(F_VULN_DEFENSE)

        for a in range(0, len(F_VULN_REPORT)):
            for b in range(0, len(F_VULN_REPORT[a][6])):

                width, height = 449, 313

                width_t = (width/2)*0.25
                height_t = height*0.75

                inc_Image_Col = F_VULN_REPORT[a][0]+b

                ws1.Range("O:O").ColumnWidth = width_t
                ws1.Range("%s:%s" % (inc_Image_Col, inc_Image_Col)
                          ).RowHeight = height_t

                ws1.Range('O%s' % inc_Image_Col).CopyPicture(1, 2)
                im_copy = ImageGrab.grabclipboard()

                insert_image_fpath = os.getcwd()+"\\Input_Files\\" + \
                    F_VULN_REPORT[a][6][b]
                im_copy.save(insert_image_fpath, 'jpeg')

        wd1 = win32com.client.gencache.EnsureDispatch('Word.Application')
        wd1.Visible = 0

        for a in range(0, len(F_VULN_REPORT)):

            wdoc1 = wd1.Documents.Open(os.getcwd()+'\\Sample_상세보고서.docx')

            for f, shape in enumerate(wdoc1.Sections(1).Headers(1).Shapes):
                if shape.Name.startswith('Text'):
                    shape.TextFrame.TextRange.Text = "URL : " + \
                        F_VULN_REPORT[a][2]

            for f, shape in enumerate(wdoc1.Shapes):
                if shape.Name.startswith('Text'):
                    shape.TextFrame.TextRange.Text = '(' + \
                        F_TARGET_COMP[i]+')\n'+F_VULN_REPORT[a][1]+'\n'+'상세보고서'

            wdoc1.Tables(1).Range.Copy()

            for b in range(0, len(F_VULN_REPORT[a][3])):

                FromTo = {"F_VULN_NAME": str(F_VULN_REPORT[a][3][b]), "F_VULN_DETAIL": F_VULN_REPORT[a][4][b], "F_VULN_URL": F_VULN_REPORT[a]
                          [5][b], "F_VULN_REASON": F_VULN_REPORT[a][7][b], "F_VULN_DEFENSE": F_VULN_REPORT[a][8][b]}

                for From in FromTo.keys():
                    wd1.Selection.Find.Text = From
                    wd1.Selection.Find.Replacement.Text = FromTo[From]
                    wd1.Selection.Find.Execute(Replace=2)

                insert_image_fpath = os.getcwd()+"\\Input_Files\\" + \
                    F_VULN_REPORT[a][6][b]

                wdoc1_table = wdoc1.Tables(b+1)

                wdoc1_table.Cell(2, 1).Split(2, 1)
                cell_range = wdoc1_table.Cell(3, 1).Range

                cell_range.ParagraphFormat.LeftIndent = 19

                current_pic = cell_range.InlineShapes.AddPicture(
                    insert_image_fpath)
                current_pic.Height = height
                current_pic.Width = width

                wdoc1_table.Cell(2, 1).Merge(wdoc1_table.Cell(3, 1))

                wdoc1_table.Cell(2, 1).Split(2, 1)
                cell_range = wdoc1_table.Cell(3, 1).Range

                cell_range.ParagraphFormat.Alignment = win32com.client.constants.wdAlignParagraphCenter
                cell_range.ParagraphFormat.LeftIndent = 19
                cell_range.Font.Bold = False

                cell_range.InsertAfter("[ 이미지 설명 ]")
                wdoc1_table.Cell(2, 1).Merge(wdoc1_table.Cell(3, 1))

                os.remove(insert_image_fpath)

                if(b < len(F_VULN_REPORT[a][3])-1):
                    location = wdoc1.Content
                    location.Paragraphs.Add()
                    location.Collapse(0)
                    location.Paste()

            wdoc1.SaveAs(os.getcwd()+"\\Output_Files\\[%s]%s_상세보고서_%s.docx" % (
                F_TARGET_COMP[i], F_VULN_REPORT[a][1], datetime.today().strftime("%Y%m%d")))
            print("[+] Output - [%s]%s_상세보고서_%s.docx" % (
                F_TARGET_COMP[i], F_VULN_REPORT[a][1], datetime.today().strftime("%Y%m%d")))

            wdoc1.Close(SaveChanges=False)
        print("######################################################")
        wb1.Close(False)
    wd1.Quit()
    xl.Quit()

    print('[+] Done')


if __name__ == "__main__":
    Main()
