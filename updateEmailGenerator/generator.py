# -*- coding: utf-8 -*-

import sys
reload(sys)
sys.setdefaultencoding('utf-8')
from os.path import abspath
import re
from cgi import escape

def htmlencode(string):
    return escape(string).replace('\n','<br/>')

def writeFile(string) :
    global email
    f = open(email_html,'a')
    f.write(string)
    f.close()

def readFile(fileName) :
    f = open(fileName, 'r')
    return f.read()

def replaceTarget(source, target, replace) :
    return source.replace(target, replace)

def addChineseTag(string) :
    inFlag = 0;
    output = ''
    for c in string.decode('utf-8') :
        if ord(c) >= 128 and inFlag == 0 :
            output += u'<span style="font-size:14.0pt;font-family:\'Times New Roman\',\xbc\xd0\xb7\xa2\xc5\xe9">'
            inFlag = 1
        if ord(c) < 128 and inFlag == 1 :
            output += u'</span>'
            inFlag = 0
        output += c
    if inFlag == 1 :
        output += u'</span>'
    return output
        

def main():
    global email_html
    global email_doc
    global excelFilePath
    global begin
    global newsBegin
    global newsTitle
    global newsUrl
    global newsContent
    global newline
    global updatesBegin
    global updatesTable
    global updatesCVE
    global updatesVersion
    global updatesTitle
    global updatesTableBegin
    global newline
    global end
    
    # init
    f = open(email_html,'w')
    f.close()

    # openExcel
    excelapp = win32com.client.Dispatch("Excel.Application")
    excelapp.Visible = 0
    excelxls = excelapp.Workbooks.Open(excelFilePath)

    # begin
    writeFile(readFile(begin))

    # news
    writeFile(readFile(newsBegin))

    news = excelxls.Worksheets("news")
    used = news.UsedRange
    nrows = used.Row + used.Rows.Count

    for i in range(2, nrows) :
        print ('news : ' + str(i - 1))
        writeFile(replaceTarget(readFile(newsTitle), '[newsTitle]', addChineseTag(str(news.Cells(i, 1)))))
        writeFile(replaceTarget(readFile(newsUrl), '[newsUrl]', str(news.Cells(i, 2))))
        writeFile(replaceTarget(readFile(newsContent), '[newsContent]', addChineseTag(str(news.Cells(i, 3)))))
        writeFile(readFile(newline))

    # updates
    writeFile(readFile(updatesBegin))

    updates = excelxls.Worksheets("updates")
    used = updates.UsedRange
    nrows = used.Row + used.Rows.Count

    title = ''
    for i in range(2, nrows) :
        print ('tables : ' + str(i - 1))
        table = readFile(updatesTable)
        CVEs = ''
        CVEtemp = readFile(updatesCVE)
        for CVE in re.split('\n|,',str(updates.Cells(i, 2))):
            if (len(CVE) == 0 or type(CVE) == 'NoneType') : continue
            cve = CVE.split('@')
            temp = CVEtemp
            temp = replaceTarget(temp, '[CVEnumber]', htmlencode(cve[0]))
            temp = replaceTarget(temp, '[CVEurl]', htmlencode(cve[1]))
            CVEs += temp

        table = replaceTarget(table, '[content]', str(updates.Cells(i, 3)))
        table = replaceTarget(table, '[suggest]', str(updates.Cells(i, 4)))
        
        Versions = ''
        versionTemp = readFile(updatesVersion)
        for version in str(updates.Cells(i, 5)).split(',') : 
            Versions += replaceTarget(versionTemp, '[version]', htmlencode(version))
        table = replaceTarget(table, '[risk]', str(updates.Cells(i, 6)))
        table = replaceTarget(table, '[CVEs]', CVEs)
        table = replaceTarget(table, '[Versions]', Versions)

        if title != str(updates.Cells(i, 1)) : 
            if (title != '') : 
                writeFile('</tbody></table>')
                writeFile(readFile(newline))
            title = str(updates.Cells(i, 1))
            writeFile(replaceTarget(readFile(updatesTitle), '[updatesTitle]', title))
            writeFile(readFile(updatesTableBegin))
            writeFile(table)
        else :     
            writeFile(table)

    writeFile('</tbody></table>')
    writeFile(readFile(newline))

    # end
    writeFile(readFile(end))

    #save as word
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Add(abspath(email_html))
    doc.SaveAs(abspath(email_doc),FileFormat=0)
    doc.Close()
    word.Quit()

    excelapp.Quit()
if __name__ == '__main__':
    try :
        import win32com.client
    except ImportError:
        print ('plz install win32com')
    
    #output
    email_html = abspath(r'../test/email.html')
    email_doc = abspath(r'../test/email.doc')
    
    #news
    excelFilePath = abspath(r'../test/email.xlsx')
    email = abspath(r'../test/email.html')
    begin = abspath(r'../sources/begin.html')
    newsBegin = abspath(r'../sources/news/newsBegin.html' )
    newsTitle = abspath(r'../sources/news/newsTitle.html' )
    newsUrl = abspath(r'../sources/news/newsUrl.html' )
    newsContent = abspath(r'../sources/news/newsContent.html' )
    newline = abspath(r'../sources/newline.html' )
    
    #updates
    updatesBegin = abspath(r'../sources/updates/updatesBegin.html' )
    updatesTable = abspath(r'../sources/updates/updatesTable.html' )
    updatesCVE = abspath(r'../sources/updates/updatesCVE.html' )
    updatesVersion = abspath(r'../sources/updates/updatesVersion.html' )
    updatesTitle = abspath(r'../sources/updates/updatesTitle.html' )
    updatesTableBegin = abspath(r'../sources/updates/updatesTableBegin.html' )
    newline = abspath(r'../sources/newline.html')
    end = abspath(r'../sources/end.html')
    
    
    
    main()
