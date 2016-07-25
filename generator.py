# -*- coding: utf-8 -*-

def writeFile(string) :
    f = open('email.html','a')
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
            output += u'<span style="font-size:14.0pt;font-family:標楷體">'
            inFlag = 1
        if ord(c) < 128 and inFlag == 1 :
            output += u'</span>'
            inFlag = 0
        output += c
    if inFlag == 1 :
        output += u'</span>'
    return output
        

def main():
    # init
    f = open('email.html','w')
    f.close()

    # openExcel
    excelFilePath = abspath(r'../updateEmailGenerator/email.xlsx')
    excelapp = win32com.client.Dispatch("Excel.Application")
    excelapp.Visible = 0
    excelxls = excelapp.Workbooks.Open(excelFilePath)

    # begin
    writeFile(readFile('begin.html'))

    # news
    writeFile(readFile('news/newsBegin.html'))

    news = excelxls.Worksheets("news")
    used = news.UsedRange
    nrows = used.Row + used.Rows.Count

    for i in range(2, nrows) :
        print 'news : ' + str(i - 1)
        writeFile(replaceTarget(readFile('news/newsTitle.html'), '[newsTitle]', addChineseTag(str(news.Cells(i, 1)))))
        writeFile(replaceTarget(readFile('news/newsUrl.html'), '[newsUrl]', str(news.Cells(i, 2))))
        writeFile(replaceTarget(readFile('news/newsContent.html'), '[newsContent]', addChineseTag(str(news.Cells(i, 3)))))
        writeFile(readFile('newline.html'))

    # updates
    writeFile(readFile('updates/updatesBegin.html'))

    updates = excelxls.Worksheets("updates")
    used = updates.UsedRange
    nrows = used.Row + used.Rows.Count

    title = ''
    for i in range(2, nrows) :
        print 'tables : ' + str(i - 1)
        table = readFile('updates/updatesTable.html')
        CVEs = ''
        CVEtemp = readFile('updates/updatesCVE.html')
        for CVE in str(updates.Cells(i, 2)).split(',') : 
            cve = CVE.split('@')
            temp = CVEtemp
            temp = replaceTarget(temp, '[CVEnumber]', cve[0])
            temp = replaceTarget(temp, '[CVEurl]', cve[1])
            CVEs += temp

        table = replaceTarget(table, '[content]', str(updates.Cells(i, 3)))
        table = replaceTarget(table, '[suggest]', str(updates.Cells(i, 4)))
        
        Versions = ''
        versionTemp = readFile('updates/updatesVersion.html')
        for version in str(updates.Cells(i, 5)).split(',') : 
            Versions += replaceTarget(versionTemp, '[version]', version)
        table = replaceTarget(table, '[risk]', str(updates.Cells(i, 6)))

        table = replaceTarget(table, '[CVEs]', CVEs)
        table = replaceTarget(table, '[Versions]', Versions)

        if title != str(updates.Cells(i, 1)) : 
            if (title != '') : 
                writeFile('</tbody></table>')
                writeFile(readFile('newline.html'))
            title = str(updates.Cells(i, 1))
            writeFile(replaceTarget(readFile('updates/updatesTitle.html'), '[updatesTitle]', title))
            writeFile(readFile('updates/updatesTableBegin.html'))
            writeFile(table)
        else :     
            writeFile(table)

    writeFile('</tbody></table>')
    writeFile(readFile('newline.html'))

    # end
    writeFile(readFile('end.html'))

    #save as word
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Add(abspath('email.html'))
    doc.SaveAs(abspath('email.doc'),FileFormat=0)
    doc.Close()
    word.Quit()

    excelapp.Quit()
if __name__ == '__main__':
    import sys
    reload(sys)
    sys.setdefaultencoding('utf-8')
    from os.path import abspath
    import re
    
    try :
        import win32com.client
    except ImportError:
        print ('plz install win32com')

    main()
