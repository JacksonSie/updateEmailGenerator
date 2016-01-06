import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import win32com.client

def writeFile(string) :
	f = open('email.html','a')
	f.write(string)
	f.close()

def readFile(fileName) :
	f = open(fileName, 'r')
	return f.read()

def replaceTarget(source, target, replace) :
	return source.replace(target, replace)
# init
f = open('email.html','w')
f.close()

# openExcel
excelFilePath = '../updateEmailGenerator/email.xlsx'
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
	writeFile(replaceTarget(readFile('news/newsTitle.html'), '[newsTitle]', str(news.Cells(i, 1))))
	writeFile(replaceTarget(readFile('news/newsUrl.html'), '[newsUrl]', str(news.Cells(i, 2))))
	writeFile(replaceTarget(readFile('news/newsContent.html'), '[newsContent]', str(news.Cells(i, 3))))
	writeFile(readFile('newline.html'))

# updates
writeFile(readFile('updates/updatesBegin.html'))

writeFile(readFile('updates/updatesTitle.html'))
writeFile(readFile('updates/updatesTableBegin.html'))

table = readFile('updates/updatesTable.html')
cve = readFile('updates/updatesCVE.html')
cve += readFile('updates/updatesCVE.html')
cve += readFile('updates/updatesCVE.html')

version = readFile('updates/updatesVersion.html')
version += readFile('updates/updatesVersion.html')

table = replaceTarget(table, '[CVEs]', cve)
table = replaceTarget(table, '[Versions]', version)

writeFile(table)
writeFile(readFile('updates/updatesTable.html'))
writeFile('</tbody></table>')

writeFile(readFile('newline.html'))

writeFile(readFile('updates/updatesTitle.html'))
writeFile(readFile('updates/updatesTableBegin.html'))
writeFile(readFile('updates/updatesTable.html'))
writeFile('</tbody></table>')

# end
writeFile(readFile('end.html'))

excelapp.Quit()
