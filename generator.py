

def writeFile(string) :
	f = open('email.html','a')
	f.write(string)
	f.close()

def readFile(fileName) :
	f = open(fileName, 'r')
	return f.read()

def replaceTarget(source, target, replace) :
	return source.replace(target, replace)

f = open('email.html','w')
f.close()

# begin
writeFile(readFile('begin.html'))

# news
writeFile(readFile('news/newsBegin.html'))

writeFile(replaceTarget(readFile('news/newsTitle.html'), '[newsTitle]', 'AAAAAAAAA'))
writeFile(readFile('news/newsUrl.html'))
writeFile(readFile('news/newsContent.html'))

writeFile(readFile('newline.html'))

writeFile(readFile('news/newsTitle.html'))
writeFile(readFile('news/newsUrl.html'))
writeFile(readFile('news/newsContent.html'))

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
writeFile(readFile('updates/updatesTable.html'))
writeFile('</tbody></table>')

# end
writeFile(readFile('end.html'))
