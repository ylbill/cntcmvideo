# -*- coding: utf-8 -*-  
import sys
import os
import re
import urllib.request
from html.parser import HTMLParser
import win32com
from win32com.client import Dispatch, constants


class ContentHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.links = []
    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            if len(attrs) == 0:
                pass
            else:
                for (variable, value) in attrs:
                    if variable == 'href':
                        if value.find('content', 0) != -1:
                            self.links.append(value)


class ArticleHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.content = []
    def handle_data(self, data):
        self.content.append(data)

# 源网址
# source url
urlSource = 'http://www.cntcmvideo.com/zgzyyb/html'
urlDate = input('Enter date(format[YYYYMMDD], i.e. "20140519"): ')
#urlInput = '{}/{}/{}'.format(urlSource, urlDate[0:6], urlDate[6:8])
urlInput = urlSource+'/'+urlDate[0:4]+'-'+urlDate[4:6]+'/'+urlDate[6:8]
print(urlInput)

# word文件存放目录
# dir in which put word documents
docDir = input('Enter saving directory: ')

# read all content links
contentPage = urllib.request.urlopen(urlInput).read().decode('utf-8')
contentParser = ContentHTMLParser()
contentParser.feed(contentPage)

targetList = ['学术与临床', '农村与社区', '养生保健', '健康关注', '视点']
for idx in contentParser.links:
    articlePage = urllib.request.urlopen(urlInput + '/' + idx).read().decode('utf-8')
    targetFind = False
    pageName = ''
    for target in targetList:
        if(articlePage.find('<STRONG>'+target+'</STRONG>') != -1):
            targetFind = True
            pageName = target
            break

    if targetFind == False:
        continue

    articleBlock = re.findall(u'<!----------文章部分开始---------->([\s\S]*?)<!----------文章部分结束---------->', articlePage)
    articleParser = ArticleHTMLParser()
    for art in articleBlock:
        articleHeader = re.findall(u'<tr valign=top> <td [\s\S]*?>([\s\S]*?)</td> </TR>', art)
        articlePicture = re.findall(u'<IMG src="([\s\S]*?)">', art)
        articleContent = re.findall(u'<content>([\s\S]*?)</content>', art)

        articleParser.feed(articleContent[0])

        w = win32com.client.Dispatch('Word.Application')

        w.Visible = 0
        w.DisplayAlerts = 0
        doc = w.Documents.Add()

        wrange = doc.Range()

        for header in articleHeader:
            wrange.InsertAfter(header + '\n')

        #for picture in articlePicture:
        #    wrange.InlineShapes.AddPicture(urlInput+'/'+idx+'/'+picture)
        #    w.Selection.InlineShapes.AddPicture(FileName=urlInput+'/'+idx+'/'+picture,LinkToFile= False,SaveWithDocument=True)
            
        for content in articleParser.content:
            wrange.InsertAfter(content + '\n')

        
        
        #docName = u'{}-{}.docx'.format(pageName.decode('utf-8'), articleHeader[1])
        docName = idx+'.docx'
        #candidate = os.path.join(docDir, docName)
        #candidate = docName
        
        # if file exists, add a postfix after filename, i.e "_2", "_3"
        #i = 2
        #while i > 0:
         #   if os.path.isfile(candidate):
                #docName = u'{}-{}_{}.docx'.format(pageName.decode('utf-8'), articleHeader[1], i)
          #      docName = idx+'-'+i+'.docx'
                #candidate = os.path.join(docDir, docName)
           #     candidate = docName
            #    i += 1
            #else:
            #    i = -1

        doc.SaveAs(docName)
        print(idx+': '+docName)

        w.Documents.Close()
        w.Quit()


