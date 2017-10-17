import requests
from lxml import etree
import xlwings as xw
url='http://hanyu.baidu.com/s'
payload={'wd':'','ptype':'zici'}
wb=xw.Book(r'hanzibiao.xlsx')
sht=wb.sheets[0]
i=1
word_list=sht.range('A%d:A7007'%(i)).value
for wd in word_list:
    payload['wd']=wd
    r=requests.get(url,payload)
    content=r.text
    tree=etree.HTML(content)
    nodes=tree.xpath('//img[@id="word_bishun"]')
    bishun_url=""
    if len(nodes)>0:
        bishun_url=(nodes[0].attrib)['data-gif']
    nodes=tree.xpath('//div[@id="pinyin"]/span/a')
    pinyin_url=""
    if len(nodes)>0:
        pinyin_url=(nodes[0].attrib)['url']
    sht.range("B%d"%(i)).value=bishun_url
    sht.range("C%d"%(i)).value=pinyin_url
    i=i+1