import xlwings as wx

def main():
	'''
	#创建目录页面，并建立超链接并为每个sheet 建立一个跳转到目录的超链接
	'''
	wb = wx.Book.caller()
	#sheet 数据量
	#nsheet = wb.sheets.count
	sName = '目录'
	#增加一个目录页
	if sName not in ( s.name for s in wb.sheets):
		wb.sheets.add(name='目录',before=wb.sheets[0])
	else:
		wb.sheets[sName].clear()
	cntPage = 0

	for page in wb.sheets:
		pagename = page.name
		cntPage += 1
		#创建目录页面，并建立超链接
		wb.sheets['目录'].api.Hyperlinks.Add(Anchor=wb.sheets['目录'].range((cntPage+1,2)).api,Address="",SubAddress=pagename+"!A1",ScreenTip="",TextToDisplay=pagename)
		#为每个sheet 建立一个跳转到目录的超链接
		wb.sheets[pagename].api.Hyperlinks.Add(Anchor=wb.sheets[pagename].range((1,1)).end('left').api,Address="",SubAddress=sName+"!A1",ScreenTip="",TextToDisplay="返回"+sName)

