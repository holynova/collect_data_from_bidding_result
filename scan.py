# -*- coding: utf-8 -*-
from openpyxl import Workbook,load_workbook
import xlrd
import os,datetime
import json

class Bid:
	def __init__(self):
		self.is_blank = True
class Pack:pass

#写入当前文件夹下以时间命名的函数
def save_to_file(common_name,content,file_type='.json'):
	str_now = datetime.datetime.now().strftime('%y%m%d_%H-%M-%S')
	result_file_name = os.path.dirname(os.path.abspath(__file__)) +'\\'+common_name +str_now+file_type

	with open(result_file_name,"w") as result_file:
		result_file.write(content.encode('utf-8'))

	print "Saved to file: " + result_file_name
	return result_file_name


class Scanner():
	def __init__(self,folder):
		self.folder = folder
		self.bids = []
	def scan(self):
		for root, dirs, files in os.walk(self.folder, topdown=True):
			for f in files:
				if f.find('.xls') != -1:
					wb_name = os.path.join(root,f)
					#warning what if file can not be open
					try:
						wb = xlrd.open_workbook(wb_name)
					except:
						print 'Error:fail to open ',wb_name
						continue
					# print wb_name
					
					bid = Bid()
					cnt_bid = 0
					for sht in wb.sheets():
						try:
							sht.name,sht.cell_value(0,0)
						except:
							print 'Error: blank worksheet found in ',wb_name,sht.name
							continue

						if sht.name.find(u'基本信息')!=-1 or ( sht.cell_value(0,0) and sht.cell_value(0,0) == u"投标信息组"):
							# print '--',sht.name
							# bid = Bid()
							bid.is_blank = False
							cnt_bid += 1
							if cnt_bid > 1:
								#出现了第二个基本信息
								print 'Warning: find two bids in one file :' ,wb_name
							# print f,isinstance(f,str),isinstance(f,unicode)
							# bid.file_name = f

							bid.work_num = sht.cell_value(1,1)
							bid.date = sht.cell_value(2,1)
							bid.sr = sht.cell_value(3,1)
							bid.company = sht.cell_value(4,1)
							bid.project = sht.cell_value(5,1)
							bid.csc = sht.cell_value(6,1)
							bid.scoring = sht.cell_value(7,1)
							# bid.
							bid.packs = []
							# self.bids.append(bid)

						elif sht.cell_value(0,0) == u"序号" and sht.cell_value(0,1) ==u"厂家":
							# print '----',sht.name
							pack = Pack()
							pack.pack_num = sht.cell_value(0,19)
							pack.nkt_gm3 = sht.cell_value(1,19)
							pack.num_company = sht.cell_value(2,19)
							pack.winner = sht.cell_value(3,19)
							pack.min = sht.cell_value(4,19)
							pack.max = sht.cell_value(5,19)
							pack.average = sht.cell_value(6,19)
							pack.average_no_peak = sht.cell_value(7,19)
							pack.median = sht.cell_value(8,19)
							pack.winner_price = sht.cell_value(9,19)
							pack.nkt_price = sht.cell_value(10,19)
							bid.packs.append(pack)


					if not bid.is_blank:
						self.bids.append(bid)

	def to_json(self):
		json_str = 	json.dumps(self.bids,default = lambda o:o.__dict__,indent = 4)
		return json_str

					

folder = r"E:\kuaipan\github\collect_data_from_bidding_result\bid_xls"
# folder = r"E:\kuaipan\github\collect_data_from_bidding_result\test"


s = Scanner(folder)
s.scan()
file_content = s.to_json()
save_to_file('bid',file_content,'.json')


def json_load(json_file_name):
	try:
		f = file(json_file_name,'r')	
	except:
		print 'Error:failed to open file ',json_file_name
	return json.load(f)


json_file_name = r'E:\kuaipan\github\collect_data_from_bidding_result\bid.json'
j = json_load(json_file_name)

print len(j)

output_str = 'date,winner,winner_price,average_no_peak,median,num_company,nkt_price,nkt_gm3,project,company,pack_num\n'
for bid in j:
	for pack in bid['packs']:
		output_str += "%s,%s,%.2f,%.2f,%.2f,%s,%.2f,%.4f,%s,%s,%s\n" %(bid['date'],pack['winner'],pack['winner_price'],pack['average_no_peak'],pack['median'],pack['num_company'],pack['nkt_price'],pack['nkt_gm3'],bid['project'],bid['company'],pack['pack_num'])
save_to_file('sum',output_str,'.txt')

