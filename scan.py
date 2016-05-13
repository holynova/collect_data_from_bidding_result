# -*- coding: utf-8 -*-
from openpyxl import Workbook,load_workbook
import xlrd
import os,datetime
import json

class Bid:pass
class Pack:pass

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
					print wb_name
					
					bid = Bid()
					cnt_bid = 0
					for sht in wb.sheets():
						if sht.name.find(u'基本信息')!=-1 or sht.cell_value(0,0) == u"投标信息组":
							# print '--',sht.name
							# bid = Bid()
							cnt_bid += 1
							if cnt_bid > 1:
								#出现了第二个基本信息
								
							bid.work_num = she.cell_value(1,0)
							bid.date = she.cell_value(2,0)
							bid.sr = she.cell_value(3,0)
							bid.company = she.cell_value(4,0)
							bid.project = she.cell_value(5,0)
							bid.csc = she.cell_value(6,0)
							bid.scoring = she.cell_value(7,0)

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


					self.bids.append(bid)

	def save_as_json(self):
		pass

											



folder = r"E:\kuaipan\github\collect_data_from_bidding_result\bid_xls"
folder = r"E:\kuaipan\github\collect_data_from_bidding_result\test"


s = Scanner(folder)
s.scan()
