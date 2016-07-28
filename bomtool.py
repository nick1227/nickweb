
import argparse
import re
import xlrd
import xlwt

	
def intb_format(invalue):
	cellnum_str = re.findall(ur'[A-Z]+\d+-\d+',invalue)
	#print cellnum_str
	invalue_list =  invalue.split()
	for strtmp in cellnum_str:
		#print strtmp
		str_value = re.findall(ur'\d+',strtmp)
		#print str_value
		str_begin = str_value[0]
		str_end = str_value[1]
		str_promt = re.findall(r'[A-Z]+',strtmp)
		#print str_promt,str_begin,str_end

		content = [(''.join(str_promt) + str(i))  for i in xrange(int(str_value[0]),int(str_value[1])+1)]  
		#print content
		
		invalue_list.remove(strtmp)
		invalue_list.extend(content)
		
	cellvalue = ','.join(invalue_list)
	#print cellvalue
	return cellvalue
	
	
def run_xls(infile, outfile):
	inputfile = xlrd.open_workbook(infile)
	outputbook = xlwt.Workbook() 
	inputtb = inputfile.sheets()[0] 
	inputrows = inputtb.nrows
	inputcols = inputtb.ncols
	
	outtable = outputbook.add_sheet('bom new sheet',cell_overwrite_ok=True)
	
	for i in range(inputcols):
		if i == 2:
			for k in range(inputrows) :
				outtable.write(k,i, intb_format(inputtb.cell(k,i).value)) 
		else:
			[outtable.write(k,i, inputtb.cell(k,i).value) for  k in range(inputrows) ]
		
	#print outputbook
	outputbook.save(outfile)
	
if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='Unionman Bom Tool To Foramt Bom File')
	parser.add_argument('--input',action="store", dest="input",required=True)
	parser.add_argument('--output',action="store", dest="output",default='bom_out.xlsx')

	given_args = parser.parse_args()
	inputfile,outputfile = given_args.input,given_args.output

	print "Hello Luo, out file is %s" % outputfile
	#intb_format(u'L1 L3-6 L9 CA12-20')		
	run_xls(inputfile,outputfile)
	
	print "It's OK, Bye"



