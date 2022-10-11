import os, sys
import pandas as pd
import subprocess
from pathlib import Path
import argparse 
import io


# Creates a new sheet in the XLSX file with the content of a pandas dataframe
def xlsx_preparer(df,writer,sheetname):

	workbook = writer.book
	####### Bookmarking colors #########
	# Light red fill for "evidence"
	format1 = workbook.add_format({'bg_color':   '#FFC7CE'})

	# Light yellow fill for "of interest"
	format2 = workbook.add_format({'bg_color':   '#FFEB9C'})

	# Green fill for "bookmark"
	format3 = workbook.add_format({'bg_color':   '#C6EFCE'})
	####################################


	(max_row, max_col) = df.shape
	df.insert(0,'Tag',[None]*max_row)
	df.to_excel(writer, sheet_name=sheetname, startrow=1, header=False, index=False)
	worksheet = writer.sheets[sheetname]
	column_settings = [{'header': column} for column in df.columns]
	worksheet.add_table(0, 0, max_row, max_col, {'columns': column_settings})
	worksheet.set_column(0, max_col, 17)
	worksheet.data_validation(1,0,max_row,0, {'validate': 'list', 'source': ['Evidence', 'Of interest', 'Bookmark']})
	worksheet.conditional_format(0,0,max_row,0, {'type':'text','criteria':'containing','value':'Evidence','format':format1})
	worksheet.conditional_format(0,0,max_row,0, {'type':'text','criteria':'containing','value':'Of interest','format':format2})
	worksheet.conditional_format(0,0,max_row,0, {'type':'text','criteria':'containing','value':'Bookmark','format':format3})


# Runs the given volatility command and returns the csv output as a pandas dataframe
def run_cmd_get_df(cmd):
	try:
		process = subprocess.Popen(cmd, stdout=subprocess.PIPE)
		csv = io.StringIO()
		for line in process.stdout:
			csv.write(line.decode().strip('"\n') + '\n')
		csv.seek(0)
		df = pd.read_csv(csv, index_col=0)
		csv.close()
		return df
	except Exception as e:
		pass


def pslist(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.pslist.PsList')
		df = run_cmd_get_df(cmd)
		df['CreateTime'] = pd.to_datetime(df['CreateTime'])
		df['ExitTime'] = pd.to_datetime(df['ExitTime'])
		xlsx_preparer(df,writer,'PsList')
	except Exception as e:
		pass


def psscan(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.psscan.PsScan')
		df = run_cmd_get_df(cmd)
		df['CreateTime'] = pd.to_datetime(df['CreateTime'])
		df['ExitTime'] = pd.to_datetime(df['ExitTime'])
		xlsx_preparer(df,writer,'PsScan')
	except Exception as e:
		pass

def netscan(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.netscan.NetScan')
		df = run_cmd_get_df(cmd)
		df['Created'] = pd.to_datetime(df['Created'])
		xlsx_preparer(df,writer,'NetScan')
	except Exception as e:
		pass

def cmdline(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.cmdline.CmdLine')
		df = run_cmd_get_df(cmd)
		#df['CreateTime'] = pd.to_datetime(df['CreateTime'])
		xlsx_preparer(df,writer,'CmdLine')
	except Exception as e:
		pass

def getsids(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.getsids.GetSIDs')
		df = run_cmd_get_df(cmd)
		xlsx_preparer(df,writer,'GetSIDs')
	except Exception as e:
		pass

def privs(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.privileges.Privs')
		df = run_cmd_get_df(cmd)
		xlsx_preparer(df,writer,'Privs')
	except Exception as e:
		pass

def ssdt(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.ssdt.SSDT')
		df = run_cmd_get_df(cmd)
		xlsx_preparer(df,writer,'SSDT')
	except Exception as e:
		pass

def mutantscan(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.mutantscan.MutantScan')
		df = run_cmd_get_df(cmd)
		xlsx_preparer(df,writer,'MutantScan')
	except Exception as e:
		pass

def driverscan(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.driverscan.DriverScan')
		df = run_cmd_get_df(cmd)
		xlsx_preparer(df,writer,'DriverScan')
	except Exception as e:
		pass

def driverirp(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.driverirp.DriverIrp')
		df = run_cmd_get_df(cmd)
		xlsx_preparer(df,writer,'DriverIrp')
	except Exception as e:
		pass

def dlllist(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'windows.dlllist.DllList')
		df = run_cmd_get_df(cmd)
		df['LoadTime'] = pd.to_datetime(df['LoadTime'])
		xlsx_preparer(df,writer,'DllList')
	except Exception as e:
		pass

def timeliner(target, writer):
	try:
		cmd = ('vol.exe', '-f', target, '-q', '-r', 'csv', 'timeliner.Timeliner')
		df = run_cmd_get_df(cmd)
		df['Created Date'] = pd.to_datetime(df['Created Date'])
		df['Modified Date'] = pd.to_datetime(df['Modified Date'])
		df['Accessed Date'] = pd.to_datetime(df['Accessed Date'])
		df['Changed Date'] = pd.to_datetime(df['Changed Date'])
		xlsx_preparer(df,writer,'Timeline')
	except Exception as e:
		pass


def create_xlsx_report(target, output):

	with pd.ExcelWriter(output, engine="xlsxwriter") as writer:  
		
		print("Launching pslist")
		pslist(target, writer)
		print("Launching psscan")
		psscan(target, writer)
		print("Launching netscan")
		netscan(target, writer)
		print("Launching cmdline")
		cmdline(target, writer)
		print("Launching getsids")
		getsids(target, writer)
		print("Launching privs")
		privs(target, writer)
		print("Launching ssdt")
		ssdt(target, writer)
		print("Launching mutantscan")
		mutantscan(target, writer)
		print("Launching driverscan")
		driverscan(target, writer)
		print("Launching driverirp")
		driverirp(target, writer)
		print("Launching dlllist")
		dlllist(target, writer)
		print("Launching timeliner")
		timeliner(target,writer)
		

def main():
	parser = argparse.ArgumentParser(description="VolXLSX by Martendal.", formatter_class=argparse.ArgumentDefaultsHelpFormatter)
	parser.add_argument("-t", "--target", required=True, help="Absolute path of the target memory dump.")
	parser.add_argument("-o", "--output", required=True, help="Absolute path of desired output report.")

	args = parser.parse_args()
	config = vars(args)
	print(config)

	target = config['target']
	output = config['output']

	print("Launching VolXLSX")
	
	
	print("Checking if the specified dump file exists.")
	if(not os.path.exists(target)):
		print("The dump file does not exist, check the given path (must be absolute).")
		quit()
	else:
		print("The dump file exists, proceeding...")
		if(not output.endswith(".xlsx")):
			output = output+".xlsx"
		create_xlsx_report(target, output)




if __name__ == "__main__":
	main()