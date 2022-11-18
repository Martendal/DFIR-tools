'''
Win_Disset by Martendal.

This is just a script to launch Eric Zimmerman's tools on Windows acquisitions. It converts Windows artefacts to CSV files, as well as a SQLite database and a XLSX report that contains only some interesting artefacts.
Some important details to use this tool:
	- The target acquisitions must respect Windows folders hierarchy (ex: evtx in C:/Windows/system32/winevt/logs). If it doesn't, this script won't work.
	- It partially works on local machine's artefacts but some stuff won't be converted (Registry for instance). It's really designed for KAPE acquisitions.
	- Use absolute paths in arguments.
	- There is a powershell setup script available on the Github repository of this script, that should get all the requirements. If you don't use it, make sure to follow instructions on the repository.

https://github.com/Martendal/DFIR-tools
'''



import os, sys
import pandas as pd
import sqlite3
import subprocess
from pathlib import Path
import argparse 
import logging


logging.basicConfig(filename="windissect_log.log", format='%(asctime)s %(message)s', filemode='w')
Logger = logging.getLogger()
Logger.setLevel(logging.DEBUG)

#EVTX
def evtx_to_csv(evtx_folder, output):
	Logger.info("Converting EVTX to CSV.")
	path = output+"\\EVTX"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	for filename in os.listdir(evtx_folder):
		Logger.info("Converting "+filename+" to CSV.")
		subprocess.call(['Utils\\EvtxECmd\\EvtxECmd.exe', '-f', evtx_folder+"\\"+filename, '--csv', path, '--csvf', path+"\\"+filename+".csv"])
	Logger.info("Converting all EVTX to a single CSV file.")
	subprocess.call(['Utils\\EvtxECmd\\EvtxECmd.exe', '-d', evtx_folder, '--csv', path, '--csvf', path+"\\"+"global.csv"])

#EVTX
def evtx_to_csv_partial(evtx_folder, output):
	Logger.info("Converting Security, System and Application EVTX to CSV.")
	path = output+"\\EVTX"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")

	subprocess.call(['Utils\\EvtxECmd\\EvtxECmd.exe', '-f', evtx_folder+"\\Security.evtx", '--csv', path, '--csvf', path+"\\"+"Security.evtx.csv"])
	subprocess.call(['Utils\\EvtxECmd\\EvtxECmd.exe', '-f', evtx_folder+"\\System.evtx", '--csv', path, '--csvf', path+"\\"+"System.evtx.csv"])
	subprocess.call(['Utils\\EvtxECmd\\EvtxECmd.exe', '-f', evtx_folder+"\\Application.evtx", '--csv', path, '--csvf', path+"\\"+"Application.evtx.csv"])
#TODO: complete with other needed logs


#Prefetch
def prefetch_to_csv(prefetch_folder, output):
	Logger.info("Converting Preftech files to CSV.")
	path = output+"\\Prefetch"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\PECmd.exe', '-d', prefetch_folder, '--csv', path, '--csvf', 'prefetch.csv', '-q'])

#Amcache
def amcache_to_csv(amcache_folder, output):
	Logger.info("Converting Amcache to CSV.")
	path = output+"\\Amcache"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\AmcacheParser.exe', '-f', amcache_folder+"\\Amcache.hve", '--csv', path, '-i'])
	for filename in os.listdir(path):
		os.rename(os.path.join(path, filename), os.path.join(path, filename.replace(filename[0:15], '')))


#AppCompatCache
def appcompatcache_to_csv(appcompatcache_folder, output):
	Logger.info("Converting AppCompatCache to CSV.")
	path = output+"\\AppCompatCache"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\AppCompatCacheParser.exe', '-f', appcompatcache_folder+"\\SYSTEM", '--csv', path, '--csvf', 'appcompatcache.csv'])


#USNjournal
def usnjournal_to_csv(usnjournal_folder, output):
	Logger.info("Converting USNJournal to CSV.")
	path = output+"\\USNjournal"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\MFTEcmd.exe', '-f', usnjournal_folder+"\\$J", '--csv', path, '--csvf', 'USNjournal.csv'])

#LNK
def lnk_to_csv(lnk_folder, output):
	Logger.info("Converting LNK files to CSV.")
	path = output+"\\LNK"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\LECmd.exe', '-d', lnk_folder, '--csv', path, '--csvf', 'lnk.csv', '-q'])


#RecycleBin
def recyclebin_to_csv(recyclebin_folder, output):
	Logger.info("Converting RecycleBin to CSV.")
	path = output+"\\RecycleBin"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\RBCmd.exe', '-d', recyclebin_folder, '--csv', path, '--csvf', 'RecycleBin.csv', '-q'])


#SRUM
def srum_to_csv(srum_folder, output):
	Logger.info("Converting SRUM to CSV.")
	path = output+"\\SRUM"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\SrumECmd.exe', '-d', srum_folder, '--csv', path])
	for filename in os.listdir(path):
		os.rename(os.path.join(path, filename), os.path.join(path, filename.replace(filename[0:15], '')))


#Registries
def registries_to_csv(registries_folder, output):
	Logger.info("Converting Windows Registries to CSV.")
	path = output+"\\Registries"
	try:
		os.mkdir(path)
	except FileExistsError:
		Logger.info("Output folder already exists, results will be overwritten.")
	subprocess.call(['Utils\\RECmd\\RECmd.exe', '-d', registries_folder, '--csv', path, '--csvf', 'registry.csv', '--bn', 'Utils\\RECmd\\BatchExamples\\Kroll_Batch.reb'])


#Remove whitespaces in filenames
def remove_whitespaces_filename(root):
    for path, folders, files in os.walk(root):
        for f in files:
            os.rename(os.path.join(path, f), os.path.join(path, f.replace(' ', '_')))


#Create DB
def create_database(root, db_name):
	for p in Path(root).rglob('*.csv'):
		table_name = str(os.path.basename(p))
		import_string = '.import '+str(Path(p).resolve())+' '+table_name
		subprocess.call(['Utils\\SQLite\\sqlite3.exe', root+'\\'+db_name, '--cmd', '.mode csv', '.separator ,',import_string.replace("\\","/")])


def create_database_alt(root, db_path):

	chunksize = 1000
	

	#Create DB first
	con = None
	try:
		con = sqlite3.connect(db_path)
		Logger.info("Creating database")
	except Error as e:
		Logger.info(e)
	finally:
		if con:
			for p in Path(root).rglob('*.csv'):
				index_offset = 1
				Logger.info("Adding "+str(p)+" to database.")
				table_name = str(os.path.basename(p))
				csv_path = str(Path(p).resolve())
				Logger.info("Table name: "+table_name)
				Logger.info("CSV path: "+csv_path)
				for df in pd.read_csv(csv_path, chunksize = chunksize, iterator = True, dtype=str):
					df.index += index_offset
					df.to_sql(table_name, con, if_exists = 'append')
					index_offset += 1
			con.close()

	



############################################################################### Reporting ################################################################
def xlsx_preparer(df,writer,sheetname):
	(max_row, max_col) = df.shape
	df.insert(0,'Tag',[None]*max_row)
	df.to_excel(writer, sheet_name=sheetname, startrow=1, header=False, index=False)
	worksheet = writer.sheets[sheetname]
	column_settings = [{'header': column} for column in df.columns]
	worksheet.add_table(0, 0, max_row, max_col, {'columns': column_settings})
	worksheet.set_column(0, max_col, 17)
	worksheet.data_validation(0,0,max_row,0, {'validate': 'list', 'source': ['Evidence', 'Of interest', 'Bookmark']})
	worksheet.conditional_format(0,0,max_row,0, {'type':'text','criteria':'containing','value':'Evidence','format':format1})
	worksheet.conditional_format(0,0,max_row,0, {'type':'text','criteria':'containing','value':'Of interest','format':format2})
	worksheet.conditional_format(0,0,max_row,0, {'type':'text','criteria':'containing','value':'Bookmark','format':format3})



def cleared_audit_log(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=1102)", con)
		df2 = pd.read_sql_query("SELECT * from \"System.evtx.csv\" WHERE (EventId=102)", con)
		frame = [df,df2]
		result = pd.concat(frame)
		result['TimeCreated'] = pd.to_datetime(result['TimeCreated'])
		xlsx_preparer(result,writer,'Event log cleared')
	except Exception as e:
		pass

def logon_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=4768) OR (EventId=4769) OR (EventId=4770) OR (EventId=4771) OR (EventId=4776) OR (EventId=4624) OR (EventId=4625) OR (EventId=4634) OR (EventId=4647) OR (EventId=4648) OR (EventId=4672) OR (EventId=4778) OR (EventId=4779)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'Logon Events')
	except Exception as e:
		pass


def account_management_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=4720) OR (EventId=4722) OR (EventId=4723) OR (EventId=4724) OR (EventId=4725) OR (EventId=4726) OR (EventId=4727) OR (EventId=4728) OR (EventId=4729) OR (EventId=4730) OR (EventId=4731) OR (EventId=4732) OR (EventId=4733) OR (EventId=4734) OR (EventId=4735) OR (EventId=4737) OR (EventId=4738) OR (EventId=4741) OR (EventId=4742) OR (EventId=4743) OR (EventId=4754) OR (EventId=4755) OR (EventId=4756) OR (EventId=4757) OR (EventId=4758) OR (EventId=4798) OR (EventId=4799)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'Account mgt Events')
	except Exception as e:
		pass	
	

def scheduled_tasks_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=4698) OR (EventId=4699) OR (EventId=4700) OR (EventId=4701) OR (EventId=4702)", con)
		df2 = pd.read_sql_query("SELECT * from \"Microsoft-Windows-TaskScheduler%4Operational.evtx.csv\" WHERE (EventId=106) OR (EventId=140) OR (EventId=141) OR (EventId=200) OR (EventId=201)", con)
		frame = [df,df2]
		result = pd.concat(frame)
		result['TimeCreated'] = pd.to_datetime(result['TimeCreated'])
		xlsx_preparer(result,writer,'Sched. Tasks Events')
	except Exception as e:
		pass

def RDP_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=4778) OR (EventId=4779) OR (EventId=4624)", con)
		try:
			df2 = pd.read_sql_query("SELECT * from \"RemoteDesktopServices-RDPCoreTS%4Operational.evtx.csv\" WHERE (EventId=131)", con)
		except Exception as e:
			df2 = pd.DataFrame()
		try:
			df3 = pd.read_sql_query("SELECT * from \"TerminalServices-RemoteConnectionManager%4Operational.evtx.csv\" WHERE (EventId=1149)", con)
		except Exception as e:
			df3 = pd.DataFrame()
		frame = [df,df2,df3]
		result = pd.concat(frame)
		result['TimeCreated'] = pd.to_datetime(result['TimeCreated'])
		xlsx_preparer(result,writer,'RDP Events')
	except Exception as e:
		pass

def application_installation_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Application.evtx.csv\" WHERE (EventId=1033) OR (EventId=1034) OR (EventId=11707) OR (EventId=11708) OR (EventId=11724)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'App Install Events')
	except Exception as e:
		pass

def services_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"System.evtx.csv\" WHERE (EventId=7034) OR (EventId=7035) OR (EventId=7036) OR (EventId=7040) OR (EventId=7045)", con)
		df2 = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=4697)", con)
		frame = [df,df2]
		result = pd.concat(frame)
		result['TimeCreated'] = pd.to_datetime(result['TimeCreated'])
		xlsx_preparer(result,writer,'Services Events')
	except Exception as e:
		pass

def WMI_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Microsoft-Windows-WMI-Activity%4Operational.evtx.csv\" WHERE (EventId=5857) OR (EventId=5858) OR (EventId=5859) OR (EventId=5860) OR (EventId=5861)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'WMI Events')
	except Exception as e:
		pass

def PowerShell_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Microsoft-Windows-PowerShell%4Operational.evtx.csv\" WHERE (EventId=4103) OR (EventId=4104)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'PowerShell Events')
	except Exception as e:
		pass

def Process_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=4688) OR (EventId=4689)", con)
		df2 = pd.read_sql_query("SELECT * from \"Application.evtx.csv\" WHERE (EventId=1000) OR (EventId=1001) OR (EventId=1002)", con)
		df3 = pd.read_sql_query("SELECT * from \"System.evtx.csv\" WHERE (EventId=1001)", con)
		frame = [df,df2,df3]
		result = pd.concat(frame)
		result['TimeCreated'] = pd.to_datetime(result['TimeCreated'])
		xlsx_preparer(result,writer,'Process Events')
	except Exception as e:
		pass

def Network_Share_events(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Security.evtx.csv\" WHERE (EventId=5140) OR (EventId=5142) OR (EventId=5143) OR (EventId=5144) OR (EventId=5145)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'Share Events')
	except Exception as e:
		pass

def amcache_program_presence(con, writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Amcache_UnassociatedFileEntries.csv\"", con)
		df2 = pd.read_sql_query("SELECT * from \"Amcache_AssociatedFileEntries.csv\"", con)
		frame = [df,df2]
		result = pd.concat(frame)
		result['FileKeyLastWriteTimestamp'] = pd.to_datetime(result['FileKeyLastWriteTimestamp'])
		result['LinkDate'] = pd.to_datetime(result['LinkDate'])
		xlsx_preparer(result,writer,'Present programs (amcache)')
	except Exception as e:
		pass
	
	

def recyclebin(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"RecycleBin.csv\"", con)
		df['DeletedOn'] = pd.to_datetime(df['DeletedOn'])
		xlsx_preparer(df,writer,'RecycleBin')
	except Exception as e:
		pass
	

def prefetch_program_execution(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"prefetch_Timeline.csv\"", con)
		df['RunTime'] = pd.to_datetime(df['RunTime'])
		xlsx_preparer(df,writer,'Program execution (prefetch)')
	except Exception as e:
		pass
	

def amcache_plugged_devices(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Amcache_DevicePnps.csv\"", con)
		df['FileKeyLastWriteTimestamp'] = pd.to_datetime(df['FileKeyLastWriteTimestamp'])
		xlsx_preparer(df,writer,'Amcache plugged devices')
	except Exception as e:
		pass
	

def drivers(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Amcache_DriveBinaries.csv\"", con)
		df['FileKeyLastWriteTimestamp'] = pd.to_datetime(df['FileKeyLastWriteTimestamp'])
		df['DriverTimeStamp'] = pd.to_datetime(df['DriverTimeStamp'])
		df['DriverLastWriteTime'] = pd.to_datetime(df['DriverLastWriteTime'])
		xlsx_preparer(df,writer,'Amcache drivers')
	except Exception as e:
		pass
	

def installed_programs(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Amcache_ProgramEntries.csv\"", con)
		df['FileKeyLastWriteTimestamp'] = pd.to_datetime(df['FileKeyLastWriteTimestamp'])
		df['InstallDateArpLastModified'] = pd.to_datetime(df['InstallDateArpLastModified'])
		df['InstallDate'] = pd.to_datetime(df['InstallDate'])
		df['InstallDateMsi'] = pd.to_datetime(df['InstallDateMsi'])
		df['InstallDateFromLinkFile'] = pd.to_datetime(df['InstallDateFromLinkFile'])
		xlsx_preparer(df,writer,'Amcache installed programs')
	except Exception as e:
		pass
	

def shortcuts(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"lnk.csv\"", con)
		df['SourceCreated'] = pd.to_datetime(df['SourceCreated'])
		df['SourceModified'] = pd.to_datetime(df['SourceModified'])
		df['SourceAccessed'] = pd.to_datetime(df['SourceAccessed'])
		df['TargetCreated'] = pd.to_datetime(df['TargetCreated'])
		df['TargetModified'] = pd.to_datetime(df['TargetModified'])
		df['TargetAccessed'] = pd.to_datetime(df['TargetAccessed'])
		df['TrackerCreatedOn'] = pd.to_datetime(df['TrackerCreatedOn'])
		xlsx_preparer(df,writer,'Shortcuts')
	except Exception as e:
		pass
	

def service_crash(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"System.evtx.csv\" WHERE (EventId=7034)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'Service crash')
	except Exception as e:
		pass
	

def application_error(con,writer):
	try:
		df = pd.read_sql_query("SELECT * from \"Application.evtx.csv\" WHERE (EventId=1000) OR (EventId=1001) OR (EventId=1002)", con)
		df['TimeCreated'] = pd.to_datetime(df['TimeCreated'])
		xlsx_preparer(df,writer,'Application error')
	except Exception as e:
		pass
	
	

def registry_program_execution(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"registry.csv\" WHERE \"Category\" LIKE 'Program Execution'", con)
		df['LastWriteTimestamp'] = pd.to_datetime(df['LastWriteTimestamp'])
		xlsx_preparer(df,writer,'Program execution (registry)')
	except Exception as e:
		pass


def registry_network_shares(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"registry.csv\" WHERE \"Category\" LIKE 'Network Shares'", con)
		df['LastWriteTimestamp'] = pd.to_datetime(df['LastWriteTimestamp'])
		xlsx_preparer(df,writer,'Network shares (registry)')
	except Exception as e:
		pass

def registry_system_info(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"registry.csv\" WHERE \"Category\" LIKE 'System Info'", con)
		df['LastWriteTimestamp'] = pd.to_datetime(df['LastWriteTimestamp'])
		xlsx_preparer(df,writer,'System info')
	except Exception as e:
		pass


def registry_third_party_applications(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"registry.csv\" WHERE \"Category\" LIKE 'Third Party Applications'", con)
		df['LastWriteTimestamp'] = pd.to_datetime(df['LastWriteTimestamp'])
		xlsx_preparer(df,writer,'3rd party applications')
	except Exception as e:
		pass


def registry_user_activity(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"registry.csv\" WHERE \"Category\" LIKE 'User Activity'", con)
		df['LastWriteTimestamp'] = pd.to_datetime(df['LastWriteTimestamp'])
		xlsx_preparer(df,writer,'User Activity (registry)')
	except Exception as e:
		pass

def registry_vss_info(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"registry.csv\" WHERE \"Category\" LIKE 'Volume Shadow Copies'", con)
		df['LastWriteTimestamp'] = pd.to_datetime(df['LastWriteTimestamp'])
		xlsx_preparer(df,writer,'VSS info (registry)')
	except Exception as e:
		pass



def srum_network_usages(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"SrumECmd_NetworkUsages_Output.csv\"", con)
		df['Timestamp'] = pd.to_datetime(df['Timestamp'])
		xlsx_preparer(df,writer,'Network usage')
	except Exception as e:
		pass


def srum_app_resource_use(con,writer):
	try:
		df = pd.read_sql_query("SELECT * FROM \"SrumECmd_AppResourceUseInfo_Output.csv\"", con)
		df['Timestamp'] = pd.to_datetime(df['Timestamp'])
		xlsx_preparer(df,writer,'Apps resource use')
	except Exception as e:
		pass



def create_xlsx_report(db_path, output_report, Logger):
	con = sqlite3.connect(db_path)
	Logger.info("Creating XLSX report.")
	with pd.ExcelWriter(output_report, engine="xlsxwriter") as writer:  
		workbook = writer.book

		####### Bookmarking colors #########
		# Light red fill for "evidence"
		format1 = workbook.add_format({'bg_color':   '#FFC7CE'})

		# Light yellow fill for "of interest"
		format2 = workbook.add_format({'bg_color':   '#FFEB9C'})

		# Green fill for "bookmark"
		format3 = workbook.add_format({'bg_color':   '#C6EFCE'})
		####################################

		cleared_audit_log(con,writer)
		logon_events(con, writer)
		account_management_events(con, writer)
		scheduled_tasks_events(con, writer)
		RDP_events(con, writer)
		application_installation_events(con, writer)
		services_events(con, writer)
		WMI_events(con, writer)
		PowerShell_events(con, writer)
		Process_events(con, writer)
		Network_Share_events(con, writer)
		amcache_program_presence(con,writer)
		recyclebin(con,writer)
		prefetch_program_execution(con,writer)
		registry_program_execution(con,writer)
		amcache_plugged_devices(con,writer)
		drivers(con,writer)
		installed_programs(con,writer)
		shortcuts(con,writer)
		registry_program_execution(con,writer)
		registry_system_info(con,writer)
		registry_network_shares(con,writer)
		registry_third_party_applications(con,writer)
		srum_network_usages(con,writer)


def convert_target(target_root,output_dir,db_path):
	evtx_folder = target_root+"\\Windows\\System32\\winevt\\logs"
	prefetch_folder = target_root+"\\Windows\\prefetch"
	amcache_folder = target_root+"\\Windows\\AppCompat\\Programs"
	appcompatcache_folder = target_root+"\\Windows\\System32\\Config"
	usnjournal_folder = target_root+"\\$Extend"
	lnk_folder = target_root
	recyclebin_folder = target_root
	srum_folder = target_root
	registries_folder = target_root

	Logger.info("Converting artefacts to CSV.")

	evtx_to_csv(evtx_folder, output_dir)
	#evtx_to_csv_partial(evtx_folder, output_dir)
	prefetch_to_csv(prefetch_folder, output_dir)
	amcache_to_csv(amcache_folder,output_dir)
	appcompatcache_to_csv(appcompatcache_folder,output_dir)
	usnjournal_to_csv(usnjournal_folder, output_dir)
	lnk_to_csv(lnk_folder, output_dir)
	recyclebin_to_csv(recyclebin_folder, output_dir)
	srum_to_csv(srum_folder,output_dir)
	registries_to_csv(registries_folder, output_dir)
	remove_whitespaces_filename(output_dir)

	#create_database(output_dir, db_name)
	create_database_alt(output_dir, db_path)

def dir_path(string):
	if os.path.isdir(string):
		return string
	else:
		raise NotADirectoryError(string)


########################################################################  MAIN  ####################################################################""

def main():

	parser = argparse.ArgumentParser(description="Win_Dissect by Martendal. It's just an orchestrator of Eric Zimmerman's tools that can be launched to convert a KAPE acquisition to CSV and to a XLSX \"portable case\".", formatter_class=argparse.ArgumentDefaultsHelpFormatter)
	parser.add_argument("-t", "--target", required=False, help="Absolute path to the target folder containing the OS to parse.", default="C:\\", type=dir_path)
	parser.add_argument("-o", "--output", required=False, help="Absolute path to the output folder. The folder must exist.", default=os.getcwd()+"\\OUTPUT", type=dir_path)
	parser.add_argument("-n", "--name", required=False, help="Name of the target.", default="windissect_output")
	args = parser.parse_args()
	config = vars(args)
	Logger.info(config)

	target_root = config['target']
	output_dir = config['output']
	hostname = config['name']
	db_name = hostname+".db"
	db_path = output_dir+"\\"+db_name
	output_report = output_dir+"\\"+hostname+".xlsx"


	Logger.info("Launching Win_Dissect, go make yourself a coffee.")
	Logger.info("Checking if the specified output directory exists.")
	if(not os.path.exists(output_dir)):
		Logger.info("Creating output directory.")
		os.makedirs(output_dir)

	convert_target(target_root,output_dir,db_path)

	create_xlsx_report(db_path, output_report,Logger)


if __name__ == "__main__":
	main()