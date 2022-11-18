# DFIR-tools

A small collection of tools related to digital forensics (_updated little by little_)

## Lazy_setup.ps1

#### Description
A very simple PowerShell script that should help you to install what is required to use the tools provided in this repository. 

You can also read it just to follow what it contains if you want to install everything by hand, it's pretty straightforward.

DISCLAIMER: Since this script downloads and installs stuff from Internet, maybe it could trigger alerts depending on your environment. So it may be better to just read the content of the script as a guideline to install everything manually. 


## Win-Dissect

#### Description
A script that launches Eric Zimmerman's tools (https://ericzimmerman.github.io/#!index.md) to parse Windows artefacts to CSV, and creates a SQLite database as well as a XLSX "portable case" based on some key artefacts. 

It is primarily designed to be launched on KAPE acquisitions in the beginning of an analysis in order to get artefacts in a common format.

#### Requirements
Python 3 and Pandas + Eric Zimmerman tools.

For easy setup, you can use the Lazy_setup.ps1 script provided in this repository.

In order to run properly, Eric Zimmerman tools need to be in a Utils folder located in the folder from where you will execute the script (reason is the script will look for ./Utils). To setup all the tools properly in this Utils folder, you can download Eric Zimmerman's "Get-ZimmermanTools" script from https://f001.backblazeb2.com/file/EricZimmermanTools/Get-ZimmermanTools.zip, unzip it in Utils and run it. All the tools will be downloaded in the Utils folder and it should be OK.

#### Usage
From CLI, go in the folder where you put win_dissect.py and where you also have the Utils folder described above and run:
```
python win_dissect.py -t [absolute path to target root] -o [absolute path to existing output folder] -n [target name]
```
IMPORTANT: The target must respect the normal windows folders hierarchy with correct folders and subfolders (ex: evtx in C:/Windows/system32/winevt/logs). The target path must be to the root folder of the acquired system (ex: "C" root folder of a KAPE acquisition)

DISCLAIMER: Don't rely on the XLSX report too much, it only contains some specific artefacts that are commonly used. It is just provided as a way to have a "portable case" that can be shared and that should contain useful data.

## VolXLSX

#### Description
A simple script relying on Volatility 3 and Pandas. It runs a set of Volatility 3 plugins against a Windows memory dump and puts all the results in a single XLSX file, so the analysis can be done through a "portable case" and with Excel filtering capacity (also it's convenient for copy/pasting into a timeline maybe?).
There is also a "tag" column on each table so interesting information can be bookmarked.

#### Requirements
Volatility 3 needs to be available, built and installed following the available procedure: https://github.com/volatilityfoundation/volatility3
To make sure that everything is alright, a simple test consisting on running "vol.exe --help" from PowerShell CLI should do the trick. 

Runs on Windows only (tested on Win10). 
Only parses Windows memory dumps (tested on dumps made with WinPMem).

You need to install the Pandas library:

```
pip install pandas
```


#### Usage

```
python3 volxlsx.py -t [absolute path to memory dump] -o [absolute path to output file]
```

For now, only the following modules are run (probably more to come later, let's see):
- windows.pslist.PsList
- windows.psscan.PsScan
- windows.netscan.NetScan
- windows.cmdline.CmdLine
- windows.getsids.GetSIDs
- windows.privileges.Privs
- windows.ssdt.SSDT
- windows.mutantscan.MutantScan
- windows.driverscan.DriverScan
- windows.driverirp.DriverIrp
- windows.dlllist.DllList
- timeliner.Timeliner

DISCLAIMER: Since memory parsing is somehow not the most stable process in the IT world, you can expect that some modules will fail depending on the memory dump. So be careful with the results. You can check the console to see where Volatility encountered errors. It will either skip a whole module or you will get partial results (if the last line of a table is empty on Excel, you can assume that the module crashed before the end).
