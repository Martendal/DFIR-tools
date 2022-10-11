# DFIR-tools
"We have DFIR at home"

A collection of small tools related to digital forensics.


## volxlsx
#### Description
A simple script relying on Volatility 3 and Pandas. It runs a set of Volatility 3 plugins against a Windows memory dump and put all the results in a single XLSX file, so the analysis can be done through a "portable case" and with Excel filtering capacity (also it's convenient for copy/pasting into a timeline maybe?).

#### Requirements
Volatility 3 needs to be available and installed following the available procedure: https://github.com/volatilityfoundation/volatility3
To make sure that everything is alright, vol.exe needs to be available from CLI.

Runs on Windows only (tested on Win10). 
Only parses Windows memory dumps (tested on dumps made with WinPMem).

Disclaimer: Since memory parsing is somehow not the most stable process in the IT world, you can expect that some modules will fail depending on the memory dump. So be careful with the results. You can check the console to see where Volatility encountered errors. It will either skip a whole module or you will get partial results (if the last line of a table is empty on Excel, you can assume that the module crashed before the end).


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
