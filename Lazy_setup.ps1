<#
    Setup script to install the environment required for win_dissect and volxlsx.
#>


# .NET 4.6.2 check
if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 394802) {
    Write-Output(".NET 4.6.2 or later seems to be installed.")
}
else{
    Write-Output("Please install .NET 4.6.2 or later. See https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net462-web-installer")
    Start-Process "https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net462-web-installer"
}

mkdir ./Utils

#Downloading Get-ZimmermanTools script
Write-Output("Downloading Get-Zimmerman script.")
Invoke-WebRequest -Uri "https://f001.backblazeb2.com/file/EricZimmermanTools/Get-ZimmermanTools.zip" -OutFile ./Utils/Get-ZimmermanTools.zip
if($PSVersionTable.PSVersion.Major -eq 5){
    Expand-Archive ./Utils/Get-ZimmermanTools.zip -DestinationPath ./Utils
    Remove-Item -Path ./Utils/*.zip
}
else{
    Write-Output("PowerShell version is too low, please uncompress ./Utils/Get-ZimmermanTools.zip manually to ./Utils")
}


#Checking python 3
Write-Output("Checking if python 3 is installed.")
if((python --version | Select-String "3.")) {
    Write-Output("Python 3 seems to be installed.")
    Write-Output("Installing required dependencies through pip.")
    pip3 install -r requirements.txt
}
else{
    Write-Output("Please install Python 3. See https://www.python.org/downloads/")
    Start-Process "https://www.python.org/downloads/"
}


#Installing volatility 3
Write-Output("Checking if Volatility 3 is installed.")
if((vol.exe --help | Select-String "Volatility 3 Framework")){
    Write-Output("Volatility 3 seems to be installed.")
}
else{
    Write-Output("Downloading Volatility 3 from GitHub.")
    Invoke-WebRequest -Uri "https://github.com/volatilityfoundation/volatility3/archive/refs/heads/develop.zip" -OutFile ./Utils/volatility.zip
    if($PSVersionTable.PSVersion.Major -eq 5){
        Expand-Archive ./Utils/volatility.zip -DestinationPath ./Utils
        Remove-Item -Path ./Utils/*.zip
        pip3 install -r ./Utils/volatility3-develop/requirements.txt
        python ./Utils/volatility3-develop/setup.py build
        python ./Utils/volatility3-develop/setup.py install
    }
    else{
        Write-Output("PowerShell version is too low, please uncompress ./Utils/volatility.zip manually to ./Utils/Volatility and follow installation guide from github")
    }
}

#Downloading all Zimmerman's tools
./Utils/Get-ZimmermanTools.ps1 -NetVersion 4 -Dest ./Utils
