# PowershellVsRedistInstaller
Powershell script that checks and installs specified version of the c++ redist for windows.

### Details
This script is intended to be used in CI (Continuous Integration) pipelines, if you have a dependency on one of the redist dll's that you want to automatically install on target systems that don't have it.

The script checks if the selected versions are installed (via regedit), then silently installs them if they are not.

### Usage
* Open Powershell as administrator.
* Ensure you have write access to working directory eg. `cd ~\Downloads`
* Run: `iwr -Uri <PathToRawVersionOfScriptOnGitHub> -OutFile .\vscppredistcheck.psm1`
* Run: `Import-Module .\vscppredistcheck.psm1`
* Run: `Install-VSRedistPackages -VersionList "mscpp2012x64,mscpp2015x86"` 

#### -VersionList info
(More detail can be found in the script itself or in this [StackOverflow Answer](https://stackoverflow.com/a/34209692/665879))
* mscpp2005x64
* mscpp2005x86
* mscpp2008sp1x64
* mscpp2008sp1x86
* mscpp2010x64
* mscpp2010x86
* mscpp2012x64
* mscpp2012x86
* mscpp2013x64
* mscpp2013x86
* mscpp2015x64
* mscpp2015x86
* mscpp2017x64
* mscpp2017x86
