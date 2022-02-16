### Author:
# Jim Wolff @ https://github.com/JimWolff/
### Credit:
# List of redist data by kayleeFrye_onDeck and others:
# https://stackoverflow.com/a/34209692/665879
#
# List of arguments for silently installing by Rens Hollanders:
# https://social.technet.microsoft.com/Forums/en-US/0408f4e0-0f06-4435-82e6-bb4d3ad38357/silent-installs?forum=mdt
#
#

function Test-RegistryValue {
param (
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)
    try {
    Get-ItemProperty -Path "$($Path+$Value)" -ErrorAction Stop | Out-Null # | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Install-VSRedistPackages {
param (
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$VersionList
)
	$list = @{
	'mscpp2005x64' = [PSCustomObject]@{Name = 'mscpp2005x64'; Configuration = 'x64'; SilentInstallArgs = '/q'; Version = '6.0.2900.2180'; Description = 'Microsoft Visual C++ 2005 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'; RegValue = '1af2a8da7e60d0b429d7e6453b3d0182'; DownloadUrl = 'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.EXE'}
	'mscpp2005x86' = [PSCustomObject]@{Name = 'mscpp2005x86'; Configuration = 'x86'; SilentInstallArgs = '/q'; Version = '6.0.2900.2180'; Description = 'Microsoft Visual C++ 2005 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'; RegValue = 'c1c4f01781cc94c4c8fb1542c0981a2a'; DownloadUrl = 'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x86.EXE'}

	# missing info on 2008

	# (Actual $Version data in registry: 0x9007809 [DWORD])
	'mscpp2008sp1x64' = [PSCustomObject]@{Name = 'mscpp2008sp1x64'; Configuration = 'x64'; SilentInstallArgs = '/q'; Version = '9.0.30729.6161'; Description = 'Microsoft Visual C++ 2008 Redistributable (SP1) (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'; RegValue = '67D6ECF5CD5FBA732B8B22BAC8DE1B4D'; DownloadUrl = 'https://download.microsoft.com/download/2/d/6/2d61c766-107b-409d-8fba-c39e61ca08e8/vcredist_x64.exe'}
	'mscpp2008sp1x86' = [PSCustomObject]@{Name = 'mscpp2008sp1x86'; Configuration = 'x86'; SilentInstallArgs = '/q'; Version = '9.0.30729.6161'; Description = 'Microsoft Visual C++ 2008 Redistributable (SP1) (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'; RegValue = '6E815EB96CCE9A53884E7857C57002F0'; DownloadUrl = 'https://download.microsoft.com/download/d/d/9/dd9a82d0-52ef-40db-8dab-795376989c03/vcredist_x86.exe'}

	'mscpp2010x64' = [PSCustomObject]@{Name = 'mscpp2010x64'; Configuration = 'x64'; SilentInstallArgs = '/q'; Version = '10.0.40219.325'; Description = 'Microsoft Visual C++ 2010 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'; RegValue = '1926E8D15D0BCE53481466615F760A7F'; DownloadUrl = 'https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe'}
	'mscpp2010x86' = [PSCustomObject]@{Name = 'mscpp2010x86'; Configuration = 'x86'; SilentInstallArgs = '/q'; Version = '10.0.40219.325'; Description = 'Microsoft Visual C++ 2010 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'; RegValue = '1D5E3C0FEDA1E123187686FED06E995A'; DownloadUrl = 'https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe'}

	## version caveat: Per user Wai Ha Lee's findings, "...the binaries that come with VC++ 2012 update 4 (11.0.61030.0) have version 11.0.60610.1 for the ATL and MFC binaries, and 11.0.51106.1 for everything else, e.g. msvcp110.dll and msvcr110.dll..."
	'mscpp2012x64' = [PSCustomObject]@{Name = 'mscpp2012x64'; Configuration = 'x64'; SilentInstallArgs = '/quiet /norestart'; Version = '11.0.61030.0'; Description = 'Microsoft Visual C++ 2012 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\'; RegValue = '{ca67548a-5ebe-413a-b50c-4b9ceb6d66c6}'; DownloadUrl = 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe'}
	'mscpp2012x86' = [PSCustomObject]@{Name = 'mscpp2012x86'; Configuration = 'x86'; SilentInstallArgs = '/quiet /norestart'; Version = '11.0.61030.0'; Description = 'Microsoft Visual C++ 2012 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\'; RegValue = '{33d1fd90-4274-48a1-9bc1-97e33d9c2d6f}'; DownloadUrl = 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe'}

	'mscpp2013x64' = [PSCustomObject]@{Name = 'mscpp2013x64'; Configuration = 'x64'; SilentInstallArgs = '/quiet /norestart'; Version = '12.0.30501.0'; Description = 'Microsoft Visual C++ 2013 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\'; RegValue = '{050d4fc8-5d48-4b8f-8972-47c82c46020f}'; DownloadUrl = 'https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe'}
	'mscpp2013x86' = [PSCustomObject]@{Name = 'mscpp2013x86'; Configuration = 'x86'; SilentInstallArgs = '/quiet /norestart'; Version = '12.0.30501.0'; Description = 'Microsoft Visual C++ 2013 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\'; RegValue = '{f65db027-aff3-4070-886a-0d87064aabb1}'; DownloadUrl = 'https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x86.exe'}

	'mscpp2015x64' = [PSCustomObject]@{Name = 'mscpp2015x64'; Configuration = 'x64'; SilentInstallArgs = '/quiet /norestart'; Version = '14.0.24215'; Description = 'Microsoft Visual C++ 2015 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\'; RegValue = '{d992c12e-cab2-426f-bde3-fb8c53950b0d}'; DownloadUrl = 'https://download.microsoft.com/download/6/A/A/6AA4EDFF-645B-48C5-81CC-ED5963AEAD48/vc_redist.x64.exe'}
	'mscpp2015x86' = [PSCustomObject]@{Name = 'mscpp2015x86'; Configuration = 'x86'; SilentInstallArgs = '/quiet /norestart'; Version = '14.0.24215'; Description = 'Microsoft Visual C++ 2015 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\'; RegValue = '{e2803110-78b3-4664-a479-3611a381656a}'; DownloadUrl = 'https://download.microsoft.com/download/6/A/A/6AA4EDFF-645B-48C5-81CC-ED5963AEAD48/vc_redist.x86.exe'}

	'mscpp2017x64' = [PSCustomObject]@{Name = 'mscpp2017x64'; Configuration = 'x64'; SilentInstallArgs = '/quiet /norestart'; Version = '14.16.27012'; Description = 'Microsoft Visual C++ 2017 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,redist.x64,amd64,14.16,bundle\Dependents\'; RegValue = '{427ada59-85e7-4bc8-b8d5-ebf59db60423}'; DownloadUrl = 'https://download.visualstudio.microsoft.com/download/pr/9fbed7c7-7012-4cc0-a0a3-a541f51981b5/e7eec15278b4473e26d7e32cef53a34c/vc_redist.x64.exe'}
	'mscpp2017x86' = [PSCustomObject]@{Name = 'mscpp2017x86'; Configuration = 'x86'; SilentInstallArgs = '/quiet /norestart'; Version = '14.16.27012'; Description = 'Microsoft Visual C++ 2017 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,redist.x64,x86,14.16,bundle\Dependents\'; RegValue = '{67f67547-9693-4937-aa13-56e296bd40f6}'; DownloadUrl = 'https://download.visualstudio.microsoft.com/download/pr/d0b808a8-aa78-4250-8e54-49b8c23f7328/9c5e6532055786367ee61aafb3313c95/vc_redist.x86.exe'}

	## update from https://stackoverflow.com/a/34209692/665879 on 2022-02-16
	'mscpp2019x64' = [PSCustomObject]@{Name = 'mscpp2019x64'; Configuration = 'x64'; SilentInstallArgs = '/quiet /norestart'; Version = '14.25.28508'; Description = 'Microsoft Visual C++ 2015-2019 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,redist.x64,amd64,14.25,bundle\Dependents\'; RegValue = '{6913e92a-b64e-41c9-a5e6-cef39207fe89}'; DownloadUrl = 'https://aka.ms/vs/16/release/vc_redist.x64.exe'}
	'mscpp2019x86' = [PSCustomObject]@{Name = 'mscpp2019x86'; Configuration = 'x86'; SilentInstallArgs = '/quiet /norestart'; Version = '14.25.28508'; Description = 'Microsoft Visual C++ 2015-2019 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,redist.x64,x86,14.25,bundle\Dependents\'; RegValue = '{65e650ff-30be-469d-b63a-418d71ea1765}'; DownloadUrl = 'https://aka.ms/vs/16/release/vc_redist.x86.exe'}

	'mscpp2022x64' = [PSCustomObject]@{Name = 'mscpp2022x64'; Configuration = 'x64'; SilentInstallArgs = '/quiet /norestart'; Version = '14.31.31103'; Description = 'Microsoft Visual C++ 2015-2022 Redistributable (x64)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,redist.x64,amd64,14.31,bundle\Dependents\'; RegValue = '{2aaf1df0-eb13-4099-9992-962bb4e596d1}'; DownloadUrl = 'https://aka.ms/vs/17/release/vc_redist.x64.exe'}
	'mscpp2022x86' = [PSCustomObject]@{Name = 'mscpp2022x86'; Configuration = 'x86'; SilentInstallArgs = '/quiet /norestart'; Version = '14.31.31103'; Description = 'Microsoft Visual C++ 2015-2022 Redistributable (x86)'; RegPath = 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,redist.x64,x86,14.31,bundle\Dependents\'; RegValue = '{41d7b770-418a-43b7-95a5-f925fff05789}'; DownloadUrl = 'https://aka.ms/vs/17/release/vc_redist.x86.exe'}
	}

    try {
        $versionkeys = $VersionList -split ","

        foreach ($key in $versionkeys) {
           $redistInfo = $list[$key]
           $redistInstalled = Test-RegistryValue -Path $redistInfo.RegPath -Value $redistInfo.RegValue
           if($redistInstalled -eq $False) {
            Write-Output "$($redistInfo.Description) installing."
            $path = gi -Path "~\Downloads"
            $lastPartOfUrl = $redistInfo.DownloadUrl -split "/" | select-object -Last 1
            Write-Output "Downloading: $($redistInfo.DownloadUrl) to $($path.FullName)"
            $downloadTargetPath = "$path\$lastPartOfUrl"
            Invoke-WebRequest -Uri $redistInfo.DownloadUrl -OutFile $downloadTargetPath
            Write-Output "Running: '$lastPartOfUrl' with arguments: '$($redistInfo.SilentInstallArgs)'"
            Start-Process -FilePath $downloadTargetPath -ArgumentList "$($redistInfo.SilentInstallArgs)" -Wait -NoNewWindow | Wait-Process

            Write-Output "Rechecking registry." 
            $redistInstalledNow = Test-RegistryValue -Path $redistInfo.RegPath -Value $redistInfo.RegValue
            if($redistInstalledNow -eq $True){
                Write-Output "$($redistInfo.Description) was installed succesfully."
            } Else {
                Write-Error "Registry key was not found, installation unsuccessful."
            }
           } Else {
            Write-Output "$($redistInfo.Description) already installed."
           }
        }
    }
    catch {
        return $false
    }
}
Export-ModuleMember Install-VSRedistPackages

# Install-VSRedistPackages -VersionList  "mscpp2005x64,mscpp2005x86,mscpp2008sp1x64,mscpp2008sp1x86,mscpp2010x64,mscpp2010x86,mscpp2012x64,mscpp2012x86,mscpp2013x64,mscpp2013x86,mscpp2015x64,mscpp2015x86,mscpp2017x64,mscpp2017x86"
#Install-VSRedistPackages -VersionList  "mscpp2017x64"
