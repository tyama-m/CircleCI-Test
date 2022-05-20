#Parameters
$DownloadURL = "https://download.microsoft.com/download/B/3/D/B3DA6839-B852-41B3-A9DF-0AFA926242F2/sharepointclientcomponents_16-6906-1200_x64-en-us.msi"
$Assemblies= @(
        "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll",
        "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    )
 
#Check if all assemblies given in the list are found
$FileExist = $True
ForEach ($File in $Assemblies)
{
    #Check if CSOM Assemblies are Found
    If(!(Test-Path $File))
    {
        $FileExist = $False; Break;
    }
}
 
#Download and Install CSOM Assemblies
If(!$FileExist)
{
    #Download the SharePoint Online Client SDK
    #Write-host "Downloading SharePoint Online Client SDK..." -f Yellow -NoNewline
    $InstallerPath = "$Env:TEMP\SharePointOnlineClientComponents16.msi"
    Invoke-WebRequest $DownloadURL -OutFile $InstallerPath
    #Write-host "Done!" -f Green
     
    #Start Installation
    #Write-host "Installing SharePoint Online Client SDK..." -f Yellow -NoNewline
    Start-Process MSIExec.exe -ArgumentList "/i $InstallerPath /quiet /norestart" -Wait
    #Write-host "Done!" -f Green
}
Else
{
    Write-host "SharePoint Online CSOM assemblies are already installed!" -f Yellow
}
