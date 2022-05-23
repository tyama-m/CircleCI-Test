Param(  [string]$ParentName,
        [string]$CommitKey,
        [string]$UserName,
        [string]$Pword,
        [string]$CopyFileName,
        [string]$FileExt)

Write-host "ParentName = " + $ParentName.Substring(0,10) -f Green
Write-host "CommitKey = " + $CommitKey.Substring(0,10) -f Green
Write-host "UserName = " + $UserName.Substring(0,10) -f Green
Write-host "Pword = " + $Pword.Substring(0,10) -f Green
Write-host "CopyFileName = " + $CopyFileName.Substring(0,10) -f Green
Write-host "FileExt = " + $FileExt.Substring(0,3) -f Green

#.NET CSOM モジュールの読み込み
# SharePoint Online Client Components SDK をダウンロードする
# https://www.microsoft.com/en-us/download/details.aspx?id=42038
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null


#SharePointに接続する
$siteUrl = "https://fonts.sharepoint.com/sites/devdep"
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)

$accountName = $UserName
$password = ConvertTo-SecureString -AsPlainText -Force $Pword
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($accountName, $password)
$ctx.Credentials = $credentials

#ドキュメントライブラリに接続する
$folderURL = $siteUrl + "/" + $ParentName
$folder = $ctx.Web.GetFolderByServerRelativeUrl($folderURL)
$ctx.Load($folder)
$ctx.ExecuteQuery()

# フォルダーを追加する
# $newfolder = "0200-1-01_MDM-Installer"
# $subfolder = $folder.Folders.Add($newfolder)
# $ctx.Load($subfolder)
# $ctx.ExecuteQuery()

# アップロードファイル名を作成する
$Commit = $CommitKey.Substring(0, 7)
$UploadFileName = "MorisawaDesktopManager_" + $Commit + "." + $FileExt
# $UploadFileName = "MorisawaDesktopManager_" + $(Get-Date).ToString("yyyyMMddHHmmss") + ".zip"
# Write-host $UploadFileName -f Green

# ファイルを追加する
$FileStream = ([System.IO.FileInfo] (Get-Item $CopyFileName)).OpenRead()

$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL = $UploadFileName
$FileUploaded = $folder.Files.Add($FileCreationInfo)

$ctx.Load($FileUploaded)
$ctx.ExecuteQuery()
 
$FileStream.Close()

$ctx.Dispose()