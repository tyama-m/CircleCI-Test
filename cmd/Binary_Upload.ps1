Param([string]$CommitKey, [string]$UserName, [string]$Pword)

Write-host $CommitKey.Substring(0,10) -f Green
Write-host $UserName.Substring(0,10) -f Green
Write-host $Pword.Substring(0,10) -f Green

#.NET CSOM モジュールの読み込み
# SharePoint Online Client Components SDK をダウンロードする
# https://www.microsoft.com/en-us/download/details.aspx?id=42038
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
# [System.Net.WebRequest]::GetSystemWebProxy()
# [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials


#SharePointに接続する
$siteUrl = "https://fonts.sharepoint.com/sites/devdep"
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)

$accountName = $UserName
$password = ConvertTo-SecureString -AsPlainText -Force $Pword
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($accountName, $password)
$ctx.Credentials = $credentials

#ドキュメントライブラリに接続する
$ParentName = "Shared%20Documents%2F10%5FProject%2F0900%2D0%2D01%5FMP%2DCloud%280008%2C0011%2C0013%29%2FDesktop%20Manager%20%28DM%29%2Ftmp%2FCICD"
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
$UploadFileName = "MorisawaDesktopManager_" + $Commit + ".zip"
# $UploadFileName = "MorisawaDesktopManager_" + $(Get-Date).ToString("yyyyMMddHHmmss") + ".zip"
# Write-host $UploadFileName -f Green

# ファイルを追加する
#$FileStream = ([System.IO.FileInfo] (Get-Item "D:/GitHub/CircleCI-Test/main.zip")).OpenRead() DEBUG用
$FileStream = ([System.IO.FileInfo] (Get-Item "c:/project/main.zip")).OpenRead()

$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL = $UploadFileName
$FileUploaded = $folder.Files.Add($FileCreationInfo)

$ctx.Load($FileUploaded)
$ctx.ExecuteQuery()
 
$FileStream.Close()

$ctx.Dispose()