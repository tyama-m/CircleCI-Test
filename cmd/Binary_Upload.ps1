Param( [string]$ParentName = "",
       [string]$CommitKey = "",
       [string]$AccountName = "",
       [string]$Pword = "" )


Write-host $ParentName -f Green
Write-host $AccountName -f Green
Write-host $Pword -f Green

exit

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

$accountName = $AccountName
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
# $Commit = $CommitKey.Substring(0, 7)
# $UploadFileName = "main_" + $Commit + ".zip"
# Write-host $UploadFileName -f Green
$UploadFileName = "MorisawaDesktopManager_" + $(Get-Date).ToString("yyyyMMddHHmmss") + ".zip"

# ファイルを追加する
$FileStream = ([System.IO.FileInfo] (Get-Item "D:/GitHub/CircleCI-Test/main.zip")).OpenRead()

$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL = $UploadFileName
$FileUploaded = $folder.Files.Add($FileCreationInfo)

$ctx.Load($FileUploaded)
$ctx.ExecuteQuery()
 
$FileStream.Close()

$ctx.Dispose()