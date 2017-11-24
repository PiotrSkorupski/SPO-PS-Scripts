#Excel static values

Write-Host "Setting variables" -ForegroundColor Gray

$excelFilePath = "C:\<path_to_input_file>\Videos.xlsx"
$errorLogFile = "C:\<path_to_error_file>\errors.txt"

$rowStart = 0
$rowEnd = 205 # n - 1

$rowNum = 2 # Start from second row as first one is header

#Declare column numbers
$groupNameColNum = 1
$groupDescColNum = 2
$videoTitleColNum = 3
$videoDescColNum = 4
$videoLinkColNum = 6
$videoRequiredColNum = 7
$videoSourceColNum = 9

Write-Host "Initializig Excel" -ForegroundColor Gray

$excel = new-object -com excel.application
$workbook = $excel.workbooks.open($excelFilePath)

#SPO Static values

$user = "spouser@tenant.onmicrosoft.com"
$password = "password"
$securePassword = $password | ConvertTo-SecureString -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential ($user, $securePassword)
$siteURL = "https://tenant.sharepoint.com/sites/SPSite/SPWeb"
$listName = "Videos"
$listTitle = "Videos"

#Connect to SPO using PnP PowerShell module
Connect-PnPOnline -Url $siteURL -Credentials $credentials
#Get SPWeb context
$context = Get-PnPContext
$web = $context.Web
$lists = $web.Lists

#Load SPWeb and lists collection
$context.Load($web)
$context.Load($lists)
$context.ExecuteQuery()

Write-Host "Obtained web: " $web.Url

#List all lists just for reference
Write-Host "Obtained lists: "
foreach($list in $web.lists){ 
    Write-Host $list.Title " " $list.Id -ForegroundColor Gray
}

$List = $context.Web.Lists.GetByTitle($listTitle)
$context.Load($List)
$context.ExecuteQuery()

Write-Host "Obtained list: " $List.Title -ForegroundColor Green

#Get first and only sheet in workbook
$sheet = $workbook.Sheets.Item(1);

#Setting permissions for region
Write-Host "Adding videos" -BackgroundColor blue
for ($i =$rowStart; $i -le $rowEnd-1; $i++) {
    $groupName = $sheet.Cells.Item($rowNum+$i, $groupNameColNum).text
    $groupDesc = $sheet.Cells.Item($rowNum+$i, $groupDescColNum).text
    $videoTitle = $sheet.Cells.Item($rowNum+$i, $videoTitleColNum).text
    $videoDesc = $sheet.Cells.Item($rowNum+$i, $videoDescColNum).text
    $videoURL = $sheet.Cells.Item($rowNum+$i, $videoLinkColNum).text
    $videoSource = $sheet.Cells.Item($rowNum+$i, $videoSourceColNum).text

    Write-Host "Veryfying or creating folder " $groupName -ForegroundColor Blue

    Write-Host "Adding video name: " $videoTitle -ForegroundColor Yellow
    Write-Host "To group: " $groupName -ForegroundColor Yellow

    $folder = $groupName

    #Add new list item if Video Item Content Type to list
    #Sets following fields: Title, VideoUrl, VideoSource, VideoDesc, Required, VideoGroup
    #Required is Yes / No field
    #Adds item to folder $groupName using -Folder $folder switch
    Add-PnPListItem -List "Filmy" -ContentType "Video Item" -Values @{"Title"=$videoTitle; "VideoUrl"=$videoURL; "VideoSource"=$videoSource; "VideoDescription"=$videoDesc; "Required"=$false; "VideoGroup"=$groupName} -Folder $folder

    Write-Host "Done" -ForegroundColor Green
}