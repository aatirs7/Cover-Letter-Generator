Copy-Item "Enter Template Path" -Destination "Enter Destination Path"
#$template = Read-Host "Please enter file path of CV Template"
$Company = Read-Host "Please enter the company name"
$title = Read-Host "Please enter the postition title"
$description = Read-Host "Please describe your interest in this company"

# Start Word Object
$Word = New-Object -ComObject Word.Application


# Open Word doc
$OpenFile = $Word.Documents.Open("Enter Template Path")
#$template


# Get the content of the doc
$Content = $OpenFile.Content
# My name is $name and the reason is $reason.

# New variable for new text and variables to to replace the ones from the doc.
$newText = ""
$companyName = $Company
$date =  Get-Date -Format " MM/dd/yyyy"
$title = $title
$description = $description

# Store the current text in the var
$newText = $Content.Text

# Replace the template vars with the new values
$newText = $newText  -replace '\$companyName', $companyName
$newText = $newText  -replace '\$title', $title
$newText = $newText  -replace '\$datee', $date
$newText = $newText  -replace '\$description', $description
# Make the modified text the new content and Save
$Content.Text = $newText
$OpenFile.Save()
$Word.Quit()
$path = "Enter Destination Path"



$wd = New-Object -ComObject Word.Application

Get-ChildItem -Path $path -Include *.doc, *.docx -Recurse |

ForEach-Object {

$doc = $wd.Documents.Open($_.Fullname)

$pdf = $_.FullName -replace $_.Extension, '.pdf'

$doc.ExportAsFixedFormat($pdf,17,$false,0,3,1,1,0,$false, $false,0,$false, $true)

$doc.Close()

}

$wd.Quit()
$cv = $companyName + " Cover Letter.pdf"

Start-Sleep -Seconds .5
Rename-Item "Enter path" $cv 
Remove-Item "Enter path"

Write-Host "Process Complete"
