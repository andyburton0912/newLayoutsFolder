$Host.UI.RawUI.WindowTitle = 'New Layouts'

Stop-Process -name cmd

[string]$SelectFont = Read-Host "Select Font"
[float]$SelectSize = Read-Host "
Select Size"

[float]$FontInvCen = $SelectSize + 3

[string]$Folder = -Join ("$SelectFont" + " " + "$SelectSize")
$test = Test-Path "I:\Shared installs - Non-Confidential\Install Files\DGL - Master Layouts\Layout templates\Single consultant layouts\$Folder\"

if($test -eq $False) {

Write-Host "
Creating Font..." -ForegroundColor Yellow

Copy-Item -Path "I:\Shared installs - Non-Confidential\Install Files\DGL - Master Layouts\Layout templates\Single consultant layouts\Arial 10" -Destination "C:\Users\andy.burton\Desktop\Layouts\$Folder\" -Recurse

$Files = @(gci -path "C:\Users\andy.burton\Desktop\Layouts\$Folder\layouts" -name -Exclude "*.jpg","InvCen.doc")

$Word = New-Object -ComObject Word.Application
$Word.visible = $false

foreach($file in $Files)
{

$Doc = $word.Documents.Open("C:\Users\andy.burton\Desktop\Layouts\$Folder\Layouts\$File")
$Selection = $word.Selection
$Doc.Select()
$Selection.Font.Name = $SelectFont
$Selection.Font.Size = $SelectSize
$Doc.Close()

}
$Doc = $word.Documents.Open("C:\Users\andy.burton\Desktop\Layouts\$Folder\Layouts\InvCen.doc")
$Selection = $word.Selection
$Doc.Select()
$Selection.Font.Name = $SelectFont
$Selection.Font.Size = $FontInvCen
$Doc.Close()
$Word.Quit()

Write-Host "
Layouts Folder Created
" -ForegroundColor Green 
Start-Sleep -s 3
Stop-Process -name powershell
}
Else {Write-Host " "
Write-Host "Folder Already Exists!!!" -ForegroundColor Red
Start-Sleep -s 1
Write-Host "
Closing..." -ForegroundColor Red} 
Start-Sleep -s 3
Stop-Process -name powershell
