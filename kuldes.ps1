# Beállítások
$downloadPath = "C:\Users\BaloghI\Downloads"
$fileName = "postakonyv.xml"
$fullPath = Join-Path $downloadPath $fileName

# Ellenőrzés: létezik-e a fájl

if (-not (Test-Path $fullPath)) {
    Write-Host "Hiba! A fájl nem található: $fullPath"
    Read-Host "Nyomd le az ENTER-t a kilépéshez!"
    exit
}

# Dátum: hónap (HH) és nap (NN)
$month = (Get-Date).ToString("MM")
$day   = (Get-Date).ToString("dd")
$newName = "postakonyv$month$day.xml"
$newPath = Join-Path $downloadPath $newName

# Átnevezés
Rename-Item -Path $fullPath -NewName $newName -Force

# Outlook COM objektum
$Outlook = New-Object -ComObject Outlook.Application
$mail = $Outlook.CreateItem(0)  # 0 = MailItem

# Beállítások
$mail.To = "efeladas@posta.hu"
$mail.Subject = "$month$day"
$mail.Attachments.Add($newPath)

# Megnyitás szerkesztőben (nem küldi el automatikusan!)
$mail.HTMLBody = ""
$mail.Display($false)
