try{
    Write-Host "Browser wird gestartet" -ForegroundColor Yellow
    $i_1 = New-Object -ComObject 'internetExplorer.Application' -ErrorAction Ignore -ErrorVariable global:Fehler
    $i_1.Visible = $true
    $i_1.Navigate("https://www.ebay-kleinanzeigen.de/m-einloggen.html")
    do{Sleep 1}until($i_1.Busy -eq $false)
}catch{
    Write-Host "Browser konnte nicht gestartet werden" -ForegroundColor Red
    $try++
    sleep 10
}
if($i_1 -ne 0){
    $try = 6
    Write-Host "Browser wurde gestartet" -ForegroundColor Green
}
$cookie = $i_1.Document.IHTMLDocument3_getElementById("gdpr-banner-accept")
$cookie.click()
$email = $i_1.Document.IHTMLDocument3_getElementById("login-email")
$email.value = "" #Email Adresse eingeben
$password = $i_1.Document.IHTMLDocument3_getElementById("login-password")
$password.value = "" #Passwort eingeben
$login = $i_1.Document.IHTMLDocument3_getElementById("login-submit")
$login.click()
$i_1.Navigate("https://www.ebay-kleinanzeigen.de/m-nachrichten.html")
do{Sleep 1}until($i_1.Busy -eq $false)
#$all_chats  = $i_1.Document.body.getElementsByClassName("ConversationListItem")
$new_chats = $i_1.Document.body.getElementsByClassName("ConversationListItem--Header--From--New")
foreach($new_chat in $new_chats){
sleep 1
$new_chat.click()
sleep 1
$chat_1 = $i_1.Document.body.getElementsByClassName("jsx-3433089031")
$chat_2 = $i_1.Document.IHTMLDocument3_getElementById("Reply-Text")
$random_text = Get-Random "Hallo, danke für deine Nachricht und dein damit Verbundenes Interesse an dem Artikel.","Moin, der Artikel ist noch zu haben.","Guten Tag. Danke für die Nachricht! Der Artikel ist neu und ungeöffnet","Guten Tag. Ich melde mich später","Danke für die Nachricht!"
$chat_1[0].innerText = $random_text
sleep 1
$send = $i_1.Document.body.getElementsByClassName("Button-primary")
$send[0].click()
}
