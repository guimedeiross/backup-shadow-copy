$testServerOnline = Test-Connection -ComputerName ncomp-wks-02 -Count 3 -Quiet;
$smtpServer = "seu provedor smtp";
$port = port;
$username = "seuemail@email";
$password = "suaSenhadoemail";
$from = "o que vc quer que apareça como";
$to = "pra que email enviar";
$subject = "assunto";
$body = "corpo da mensagem";

If(-Not $testServerOnline) {
    $Message = New-Object System.Net.Mail.MailMessage $from,$to;
    $Message.IsBodyHTML = $true;
    $Message.Subject = $subject;
    $Message.Body = $body;

    $Smtp = New-Object Net.Mail.SmtpClient($smtpServer,$port);
    $Smtp.Credentials = New-Object System.Net.NetworkCredential($username,$password);
    $Smtp.Send($Message);

    break;
}

$s1 = (Get-WmiObject -List Win32_ShadowCopy).Create("C:\", "ClientAccessible");
$s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $s1.ShadowID };

$d  = $s2.DeviceObject + "\";

cmd /c mklink /d C:\shadowcopy "$d";

robocopy "C:\shadowcopy\Users\evandro\Documents\Arquivos do Outlook" "C:\Users\evandro\Documents\teste\" /E /ZB;

"vssadmin delete shadows /Shadow=""$($s2.ID.ToLower())"" /Quiet" | iex;

Invoke-Expression -Command "cmd /c rmdir /S /Q C:\shadowcopy";