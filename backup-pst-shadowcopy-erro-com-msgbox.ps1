$testServerOnline = Test-Connection -ComputerName ncomp-wks-02 -Count 3 -Quiet;

If(-Not $testServerOnline) {
    Add-Type -AssemblyName PresentationFramework;
    $msgBox = [System.Windows.MessageBox]::Show('Servidor inacessível','Verificar Servidor Online','OK','Error');
    $msgBox;
    If($msgBox -eq "OK") {
        break;
    }
}

$s1 = (Get-WmiObject -List Win32_ShadowCopy).Create("C:\", "ClientAccessible");
$s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $s1.ShadowID };

$d  = $s2.DeviceObject + "\";

cmd /c mklink /d C:\shadowcopy "$d";

robocopy "C:\shadowcopy\Users\evandro\Documents\Arquivos do Outlook" "C:\Users\evandro\Documents\teste\" /E /ZB;

"vssadmin delete shadows /Shadow=""$($s2.ID.ToLower())"" /Quiet" | iex;

Invoke-Expression -Command "cmd /c rmdir /S /Q C:\shadowcopy";