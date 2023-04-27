<#
    Powershell - Description of Script here

    Return codes:
#>

#region Parameter
<#
    DESCRIPTION
    WE HAVE TO CHCEK THE PARAMETER DELIVERED FROM TWS

    What does Parameters do

#>
#endregion

#region LocalTesting
<# 
#>
#endregion

#region ActionPreference
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
$VerbosePreference     = [System.Management.Automation.ActionPreference]::SilentlyContinue
#endregion
  
#region Script Variables
$ScriptPath            = if ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } elseif ($psIse) { $psISE.CurrentFile.FullPath } elseif (Get-Location) { ((get-Location).path + '\*' | Get-ChildItem -Include *.ps1 | Select-Object FullName).FullName } else { Write-Error 'Could not get script path!' }
$ScriptFolder          = Split-Path -Path $ScriptPath
$ScriptTemp            = Join-Path  -Path $ScriptFolder -ChildPath 'temp'
$ScriptLogFolder       = Join-Path  -Path $ScriptFolder -ChildPath 'log'
$ScriptXml             = Join-Path  -Path $ScriptFolder -ChildPath ('{0}.xml' -f [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath))
$ScriptStartDate       = Get-Date
$ScriptTranscript      = $true
#endregion

#region load xml
$xml = New-Object xml
$xml.Load($ScriptXml)


#endregion

#region Start Transcript
if ($ScriptTranscript)
{
  $ScriptLogFile = Join-Path $ScriptLogFolder ('{0}_{1:yyyy-MM-dd_HHmmss.fff}.txt' -f [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath), $ScriptStartDate)
    
  if (-not (Test-Path $ScriptLogFolder))
  {
    $null = mkdir $ScriptLogFolder
  }
  
  Start-Transcript -LiteralPath $ScriptLogFile  
}
#endregion

#Trace-Output für Transcript
function Trace-Output($txt)
{
  <# 

      This is only for Display in SARA

  #>
  
  Write-Host (Get-Date) : $txt
  return

}

#################################################################### SCRIPT MAIN ####################################################################

## Server Base Lookup ##

$searchbase = $xml.Values.searchbase
$not_like_servernames = $xml.Values.excluded_servers

## mail ##

$smtpserver = $xml.Values.smtp
$mailfrom = $xml.Values.sender

## filepath for .xlsx ##
$filepath = $xml.Values.filepath


#PSObject Output fuer Custom PS Object mit Werten fuer in die Tabelle und zum Versenden der Mails
$output = @()

#Hole alle Server welche die Eigenschaft haben, Enabled, OperatingSystem wie Windows Server und nicht in der liste von $not_like_serveranems sind, sortieren, und nur die Objecte Name,OS,Mangedby selektieren.
$servers = Get-ADComputer -SearchBase $searchbase -Filter {(OperatingSystem -like "*windows*server*") -and (Enabled -eq "True") -and (name -notlike $not_like_servernames)} -Properties * | Sort Name | select -Unique Name, OperatingSystem, managedby
$servercount = $servers.Count
Trace-Output -txt ("Es wurden {0} gefunden" -f $servercount)

#Function fuer OUPfad (Rueckgabewert aus Managedby) abfragen ob User oder Gruppe und Name sowie Mail des Objekt zurÃ¼ckgeben.
function check_ManagedBy()
{

    param(
        [string]$OUPfad
        )
    
    if($OUPfad -ne $null)
    {  
      $isuser = (Get-ADObject -Filter 'ObjectClass -eq "User"' -SearchBase $OUPfad -Properties Name,Mail)
      if($isuser -eq $null)
        {
            $isuser = (Get-ADObject -Filter 'ObjectClass -eq "Group"' -SearchBase $OUPfad -Properties Name,Mail)

        }
        return $isuser
    }
    else
    {
    $noentry = "No Entry in ManagedBy"
    return $noentry
    }
    
}

#Prüfen ob die Objekte in $output älter als 30Tage sind, wenn ja eine E-Mail an die Aufgelöste E-Mail aus der OU von Managedby.
function check_30days()
{
    [CmdletBinding()]
    param (
        [psobject]$data,
        [string]$smtpserver,
        [string]$mailfrom
    )

    foreach($item in $data){
        $Mail = $item.Mail
        $ServerName = $item.ServerName
        $Patches = $item.Patches
        
        $LastInstallationSuccessDate = $item.LastInstallationSuccessDate
        if($LastInstallationSuccessDate -ne $Null){
            $FormatedDate = Get-Date($LastInstallationSuccessDate) -Format dd.MM.yyyy
        }
        else
        {
            $FormatedDate = "No Date found"
        }        
        $mailsubject = "Warnung - Patchstand von {0} aelter als 30 Tage" -f $ServerName
        $body = @"
        Patches auf Server: {0} wurden das letzte mal installiert am {1}

        {2}
"@ -f $ServerName,$FormatedDate,$Patches

        if($LastInstallationSuccessDate -lt (Get-Date).AddDays(-30).Date)
        {
            try
            {
                Write-Host "Mail send"
                Send-MailMessage -From $mailfrom -To $Mail <#Hier $Mail einfügen od. TestEMail "" #>  -Subject $mailsubject -smtpserver $smtpserver -body $body
                Trace-Output -txt ("Mail für Server {0} gesendet" -f $item.ServerName)
            }
            catch
            {
                Trace-Output ("Fehler bei Mail versand für Server: {0}" -f $item.ServerName)
            }
        }

    }
}


#Gehe durch alle objects in $servers und mache ein invoke command auf dem dem $server und suche nach dem patchstandt
foreach ($server in $servers){

    Trace-Output ("Server:             {0}" -f $server.name)

    $porttest = (Test-NetConnection $server.Name -Port 5985 -InformationLevel "Detailed")

    if(Test-Connection $server.Name)
    {
        Trace-Output "Ping ok"
        if($porttest.TcpTestSucceeded)
        {   
            Trace-Output "RemoteShell ist erreichbar"
            $WindowsUpdateInfo = Invoke-Command -ComputerName $server.name -ScriptBlock{
                (New-Object -com "Microsoft.Update.AutoUpdate").Results}

            $patch = Invoke-Command -ComputerName $server.Name -ScriptBlock {

            $updateInfoMsg = "Windows Update Status: `n";
    
            $UpdateSession = New-Object -ComObject Microsoft.Update.Session;
            $UpdateSearcher = $UpdateSession.CreateupdateSearcher();
            $Updates = @($UpdateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'").Updates);
            $Found = ($Updates | Select-Object -Expand Title);
    
            If ($Found -eq $Null) {
                $updateInfoMsg += "Up to date";
            } Else {
                $Found = ($Updates | Select-Object -Expand Title) -Join "`n";
                $updateInfoMsg += "Updates available:`nema"
                $updateInfoMsg += $Found;
            }

            Return $updateInfoMsg;
            }

            $check = check_ManagedBy -OUPfad $server.managedby
            Trace-Output ("Wird verwaltet von: {0}" -f $check.Name)
        
            $output += New-Object -TypeName psobject -Property @{
                ServerName = $server.Name
                ServerOS = $server.OperatingSystem
                ManagedBy = $check.Name
                Mail = $check.Mail
                Patches = $patch
                LastSearchSuccessDate = $WindowsUpdateInfo.LastSearchSuccessDate
                LastInstallationSuccessDate = $WindowsUpdateInfo.LastInstallationSuccessDate
            }
        }
        else {
             Trace-Output "RemoteShell auf Port 5985 nicht erreichbar"
             Trace-Output ""
        }
    }
    else {
        Trace-Output "Server per Ping nicht erreichbar"
        Trace-Output ""
    }
    
}

check_30days -data $output -smtpserver $smtpserver -mailfrom $mail

try {
    $output | Export-Excel -Path ("{0}updatelist{1}.xlsx" -f $filepath, $(get-date -f yyyy-MM-dd))
    Trace-Output ("Excel Liste wurde in den Pfad: {0} geschrieben" -f $filepath)
}
catch {
    Trace-Output "Fehler beim schreiben der Excle Liste"
}


#region Stop Transcript
if ($ScriptTranscript)
{
  Stop-Transcript
}
#endregion
Exit $exitcode