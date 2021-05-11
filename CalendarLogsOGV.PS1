# CONNECT Exchange Online
# [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Register-PSRepository -Default
# install-module ExchangeOnlineManagement -SkipPublisherCheck
# install-module -name PowershellGet -Force -SkipPublisherCheck
# Uninstall-Module PowershellGet -MaximumVersion "1.0.0.1" -Force -Confirm:$false -EA stop

 IF (!(get-accepteddomain -EA silentlycontinue)) { Connect-ExchangeOnline }

$ts = Get-Date -Format yyyyMMdd_hhmmss ; $FormatEnumerationLimit = -1
$DesktopPath = ([Environment]::GetFolderPath('Desktop'))
[String]$logsPATH = [String]$(mkdir "$DesktopPath\MS-Logs\Calendar-LOGS-$ts")
Start-Transcript "$logsPATH\Calendar-LOGS-$ts.txt" -Verbose

#     $user = "user@domain.com"
$MBXs = Try {get-EXOmailbox -resultsize unlimited } catch { get-mailbox -resultsize unlimited }

$Select = $MBXs | select userprincipalname,Displayname,windows*,Alias,Exchangeguid,externaldirectoryobjectid | Out-Gridview -OutputMode single -Title "Select Mailbox for Calendar Logs"

$User = "$($Select.userprincipalname)"

$startDays = "-30" # [-] days in past
  $endDays = "30"  # [+] days in future
 
$path = $logsPATH

      $st = (get-date).AddDays($startDays).tostring(“MM/dd/yyyy”); 
      $ed = (get-date).AddDays($endDays).tostring(“MM/dd/yyyy”);
$Calendar = Get-CalendarDiagnosticLog $user -StartDate $start -EndDate $end -WA silentlycontinue ;

$Cal = $Calendar | sort-Object NormalizedSubject,LogDate,CleanGlobalObjectId -Unique | Out-Gridview -PassThru -Title "Select Events for Calendar Logs"


   $count = $Cal.count ; $label = "Calendarlogs for [$user]"
  for ($C = 0; $C -lt $Cal.count; $C++) { 
       $S =" [Event Count] ($($C+1)/$count)  [Time]"
       $A = "$label [Subject] $($Cal[$C].NormalizedSubject) [Date] $($Cal[$C].LogDate)" ; 
      $ID = $Cal[$C].CleanGlobalObjectId
       
       Write-Progress -Activity $A -Status $S -PercentComplete (($C/$count)*100) -SecondsRemaining ($count-$C) ;

             $Out = Get-CalendarDiagnosticObjects -Identity $user -MeetingId $ID -StartDate $start -EndDate $end -WA silentlycontinue ;
            $DayN = (get-date $Cal[$C].LogDate).DayOfWeek ; 
         $Datestr = get-date $Cal[$C].LogDate -Format yyyy-MM-dd-HHmm ;
   [string]$FileN =  $Cal[$C].CleanGlobalObjectId + '_' + $DayN + '_' + $Datestr + '-' + $Cal[$C].NormalizedSubject 
[string]$FileName = $FileN.Split([IO.Path]::GetInvalidFileNameChars()) -join '_' ; [int]$Lgt = (250 - $path.Length)
IF ($FileName.length -gt $Lgt) { $File = $FileName.substring(0, $Lgt) + '.Csv' } else { $File = $FileName + '.Csv' } ;
IF(!(Test-Path $path)){ mkdir $path} ; $Out | Export-csv $path\$File -NoTypeInformation -force}

Stop-Transcript
Compress-Archive -Path $logsPATH -DestinationPath "$DesktopPath\MS-Logs\Calendar-LOGS-$ts.zip" # Zip Logs
Invoke-Item $DesktopPath\MS-Logs # open Logs Folder in Filemanager
###### END ZIP Logs ########################