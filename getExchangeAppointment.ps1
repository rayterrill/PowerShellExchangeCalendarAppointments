# NOTE : find the DLL path for your Microsoft.Exchange.WebServices.dll. Mine was at C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll

function Get-OutlookInfo($mailboxName, $startDate, $endDate) {
   $dllpath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll"
   [void][Reflection.Assembly]::LoadFile($dllpath)
   $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

   $windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
   $sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
   $user = [ADSI]$sidbind

   $service.AutodiscoverUrl($user.mail.ToString())

   $folder = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)
   $calendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service,$folder)
   $calendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,2000)

   $calendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
   $calendarResult = $calendarFolder.FindAppointments($calendarview)

   $status = ""

   foreach ($appointment in $calendarResult.Items){
      if ($appointment.Subject.ToString() -like '*WFH*' -OR $appointment.Subject.ToString() -like '*Work from home*') {
         $status = "working from home"
      } elseif ($appointment.Subject.ToString() -like '*vacation*') {
         $status = "on vacation"
      }
   }

   return $status
}

#provide a list of employees to check out their schedules
$employees = ('Peter.Gibbons@Initech.com','Michael.Bolton@Initech.com','Samir.Nagheenanajar@Initech.com')

$startDate = [datetime]::Today
$endDate = ([datetime]::Today).AddDays(1)

foreach ($e in $employees) {
   $status = Get-OutlookInfo -mailbox $e -startDate $startDate -endDate $endDate
   if ($status -eq "") {
      Write-Host "$e is here today."
   } else {
      Write-Host "$e is $($status)"
   }

   #do something interesting with their status
   #maybe update a database that feeds a "where are my coworkers?" page
   Push-Location
   $updateStatus = Invoke-SQLCmd -ServerInstance MYSQLSERVER -Database MYDATABASE -Query "update dbo.teamstatus set status = '$($status)' where email = '$($e)'"
   Pop-Location
}