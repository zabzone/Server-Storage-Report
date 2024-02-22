<#
.SYNOPSIS
  A script that is meant to pull data drive D: disk space to create weekly reports on usage. This is done so that the endpoints team has glance visibility every week at the changing currents without needing access to multiple Solarwinds instances throughout the OPI domains.
.DESCRIPTION
  Pull data > send report via email
.PARAMETER <Parameter_Name>
    At the moment, no inputs required - everything is hard coded.
.INPUTS
  None
.OUTPUTS
  HTML file put into the scripting servers local C:\temp > outputs into email
.NOTES
  Version:        1.0
  Author:         Zab Rivera
  Creation Date:  01.17.2024
  Purpose/Change: The beginning
.EXAMPLE
  None
#>

# Module Setups
Install-Module PSWriteHTML -Force
Import-Module PSWriteHTML

# Script Run Setup
Set-ExecutionPolicy Bypass

$html_output = "C:\temp\Ivanti_Pref_Srv_Data_Drive_Space_Report_$(Get-Date -Format yyyyMMdd).html" 

# Hash of all Preferred Servers
$servers = @('SRVAKR2','SRVAMH2','SRVANT2','SRVATH2','SRVATL2','SRVAUC2','SRVAUD2','SRVAUG2','SRVAUR2','SRVBAK2','SRVBAL2','SRVBAP2','SRVBED2','SRVBKP2','SRVBRI2','SRVBRM2','SRVBTR2','SRVCBL2','SRVCGA2','SRVCGF2','SRVCHA2','SRVCHE2','SRVCHI2','SRVCHK2','SRVCHS2','SRVCHV2','SRVCLB2','SRVCLC2','SRVCLE2','SRVCLR2','SRVCLT2','SRVCLV2','SRVCNH2','SRVCOL2','SRVCON2','SRVCSB2','SRVCTN2','SRVCVS2','SRVDAL2','SRVDAY2','SRVDDN2','SRVDEC2','SRVDEN2','SRVDMO2','SRVDNT2','SRVDUR2','SRVERI2','SRVESC2','SRVEVN2','SRVFIF2','SRVFLO2','SRVFLP2','SRVFOR2','SRVFRS2','SRVFTH2','SRVFTJ2','SRVFTW2','SRVGAR2','SRVGGV2','SRVGLA2','SRVGPR2','SRVGRE2','SRVGRT2','SRVHAZ2','SRVHFM2','SRVHLY2','SRVHOU2','SRVHUD2','SRVIDO2','SRVIDP2','SRVIND2','SRVINL2','SRVJAX2','SRVJOP2','SRVKCK2','SRVKEN2','SRVKNO2','SRVKZO2','SRVLAB2','SRVLAN2','SRVLEW2','SRVLEX2','SRVLOU2','SRVLUB2','SRVLVC2','SRVLVD2','SRVLYN2','SRVMAD2','SRVMAP2','SRVMEA2','SRVMEP2','SRVMEW2','SRVMID2','SRVMIG2','SRVMIL2','SRVMOL2','SRVMON2','SRVMPW2','SRVMSA2','SRVMTG2','SRVMWO2','SRVNEW2','SRVNFM2','SRVNLD2','SRVNLK2','SRVNLR2','SRVNOR2','SRVNRV2','SRVOAK2','SRVOKC2','SRVOMA2','SRVORB2','SRVORL2','SRVORS2','SRVOSS2','SRVPAS2','SRVPEM2','SRVPEN2','SRVPET2','srvpln2','SRVPON2','SRVPOR2','SRVPRT2','SRVRAL2','SRVRED2','SRVREY2','SRVRIB2','SRVRIC2','SRVRIV2','SRVRKM2','SRVROC2','SRVRVC2','SRVSAN2','SRVSAW2','SRVSAZ2','SRVSBE2','SRVSDG2','SRVSDS2','SRVSFS2','srvsio2','SRVSLC2','SRVSOU2','SRVSPA2','SRVSPB2','SRVSPK2','SRVSPO2','SRVSPR2','SRVSTC2','SRVSTL2','SRVSTP2','SRVSUM2','SRVSZA2','SRVTER2','SRVTPA2','SRVTUL2','SRVTYL2','SRVVCR2','SRVVIR2','SRVVIS2','SRVVLD2','SRVVNC2','SRVVRB2','SRVWAC2','SRVWCL2','SRVWCO2','SRVWHL2','SRVWIC2','srvwik2','SRVWML2','SRVWPB2','SRVWSL2')

# Define a script block to collect data
$driveDataScriptBlock = {

    param ($server)

    # Check if server online > continue > else > return object to report that states node is offline
    if(Test-Connection -ComputerName $server -Quiet -Count 1){
            # If WinRM services are enabled > try to grab data and pass object to function caller > or catch it and identify if its a winRM issue or a connectivity issue and return object stating which on that node
            # This entire scriptblock has 5 potential returns
            if([bool](Test-WSMan -ComputerName $server -ErrorAction SilentlyContinue)){
                try {
                    # Pull drive data and create your own percentage and totals using built-in math functions
                    $data = Invoke-Command -ComputerName $server {Get-PSDrive D | Select-Object Name, Used, @{Name='UsedInGB';Expression={[math]::Round($_.Used / 1GB)}}, Free, @{Name='FreeInGB';Expression={[math]::Round($_.Free / 1GB, 2)}}, @{Name='PercentUsed';Expression={[math]::Round($_.Used / 1GB, 2)}}}

                    # create percentage metric based on pulled data
                    $PercentUsed = ( ($data).UsedInGB / ( ($data).FreeInGB + ($data).UsedInGB ) )
                    $RoundedPercentUsed = [math]::Round($PercentUsed * 100, 0)

                    $result = [PSCustomObject]@{
                        HostName = $server
                        DriveLetter = ($data).Name
                        UsedInGB = ($data).usedInGB
                        PercentUsed = "$RoundedPercentUsed%"
                        FreeInGB = ($data).FreeInGB
                    }
                    return $result
                }
                catch {
                    if ($_.Exception.Message -like '*The client cannot connect to the destination*') {
                        Write-Host "Unable to connect to $server. Zabbie connection issue."

                        $result = [PSCustomObject]@{
                            HostName = $server
                            DriveLetter = "Cannot Detect WinRM Services"
                            UsedInGB = "Cannot Detect WinRM Services"
                            PercentUsed = "Cannot Detect WinRM Services"
                            FreeInGB = "Cannot Detect WinRM Services"
                            }
                            return $result
                    }
                    else {
                        Write-Host "Error: $($_.Exception.Message)"

                        $result = [PSCustomObject]@{
                            HostName = $server
                            DriveLetter = "Cannot establish a network connection to server"
                            UsedInGB = "Cannot establish a network connection to server"
                            PercentUsed = "Cannot establish a network connection to server"
                            FreeInGB = "Cannot establish a network connection to server"
                        }
                        return $result
                    }
                }
            }
        else {
            # If WinRM services are not enabled, pass object saying not enabled - return this value - exit method
            $result = [PSCustomObject]@{
            HostName = $server
            DriveLetter = "Cannot Detect WinRM Services"
            UsedInGB = "Cannot Detect WinRM Services"
            PercentUsed = "Cannot Detect WinRM Services"
            FreeInGB = "Cannot Detect WinRM Services"
            }

            return $result
        }
    }
    else {
        $result = [PSCustomObject]@{
            HostName = $server
            DriveLetter = "Cannot establish a network connection to server"
            UsedInGB = "Cannot establish a network connection to server"
            PercentUsed = "Cannot establish a network connection to server"
            FreeInGB = "Cannot establish a network connection to server"
        }
        return $result
    }
}

# Invoke the script block on multiple servers in parallel
$jobs = $servers | ForEach-Object {
    $server = $_
    Start-Job -ScriptBlock $driveDataScriptBlock -ArgumentList $server
}

# Wait for all jobs to complete
$jobs | Wait-Job

# Retrieve and format the results from the completed jobs
$Output = $jobs | Receive-Job | Where-Object {$_ -ne $null} | Select-Object Hostname, DriveLetter, UsedInGB, PercentUsed, FreeInGB | Sort-Object PercentUsed -Descending

# Cleanup jobs
$jobs | Remove-Job

# Grab all the jobs and output them as an HTML file with basic CSS styling and html page configs
$Output | Out-HtmlView  {
    New-TableCondition -Name 'PercentUsed' -Operator ge -Value 80 -BackgroundColor Yellow -Color Black -Inline -ComparisonType number
    New-TableCondition -Name 'PercentUsed' -Operator lt -Value 80 -BackgroundColor Green -Color White -Inline -ComparisonType number
    New-TableCondition -Name 'PercentUsed' -Operator ge -Value 95 -BackgroundColor Red -Color White -Inline -ComparisonType number
    New-TableHeader -Title "Ivanti Preferred Servers D: Data Drive Report ($(Get-Date))" -Alignment center -BackGroundColor BuddhaGold -Color White -FontWeight bold
    New-TableHeader -Names 'DriveLetter','UsedInGB','PercentUsed','FreeInGB' -Title 'Disk Information' -Color White -Alignment center -BackGroundColor Gray
} -HideFooter -PagingLength 250 -FilePath $html_output

$Username = "anonymous"
$Password = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$SmtpServer = "exchange.subdomain.domain.net"
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $Password
$Recipient = "zab@thatsme.com"
$Sendor = "iwata@nintendo.com"
$Body = Get-Content -Path $html_output
$Attachment = $html_output
$Subject = "Ivanti_Pref_Srv_Data_Drive_Space_Report $(Get-Date)"

Send-MailMessage -To $Recipient -From $Sendor -Subject $Subject -Attachments $Attachment -Body "$Body" -BodyAsHtml -SmtpServer $SmtpServer -Credential $Credentials