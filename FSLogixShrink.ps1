#Requires -Version 5
#Requires -RunAsAdministrator
#Requires -Module Hyper-V
 

<#
Written by Florian Benz
I stole some code from www.thesurlyadmin.com (set-alternatingrows function) and www.citrixirc.com
This script will compact FSLogix VHD/VHDX profiles in the profile share.  It would also work for
any directory containing VHD/VHDX files.
Test before using!!
Search for "#####" to find the sections you need to edit for your environment
#>

 

# Declaration of Variables #
$DriveLetter = "" ##### Diskletter used to mount the vhd(x) files
$smtpserver = "" ##### SMTP Server
$to = "" ##### email report to - "email1","email2" for multiple
$from = "" ##### email from
$rootfolder = "" ##### root path to vhd(x) files


# DO NOT CHANGE #
$t = "0" # Variable declared for further use
$i = "0" # Variable declared for further use
$dismount = "0" # Variable declared for further use
[System.Collections.ArrayList]$info = @() # Array to be processed for report generation
# DO NOT CHANGE #


Function Set-AlternatingRows 
    {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)]
        [string]$Line,
      
        [Parameter(Mandatory)]
        [string]$CSSEvenClass,
      
        [Parameter(Mandatory)]
        [string]$CSSOddClass
        )
    Begin 
        {
        $ClassName = $CSSEvenClass
        }
    Process 
        {
        If ($Line.Contains("<tr><td>"))
            {   
            $Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
            If ($ClassName -eq $CSSEvenClass)
                {
                $ClassName = $CSSOddClass
                }
            Else
                {
                $ClassName = $CSSEvenClass
                }
            }
        Return $Line
        }
    }

 

Function checkFileStatus($filePath)
    {
        $fileInfo = New-Object System.IO.FileInfo $filePath
        try
        {
            $fileStream = $fileInfo.Open( [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read )
            $filestream.Close()
            return $false
        }
        catch
        {
        return $true
        }
    }
     

Function vhdmount($vhd) 
    {
    try 
        {
        Mount-VHD -Path $vhd -Passthru -ErrorAction Stop | Get-Disk | Get-Partition |Set-Partition -NewDriveLetter $DriveLetter
        return "0"
        } 
    catch
        {
        return "1"
        }
    }


Function vhdoptimize($vhd) 
    {
    $o = [math]::Round((Get-Item $vhd | Select-Object -expand length)/1mb,2)
    try 
        {
        Optimize-Volume -DriveLetter $DriveLetter -ReTrim -Verbose -ErrorAction Stop
        $SizeMin = (Get-PartitionSupportedSize -DriveLetter $DriveLetter).SizeMin
        $SizeMinRounded = [math]::Round($SizeMin/1mb,2)
        $SizeActual = [math]::Round(((Get-Partition -DriveLetter Z).Size)/1mb,2)
        If ($SizeMinRounded -lt $SizeActual)
            {
            Resize-Partition -DriveLetter $DriveLetter -Size $SizeMin -ErrorAction Stop
            Resize-VHD -Path $vhd -ToMinimumSize -ErrorAction Stop
            }
        Get-Volume -Drive $DriveLetter | Get-Partition | Remove-PartitionAccessPath -AccessPath "$DriveLetter`:\"
        Dismount-VHD $vhd
        Optimize-VHD -Path $VHD -Mode Full
        $r = 0
        }
    catch
        {
        $r = 1
        }
    $n = [math]::Round((Get-Item $vhd | Select-Object -expand length)/1mb,2)
    $dif = [math]::Round(($o-$n),2)
    $i | Select-Object @{n='VHD';e={Split-Path $vhd -Leaf}},@{n='Before_MB';e={$o}},@{n='After_MB';e={$n}},@{n='Reduction_MB';e={$dif}},@{n='Success';e={if ($r -eq "0"){$true} else {$false}}},@{n='VHD_Fullname';e={$vhd}}
    }


$vhds = (Get-ChildItem $rootfolder -recurse -Include *.vhd,*.vhdx).fullname

foreach ($vhd in $vhds) 
    {
    $locked = checkFileStatus -filePath $vhd
    if ($locked -eq $true) 
        {
        "$vhd in use, skipping."
        $info.add(($t | Select-Object @{n='VHD';e={Split-Path $vhd -Leaf}},@{n='Before_MB';e={0}},@{n='After_MB';e={0}},@{n='Reduction_MB';e={0}},@{n='Success';e={"Locked"}},@{n='VHD_Fullname';e={$vhd}})) | Out-Null
        continue
        }
        
    $mount = vhdmount -v $vhd
    if ($mount -eq "1")
        {
        $e = "Mounting $vhd failed "+(get-date).ToString()
        Write-Host $e
        Send-MailMessage -SmtpServer $smtpserver -From $from -To $to -Subject "FSLogix VHD(X) ERROR" -Body "$e" -Priority High -BodyAsHtml
        break
        }

    $info.add((vhdoptimize -v $vhd)) | Out-Null

    $Attached = Get-VHD $vhd | Select-Object Attached
        If ($Attached.Attached -eq "True")
            {
            $dismount = "1"
            }
        Else
            {
            $dismount = "0"
            }
    If ($dismount -eq "1")
        {
        $e = "Failed to dismount $vhd "+(get-date).ToString()
        Send-MailMessage -SmtpServer $smtpserver -From $from -To $to -Subject "FSLogix VHD(X) ERROR" -Body "$e" -Priority High -BodyAsHtml
        break
        }
    }

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;width: 95%}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
"@
$date = Get-Date
##### uncomment the next 2 lines if you would like to save a .htm report (also 2 more at the end)
#$out = Join-Path ([environment]::GetFolderPath("mydocuments")) ("FSLogix_Reports\VHD_Reduction_Report_$timestamp.htm")
#if (!(test-path (Split-Path $out -Parent))) {New-Item -Path (Split-Path $out -Parent) -ItemType Directory -Force | Out-Null}
$before = ($info.before_mb | Measure-Object -Sum).Sum
$after = ($info.after_mb | Measure-Object -Sum).Sum
$reductionmb = ($info.reduction_mb | Measure-Object -Sum).sum
$message = $info | Sort-Object After_MB -Descending | ConvertTo-Html -Head $header -Title "FSLogix VHD(X) Reduction Report" -PreContent "<center><h2>FSLogix VHD(X) Reduction Report</h2>" -PostContent "</center><br><h3>Pre-optimization Total MB: $before<br>Post-optimization Total MB: $after<br>Total Reduction MB: $reductionmb</h3>"| Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd
##### comment the next line if you do not wish the report to be emailed
Send-MailMessage -SmtpServer $smtpserver -From $from -To $to -Subject ("FSLogix VHD(X) Reduction Report "+($date).ToString()) -Body "$message" -BodyAsHtml
##### uncomment the next 2 lines to save the report to your My Documents\FSLogix_Reports directory, and open it in your default browser
#$message | Out-File $out
#Invoke-Item $out