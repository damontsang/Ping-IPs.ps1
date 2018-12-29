<#
.SYNOPSIS
    Ping all the hosts read from a .CSV file.
.DESCRIPTION
    The default action is to ping 5 times each of the hosts simultaneously. 
    The ping results will show at the end.For each successful ping, the 
    roundtrip time will be used for the calculation of average roundtrip time.
.PARAMETER .Gridview
    Specify this switch to generate gridview at the end of the execution.
.PARAMETER .ShowTime
    Display individual ping test finish time.
.PARAMETER .Flood
    ICMP packet sent once every second if not used. Otherwise, ICMP send continuously 
.PARAMETER .Detail
    To show statistics of the ping tests at the end of the return table.
    The statistics includes:
    N   : number of Tries
    Min : minimum roundtrip time in milliseconds
    AVG : average roundtrip time in milliseconds
    SD  : standard deviation
    Max : maximum roundtrip time in milliseconds
.PARAMETER .Wait
    Time in millisecond for waiting echo reply from destination. Echo reply packet 
    return in longer then the setting will be considered ping failure. 
    Default set to 1000 milliseconds
.PARAMETER .Tries
    Number of ping for each host.
.PARAMETER .PacketSize
    Buffer size in bytes for an ICMP packet.
.PARAMETER .Version
    Show the version of the script and quite the script.
.PARAMETER .RelNotice
    Show the release notice and quite the script.
.PARAMETER .Verbose
    Specify this switch to show running time of each phases of the execution of 
    the script.
.PARAMETER .FileName
    Specify file name for hosts in a .CSV file in following format. 
    Columns 'System' and 'Name' are names only. Column 'Host' is used for ping.

    System,Name,Host
    External,www.yahoo.com,www.yahoo.com
    External,www.google.com,www.google.com
    Internal,Dev,192.168.1.10
.PARAMETER .ExportFile
    Export ping results to a file name by appending '_Results' to file name assigned by parameter -FileName.

    Ping-IPs.ps1 -FileName hosts.csv -ExportFile

    hosts.csv: host to ping
    hosts_Results.csv: created by the script for ping results

.EXAMPLE
    Ping-IPs.ps1 -FileName hosts.csv  
    
    2018-Dec-29 11:46:29.886 AM: All pings started and waiting for their complete.
    2018-Dec-29 11:46:30.411 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-Dec-29 11:46:30.455 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-Dec-29 11:46:40.005 AM: All ping tests finished

    System   Name           Host           Loss  Average RTT (ms)
    ------   ----           ----           ----  ----------------
    External www.yahoo.com  www.yahoo.com  0/5  52.7 ms         
    External www.google.com www.google.com 0/5  20 ms           
    Internal Dev            192.168.1.10   10/5 0 ms

.EXAMPLE
    Ping-IPs.ps1 -FileName hosts.csv -Detail

    2018-Dec-29 11:46:29.886 AM: All pings started and waiting for their complete.
    2018-Dec-29 11:46:30.411 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-Dec-29 11:46:30.455 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-Dec-29 11:46:40.005 AM: All ping tests finished

    System   Name           Host           Loss  Average RTT (ms)
    ------   ----           ----           ----  ----------------
    External www.yahoo.com  www.yahoo.com  0/5  52.7 ms         
    External www.google.com www.google.com 0/5  20 ms           
    Internal Dev            192.168.1.10   10/5 0 ms         

.EXAMPLE
    Ping-IPs.ps1 -FileName hosts.csv -ShowTime

    2018-Dec-29 11:55:28.973 AM: All pings started and waiting for their complete.
    2018-Dec-29 11:55:29.478 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-Dec-29 11:55:29.482 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-Dec-29 11:55:39.535 AM: All ping tests finished

    Time                        System   Name           Host           Loss  Average RTT (ms)
    ----                        ------   ----           ----           ----  ----------------
    2018-Dec-29 11:55:39.018 AM External www.yahoo.com  www.yahoo.com  0/5  52.5 ms         
    2018-Dec-29 11:55:39.067 AM External www.google.com www.google.com 0/5  20.6 ms         
    2018-Dec-29 11:55:39.153 AM Internal Dev            192.168.1.10   10/5 0 ms      

.EXAMPLE
    ping-IPs.ps1 -FileName hosts.csv -Gridview
.EXAMPLE
    ping-IPs.ps1 -FileName hosts.csv -Verbose

    2018-Dec-29 11:59:25.888 AM: Script Ping-IPs.ps1 started
    2018-Dec-29 11:59:25.898 AM: [Verbose] Create running spool.
    2018-Dec-29 11:59:25.936 AM: [Verbose] Create host list.
    2018-Dec-29 11:59:25.968 AM: [Verbose] Initialize ping tests.
    2018-Dec-29 11:59:26.040 AM: All pings started and waiting for their complete.
    2018-Dec-29 11:59:26.065 AM: [Verbose] All pings started and waiting for their complete. Waiting.
    2018-Dec-29 11:59:26.577 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-Dec-29 11:59:26.601 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    .........
    2018-Dec-29 11:59:31.215 AM: [Verbose] 5.31779s All jobs completed!
    2018-Dec-29 11:59:31.218 AM: All ping tests finished

    System   Name           Host           Loss Average RTT (ms)
    ------   ----           ----           ---- ----------------
    External www.yahoo.com  www.yahoo.com  0/5  53.2 ms         
    External www.google.com www.google.com 0/5  20.8 ms         
    Internal Dev            192.168.1.10   5/5  0 ms 

.EXAMPLE
    ping-IPs.ps1 -FileName hosts.csv -PacketSize 1500

    2018-Dec-29 12:10:46.169 PM: All pings started and waiting for their complete.
    2018-Dec-29 12:10:46.760 PM: @{System=Internal; Name=Dev; Host=192.168.1.10} ping success.
    2018-Dec-29 12:10:51.784 PM: All ping tests finished

    System   Name           Host           Loss Average RTT (ms)
    ------   ----           ----           ---- ----------------
    External www.yahoo.com  www.yahoo.com  5/5  0 ms            
    External www.google.com www.google.com 5/5  0 ms            
    Internal Dev            192.168.1.10   0/5  3.4 ms       

.EXAMPLE
    PS C:\Users\Damon\desktop> .\Ping-IPs.ps1 -ExportFile 
    2018-Dec-29 15:17:10.601 PM: All pings started and waiting for their complete.
    2018-Dec-29 15:17:11.109 PM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-Dec-29 15:17:11.139 PM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-Dec-29 15:17:11.149 PM: @{System=Internal; Name=Dev; Host=192.168.1.10} ping success.
    2018-Dec-29 15:17:16.195 PM: All ping tests finished

    System   Name           Host           Loss Average RTT (ms)
    ------   ----           ----           ---- ----------------
    External www.yahoo.com  www.yahoo.com  0/5  54 ms           
    External www.google.com www.google.com 0/5  20 ms           
    Internal Dev            192.168.1.10   0/5  0 ms            



    PS C:\Users\Damon\desktop> dir *.csv


        Directory: C:\Users\Damon\desktop


    Mode                LastWriteTime         Length Name                                                                                                                                                                                                            
    ----                -------------         ------ ----                                                                                                                                                                                                            
    -a----       29/12/2018   2:12 AM            121 hosts.csv                                                                                                                                                                                                       
    -a----       29/12/2018   3:17 PM            560 hosts_results.csv                                                                                                                                                                                               



    PS C:\Users\Damon\desktop> type .\hosts_results.csv
    2018-Dec-29 15:17:11.130 PM,@{System=External; Name=www.yahoo.com; Host=www.yahoo.com},success
    2018-Dec-29 15:17:11.143 PM,@{System=External; Name=www.google.com; Host=www.google.com},success
    2018-Dec-29 15:17:11.152 PM,@{System=Internal; Name=Dev; Host=192.168.1.10},success

.NOTES
    Author: Damon
    Date:   December 29, 2018    
#>
 
Param (
   [switch]$Verbose,
   [switch]$Gridview,
   [switch]$ShowTime,
   [switch]$Detail,
   [switch]$Version,
   [switch]$RelNotice,
   [switch]$Flood,
   [string]$Wait="1000",
   [string]$Tries="5",
   [uint16]$PacketSize=32,
   [ValidateScript({
      if(-Not ($_ | Test-Path) ){
         throw "File or folder does not exist" 
      }
      if(-Not ($_ | Test-Path -PathType Leaf) ){
         throw "The Path argument must be a file. Folder paths are not allowed."
      }
      return $true
   })]
   [System.IO.FileInfo]$FileName="./hosts.csv",
   [switch]$ExportFile
)

$PingStatusDescription=@{}
$PingStatusDescription[0]="failure"
$PingStatusDescription[1]="success"

if ($Verbose){Write-Host ("{0}: Script {1} started" -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"), $MyInvocation.MyCommand.Name)}
if ($Version -or $RelNotice){Write-Host ("{0} Version 0.1" -f $MyInvocation.MyCommand.Name)}
if ($RelNotice){
   Write-Host ("...")
   Exit
}
if ($Version -or $RelNotice){ Exit }

$StopWatch = [system.diagnostics.stopwatch]::startNew()

$Throttle = 100 #threads
 
if([bool]($Wait -as [int32] -is [int32]) -and [bool]($Tries -as [int32] -is [int32])) {
   [int32]$nTries = [convert]::ToInt32($Tries,10)
   [int32]$nWait = [convert]::ToInt32($Wait,10)
} else {
   Write-Host "Input parameters format error, please try again"
   Exit
}
 
# script block to ping a host and return an object of Host & Conn
$TestScriptBlock = {
   Param (
      # [string]$TargetHost
      $TargetHost,
      $Tries,
      $Wait,
      $Detail,
      $Flood,
      $PingHistory,
      [uint16]$PacketSize = 32
   )
   $ICMPClient = New-Object System.Net.NetworkInformation.Ping
   $PingOptions = New-Object System.Net.NetworkInformation.PingOptions
   $PingOptions.DontFragment = $True
   $nos = 0
   $not = 0
   $tt  = 0
   $ps = [uint16]0
   
   if($Tries -eq 0) {
      $Tries = [int]::MaxValue
   }
   $rrts = 1..$Tries | % {
      $pingStartTime = (get-date)
      $PingResult = ($ICMPClient.Send($TargetHost.Host, $Wait, [System.Byte[]]::CreateInstance([System.Byte],$PacketSize)))
      # ??? $not++
      if($PingResult.Status -eq "Success") { 
         $ps = ($ps -shl 1) -bor 1
         $nos++
      } else {
         $ps = $ps -shl 1
      }
      $PingHistory[$TargetHost] = $ps
      $pingEndTime = (get-date)
      if(!$Flood) {sleep -Milliseconds (1000 - (($pingEndTime - $pingStartTime).TotalMilliseconds))}
      $PingResult.RoundtripTime
   }
   $rrtstats = $rrts | Measure-Object -Average -Maximum -Minimum | select Count, Average, Maximum, Minimum
   $popdev = 0
   foreach ($rrt in $rrts) {
      $popdev +=  [math]::pow(($rrt - $rrtstats.Average), 2)            
   }
   if($rrtstats.Count -gt 1) {
      $sd = [math]::sqrt($popdev / ($rrtstats.Count-1))
   } esle {
      $rrtstats.Maximum = ""
      $rrtstats.Minimum = ""
      $rrtstats.Average = ""
      $rrtstats.Count = ""
      $sd = ""
   }
   $rrtsRS = ""
   if(!$Detail) {
      $rrtsRS = [string]::Format("{0} ms", $rrtstats.Average)
   } else {
      $rrtsRS = [string]::Format("{0} / {1:N1} / {2:N1} / {3:N1} / {4:N1}", $rrtstats.Count, $rrtstats.Minimum, $rrtstats.Average, $sd, $rrtstats.Maximum)
   }
   Return New-Object PSObject -Property @{
      Time = Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss.fff tt"
      System = $TargetHost.System
      Name = $TargetHost.Name
      Host = $TargetHost.Host
      Loss = [string]($Tries-$nos)+'/'+$Tries
      RoundtripTimeStats = $rrtsRS
   }
}
 
if ($Verbose){Write-Host ("{0}: [Verbose] Create running spool." -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"))}
$PreviousPingHistory = @{}
$PingHistory = [hashtable]::Synchronized(@{})

# create running spool having $Throttle
$SessionState = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $SessionState, $Host)
$RunspacePool.Open()
# $RunspacePool.SessionStateProxy.SetVariable('Hash',$hash)
$Jobs = @()
 
if ($Verbose){Write-Host ("{0}: [Verbose] Create host list." -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"))}

$Seq = 0
# host list
$c = @()
 
# host/ip to ping
$c = import-csv $FileName

if ($Verbose){Write-Host ("{0}: [Verbose] Initialize ping tests." -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"))}
$c | % {
   $PingHistory[$_] = $null
   $Job = [powershell]::Create().AddScript($TestScriptBlock).AddArgument($_).AddArgument($nTries).AddArgument($nWait).AddArgument($Detail).AddArgument($Flood).AddArgument($PingHistory).AddArgument($PacketSize)
   $Job.RunspacePool = $RunspacePool
   $Jobs += New-Object PSObject -Property @{
      RunNum = $_
      Pipe = $Job
      Result = $Job.BeginInvoke()
   }
}

Write-Host ("{0}: All pings started and waiting for their complete." -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"))

if ($Verbose){Write-Host -NoNewline ("{0}: [Verbose] All pings started and waiting for their complete. Waiting" -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"))}

# $PreviousPingHistory = [hashtable]
# wait until all ping jobs completed.
$p = '(.+)(\.csv)'
$r = '$1_results$2'
if(test-path ($FileName -replace $p, $r)){ remove-item -Force -Confirm:$false ($FileName -replace $p, $r)}
Do {
   if ($Verbose){Write-Host "." -NoNewline}
   Start-Sleep -Milliseconds 500
   foreach($k in $($PingHistory.Keys)) {
      if(($PreviousPingHistory[$k] -band [uint16]1) -ne ($PingHistory[$k] -band [uint16]1)) {
         Write-Host ("{0}: {1} ping {2}." -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"), $k, $PingStatusDescription[$PingHistory[$k] -band [uint16]1]) # , [System.Convert]::ToString($PingHistory[$k] -band [uint16]65535,8).padleft(5,'0')) 
         if($ExportFile){ 
            ("{0},{1},{2}" -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"), $k, $PingStatusDescription[$PingHistory[$k] -band [uint16]1]) | Out-File -Append -FilePath $($FileName -replace $p, $r)
            # Write-Host ("{0}: {1} {2}." -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"), $PingStatusDescription[$PingHistory[$k] -band [uint16]1], $($FileName -replace $p, $r))
         }
         $PreviousPingHistory[$k] = $PingHistory[$k]
      }
   }
} While ( $Jobs.Result.IsCompleted -contains $false)
if ($Verbose){Write-Host}
if ($Verbose){Write-Host ("{0}: [Verbose] {1:N5}s All jobs completed!" -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"), $StopWatch.Elapsed.TotalSeconds)}

$StopWatch.Stop()

Write-Host ("{0}: All ping tests finished" -f (Get-Date -Format "yyyy'-'MMM'-'dd HH':'mm':'ss'.'fff tt"))

# print out the result
$Results = @()
ForEach ($Job in $Jobs) {
   $Results += $Job.Pipe.EndInvoke($Job.Result)
}

if($Detail) {
   $StatsLabel = 'N / Min / AVG / SD / Max'
} else {
   $StatsLabel = 'Average RTT (ms)'
}

if ($Gridview){
   $Results | Select-Object @{label='System';expression={$_.system}}, Name, Host, Loss, @{label=$StatsLabel; expression={$_.RoundtripTimeStats}} | Out-GridView -Title "Ping Results"
} else {
   if($ShowTime) {
      $Results | Format-Table Time, System, Name, Host, Loss, @{label=$StatsLabel; expression={$_.RoundtripTimeStats}} -AutoSize
   } else {
      $Results | Format-Table System, Name, Host, Loss, @{label=$StatsLabel; expression={$_.RoundtripTimeStats}} -AutoSize
   }
}
