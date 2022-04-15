# List all vss writers from a computer

Function Get-VssWriters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$False)]
        [switch]
        $DisplayWriters,

        [Parameter(Mandatory=$False)]
        [switch]
        $FilterUnreliable,

        [Parameter(Mandatory=$False)]
        [switch]
        $Restart,

        [Parameter(Mandatory=$False)]
        [switch]
        $Log
    )
    # Get the vss writer list
    $vssadmin = vssadmin list writers

    # Counter for the for loop. 
    $counter = ($vssadmin | Select-String -Pattern "Writer name").count
    
    # Select the strings for the array
    $writers = $vssadmin | Select-String -Pattern "Writer name","State","Last error"

    # array to make objects
    $DisplayCollection = @()

    # linecounters
    $a = 0
    $b = 1
    $c = 2
    
    for ($i=0; $i -lt $counter ;$i++){
        
        # Name of the writer, this is NOT the name of the service. Formatting the input. 
        $nameWriter = $writers[$a] -replace 'Writer name: ',''
        $nameWriter = $nameWriter -replace "\'",""

        # State of the writer. Formatting the input. 
        $nameState = $writers[$b] -replace "\s{3}\w{5}.{6}",""

        # Last error of the writer.Formatting the input. 
        $nameError = $writers[$c] -replace "\s{3}\w{4}.{8}",""

        # corresponding display name of the service. You can add other writers and service displaynames here. 
        $DisplayServicename = switch ($nameWriter){
           'ASR Writer' {"Volume Shadow Copy"}
           'BITS Writer' {"Background Intelligent Transfer Service"}
           'Certificate Authority' {"Active Directory Certificate Services"}
           'COM+ REGDB Writer' {"Volume Shadow Copy"}
           'DFS Replication service writer' {"DFS Replication"}
           'DHCP Jet Writer' {"DHCP Server"}
           'FRS Writer' {"File Replication"}
           'FSRM writer' {"File Server Resource Manager"}
           'IIS Config Writer' {"Application Host Helper Service"}
           'IIS Metabase Writer' {"IIS Admin Service"}
           'Microsoft Exchange Replica Writer' {"Microsoft Exchange Replication Service"}
           'Microsoft Exchange Writer' {"Microsoft Exchange Information Store"}
           'Microsoft Hyper-V VSS Writer' {"Hyper-V Virtual Machine Management"}
           'MSMQ Writer (MSMQ)' {"Message Queuing"}
           'MSSearch Service Writer' {"Windows Search"}
           'NPS VSS Writer' {"COM+ Event System"}
           'NTDS' {"NTDS"}
           'OSearch VSS Writer' {"Office SharePoint Server Search"}
           'OSearch14 VSS Writer' {"SharePoint Server Search 14"}
           'OSearch15 VSS Writer' {"SharePoint Server Search 15"}
           'Performance Counters Writer' {"Cryptographic Services"}
           'Registry Writer' {"Volume Shadow Copy"}
           'Shadow Copy Optimization Writer' {"Volume Shadow Copy"}
           'SharePoint Services Writer' {"Windows SharePoint Services VSS Writer"}
           'SMS Writer' {"SMS_SITE_VSS_WRITER"}
           'SPSearch VSS Writer' {"Windows SharePoint Services Search"}
           'SPSearch4 VSS Writer' {"SharePoint Foundation Search V4"}
           'SqlServerWriter' {"SQL Server VSS Writer"}
           'System Writer' {"Cryptographic Services"}
           'TermServLicensing' {"Remote Desktop Licensing"}
           'WDS VSS Writer' {"Windows Deployment Services Server"}
           'WIDWWriter' {"Windows Internal Database"}
           'WINS Jet Writer' {"Windows Internet Name Service (WINS)"}
           'WMI Writer' {"Windows Management Instrumentation"}
           'Task Scheduler Writer' {"Not Supported, writer is obsolete."}
           'VSS Metadata Store Writer' {"Not Supported, writer is obsolete."}
           #'' {""}
       }
       # Servicename so it can Pipe into Restart-Service
       $Servicename = switch ($nameWriter){
        'ASR Writer' {"VSS"}
        'BITS Writer' {"BITS"}
        'Certificate Authority' {"CertSvc"}
        'COM+ REGDB Writer' {"VSS"}
        'DFS Replication service writer' {"DFSR"}
        'DHCP Jet Writer' {"DHCPServer"}
        'FRS Writer' {"NtFrs"}
        'FSRM writer' {"srmsvc"}
        'IIS Config Writer' {"AppHostSvc"}
        'IIS Metabase Writer' {"IISADMIN"}
        'Microsoft Exchange Replica Writer' {"MSExchangeRepl"}
        'Microsoft Exchange Writer' {"MSExchangeIS"}
        'Microsoft Hyper-V VSS Writer' {"vmms"}
        'MSMQ Writer (MSMQ)' {"MSMQ"}
        'MSSearch Service Writer' {"WSearch"}
        'NPS VSS Writer' {"EventSystem"}
        'NTDS' {"NTDS"}
        'OSearch VSS Writer' {"OSearch"}
        'OSearch14 VSS Writer' {"OSearch14"}
        'OSearch15 VSS Writer' {"OSearch15"}
        'Performance Counters Writer' {"Cryptographic Services"}
        'Registry Writer' {"VSS"}
        'Shadow Copy Optimization Writer' {"VSS"}
        'SharePoint Services Writer' {"SPWriter"}
        'SMS Writer' {"SMS_SITE_VSS_WRITER"}
        'SPSearch VSS Writer' {"SPSearch"}
        'SPSearch4 VSS Writer' {"SPSearch4"}
        'SqlServerWriter' {"SQLWriter"}
        'System Writer' {"CryptSvc"}
        'TermServLicensing' {"TermServLicensing"}
        'WDS VSS Writer' {"WDSServer"}
        'WIDWWriter' {"WIDWriter"}
        'WINS Jet Writer' {"WINS"}
        'WMI Writer' {"Winmgmt"}
        'Task Scheduler Writer' {"Not Supported, writer is obsolete."}
        'VSS Metadata Store Writer' {"Not Supported, writer is obsolete."}
        #'' {""}
    }
       
    # Custom object array   
    $DisplayCollection += [PSCustomObject]@{
        Writer = $nameWriter
        State = $nameState
        Error = $nameError
        DisplayServiceName = $DisplayServicename
        ServiceName = $Servicename
    }
        # Linecounter
        $a = $a + 3
        $b = $b + 3
        $c = $c + 3
        
    }
      
    # Display all vss writers
    if ($DisplayWriters.IsPresent){Write-Output $DisplayCollection}

    # Display vss writers that are not working properly. 
    if ($FilterUnreliable.IsPresent){
        $UnreliableWriters = $DisplayCollection | Where-Object {($_.Error -ne "No error") -or ($_.state -ne "Stable")}
        if ($UnreliableWriters) {Write-Output $UnreliableWriters}
        Else {Write-Host "No Errors detected. All vss writers are operational."}
               
    }
    # restart services of Vss Writers that aren't working properly. 
    if ($Restart.IsPresent){
        $RestartWriters = $DisplayCollection | Where-Object {($_.Error -ne "No error") -or ($_.state -ne "Stable") -and ($_.DisplayServiceName -notlike "*obsolete*")}
        $RestartWriters | Restart-Service -Verbose

        # Compare Previous vss writer State with current. 
        $CompareWriters = $DisplayCollection | Where-Object {($_.Error -ne "No error") -or ($_.state -ne "Stable") -and ($_.DisplayServiceName -notlike "*obsolete*")}

        # Write an event if the service restart didn't solve the issue.  
        if ($CompareWriters){
            try {New-EventLog –LogName Application –Source “Get-Vsswriters -Restart”}
            catch {Write-Warning -Message "Source already exists."}
            Write-EventLog -LogName "Application" -Source "Get-Vsswriters -Restart" -EventID 3201 -EntryType Warning -Message "One or more vss writers do not work properly. Restarting the service(s) did not solve the problem. Advice: Reboot the system." 
        }


    }
    # Show latest logs of this script
    if ($log.IsPresent) {
        try {
            Get-EventLog -LogName Application -Source “Get-Vsswriters -Restart” | Sort-Object -Property TimeGenerated
        }
        catch {Write-Host "No log has been found"}
    }
   
}
