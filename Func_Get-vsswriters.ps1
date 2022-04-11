# List all vss writers from a computer

Function Get-VssWriters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$False)]
        [Boolean]
        $DisplayService = $true
    )
    # Get the vss writer list
    $vssadmin = vssadmin list writers

    # Counter for the for loop. 
    $counter = ($vssadmin | Select-String -Pattern "Writer name").count
    
    # Select the string 
    $writers = $vssadmin | Select-String -Pattern "Writer name","State","Last error"
    # array to make objects
    $DisplayCollection = @()

    # linecounters
    $a = 0
    $b = 1
    $c = 2
    
    for ($i=0; $i -lt $counter ;$i++){
        
        # Name of the writer, this is NOT the name of the service. 
        $nameWriter = $writers[$a] -replace 'Writer name: ',''
        $nameWriter = $nameWriter -replace "\'",""

        # State of the writer
        $nameState = $writers[$b] -replace "\s{3}\w{5}.{6}",""

        # Last error of the writer
        $nameError = $writers[$c] -replace "\s{3}\w{4}.{8}",""

        # corresponding display name of the service. You can add other writers and service displaynames here. 
        $servicename = switch ($nameWriter){
           'ASR Writer' {"Volume Shadow Copy"}
           'BITS Writer' {"Background Intelligent Transfer Service"}
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
           'NTDS' {"NTDS"}
           'OSearch VSS Writer' {"Office SharePoint Server Search"}
           'OSearch14 VSS Writer' {"SharePoint Server Search 14"}
           'Performance Counters Writer' {"Cryptographic Services"}
           'Registry Writer' {"Volume Shadow Copy"}
           'Shadow Copy Optimization Writer' {"Volume Shadow Copy"}
           'SPSearch VSS Writer' {"Windows SharePoint Services Search"}
           'SPSearch4 VSS Writer' {"SharePoint Foundation Search V4"}
           'SqlServerWriter' {"SQL Server VSS Writer"}
           'System Writer' {"Cryptographic Services"}
           'TermServLicensing' {"Remote Desktop Licensing"}
           'WIDW Writer' {"Windows Internal Database"}
           'WINS Jet Writer' {"Windows Internet Name Service (WINS)"}
           'WMI Writer' {"Windows Management Instrumentation"}
           'Task Scheduler Writer' {"Not Supported"}
           'VSS Metadata Store Writer' {"Not Supported"}
           #'' {""}
       }
       
    # Custom object array   
    $DisplayCollection += [PSCustomObject]@{
        Writer = $nameWriter
        State = $nameState
        Error = $nameError
        DisplayServiceName = $servicename
    }
        # Linecounter
        $a = $a + 3
        $b = $b + 3
        $c = $c + 3
        
    }
      
    # Output
   Write-Output $DisplayCollection
}
