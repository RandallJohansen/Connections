

###########################################################
#                   SLEEP FUNCTION
###########################################################

function SleepFor([single]$loopdelay) {

    # number of seconds to pause on each picture
    # $loopdelay = 10

    # milliseconds to respond to input
    $responsetime = 200
    $loopiterations = 1000 * $loopdelay / $responsetime

    # pause between loops in case there's nothing new to prevent racing
    for($i = 0; $i -lt $loopiterations; $i++)
        {

        #Exit the loop
        if($script:CancelLoop -eq $true) { break }

        Sleep -m $responsetime

#        [System.Windows.Forms.Application]::DoEvents()
    }
}


###########################################################
#                 Main{} INLINE CODE
###########################################################


# alternate one liners:
#
# Get-NetTCPConnection | ? State -eq Established | ? RemoteAddress -notlike 127* #| % { $_; Resolve-DnsName $_.RemoteAddress -type PTR -ErrorAction SilentlyContinue }
#
# filter timestamp {"$(Get-Date -Format G): $_"};netstat -abno 1 | Select-String -Context 0,1 -Pattern LISTENING|timestamp
#



# cast and set default
$script:CancelLoop = $False

while ($script:CancelLoop -eq $false) {

    Clear-Host

    get-nettcpconnection | 
        ? State -eq Established | 
        ? RemoteAddress -notlike 127* | 
        # pulling all too many columns here, but will filter them out during formating at the end
        select local*,remote*,state,AppliedSetting,
        @{Name="DNS";Expression={  (Resolve-DnsName $_.RemoteAddress -type PTR -ErrorAction SilentlyContinue) | % {  $_.NameHost }  }},
        @{Name="Process";Expression={(Get-Process -Id $_.OwningProcess).ProcessName}} | 
        Sort-Object -Property Process, RemoteAddress | 
        Format-Table -Property RemoteAddress,RemotePort,Process,DNS -AutoSize

# object returned by this query
#Resolve-DnsName 45.60.11.212 -type PTR | Get-Member 
#   TypeName: Microsoft.DnsClient.Commands.DnsRecord_PTR
#
#Name         MemberType    Definition                                                     
#----         ----------    ----------                                                     
#QueryType    AliasProperty QueryType = Type                                               
#Server       AliasProperty Server = NameHost                                              
#Equals       Method        bool Equals(System.Object obj)                                 
#GetHashCode  Method        int GetHashCode()                                              
#GetType      Method        type GetType()                                                 
#ToString     Method        string ToString()                                              
#CharacterSet Property      Microsoft.DnsClient.Commands.DNSCharset CharacterSet {get;set;}
#DataLength   Property      uint16 DataLength {get;set;}                                   
#Name         Property      string Name {get;set;}                                         
#NameHost     Property      string NameHost {get;set;}                                     
#Section      Property      Microsoft.DnsClient.Commands.DNSSection Section {get;set;}     
#TTL          Property      uint32 TTL {get;set;}                                          
#Type         Property      Microsoft.DnsClient.Commands.RecordType Type {get;set;}        



    SleepFor 10
    # either -or- for debug
    #$script:CancelLoop = $true

}
