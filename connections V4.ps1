
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

        [System.Windows.Forms.Application]::DoEvents()
    }
}



###########################################################
#              EXPIREFILES SCRIPTBLOCK
###########################################################

$ShowConnections = { param( $PutArgHere, $HereToo )




  #  Set-Location -Path $DirToExpire



        ###############
        #             #
        #    GUI      #
        #             #
        ###############

        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Application]::EnableVisualStyles();
        $Form = new-object System.Windows.Forms.Form
     #   $Form.Size = New-Object System.Drawing.Size(640,480)         # <- handled below with Height / Width

        $TextBox = New-Object System.Windows.Forms.RichTextBox
        $Form.controls.add($TextBox)

        $TextBox.Size = New-Object System.Drawing.Size($Form.ClientRectangle.Width,$Form.ClientRectangle.Height)
        $TextBox.Anchor = "top", "left", "Right", "Bottom"
        $TextBox.BringToFront()
        $TextBox.Multiline = $true
        $TextBox.WordWrap = $False
     #   $TextBox.ReadOnly = $True
        $TextBox.ScrollBars = "Vertical"
        $TextBox.font = New-Object System.Drawing.Font("Courier New",12,0,3,0)   # *** monospaced font ***
     #   $TextBox.ForeColor = [Drawing.Color]::Blue

        # find the font used and make a new version using bold
        $PlainFont =  $TextBox.Font
        $BoldFont = New-Object Drawing.Font($PlainFont.FontFamily, $PlainFont.Size, [Drawing.FontStyle]::Bold) 

        $script:CancelLoop = $false    # cast and set default
        $Form.Add_Closing({$script:CancelLoop = $true})

        $Form.Text = "Connections"
        $Form.StartPosition = 'CenterScreen'
        $Form.Width = 1400
        $Form.Height =  700
        $Form.Add_Shown( { $Form.Activate()} )
        $Form.Show();


        



        $TextBox.Text = -join("Setting up, hold please..." , "`r`n") 
            
        [System.Windows.Forms.Application]::DoEvents()   # show the message

 
 
 
        ###############
        #             #
        #   Vars      #
        #             #
        ###############

 
 
        # create an empty "list" collection
        $table = [System.Collections.Generic.List[PSObject]]::new()

        $ActiveLocalPorts = @{}
        $WatcingLocalPorts = @{}
        $OriginalOwningProcess = @{}


        $RejectedDestinations = @( '127.0.0.1' , '0.0.0.0' , '::' )

        $RejectedApps = @( 'some app' , 'other app' )
        

        
        $OtherCountries = @( 'China' )




        ###############
        #             #
        # scriptblock #
        #   Main{}    #
        #             #
        ###############



        while ($script:CancelLoop -eq $false) {


 
            ###############
            #             #
            #   Gather    #
            #             #
            ###############


                Clear-Host # fresh for showing current connections in text window

 
                foreach ( $CurrConn in Get-NetTCPConnection | ? { $RejectedDestinations -NotContains $_.RemoteAddress } ) {





                        if($script:CancelLoop -eq $true) { break }      # time to go?

                        [System.Windows.Forms.Application]::DoEvents()  # keep form responsive





                        $CurrConn # processing feedback to user
                        
                
                        $RemoteTuple = -join( $CurrConn.OwningProcess, ":", $CurrConn.RemoteAddress, ":", $CurrConn.RemotePort )
                
                        $Match = $table.FindIndex( { $args[0].RemoteTuple -eq $RemoteTuple } )
                
                        if( $Match -eq -1 ) {



                                if($script:CancelLoop -eq $true) { break }      # time to go?

                                [System.Windows.Forms.Application]::DoEvents()  # keep form responsive



                                $App = Get-Process -Id $CurrConn.OwningProcess | % { $_.ProcessName }


                                ###################################################
                                #  some specific conditions to reject connections #
                                ###################################################

                                # ignore uniteresting apps, orphans, AWS, and ip2c.org
                                if( 
                                    ( $RejectedApps -NotContains $App ) `
                                    -and
                                    (  ( -Not ( ( $CurrConn.OwningProcess -eq 0 )               -and ( $App -eq "Idle"       ) ) ) ) `
                                    -and
                                    (  ( -Not ( ( $CurrConn.RemoteAddress -eq "77.55.235.217" ) -and ( $App -eq "powershell" ) ) ) ) `
                                    -and
                                    (  ( -Not ( ( $CurrConn.RemoteAddress -eq "18.191.49.147" ) -and ( $App -eq "powershell" ) ) ) ) `
                                    ) {



                                        ###############
                                        #             #
                                        #   Update    #
                                        #   Object    #
                                        # properties  #
                                        #             #
                                        ###############




                                        if($script:CancelLoop -eq $true) { break }      # time to go?

                                        [System.Windows.Forms.Application]::DoEvents()  # keep form responsive



                                        $IPAddress = $CurrConn.RemoteAddress  # need this for geolocation call, using $CurrConn.RemoteAddress in uri does not work
                                        $Octets = $IPAddress.Split(".")   # less server-side processing to convert IPv4 to decimal
                                        $DecimalIP = [int]$Octets[0]*16777216 + [int]$Octets[1]*65536 + [int]$Octets[2]*256 + [int]$Octets[3]
                                        $Geo = Invoke-RestMethod -Method Get -Uri "http://77.55.235.217/$DecimalIP" # ip2c.org

                                        $SomeCode = ($Geo -split ';')[0]         # always '1'
                                        $Country2letter = ($Geo -split ';')[1]   # US
                                        $Country3letter = ($Geo -split ';')[2]   # USA
                                        $CountryFullName = ($Geo -split ';')[3]  # United States


                                        if( ( $OtherCountries -notcontains $CountryFullName ) -and ( $Country2letter -ne 'US' ) ) { $OtherCountries += $CountryFullName }


                                        [void]$table.add([pscustomobject]@{
                
                                          #  LocalAddress  = $CurrConn.LocalAddress
                                          #  LocalPort     = $CurrConn.LocalPort
                                            RemoteAddress = $CurrConn.RemoteAddress
                                            RemotePort    = $CurrConn.RemotePort
                                            RemoteTuple   = $RemoteTuple
                                            RemoteName    = Resolve-DnsName $CurrConn.RemoteAddress -type PTR -ErrorAction SilentlyContinue | % { $_.NameHost }  # possible multiple values...
                                            GeoLocation   = $CountryFullName
                                            PID           = $CurrConn.OwningProcess
                                            ProcessName   = $App
                                            State         = $CurrConn.State
                                            HitCOunt      = 0  # will increment as local ports are freed up
                        
                                        } # PSCustomObject
                                       ) # .add

                   #
                   # OUTPUT handled later so can be sorted
                   #                            
                   #                    $TextBox.AppendText( -join(   $table[ ( $table.Count -1 ) ].ProcessName    , "   ",
                   #                                                             $table[ ( $table.Count -1 ) ].PID            , "   ",
                   #                                                             $table[ ( $table.Count -1 ) ].RemoteName     , "   ",
                   #                                                             $table[ ( $table.Count -1 ) ].RemoteAddress  , "   ",
                   #                                                             $table[ ( $table.Count -1 ) ].RemotePort     , "   ",
                   #                                                             $table[ ( $table.Count -1 ) ].HitCount       , "   ",
                   #                                                             $table[ ( $table.Count -1 ) ].GeoLocation    , "   ",
                   #                                                   "`r`n" ))

 
                                } # if
                        } # if


                        ###############
                        #             #
                        # gather for: #
                        #  Hit Count  #
                        #             #
                        ###############

                        
                        #
                        #     ***  OVERRIDE TUPPLE  ***
                        #
                        # keep track of original owner to temporaty override orphaned connections, for correct tupple match with hash values
                        if( $CurrConn.OwningProcess -eq 0 ) { $RemoteTuple = $OriginalOwningProcess[$CurrConn.LocalPort] }
                        else{ $OriginalOwningProcess[$CurrConn.LocalPort] = $RemoteTuple }
                        

                        
                        # two ways to update a hash.  1st is add only, 2nd way is update/add if not found.
                        
                        # hash key for both is 'local port', with value = tupple (so we can find the tupple's object index later)

                        $ActiveLocalPorts.Add( $CurrConn.LocalPort , $RemoteTuple )
                        $WatcingLocalPorts[$CurrConn.LocalPort] = $RemoteTuple



                        # for connections being tracked, update state
                        $target = $table.FindIndex( { $args[0].RemoteTuple -eq $RemoteTuple } )
                        if( $target -ne -1 ) { $table[ $target ].State = $CurrConn.State }

                
                } # foreach
 
 

 
                ###############
                #             #
                #   Update    #
                #  Hit Count  #
                #             #
                ###############



 
                 # track local ports to update counters when that port is closed

                $Gone = ( $WatcingLocalPorts.Keys | ? { $ActiveLocalPorts.Keys -NotContains $_  } )


                $Gone | % { 



                        if($script:CancelLoop -eq $true) { break }      # time to go?

                        [System.Windows.Forms.Application]::DoEvents()  # keep form responsive



                        $Tuple = $WatcingLocalPorts[ $_ ]
 
                        $target = $table.FindIndex( { $args[0].RemoteTuple -eq $Tuple } )

                        # coursely weed out hit count for AWS
                        if( $target -ne -1 ) { 
                        
                                $table[ $target ].HitCount++

                                $table[ $target ].State = "Closed"
                                
                        } # if
            
                        $WatcingLocalPorts.Remove($_) 

                } # %

                $ActiveLocalPorts.Clear()





                ###############
                #             #
                #   Output    #
                #             #
                ###############




                $TextBox.Text = ( 
                        
                        $table | Sort-Object -Property ProcessName,PID,RemoteAddress | 
                        
                        Format-Table -autosize -GroupBy ProcessName -Property PID,State,RemoteName,RemoteAddress,RemotePort,HitCount,GeoLocation | 
                        
                        Out-String -width 2048 )  # width defaults to width of console window




                0..( $OtherCountries.Count-1 ) | % { 

                        $SearchString = $OtherCountries[$_]


                
                                   # the unary form of ",", the array constructor operator ensures (by way of a transient aux. array) that 
                                   # the collection returned by the method call is output as a whole. If you omit the "," then the indices are 
                                   # output one by one, resulting in a flat array when collected in a variable.

                       # $indices = , [regex]::Matches( $TextBox.Text , $SearchString ).Index 
                        $indices =   [regex]::Matches( $TextBox.Text , $SearchString ).Index 



                        $indices | % {

                                if( $_ -gt 0 ) {

                                        $TextBox.SelectionStart = $_
                                        $TextBox.SelectionLength = $SearchString.length
                                        $TextBox.SelectionFont = $BoldFont
                                        $TextBox.SelectionColor = 'red'
                                       # $TextBox.SelectionBackColor = 'black'
                                        $TextBox.DeselectAll()
                                        
                              } # if
                        } # $indices
                        
                } # %

                


                ###############
                #             #
                #    Pause    #
                #             #
                ###############



                sleepfor 10  # process interrupts

 
        }  # while

    
    ###############
    #             #
    #   Cleanup   #
    #             #
    ###############


    $Form.DialogResult = 'ok'
    $Form.Close()


}  # scriptblock




###########################################################
#                 Main{} INLINE CODE
###########################################################


Invoke-Command -ScriptBlock $ShowConnections

Exit   # close the text window  <--  not working...
