# EP (ExchangePowerShell) Powershell Module 0.11.0
# Author: Pietro Ciaccio | LinkedIn: https://www.linkedin.com/in/pietrociaccio | Twitter: @PietroCiac
# Pre-requisites: Exchange Server 2016 Exchange Management Shell

# Checks if the powershell session has the Exchange 2016 powershell module loaded

if (!($env:ExchangeInstallPath)) {
    throw "Exchange Server system variable ExchangeInstallPath missing."
}

if ($env:ExchangeInstallPath -notmatch '\\V15\\') {
    write-warning "The Microsoft Exchange Management Shell will be loaded from '$($env:ExchangeInstallPath)'"
    write-warning "Exchange Server 2016 powershell module not detected. There may be issues."
}

try {
    $cmd = $null; $cmd = get-command 'get-mailbox' -erroraction stop
} catch {
    $invexpr = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell"
    Invoke-Expression $invexpr
}

try {
    $cmd = $null; $cmd = get-command 'get-mailbox' -erroraction stop
} catch {
    throw "Unable to load the Microsoft Exchange Management Shell."
}

# Non Exchange specific functions and cmdlets
function Show {

    write-host ""
    write-host "EP (ExchangePowerShell) PowerShell Module 0.11.0" -foreground yellow
    #write-host ""
    write-host "Contribute via BTC: 1FjTBFdQEpfa1rRPCFTG74iDNrvdH5xU9K XRP: rhYuiuxpQLQoVtGrD52s79sGV5kApe5UqY"
    #write-host ""
    write-host "Please contact the author with any comments, issues, or requests for features and improvements via Twitter " -nonewline 
    write-host "@PietroCiac" -foreground white -nonewline
    write-host " or PowerShell Gallery."
    #write-host ""
    write-host "Please read the release notes for information on changes. Use Get-Help with cmdlets for guidance on usage."        
    write-host ""   

}

show

function Update-EPModule {
    <#
    .SYNOPSIS
        Checks for the latest module version on PSGallery.

    .DESCRIPTION
        Checks for the latest module version on PSGallery.

    #>

    Process {
        
        try {

            $FilePath = "$env:temp\EPModuleCheck.txt"
            $LastChecked = $null

            if (test-path $FilePath) {
                $LastChecked = (gi $FilePath).lastwritetime
            } else {
                new-item $FilePath -erroraction stop | out-null
                $LastChecked = (gi $FilePath).lastwritetime
            }

            $dateDiff = $null; $dateDiff = new-timespan $($LastChecked) $(get-date)

            if ($dateDiff.totaldays -ge 7) {
                $InstalledModuleVersion = $null
                try {
                    $InstalledModule = $null; $InstalledModule = Get-InstalledModule ExchangePowershell -erroraction stop
                    $InstalledModuleVersion = $InstalledModule.version.tostring()
                } catch {
                    $InstalledModuleVersion = "0.0.0"
                }
    
                $GalleryModule = $null; $GalleryModule = Find-Module exchangepowershell -Repository psgallery -erroraction stop
                $GalleryModuleVersion = $null; $GalleryModuleVersion = $GalleryModule.version.tostring()
    
                if ($GalleryModuleVersion -ne $InstalledModuleVersion) {
                    write-warning "EP module version $InstalledModuleVersion installed. $GalleryModuleVersion has been published to the PowerShell Gallery. Please update when possible."
                    write-host ""
                } else {
                    write-host "You are running the latest EP module version $InstalledModuleVersion."
                    write-host ""
                } 

                (gi $FilePath).lastwritetime = get-date
            }    

        } catch {
            
        }
    }

}

Update-EPModule

function Get-EPDate {
    <#
    .SYNOPSIS
        Checks date format for cmdlets.

    .DESCRIPTION
        Checks date format for cmdlets

    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$false)][PSCustomObject]$Date
    )

    Process {        
        
        # validate date time
        if ($date.gettype().fullname -ne "System.DateTime") {
            try {
                $date = $date.tostring()
                if ($date -match "\D") {
                    throw "DateTime and numerical value formats supported only."
                }
                switch ($date.length) {
                    7           {$date = "0" + $date}
                    13          {$date = "0" + $date}
                    default     {}
                }
                switch ($date.length) {
                    8           {$date = [datetime]::ParseExact($date,'ddMMyyyy',$null)}
                    14          {$date = [datetime]::ParseExact($date,'ddMMyyyyHHmmss',$null)}
                    default     {throw "Unsupported date provided. Must be ddMMyyyy or ddMMyyyyHHmmss. E.g. 15112019150329 for 15th November 2019 15:03:29."}
                }
            } catch {
                throw $_.exception.message
            }
        }

        if ($date.gettype().fullname -ne "System.DateTime") {
            throw "'Date' format is unsupported. This must be of format System.DateTime or a string in the format of ddMMyyyy or ddMMyyyyHHmmss."
        }

        return $Date

    }
}

# Exchange Server functions and cmdlets
function Get-EPExchangeServer{
    <#
    .SYNOPSIS
        Checks an Exchange 2016 servers identity.

    .DESCRIPTION
        Checks an Exchange 2016 servers identity.

    .PARAMETER Identity
        Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string.
    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity
    )

    Process {        
        
        # Validate Exchange Server
        if ($input) {
            if ($input.objectcategory.name -ne "ms-Exch-Exchange-Server"){
                throw "Unable to validate Exchange server identity."
            } else {
                $ExchangeServer = $null; $ExchangeServer = $input
            }
        }

        if (!($input)) {
            if ($identity.gettype().fullname -ne "System.String") {
                throw "Unable to use parameter 'Identity' of type '$($identity.gettype().fullname)'." 
            } else {
                try {
                    $ExchangeServer = $null; $ExchangeServer = Get-ExchangeServer -Identity $identity -erroraction stop
                } catch {
                    throw $_.exception.message
                }
            }
        }

        if ($ExchangeServer.Admindisplayversion.tostring() -notmatch "^version 15\." ) {
            write-warning "Exchange version is not 15 for '$($ExchangeServer.identity)'. There may be issues."
        }

        return $ExchangeServer
    }
}

function Start-EPRequiredServices {
    <#
    .SYNOPSIS
        Starts Exchange Server services that are required but are not running.

    .DESCRIPTION
        Starts Exchange Server services that are required but are not running. This cmdlet uses information from the Test-ServiceHealth cmdlet.

    .PARAMETER Identity
        Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string.

        
    .EXAMPLE
        Get-ExchangeServer Server1  | Start-EPRequiredServices

        This will start any required Exchange services that are not running on Server1.


    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity
    )
    Process {     

        # Validate identity

        $ExchangeServer = $null
        try {
            if ($input) {            
                $ExchangeServer = $input | Get-EPExchangeServer
            } else {
                $ExchangeServer = Get-EPExchangeServer -Identity $Identity
            }
        } catch {
            throw $_.exception.message         
        }

        write-host "$($ExchangeServer.fqdn.toupper())." 

        try {
            $SH = $null; $SH = Test-ServiceHealth -server $($ExchangeServer.fqdn) -erroraction stop
            $NR = $null; $NR = ($SH.servicesnotrunning | sort-object -unique)
            if ($NR) {
                write-Warning "The following services are not running:"
                $NR
                write-host "Starting required services..."
                Invoke-Command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock {$using:NR | start-service}
                write-host "Done."
            } else {
                write-host "All required services are running on $($ExchangeServer.fqdn.toupper())."
            }
        } catch {
            throw $_.exception.message
        }

    }
}

function Clear-EPExchangeLogs {
    <#
    .SYNOPSIS
        Clears Exchange Server 2016 logs older than a specified date. 

    .DESCRIPTION
        Clears Exchange Server 2016 logs older than a specified date. Deletes files with extensions .log .blg and .etl. The cmdlet will determine the Exchange and IIS logging directories automatically. 

    .PARAMETER Identity
        Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string.

    .PARAMETER StartDate
        Specify the date from which to clear logs. This can be of type Date.Time or a string in the format of ddMMyyyy or ddMMyyyyHHmmss.
    
    .EXAMPLE
        Get-ExchangeServer Server1  | Clear-EPExchangeLogs -Date (get-date).adddays(-30)

        This will clear logs older than 30 days on the Exchange server Server1.

    .EXAMPLE
        Clear-EPExchangeLogs -Identity Server2 -Date 01112019

        This will clear logs older than the 1st November 2019 on the Exchange server Server2.

    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity,
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$false)][PSCustomObject]$StartDate
    )
    Process {      
        
        # Validate identity

        $ExchangeServer = $null
        try {
            if ($input) {            
                $ExchangeServer = $input | Get-EPExchangeServer
            } else {
                $ExchangeServer = Get-EPExchangeServer -Identity $Identity
            }
        } catch {
            throw $_.exception.message         
        }

        # validate date time
        $Date = Get-EPDate -Date $StartDate

        # Script blocks
        $Scriptblock = $null; $Scriptblock = [scriptblock]::Create('
            Param($thisdate,$thisserver)   
            
            $files = $null
            $scopecount = 0
            $successcount = 0
            $errorcount = 0
            $Status = $null
            $Comment = $null

            try {
                $ExchangeLoggingPath = $null; $ExchangeLoggingPath = $env:exchangeinstallpath + "Logging"
                $ExchangeTempUCPath = $null; $ExchangeTempUCPath = $env:exchangeinstallpath + "TransportRoles\data\Temp\UnifiedContent"

                if (test-path $ExchangeLoggingPath) {

                    $Files += [System.IO.Directory]::GetFiles($ExchangeLoggingPath,"*.log","AllDirectories")
                    $Files += [System.IO.Directory]::GetFiles($ExchangeLoggingPath,"*.blg","AllDirectories")
                    $Files += [System.IO.Directory]::GetFiles($ExchangeLoggingPath,"*.etl","AllDirectories")
                    $Files += [System.IO.Directory]::GetFiles($ExchangeTempUCPath,"*","AllDirectories")

                    $IISLoggingPath = $null; 
                    try {
                        $IISLoggingPath = (get-iissite).logfile.directory
                    } catch {
                        throw "Unable to determine IIS logging paths."
                    } 
                    if ($IISLoggingPath) {
                        $IISLoggingPath | . {process 
                            {
                                $Files += [System.IO.Directory]::GetFiles([System.Environment]::ExpandEnvironmentVariables($_),"*.log","AllDirectories")
                            }
                        }
                    }
                } else {
                    throw "$ExchangeLoggingPath does not exist."
                }
    
                if ($Files) {
                    $Files | . {process 
                        {
                            if ($([System.IO.File]::GetLastWriteTime($_)) -lt $thisdate) {
                                $scopecount += 1
                                try {
                                    [System.IO.File]::Delete($_)
                                    $successcount += 1
                                } catch {
                                    $errorcount += 1
                                }                            
                            }
                        }
                    }
                }
                $Status = "OK"
            } catch {
                $Status = "Error"
                $Comment += $_.exception.message
            }   

            [pscustomobject]@{
                "Identity" = $thisServer
                "TotalFound" = $(($Files | measure).count)
                "InScope" = $scopecount
                "Deleted" = $successcount
                "Skipped" = $errorcount
                "Status" = $Status
                "Comment" = $null
            } 
        ')      

        # Local or Remote execution
        try {            
            $localhostname = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
            if ($localhostname -eq $ExchangeServer.fqdn) {
                try {
                    $scriptblock.invoke($date,$($ExchangeServer.fqdn))
                } catch {
                    throw $_.exception.message
                }
            } else {
                try {
                    Invoke-Command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock $Scriptblock -ArgumentList $date,$($ExchangeServer.fqdn) -ErrorAction stop | select * -ExcludeProperty pscomputername,runspaceid
                } catch {
                    throw $_.exception.message
                }
            }            
        } catch {
            [pscustomobject]@{
                "Identity" = $($ExchangeServer.fqdn)
                "TotalFound" = 0
                "InScope" = 0
                "Deleted" = 0
                "Skipped" = 0
                "Status" = "Error"
                "Comment" = $_.exception.message
            }  
        } 
    }
}

function Get-EPMaintenanceMode {
    <#
    .SYNOPSIS
        Checks if a Microsoft Exchange Server 2016 computer is in maintenance mode.
 
    .DESCRIPTION
        Checks if a Microsoft Exchange Server 2016 computer is in maintenance mode.
          
    .PARAMETER Identity
        Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string. 
 
    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity
    )
    Process {  

        # Validate identity
        if ($input) {            
            $ExchangeServer = $null; $ExchangeServer = $input | Get-EPExchangeServer
        } else {
            $ExchangeServer = $null; $ExchangeServer = Get-EPExchangeServer -Identity $Identity
        }

        write-host "$($ExchangeServer.fqdn.toupper())." 

        # Determine DAG membership
        $isDAGMember = $false
        try {
            $RecipientServer = $null; $RecipientServer = Get-MailboxServer -identity $Exchangeserver.fqdn -erroraction stop
            if ($($RecipientServer | measure).count -ne 1) {
                throw "$($($RecipientServer | measure).count) servers returned from query. Unable to continue."
            }
            if ($RecipientServer.DatabaseAvailabilityGroup -ne $null) {
                $isDAGMember = $true
            }
        } catch {
            throw $_.exception.message
        } 

        try {
            # DAG members only
            if ($isDAGMember) {
                $MBServer = $null; $MBServer = $ExchangeServer | Get-MailboxServer -erroraction stop

                if ($MBServer.DatabaseCopyActivationDisabledAndMoveNow -eq $false) {
                    write-host "DatabaseCopyActivationDisabledAndMoveNow is False."
                } else {
                    write-warning "DatabaseCopyActivationDisabledAndMoveNow is True."
                }

                $cn = $null; $cn = invoke-command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock {Get-ClusterNode $($using:ExchangeServer.fqdn)} -ErrorAction Stop
                if ($cn.state -eq "up") {
                    write-host "Cluster node is Up."
                } else {
                    write-warning "Cluster node is not Up."
                }

                if ($MBServer.DatabaseCopyAutoActivationPolicy -eq "unrestricted") {
                    write-host "DatabaseCopyAutoActivationPolicy is Unrestricted."
                } else {
                    write-warning "DatabaseCopyAutoActivationPolicy is $($MBServer.DatabaseCopyAutoActivationPolicy)."
                }

                $Copies = $null; $Copies = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name) 
                if ($Copies) {
                    $Copies | . { process {
                        if ($_.status -notmatch "^healthy$|^mounted$") {
                            write-warning "$($_.name) database copy status is $($_.status)."
                        } else {
                            write-host "$($_.name) database copy status is $($_.status)."
                        }
                    }}
                }

                $CS = $null; $CS = Get-ServerComponentState -Identity $($ExchangeServer.fqdn) -erroraction stop | ? {$_.state -ne "active"}
                if (-not($CS)) {
                    write-host "Server component states active."
                } else {
                    write-warning "Server component states inactive: $($CS.component -join ';')"
                }

                write-host "Done."

            }
        } catch {
            throw $_.exception.message
        }
    }
}

function Enable-EPMaintenanceMode {
    <#
    .SYNOPSIS
        Puts a Microsoft Exchange Server 2016 computer into maintenance mode.
 
    .DESCRIPTION
        Puts a Microsoft Exchange Server 2016 computer into maintenance mode. CmdLet will -
         
            - drain queues
            - restart transport services
            - redirect messages to a redirection server
            - move off active database copies to an available DAG member
            - suspend the cluster node
            - prevent database activation on the server
            - suspend passive copies
            - set all server component states to inactive
 
    .PARAMETER Identity
        Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string.
 
    .PARAMETER RedirectionTarget
        Specify the identity of the computer you wish to redirect pending messages to.
         
    .PARAMETER MoveActiveDatabaseCopies
        Specify whether to move active database copies to other DAG members, if possible. The default is false.
 
    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$false)][PSCustomObject]$RedirectionTarget,
        [switch]$MoveActiveDatabaseCopies
    )
    Process {        

        # Validate identity
        if ($input) {            
            $ExchangeServer = $null; $ExchangeServer = $input | Get-EPExchangeServer
        } else {
            $ExchangeServer = $null; $ExchangeServer = Get-EPExchangeServer -Identity $Identity
        }

        # Validate Redirection Server
        $RedirectionServer = $null; 
        if ($RedirectionTarget) {
            try {
                $RedirectionServer = Get-EPExchangeServer -Identity $RedirectionTarget
            } catch {
                throw $_.exception.message
            }
        }

        # Determine DAG membership
        $isDAGMember = $false
        try {
            $RecipientServer = $null; $RecipientServer = Get-MailboxServer -identity $Exchangeserver.fqdn -erroraction stop
            if ($($RecipientServer | measure).count -ne 1) {
                throw "$($($RecipientServer | measure).count) servers returned from query. Unable to continue."
            }
            if ($RecipientServer.DatabaseAvailabilityGroup -ne $null) {
                $isDAGMember = $true
            }
        } catch {
            throw $_.exception.message
        } 

        # Draining queues
        Write-Host "Putting '$($ExchangeServer.fqdn.toupper())' into maintenance mode."
        if ($RedirectionServer){
            Write-Host "Using '$($RedirectionServer.fqdn.toupper())' for message redirection."
        }
        Write-Host "Draining mail queues."
        try {
            Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component HubTransport -State Draining -Requester Maintenance -erroraction stop
        } catch {
            throw $_.exception.message
        }

        # Restarting transport services
        Write-Host "Restarting MSExchangeTransport and MSExchangeFrontEndTransport services."
        $n = 0
        Do {
            try {
                invoke-command -ComputerName $($ExchangeServer.fqdn) -scriptblock {"MSExchangeTransport","MSExchangeFrontEndTransport" | restart-service -WarningAction SilentlyContinue} -ErrorAction stop -WarningAction SilentlyContinue
                break
            } catch {
                $n++
                write-host "WARNING: Issue restarting MSExchangeTransport and MSExchangeFrontEndTransport services. Waiting 60 seconds then retrying." -nonewline -ForegroundColor Yellow
                start-sleep -Seconds 60
                write-host " Retry attempt $n of 3." -ForegroundColor Yellow      
            }
            if ($n -eq 3) {
                write-warning "Issue restarting MSExchangeTransport and MSExchangeFrontEndTransport services. Continuing."
                break
            }
        } while ($true)

        # Redirect messages
        if ($RedirectionServer) {
            Write-Host "Redirecting messages."      
            try {
                Redirect-Message -Server $($ExchangeServer.fqdn) -Target $($RedirectionServer.fqdn) -confirm:$false -erroraction stop -WarningAction SilentlyContinue
            } catch {
                throw $_.exception.message
            }
        }

        # DAG members only
        if ($isDAGMember) {

            # Move active database copies off
            try {
                Write-Host "Setting DatabaseCopyActivationDisabledAndMoveNow to 'True'."
                Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyActivationDisabledAndMoveNow $True -erroraction Stop -confirm:$false
            } catch {
                throw $_.exception.message
            }

            # Move active copies immediately
            try {
                $actives = $null; $actives = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name) | ? {$_.activecopy -eq $true}
                write-host "$($($actives | measure).count) active database copies found."

                if ($($($actives | measure).count) -eq 0 -and $MoveActiveDatabaseCopies) {
                    Write-host "No active database copies to move."
                }

                if ($actives -and $MoveActiveDatabaseCopies) {
                    write-host "Moving active databases to other DAG members."
                    $actives | . {
                        process {
                            if ($($($($_ | . { process {(get-Mailboxdatabase $_.databasename).servers}}) | measure).count) -lt 2) {
                                Write-Warning "No other database copies exist. Unable to move active database copy."
                            } else {
                                $move = $null;
                                try {
                                    $move = Get-Mailboxdatabase $($_.databasename) |  Move-ActiveMailboxDatabase -MountDialOverride lossless -SkipClientExperienceChecks -SkipMaximumActiveDatabasesChecks -confirm:$false -erroraction stop
                                    if ($move.status -ne "Succeeded") {
                                        throw "$($move.identity) Issue moving active database copy."
                                    }
                                } catch {
                                    write-warning $_.exception.message
                                }
                            }
                        }                        
                    }
                }

            } catch {
                throw $_.exception.message
            }

            # Suspend cluster node
            Write-Host "Suspending cluster node."      
            try {
                invoke-command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock {
                    if ((Get-ClusterNode $($using:ExchangeServer.fqdn)).state -ne "Paused") {
                        Suspend-ClusterNode $($using:ExchangeServer.fqdn)
                    }     
                } -ErrorAction Stop | out-null
            } catch {
                throw $_.exception.message
            }

            # Set activation policy to blocked
            try {
                Write-Host "Setting DatabaseCopyAutoActivationPolicy to 'Blocked'."
                Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyAutoActivationPolicy Blocked -erroraction Stop -confirm:$false
            } catch {
                throw $_.exception.message
            }

            # Suspend passive copies
            try {                
                $Copies = $null; $Copies = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name) | ? {$_.activecopy -eq $false}
                if ($Copies) {
                    Write-Host "Suspending passive copies."                
                    $Copies | . { process {
                            $_  | Suspend-MailboxDatabaseCopy -confirm:$false -erroraction stop
                        }                
                    }
                }
            } catch {
                throw $_.exception.message
            }

        }

        # Complete maintenance mode
        try {
            Write-Host "Completing maintenance mode."
            Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component ServerWideOffline -State Inactive -Requester Maintenance -erroraction stop
        } catch {
            throw $_.exception.message
        }
        
        write-host "Done."     
        
    }
}

function Disable-EPMaintenanceMode {
    <#
    .SYNOPSIS
        Removes a Microsoft Exchange Server 2016 computer from maintenance mode.
 
    .DESCRIPTION
        Removes a Microsoft Exchange Server 2016 computer from maintenance mode. Cmdlet will -
 
            - set all server component states to active
            - resume the cluster node
            - enable database activation on the server
            - resume passive database copies
            - resume transport
            - restart transport services
 
    .PARAMETER Identity
        Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string.
 
    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity
    )
    Process {        

        # Validate identity
        if ($input) {            
            $ExchangeServer = $null; $ExchangeServer = $input | Get-EPExchangeServer
        } else {
            $ExchangeServer = $null; $ExchangeServer = Get-EPExchangeServer -Identity $Identity
        }

        # Determine DAG membership
        $isDAGMember = $false
        try {
            $RecipientServer = $null; $RecipientServer = Get-MailboxServer -identity $Exchangeserver.fqdn -erroraction stop
            if ($($RecipientServer | measure).count -ne 1) {
                throw "$($($RecipientServer | measure).count) servers returned from query. Unable to continue."
            }
            if ($null -ne $RecipientServer.DatabaseAvailabilityGroup) {
                $isDAGMember = $true
            }
        } catch {
            throw $_.exception.message
        } 

        # Remove from maintenance mode
        try {
            Write-Host "Removing '$($ExchangeServer.fqdn.toupper())' from maintenance mode."
            Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component ServerWideOffline -State Active -Requester Maintenance -erroraction stop
        } catch {
            throw $_.exception.message
        } 
        
        # DAG members only
        if ($isDAGMember) {
            
            # Resume cluster node
            Write-Host "Resuming cluster node."       
            try {
                invoke-command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock {
                    if ((Get-ClusterNode $($using:ExchangeServer.fqdn)).state -ne "Up") {
                        Resume-ClusterNode $($using:ExchangeServer.fqdn)
                    }     
                } -ErrorAction Stop | out-null
            } catch {
                throw $_.exception.message
            }

            # Move active database copies on
            try {
                Write-Host "Setting DatabaseCopyActivationDisabledAndMoveNow to 'False'."
                Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyActivationDisabledAndMoveNow $false -erroraction Stop -confirm:$false
            } catch {
                throw $_.exception.message
            }     

            # Set activation policy to unrestricted
            try {
                Write-Host "Setting DatabaseCopyAutoActivationPolicy to 'Unrestricted'."
                Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyAutoActivationPolicy Unrestricted -erroraction Stop -confirm:$false
            } catch {
                throw $_.exception.message
            }   
            
            # Resume passive copies
            try {                
                $Copies = $null; $Copies = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name) | ? {$_.activecopy -eq $false}
                if ($Copies) {
                    Write-Host "Resuming passive copies."
                    $Copies | . { process {
                            $_  | Resume-MailboxDatabaseCopy -confirm:$false -erroraction stop
                        }                
                    }
                }
            } catch {
                throw $_.exception.message
            }

        }     
        
        # Resume transport
        Write-Host "Resuming transport."
        try {
            Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component HubTransport -State Active -Requester Maintenance -erroraction stop
        } catch {
            throw $_.exception.message
        }

        # Restarting transport services
        Write-Host "Restarting MSExchangeTransport and MSExchangeFrontEndTransport services."
        $n = 0
        Do {
            try {
                invoke-command -ComputerName $($ExchangeServer.fqdn) -scriptblock {"MSExchangeTransport","MSExchangeFrontEndTransport" | restart-service -WarningAction SilentlyContinue} -ErrorAction stop -WarningAction SilentlyContinue
                break
            } catch {
                $n++
                write-host "WARNING: Issue restarting MSExchangeTransport and MSExchangeFrontEndTransport services. Waiting 60 seconds then retrying." -nonewline -ForegroundColor Yellow
                Start-Sleep -Seconds 60
                write-host " Retry attempt $n of 3." -ForegroundColor Yellow      
            }
            if ($n -eq 3) {
                write-warning "Issue restarting MSExchangeTransport and MSExchangeFrontEndTransport services. Continuing."
                break
            }
        } while ($true)

        write-host "Done."

    }
}

function Get-EPAntimalwareStatus {
    <#
    .SYNOPSIS
        Gets the native antimalware status of an Exchange Server 2016 server. 

    .DESCRIPTION
        Gets the native antimalware status of an Exchange Server 2016 server. Cmdlet searches the Windows Server Application logs for events and presents the results in an object oriented format. 

    .PARAMETER Identity
        Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string.

    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity
    )
    Process {      
        
        # Validate identity

        $ExchangeServer = $null
        try {
            if ($input) {            
                $ExchangeServer = $input | Get-EPExchangeServer
            } else {
                $ExchangeServer = Get-EPExchangeServer -Identity $Identity
            }
        } catch {
            throw $_.exception.message         
        }
  
        $Check = $null
        $CheckDateTime = $null
        $CheckInfo = $null
        try {
            $Check = Get-WinEvent -ComputerName $($ExchangeServer.fqdn) -FilterHashtable @{logname='application';ProviderName='Microsoft-Filtering-FIPFS';ID='6023','6024';Level='4'} -MaxEvents 1 -erroraction stop
            $CheckDateTime = $Check.TimeCreated
            $CheckInfo = $Check.message
        } catch {}


        $Update = $null
        $UpdateDateTime = $null
        $UpdateInfo = $null
        try {
            $Update = Get-WinEvent -ComputerName $($ExchangeServer.fqdn) -FilterHashtable @{logname='application';ProviderName='Microsoft-Filtering-FIPFS';ID='6033';Level='4'} -MaxEvents 1 -erroraction stop
            $UpdateDateTime = $Update.TimeCreated
            $UpdateInfo = $Update.message
        } catch {}

        $Problem = $null
        $ProblemDateTime = $null
        $ProblemInfo = $null
        try {
            $Problem = Get-WinEvent -ComputerName $($ExchangeServer.fqdn) -FilterHashtable @{logname='application';ProviderName='Microsoft-Filtering-FIPFS';Level='1','2','3'} -MaxEvents 1 -erroraction stop
            $ProblemDateTime = $Problem.TimeCreated
            $ProblemInfo = $Problem.message
        } catch {}

        try {
            [pscustomobject]@{
                "Identity" = $($ExchangeServer.fqdn)
                "CheckInfo" = $CheckInfo
                "UpdateInfo" = $UpdateInfo
                "ErrorInfo" = $ProblemInfo
                "WhenChecked" = $CheckDateTime
                "WhenUpdated" = $UpdateDateTime
                "WhenError" = $ProblemDateTime
            }
        } catch {
            throw $_.exception.message
        }

    }
}


# Mail enabled objects cmdlets and functions

function Remove-EPProxyAddress{
    <#
    .SYNOPSIS
        Removes secondary SMTP addresses from mail enabled objects that match a specific domain.

    .DESCRIPTION
        Removes secondary SMTP addresses from mail enabled objects that match a specific domain.

    .PARAMETER Identity
        Specify the identity of the mail enabled object. This can be piped from Get-Recipient, Get-Recipient, Get-DistributionGroup etc or specified explicitly using a string.

    .PARAMETER SMTPDomain
        Specify the SMTP domain that you wish to have removed from the object's proxy addresses.
    
    .PARAMETER X500Keyword
        Specify a keyword that can be matched to an X500 address that you wish to have removed from the object's proxy addresses.

    .PARAMETER Confirm
        Specify if you want to be prompted. The default is true.

    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true)][string]$SMTPDomain,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true)][string]$X500Keyword,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true)][Boolean]$Confirm=$true
    )

    Process {

        $Pxys = $null
        $Recipient = $null
        $Status = $null
        $Comment = $null
        $SMTPDomain = $SMTPDomain.toupper()

        $MailObj = [pscustomobject]@{
            "Identity" = $null
            "GUID" = $null
            "RecipientTypeDetails" = $null
            "Status" = $Status
            "Comment" = $Comment
        }

        try {
            if ($input) {
                $Recipient = $Input
                if (!($Recipient.recipienttypedetails)) {
                    $MailObj.identity = $Input.identity
                    throw "Issue with input object."
                }
            }

            if (!($input)) {
                $Recipient = Get-Recipient -identity $Identity -erroraction stop                
            }

            if (!($Recipient)) {
                throw "Recipient not found."
            }

            if (($Recipient | measure).count -gt 1) {
                throw "Too many matches found."
            }

            $MailObj.identity = $Recipient.identity
            $MailObj.RecipientTypeDetails = $Recipient.RecipientTypeDetails
            $MailObj.GUID = $Recipient.guid.guid

            if ($Recipient.exchangeversion.exchangebuild.major -ne 15) {
                write-warning "Exchange version is not 15 for '$($Recipient.identity)'. There may be issues."
                $Comment += "Exchange version is not 15. There may be issues."
            }

            $Pxys = ($Recipient).emailaddresses.proxyaddressstring | sort 

            if (!($Pxys)) {
                throw "No proxy addresses found."
            }

            $PxysToRemove = $null

            if (!($SMTPDomain -or $X500Keyword)) {
                throw "No parameters passed to cmdlet."
            }

            if ($SMTPDomain) {
                $PxysToRemove = $Pxys | sort | ? {($_ -match "$SMTPDomain$" -and $_ -cmatch "^smtp:")}
                $Pxys = $Pxys | sort | ? {!($_ -match "$SMTPDomain$" -and $_ -cmatch "^smtp:")}           
            }

            if ($X500Keyword) {
                $PxysToRemove += $Pxys | ? {($_ -match "^X500:" -and $_ -match $X500Keyword)}
                $Pxys = $Pxys | sort | ? {!($_ -match "^X500:" -and $_ -match $X500Keyword)} 
            }


            if (!($PxysToRemove)) {
                $MailObj.Status = "OK"
            } else {
                $ADObj = $null
                $GUID = $null
                $GUID = $Recipient.guid.guid
                $ADObj = Get-ADObject -Filter {objectguid -eq $GUID} -Properties proxyaddresses -Server $recipient.originatingserver
                if (!($ADObj)) {
                    throw "Recipient not found in Active Directory."
                }

                $Pxys = $Pxys | . { process{ $_.tostring()}}
                $ADObj.proxyaddresses = $Pxys
                
                if ($Confirm) {
                    write-host "Are you sure you want to remove proxy addresses from '$($Recipient.identity)'?"
                    write-host ""
                    write-host "Proxy addresses to keep -`n"
                    $Pxys
                    write-host ""
                    Write-host "Proxy addresses to remove -`n"
                    $PxysToRemove
                    write-host ""
                    write-host "[Y] Yes " -ForegroundColor Yellow -NoNewline
                    write-host '[N] No (default is "Y"): ' -NoNewline                    
                    $Read = Read-Host
                }
                if (!($Read)) {
                    $Read = "Y"
                }
                if ($Read -match "^Y$|^Yes$") {
                    Set-ADObject -Instance $ADObj -Server $recipient.originatingserver -erroraction stop
                } else {
                    throw "Cancelled by user."
                }

                $MailObj.status = "Updated"
                $MailObj.Comment += "Removed $($PxysToRemove -join ";")"
            }

        } catch {
            $MailObj.Status = "Error"
            $MailObj.Comment += $_.exception.message
        }
   
        $MailObj

    }
}

function ConvertTo-EPMailUser{
    <#
    .SYNOPSIS
        Converts a mailbox to a mail user.

    .DESCRIPTION
        Converts a mailbox to a mail user.

    .PARAMETER Identity
        Specify the identity of the mailbox enabled object. This can be piped from Get-Recipient, Get-Mailbox etc or specified explicitly using a string.

    .PARAMETER ExternalEmailAddress
        Specify the ExternalEmailAddress of the mail user. If this is not specified the primary SMTP address will be used.
    
    .PARAMETER Confirm
        Specify if you want to be prompted. The default is true.

    #>    
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$false)][string]$ExternalEmailAddress,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true)][Boolean]$Confirm=$true
    )

    Process {

        $MailObj = [pscustomobject]@{
            "Identity" = $null
            "GUID" = $null
            "RecipientTypeDetails" = $null
            "ExternalEmailAddress" = $null
            "Status" = $Status
            "Comment" = $Comment
        }

        try {

            $Recipient = $null

            if ($input) {
                $Recipient = $Input
                if (!($Recipient.recipienttypedetails)) {
                    $MailObj.identity = $Input.identity
                    throw "Issue with input object."
                } else {
                    if (!($Recipient.recipienttypedetails -match "mailbox$" -or $Recipient.recipienttypedetails -match "mailuser$")) {
                        $MailObj.identity = $Input.identity
                        throw "Issue with input object."
                    }
                }
            }

            if (!($input)) {
                $Recipient = Get-Recipient -identity $Identity -erroraction stop                
            }

            if (!($Recipient)) {
                throw "Recipient not found."
            }

            if (($Recipient | measure).count -gt 1) {
                throw "Too many matches found."
            }

            if ($Recipient.exchangeversion.exchangebuild.major -ne 15) {
                write-warning "Exchange version is not 15 for '$($Recipient.identity)'. There may be issues."
                $Comment += "Exchange version is not 15. There may be issues."
            }

            $MailObj.identity = $Recipient.identity
            $MailObj.RecipientTypeDetails = $Recipient.RecipientTypeDetails
            $MailObj.GUID = $Recipient.guid.guid


            if ($ExternalEmailAddress) {               
                $MailObj.ExternalEmailAddress = $ExternalEmailAddress
            } else {
                $MailObj.ExternalEmailAddress = $Recipient.primarysmtpaddress
            }
            
            $TargetAddress = $null
            $TargetAddress = $MailObj.ExternalEmailAddress
            if ($TargetAddress -notmatch "^smtp:") {
                $TargetAddress = "SMTP:" + $TargetAddress
            }

            if ($Recipient.recipienttypedetails -eq "mailuser" -and $Recipient.ExternalEmailAddress -ceq $TargetAddress) {

                $MailObj.Status = "OK"

            } else {

                $ADObj = $null
                $ADObj = Get-ADObject -Identity $MailObj.GUID -Properties homeMDB,msExchHomeServerName,msExchMailboxGUID,msExchArchiveGUID,msExchArchiveName,msExchRecipientDisplayType,msExchRecipientTypeDetails,targetaddress -Server $recipient.originatingserver
                $ADObj.homeMDB = $null
                $ADObj.msExchHomeServerName = $null
                $ADObj.msExchMailboxGUID = $null
                $ADObj.msExchArchiveGUID = $null
                $ADObj.msExchArchiveName = $null
                $ADObj.msExchRecipientDisplayType = 6
                $ADObj.msExchRecipientTypeDetails = 128
                $ADObj.targetAddress = $TargetAddress

                if ($Confirm) {
                    write-host "Are you sure you want to convert the mailbox to a mailuser using ExternalEmailAddress '$TargetAddress' for '$($Recipient.identity)'?"
                    write-host "[Y] Yes " -ForegroundColor Yellow -NoNewline
                    write-host '[N] No (default is "Y"): ' -NoNewline                    
                    $Read = Read-Host
                }
                if (!($Read)) {
                    $Read = "Y"
                }
                if ($Read -match "^Y$|^Yes$") {
                    Set-ADObject -Instance $ADObj -Server $recipient.originatingserver -erroraction stop
                } else {
                    throw "Cancelled by user."
                }

                $MailObj.Status = "Updated"
                $MailObj.Comment += "ExternalEmailAddress set to '$TargetAddress'"
            }

        } catch {
            $MailObj.Status = "Error"
            $MailObj.Comment += $_.exception.message
        }

        $MailObj

    }
}

function ConvertTo-EPMailContact{
    <#
    .SYNOPSIS
        Converts a mailbox or mail user to a mail contact.

    .DESCRIPTION
        Converts a mailbox or mail user to a mail contact. The user object will have Exchange attributes removed and a mail contact object will be created in the same organizational unit unless specified with the OrganizationalUnit parameter.

    .PARAMETER Identity
        Specify the identity of the mailbox or mail enabled object. This can be piped from Get-Recipient, Get-Mailbox etc or specified explicitly using a string.

    .PARAMETER ExternalEmailAddress
        Specify the ExternalEmailAddress of the mail contact. If this is not specified for a mailbox the primary SMTP address will be used.
        If this is not specified for a mail user to ExternalEmailAddress of the mail user will be used.

    .PARAMETER OrganizationalUnit
        Specify where you would like the mail contact to be created. If not specified the same organizationalunit as the user object will be used. 

    .PARAMETER Confirm
        Specify if you want to be prompted. The default is true.

    #>    
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$false)][string]$ExternalEmailAddress,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$false)][string]$OrganizationalUnit,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true)][Boolean]$Confirm=$true
    )

    Process {

        $MailObj = [pscustomobject]@{
            "Identity" = $null
            "GUID" = $null
            "RecipientTypeDetails" = $null
            "ExternalEmailAddress" = $null
            "Status" = $Status
            "Comment" = $Comment
        }

        try {

            $Recipient = $null

            if ($input) {
                $Recipient = $Input
                if (!($Recipient.recipienttypedetails)) {
                    $MailObj.identity = $Input.identity
                    throw "Issue with input object."
                } else {
                    if (!($Recipient.recipienttypedetails -match "mailbox$" -or $Recipient.recipienttypedetails -match "mailuser$" -or $Recipient.recipienttypedetails -match "mailcontact$")) {
                        $MailObj.identity = $Input.identity
                        throw "Issue with input object."
                    }
                }
            }

            if (!($input)) {
                $Recipient = Get-Recipient -identity $Identity -erroraction stop                
            }

            if (!($Recipient)) {
                throw "Recipient not found."
            }

            if (($Recipient | measure).count -gt 1) {
                throw "Too many matches found."
            }

            if ($Recipient.exchangeversion.exchangebuild.major -ne 15) {
                write-warning "Exchange version is not 15 for '$($Recipient.identity)'. There may be issues."
                $Comment += "Exchange version is not 15. There may be issues."
            }

            $MailObj.identity = $Recipient.identity
            $MailObj.RecipientTypeDetails = $Recipient.RecipientTypeDetails
            $MailObj.GUID = $Recipient.guid.guid

            if ($ExternalEmailAddress) {
                $MailObj.ExternalEmailAddress = $ExternalEmailAddress
            } else {
                if ($Recipient.recipienttypedetails -match "mailbox$") {
                    $MailObj.ExternalEmailAddress = $Recipient.primarysmtpaddress.address
                }
                if ($Recipient.recipienttypedetails -match "mailuser$") {
                    if ($Recipient.ExternalEmailAddress) {
                        $MailObj.ExternalEmailAddress = $Recipient.ExternalEmailAddress
                    }
                }
            }

            $TargetAddress = $null
            $TargetAddress = $MailObj.ExternalEmailAddress
            if ($TargetAddress -notmatch "^smtp:") {
                $TargetAddress = "SMTP:" + $TargetAddress
            }

            if ($Recipient.recipienttypedetails -eq "mailcontact") {

                $MailObj.Status = "Error"
                $MailObj.Comment += "MailContact objects not supported."

            } else {

                if ($Confirm) {
                    write-host "Are you sure you want to convert to a mail contact using ExternalEmailAddress '$TargetAddress' for '$($Recipient.identity)'?"
                    write-host "[Y] Yes " -ForegroundColor Yellow -NoNewline
                    write-host '[N] No (default is "Y"): ' -NoNewline                    
                    $Read = Read-Host
                }
                if (!($Read)) {
                    $Read = "Y"
                }
                if ($Read -match "^Y$|^Yes$") {    
        
                    if (!($Recipient.primarysmtpaddress.address)) {
                        throw "Missing primary SMTP address."
                    }

                    $TempSMTPAddress = $null
                    $TempSMTPAddress = [guid]::NewGuid().guid + "@tempaddr.local"

                    $Pxys = $null; $Pxys = @()
                    $Pxys += $Recipient.emailaddresses.proxyaddressstring

                    $X500 = (get-adobject -Identity $recipient.guid.guid -Properties legacyexchangedn -Server $recipient.originatingserver).legacyexchangedn
                    
                    if (!($X500)) {
                        throw "LegacyExchangeDN is null."
                    }

                    $X500 = "X500:" + $X500
                    $Pxys += $X500
                    $Pxys = $Pxys | . {process{$_.tostring()}}

                    $ContactName = $null
                    $ContactName = $Recipient.name + "-MailContact"
                    $ConObj = $null

                    if (!($OrganizationalUnit)) {
                        $ConObj = New-MailContact -Name $ContactName -Displayname $($Recipient.Displayname) -OrganizationalUnit $($Recipient.OrganizationalUnit) -PrimarySMTPAddress $TempSMTPAddress -ExternalEmailAddress $TargetAddress -erroraction stop
                    } else {
                        $ConObj = New-MailContact -Name $ContactName -Displayname $($Recipient.Displayname) -OrganizationalUnit "$OrganizationalUnit" -PrimarySMTPAddress $TempSMTPAddress -ExternalEmailAddress $TargetAddress -erroraction stop
                    }

                    if (!($ConObj)) {
                         throw "Issue creating mail contact."
                    } else {
                        if ($ConObj.EmailAddressPolicyEnabled) {
                            Set-MailContact -Identity $ConObj.guid.guid -EmailAddressPolicyEnabled $false -domaincontroller $recipient.originatingserver -erroraction stop
                        }
                    }

                    switch -regex ($Recipient.recipienttypedetails) {
                        'mailbox$'      {
                                            Disable-Mailbox -Identity $Recipient.guid.guid -confirm:$false -erroraction stop
                                            Set-ADObject -Identity $recipient.guid.guid -clear legacyexchangedn -Server $recipient.originatingserver -erroraction stop
                                        }
                        'mailuser$'     {
                                            Disable-MailUser -Identity $Recipient.guid.guid -confirm:$false -erroraction stop    
                                            Set-ADObject -Identity $recipient.guid.guid -clear legacyexchangedn -Server $recipient.originatingserver -erroraction stop                                      
                                        }
                        default { throw "Unsupported '$($Recipient.recipienttypedetails)'"}
                    }
                    
                    Set-MailContact -identity $ConObj.guid.guid -EmailAddresses $Pxys -domaincontroller $recipient.originatingserver -erroraction stop                

                } else {
                    throw "Cancelled by user."
                }

                $MailObj.Status = "Updated"
                $MailObj.Comment += "Created '$ContactName'. ExternalEmailAddress set to '$TargetAddress'"
            }

        } catch {
            $MailObj.Status = "Error"
            $MailObj.Comment += $_.exception.message
        }

        $MailObj

    }
}

function Clear-EPAutoMapping{
    <#
    .SYNOPSIS
        Removes all automapped mailboxes.

    .DESCRIPTION
        Removes all automapped mailboxes. Please note this does not remove any mailbox permissions.

    .PARAMETER Identity
        Specify the identity of the mailbox or mail enabled object. This can be piped from Get-Recipient, Get-Mailbox etc or specified explicitly using a string.

    .PARAMETER Confirm
        Specify if you want to be prompted. The default is true.

    #>    
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity,
        [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true)][Boolean]$Confirm=$true
    )

    Process {

        $MailObj = [pscustomobject]@{
            "Identity" = $null
            "GUID" = $null
            "RecipientTypeDetails" = $null
            "Status" = $Status
            "Comment" = $Comment
        }

        try {

            $Recipient = $null

            if ($input) {
                $Recipient = $Input
                if (!($Recipient.recipienttypedetails)) {
                    $MailObj.identity = $Input.identity
                    throw "Issue with input object."
                } else {
                    if (!($Recipient.recipienttypedetails -match "mailbox$" -or $Recipient.recipienttypedetails -match "mailuser$")) {
                        $MailObj.identity = $Input.identity
                        $MailObj.RecipientTypeDetails = $Recipient.RecipientTypeDetails
                        $MailObj.GUID = $Recipient.guid.guid
                        throw "Issue with input object."
                    }
                }
            }

            if (!($input)) {
                $Recipient = Get-Recipient -identity $Identity -erroraction stop                
            }

            if (!($Recipient)) {
                throw "Recipient not found."
            }

            if (($Recipient | measure).count -gt 1) {
                throw "Too many matches found."
            }

            if ($Recipient.exchangeversion.exchangebuild.major -ne 15) {
                write-warning "Exchange version is not 15 for '$($Recipient.identity)'. There may be issues."
                $Comment += "Exchange version is not 15. There may be issues."
            }

            $MailObj.identity = $Recipient.identity
            $MailObj.RecipientTypeDetails = $Recipient.RecipientTypeDetails
            $MailObj.GUID = $Recipient.guid.guid

            $AutoBL = $null; $AutoBL = (Get-ADObject $($MailObj.GUID) -properties msExchDelegateListBL).msExchDelegateListBL

            if (!($AutoBL)) {
                $MailObj.Status = "OK"
                $MailObj.Comment += "No automapping found."
            } else {

                if ($Confirm) {
                    write-host "Are you sure you want to remove all automapping for '$($Recipient.identity)'?"
                    write-host ""
                    write-host "The following will be removed -"
                    write-host ""
                    $AutoBL
                    write-host ""
                    write-host "[Y] Yes " -ForegroundColor Yellow -NoNewline
                    write-host '[N] No (default is "Y"): ' -NoNewline            
                    $Read = Read-Host
                }
                if (!($Read)) {
                    $Read = "Y"
                }

                if ($Read -notmatch "^Y$|^Yes$") {
                    throw "Cancelled by user."
                }

                $AutoBL | . { process {
                    $ADObj = $null;
                    $DN = $null; $DN = $_
                    try {
                        $ADObj = Get-ADObject $DN -Properties msExchDelegateListLink -Server $recipient.originatingserver
                        $ADObj.msExchDelegateListLink = $ADObj.msExchDelegateListLink | . { process {if ($_ -ne $recipient.distinguishedname){$_}}} | . { process {$_.tostring()}}
                        Set-ADObject -Instance $ADObj -Server $recipient.originatingserver
                        $MailObj.Comment += "Removed: $DN`n"
                    } catch {
                        $MailObj.Comment += "Error: $DN $($_.exception.message)`n"
                    }
                }}
                if ($MailObj.comment -notmatch "Error\:") {
                    $MailObj.Status = "OK"
                } else {
                    if ($MailObj.comment -notmatch "Removed\:") {
                        $MailObj.Status = "Error"
                    } else {
                        $MailObj.Status = "Warn"
                    }
                }
            }

        } catch {
            $MailObj.Status = "Error"
            $MailObj.Comment += $_.exception.message
        }

        $MailObj

    }
}

function Convert-EPIMCEAEXtoX500{
    <#
    .SYNOPSIS
        Converts IMCEAEX to X500 format.

    .DESCRIPTION
        Converts IMCEAEX to X500 format which is a useable proxy address string for recipients. Note the following -

        - If a recipient was accidentally mail or mailbox disabled the legacyexchangedn property would be cleared on the recipient object.
        - If the recipient is mail enabled again, Exchange would generate a new legacyexchangedn.
        - This creates a problem for cached Outlook lookups that used the old legacyexchangedn value.
        - Exchange will generate an undeliverable report if a stale cached lookup is used and presents the address in IMCEAEX format. E.g. IMCEAEX-_o=Example+20Org_ou=First+20Administrative+20Group_cn=Recipients_cn=Example+2ERecipient@domain.com
        - You can convert the IMCEAEX to X500 format and add it to the recipients emailaddresses property. Cached Outlook lookups will work again. 

    .PARAMETER IMCEAEX
        Specify the IMCEAEX address as a string.

    #>    
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true)][string]$IMCEAEX
    )

    Process {
        try {
            $X500 = $null; $X500 = $IMCEAEX.replace("+20", " ").replace("+28", "(").Replace("+29", ")").replace("IMCEAEX-", "X500:").replace("_", "/").replace("+2E", ".").replace("+2C", ",").split("@")[0]
        } catch {
            throw $_.exception.message
        }
        return $X500
    }
}

function Import-EPPSTs {
    <#
    .SYNOPSIS
        Bulk import multiple PST files.

    .DESCRIPTION
        Used to import multiple PST files from a directory into a mailbox or online archive mailbox.

    .PARAMETER Identity
        Specify the identity of the mailbox. This can be piped from Get-ExchangeServer or specified explicitly using a string.
    
    .PARAMETER Path
        Specify the unc path to the PST files. Cmdlet will recursively locate PST files to import. The Exchange Trusted Subsystem must have permissions.

    .PARAMETER ToArchiveMailbox
        Specify whether to import PST files to an online archive mailbox. This is enabled by default.

    .PARAMETER Confirm
        Specify whether a confirmation is required before importing. Default is True.
        
    .EXAMPLE
        Import-EPPSTs -Identity User01 -Path C:\PSTs -ToArchiveMailbox $false

        The above command will import the PSTs found in C:\PSTs into the primary mailbox for the identity User01.

    #>
    [cmdletbinding()]
    Param (
        [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)][PSCustomObject]$Identity,
        [Parameter(mandatory=$true)][String]$Path,
        [Parameter(mandatory=$false)][Boolean]$ImportToArchive=$true,
        [Parameter(mandatory=$false)][Boolean]$Confirm=$true
    )

    Process {

        $MailObj = [pscustomobject]@{
            "Identity" = $null
            "GUID" = $null
            "RecipientTypeDetails" = $null
            "Path" = $Path
            "ImportToArchive" = $ImportToArchive
            "PSTs" = @()
            "TotalSize(MB)" = $null
            "Commands" = @()
            "Status" = $Status
            "Comment" = $Comment
        }

        try {

            $Recipient = $null
            $timestamp = ("{0:yyyyMMddHHmmss}" -f (get-date)).tostring()

            if ($input) {
                $Recipient = $Input
                if (!($Recipient.recipienttypedetails)) {
                    $MailObj.identity = $Input.identity
                    throw "Issue with input object."
                } else {
                    if (!($Recipient.recipienttypedetails -match "mailbox$" -or $Recipient.recipienttypedetails -match "mailuser$")) {
                        $MailObj.identity = $Input.identity
                        $MailObj.RecipientTypeDetails = $Recipient.RecipientTypeDetails
                        $MailObj.GUID = $Recipient.guid.guid
                        throw "Issue with input object."
                    }
                }
            }

            if (!($input)) {
                $Recipient = Get-Recipient -identity $Identity -erroraction stop                
            }

            if (!($Recipient)) {
                throw "Recipient not found."
            }

            if (($Recipient | measure).count -gt 1) {
                throw "Too many matches found."
            }

            if ($Recipient.exchangeversion.exchangebuild.major -ne 15) {
                write-warning "Exchange version is not 15 for '$($Recipient.identity)'. There may be issues."
                $Comment += "Exchange version is not 15. There may be issues."
            }

            $MailObj.identity = $Recipient.identity
            $MailObj.RecipientTypeDetails = $Recipient.RecipientTypeDetails
            $MailObj.GUID = $Recipient.guid.guid

            $dbname = $null;
            if ($importtoarchive) {
                $dbname = $recipient.archivedatabase.name
            } else {
                $dbname = $recipient.database.name
            }

            if (!($dbname)) {
                if ($importtoarchive) {
                    throw "No archive database has been identified for this recipient."
                } else {
                    throw "No database has been identified for this recipient."
                }
            }

            if ($dbname) {
                $circ = $null; $circ = (Get-MailboxDatabase $($dbname)).circularloggingenabled
                if (!($circ)){
                    write-warning "Circular logging is not enabled for database '$($dbname)'. If you have a large amount of mail data to import this may result in significant transaction log growth for this database."
                    $MailObj.Comment += "Circular logging is not enabled for database '$($dbname)'. "
                }
            }

            if (!(test-path $Path)) {
                $Comment += "Issue with path."
                throw "Issue with path."
            } else {
                if ($Path -notmatch "^\\\\") {
                    $Comment += "Path is not a UNC path."
                    throw "Path is not a UNC path."
                }
            }

            $PSTs = $null; $PSTs = gci $Path *.pst -recurse
            $PSTs | % {$MailObj.PSTs += $_.versioninfo.filename}
            $MailObj."totalsize(mb)" = ($PSTs | measure length -sum ).sum / 1000000

            if ($MailObj.PSTs) {
                $n = 0
                $MailObj.PSTs | . {process {
                    $PST = $null; $PST = $_
                    $thiscmd = $null; $thiscmd = "new-mailboximportrequest -BadItemLimit 1000 -AcceptLargeDataLoss -name ""$($mailobj.guid)`-$timestamp`-$n"" -mailbox $($mailobj.guid) -filepath $_ -erroraction stop"
                    if ($importtoarchive) {
                        $thiscmd += " -isarchive"
                    }
                    $MailObj.Commands += $thiscmd
                }}
            }

            if (!($MailObj.Commands)) {
                throw "No PSTs found to import."
            }

            if ($Confirm) {
                if ($importtoarchive) {
                    write-host "Are you sure you want to import $($MailObj."totalsize(mb)") MB of PST mail data into the archive mailbox for '$($Recipient.identity)'?"
                } else {
                    write-host "Are you sure you want to import $($MailObj."totalsize(mb)") MB of PST mail data into the primary mailbox for '$($Recipient.identity)'?"
                }
                write-host "[Y] Yes " -ForegroundColor Yellow -NoNewline
                write-host '[N] No (default is "Y"): ' -NoNewline                    
                $Read = Read-Host
            }
            if (!($Read)) {
                $Read = "Y"
            }
            if ($Read -match "^Y$|^Yes$") {
                
            } else {
                throw "Cancelled by user."
            }
            
            if ($MailObj.Commands) {
                $MailObj.Commands | . { process {
                    $thiscmd = $null; $thiscmd = $_
                    try {                        
                        invoke-expression "[void]($_)" -erroraction stop
                    } catch {
                        $MailObj.comment += "Issue invoking '$thiscmd'."
                        $MailObj.status = "Warning"
                        $_                        
                    }
                }}
            }

        } catch {
            $MailObj.Status = "Error"
            $MailObj.Comment += $_.exception.message
        }

        if (!($MailObj.status)) {
            $MailObj.status = "OK"
        }

        $MailObj

    }
}



Function Compare-ObjectProperties {
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject  
    )
    $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops = $objprops | Sort | Select -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff) {            
            $diffprops = @{
                PropertyName=$objprop
                RefValue=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                DiffValue=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }        
    }
    if ($diffs) {return ($diffs | Select PropertyName,RefValue,DiffValue)}     
}




