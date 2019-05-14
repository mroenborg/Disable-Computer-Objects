<#

    .SYNOPSIS

    .DESCRIPTION

    .PARAMETER

    .EXAMPLE

    .NOTES
    Author: Morten Rønborg
    Date: 16-04-2018
    Last Updated: 16-04-2018

#>
#Requires –Modules ActiveDirectory
################################################
##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region Function Write-Log
Function Write-Log
{
    param 
    (
        [Parameter(Mandatory=$true, HelpMessage="Provide a message")][string]$LogOutput,
        [Parameter(Mandatory=$true, HelpMessage="Provide the function name")][string]$FunctionName,
        [Parameter(Mandatory=$false, HelpMessage="Provide the scriptlinenumber")][string]$ScriptLine,
        [Parameter(Mandatory=$false, HelpMessage="Provide path, default is .\Logs")][string]$Path,
        [Parameter(Mandatory=$false, HelpMessage="Provide name for the logs")][string]$Name,
        [Parameter(Mandatory=$false, HelpMessage="Provide level, 1 = default, 2 = warning 3 = error")][ValidateSet(1, 2, 3)][int]$LogLevel = 1
    )

    #If the scriptline is not defined then use from the invocation
    If(!($ScriptLine)){
        $ScriptLine = $($MyInvocation.ScriptLineNumber)
    }

    if($LogOutput){

        #Date for the lognaming
        $FullLogName = ($Path + "\" + $Name + ".log")
        $FullSecodaryLogName = ($FullLogName).Replace(".log",".lo_")

        #If the log has reached over xx mb then rename it
        if(Test-Path $FullLogName){
            if((Get-Item $FullLogName).Length -gt 5000kb){
                if(Test-Path $FullSecodaryLogName){
                    Remove-Item -Path $FullSecodaryLogName -force
                }
                Rename-Item -Path $FullLogName -NewName $FullSecodaryLogName
            }
        }

        #First check if folder/logfile exists, if not then create it
        if(!(test-path $Path)){
            New-Item -ItemType Directory -Force -Path $Path -ErrorAction SilentlyContinue
        }

        #Get current date and time to write to log
        $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"

        #Construct the logline format
        $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'

        #Define line
        $LineFormat = $logOutput, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($FunctionName):$($Scriptline)", $LogLevel

        #Append line
        $Line = $Line -f $LineFormat

        #Write log
        try {
            Write-Host ("[$($FunctionName):$($Scriptline)]" + $logOutput)
            $Line | Out-File -FilePath ($Path + "\" + $Name + ".log") -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Stop'
        }
        catch {
            Write-Host "$_"
        }
    }
}
#endregion
##*=============================================
##* END FUNCTION LISTINGS
##*=============================================
##*=============================================
##* VARIABLES LISTINGS
##*=============================================
$LogLocation = "$PSScriptRoot\Logs"
$LogName = "Disable-Workstations"

$DomainName = "mroenborg.dk"
$SearchOUs = @("OU=AD Computers,DC=mroenborg,DC=dk","CN=Computers,DC=mroenborg,DC=dk")
$MoveToOU =  "OU=Disabled Computers,DC=mroenborg,DC=dk"

#Number of days from today since the last logon. This will impact on when to disable and move the object.
$Days = -120

##*=============================================
##* END VARIABLES LISTINGS
##*=============================================
Write-Log -LogOutput ("*********************************************** SCRIPT START ***********************************************") -FunctionName $($MyInvocation.MyCommand)-Path $LogLocation -Name $LogName | Out-Null
Write-Log -LogOutput ("Fetching information from objects in AD and moving objects to '$($MoveToOU)' for workstation objects that have lastLogon '$($Days)' days ago..") -FunctionName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName | Out-Null

#Get the date between now and the $Days.
$LastLogonDate = (Get-Date).AddDays($Days)
try{

    #Ensure that the variable is defined as an array
    $Computers = @()

    #Add to the array from the defined OUs
    foreach($OU in $SearchOUs){

        #Add all computer objects that is enabled, and last logon is XXX days.
        $Computers += Get-ADComputer -Property Name,lastLogonDate -Filter {lastLogonDate -lt $LastLogonDate -AND Enabled -eq $true} -SearchBase $OU

        #Add all computers that have been created but have never logged on, and is created more than xx days ago
        $Computers += Get-ADComputer -Property Name,lastLogonDate,whenCreated -Filter  {lastlogondate -notlike "*" -AND whenCreated -lt $LastLogonDate -AND Enabled -eq $true} -SearchBase $OU
    }

    #Disable the computer object.
    Write-Log -LogOutput  ("Number og computers to disable is '$($Computers.Count)'. Computers:`n$($Computers.Name -join "`n")") -FunctionName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName | Out-Null
    $Computers | Set-ADComputer -Server $DomainName -Enabled $false -Description ("Disabled (Script) - " + (Get-Date -UFormat "%d-%m-%Y %H:%M:%S")) 

    #Move all the computers to the disabled OU.
    $Computers | Move-ADObject -Server $DomainName -TargetPath $MoveToOU
}
catch {
    Write-Log -LogOutput ("$_") -FunctionName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName -LogLevel 3 | Out-Null
}
Write-Log -LogOutput ("*********************************************** SCRIPT END ***********************************************") -FunctionName $($MyInvocation.MyCommand)-Path $LogLocation -Name $LogName | Out-Null
