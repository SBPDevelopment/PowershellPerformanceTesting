#Script created by Steve Patterson
#created on 10/20/2014
#special thanks to  Prashanth Jayaram for code excerpts for SQL calls
#special thanks to Praveen Hebbar for code excerpts on powershell functions and wmi calls
#much of the code below is straight from Preshanth and Praveen examples for various other functions
#there is a lot left to do, such as better handling of the SQL connections/curser opens/closes
#better documentation, etc.  

#Define Input and output filepath
#legacy info in case we want to put the information in a log file rather than SQL database
#this script runs locally on the computer rather than aggregated accross a network.
#this way it can be pushed out and run on a schedule for short bursts in measurements.
$servers=$env:computername  
$outfile="\\Fileshare\ClientPerfReport.log"

################################################################################################################

#Function to have the customized output in CSV format
function Export-CsvFile {
[CmdletBinding(DefaultParameterSetName='Delimiter',
  SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
[Parameter(Mandatory=$true, ValueFromPipeline=$true,
           ValueFromPipelineByPropertyName=$true)]
[System.Management.Automation.PSObject]
${InputObject},

[Parameter(Mandatory=$true, Position=0)]
[Alias('PSPath')]
[System.String]
${Path},

#region -Append 
[Switch]
${Append},
#endregion 

[Switch]
${Force},

[Switch]
${NoClobber},

[ValidateSet('Unicode','UTF7','UTF8','ASCII','UTF32','BigEndianUnicode','Default','OEM')]
[System.String]
${Encoding},

[Parameter(ParameterSetName='Delimiter', Position=1)]
[ValidateNotNull()]
[System.Char]
${Delimiter},

[Parameter(ParameterSetName='UseCulture')]
[Switch]
${UseCulture},

[Alias('NTI')]
[Switch]
${NoTypeInformation})

begin
{
# This variable will tell us whether we actually need to append
# to existing file
$AppendMode = $false

 try {
  $outBuffer = $null
  if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
  {
      $PSBoundParameters['OutBuffer'] = 1
  }
  $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Export-Csv',
    [System.Management.Automation.CommandTypes]::Cmdlet)
        
        
                #String variable to become the target command line
                $scriptCmdPipeline = ''

                # Add new parameter handling
                #region Dmitry: Process and remove the Append parameter if it is present
                if ($Append) {
  
                                $PSBoundParameters.Remove('Append') | Out-Null
    
  if ($Path) {
   if (Test-Path $Path) {        
    # Need to construct new command line
    $AppendMode = $true
    
    if ($Encoding.Length -eq 0) {
     # ASCII is default encoding for Export-CSV
     $Encoding = 'ASCII'
    }
    
    # For Append we use ConvertTo-CSV instead of Export
    $scriptCmdPipeline += 'ConvertTo-Csv -NoTypeInformation '
    
    # Inherit other CSV convertion parameters
    if ( $UseCulture ) {
     $scriptCmdPipeline += ' -UseCulture '
    }
    if ( $Delimiter ) {
     $scriptCmdPipeline += " -Delimiter '$Delimiter' "
    } 
    
    # Skip the first line (the one with the property names) 
    $scriptCmdPipeline += ' | Foreach-Object {$start=$true}'
    $scriptCmdPipeline += '{if ($start) {$start=$false} else {$_}} '
    
    # Add file output
    $scriptCmdPipeline += " | Out-File -FilePath '$Path' -Encoding '$Encoding' -Append "
    
    if ($Force) {
     $scriptCmdPipeline += ' -Force'
    }

    if ($NoClobber) {
     $scriptCmdPipeline += ' -NoClobber'
    }   
   }
  }
} 
  

  
 $scriptCmd = {& $wrappedCmd @PSBoundParameters }

 if ( $AppendMode ) {
  # redefine command line
  $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock(
      $scriptCmdPipeline
    )
} else {
  # execute Export-CSV as we got it because
  # either -Append is missing or file does not exist
  $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock(
      [string]$scriptCmd
    )
}

# standard pipeline initialization
$steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
$steppablePipeline.Begin($PSCmdlet)

 } catch {
   throw
}
    
}

process
{
  try {
      $steppablePipeline.Process($_)
  } catch {
      throw
  }
}

end
{
  try {
      $steppablePipeline.End()
  } catch {
      throw
  }
}

}

################################################################################################################

#The actual script start here to gather information

function Global:Convert-HString {      
[CmdletBinding()]            
 Param             
   (
    [Parameter(Mandatory=$false,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [String]$HString
   )#End Param

Begin 
{
    Write-Verbose "Converting Here-String to Array"
}#Begin
Process 
{
    $HString -split "`n" | ForEach-Object {
    
        $ComputerName = $_.trim()
        if ($ComputerName -notmatch "#")
            {
                $ComputerName
            }    
        }
}#Process
End 
{
    # Nothing to do here.
}#End
}#Convert-HString


#Performance counters declaration
function Get-CounterStats { 
param 
    ( 
    [String]$ComputerName = $ENV:ComputerName
    ) 

$Object =@()


$Counter = @" 
Processor(_total)\% processor time 
system\processor queue length
Memory\Available MBytes 
Memory\% committed bytes in use
PhysicalDisk(*)\Current Disk Queue Length 
PhysicalDisk(*)\Avg. Disk sec/Transfer 
PhysicalDisk(*)\% Disk Time 
PhysicalDisk(*)\Disk Read Bytes/sec 
PhysicalDisk(*)\Disk Write Bytes/sec 
Network Interface(*)\Bytes Total/sec
<#Hyper-V Hypervisor Virtual Processor(*)\% Guest Run Time #>
"@ 

        (Get-Counter -ComputerName $ComputerName -Counter (Convert-HString -HString $Counter)).counterSamples |  
        ForEach-Object { 
        $path = $_.path 
        New-Object PSObject -Property @{
        computerName=$ComputerName;
        Counter        = ($path  -split "\\")[-2,-1] -join "-" ;
        Item        = $_.InstanceName ;
        Value = [Math]::Round($_.CookedValue,2) 
        datetime=(Get-Date -format "yyyy-MM-d hh:mm:ss")
        } 
        }
} 

#Collecting counter information for target servers
foreach($server in $Servers)
{
$d=Get-CounterStats -ComputerName $server |Select-Object computerName,Counter,Item,Value,datetime
$d1=Get-CounterStats -ComputerName $server |Select-Object computerName
$d2=Get-CounterStats -ComputerName $server |Select-Object Counter
$d3=Get-CounterStats -ComputerName $server |Select-Object Item
$d4=Get-CounterStats -ComputerName $server |Select-Object Value
$d5=Get-CounterStats -ComputerName $server |Select-Object datetime

#Put into SQL Here:
function Do-FilesInsertRowByRow ([Data.SqlClient.SqlConnection] $OpenSQLConnection) {
    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand
    $sqlCommand.Connection = $sqlConnection

$sqlCommand.CommandText = "SET NOCOUNT ON; " + 
        "INSERT INTO dbo.tblClientPerfMonInfo (ComputerName,MeasureName,MeasureSubName,MeasureValue,MeasureDateTime) " + 
        "VALUES (@ComputerName,@MeasureName,@MeasureSubName,@MeasureValue,@MeasureDateTime); " +
        "SELECT SCOPE_IDENTITY() as [InsertedID]; "
    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@ComputerName",[Data.SQLDBType]::NVarChar,100))) | Out-Null
    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@MeasureName",[Data.SQLDBType]::NVarChar,100))) | Out-Null
    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@MeasureSubName",[Data.SQLDBType]::DateTime2))) | Out-Null
    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@MeasureValue",[Data.SQLDBType]::Money))) | Out-Null    
    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@MeasureDateTime",[Data.SQLDBType]::DateTime2))) | Out-Null
    $sqlCommand.Parameters[0].Value = $d1
    $sqlCommand.Parameters[1].Value = $d2
    $sqlCommand.Parameters[2].Value = $d3
    $sqlCommand.Parameters[3].Value = $d4
    $sqlCommand.Parameters[3].Value = $d5
     
    # Run the query and get the scope ID back into $InsertedID
    $InsertedID = $sqlCommand.ExecuteScalar()
}

# Open SQL connection (you have to change these variables)
# First off, let me state that this is an insert statement, and not calling a proc. 
# yes yes, bad bad
# Fortunately, this is a proof of concept right now.  
$DBServer = "SQLServer\default"
$DBName = "DBName"
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$sqlConnection.ConnectionString = "Server=$DBServer;Database=$DBName;Integrated Security=True;"
$sqlConnection.Open()
# Quit if the SQL connection didn't open properly.
if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
    "Connection to DB is not open."
    Exit
}
# Call the function that does the inserts.
Do-FilesInsertRowByRow ($sqlConnection)
# Close the connection.
if ($sqlConnection.State -eq [Data.ConnectionState]::Open) {
    $sqlConnection.Close()
} 


$d |Export-CsvFile $outfile  -Append -NoTypeInformation
}
#End of Script