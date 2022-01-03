<#
.SYNOPSIS
  Name: Check-NspPrivateCertificates.ps1
  Validate all private certificates found in the NoSpamProxy certificate store.

.DESCRIPTION
    This Script can be used to evaluate possible problems with the usage which forces a validation. If a certificate can not be validated, the script will show the exact issue and which system it occured on.

.PARAMETER SqlCredential
	Receives a PowerShell PowerShell-Credential-Object for logging in to the database.
    By default, the script will use the credentials of the current user.
    
.PARAMETER SqlDatabase
	Specifies the name of the Database of the Intranet Role the script will use to gather private key thumbprints (default value: "NoSpamProxyAddressSynchronization").

.PARAMETER SqlInstance
    Specifies the SQL-instance name the script will connect to (default value: "NoSpamProxy").

.PARAMETER SqlServer
    Specifies the SQL-server the script will connect to (default value: "(local)").

.OUTPUTS
	This script outputs an overview of all unsuccessfully validated private keys and the respective reason why it could not succeed validation.

.NOTES
  Version:        1.0
  Author:         Finn Schulte
  Creation Date:  2021-12-20
  Purpose/Change: Initial release
  
.LINK
  https://www.nospamproxy.de
  https://forum.nospamproxy.de
  https://github.com/noSpamProxy

.EXAMPLE
  .\Check-NspPrivateCertificates.ps1
  This will check all certificates based on a standard installation with a local SQL server.

.EXAMPLE
  .\Check-NspPrivateCertificates.ps1 -SqlServer 10.0.0.2 -SqlInstance Mail
  This will check all certificates based on an installation with outsourced SQL server 10.0.0.2 with a named instance "Mail" containing the database of the Intranet Role.

.EXAMPLE
	.\Check-NspPrivateCertificates.ps1 -SqlServer 10.0.0.2 -SqlInstance Mail -SqlDatabase NSPIntranet -SqlCredential $Credentials
    This will check all certificates based on an installation with outsourced SQL server 10.0.0.2 with a named instance "Mail" containing the database of the Intranetrole named "NSPIntranet" using the provided credentials. 
    The credentials provided can be created by using e.g. $Credentials = Get-Credentials.
#>

param (
# userParams are used for SQL-Connection
	# sql credentials
	[Parameter(Mandatory=$false)][pscredential] $SqlCredential,
	# database name
	[Parameter(Mandatory=$false)][string] $SqlDatabase = "NoSpamProxyAddressSynchronization",
	# sql server instance
	[Parameter(Mandatory=$false)][string] $SqlInstance = "NoSpamProxy",
	# sql server
	[Parameter(Mandatory=$false)][string] $SqlServer = "(local)"
)

#----Functions----

# create database connection
function New-DatabaseConnection() {
	$connectionString = "Server=$SqlServer\$SqlInstance;Database=$SqlDatabase;"
	if ($SqlCredential) {
		$networkCredential = $SqlCredential.GetNetworkCredential()
		$connectionString += "uid=" + $networkCredential.UserName + ";pwd=" + $networkCredential.Password + ";"
	}
	else {
		$connectionString +="Integrated Security=True";
	}
	$connection = New-Object System.Data.SqlClient.SqlConnection
	$connection.ConnectionString = $connectionString
	
	$connection.Open()

	return $connection;
}

# run sql query
function Invoke-SqlQuery([string] $queryName, [bool] $isInlineQuery = $false, [bool] $isSingleResult) {
	try {
		$connection = New-DatabaseConnection
		$command = $connection.CreateCommand()
		if ($isInlineQuery) {
			$command.CommandText = $queryName;
		}
		else {
			$command.CommandText = (Get-Content "$PSScriptRoot\$queryName.sql")
		}
		if ($isSingleResult) {
			return $command.ExecuteScalar();
		}
		else {
			$result = $command.ExecuteReader()
			$table = new-object "System.Data.DataTable"
			$table.Load($result)
			return $table
		}
	}
	finally {
		$connection.Close();
	}
}

#----Main----
$Roles = Get-NspGatewayRole
$Thumbprints = Invoke-SqlQuery "PrivateKeyThumbprints"
Foreach ($Thumbprint in $Thumbprints)
{
    $Validationresult = "Unknown"
    [String[]]$Gateways=@()
    [String[]]$Errors=@()
    #Reset for Each Certificate
	Foreach ($Role in $Roles) 
	#Test on each gateway
	{
		$Result=Test-NspCertificate -Thumbprint $Thumbprint.Thumbprint -GatewayRole $Role.Name 
		#Check the certificate on each gateway
		Foreach ($Partialresult in $Result)
		#Split the result in root, intermediate and public or private certificate
		{
			$Subject=((($Partialresult.Subject).Split(","))[0]).Split("=")
			#Cut the address or common name out of the subject 
			If ($Partialresult.ValidationResult -notlike "NoError" -and $Thumbprint.Thumbprint -like $Partialresult.Thumbprint)
			{
			    #Outputs common name and Thumbprint
                $Gateways+=$Role.Name
                $Errors+=$Partialresult.ValidationResult
                $Validationresult = "Fail"
                #Adds Results to Cache for later integration
            }
        }
	}
    if ($Validationresult -like "Fail")
    {
     Write-Host -ForegroundColor White $Subject[1] "(" $PartialResult.Thumbprint ")  "
        For ($x=0;$x -lt $Gateways.Count;$x++)
        {           
            Write-Host -ForegroundColor Yellow $Gateways[$x]": " -NoNewline
            Write-Host -ForegroundColor Red $Errors[$x]
        }
        #Integrates Results in a list for this specific Certificate
        Write-Host " "
    }
}