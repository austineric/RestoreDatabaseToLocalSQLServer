
####################################
# Author:      Eric Austin
# Create date: January 2020
# Description: Presents a list of SQL Servers (from $Env:ServerList) and databases to back up and restore to local SQL Server
####################################

using namespace System.Data
using namespace System.Data.SqlClient

#set common variables
$ErrorActionPreference="Stop"

#set script-specific variables
$ServerList=@()
$Server=""
$Database=""
$DatabasesToExclude=@("master", "model", "msdb", "tempdb", "dbadmin")
$BackupFileName=""
$BackupFileDirectory=""
$LocalBackupFileDirectory=""

#SQL Server parameters
$Connection1=New-Object SqlConnection
$Command1=$Connection1.CreateCommand()
$Connection2=New-Object SqlConnection
$Connection2.ConnectionString="Server=.;Database=master;Trusted_Connection=true;"
$Command2=$Connection2.CreateCommand()

Try {

    Clear-Host
    Write-Host ""

    #ensure the SqlServer module is installed
    Write-Host "Checking for SqlServer module..."
    if (-Not (Get-Module -ListAvailable -Name "SqlServer")){
        Throw "The SqlServer module is required to run this script. Run `"Install-Module -Name SqlServer -Scope CurrentUser`" to install the module and then run the script again."
    }
    Write-Host "SqlServer module is installed"

    #retrieve server list from environmental variable
    Write-Host "Retrieving SQL Server list from environmental variable..."
    if ([string]::IsNullOrWhiteSpace($Env:SQLServerList))
    {
        Throw "The environmental variable `"SqlServerList`" either does not exist or is empty, exiting now."
    }
    else
    {
        $ServerList=$Env:SQLServerList.Split(",")
    }

    #receive server selection
    $Server=($ServerList | Sort-Object | Out-GridView -PassThru)
    if ([string]::IsNullOrWhiteSpace($Server))
    {
        Throw "No server was selected."
    }
    
    #receive database selection
    Write-Host "Retrieving database list for $($Server)..."
    $Database=Compare-Object -ReferenceObject (Get-SqlDatabase -ServerInstance $Server | Select-Object -Property Name -ExpandProperty Name) -DifferenceObject $DatabasesToExclude | Select-Object -Property InputObject -ExpandProperty InputObject | Out-GridView -PassThru
    if ([string]::IsNullOrWhiteSpace($Database))
    {
        Throw "No database was selected."
    }

    #set backup directory
    #this gets the location of the last backup for the database
    #if no backup for the database exists in the directory throw an error
    Write-Host "Looking up backup location..."
    $Connection1.ConnectionString="Server=$($Server); Database=msdb; Trusted_Connection=true;"
    $Command1.CommandText="
        SELECT TOP 1 f.physical_device_name
        FROM dbo.backupset s
        JOIN dbo.backupmediafamily f ON (s.media_set_id=f.media_set_id)
        WHERE s.database_name='$($Database)'
        ORDER BY s.backup_finish_date DESC;"
    $Connection1.Open()
    $BackupFileDirectory=([System.IO.DirectoryInfo]$Command1.ExecuteScalar()).Parent.FullName
    if (!$BackupFileDirectory)
    {
        Throw "No backup exists to retrieve the backup directory path from"
    }

    #set backup file name
    $BackupFileName="$($Database)_$(Get-Date -Format "MMddyyyy_hhmmss")_$((New-Guid).Guid).bak"

    #create new backup
    Write-Host "Creating new backup..."
    Backup-SqlDatabase -ServerInstance $Server -Database $Database -BackupFile (Join-Path -Path $BackupFileDirectory -ChildPath $BackupFileName)

    #set local backup directory
    $Command2.CommandText="EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer',N'BackupDirectory';"
    $Datatable=New-Object DataTable
    $Adapter=New-Object SqlDataAdapter $Command2
    $Connection2.Open()
    $Adapter.Fill($Datatable) | Out-Null
    $LocalBackupFileDirectory=$Datatable.Rows[0].Data

    #copy backup to local directory
    Write-Host "Copying backup to local backup location..."
    Copy-Item -Path (Join-Path -Path $BackupFileDirectory -ChildPath $BackupFileName) -Destination (Join-Path -Path $LocalBackupFileDirectory -ChildPath $BackupFileName)

    #delete original backup (comment out if no permission to delete)
    #Write-Host "Deleting original backup..."
    #Remove-Item -Path (Join-Path -Path $BackupFileDirectory -ChildPath $BackupFileName)

    #drop existing local database
    Write-Host "Dropping local database if it exists..."
    $Command2.CommandText="
        IF EXISTS (SELECT * FROM sys.databases WHERE name='$($Database)')
        BEGIN
            ALTER DATABASE [$($Database)] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [$($Database)]
        END;"
    $Command2.ExecuteNonQuery() | Out-Null

    #restore backup locally
    Write-Host "Restoring backup locally..."
    Restore-SqlDatabase -ServerInstance . -Database $Database -BackupFile (Join-Path -Path $LocalBackupFileDirectory -ChildPath $BackupFileName) -AutoRelocateFile

    Write-Host "Success."

}

Catch {

    Write-Host $Error[0]
    
}

Finally {

    $Connection1.Close()
    $Connection2.Close()

}