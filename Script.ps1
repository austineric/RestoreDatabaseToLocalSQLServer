
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

    #ensure the SqlServer module is installed
    if (-Not (Get-Module -ListAvailable -Name "SqlServer")){
        Throw "The SqlServer module is required to run this script. Run `"Install-Module -Name SqlServer -Scope CurrentUser`" to install the module and then run the script again."
    }

    #retrieve server list from environmental variable
    $ServerList=$Env:ServerList.Split(",")

    #receive server selection
    $Server=($ServerList | Sort-Object | Out-GridView -PassThru)
    if (!$Server)
    {
        Throw "No server was selected."
    }
    
    #receive database selection
    $Database=Compare-Object -ReferenceObject (Get-SqlDatabase -ServerInstance $Server | Select-Object -Property Name -ExpandProperty Name) -DifferenceObject $DatabasesToExclude | Select-Object -Property InputObject -ExpandProperty InputObject | Out-GridView -PassThru

    #set backup directory
    #this gets the location of the last backup for the database
    #if no backup for the database exists in the directory throw an error
    $Connection1.ConnectionString="Server=$($Server); Database=msdb; Trusted_Connection=true;"
    $Command1.CommandText="
        SELECT TOP 1 f.physical_device_name
        FROM dbo.backupset s
        JOIN dbo.backupmediafamily f ON (s.media_set_id=f.media_set_id)
        WHERE s.database_name='ContinuousImprovement'
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
    Backup-SqlDatabase -ServerInstance $Server -Database $Database -BackupFile (Join-Path -Path $BackupFileDirectory -ChildPath $BackupFileName)

    #set local backup directory
    $Command2.CommandText="EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer',N'BackupDirectory';"
    $Datatable=New-Object DataTable
    $Adapter=New-Object SqlDataAdapter $Command2
    $Connection2.Open()
    $Adapter.Fill($Datatable) | Out-Null
    $LocalBackupFileDirectory=$Datatable.Rows[0].Data

    #move backup to local backup directory (move-item removes the item from the source directory once copied over)
    Move-Item -Path (Join-Path -Path $BackupFileDirectory -ChildPath $BackupFileName) -Destination (Join-Path -Path $LocalBackupFileDirectory -ChildPath $BackupFileName)

    #robocopy the backup to the new location (use instead of move-item because if I don't have permission to delete from the backup file directory robocopy still does the copy but not the delete; move-item just fails)
    #if I have permissions to delete from the backup file directory then /mov will delete the file after copying it
    robocopy "$($BackupFileDirectory)" "$($LocalBackupFileDirectory)" "$($BackupFileName)" /mov

    #drop existing local database
    $Command2.CommandText="
        IF EXISTS (SELECT * FROM sys.databases WHERE name='$($Database)')
        BEGIN
            ALTER DATABASE [$($Database)] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [$($Database)]
        END;"
    $Command2.ExecuteNonQuery() | Out-Null

    #restore backup locally
    Restore-SqlDatabase -ServerInstance . -Database $Database -BackupFile (Join-Path -Path $LocalBackupFileDirectory -ChildPath $BackupFileName) -AutoRelocateFile

}

Catch {

    Write-Host $Error[0]
    
}

Finally {

    $Connection1.Close()
    $Connection2.Close()

}