d#==============================================================================
# FileTek Data Migration Toolkit 2.0
#
# Copies data from the Compellent volumes into StorHouse/RFS.
#
# © 2011 FileTek, Inc. All rights reserved.
#==============================================================================

#------------------------------------------------------------------------------
# Function declaration: Load-GlobalVariables
#------------------------------------------------------------------------------
function Load-GlobalVariables
{
    $Global:RootDir               = "C:\FileTek\Migration";
    $Global:ConfigDir             = Join-Path $RootDir "Config";
    $Global:BinDir                = Join-Path $RootDir "Bin";
    $Global:MountDir              = Join-Path $RootDir "Mounts";
    $Global:ScriptsDir            = Join-Path $RootDir "Scripts";

    $Global:RoboCopyCmd           = Join-Path $BinDir "robocopy.exe";
    $Global:HashVerifier          = Join-Path $BinDir "filehashverifier.exe";

    
    $Global:LogDir                = Join-Path $RootDir "Logs";
    $Global:FileTekLogDate        = Get-Date -Format "yyyy.MM.dd.HH.mm.ss";
    $Global:FileTekLogFileName    = "FileTek.Migration.Hash.{0}" -f $FileTekLogDate;
    $Global:FileTekLogFilePath    = Join-Path $LogDir $LogFileName;

    $Global:HostName              = "10.1.14.240";
    $Global:UserName              = "Admin";
    $Global:SCServerName          = "rfsserv";
    $Global:PhysicalName          = "cs-rfsserv-64";
    $Global:PasswordFile          = "password.txt";
    $Global:PasswordFullPath      = Join-Path $RootDir $PasswordFile;    
    
    $Global:TargetRootDir         = "V:\Repository";    

    $Global:DiskRescanDelay       = 20;
    $Global:DiskRescanAttempts    = 5;
    $Global:SleepSetting          = 10;
    
    $Global:FatalError            = $False;
    $Global:MenuResponse          = $Null;
    
    Write-Host ("Global variables loaded.") -ForegroundColor White;
}

#------------------------------------------------------------------------------
# Function declaration: Get-Password
#------------------------------------------------------------------------------
function Get-Password
{
    try
    {
        $PasswordFile = Test-Path $PasswordFullPath
        if (!$PasswordFile)
        {
            Write-Host "`nError: Password file not found."  -ForegroundColor Red;            
            Read-Host -Prompt "Enter new password" -AsSecureString | ConvertFrom-SecureString | Out-File $PasswordFullPath;      
            $Password = Get-Content password.txt | ConvertTo-SecureString;      
            return $Password;
        } 
        else
        {
            $Password = Get-Content password.txt | ConvertTo-SecureString;            
            return $Password;
        }
        
    }    
    catch
    {
        Write-Host "ERROR in Get-Password function: " $_ -ForegroundColor Red;
        Exit;
    }
}

#------------------------------------------------------------------------------
# Function declaration: ConnectToCompellent
#------------------------------------------------------------------------------
function Get-Connection($Password)
{    
    Write-Host ("`nConnecting to Storage Center {0} as {1} ... " -f $HostName, $UserName) -ForegroundColor White -NoNewLine;
    
    try
    {
        $Connection = Get-SCConnection -HostName $HostName -User $UserName -Password $Password;
        if ($Connection.AccessLevel -eq "Admin")
        {
            Write-Host "OK!" -ForegroundColor Green;
            $myShell.WindowTitle =  ("FileTek Compellent RFS Migration - Connected as User: {0} - Host: {1}" -f $Connection.User, $Connection.Host);
            return $Connection;
        }
    }
    
    catch
    {
        Write-Host "Error in Get-Connection function: " $_ -ForegroundColor Red;
        Exit;
    }           
}

#------------------------------------------------------------------------------
# Function declaration: Show-Menu
#------------------------------------------------------------------------------
function Show-Menu
{
    $Exit = New-Object System.Management.Automation.Host.ChoiceDescription "E&xit", "Used to exit script.";
    $NewVolumeFile = New-Object System.Management.Automation.Host.ChoiceDescription "Create &new volume file only.", "Only creates volume file, does not start migration.";
    $PreviousVolumeFile = New-Object System.Management.Automation.Host.ChoiceDescription "Migrate &previous volume files.", "Start migration with previous files in COnfig directory.";
    
    $Options = [System.Management.Automation.Host.ChoiceDescription[]] ($Exit, $NewVolumeFile, $PreviousVolumeFile);  
    $Response = $host.ui.PromptForChoice("Migration", "Enter your selection", $Options, 0);
    
    return $Response;
}

#------------------------------------------------------------------------------
# Function declaration: Create-VolumeFile
#------------------------------------------------------------------------------
function Create-VolumeFile ($RootVolumeFolder, $SubVolumeFolder)
{    
    try
    {
    if($SubVolumeFolder)
        {
            Write-Host ("`nCreating volume file {0}.csv ... " -f $SubVolumeFolder) -NoNewline -ForegroundColor White;
            $VolumeData = Get-SCVolume -Connection $Connection | where { $_.LogicalPath -match '^' + $RootVolumeFolder + '\\' + $SubVolumeFolder + '\\' } | select Name,SerialNumber,LogicalPath,Index;
            if(!$VolumeData)
            {
                Write-Host "FAILED!" -ForegroundColor Red;
                Write-Host "No volumes found. Check if '$SubVolumeFolder' is spelled correctly or is available in Compellent Manager.`n" -ForegroundColor White;
            }
            else
            {
                $NewVolumeFile = $SubVolumeFolder + ".csv";
                $NewVolumeFile = Join-Path $ConfigDir $NewVolumeFile;                
                $VolumeData | Export-Csv $NewVolumeFile;
                Write-Host "OK!`n" -ForegroundColor Green;       
                return $NewVolumeFile;
            }
        }
    }    
    catch
    {
        Write-Host "ERROR in Create-VolumeFile: " $_ -ForegroundColor Red;
        Continue;
    }
}

#------------------------------------------------------------------------------
# Function declaration: Map-Volume
#------------------------------------------------------------------------------
function Map-Volume ($Volume, $Connection)
{
    try
    {
        $SCServerObject = Get-SCServer -Connection $Connection | where { $_.Name -eq $SCServerName };
        $SelectedVolume = Get-SCVolume -Connection $Connection -Name $Volume.Name -Index $Volume.Index;        
        
        $VolumeMap = Get-SCVolumeMap -Connection $Connection | where { $_.VolumeName -eq $Volume.Name -and $_.ServerName -eq $SCServerName  -and $_.LogicalPath -eq $Volume.LogicalPath};
        if($VolumeMap)
        {
            Write-Host ("Using existing SCVolumeMap for Server: {0} Volume: {1}" -f $SCServerObject.Name, $SelectedVolume.Name) -ForegroundColor White;
            return $VolumeMap;
        }
        else
        {
            Write-Host ("Mapping volume ... " -f $SCServerObject.Name, $SelectedVolume.Name) -NoNewline -ForegroundColor White;
            $VolumeMap = New-SCVolumeMap -Connection $Connection -SCServer $SCServerObject -SCVolume $SelectedVolume -ReadOnly;
            Write-Host "OK!" -ForegroundColor Green;
            return $VolumeMap;
        }
    }    
    catch
    {
        Write-Host "ERROR in Map-Volume: " $_ -ForegroundColor Red;
        Exit;
    }
}

#------------------------------------------------------------------------------
# Function declaration: Scan-Disk
#------------------------------------------------------------------------------
function Scan-Disk ($SelectedVolume, $Action)
{
    try
    {
        Write-Host ("Rescanning Disk Devices ... " -f $SelectedVolume.Name) -ForegroundColor White -NoNewLine;
        $Count = 0;
        do
        {
            Rescan-DiskDevice -Server $PhysicalServerName -RescanDelay $DiskRescanDelay;
            $Disk = Get-DiskDevice -SerialNumber $SelectedVolume.SerialNumber -WarningAction SilentlyContinue;
            $Count++;
        }       
        until (($Disk -ne $null) -or ($Count -eq $DiskRescanAttempts))
        
        if(($Disk -eq $null) -and ($Action -eq "Add")) 
        {
            throw "Unable to find the disk on server";
        }
        else
        {  
            Set-DiskDevice -DeviceName $Disk.devicename -Online
            Write-Host "OK!" -ForegroundColor Green;
            return $Disk;
        }
    }    
    catch
    {
        Write-Host "ERROR in Scan-Disk: " $_ -ForegroundColor Red;   
        Continue;
    }
}

#------------------------------------------------------------------------------
# Function declaration: Create-MountDir
#------------------------------------------------------------------------------
function Create-MountDir($AccessPath)
{
    try
    {
        if (!(Test-Path $AccessPath -PathType Container))
        {
            Write-Host ("Creating mount folder {0} ... " -f $AccessPath) -ForegroundColor White -NoNewLine;
            $null = New-Item -Path $AccessPath -ItemType "Directory";
            Write-Host "OK!" -ForegroundColor Green;
            return $True
        }
    }
    
    catch
    {
        Write-Host "ERROR in Create-MountDir: " $_ -ForegroundColor Red;
        Continue;
    }
}

#------------------------------------------------------------------------------
# Function declaration: Update-Log
#------------------------------------------------------------------------------
function Update-Log($LogPath, $Message)
{
    try
    {
        $LogStream = New-Object System.IO.StreamWriter($LogPath, $True);
        $LogStream.WriteLine($Message);
    }
    
    catch
    {
        Write-Host "ERROR in Update-Log: " $_ -ForegroundColor Red;
        Continue;
    }
    
    finally
    {
        $LogStream.Close();
    }
}

function Hash-Volume($SourceFolder, $DestinationFolder, $HashType, $SourceVolumeName)
{
        $HashLogFileName = "Hash - {0} - {1} - {2}" -f $SourceVolumeName.ToString().Replace("\"," - ").ToUpper(), $LogDate,  $HashType;
        $HashLogFileName = Join-Path $LogDir  $HashLogFileName
        $CmdArgs = (" -src ""{0}""  -dest ""{1}"" -l ""{2}.xlsx"" -lf XLSX -hash {3} -ex ""System Volume Information"" ""Recycler"" ""Recycled"" -sv ""{4}"" " -f $SourceFolder,  $DestinationFolder, $HashLogFileName, $HashType, $SourceVolumeName);      
        write-host "Hash CmdArgs: " $CmdArgs  -ForegroundColor Cyan;
        #Start-Process -NoNewWindow -Wait -FilePath $FileHashVerifierCmd -ArgumentList $CmdArgs ;
        $procReturnValue = run-myprocess $HashVerifier $CmdArgs;
        #Check last command is succeeded.
        #Write-Host("Comparison Return Value: " + $procReturnValue) -ForegroundColor Blue;
        return $procReturnValue;
}


function run-myprocess ($cmd, $params) {
    $p = new-object System.Diagnostics.Process;
    # $p.StartInfo = new-object System.Diagnostics.ProcessStartInfo;
    $exitcode = -100      ;
    $p.StartInfo.FileName = $cmd;
    $p.StartInfo.Arguments = $params;
    $p.StartInfo.UseShellExecute = $shell;
    $p.StartInfo.WindowStyle = 1; #hidden.  Comment out this line to show output in separate console
    $null = $p.Start();
    $p.WaitForExit();
    $exitcode = $p.ExitCode;
    #$p.Dispose();
    return $exitcode;
}



#------------------------------------------------------------------------------
# Function declaration: Copy-Volume
#------------------------------------------------------------------------------
function Copy-Volume ($AccessPathString, $Volume, $LogDate)
{   
    Write-Host ("`nCopying volume {0}, please wait ..." -f $Volume.Name) -ForegroundColor White;
    try
    { 
        Start-Sleep -s $SleepSetting;
        Foreach ($Directory in [System.IO.Directory]::GetDirectories($AccessPathString))
        {    
            Write-Host ("{0}" -f $Directory) -ForegroundColor Gray;
            $Type = "COPY";                
            $RobocopyLogFileName = "{0} - {1} - {2}.log" -f $Volume.LogicalPath.ToString().Replace("\"," - ").ToUpper(), $LogDate, $Type;
            $RobocopyLogFilePath = Join-Path $LogDir $RobocopyLogFileName
            $Destination = Join-Path $TargetRootDir $Volume.LogicalPath;       
            
            $ExcludeDirectoriesRoot = Join-Path $MountDir $Volume.Name;
            $SystemVolume = "System Volume Information";
            $ExcludeSystemVolume = Join-Path $ExcludeDirectoriesRoot $SystemVolume;
            $Recycled = "Recycled";
            $ExcludeRecycled = Join-Path $ExcludeDirectoriesRoot $Recycled;
            $Recycler = "Recycler";
            $ExcludeRecycler = Join-Path $ExcludeDirectoriesRoot $Recycler; 
            $ExcludeDirectories = $ExcludeSystemVolume + " " + $ExcludeRecycled + " " + $ExcludeRecycler;
            
            $CmdArgs = '"{0}" "{1}" /COPY:DAT /E /V /NP /ZB /LOG:"{2}" /R:1 /W:2 /TEE /XD "{3}" "{4}" "{5}"' -f $AccessPathString, $Destination, $RobocopyLogFilePath, $ExcludeSystemVolume, $ExcludeRecycled, $ExcludeRecycler;  
            $RoboCopyResponse = Start-Process -NoNewWindow -Wait -FilePath $RoboCopyCmd -ArgumentList $CmdArgs;              
        }
        Start-Sleep -s $SleepSetting;
    }
    
    catch
    {
        Write-Host "ERROR in Copy-Volume: " $_ -ForegroundColor Red;
        $FatalError=$True;
        Exit;
    }
}


#------------------------------------------------------------------------------
# Function declaration: Start-Migration
#------------------------------------------------------------------------------
function Start-Migration ($VolumeFile)
{
    try
    {
        if(!$VolumeFile)
        {
            $RootVolumeFolder = Read-Host -Prompt "`nEnter the name of the top-most folder of the volume you want to migrate";
            $SubVolumeFolder = Read-Host -Prompt "Enter the name of the volume folder you want to migrate";
            $VolumeFile = Create-VolumeFile $RootVolumeFolder $SubVolumeFolder;
        }
        
        $VolumeData = Import-Csv $VolumeFile;
        foreach($Volume in $VolumeData)
        {
            $VolumeMap = Map-Volume $Volume $Connection;
            $Disk = Scan-Disk $Volume "Add";           
            if($Disk)
            {
                $AccessPathString = Join-Path $MountDir $Volume.Name;
                $MountDirCreated = $False;
                $MountDirCreated = Create-MountDir $AccessPathString;
                
                Write-Host ("Adding access path {0} ... " -f $AccessPathString) -ForegroundColor White -NoNewLine;
                $Temp = Add-VolumeAccessPath -DiskSerialNumber $Volume.SerialNumber -AccessPath $AccessPathString -Confirm:$False;
                Write-Host "OK!" -ForegroundColor Green;
                
                #-----------------------------------------------------------------------------------------------------------  
                $LogDate = Get-Date -Format "yyyyMMddHHmm"; 
                                
                $Temp = Copy-Volume $AccessPathString $Volume $LogDate;
                $DestinationDir = Join-Path $TargetRootDir $Volume.LogicalPath
                $HashResponse = Hash-Volume $AccessPathString $DestinationDir "MD5andSHA1" $Volume.LogicalPath              
                #-----------------------------------------------------------------------------------------------------------
                
                Write-Host ("Removing access path {0} ... " -f $accessPathString) -ForegroundColor White -NoNewLine;
                $Temp = Remove-VolumeAccessPath -AccessPath $AccessPathString -Confirm:$False;
                Write-Host "OK!" -ForegroundColor Green;
                
                if($MountDirCreated)
                {
                    Write-Host ("Removing mount folder {0} ... " -f $AccessPathString) -ForegroundColor White -NoNewLine;
                    Remove-Item -Path $AccessPathString;
                    Write-Host "OK!" -ForegroundColor Green;
                }
            }
            
            If($VolumeMap -ne $null)
            {
                Write-Host ("Removing disk from server: {0} ... " -f $PhysicalName) -ForegroundColor White -NoNewLine;
                $Temp = Remove-SCVolumeMap -SCVolumeMap $VolumeMap -Connection $Connection -Confirm:$False;
                Write-Host "OK!" -ForegroundColor Green;
            }
            
            $Disk = Scan-Disk $Volume "Remove";                        
            if($HashResponse -eq 0)
            {
                Write-Host ("Migration for volume: {0} complete.`n" -f $Volume.LogicalPath) -ForegroundColor Green;           
            }
            else
            {
                Write-Host ("Migration for volume: {0} failed.`n" -f $Volume.LogicalPath) -ForegroundColor Red;
                return $False;
                Exit;
            }        
        }
        return $True;     
    }
    

    catch
    {
        Write-Host "ERROR in Start-Migration: " $_ -ForegroundColor Red;
        return $False;
        Exit;
    }
}

#==============================================================================
# Script Entry Point
#==============================================================================

# Update the window title
Clear-Host;
$myShell = (Get-Host).UI.RawUI;
$myShell.WindowTitle = "FileTek Volume Migration for FTC 2.0";
Write-Host "";

Load-GlobalVariables;
$Password = Get-Password;
$Connection = Get-Connection $Password;

do
{
    $MenuResponse = Show-Menu;
    
    if ($MenuResponse -eq 1)
    {
        $RootVolumeFolder = Read-Host -Prompt "`nEnter the name of the top-most folder of the volume you want to migrate";
        $SubVolumeFolder = Read-Host -Prompt "Enter the name of the volume folder you want to migrate";
        $VolumeFile = Create-VolumeFile $RootVolumeFolder $SubVolumeFolder;               
    }
    elseif ($MenuResponse -eq 2)
    {
        $Files = [System.IO.Directory]::GetFiles($ConfigDir)
        foreach ($File in $Files)
        {
            $Completed = $False;
            $FileExtension = [System.IO.Path]::GetExtension($File);
            if ($FileExtension -eq ".csv")
            {
                Write-Host $File -ForegroundColor Yellow;
                $Completed = Start-Migration $File;            
                if ($Completed)
                {
                    $NewFile = [System.IO.Path]::ChangeExtension($File, ".done");
                    [System.IO.File]::Move($File, $NewFile); 
                }
            }    
        }
    }
}
until (($FatalError) -or ($MenuResponse -eq 0))

#------------------------------------------------------------------------------
# Last line in script
#------------------------------------------------------------------------------
Write-Host "";