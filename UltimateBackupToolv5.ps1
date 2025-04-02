[cmdletbinding(SupportsShouldProcess=$True)]

Param([Parameter(Position=0,Mandatory=$False,
      HelpMessage="Enter the virtual machine name or names",
      ValueFromPipeline=$False,ValueFromPipelineByPropertyName=$False)]
      [ValidateNotNullorEmpty()]
       $WeekVMnames = (@(Select-Xml -Path "C:\BackupXML\VM-List-UBToolv5-Demo.xml" -XPath "/All/$hostname/VMList/Week/VM/Name" | ForEach-Object { $_.Node.InnerXML })),
       $hostname = (hostname),
       $S0 = (Select-xml -Path "C:\BackupXML\VM-List-UBToolv5-Demo.xml" -XPath "/All/S0" | ForEach-Object { $_.Node.InnerXml}),
       

      [Parameter(Position=1)]
      [ValidateNotNullorEmpty()]

      $WeeklyPath = (Select-Xml -Path "C:\BackupXML\VM-List-UBToolv5-Demo.xml" -XPath "/All/$hostname/Exportpath/Weekly/pathname" | ForEach-Object { $_.Node.InnerXML }),
     
    
      [Parameter(Position=2)]
      [string]$Weekly,
      [string]$Monthly,

      [Parameter(Position=3)]
      [switch]$AsJob
)

Begin {

 
  if ($S0=$hostname){
 
    $type1 = "Weekly"
    $retain1 = 1
  }

  else {
  Write-Verbose "Error in Hostname, So Backup failed."
  }


  Write-Verbose "Processing $type1 backups. Retaining last $retain1."

 
  $now1 = Get-Date

 
  $childPath1 = "{0}_{1}_{2:D2}_{3:D2}_{4:D2}{5:D2}" -f $type1,$now1.year,$now1.month,$now1.day,$now1.hour,$now1.minute

  $new1 = Join-Path -Path $WeeklyPath -ChildPath $childPath1

  Try {
      Write-Verbose "Creating $new1"
    
      $BackupFolder1 = New-Item -Path $new1 -ItemType directory -ErrorAction Stop 
      $LogFolderWeekly = New-Item -Path "$BackupFolder1\Log Files"  -ItemType directory -ErrorAction Stop 
  }
  Catch {
    Write-Warning "Failed to create folder $new1. $($_.exception.message)"
   
    Return
  }
} 

Process {
  

if ($BackupFolder1) {

 Out-File -FilePath "$LogFolderWeekly\Weekly Log.txt"
 Clear-Content "$LogFolderWeekly\Weekly Log.txt"
 Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date) 
 
  $exportParam1 = @{
   Path = $new1
   Name=$Null
   ErrorAction="Stop"
  }
  if ($asjob) {
    Write-Verbose "Exporting as background job"
    $exportParam1.Add("AsJob",$True)
  }

  Write-Verbose "Exporting virtual machines"

  foreach ($name1 in $WeekVMnames) {
    $exportParam1.Name=$name1
    
    if ($PSCmdlet.shouldProcess($name1)) {
       Try {
           Export-VM @exportParam1 | Wait-Job -Force
           Add-Content "$LogFolderWeekly\Weekly Log.txt" "ExportSuccessfull"
           Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)   
           Write-Host "Export Successful" -ForegroundColor Green 
           
       }
       Catch {
        Write-Warning "Failed to export virtual machine(s). $($_.Exception.Message)"
        Add-Content "$LogFolderWeekly\Weekly Log.txt" "Failed to export virtual machine(s). $($_.Exception.Message)"
        Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)  
       }
                 
      }
      Start-Sleep -Seconds 60
      $VhdSizeWS = (Get-VM -Name $name1 | Select-Object VMId | Get-VHD | Select -Property @{expression={$_.filesize/1kb -as [int]}} | Out-String)
      $size1S= (Get-ChildItem -Path "$new1\$name1\Virtual Hard Disks\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property @{expression={$_.filesize/1kb -as [int]}} | Out-String)         
      $VhdSizeW = (Get-VM -Name $name1 | Select-Object VMId | Get-VHD | Select -Property @{label='VMNames';expression ={($name1) -as [String]}}, @{label='Path';expression ={$_.Path -as [String]}},@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-String -Width 800)
      $size1= (Get-ChildItem -Path "$new1\$name1\Virtual Hard Disks\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property @{label='VMNames';expression ={($name1) -as [String]}}, @{label='Path';expression ={$_.Path -as [String]}},@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-String -Width 800)   

       
    if($VhdSizeWS -eq $size1S){
    Add-Content "$LogFolderWeekly\Weekly Log.txt" "$VhdsizeW $size1 Verified"
    Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)  
    Add-Content "$LogFolderWeekly\Weekly Log.txt" "`n `n" 
    }
    else{
    Add-Content "$LogFolderWeekly\Weekly Log.txt" "$VhdsizeW $size1 Not-Verified"
    Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)  
     Add-Content "$LogFolderWeekly\Weekly Log.txt" "`n `n"
    } 

    } 
   
  } 
  
   
} 

End {
     
    Add-Content "$LogFolderWeekly\Weekly Log.txt" "`nBackup has Finshed"
    Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)  
    Write-Host "Export has Finished" -ForegroundColor Green
    }
