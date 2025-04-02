[cmdletbinding(SupportsShouldProcess=$True)]

Param([Parameter(Position=0,Mandatory=$False,
      HelpMessage="Enter the virtual machine name or names",
      ValueFromPipeline=$False,ValueFromPipelineByPropertyName=$False)]
      [ValidateNotNullorEmpty()]
       $WeekVMnames = (@(Select-Xml -Path "E:\Important Scripts\Export_VMs\VM-List-2022.xml" -XPath "/All/$hostname/VMList/Week/VM/Name" | ForEach-Object { $_.Node.InnerXML })),
       $hostname = (hostname),
       $S0 = (@(Select-xml -Path "E:\Important Scripts\Export_VMs\VM-List-2022.xml" -XPath "/All/S0" | ForEach-Object { $_.Node.InnerXml})),
       

      [Parameter(Position=1)]
      [ValidateNotNullorEmpty()]

      $WeeklyPath = (Select-Xml -Path "E:\Important Scripts\Export_VMs\VM-List-2022.xml" -XPath "/All/$hostname/Exportpath/Weekly/pathname" | ForEach-Object { $_.Node.InnerXML }),
      
    
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
 #Clear-Variable -Name subFolders1 -ErrorAction Ignore
Clear-Variable -Name oldest2 -ErrorAction Ignore
 $subFolders1 =  dir -Path $WeeklyPath\$type1* -Directory -ErrorAction Stop
   if ($subFolders1) {
    
      if ($subFolders1.count -ge $retain1) {

         
         $oldest2 = $subFolders1 | sort CreationTime | Select -First 1 
         Write-Verbose "Found OldFolder $($oldest2.FullName)" 

             }
       
    else {
    
      Write-Verbose "No matching folders found. Creating the first folder"   

  }
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
  Add-Content "$LogFolderWeekly\Weekly Log.txt" "Exporting virtual machines"
  Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)

  foreach ($name1 in $WeekVMnames) {
    $exportParam1.Name=$name1
    
    if ($PSCmdlet.shouldProcess($name1)) {
    try{
    
           Export-VM @exportParam1 | Wait-Job -Force
           Add-Content "$LogFolderWeekly\Weekly Log.txt" "ExportSuccessfull"
           Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)   
           Write-Host "Export Successful" -ForegroundColor Green 
           
       }
       Catch{
        Write-Warning "Failed to export virtual machine(s). $($_.Exception.Message)"
        Add-Content "$LogFolderWeekly\Weekly Log.txt" "Failed to export virtual machine(s). $($_.Exception.Message)" 
        Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)
        
       } 
       }      
      Start-Sleep -Seconds 60
      $VhdSizeWS = (Get-VM -Name $name1 | Select-Object VMId | Get-VHD | Select -Property @{expression={$_.filesize/1kb -as [int]}} | Out-String)
      $size1S= (Get-ChildItem -Path "$new1\$name1\Virtual Hard Disks\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property @{expression={$_.filesize/1kb -as [int]}} -ErrorAction Ignore | Out-String )   
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
      Try {
   Write-Verbose "Checking $WeeklyPath for subfolders"
   Add-Content "$LogFolderWeekly\Weekly Log.txt" "Checking $WeeklyPath for subfolders"
   
  }
  Catch {
      Write-Warning "Failed to enumerate folders from $WeeklyPath"
      Add-Content "$LogFolderWeekly\Weekly Log.txt" "Failed to enumerate folders from $WeeklyPath" 
   
      return
  } 
     
    if($oldest2.FullName){
    $size1old= (Get-ChildItem -Path "$($oldest2.fullname)\$name1\Virtual Hard Disks\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property @{expression={$_.filesize/1kb -as [int]}} -ErrorAction Ignore | Out-String )      
     
   if ($subFolders1.Count -ge $retain1) {
   
         Write-Verbose "Found $($subfolders1.count) folder(s)"
         Add-Content "$LogFolderWeekly\Weekly Log.txt" "Found $($subfolders1.count) folder(s)"
         Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date) 
         Write-Verbose "Found Oldest Backup folder -: $($oldest2.fullname)"
         Add-Content "$LogFolderWeekly\Weekly Log.txt" "Found Oldest Backup folder -: $($oldest2.fullname)"
         Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date) 

   if($oldest2.FullName -ne $BackupFolder1){

         $oldestVM = "$oldest2\$name1"
         $oldestlog = "$oldest2\Log Files"
         Add-Content "$LogFolderWeekly\Weekly Log.txt" "Deleting Oldest Backed up VM -: $oldest2\$name1" 
         Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)
         $oldestVM | Remove-Item -Recurse -Force | Wait-Job -ErrorAction Ignore
         $oldestlog | Remove-Item -Recurse -Force | Wait-Job -ErrorAction Ignore
         Add-Content "$LogFolderWeekly\Weekly Log.txt" "Deleted Oldest Backed up VM -: $oldest2\$name1" 
         Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)
         #Start-Sleep -Seconds 30
         }
         else{
         Add-Content "$LogFolderWeekly\Weekly Log.txt" "Failed to Delete Oldest Backed up VM -: $oldest2\$name1"
      }
  
  }
  else{
         Write-Verbose "No matching folders found. Creating the first folder"   
         Add-Content "$LogFolderWeekly\Weekly Log.txt" "No matching folders found. Creating the first folder"
         Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date) 
         } 
         }

      } #for loop week VM names ends.
      #Start-Sleep -Seconds 30
      if((Get-ChildItem $oldest2.FullName | Measure-Object).Count -eq 0){
      
      
      $oldest2 | Remove-Item -Recurse -Force | Wait-Job
      Add-Content "$LogFolderWeekly\Weekly Log.txt" "Deleting oldest folder $($oldest2.fullname)"
      Write-Host "Deleting Oldest Backup Folder $($oldest2.fullname)"
      }
 
    else {
    Add-Content "$LogFolderWeekly\Weekly Log.txt" "Fail to delete and oldest folder $($oldest1.fullname) is not empty or non-exist folder"
    Write-Host "Fail to delete and oldest folder $($oldest1.fullname) is not empty or non-exist folder"
    }
    
    }
    }


End {
     
    Add-Content "$LogFolderWeekly\Weekly Log.txt" "`nBackup has Finshed"
    Add-Content "$LogFolderWeekly\Weekly Log.txt" -Value(Get-Date)  
    Write-Host "Export has Finished" -ForegroundColor Green
    }
