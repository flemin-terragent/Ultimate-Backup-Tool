[cmdletbinding(SupportsShouldProcess=$True)]

Param([Parameter(Position=0,Mandatory=$False,
      HelpMessage="Enter the virtual machine name or names",
      ValueFromPipeline=$False,ValueFromPipelineByPropertyName=$False)]
      [ValidateNotNullorEmpty()]
       #[Alias("name")]
       $WeekVMnames = (@(Select-Xml -Path "D:\FleminScripts(Don'tDelete)\Export_VMs\Example_VM_Details.xml" -XPath "/All/$hostname/VMList/Week/VM/Name" | ForEach-Object { $_.Node.InnerXML })),
       #$DynamicPath="D:\FleminScripts(Don'tDelete)",
       #$WeekVMnames = @(Select-Xml -Path "$DynamicPath\Export_VMs\Example_VM_Details.xml" -XPath "/All/$hostname/VMList/Week/VM/Name" | ForEach-Object { $_.Node.InnerXML }),
       $MonthVMnames = (@(Select-Xml -Path "D:\FleminScripts(Don'tDelete)\Export_VMs\Example_VM_Details.xml" -XPath "/All/$hostname/VMList/Month/VM/Name" | ForEach-Object { $_.Node.InnerXML })),
       $hostname = (hostname),
       $S0 = (Select-xml -Path "D:\FleminScripts(Don'tDelete)\Export_VMs\Example_VM_Details.xml" -XPath "/All/S0" | ForEach-Object { $_.Node.InnerXml}),
       $fileW = (@(Select-Xml -Path "D:\FleminScripts(Don'tDelete)\Export_VMs\Example_VM_Details.xml" -XPath "/All/$hostname/Exportpath/Weekly/pathname" | ForEach-Object { $_.Node.InnerXML })),
       #$file1 = ("$fileW\$BackupFolder1\$WeekVMnames\Virtual Hard Disks"),
       $fileM = (@(Select-Xml -Path "D:\FleminScripts(Don'tDelete)\Export_VMs\Example_VM_Details.xml" -XPath "/All/$hostname/Exportpath/Monthly/pathname" | ForEach-Object { $_.Node.InnerXML })),
       #$file2 = ("$fileM\$BackupFolder2\$MonthVMnames\Virtual Hard Disks"),

      [Parameter(Position=1)]
      [ValidateNotNullorEmpty()]
      #$Path = "D:\FleminScripts(Don'tDelete)\Sample_Exports",
      #$DynamicPath = "D:\FleminScripts(Don'tDelete)",

      $MonthlyPath = (Select-Xml -Path "D:\FleminScripts(Don'tDelete)\Export_VMs\Example_VM_Details.xml" -XPath "/All/$hostname/Exportpath/Monthly/pathname" | ForEach-Object { $_.Node.InnerXML }),
      $WeeklyPath = (Select-Xml -Path "D:\FleminScripts(Don'tDelete)\Export_VMs\Example_VM_Details.xml" -XPath "/All/$hostname/Exportpath/Weekly/pathname" | ForEach-Object { $_.Node.InnerXML }),
      $VhdSizeM = (Get-VM -Name $MonthVMnames | Select-Object VMId | Get-VHD | Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-String ),
      $VhdSizeW = (Get-VM -Name $WeekVMnames | Select-Object VMId | Get-VHD | Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-String ),
      $size1= (Get-ChildItem -Path "D:\FleminScripts(Don'tDelete)\Sample_Exports\Week\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-String),
      $size2= (Get-ChildItem -Path "D:\FleminScripts(Don'tDelete)\Sample_Exports\Month\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-String),
      $VhdSizeMS = (Get-VM -Name $MonthVMnames | Select-Object VMId | Get-VHD | Select -Property @{expression={$_.filesize/1mb -as [int]}} | Out-String ),
      $VhdSizeWS = (Get-VM -Name $WeekVMnames | Select-Object VMId | Get-VHD | Select -Property @{expression={$_.filesize/1mb -as [int]}} | Out-String ),
      $size1S= (Get-ChildItem -Path "D:\FleminScripts(Don'tDelete)\Sample_Exports\Week\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property @{expression={$_.filesize/1mb -as [int]}} | Out-String),
      $size2S= (Get-ChildItem -Path "D:\FleminScripts(Don'tDelete)\Sample_Exports\Month\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property @{expression={$_.filesize/1mb -as [int]}} | Out-String),

    
      [Parameter(Position=2)]
      [string]$Weekly,
      [string]$Monthly,

      [Parameter(Position=3)]
      [switch]$AsJob
)


Function okButton1_Click(){
Begin {

  if ($S0=$hostname){
    $type1 = "Weekly"
    $retain1 = 1
    }
  else {
  Write-Verbose "Error in Hostname, So Backup failed."
  }
  Write-Verbose "Processing $type1 backups. Retaining last $retain1."

  Try {
   Write-Verbose "Checking $WeeklyPath for subfolders"

   $subFolders1 =  dir -Path $WeeklyPath\$type1* -Directory -ErrorAction Stop
  }
  Catch {
      Write-Warning "Failed to enumerate folders from $WeeklyPath"
      #bail out of the script
      return
  }

  #check if any backup folders
  if ($subFolders1) {
      #if found, get count
      Write-Verbose "Found $($subfolders1.count) folder(s)"
      
      #if more than the value of $retain, delete oldest one
      if ($subFolders1.count -ge $retain1 ) {
         #get oldest folder based on its CreationTime property
         $oldest1 = $subFolders1 | sort CreationTime | Select -first 1 
         Write-Verbose "Deleting oldest folder $($oldest1.fullname)"
         #delete it
         $oldest1 | Remove-Item -Recurse -Force
      }
        
   } #if $subfolders
  else {
      #if none found, create first one
      Write-Verbose "No matching folders found. Creating the first folder"    
  }

  #create the folder
  #get the current date
  $now1 = Get-Date

  #name format is Type_Year_Month_Day_HourMinute
  $childPath1 = "{0}_{1}_{2:D2}_{3:D2}_{4:D2}{5:D2}" -f $type1,$now1.year,$now1.month,$now1.day,$now1.hour,$now1.minute

  #create a variable that represents the new folder path
  $new1 = Join-Path -Path $WeeklyPath -ChildPath $childPath1

  Try {
      Write-Verbose "Creating $new1"
      #Create the new backup folder
      $BackupFolder1 = New-Item -Path $new1 -ItemType directory -ErrorAction Stop 
  }
  Catch {
    Write-Warning "Failed to create folder $new1. $($_.exception.message)"
    #failed to create folder so bail out of the script
    Return
  }
} #end begin

Process {

#only process if a backup folder was created
if ($BackupFolder1) {
  #export VMs
  #define a hashtable of parameters to splat to Export-VM
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
  <#
   Go through each virtual machine name, and export it using Export-VM
  #>
  foreach ($name1 in $WeekVMnames) {

    $exportParam1.Name=$name1
    #if the user did not include -WhatIf then the machine will be exported
    #otherwise they will get a WhatIf message
    if ($PSCmdlet.shouldProcess($name1)) {
       Try {
            Export-VM @exportParam1 | Out-GridView -Wait
            
            if($VhdSizeWS -eq $size1S){

    Write-Host "Export script Finished." -ForegroundColor Green
}
else{
Write-Host "Export Failed." -ForegroundColor Red
}
       }
       Catch {
        Write-Warning "Failed to export virtual machine(s). $($_.Exception.Message)"
       }
    }
    } #close foreach
  } #if backup folder exists
} #Process

End {
    Write-Host "Export script Successfull." -ForegroundColor Green

}

}


Function okButton2_Click(){
Begin {

  #define some variables if we are doing weekly or monthly backups
  if ($S0=$hostname){
  #if ($Monthly) {
    $type2 = "Monthly"
    $retain2 = 1
  }
 # else {
    # $type = "Weekly"
     #$retain = 1
  #}
  #}
  else {
  Write-Verbose "Error in Hostname, So Backup failed."
  }
 # }

  Write-Verbose "Processing $type2 backups. Retaining last $retain2."

  #get backup directory list
  Try {
   Write-Verbose "Checking $MonthlyPath for subfolders"
   
   #get only directories under the path that start with Weekly or Monthly
   $subFolders2 =  dir -Path $MonthlyPath\$type2* -Directory -ErrorAction Stop
  }
  Catch {
      Write-Warning "Failed to enumerate folders from $MonthlyPath"
      #bail out of the script
      return
  }

  #check if any backup folders
  if ($subFolders2) {
      #if found, get count
      Write-Verbose "Found $($subfolders2.count) folder(s)"
      
      #if more than the value of $retain, delete oldest one
      if ($subFolders2.count -ge $retain2 ) {
         #get oldest folder based on its CreationTime property
         $oldest2 = $subFolders2 | sort CreationTime | Select -first 1 
         Write-Verbose "Deleting oldest folder $($oldest2.fullname)"
         #delete it
         $oldest2 | Remove-Item -Recurse -Force
      }
        
   } #if $subfolders
  else {
      #if none found, create first one
      Write-Verbose "No matching folders found. Creating the first folder"    
  }

  #create the folder
  #get the current date
  $now2 = Get-Date

  #name format is Type_Year_Month_Day_HourMinute
  $childPath2 = "{0}_{1}_{2:D2}_{3:D2}_{4:D2}{5:D2}" -f $type2,$now2.year,$now2.month,$now2.day,$now2.hour,$now2.minute

  #create a variable that represents the new folder path
  $new2 = Join-Path -Path $MonthlyPath -ChildPath $childPath2

  Try {
      Write-Verbose "Creating $new2"
      #Create the new backup folder
      $BackupFolder2 = New-Item -Path $new2 -ItemType directory -ErrorAction Stop 
  }
  Catch {
    Write-Warning "Failed to create folder $new2. $($_.exception.message)"
    #failed to create folder so bail out of the script
    Return
  }
} #end begin

Process {

#only process if a backup folder was created
if ($BackupFolder2) {
  #export VMs
  #define a hashtable of parameters to splat to Export-VM
  $exportParam2 = @{
   Path = $new2
   Name=$Null
   ErrorAction="Stop"
  }
  if ($asjob) {
    Write-Verbose "Exporting as background job"
    $exportParam2.Add("AsJob",$True)
  }

  Write-Verbose "Exporting virtual machines"
  <#
   Go through each virtual machine name, and export it using Export-VM
  #>
  foreach ($name2 in $MonthVMnames) {
    $exportParam2.Name=$name2
    #if the user did not include -WhatIf then the machine will be exported
    #otherwise they will get a WhatIf message
    if ($PSCmdlet.shouldProcess($name2)) {
       Try {
            Export-VM @exportParam2 |Out-GridView -Wait
                  if($VhdSizeMS -eq $size2S){

    Write-Host "Export script Finished." -ForegroundColor Green
}
else{
Write-Host "Export Failed." -ForegroundColor Red
}
       }
       Catch {
        Write-Warning "Failed to export virtual machine(s). $($_.Exception.Message)"
       }
  
      } #whatif

    } #close foreach

  } #if backup folder exists
   
} #Process

End {
    Write-Host "Export script Successful." -ForegroundColor Green
    }
}



Function Generate-Form
{
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Computer'
$form.Size = New-Object System.Drawing.Size(500,300)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(50,200)
$okButton.Size = New-Object System.Drawing.Size(80,25)
$okButton.Text = 'ÓK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$okButton1 = New-Object System.Windows.Forms.Button
$okButton1.Location = New-Object System.Drawing.Point(150,150)
$okButton1.Size = New-Object System.Drawing.Size(80,25)
$okButton1.Text = 'Weekly'
$okButton1.Add_Click({okButton1_Click})
$form.Controls.Add($okButton1)

$okButton2 = New-Object System.Windows.Forms.Button
$okButton2.Location = New-Object System.Drawing.Point(250,150)
$okButton2.Size = New-Object System.Drawing.Size(80,25)
$okButton2.Text = 'Monthly'
$okButton2.Add_Click({okButton2_Click})
$form.Controls.Add($okButton2)



$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(350,200)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Your Computer Name is'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(450)
$listBox.Height = 80


[void] $listBox.Items.Add($hostname)


$form.Controls.Add($listBox)

$form.Topmost = $true
$result = $form.ShowDialog() | Out-Null 
}

#$MonthVMnames , $VhdSizeM | Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" -Append
#$VhdSizeM = (Get-VM -Name $MonthVMnames | Select-Object VMId | Get-VHD | Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} |  Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" -Append)

#$WeekVMnames , $VhdSizeW | Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDWLog.txt"  -Append
#$VhdSizeW = (Get-VM -Name $WeekVMnames | Select-Object VMId | Get-VHD | Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDWLog.txt" -Append)
 
#$WeekVMnames , $size1| Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt"  -Append
#$size1= (Get-ChildItem -Path "D:\FleminScripts(Don'tDelete)\Sample_Exports\Week\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt" -Append)

 
#$MonthVMnames , $size2| Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt"  -Append 
#$size2= (Get-ChildItem -Path "D:\FleminScripts(Don'tDelete)\Sample_Exports\Month\" -Recurse -Include *.vhd, *.vhdx, *.vhds, *.avhd, *.avhdx | Get-VHD |  Select -Property Path,@{label='Size(MB)';expression={$_.filesize/1mb -as [int]}} | Out-File -filepath "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt" -Append)




#$WeekVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt" 

#$MonthVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" 

#$MonthVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt" 


#$VhdSizeW | Out-File "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDWLog.txt" -Append default
#$size1 |Out-File "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt" -Append default

#$MonthVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" 
#$MonthVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt" 
#$VhdSizeM | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt"
#$size2 |Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt"
Generate-Form

Clear-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDWLog.txt" 
Clear-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" 
Clear-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt"
Clear-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt"


Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDWLog.txt" -Value(Get-Date)  
Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt" -Value(Get-Date)

Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" -Value(Get-Date) 
Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt" -Value(Get-Date)


$WeekVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt" 

$MonthVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" 

$WeekVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDWLog.txt"

$MonthVMnames | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt"

$size1 | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupWLog.txt" 

$VhdSizeM | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDMLog.txt" 

$VhdSizeW | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\VHDWLog.txt"

$size2 | Add-Content "D:\FleminScripts(Don'tDelete)\Sample_Exports\BackupMLog.txt"


