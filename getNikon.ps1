# set the camera folder
$Source = "D5300\Removable storage\DCIM\103D5300"

# set the destination folder
$Destination = "D:\Camera\2021"

# do we want to delete files from the camera?
$cleanup = $true

Clear-Host
Write-Host "Copy from " $Source " to " $Destination

# Create a shell application  
$Shell = New-Object -ComObject Shell.Application

# Get the 'This PC' list of items 
# 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
$ShellItem = $Shell.NameSpace(17).Self

# now make the camera folder a list of sub-folders
$PathArray = $Source -split "\\"

# get to the camera folder
$CameraFolder = $null
foreach($item in $PathArray){
  if(!($CameraFolder)){
    $CameraFolder = $Script:ShellItem.GetFolder.Items() | Where-Object {$_.Name -eq $item}
  }
  else{
    $CameraFolder = $CameraFolder.GetFolder.Items() | Where-Object {$_.Name -eq $item}
  } 
} 

if(!$CameraFolder){
  Write-Host "Camera folder not found: " $Source
  pause
  return
}

# get the items in the folder
$CameraItems = $CameraFolder.GetFolder.Items()

# set up the destination
$DestinationFolderShell = $Shell.NameSpace($Destination).self
$DestinationFolderShell.Path
if(!(Test-Path -Path $DestinationFolderShell.Path)){
  Write-Host "Unable to find destination folder:"  $Destination
  pause
  return
}

# now copy everything that doesn't already exist
$skipped = 0
$copied  = 0
$deleted = 0

foreach($File in ($CameraItems | Sort-Object -Property Name)){

  $FilePath = Join-Path -Path $DestinationFolderShell.Path -ChildPath $File.Name
  if(Test-Path -Path $FilePath){
    ++$skipped
  }
  else{
    $File.Name
    $DestinationFolderShell.GetFolder.CopyHere($File)
    ++$copied

    if($cleanup){
      $File.InvokeVerbEx("Deletey")
      ++$deleted
    }
  }
}
Write-Host $copied " Files copied  " $skipped " Files skipped " $deleted " Files removed"
pause