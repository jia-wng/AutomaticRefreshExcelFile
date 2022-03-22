

#Orginal Link: https://github.com/TylerNielsen/powershell-refresh-excel/blob/master/RefreshExcelFiles.ps1

#VARIABLES
$files = @(
	#List of file paths go here:
	"C:\Users\DataSource\Test1.xlsx",
	"C:\Users\DataSource\Test2.xlsx"
	

)
$Date = (Get-Date -Format 'dd-MM-yyyy')
$errorFile = "C:\Temp\RefreshExcelError_" + $Date + ".txt" #Where you want an error file to be generated.
$isError = $false

#Get-TimeStamp, What time period did the file have a failure?
function Get-TimeStamp{
    return "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)

}

#Function to test filelock. Found on http://stackoverflow.com/questions/24992681/powershell-check-if-a-file-is-locked
function Test-FileLock {
  param (
    [parameter(Mandatory=$true)][string]$Path
  )

  $oFile = New-Object System.IO.FileInfo $Path

  if ((Test-Path -Path $Path) -eq $false) {
    return $false
  }

  try {
    $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

    if ($oStream) {
      $oStream.Close()
    }
    $false
  } catch {
    # file is locked by a process.
    return $true
  }
}

#Loop through all files, attempt to grab lock and refresh.
foreach ($file in $files){

	#Ensure file exists
	IF (!(Test-Path $file)) {
		Write-Host $File" not found. `n" -foregroundcolor Red
		
		$errMsg = "FILE NOT FOUND: " + $file + " _Time:" +$(Get-TimeStamp)
		Add-Content $errorFile $errMsg
		
		$isError = $true;
		CONTINUE;
	}
	
	Write-Host $file -foregroundcolor Green
	Write-Host "Checking for lock in file..." -nonewline
	
	#Check if file is locked
	IF (Test-FileLock $file){
		#File is locked.
		
		#Check if there is an error file yet.
		IF (!(Test-Path $errorFile)){ 
			#Error file doesn't exist, create one
			New-Item $errorFile -type file
		}
		
		#Add entry to log file
		$errMsg = "FILE LOCKED: " + $file + " _Time:" +$(Get-TimeStamp)
		Add-Content $errorFile $errMsg
		Write-Host "file locked." -foregroundcolor Magenta
		Write-Host "Error added to"+$errorFile -foregroundcolor Magenta
		
		$isError = $true;
	
	} ELSE {	
		#File is NOT locked.
		Write-Host "file available."
		
		$excelObj = New-Object -ComObject Excel.Application
		$excelObj.Visible = $false

		#Open the workbook
		$workBook = $excelObj.Workbooks.Open($file)
		
		#Refresh all data in workbook.
		Write-Host "Starting refresh..." -nonewline
		$workBook.RefreshAll()
		Write-Host "done." 
		
		Write-Host "Saving file..." -nonewline
		$workBook.Save()
		Write-Host "done." 
		
		#Close workbook.
		$workBook.Close()
		$excelObj.Quit()
		
		#We must decrement the CLR reference count (to prevent the process from continuing to run in the background, which causes memory and lock problems).
		#https://technet.microsoft.com/en-us/library/ff730962.aspx
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObj)
		Remove-Variable excelObj
	}
	
	Write-Host "`n"
}

Write-Host "`n"

#If an anticipated error found above, open the error file.
IF ($isError){
	Invoke-Item $errorFile
}
