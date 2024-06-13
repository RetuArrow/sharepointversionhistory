#Config Variables

param (
# Default number of old versions to keep
    [string]$oldversions = 50,
    [string]$delete = $( Read-Host "Delete previous versions (y/N)?" ).ToUpper()
)

write-host "Number of old versions to leave ('-oldversions n', default " $oldversions "): " $oldversions
write-host "* Authenticating user"

$SiteURL = "https://acme365.sharepoint.com/sites/sitenam/subsite/"
$ListName = "Documents"

# Use -Interactive (or none) for other systems than Windows, or -UserWebLogin
Connect-PnPOnline -Url $SiteURL -Interactive

#To Get All Lists
# Get-PnPList

#Get the Context
$Ctx= Get-PnPContext


$global:counter=0
#Get All Items from the List, Sharepoint limits queries to 5000 items - fixme

$ListItems = Get-PnPListItem -List $ListName -Fields File_x0020_Type, Created -PageSize 4999 | Where {($_.FileSystemObjectType -eq "File")}

# Filter for specific file types?
$DocumentItems = $ListItems 
#| Where-Object { $_["File_x0020_Type"] -in ("docx", "xlsx", "pptx") }

ForEach ($Item in $DocumentItems)
{
    #Get File Versions
    $File = $Item.File
    $Versions = $File.Versions
    $Ctx.Load($File)
    $Ctx.Load($Versions)
    $Ctx.ExecuteQuery()

    Write-host -f Yellow "* Scanning File:"$File.Name
    $VersionsCount = $Versions.Count
    If($VersionsCount -gt 0)
    {
        write-host -f Cyan "`t Total Number of Versions of the File:" $VersionsCount
        #Delete versions
        For($i=0; $i -lt $VersionsCount; $i++ )
        {
	    if($i -lt $VersionsCount-$oldversions) {
		$deletetag = "*"
	    } else {
		$deletetag = ""
	    }
            write-host -f Cyan "`t Found Version:" $Versions[$i].VersionLabel $Versions[$I].Created $deletetag
            
		if(($delete -eq "Y") -and ($i -lt $VersionsCount-$oldversions) -and ($VersionsCount -gt 1)) {
                write-host -f Cyan "`t * Deleting Version:" $Versions[$i].VersionLabel
        	$Versions[$i].DeleteObject()
	        $i--
		$VersionsCount--
		Write-host -f Cyan "`t Versions left: " $Versions.Count
		
            }
        }

	# Wait for it?
	#write-host -f Green "Wait 5s"
	#Start-Sleep -Seconds 5

        $Ctx.ExecuteQuery()
        Write-Host -f Green "`t Processed File:"$File.Name
    }
}
