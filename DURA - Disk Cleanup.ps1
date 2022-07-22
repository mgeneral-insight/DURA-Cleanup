####################################
########## Config Values ###########
####################################

$HName = & hostname
$SMTP_Relay = "mail.authsmtp.com"
$SMTP_User = "ac68547"
$SMTP_Email = "NoReply@insight.com"
$SMTP_Password = "ysmtwky7gpudva"
$Email_Subject = "RE: $INCnum - $HName C Drive Automated Cleanup Results"
$SMTP_Port = "587"

$TempDirs = @'
C:\Windows\Installer\$PatchCache$\Managed
C:\Windows\ccmcache
C:\temp
C:\windows\temp
C:\$Recycle.Bin
'@ -split "`n"
If($TempDirs.Count -le 1) {
$TempDirs = @'
C:\Windows\Installer\$PatchCache$\Managed
C:\Windows\ccmcache
C:\temp
C:\windows\temp
C:\$Recycle.Bin
'@ -split "`r`n"
}

###########################################
########## Get Disk Info ###########
###########################################

Function Measure-DiskSpace{
    $Volume = 'C'
	$Disk = Get-Volume $Volume
	$Disk_Total = [Math]::Round($Disk.size / 1GB)
	$Free_Space = [Math]::Round($Disk.SizeRemaining / 1GB,2)
	$Free_Space_Percentage = [Math]::Round($Free_Space/$Disk_Total * 100,2) #.ToString("P")

	$OutputObj = [PSCustomObject] @{
		Volume = $Volume
		'Capacity' = "$Disk_Total"
		'Space Free'= "$Free_Space"
		'Free %'= "$Free_Space_Percentage"
    }
$OutputObj
}

$InitVolumeScan = Measure-DiskSpace

#########################################
########## Carbon Black Bypass ##########
#########################################

if ( Test-Path -Path "C:\Program Files\Confer\Uninstall.exe" ) { & "C:\Program Files\Confer\Uninstall.exe" /bypass 1 GQ117MS8 }
Start-Sleep 1

########################################
########## GET PAGE FILE INFO ##########
########################################

if ((Get-WmiObject Win32_Pagefile) -ne $null) { 
    $PFInfo = Get-WmiObject WIN32_Pagefile | Select-Object Name, InitialSize, MaximumSize, FileSize, Drive
    if ( $PFInfo.Drive -ne "c:" ) {  
        $PFtxt = "<br><br><h3>Page File is Custom, but is on a different drive.</h3>"
    } else {
    $pname = $PFInfo.Name
    $csize1 = $PFInfo.FileSize/1GB
    $csize = "$csize1 GB"
    $isize1 = $PFInfo.InitialSize * 1024 * 1024
    $isize2 = $isize1/1GB
    $isize = "$isize2 GB"
    $msize1 = $PFInfo.MaximumSize * 1024 * 1024
    $msize2 = $msize1/1GB
    $msize = "$msize2 GB"
    $PFtxt = "<br><br><h3>Page File is <span style='color: red'>CUSTOM</span>.</h3> <br> File: $pname <br> Initial Size: $isize <br> Maximum Size: $msize <br> Current Size: $csize"
    }
}
else { $PFtxt = "<br><br><h3>Page File is <span style='color: green'>System Managed</span></h3>"  }

#############################################
########## Disabled User Profiles ###########
#############################################
##### - Get list of user profiles that exist in C:\Users and check if any of those users are in the Disabled OU - #####

Function Get-DisabledUsers {
    $Attributes1 = "Archive, Compressed, Device, Directory, Encrypted, Hidden, IntegrityStream, Normal, NoScrubData, NotContentIndexed, Offline, ReadOnly, ReparsePoint, SparseFile, System, Temporary"
    $userdirs = get-childitem -path "C:\users\" -Name -Directory -exclude public, localadmin, MSSQL*
    $DUArray =@()
    $DUReclaim = 0
    foreach($User in $userdirs) {
        $objSearcher=[adsisearcher]""
        $objSearcher.Filter = "(&(objectClass=user)(sAMAccountName=$User))"
        $objSearcher.SearchRoot = [adsi]"LDAP://OU=Disabled,OU=Accounts,DC=corp,DC=duracell,DC=com"
        $firstObject = $objSearcher.FindOne()
        if ($firstObject -ne $null) {
            $du_file_list = Get-Item -Path "C:\users\$User" -Force | Get-ChildItem -Attributes $Attributes1 -Recurse -File
            $du_file_size = ([Math]::Round((($du_file_list |Measure-Object -Property Length -Sum).Sum / 1GB),2))
            $DU_delete = Get-CimInstance -ClassName win32_userprofile | Where-Object { $_.LocalPath.split('\')[-1] -eq "$User" } | Remove-CimInstance
            if (-not (Test-Path "C:\Users\$User")) { 
            $du_deleted = "<span style='color: green'>Y</span>"
            } else { 
            Remove-Item -path "C:\Users\$User" -recurse -force 
            if (Test-Path "C:\Users\$User") { 
                $du_deleted = "<span style='color: red'>N</span>" 
            } else {
                $du_deleted = "<span style='color: green'>Y</span>"
            }
        }
        if ( $du_deleted -eq "<span style='color: green'>Y</span>" ) {
            $DUReclaim = $du_file_size
        }
            $DUObj = [PSCustomObject] @{
                User= $User
                Files= $du_file_list.Count
                'Size (GB)'= $du_file_size
                Deleted = $du_deleted
                Reclaim = $DUReclaim
            }
            $DUArray += $DUObj
        }
    }
$DUArray
}


$DU_List = Get-DisabledUsers
if ($DU_List -eq $null){
    $DU = [PSCustomObject] @{
        User= "NO Disabled Users Found"
    }
} else {
    $DU = $DU_List | Select-Object User,Files,'Size (GB)',Deleted,Reclaim
    $DU_Files = ($DU.Files | Measure-Object -Sum).Sum
    $DU_Total = ($DU.'Size (GB)' | Measure-Object -Sum).Sum
    $DU_Reclaim = ($DU.Reclaim |Measure-Object -Sum).Sum
    $Total_HTML2 = @"
        <tr style="border: 2px solid black">
            <th>TOTALS:</th>
            <th>$DU_Files</th>
            <th>$DU_Total</th>
        </tr>
        <tr>
            <td colspan='4' align='center'><br><h4>Total Space Reclaimed from Disabled Users: $DU_Reclaim GB</h4></td>
        </tr>
"@
}
##########################################
########## Delete Temp Profiles ##########
##########################################

function Get-TempUsers {
    $Attributes2 = "Archive, Compressed, Device, Directory, Encrypted, Hidden, IntegrityStream, Normal, NoScrubData, NotContentIndexed, Offline, ReadOnly, ReparsePoint, SparseFile, System, Temporary"
    $tempuserdirs = get-childitem -path "C:\users\" -Name -Directory -Include "TEMP*"
#    $tempuserdirs = "TEMP.DURACELL.002"
    $TUArray =@()
    $TUReclaim = 0
    foreach($TU in $tempuserdirs) {
        $tu_file_list = Get-Item -Path "C:\users\$TU" -Force | Get-ChildItem -Attributes $Attributes2 -Recurse -File
        $tu_file_size = ([Math]::Round((($tu_file_list |Measure-Object -Property Length -Sum).Sum / 1GB),2))

        $tu_delete = Get-CimInstance -ClassName win32_userprofile | Where-Object { $_.LocalPath.split('\')[-1] -eq "$TU" } | Remove-CimInstance

        if (-not (Test-Path "C:\Users\$TU")) { 
            $tu_deleted = "<span style='color: green'>Y</span>"
        } else { 
            Remove-Item -path "C:\Users\$TU" -recurse -force 
            if (Test-Path "C:\Users\$TU") { 
                $tu_deleted = "<span style='color: red'>N</span>" 
            } else {
                $tu_deleted = "<span style='color: green'>Y</span>"
            }
        }
        if ( $tu_deleted -eq "<span style='color: green'>Y</span>" ) {
            $TUReclaim = $tu_file_size
        }
        $TUObj = [PSCustomObject] @{
            User= $TU
            Files= $tu_file_list.Count
            'Size (GB)'= $tu_file_size
            Deleted = $tu_deleted
            Reclaim = $TUReclaim
        }
        $TUArray += $TUObj
    }
$TUArray
}

$TU_List = Get-TempUsers
if ($TU_List -eq $null ) {
    $TU = [PSCustomObject] @{
        User= "NO Temporary Profiles Found"
    }
} else {
$TU = $TU_List | Select-Object User,Files,'Size (GB)',Deleted,Reclaim
$TU_Files = ($TU.Files | Measure-Object -Sum).Sum
$TU_Total = ($TU.'Size (GB)' | Measure-Object -Sum).Sum
$TU_Reclaim = ($TU.Reclaim | Measure-object -Sum).Sum
$TU_Total_HTML = @"
    <tr style="border: 2px solid black">
       <th>TOTALS:</th>
       <th>$TU_Files</th>
       <th>$TU_Total</th>
    </tr>
    <tr>
        <td colspan='4' align='center'><br><h4>Total Reclaimed from Temp Profiles: $TU_Reclaim GB</h4></td>
    </tr>
"@
}
######################################
########## Temp Directories ##########
######################################

Function Get-FileAnalysis {
    $Attributes = "Archive, Compressed, Device, Directory, Encrypted, Hidden, IntegrityStream, Normal, NoScrubData, NotContentIndexed, Offline, ReadOnly, ReparsePoint, SparseFile, System, Temporary"
    $ObjArray =@()

    foreach ($TempDir in $TempDirs){
        If(Test-Path $TempDir) {
            $Files_List = Get-Item -Path $TempDir -Force | Get-ChildItem -Attributes $Attributes -Recurse -File
		    $TD_Size = ([Math]::Round((($Files_List |Measure-Object -Property Length -Sum).Sum / 1GB),2))

            Get-ChildItem -Path "$TempDir" -Include * -Recurse -Force | Remove-Item -Force -Recurse 

            $After_Files_List = Get-Item -Path $TempDir -Force | Get-ChildItem -Attributes $Attributes -Recurse -File
		    $After_TD_Size = ([Math]::Round((($After_Files_List |Measure-Object -Property Length -Sum).Sum / 1GB),2))            

            $OutputObj = [PSCustomObject] @{
                Path= $TempDir
                Files= $Files_List.Count
                'Size (GB)'= $TD_Size
                'Size after Cleaning'= $After_TD_Size
                'File List'= $Files_List
            }
            $ObjArray += $OutputObj
        }
    }
$ObjArray
}



#    foreach ($TempDir in $TempDirs){
#        If(Test-Path $TempDir) {
#                    Get-ChildItem -Path "$TempDir" -Include * -Recurse -Force | Remove-Item -Force -Recurse -WhatIf
#}}



$TempFiles_List = Get-FileAnalysis -Path $TempDirs
$TempFiles = $TempFiles_List | Select-Object Path,Files,'Size (GB)','Size after Cleaning'
$Count_Total = ($TempFiles.Files | Measure-Object -Sum).Sum
$Size_Total = ($TempFiles.'Size (GB)' | Measure-Object -Sum).Sum
$After_Size_Total = ($TempFiles.'Size after Cleaning' | Measure-Object -Sum).Sum
$TempFiles_TotalSaved = ( $Size_Total - $After_Size_Total )
$Total_HTML = @"

        <tr style="border: 2px solid black">
            <th>TOTALS:</th>
            <th>$Count_Total</th>
            <th>$Size_Total</th>
            <th>$After_Size_Total<th>
        </tr>
        <tr>
            <td colspan='4' align='center'><br><h4>Total Space Reclaimed from Temp Dirs: $TempFiles_TotalSaved GB</h4></td>
        </tr>
"@

################################################
########## Remove Carbon Black Bypass ##########
################################################

if ( Test-Path -Path "C:\Program Files\Confer\Uninstall.exe" ) { & "C:\Program Files\Confer\Uninstall.exe" /bypass 0 GQ117MS8 }

###########################
########### HTML ##########
###########################
$TempPre = "<h2>Temporary Directories</h2>"
$DrivePre = "<h2>Drive Analysis</h2>"
$DUPre = "<h2>Disabled Users Profiles</h2>"
$TUPre = "<h2>Temporary Users Profiles</h2>"


Function Set-HTML{
    param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)][object[]]$Array,
    [string[]]$Property,
    $PreContent,
    $PostContent,
    $InTable,
    [switch]$List,
    [switch]$Logo,
    [switch]$IncludeCSS
    )
BEGIN{
$Array2 = @()
}
PROCESS{
$Array2 += @($_)
}
END{

#Adds the Insight logo to top Right corner of HTML page being generated.
If($Logo){
$Insight_Logo = @'
<img align=Right width=100px min-height=100% padding=1px src=data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAfIAAADXCAYAAADhh9PhAAAAAXNSR0IArs4c6QAAIABJREFUeF7tnQmQVcW5xxvZmWFTYEZcGFk0CpnBl4cEkZdgSkTBegHZzGNgMFqyDInGYbFcEDWRzcREEeLKZkRkMQUoYiK4oFE0AQQ0yrAYEQaMCDPDKpxX3yUHr8O9c073191nuf9bRZn3ptff93X/ezvdNQR+IAACIAACIAACkSVQI7IlR8FBAARAAARAAAQEhBxOAAIgAAIgAAIRJgAhj7DxUHQQAAEQAAEQgJDDB0AABEAABEAgwgQg5BE2HooOAiAAAiBgh8ChLw44u1Z8KvZvLBNHyipEg/ObiKY/aClye7YTtbPrBqqlgWZuBz9yAQEQAAEQAAE1Agd3HnA++s1qsfPFj1ImULN+bdF2xGXiotuvCExPA8tYDSligQAIgAAIgIAdAv9e+7mzdthicXTfIc8MaXbeeW5/UadxPeu6aj1DTxoIAAIgAAIgEFsC23rd5zjfHJeqX+O+XcRZt/S0qleVO/Y5b1w9WxwrP+K7rM1/lCe6/Gmg1XJS4axn6JsIAoIACIAACMSOwMbGgxznmJyQn1XcS7ScUmRVr16/Zrazf8Nuaf75k3qIvMJLrZbVambSRBABBEAABEAgVgSiIOS7Xv7EWXvTEiXudZtniavXFVvVVquZKVFBJBAAARAAgdgQiIKQf1C81Nm5ZLMy827LC0XTji2t6au1jJSJICIIgAAIgEBsCERByN/oNcf5et0uZeYF064RrW7It6av1jJSJoKIIAACIAACsSEQBSFf1f0pp/yTL5WZX3zH/4h2xV2s6au1jJSJICIIgAAIgEBsCERByN/q+6zz1bufKzPPn3y1yBvc0Zq+WstImQgiggAIgAAIxIZAFIR8/dgVzo5n1ysz77roZ+KsH55nTV+tZaRMBBFBAARAAARiQyAKQl72WqnzbuFCJeZ1mtYXPTf+wqq2Ws1MiQoigQAIgAAIxIZAFIScYL9x7Rzn6/XyB97yH+wh8obgO/LYOCwqAgIgAAIg8F0CURHyiu37nDd7St7s1q2V6DJ/kPUJsvUM4dQgAAIgAAKZSyAqQk4WorvW3xuyUBw74H1N61mdzxOdZvUVdRrhrvXM9W7UHARAAAQygECUhJzMcfDz/c5Hv3ld7Pxz6tfPamXVEW1H/1BcONre52ZV3QQz8gxoOKgiCIAACISFQNSE3OVGgr7r5U/EgU17xOE9FaJhu2aJ98hzftJG1MqqE6iWBpp5WBwL5QABEAABELBDIKpCboeOWi4QcjVuiAUCIAACIKBAAEKuAM0jCoRcP1OkCAIgAAIgkIYAhFy/a0DI9TNFiiAAAiAAAhByaz4AIbeGGhmBAAiAAAhgRq7fByDk+pkiRRAAARCIFYGjnx9wDrz4iShfuU0c2bRXnCg/Juq2byayfnS+aPzTC0X9jjm+tQRCrt81fMPXnzVSBAEQAAEQCDOBb/YdcnaOfEWUL9tSbTHr5bcQ583qLeq2O9NTU0wJ+b5nNzpHd+wPM06lsjW6to2o3zG3Wq6e0JVyRiQQAAEQAIFIE/hm/2FnW/c/iSOffuWrHjWb1BOtV/+fqNumabW6YkrIt/Ve4FS+/pmvskYpUMtHeogzi/Ih5FEyGsoKAiAAAmEgsGPQi0758upn4lXLWa9Dc9H2naHBCHmv553KN/4VBnRaywAh14oTiYEACIBAZhA4+P4uZ2v3Z5Uqe+5T14omAy5JK+bGZuRxFfI/XCXOHFaAGbmSNyISCIAACGQogX8VLXX2L/qnUu2zup0nLnhpIIRcid7pkVpCyDWRRDIgAAIgkEEENuf+3jlReUy5xheX/ULUbJD6/nHMyOWwYmldjhdCgwAIgEDGEzhefsT5qOUjLA5t3ysS9S5ulnJWDiGXQ4sZuRwvhAYBEACBjCdAn5x9fP50FofWf/2ZaHBZS7tCjlPrLJshMgiAAAiAQIwIbGw0zRGOeoXafXiTqJvXxK6Q47CbusEQEwRAAARAIF4ESrs/6xx6f5dSpWqeWV9cvGMUDrsp0Ts9EpbWNYFEMiAAAiCQSQS+nPGBs3vsKqUqn3nLpaLltJ9AyJXopRByXAijiSSSAQEQAIEMI/BJx6eco6X7pGpds2k90e7vN4pazRrYF3LskUvZCoFBAARAAARiTuDwP//tbO02T5w45PMztBpC5L3YT2RfmYeb3TT6BpbWNcJEUiAAAiCQaQQOfrDL+az/EvHN3oPVVv2M+rXFuU9fKxr1buf5fgc+P5PzIgi5HC+EBgEQAAEQqELgm68OOXsnvSP2zdsoTpQfPY1Pk4EXixYTrhB1zmvsKeIU2ZSQZ7LhfIHPZECoOwiAAAiAwEkCFW/+yzny8b+Fc/gbUfd7Z4kGXc9Je4NbOmYQcv3eBCHXzxQpggAIgAAIpCEAIdfvGhBy/UyRIgiAAAiAAITcmg9AyK2hRkYgAAIgAAKYkev3AQi5fqZIEQRAAARAADNyaz4AIbeGGhmBAAiAAAhgRq7fByDk+pkiRRAAARAAAczIrfkAhNwaamQEAiAAAiAQlRn5oS8OOJXb5a6oJes2u7yVdV21niHcGARAAARAIHMJREXISx9f62ya+Jq0oXr8Y5So1yLbqrZazUyaCCKAAAiAAAjEigCEXL85IeT6mSJFEAABEACBiO+RY0YOFwYBEAABEACBFAQwI9fvFpiR62eKFEEABEAABDAjt+YDEHJrqJERCIAACIAAZuT6fQBCrp8pUgQBEAABEMCM3JoPQMitoUZGIAACIAACmJHr9wEIuX6mSBEEQAAEQAAzcms+ACG3hhoZgQAIgAAIYEau3wcg5PqZIkUQAAEQAAHMyK35AITcGmpkBAIgAAIgUPHmZkc4J6RA1D6nmajbJteqXuFCGCkTITAIgAAIgAAIhIsAhDxc9kBpQAAEQAAEQECKAIRcChcCgwAIgAAIgEC4CEDIw2UPlAYEQAAEQAAEpAhAyKVwITAIgAAIgAAIhIsAhDxc9kBpQAAEQAAEQECKAIRcChcCgwAIgAAIgEC4CNgU8qOvf+Ac/3ibELVqipoXthJ1uv2X1Kd2UoHDhRmlAQEQAAEQAAEzBEwL+fFdXzqHps0WRxf9VTiHj3ynEjWyG4i6A3uI+rcVijOaNfHUac8AZhAhVRAAARAAgbATOLpqrVN+w/iwF1O6fGfktRRN/za3Wv0zKeQ0A6+4cYJwKg9VW/YajRuKhnMfELUv61BtWSHk0i6ACCAAAiCQGQQg5K9JG7rHP0aJei2y02rrsXfWOwf6lQhx3OftdnVqi0bL/iBq51+YNk0IubSZEAEEQCAVgY3r/+FMGHObMpxFK1ejP1KmZybi0dXvO+WDxplJPMBUg5qRnyivdPZ3LRIn9nwlVfszzs0Rjd94WpzRoF7KNoKGI4UTgUEABNIRgJDHzzeOrl7rlA/C0rqMZaubkVfe/7hzePrzMsmdClt//DDR4NbBEHIleogUagIV5eXOjq2lYk/Z7sS/DvkdRbOcHJGTezYGqZYtByG3DNxCdlha17u0vq+gv3OiTG427pq55kV5osnrT/GFvGz3Lmf1q68ou0/7/ALRoeBSdLDKBBHRJUC++MK82WLVyhUpobTP7ygGFA6Fv1l0GQi5RdiWsoKQ6xPyE18fcPZ9rw/Lck23LRdn1D99eV1KVLkNdcDgoWLgkGFSebJqjcixJPDaKy870x+a7Ktu3Xv0FMUl4+FzvmjxAnH7B+yR8/ibiI09cn1CfnznHufrH9zAMlPj158StS7KO60/k+rguA0VQs6yISILIZ6f84yzYN5sKRYQcylcyoG5/QOEXBm9sYiYkesTch0z8ibvzBE1LzgHQm7M45GwcQLvrnnTmTLxbqV8MIhUwiYVCUIuhSsSgXHYTZ+Qk8G/atPb8fp2vDrHOGv3X/l75NyGis40Em03tIUcUTjIoQNtKr+s7Gzx2OznRHbDhlKrUCp5ZWocbv+AGXn4PAczcr1CXn7jBOfoS28pGbp2906i0XOTIORK9BApFAQ4s3G3AkXDR4nr+vaHkBuyKITcENgAk4WQ6xXyYx9sdg70Gq1k0UZL/yBqd2oPIVeih0ihIPDI1AdZX0xQJS7JLxD3T/s9hNyQRSHkhsAGmCwOu+kVcjJlxbiHnSOzl0pZtd6w/xVZD/5Cz81u3IaKpXUp2yFwEoG7S37pbN6wnsWkRU6umDF3PoScRTF9ZG7/gKV1Q4ZhJIs9cv1CTuY40PdXzrG3/fVntX/SWTR69jf67lrnNlQIOaNFZXhUHUJOCCEW5hyJ2z/E2TacOziat8gRV159TSAD0OPbdjpHFv7FnNMElHKNJg1F/Zv7BvZoClW78p7HnMOPL0pLoEbdOqJe8SDRYMxQT9t7BkjOhdtQIeQBeW0MstUh5A2yssTcJculfD4G6KxVgds/xFnIOWywJWTNhb+TkcnXz9yMju/Y5RxZsFIce+MDcfyjbULUqS1qXXyBoINtdfpfJWrmnOWrv/IVyM2U44yUBoQ8GIeMQ67YIw+/Fbn9A4Q8tY0h5MH4vg0h11UzCLkukkjHKAGcWjeKV0viEPL0GDlsIORa3FM6EQh5GmSYkUv7EiIkERg+eKCzd0+ZEhNaVp8xZz6+I1ei5y8SR6woB8zIMSP352l2QkHIIeR2PC3DcuHMyjGINO8sEHLMyM17mb0cIOQQcnvelmE5qdy13qlLVzF+4q+ltpEyDKuW6kLIIeRaHCkkiUDIIeQhccV4FmPp4hecWTOn+6rcj6+6WowecwdE3BctXiAIOYRc1oOOVx52yl/+uzj43ifiyMefi1otmoh6328lGvf5oahzfvNA2y2EHEIu688IL0mAvstdMHeWeO/tt8TBysrTYtMsvHfffniPXJIrJziEHELu13+OffGVs/vOeWL/n/8mnMPHUkar/99tRYuxfUWj3p0CEXQIOYTcrz8jnAYCJCCUTGVFhaDHUfJat8WhNg1cZZOAkEPI/fjM4X9+7my7aoL45ssDfoKLs6cOE81GXWtdzCHkEHJfDopAIBAnAhByCLmXPx+vOOxs6TJGHC2Ve8Uw7893ioZXdbQq5hByCLmXP+PvIBA7AhByCLmXU++ZvMgpmzjfK9hpf6/TJldc9OEjEPI05KTAcBsqPgGS9l9EAIHIEOD2D/iOPLWp43QhzEfn/9zxu6RelUarBWOt7pdjRo4ZeWQ6XxQUBHQRgJBjRl6dL1Wu+cjZetU9yu7WtLC7OPePI6Umn8qZCSEg5BByjv8gLghEkgCEHEJenePu+9Przuc3Pars2/UvbS3arpkMIU9BUAoKt6FiaV3ZhxERBEJPgNs/YGk93kvrXz39F2dn8R+V/bjW2U3FxaWPS2mWcmaYkadHByHnuBXigkC4CUDIMSOvzkP3L37H+Wzwb5WduO7F54kLP/gthBwzcmUfQkQQAAEPAhByCHl1LnJo4w5ny2Ulyu2ocd8u4vx5v4KQQ8iVfQgRQQAEIOTKPsAZ5MTp1PrH3xvpHPtsrxLH82bfKpr07wohz3Qh37bl08QzmNtKt4i9ZbvFnt27T/63rPrLCdzbwohf+4KO4oI2bUXzFjnigrbtrDmVkuf7iFRRXu5s/nB9gsn20i2J29E2bVjnGbN9fseTt6i1aRsrHp4V1xyA+O/YWio2blgnDlZUfOubPn3StUGLnFxB/9oX2L00IxkHR6wonVR75JvWr3OIjeub27ee9NF0P2LQnFjk5gryURLBnNyzA2+nHDZxEvJ9z652Pr/Z3zsJyTaud8l5ot379pbVKW+cWk/TymzvkSd3AiRO1XUAKv0zdaLUWVx2+RWh6TD81IPuKV/96iuJe8qpg9T1c3kQk06Xd9XegVK5vyxTf488LAMv6tTXvr0mId46+bt2JP402KRB5yXfL7B2XS1HrFwhp4HN2nfWJHyT/un4kbhTG6UHdILyAQ4bGizfOLyYjaJV6zbWfKG6wm7vN8kpf+kD3/U5I7ueaPPWJFHvwnOsDsi+WPaxs+2Zv/supxuw05N9RJ2m9a2WVSozjjNSJW0I+Xtvv+W4nYBu4fayKHUW9K97j55SXL3S1fV3st+CubN9zbh15EkdUPereiY60OyGDdlMVJ4wdesR9KyGBiEvLVmUECevFSAd7JPTcO3QPr/AqJCx+4fCIrF8yULtA+6qLHr36We9jXLZ6PCJiVN/F4pHhI6XH3J29J8sKt/Y5Fmtmo2zxPkLxorsbpew+w/PzCIcQAoO1xlNCTmVy51h2hbvVLanGcCAwiLrnUU6PyQ+z8ycbmT258f3aaZOA5z+g4eyZulRFHLbgycve7iibmLFhNs/eJVd59+pjRYNHyU6d+0m1QeqliEMbMIi5C7DPVMXO3unLhEnKg6nxNq4f1eRe9/PRJ1WLazYSNW2YYgnBYjrjDqFnJbgSLxpBG97huPXcNRZjCoZF9gomBi9MG+2WLZkod8iGw/XvUdPUXTLKKUZepSEnGbg06dNtrb6oWK4/6we0SBLqh+obsA4YcxtKkUJLA5tQ1AbNb2Pzu07dQAKm5BTnegRlQPL3hMH3/1EHN1aJmq1aCzqF1wgGv20s6hzbjMtfqmDXdjTkALFdUYdQk7itPzFRcaX4HQajiNequWgg32PPjQ5sFl4deWmGfqAwUX0XriU/0VByF3/pLfSo/KjAWevPv3YWyDc/iEoXuSPo24fZ3R2HgY2YRTyoGwet3ylOlKuM3KF/Pm5sxzTe2imDExLmsW3jzO6R+mWnUR8wtjbjO41cjmp7FmHXcjDPHjyYy8SNNr+uK5vf6l+wU2b2z/4KaPJMLR3PmxEsVLdvcoVBjYQci8rRffvUk7LdUaukF/f48dOdFGLxOdaE6f8zqiYR0HEyYYqnUqYhfzdNW860x+aHOrBk5+2ozLAiouQUz1o9ay4ZLxUv+iHK7fv9JOHVxiVNueVJv4eDgJSDst1xkwXcjK5STGnZd2RQ28IvZh06tJVjJ/4aynfI3ZhFfLXXnk5IeJx+GW6kJsSc27fqcO3IOQ6KIYzDanOlOuMEPKTTmBKzO8pudXxc5lL0K742JznlA4XhVHI4yTi5BcQ8pOto1ef68WNI0ZL9Y/VtStu36mjzULIdVAMZxpSjsp1Rgj5t05Ae+a0zK7j+2pKNSqCQt+Ujx5zh5TfudTCJuRRYS7T9UDIv6U1dsL92g7AcftOGRumCwsh10ExnGlIdahcZ4SQf9cJOKJW1Z1GFA5ywvoZnlvWBllZYtqMJ5Vm42FbWqc98SkT7w5nq2aUCkL+LTxaOXts9nNaBtvcvpNh0lNRIeQ6KIYzDQh5wHbR0bh0zgzpDnm62tO9Qz0VHveuevrv5g3rfRPkDuTCMiOPyoFC34ZJCggh/y41+tZ+3L0PSPWTqbhDyFW8EXH8EpByUK4zcjtynafWaXZIl0GcegQlJzflTJEOkNFDDSRam9avS1zwcbCy0i9fz3D0De+MufOl7FA10dtH3ORw7+ymDnxgYZHS5TXkF4lrcde8JehRmlQ/4j1jznzW7CYsQq6DdypG7iDKvSs93b3gNJCorKxI+GPCLzX6ZJiEnMrSIf/koJLaSToe5H/0ABJxICbpfNCzMaYJoGOwze07VcueHE9HPXSUA2noJyAlIFxnDFrIqaO8rOsVifu/VR9PSFz4sWRh4rY0XYJOl1FcefU1UrZwXYFuEBs55AaWZ3Dyr5ox+ciqlSsE3bqX/OPantIKg5BzypDOSLTFQp89dSi4VMkHKF0S91Wvrqh2MOXHSYIWcvqiwX2zQPX8CG17UPuUWS2qjg0NrO6b9rCybShtbt9JIsz95bVuyxpIc/NHfHMEpJyT64zczpwzI6d7lVUvukiFnwR9wbxZYvmSRWzrcGbl3GV13VySBxirV65IDHiysrITe+OqHbObJkdEOQKlc9CU7CxUJnrVSnVQmc7xguLE7R9Uv2ZIx4HKM/neu7QMuLmzWS6bVE+8sjseJBAbAhkj5NyGmM7iXCF101Ut39MzHnFUBxO0QjFz3vNSPiDr+bRiQG++c2abYRHyyffeldhC0PEzNYDizv44A54wihX535R77xLbt5ayzKZ694GbaRjZsIAgcqgISHXiXGcMckauKpR+rKVDzFVPsN9d8ktHdQlR97eyflhxwgQ10+SKY3Kd6azAfVMf1j4LT86D007jJuTEhVbPJoy5lS3mnBUDjk2oDpiRc3qO+MeFkGuy8SNTH0y8xqb6o09d5ixeJmUPyosj5Dr3xlXrLRMvSCHXcdmODRHnDjriKOTEhGbmJSNuYi2zc1ZRIOQyLR1hZQlICQfXGeM6I3ehDx880OGcmFVZNeB8P87pmGQdTUf4oIRcx4FCqr+KfVW4cdppXIWcOHJXzuLMRsXPECc8BCDkGm3BERoqhoqwcgYPWFr3Z3zOOQQ3BxXb+ivd6aEg5OnJcdoLZ4mbYxNOvqo+hHjRIgAh12gv2osbMWSQ8vKdyoEaztK66nK+RmRSSXEGSpzZ1JC+vZ3KigqpsiYH5uStkilHNDhl5eRrS6y4s3LVVZUosFHxNcQJBwEIuWY7TJpwp7P2nTVKqap8hsYRcipklGblQQi5jqtYOYekVByJIxpxF3IabA+9/joVrIk4qtuDHJvYGuQoQ0HEwAlAyDWbYOniF5xZM6crpyp7OlXHsu+AwiK61U3KF5QryIgYhJBz+ap+jcDAxLp8JO5CTlxvH/5zR/VzNNWBL4Sc49GI60VAqvPmOqPqaNatBOdCGNUlMS+AVf9ON2yVjLxZNtqp8LKzN+5SoZsx3V41qmSc8oMmyhWWiBiEkHMOE1LVpj32hNFPzVLh47TTTBByzuBMlQ/HJpiRS3QSGRoUQm7A8DYHHNylwuTq0555rz79RK+fXs++hc0AVutXtHJPq+e1biMemvmUVBvTwY0jGqpCReXm5GtTrIIYEEaFjQ7/Qxr2CUh1MlxnzIQZOZnQppBzlwpTuRwJOt133X/w0FDN0G13wNz9cZsn1ZPtyGmnmSDkXLvKbn9FaZBjX4KQow4CEHIdFKukYVvIdS2vp0JBgk77gjquWOWiti3knPyorrLbJFw+bnwIefUkOXxUVw6CyFOXPyGd8BOAkBuwEecwjepePvf7WC8MdKKelt3p8Bb38ROvvNL9nSOsKjNNzhcBNu6xT8eJIxoqnHQMIFQFUsWXOHxUyxlEnipsECeaBCDkBuzGEQBVITc5K09G5C670z667le7vExhW8g5746r3AngVX+/f+eIBoTcmzKW1r0ZIYRdAhByA7yDEHKqBve+d1kUdNJ9QOFQa8vutoWcs0XCPQ8ia4vk8BByLK1z/Adxo0cAQm7AZkEJua5XnmSR0LI7fYvevUdPKX+SzSdKQh7kgzQQcgi5bNtC+GgTkOp4OR0EYeLOUjgzJNUlaxXzBiXkVFb6ZGpCya2C83iLSp0pDgk6ndTu3LWblF/5zc+mkHN93aa/VeXHKTuW1r29EUvr3owQwi4BqQ6X00FAyP0ZVocABDUzd2to6nIZCLk/H+K0Uwi5N2MIuTcjhLBLAEJugHeQM3K3OiTmj06bJFTvfedioUNxAwYXid59+0n5WHX5RknIg7jRzWUHIcfSOrf9In60CEh1spwOAjNyf46hY0aenBNdfkGCfrCy0l8BNIei2fnYCfdr+WQtSkKuMmvThZ7TTjEj97aCim05NqESqeTpXROEiAsBCLkBS4ZhRp5cLZqdPzPzUbH61VcM1NY7ybw2bcXEKb9ji3mUhBwzcm+/qBrCllgFIapB5ClvAcSIKgEIuQHLhU3I3SrSQbgFc2cFIui01D71sSdYV75GSch1r6zIuClHNDAj9yatMuDg2AQzcm+bZHoICLkBDwirkCcL+uqVK8SyJQutLrlzZ+YQcn/OyhENCLk3Ywi5NyOEsEsAQm6Ad9iFPLnKdCPc8iULher7zLL46O72cfc+IOV3bh42hZzy5HzuSOcCTH2G58UcQl49IQ4f1dlxEHl6+Qn+Hh8CUh0q1xnxHbm34wS1JEu2XbVyhZVld9XLUqIk5Fxf9/aU9CE47RQzcm/ymJF7M0IIuwQg5AZ4R2lGXrX6dDCOZugk6qYulaGLY2bMnS/le1RO20LOefyGHpcZPeYO6TrqcEcIOWbkOvwIaUSHgFRHw+kgCAl3lsJZ6rQ5042ykCe7Ln26Rvvomzes1+7RKrNy20LOsaPqYEUHaE47xYzc2wKYkXszQgi7BCDkBnhzBMDmgMNv1U2cdqeDbw/NeFLK/2wLOSc/Yov3yP162MlwKgIpl8PJ0JyBjmo5g8hThQ3iRJOAVEfKdUbMyL2dJIxC7paaBJ0ul9E1Q5cVOo6wqsw0aUViysS7vY2WJgTdO39d3/5SbUw5s6SInHaqwsnNmpOvqkCq8AqinNw8g7yXQIUx4tglINXJcJ0RQu5t3DALuVv6pYtfcGbNnO5dGY8QssvrtoWcBi4jh9ygXE+61e6+aQ9LtTHlzCDkvtFx+zGVlQNunlHoF3wbAAG1E5DqZLjOCCH3tl9UGix9tjb9ocneFaomRKcuXcX4ib/27YO2hZyKPnzwQIdz6E921YEF9D+ROe0UM3JvC6gIOXdQyO07vWuFEFEm4LsTpUpyOgiKz3VGHHYLl6txhJVqIisanPxk83JJPz3jEWf5kkXK4IM4vc5pp6qcdPQPKgKpYhgOH8pPtZyc/kt20KvCBXGiSwBCbsB2cTvslg4Rfao2Ysgg5dvhZA+8BSHk3H1yYmd7Vs4RKgi5d4cQhJDTFcdzFi+T6q+9a4IQcSEg5RicDgIzcn8uE5Wldbc2kybc6XCeSpXpFIMQcqpnYZ9eDuf1ONt75Zx2CiH3bqcyPpucGmeAT+nIninxrglCxIUAhNyAJTkNNmpCzhFX2WVKTl4cgeIur1M9bZ5gh5BX36g5fGR9Nrkkj0x90OG8QBjk3QQGukkkqZEAhFwjTDcpCLl/qDKzm6Dn83txAAALyUlEQVSEnHtQyaVh6xMijlBxBjycfDkC6d/bToYMqpw6vvawOSCU5YrwwRGAkBtgn0lCzqlr8xY5Yua85337YFBCTi7CqafrYrTPSe+yX9C2ne86q7gnR6gg5N7EZQafyalt2/KpUzLyZu8MPELYGhCyC4oErBGQ6lA4HQTVCKfWve3KWVpftXKFQ6dbsxs2lLKrd6lSh6DDbkOvv041eiROrbuV4/q+TTHnlBVC7u3OqkJOKXPPW1AaNCCk/fKgXtfzJoQQtglIdficDgJC7s+0qkLu2ob20QYUFonuPXpK2dZf6b4birt33KvP9eLGEaN9lzPIGbmuWblLcNjwYtG7bz/fdZexD6edQsi9SXOEnLtPnly67j16iv6Dh4qc3LON+JE3CYQICwEpB+B0EBByfyZXFfJ7Sm51Nm1YdyoT04Ku40IY2Te7gxZyrv9X9QA6zT5s+CjtS+0c20DIvdspR8h1fM5YtYSXXX6FoH82V+O8KSGETQIQcgO0OfupKkJeXcdNgt6rTz/R6fKu2kbuz8+d5SyYO4tNbvaipVLbAEELOVWYuwqRChp1wrQ60aHgUqn2mJwWHchb+/YaQU/Q7inbrWwbCLk3Oo6QU+o6ltfTlZLae/OcXNEi9+R/k390ORFm7972jWIIqY6DOyPBHrm3i6gI+YjCQY6fzpsEg2aBKqJO++H0vTgJuJ+8vGqqclNVGIScOJSMuMnIW+3UCSdsVNBR0EHAdIfiSLS/LCsT20q3iO1bt4hN69dpsQnZDELu5bnqN7u5KXP82Lt06UOo9C2c/BDXHgEIuQHWNmfkqp+00K1qJBz03w75HU9RaJCVdeqmtsrKioRYkFAkL9vrQCa7rE55cjpAjkBVra+u08d+OZKdyBaVFRV+oyiH43DiDvS5M12/lQ66nNwbEf3Ws2o4CLkqufDHg5AbsJEtIacOYeTQG6x08DoxyX52pmMmwxGoVHVXHUDp5GgiLQ6noAXSL48wlDMI/4GQ+/WQ6IWDkBuwmS0h58xQDVTbd5KqHQqnvhyBSlcxnSeQfcMzHJDDKQwC6QdPWMrJ6Sf81FP3jJwmDstfXCRWr1xxaiuHVotoKwin51Usoi8OhFwfy1MpcRqoX5GL6myc8xpY2IScDB43MYeQe3cIurYATJ63SFULv31Lqri0nTRl4t2JMxS9+/Q7dX6DzmuQsK9auUIMG1GMb9u93cdICAi5Aaw2hDyKApLXuo2YOPVhqZPqyeYJo5DHTcwh5N4dgi4hp5xIIO8Zc6vyC4Lepf02hKqQu5OGoltGiSuvvialZrj1uG/qw9o/p5SpY6aGhZAbsLxpIdd197eBqqdNkg7RTZvxJOvzl7AKOVXaxGdpNu3j5gUh96auU8htirmqkNOkgZbQBw4ZVqNqG6Rb5lyBp89g6VGY+6Y9LKUr3sQRwouAFHDu3hI+P/MyhxB+G1uUhIMOt4279wH2SD3MQk6W5VzE4u0ZdkJAyL056xZyypEG549OmyQ2b1jvXQDFEH77lqrJ03fvM+bMT6ykuW2Q0qLfssULxeYP1596K3344IGOjrauWMWMjQYhN2B60zNyt8i0nPX0zEeNNn4uHu5yenL+YRdyd3ZFHfL2raVcdIHEh5B7Yzch5G6u5OPLliw0stSuIuQ0eXt+7ixx/7TfJ7TCbYMuA5qtv/f2W2LukuWJv0+acKdDdyGkW4L3posQKgQg5CrUPOLYEnK3GG5jMzmaV8HEXYGpmmcUhNwtM31eRJfnHKysVEEXWBwIuTd6k0JOudOe9IJ5sxIHyHT6j04hT6ZED7i4wk1tlP5Gy/DeJBFCFwEp2Fha94fdtpAnz9BpNE/7VEH+SAyKS8az9sNTlT9KQu52yHRlqqkZlgkbQ8i9qZoWcrcEJOg021316gotq24qQk6rfpPvvevUc8PJS+u0rL5ja6mY+tgTpw6wUt83sLCIdd2wtwUQoioBCLkBnwhKyKt2ANQJ0LWqtn70aRm9yMS5M7y6skZNyJPtQZ3x8sULjVztmsyMDhXSNbyqdoeQe7cWW0KeXBISdbpdcXvpFrFxwzqxd/duaV9SEXIqA+17F48Zn2jXyUvr7g11nbt2o4F7DfdZY9k3FLyJI4QXASkhJ0PR3c6qP7rEn3NpP60IqOad17qt8mdPsnnSKJau1FT56S6nO6qnToCuWt27p0ylWGnjUMdPe2Ldr+ppnC8dCNqr+CBIVlY2+7CdDnDkGyTqZAtd++iueLt7k5yVM5U78JMHLJz+wdQAsKrduP2YrXL68TeZNqHatySfRnfzcxkk/9/Jp9v9lB1h9BGQEnJ92SKloAi4Dc99aIMeQKFRvtdeHImFez+7e5tTmDq0oHhy8k2eZdGd9jT48zrnkGwHsscFdFd+lVfTONd/6j7XwOGDuOEhQCJNfQTth9Pp9aolo79TP8K5JyI8tY1eSSDk0bOZlRLTzDHd61tWCoBMEgRU7MDZgoCQw/HSEaBPXukA3pU9rkkM6ulH4k1beLQyN2x4sfFVOVgnNQEIOTwDBGJGgCPkKq/SxQwfqlMNAVrRI+GmLS56jY8EnbZ0OFumAM4nACHnM0QKIBAqApzre6c99kQozhKECigKAwIhJwAhD7mBUDwQkCVw+4ibEvuVKr8gTmSrlBNxQAAEviUAIYc3gECMCLifAKlUiW7he2jmU+gTVOAhDggESACNNkD4yBoEdBN4d82biecmVX69+lwvbhwxGn2CCjzEAYEACaDRBggfWYOAbgKc/XEcdNNtDaQHAnYIQMjtcEYuIGCcAPd5W9zIZdxEyAAEjBCAkBvBikRBwD4BzmycczWr/ZoiRxAAgWQCEHL4AwjEgADnWlaqfvILVjHAgSqAQEYRgJBnlLlR2TgSoJPqY0beLOi6XdUfltVVySEeCARPAEIevA1QAhBQJkAiPmHsbYmrMlV/9Grd6DF3oC9QBYh4IBAwATTegA2A7EFAlQDdw/7oQ5NZIk554zY3VQsgHgiEgwCEPBx2QClAQIrAssULnQXzZiXuu+b8cMiNQw9xQSAcBCDk4bADShFhAnQJy96yMtE+v8DoPeW0jL72nTViwdxZrP3wZNSYjUfY8VB0EPgPAQg5XAEEmASSXxvLys4Wea3bivYFHRNvhTdvkcMSd/o2nN4o37RhXeK5SO4MPLmquMmNaXhEB4GQEICQh8QQKEZ0Cfh5NpSee8zKyhYtcnNF85zcaitLT0Tu2b1bbN+6RatwJ2faICtLzJgzH+9HR9ftUHIQOEUAQg5nAAEmAT9CzsxCe3QsqWtHigRBIDACEPLA0CPjuBCImpAXDR8lruvbH20/Lg6IemQ8ATTmjHcBAOASiJKQ45txrrURHwTCRwBCHj6boEQRIxAVIYeIR8yxUFwQ8EkAQu4TFIKBQDoCURBynFCH/4JAfAlAyONrW9TMEoEwCzmdTh82vFhcefU1aOuW/AHZgIBtAmjctokjv9gRCKuQ061txSXjRU7u2WjnsfM6VAgEviWABg5vAAEmgbAJOWbhTIMiOghEjACEPGIGQ3HDRyAsQk63yA0oLMIyevhcBCUCAaMEIORG8SLxTCCwdPELzqyZ0wOpKs2+L7v8isS/zl27oT0HYgVkCgLBEkDDD5Y/co8JAXrQhK5U3bR+ndhWukXQNavbt5YaqR3tfdM97u3zO0K8jRBGoiAQLQIQ8mjZC6WNGAF6M7yysiIh7gcrKoT7v6tWY3vplsRd7M1zv3sPOwk2/f8Td7VnZ4sOBZeizUbMB1BcEDBNAJ2CacJIHwRAAARAAAQMEoCQG4SLpEEABEAABEDANAEIuWnCSB8EQAAEQAAEDBKAkBuEi6RBAARAAARAwDQBCLlpwkgfBEAABEAABAwSgJAbhIukQQAEQAAEQMA0AQi5acJIHwRAAARAAAQMEoCQG4SLpEEABEAABEDANAEIuWnCSB8EQAAEQAAEDBKAkBuEi6RBAARAAARAwDQBCLlpwkgfBEAABEAABAwSgJAbhIukQQAEQAAEQMA0gf8HCczqi6UTGPYAAAAASUVORK5CYII= />
'@
}Else{$Insight_Logo = $null}


#Needed to allow both PIPE and Passing Expression, otherwise only PIPE works. It checks if $Array2 (PIPE) exists, otherwise treat as expression passed to function.
if(!$Array2){
$Array2 = $Array
}

$HTML_Properties_And_Values = @()
$HTML_Properties = @()
$HTML_Lists = @()
$HTML_Values = @()
$HTML_Rows = @()
$HTML_Body = @()
$Number = 0

#If Properties are not defined, will extract from Array being fed into function
if(!$Property){
$Property = $Array2[0].psobject.Properties | Select-Object -ExpandProperty name
}

#Switch to generate HTML Code for a LIST as opposed to a Table.
If($List){
#Dynamically add values based on how many properties are defined/available.
#What happens if array has 4 properties but you define 3? it will create table based on the 3 defined and assign values accordingly
Foreach($Row in $Array2){
    $Number = 0
    $HTML_Properties_And_Values = @()
        Foreach($P in $Property){
            $Value = $Row.($Property[$Number])
            $HTML_Properties_And_Values += "
            <tr>
                <th width=1%>$P</th>
                <td>$Value</td>
            </tr>"
            $Number++
        }
    $HTML_Lists += @"
            <table class="fl-table">
            <thead>
                $HTML_Properties_And_Values
            </thead>
            </table>
"@
    }
}Else{

#Cycle through properties for each column and add to main HTML_Body array
Foreach($P in $Property){
$HTML_Properties += "
<th>$P</th>"
}

#Dynamically add values based on how many properties are defined/available.
#What happens if array has 4 properties but you define 3? it will create table based on the 3 defined and assign values accordingly
Foreach($Row in $Array2){
$Number = 0
$HTML_Values = @()
    Foreach($P in $Property){
        $Value = $Row.($Property[$Number])
        $HTML_Values += "
        <td style='padding-left: 20px; padding-right: 20px'>$Value</td>
        "
        $Number++
    }
$HTML_Rows += @"
        
        <tr>
            $HTML_Values
        </tr>
"@
}
}

# CSS code to define style of HTML table
If($IncludeCSS){
$HTML_CSS = @'
<style type="text/css">*{
    box-sizing: border-box;
    -webkit-box-sizing: border-box;
    -moz-box-sizing: border-box;
}
body{
    font-family: Helvetica;
    -webkit-font-smoothing: antialiased;
}
h2{
    text-align: center;
    font-size: 22px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: black;
    padding: 30px 0;
}
h3{
    text-align: center;
    font-size: 14px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: black;
    padding: 0px 0;
}
/* Table Styles */

.table-wrapper{
    margin: 10px 70px 70px;
    box-shadow: 0px 35px 50px rgba( 0, 0, 0, 0.2 );
}

.fl-table {
    border-radius: 5px;
    font-size: 18px;
    font-weight: normal;
    border: 2px solid black;
    border-collapse: collapse;
    width: 100%;
    max-width: 100%;
    white-space: nowrap;
    background-color: white;
}

.fl-table td, .fl-table th {
    text-align: Left;
    padding: 8px;
}

.fl-table td {
    border-right: 1px #F8F8F8;
    font-size: 12px;
}

.fl-table thead th {
    color: #FFFFFF;
    background: #FF00FF;
}


.fl-table thead th:nth-child(odd) {
    color: #ffffff;
    background: #FF00FF;
}

.fl-table tr:nth-child(even) {
    background: #FFCCFF;
}

/* Responsive */

@media (max-width: 767px) {
    .fl-table {
        display: block;
        width: 100%;
    }
    .table-wrapper:before{
        content: "Scroll horizontally >";
        display: block;
        text-align: right;
        font-size: 11px;
        color: white;
        padding: 0 0 10px;
    }
    .fl-table thead, .fl-table tbody, .fl-table thead th {
        display: block;
    }
    .fl-table thead th:last-child{
        border-bottom: none;
    }
    .fl-table thead {
        float: left;
    }
    .fl-table tbody {
        width: auto;
        position: relative;
        overflow-x: auto;
    }
    .fl-table td, .fl-table th {
        padding: 20px .625em .625em .625em;
        height: 60px;
        vertical-align: middle;
        box-sizing: border-box;
        overflow-x: hidden;
        overflow-y: auto;
        width: 120px;
        font-size: 13px;
        text-overflow: ellipsis;
    }
    .fl-table thead th {
        text-align: left;
        border-bottom: 1px solid #f7f7f9;
    }
    .fl-table tbody tr {
        display: table-cell;
    }
    .fl-table tbody tr:nth-child(odd) {
        background: none;
    }
    .fl-table tr:nth-child(even) {
        background: transparent;
    }
    .fl-table tr td:nth-child(odd) {
        background: #F8F8F8;
        border-right: 1px solid #E6E4E4;
    }
    .fl-table tr td:nth-child(even) {
        border-right: 1px solid #E6E4E4;
    }
    .fl-table tbody td {
        display: block;
        text-align: center;
    }
}</style>
'@
}Else{$HTML_CSS = $null}

# Groups everything together to generate the final HTML code
$HTML_Body = @"
$HTML_CSS
$Insight_Logo
$PreContent
<div class="table-wrapper">
    <table class="fl-table">
        <thead>
        <tr>
            $HTML_Properties
        </tr>
        </thead>
        <tbody>
            $HTML_Rows
            $InTable
        </tbody>        
    </table>
    $HTML_Lists
</div>
$PostContent
"@
$HTML_Body
}#END OF END


}

##


####################################################################################
######          Creating Secondary Array for Before & After Overview          ######
####################################################################################

#Free Space After (Adds Current Free Space + Size_Total of all Temp Files added up)
#$Space_After = 108 ############################################ ---FIX THIS
#$Space_After += [decimal]$VolumeScan.'Space Free (GB)'+$Size_Total

#Free Space After in Percentage
#$Percent_Full_After = ($Space_After/$VolumeScan.'Capacity (GB)').ToString("P")

#If all Temp Files are deleted, provides a snapshot of how much space will be gained.
#$Before_And_After = [PSCustomObject] @{
#    Volume = $VolumeScan.Volume
##    'Total Size (GB)' = $VolumeScan.'Capacity (GB)'
#    'Free Space(Before) GB' = $VolumeScan.'Space Free (GB)'
#    '% Free (Before)' = $VolumeScan.'Free %'
#    'Free Space (After) GB' = $Space_After
#    '% Free (After)' = $Percent_Full_After
#}

###############################################
########## Calculate Reclaimed Space ##########
###############################################
$PostVolumeScan = Measure-DiskSpace

$Disk_Capacity = $InitVolumeScan.Capacity
$Disk_Init_Free = $InitVolumeScan.'Space Free'
$Disk_Init_FP = $InitVolumeScan.'Free %'
[int]$Disk_Post_Free_clean = $PostVolumeScan.'Space Free'
if ( $Disk_Post_Free_clean -le '15' ) {
    $Disk_Post_Free = "<span style='color: red'>$Disk_Post_Free_clean</span>"
} else {
    $Disk_Post_Free = "<span style='color: green'>$Disk_Post_Free_clean</span>"
}
$Disk_Post_FP = $PostVolumeScan.'Free %'
$Disk_Recl_Space = [Math]::Round($Disk_Post_Free_clean - $Disk_Init_Free,2)
$Disk_Recl_P = [Math]::Round($Disk_Post_FP - $Disk_Init_FP,2)





####################################
########## HTML Construct ##########
####################################
$HTML_EmailBody = Set-HTML -Array $TempFiles -InTable $Total_HTML -PreContent $TempPre -Logo -PostContent "<br><hr><br>"
$HTML_EmailBody += Set-HTML -Array $DU -InTable $Total_HTML2 -PreContent $DUPre -PostContent "<br><hr><br>" -Property User,Files,'Size (GB)',Deleted
$HTML_EmailBody += Set-HTML -Array $TU -InTable $TU_Total_HTML -PreContent $TUPre -PostContent "<br><hr><br>" -Property User,Files,'Size (GB)',Deleted

$HTML_EmailBody += @"
<h2>Drive Analysis</h2>
<table>
    <tr><td><b>Total Capacity:</b></td><td>$Disk_Capacity GB</td><tr>
    <tr><td colspan='2'><hr></td></tr>
    <tr><td><b>Initial Free Space:</b></td><td nowrap='nowrap'>$Disk_Init_Free GB</td></tr></tr>
    <tr><td><b>Initial Free %:</b></td><td nowrap='nowrap'>$Disk_Init_FP %</td></tr>
    <tr><td colspan='2'><hr></td></tr>
    <tr><td><b>Post Free Space:</b></td><td nowrap='nowrap'>$Disk_Post_free GB</td></tr>
    <tr><td><b>Post Free %:</b></td><td nowrap='nowrap'>$Disk_Post_FP %</td></tr>
    <tr><td colspan='2'><hr></td></tr>
    <tr><td><b>Reclaimed Space:</b></td><td nowrap='nowrap'>$Disk_Recl_Space GB</td></tr>
    <tr><td><b>Reclaimed %:</b></td><td nowrap='nowrap'>$Disk_Recl_P %</td></tr>
</table>

$PFtxt

"@

#$HTML_EmailBody += Set-HTML -Array $Before_And_After -PostContent $PFtxt -PreContent $DrivePre -List


#################################
########## SMTP Setup ###########
#################################

$socket = new-object System.Net.Sockets.TcpClient($SMTP_Relay, $SMTP_Port)
If(!$socket.Connected) { 
    $SMTP_Port = 25
}
$socket.Close()

$Secure = ConvertTo-SecureString -String $SMTP_Password -AsPlainText -Force
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($SMTP_User, $Secure)

$SNBody = "[code]
	$HTML_EmailBody
    [/code]
"

$Body = "$HTML_EmailBody"
if ( $Additional_Email -ne "none" ) { Send-MailMessage -To $Additional_Email -From $SMTP_Email -SmtpServer $SMTP_Relay -Port $SMTP_Port -Credential $Credentials -Subject "$Email_Subject" -BodyAsHtml $Body }
#Send-MailMessage -To Services-CDCT@insight.com -From $SMTP_Email -SmtpServer $SMTP_Relay -Port $SMTP_Port -Credential $Credentials -Subject "$Email_Subject" -BodyAsHtml $SNBody
Send-MailMessage -To Michael.General@insight.com -From $SMTP_Email -SmtpServer $SMTP_Relay -Port $SMTP_Port -Credential $Credentials -Subject "$Email_Subject" -BodyAsHtml $Body
