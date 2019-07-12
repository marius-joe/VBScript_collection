# ******************************************
#  Dev:  marius-joe
# ******************************************
#  Repair User Profile (Win10)
#  v1.0.0
# ******************************************


Function Main {
	$path_reg_profiles = 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
	
	$ErrorActionPreference = 'stop'
	Try {
	  	Write-Host "Check registry for broken profile keys (extension: *.bak):`n"
		$keys_profile_backups = Get-ChildItem "Registry::$path_reg_profiles" -Recurse -Include *.bak

		if ($keys_profile_backups.Count -gt 0) {
			Write-Host ($keys_profile_backups -join "`n")

			Write-Host "`n`n`nDelete temp profiles and restore the original keys:`n"
			ForEach($key_profile_backup in $keys_profile_backups) {
				$key_profile_temp = $key_profile_backup.Name.Split('.bak')[0]
				Remove-Item "Registry::$key_profile_temp"
				Write-Host $key_profile_temp

				$name_profile_original = $key_profile_temp.Split('\')[-1]
				Rename-Item "Registry::$key_profile_backup" -NewName $name_profile_original
			}
		} else {
			Write-Host "There are no broken profile keys !"
		}
	}
	Catch [System.Management.Automation.ItemNotFoundException] {
		Write-Host "Registry Key missing`n"
		Write-Host $path_reg_profiles
	}
	Finally { $ErrorActionPreference = 'Continue' }
}


Main