
# ====================================
# FUNCTION Outlook_signature_control
# ====================================
function outlook_signature_control{
	param()

	write-host "- Checking Outlook signature configuration."

	$hkcu_shellfolders	= 'Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\shell Folders'
	$general_key		= 'Registry::HKCU\Software\Microsoft\Office\%VERSION%\Common\General'
	$mail_key		= 'Registry::HKCU\Software\Microsoft\Office\%VERSION%\Common\MailSettings'
	$policy_key		= 'Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System'
	$profile		= 'Registry::HKCU\Software\Microsoft\Office\%VERSION%\Outlook\Profiles'
	$profile2		= 'Registry::HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
	$userappdata		= (get-itemproperty -path $reg_path.hkcu_shellfolders).'Appdata'

	$sign_template_path	= '\\File_server\netlogon\templates\signatures\' + $env:userdomain
	$def_template_path	= '\\File_server\netlogon\templates\signatures\default'
	$sign_folder_name	= 'Signatures'
	$sign_file_name		= 'sign'
	$user_sign_folder	= $env_path.userappdata + '\Microsoft\' + $sign_folder_name
	$sign_template_files	= ($sign_file_name + '.htm'),($sign_file_name + '.rtf'),($sign_file_name + '.txt')
	$mso_versions		= '11.0','12.0','14.0','15.0','16.0'
	$mail_client		= (get-itemproperty -path 'Registry::HKLM\SOFTWARE\Clients\Mail').'(default)'
	$apply_signature	= $false
	$template_fail		= $false
	$sign_exception		= $false


	if(!(test-path -path $sign_template_path)){$sign_template_path	= $def_template_path}

	foreach($check_file in $sign_template_files){
		if(!(test-path -path ($sign_template_path + '\' + $check_file))){
			$template_fail = $true
			break
		}
	}
	foreach($mso_version in $mso_versions){
		$outlook_profile = $profile
		$outlook_profile = $outlook_profile -replace '%version%',$mso_version
		if(test-path -path "Registry::$outlook_profile\Outlook"){
			$mail_client = 'Microsoft Outlook'
			break
		}
		if(test-path -path "registry::$outlook_profile\mail"){
			$mail_client = 'Microsoft Outlook'
			break
		}
	}
	if(test-path -path "registry::$profile2\Outlook\Setup"){$mail_client = 'Microsoft Outlook'}
	if(test-path -path "registry::$profile2\mail"){$mail_client = 'Microsoft Outlook'}
	if(($mail_client -eq 'Microsoft Outlook') -and ($template_fail -eq $false)){
		$adsi_search = "CN=Users"
		$obj_adsi = [DirectoryServices.DirectoryEntry]""
		$dn = $obj_adsi.distinguishedName
		if($dn){
			$query = [DirectoryServices.DirectoryEntry]"LDAP://$adsi_search,$dn"
			$adinfo			= @{}

			$adinfo.name		= $query.PSBase.Children | where-object{$_.name -eq $info.user} | select-object -expandproperty Name
			$adinfo.fullname	= $query.PSBase.Children | where-object{$_.name -eq $info.user} | select-object -expandproperty displayName
			$adinfo.title		= $query.PSBase.Children | where-object{$_.name -eq $info.user} | select-object -expandproperty description
			$adinfo.tell		= $query.PSBase.Children | where-object{$_.name -eq $info.user} | select-object -expandproperty telephoneNumber
			$adinfo.mobile		= $query.PSBase.Children | where-object{$_.name -eq $info.user} | select-object -expandproperty mobile
			$adinfo.email		= $query.PSBase.Children | where-object{$_.name -eq $info.user} | select-object -expandproperty email
		}

		# -------------------------------------------------------------------------
		# After we get the user information, start make the user default template
		# -------------------------------------------------------------------------
		if($adinfo.name){
			foreach($file in $sign_template_files){
#				if(($file -match '\.rtf$') -and ($adinfo.title -match '\<br\>')){
#					$adinfo.title = $adinfo.title -replace '\<br\>',"\par`n"
#				}
#				if(($file -match '\.txt$') -and ($adinfo.title -match '\<br\>')){
#					$adinfo.title = $adinfo.title -replace '\<br\>',"`n"#
#				}
#				if(($file -match '\.rtf$') -and ($adinfo.tell -match '\<br\>')){
#					$adinfo.tell = $adinfo.tell -replace '\<br\>',"\par`n"
#				}
#				if(($file -match '\.txt$') -and ($adinfo.tell -match '\<br\>')){
#					$adinfo.tell = $adinfo.tell -replace '\<br\>',"`n"
#				}
#				if(($file -match '\.rtf$') -and ($adinfo.mobile -match '\<br\>')){
#					$adinfo.mobile = $adinfo.mobile -replace '\<br\>',"\par`n"
#				}
#				if(($file -match '\.txt$') -and ($adinfo.mobile -match '\<br\>')){
#					$adinfo.mobile = $adinfo.mobile -replace '\<br\>',"`n"
#				}

				$sign_template = $sign_template_path + '\' + $file
				$tmp_record = get-content $sign_template
				$x = 0
				while($x -le ($tmp_record.length - 1)){
					if($tmp_record[$x] -cmatch '\%FULLNAME\%'){$tmp_record[$x] = $tmp_record[$x] -replace '\%FULLNAME\%',$adinfo.fullname}
					if($tmp_record[$x] -cmatch '\%JOBTITLE\%'){
						if($adinfo.title){
							$tmp_record[$x] = $tmp_record[$x] -replace '\%JOBTITLE\%',$adinfo.title
						}else{
							$tmp_record[$x] = $null
						}
					}
					if($tmp_record[$x] -cmatch '\%PHONE\%'){
						if($adinfo.tell){
							$tmp_record[$x] = $tmp_record[$x] -replace '\%PHONE\%',$adinfo.tell
						}else{
							$tmp_record[$x] = $null
						}
					}
					if($tmp_record[$x] -cmatch '\%MOBILE\%'){
						if($adinfo.mobile){
							$tmp_record[$x] = $tmp_record[$x] -replace '\%MOBILE\%',$adinfo.mobile
						}else{
							$tmp_record[$x] = $null
						}
					}
					if($tmp_record[$x] -cmatch '\%EMAIL\%'){
						if($adinfo.mobile){
							$tmp_record[$x] = $tmp_record[$x] -replace '\%EMAIL\%',$adinfo.email
						}else{
							$tmp_record[$x] = $null
						}
					}
					if($tmp_record[$x] -cmatch '\%USERNAME\%'){$tmp_record[$x] = $tmp_record[$x] -replace '\%USERNAME\%',$info.user}
					$x++
				}
				switch($file){
					{$_.Contains('.htm')}{$htm = $tmp_record | select-object}

					{$_.Contains('.rtf')}{$rtf = $tmp_record | select-object}

					{$_.Contains('.txt')}{$txt = $tmp_record | select-object}

					default{}
				}
				$tmp_record = @()
			}
			# ---------------------------------------------------------------------
			# Start to check whether need to replace with the signature template
			# ---------------------------------------------------------------------
			if(!(test-path -path $user_sign_folder)){
				$apply_signature = $true
			}else{
				foreach($sign_file in $sign_template_files){
					if($apply_signature -eq $true){break}
					if(test-path -path ($user_sign_folder + '\' + $sign_file)){
						$user_sign = get-content ($user_sign_folder + '\' + $sign_file)
						$user_sign = [System.Collections.ArrayList]$user_sign


						switch($sign_file){
							{$_.Contains('.htm')}{$user_template = $htm | select-object}

							{$_.Contains('.rtf')}{$user_template = $rtf | select-object}

							{$_.Contains('.txt')}{$user_template = $txt | select-object}

							default{}
						}
						foreach($content in $user_template){
							if(($content -eq $null) -or ($content -eq '')){
								continue
							}else{
								$check_content = &ascan -list $user_sign -search $content
								if($check_content -eq -1){
									$apply_signature = $true
									break
								}
							}
						}
						foreach($content in $user_sign){
							if(($content -eq $null) -or ($content -eq '')){
								continue
							}else{
								$check_content = &ascan -list $user_template -search $content
								if($check_content -eq -1){
									$apply_signature = $true
									break
								}
							}
						}
					}else{
						$apply_signature = $true
					}
				}
			}

			if($apply_signature -eq $true){
				write-host '- Start prepare a new signature.'
				if(test-path -path $user_sign_folder){
					remove-item -Path $user_sign_folder -Recurse -Force -ErrorAction silentlycontinue
				}
				if(!(test-path -path $user_sign_folder)){
					new-item -Path $user_sign_folder -ItemType Directory -Force -ErrorAction silentlycontinue
				}
				copy-item -Path ($sign_template_path + '\sign_files') -Destination $user_sign_folder -Recurse -Force -Erroraction silentlycontinue

				if(! $?){
					write-warning 'ERROR : Fail to copy Outlook signature images.'
				}else{
					foreach($sign_file in $sign_template_files){
						switch($sign_file){
							{$_.Contains('.htm')}{$htm | out-file ($user_sign_folder + '\' + $sign_file)}

							{$_.Contains('.rtf')}{$rtf | out-file ($user_sign_folder + '\' + $sign_file)}

							{$_.Contains('.txt')}{$txt | out-file ($user_sign_folder + '\' + $sign_file)}
							
							default{}
						}
					}
				}
			}
		}
		# --------------------------------------------
		# Apply Outlook signature registry settings
		# --------------------------------------------
		write-host "- Double check registry settings for Outlook signature"

		foreach($mso_version in $mso_versions){
			$general_path = $general_key
			if($general_path -match '\%VERSION\%'){$general_path = $general_path -replace '\%VERSION\%',$mso_version}
			$mail_settings = $mail_key
			if($mail_settings -match '\%VERSION\%'){$mail_settings = $mail_settings -replace '\%VERSION\%',$mso_version}
			$profile_path = $profile
			if($profile_path -match '\%VERSION\%'){$profile_path = $profile_path -replace '\%VERSION\%',$mso_version}
			
			if(test-path -path "Registry::$general_path"){
				$exist_sign_folder = (get-itemproperty -path "Registry::$general_path").'Signatures'
				if($exist_sign_folder -ne $sign_folder_name){
					&write_reg -reg_key $general_path -reg_name 'Signatures' -reg_value $sign_folder_name
				}
			}
			if(test-path -path "Registry::$mail_settings"){
				$check_file = (get-itemproperty -path "registry::$mail_settings").'NewSignature'
				if($check_file -ne $sign_file_name){
					&write_reg -reg_key $mail_settings -reg_name 'NewSignature' -reg_value $sign_file_name -reg_type 'REG_EXPAND_SZ'
				}
				$check_file = (get-itemproperty -path "registry::$mail_settings").'ReplySignature'
				if($check_file -ne $sign_file_name){
					&write_reg -reg_key $mail_settings -reg_name 'ReplySignature' -reg_value $sign_file_name -reg_type 'REG_EXPAND_SZ'
				}
			}
			if(test-path -path "Registry::$profile_path\Outlook"){
				$profile_list = $profile_path + '\Outlook'
			}elseif(test-path -path "registry::$profile_path\mail"){
				$profile_list = $proifle_path + '\mail'
			}
			if($profile_list){
				$keys = @()
				$keys += resolve-path -path ('Registry::' + $profile_list + '\*') -ErrorAction silentlyContinue | select-object -expandproperty path
				foreach($key in $keys){
					if(test-path -path "$key\00000001"){
						$my_profile = $key
						break
					}
				}
				if($my_profile){
					$profile_keys = @()
					$profile_keys += resolve-path -path ($my_profile + '\*') -ErrorAction silentlyContinue | select-object -expandproperty path
					foreach($profile_key in $profile_keys){
						$sign_trigger = (get-itemproperty -path $profile_key).'Identity Eid'
						$check_sign_val = (get-itemproperty -path $profile_key).'New Signature'
						if($check_sign_val -and ($check_sign_val -is [byte[]])){
							$check_sign_val = [System.Text.Encoding]::UniCode.GetString($check_sign_val)
						}
						if(($sign_trigger) -and ($check_sign_val -ne $sign_file_name)){
							&write_reg -reg_key $profile_key -reg_name 'New Signature' -reg_value $sign_file_name -reg_type 'REG_BINARY'
							&write_reg -reg_key $profile_key -reg_name 'Reply-Forward Signature' -reg_value $sign_file_name -reg_type 'REG_BINARY'
						}
					}
				}
			}
		}
	}else{
		write-host "- User may not using Outlook because cannot define email client or cannot find the signature templates."
	}
}


# =============================
# FUNCTION - ascan
# =============================
function ascan{
        param(
                [parameter(mandatory=$true)][array]$list,
                [parameter(mandatory=$true)][string]$search
        )
        $ascan = -1
        $i = 0
        do{
                if($list[$i] -eq $search){
                        $ascan = $i
                        $i = $list.Count - 1
                }
                $i++
        }until($i -gt ($list.Count - 1))

        return $ascan
}


# =================================
# FUNCTION - write_reg
# =================================
function write_reg{
	param(
		[parameter(mandatory=$true)][string]$reg_key,
		[string]$reg_value,
		[string]$reg_name,
		[string]$reg_type
	)

	$outcome = 0

	switch($reg_type){
		'REG_SZ'	{$private:type_name = 'string'}
		'REG_EXPAND_SZ'	{$private:type_name = 'ExpandString'}
		'REG_BINARY'	{$private:type_name = 'Binary'}
		'REG_DWORD'	{$private:type_name = 'DWord'}
		'REG_MUTI_SZ'	{$private:type_name = 'MultiString'}
		'REG_QWORD'	{$private:type_name = 'QWord'}
		default		{
				$private:type_name = 'string'
				$reg_type = 'REG_SZ'
				}	
	}

	if($reg_key -imatch '^Registry\:\:.*'){
		$reg_key = $reg_key -replace 'Registry\:\:',''
	}
	

	if(! $reg_name){$reg_name = '(Default)'}
	if(!(test-path Registry::$reg_key)){new-item -path ('Registry::' + $reg_key) -Force -Erroraction silentlyContinue}
	if(test-path Registry::$reg_key){
		if($reg_type -eq 'REG_BINARY'){
			$value = [byte[]][System.Text.Encoding]::UniCode.GetBytes($reg_value)
		}else{
			$value = $reg_value
		}
		try{
			set-itemproperty -path ('Registry::' + $reg_key) -name $reg_name -value $value -type $type_name | out-null
		}
		catch{
			write-warning 'ERROR : Fail to right value :' $_
			$outcome = -1
		}
	}else{
		write-warning 'ERROR : Fail create new registry value due to registry key path fail to create.'
		$outcome = -1
	}
	return $outcome
}

