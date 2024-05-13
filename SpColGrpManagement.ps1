# Special Collection Group Management Script
#
# Only users in either BornDigitalSuper or DigitalVaultSuper groups are allowed to access this
# - Allow the supervisors to add and remove users from the specific groups for folder access
#

Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory

$msg = [System.Windows.Forms.MessageBox]
#$msg::Show("Please Enter Library Crendential")

$group=""
$borndigitalsuper=$digitalvaultsuper=$false
$cred = Get-Credential -Message "Special Collection Supervisors Only. Others will be prosecuted."
if ($cred) { 
	$aduser = Get-ADUser -Credential $cred -Properties memberof $cred.UserName
	if ($?) {
		#write-host $($aduser.memberof)
		if ($($aduser.memberof) -like '*BornDigitalSuper*') {
			$borndigitalsuper=$true
			#write-host "BornDigital Admin"
		}
		if ($($aduser.memberof) -like '*DigitalVaultSuper*') {
			$digitalvaultsuper=$true
			#write-host "DigitalVault Admin"
		}
		if (!($borndigitalsuper) -and !($digitalvaultsuper)) {
			$msg::Show("$($cred.UserName) Does Not Have Permission.")
			#Write-Host "User Not Found"
			exit
		}
	}
	else {
		$msg::Show("$($cred.UserName) Not Found or Wrong Crendential.")
		#Write-Host "User Not Found"
		exit
	}
}
else {
	#Write-Host "No Credential Entered"
	exit
}

# Function to add a member to the group
function Add-MemberToGroup ($groupname, $member) {
    Add-ADGroupMember -Credential $cred -Identity $groupname -Members $member
}

# Function to remove a member from the group
function Remove-MemberFromGroup ($groupname, $member) {
    Remove-ADGroupMember -Credential $cred -Identity $groupname -Members $member -Confirm:$false
}

function VerifyGroup ($groupname) {
	$verifyuser = $false
	$dscheck =dsacls.exe CN=$groupname,OU=Groups,DC=library,DC=unlv,DC=edu
	if ($?) {
		if ( ($dscheck -like '*BornDigitalSuper*') -and ($borndigitalsuper) ) {
			$verifyuser = $true
		}
		elseif ( ($dscheck -like '*DigitalVaultSuper*') -and ($digitalvaultsuper) ) {
			$verifyuser = $true
		}
		if (!($verifyuser)) {
			Return 0
		}
		else {
			Return 1
		}
	}
	else {
		Return 2
	}
}

# Create the form
$form = New-Object System.Windows.Forms.Form 
$form.Text = 'Library AD Group Management'
$form.Size = New-Object System.Drawing.Size(500,450) 
$form.StartPosition = 'CenterScreen'
$form.Add_Shown({
	if ($borndigitalsuper) {
		$RadioButton1.Visible=$RadioButton2.Visible=$RadioButton3.Visible=$RadioButton4.Visible=$RadioButton5.Visible=$true
	}
	else {
		$RadioButton1.Visible=$RadioButton2.Visible=$RadioButton3.Visible=$RadioButton4.Visible=$RadioButton5.Visible=$false
	}
	if ($digitalvaultsuper) {
		$RadioButton6.Visible=$RadioButton7.Visible=$RadioButton8.Visible=$RadioButton9.Visible=$true
	}
	else {
		$RadioButton6.Visible=$RadioButton7.Visible=$RadioButton8.Visible=$RadioButton9.Visible=$false
	}
	$MyGroupBox.Focus()
	$Form.Activate()
})

$MyGroupBox = New-Object System.Windows.Forms.GroupBox
$MyGroupBox.Location = '10,5'
$MyGroupBox.Size = '450,120'
$MyGroupBox.Text = "Available Groups"
$MyGroupBox.Controls.AddRange(@($GroupRadioButton1, $GroupRadioButton2, $GroupRadioButton3, $GroupRadioButton4, $GroupRadioButton5, $GroupRadioButton6, $GroupRadioButton7, $GroupRadioButton8, $GroupRadioButton9))
$form.Controls.Add($MyGroupBox)

# Create Radio Buttons
$RadioButton1 = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Text = "BornWorkWrite"
$RadioButton1.Location = '20,15'
$RadioButton1.Size = '120,20'
$RadioButton1.add_CheckedChanged({
	$membersBox.Visible = $false
	$group = $($RadioButton1.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton1)

$RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioButton2.Text = "BornWorkRead"
$RadioButton2.Location = '20,35'
$RadioButton2.Size = '120,20'
$RadioButton2.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton2.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton2)

$RadioButton3 = New-Object System.Windows.Forms.RadioButton
$RadioButton3.Text = "BornDigitalWrite"  # Customize the text for your third radio button
$RadioButton3.Location = '20,55'
$RadioButton3.Size = '120,20'
$RadioButton3.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton3.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton3)

$RadioButton4 = New-Object System.Windows.Forms.RadioButton
$RadioButton4.Text = "BornDigitalRead"
$RadioButton4.Location = '20,75'
$RadioButton4.Size = '120,20'
$RadioButton4.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton4.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton4)

$RadioButton5 = New-Object System.Windows.Forms.RadioButton
$RadioButton5.Text = "BornEncRead"
$RadioButton5.Location = '20,95'
$RadioButton5.Size = '120,20'
$RadioButton5.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton5.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton5)

$RadioButton6 = New-Object System.Windows.Forms.RadioButton
$RadioButton6.Text = "DigWorkWrite"
$RadioButton6.Location = '150,20'
$RadioButton6.Size = '120,20'
$RadioButton6.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton6.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton6)

$RadioButton7 = New-Object System.Windows.Forms.RadioButton
$RadioButton7.Text = "DigWorkRead"
$RadioButton7.Location = '150,40'
$RadioButton7.Size = '120,20'
$RadioButton7.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton7.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton7)

$RadioButton8 = New-Object System.Windows.Forms.RadioButton
$RadioButton8.Text = "DigitalVaultWrite"
$RadioButton8.Location = '150,60'
$RadioButton8.Size = '120,20'
$RadioButton8.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton8.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton8)

$RadioButton9 = New-Object System.Windows.Forms.RadioButton
$RadioButton9.Text = "DigitalVaultRead"
$RadioButton9.Location = '150,80'
$RadioButton9.Size = '120,20'
$RadioButton9.add_CheckedChanged({
	$membersBox.Visible=$false
	$group = $($RadioButton9.Text)
	$result = VerifyGroup($group)
	if ($result -eq 1) {
		$members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
		$membersBox.Text = $members -join "`r`n"
		$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible =$true
		$UserBox.Focus()
	}
	elseif ($result -eq 0) {
			$membersBox.Text = "No Permission to Group $group"
			$membersBox.Visible = $true
			$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
			$MyGroupBox.Focus()
	}
	else {
		$membersBox.Text = "Group $group Not Found"
		$membersBox.Visible = $true
		$Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
		$MyGroupBox.Focus()
	}
})
$MyGroupBox.Controls.Add($RadioButton9)

# Create the TextBox for group members
$membersBox = New-Object System.Windows.Forms.TextBox 
$membersBox.Location = New-Object System.Drawing.Point(10,120) 
$membersBox.Size = New-Object System.Drawing.Size(440,180) 
$membersBox.Multiline = $true
$membersBox.ScrollBars = 'Vertical'
$membersBox.ReadOnly = $true
$membersBox.Visible = $false
$form.Controls.Add($membersBox) 

# Username Box
$Userlabel = New-Object System.Windows.Forms.Label
$Userlabel.Location = New-Object System.Drawing.Point(10,320) 
$Userlabel.Size = New-Object System.Drawing.Size(60,20) 
$Userlabel.Text = 'Username:'
$Userlabel.Visible = $false
$form.Controls.Add($Userlabel) 

$UserBox = New-Object System.Windows.Forms.TextBox 
$UserBox.Location = New-Object System.Drawing.Point(70,320) 
$UserBox.Size = New-Object System.Drawing.Size(120,20) 
$UserBox.Visible = $false
$form.Controls.Add($UserBox) 

# Create the Add button
$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Point(200,320)
$AddButton.Size = New-Object System.Drawing.Size(75,23)
$AddButton.Text = 'Add'
$AddButton.Visible = $false
$AddButton.Add_Click({
	$group = $($MyGroupBox.Controls | Where-Object { $_.Checked }).Text
	#write-host "Adding $($UserBox.Text) from $group"
    Add-MemberToGroup -Credential $cred -group $group -member $UserBox.Text
    $members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
    $membersBox.Text = $members -join "`r`n"
	$UserBox.Clear()
	$UserBox.Focus()
})
$form.Controls.Add($AddButton)

# Create the Remove button
$RemoveButton = New-Object System.Windows.Forms.Button
$RemoveButton.Location = New-Object System.Drawing.Point(280,320)
$RemoveButton.Size = New-Object System.Drawing.Size(75,23)
$RemoveButton.Text = 'Remove'
$RemoveButton.Visible = $false
$RemoveButton.Add_Click({
	$group = $($MyGroupBox.Controls | Where-Object { $_.Checked }).Text
	#write-host "Removing $($UserBox.Text) from $group"
    Remove-MemberFromGroup -Credential $cred -group $group -member $UserBox.Text
    $members = Get-ADGroupMember -Credential $cred -Identity $group | ForEach-Object { $_.SamAccountName }
    $membersBox.Text = $members -join "`r`n"
	$UserBox.Clear()
	$UserBox.Focus()
})
$form.Controls.Add($RemoveButton)

# Create the Reset button
$ResetButton = New-Object System.Windows.Forms.Button
$ResetButton.Location = New-Object System.Drawing.Point(50,360)
$ResetButton.Size = New-Object System.Drawing.Size(75,23)
$ResetButton.Text = 'Reset'
$ResetButton.Add_Click({
    $UserBox.Clear()
    $membersBox.Clear()
	$membersBox.Visible = $Userlabel.Visible = $UserBox.Visible = $AddButton.Visible = $RemoveButton.Visible = $false
	$MyGroupBox.Focus()
})
$form.Controls.Add($ResetButton)

# Create the Quit button
$QuitButton = New-Object System.Windows.Forms.Button
$QuitButton.Location = New-Object System.Drawing.Point(150,360)
$QuitButton.Size = New-Object System.Drawing.Size(75,23)
$QuitButton.Text = 'Quit'
$QuitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($QuitButton)

# Set the CancelButton property to the Quit button
$form.CancelButton = $QuitButton

# Show the form
$form.ShowDialog()
