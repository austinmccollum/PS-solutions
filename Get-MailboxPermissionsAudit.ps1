        # User data from Get-Mailbox object
        $UserMBX = Get-Mailbox $User.DistinguishedName | Select GrantSendonBehalfTo,RecipientTypeDetails,DisplayName,Office
        If (-not $UserMBX) { Write-ErrorLog "Unable to process user.  Get-Mailbox $($User.DistinguishedName) returned zero results."; continue }
        ElseIf ($UserMBX.count -gt 1) { Write-ErrorLog "Unable to process user.  Get-Mailbox $($User.DistinguishedName) returned multiple ($($UserMBX.count)) results."; continue }
        $UserRecipientType = @{Expression={$UserMBX.RecipientTypeDetails};Label="UserRecipientType"}
        $UserDisplayName = @{Expression={$UserMBX.DisplayName};Label="UserDisplayName"}
        $UserOffice = @{Expression={$UserMBX.Office};Label="UserOffice"}
       
        # User data from Get-MailboxStatistics results
        $MailboxStats = Get-MailboxStatistics $User.DistinguishedName | select TotalItemSize,TotalDeletedItemSize
        If (-not $MailboxStats) { Write-ErrorLog "Unable to process user.  Get-MailboxStatistics $($User.DistinguishedName) returned zero results."; continue }
        ElseIf ($MailboxStats.count -gt 1) { Write-ErrorLog "Unable to process user.  Get-MailboxStatistics $($User.DistinguishedName) returned multiple ($($MailboxStats.count)) results."; continue }
        $MailboxSize = @{Expression={$MailboxStats.TotalItemSize.Value.ToMB()+$MailboxStats.TotalDeletedItemSize.Value.ToMB()};Label="UserTotalMailboxSize"}   

        # Folder level permissions
        $MailboxDN = $User.DistinguishedName
        If ($FolderAccess -eq "DefaultFolders") { $Folders = Get-MailboxFolderStatistics $MailboxDN | Where {$_.FolderType -ne "User Created"} | % {$_.FolderID} }
        ElseIf ($FolderAccess -eq "CalendarOnly") { $Folders = Get-MailboxFolderStatistics $MailboxDN | Where {$_.FolderType -eq "Calendar"} | % {$_.FolderID} }
        Else { $Folders = Get-MailboxFolderStatistics $MailboxDN | % {$_.FolderID} }

        ForEach ($Folder in $Folders) {
            $FolderKey = $User.mail + ":" + $Folder
            $FolderPermissions = Get-MailboxFolderPermission -Identity $FolderKey -ErrorAction SilentlyContinue | Where {$_.User -notlike "Default" -and $_.User -notlike "Anonymous" -and $_.User -notlike "*S-1-5-21*" -and $_.AccessRights -notlike "None" }
            If ($FolderPermissions) { # PS v2.0 iterates emtpy arrays, this was fixed in 3.0
                Foreach ($FolderPerm in $FolderPermissions) {
					# removed line because not returning accurate results
                    #$DelegateName = $FolderPerm.Identity.ToString().Replace("NT User:","")
                    $Recipient = Get-Recipient $FolderPerm.user.displayname -ErrorAction SilentlyContinue
                    If (-not $Recipient) { Write-ErrorLog "Unable to process delegate.  Could not find recipient: $DelegateName" }
                    ElseIf ($Recipient.Count -gt 1) { Write-ErrorLog "Unable to process delegate.  Duplicate recipients found for $DelegateName" }
                    ElseIf (($Recipient.RecipientTypeDetails -match "Mailbox" -and $Recipient.RecipientTypeDetails -notlike "*Remote*") -or $Recipient.RecipientTypeDetails -match "Group") {
                        $Add = @()
                        $Add = Analyze-Recipient $Recipient $($FolderPerm.AccessRights) $($FolderPerm.FolderName)
                        If ($Add) {
                            $HasDelegate = $true
                            $Script:Permissions += $Add | Select $Emailaddress,$UserUPN,$UserFirstName,$UserLastName,$UserDisplayName,$MailboxSize,$UserRecipientType,$UserOffice,Delegate*,FolderName,AccessRights,GroupDependencies
                        }
                    } Else {
                        Write-ErrorLog "Unable to process delegate.  $DelegateNane is of recipient type $($Recipient.RecipientTypeDetails), which is not supported by this script."
                    }
                }
            }
        }
        # End of folder level permissions

        # Send on behalf permissions
        $Delegates = $UserMBX.GrantSendonBehalfTo.ToArray()
        If ($Delegates) { # PS v2.0 iterates emtpy arrays, this was fixed in 3.0
            Foreach ($SOBDelegate in $Delegates) {
                #$SOBDelegateName = $SOBDelegate.Name.ToString().Replace("NT User:","")
                $Recipient = Get-Recipient $SOBDelegate -ErrorAction SilentlyContinue
                If (-not $Recipient) { Write-ErrorLog "Unable to process delegate.  Could not find recipient: $SOBDelegateName" }
                ElseIf ($Recipient.Count -gt 1) { Write-ErrorLog "Unable to process delegate.  Duplicate recipients found for $SOBDelegateName" }
                ElseIf (($Recipient.RecipientTypeDetails -match "Mailbox" -and $Recipient.RecipientTypeDetails -notlike "*Remote*") -or $Recipient.RecipientTypeDetails -match "Group") {
                    $Add = @()
                    $Add = Analyze-Recipient $Recipient "SendOnBehalf" ""
                    If ($Add) {
                        $HasDelegate = $true
                        $Script:Permissions += $Add | Select $Emailaddress,$UserUPN,$UserFirstName,$UserLastName,$UserDisplayName,$MailboxSize,$UserRecipientType,$UserOffice,Delegate*,FolderName,AccessRights,GroupDependencies
                    }
                } Else {
                    Write-ErrorLog "Unable to process delegate.  $SOBDelegateName is of recipient type $($Recipient.RecipientTypeDetails), which is not supported by this script."
                }
            }
        }
        # End of send on behalf permissions


Function Get-DelegateData ($Delegate) {
    # Delegate details based on the Get-Recipient object
    $DelegateDisplayName = @{Expression={$Delegate.DisplayName};Label="DelegateDisplayName"}
    $DelegateOffice = @{Expression={$Delegate.Office};Label="DelegateOffice"}
    $DelegatePrimarySMTP = @{Expression={$Delegate.PrimarySMTPAddress.tostring()};Label="DelegatePrimarySMTP"}
    $DelegateRecipientType = @{Expression={$Delegate.RecipientTypeDetails};Label="DelegateRecipientType"}

    # Delegate details based on the Get-ADUser object
    $DelUserObj = Get-ADUser $Delegate.distinguishedname -Properties UserPrincipalName,GivenName,Surname | Select UserPrincipalName,GivenName,Surname
    If (-not $DelUserObj) { Write-ErrorLog "Unable to process delegate.  Get-ADUser $($Delegate.distinguishedname) returned zero results."; return $null }
    ElseIf ($DelUserObj.count -gt 1) { Write-ErrorLog "Unable to process delegate.  Get-ADUser $($Delegate.distinguishedname) returned $($DelUserObj.count) results."; return $null }
    $DelegateFirstName = @{Expression={$DelUserObj.givenname};Label="DelegateFirstName"}
    $DelegateLastName = @{Expression={$DelUserObj.surname};Label="DelegateLastName"}
    $DelegateUPN = @{Expression={$DelUserObj.UserPrincipalName};Label="DelegateUPN"}

    # Delegate details based on the Get-MailboxStatistics results
    $DelStats = Get-MailboxStatistics $Delegate.DistinguishedName | Select TotalItemSize,TotalDeletedItemSize
    If (-not $DelStats) { Write-ErrorLog "Unable to process delegate.  Get-MailboxStatistics $($Delegate.distinguishedname) returned zero results."; return $null }
    ElseIf ($DelStats.count -gt 1) { Write-ErrorLog "Unable to process delegate.  Get-MailboxStatistics $($Delegate.distinguishedname) returned $($DelStats.count) results."; return $null }
    $DelegateSize = @{Expression={$DelStats.TotalItemSize.Value.ToMB()+$DelStats.TotalDeletedItemSize.Value.ToMB()};Label="DelegateTotalMailboxSize"}
    
    # Compile the data into an object and return it
    return "" | Select $DelegatePrimarySMTP,$DelegateUPN,$DelegateFirstName,$DelegateLastName,$DelegateDisplayName,$DelegateSize,$DelegateRecipientType,$DelegateOffice
}

Function Analyze-Recipient ($Delegate,$FRights,$FName) { # Function used to determine if the delegate from a permission line is the appropriate recipient type and grabs all the appropriate data if they are
    $FolderName = @{Expression={$FName};Label="FolderName"}
    $AccessRights = @{Expression={$FRights};Label="AccessRights"}
    If($Delegate.RecipientTypeDetails -match "Mailbox" -and $Delegate.RecipientTypeDetails -notlike "*Remote*") { # Check if user is a UserMailbox, SharedMailbox, or ResourceMailbox
        # Compile the delegate data from Get-Recipient, Get-ADUser, and Get-MailboxStatistics commands
        $ParentGroup = @{Expression={""};Label="GroupDependencies"}
        $DelegateData = Get-DelegateData $Delegate

        return $DelegateData | Select *,$FolderName,$AccessRights,$ParentGroup
    } ElseIf ($Delegate.RecipientTypeDetails -match "Group") { # Check for group
        $Delegates = @()
        $ParentGroup = @{Expression={"Member of Group: $($Delegate.DistinguishedName)"};Label="GroupDependencies"}

        If ($Script:GroupCache.keys -contains $Delegate.Name) { # Check if this group has already been encountered
            $Delegates = $Script:GroupCache["$($Delegate.Name)"]
        } Else {
            # Enumerate the group, grab the details, and cache it for future iterations
            $Members = New-Object System.Collections.ArrayList
            $DSLookFor.Filter = "(&(memberof:1.2.840.113556.1.4.1941:=$($Delegate.DistinguishedName))(objectCategory=user))"
            $DSLookFor.PageSize  = 1000
            $DSLookFor.SearchScope = "subtree"
            $Mail = $DSLookFor.PropertiesToLoad.Add("mail")
            $GroupMembers = $DSLookFor.findall()
            ForEach ($GroupMember in $GroupMembers) {		   
                # Add the member if the mail attribute is populated
                $GroupMember.Properties["mail"] | %{ If ($_) { [void]$Members.Add($_) } }
            }
            
            # Check each group member to see if they are the appropriate recipient type and grab their data if they are
            Foreach ($Member in $Members) {
				# added necessary property to Select to populate permissions csv for delegateprimarySMTP
                $Recipient = Get-Recipient $Member | Select PrimarySMTPAddress,RecipientTypeDetails,DistinguishedName,DisplayName,Office
                If (-not $Recipient) { Write-ErrorLog "Unable to process delegate.  Get-Recipient $Member returned zero results."; continue }
                ElseIf ($Recipient.count -gt 1) { Write-ErrorLog "Unable to process delegate.  Get-Recipient $Member returned $($Recipient.count) results."; continue }
                $RecipientType = $Recipient.RecipientTypeDetails
                If ($RecipientType -match "Mailbox" -and $RecipientType -notlike "*Remote*") {
                    # Compile the delegate date from Get-Recipient, Get-ADUser, and Get-MailboxStatistics commands
                    $DelegateData = Get-DelegateData $Recipient

                    $Delegates += $DelegateData | Select *
                } ElseIf ($RecipientType -match "Group") {
                    Write-ErrorLog "Unable to process delegate.  Group $Member is a member of group $($Delegate.PrimarySMTPAddress).  This script does not currently support nested groups."
                } Else {
                    Write-ErrorLog "Unable to process delegate.  $Member is of recipient type $RecipientType, which is not supported by this script."
                }
            }

            # Store this group with its members and data to speed up the process if this group is encountered again later in the script
            $Script:GroupCache.Add("$($Delegate.Name)",$Delegates)
        }
        
        return $Delegates | Select *,$FolderName,$AccessRights,$ParentGroup
    }
}

	# austinmc added quick logic to export just the unique relationships found outside the group originally specified
    Write-Action "Finding delegates not in the list proposed mailboxes to move..."
    $delegatesListSMTP=@()
    Foreach ($line in $Script:Permissions | Where {$_.DelegatePrimarySMTP -ne "n/a"}) {

	# austinmnc updated check to validate proper object property in the userlist array 
	if ($Script:Userlist.mail -notcontains $line.DelegatePrimarySMTP) {
		$delegatesListSMTP += $line
	}
    }	
    $delegatesListSMTP | Export-CSV ".\$delegateoutput" -NoTypeInformation