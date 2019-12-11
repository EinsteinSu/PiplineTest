#Copyright 2019 Quest Software Inc.  ALL RIGHTS RESERVED.

#region Pre-check AM installation

# Check if Archive Manager installed in this server and Add 'Quest.AM.BusinessLayer.dll' into current session.
Write-Verbose "Connecting to Archive Manager ...";

$dll = 'Quest.AM.BusinessLayer.dll';
$version = "";
$amInstallationPath = "";
If (Test-Path -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Quest\ArchiveManager\Installer) {
    $item = Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Quest\ArchiveManager\Installer;

    $amInstallationPath = $item.GetValue('InstallDirectory');
    $version = $item.GetValue('Version');
}
ElseIf (Test-Path -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Quest\ArchiveManager\Installer) {
    $item = Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Quest\ArchiveManager\Installer;

    $amInstallationPath = $item.GetValue('InstallDirectory');
    $version = $item.GetValue('Version');
}

$dllPath = [System.IO.Path]::Combine($amInstallationPath, $dll);

If (Test-Path -Path $dllPath) {
    Add-Type -Path $dllPath;

    Write-Host "Archive Manager $version installed under: $amInstallationPath.`nConnected to it successfully!`n" -ForegroundColor Green;
}
Else {
    Throw "Archive Manager does not be installed, cmdlets in this module is un-usable.";
}

#endregion

#region Non-publish functions

<#
 .Synopsis
  Execute an SQL command in Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Execute an SQL command in Archive Manager database and return an database.

 .PARAMETER script
  The executed script.

 .PARAMETER parameters
  The parameters.
#>
Function Execute-SQLDataTable {
    PARAM(
    [Parameter(Mandatory=$true,
            Position=0)]
    [string]
    $script, 
    [System.Collections.IDictionary]
    $parameters = $null
    )
    $cmd = New-Object System.Data.SqlClient.SqlCommand;
    $cmd.CommandText = $script;
    $cmd.CommandType = [System.Data.CommandType]::Text;
    If ($parameters -ne $null) {
        $cmd.Parameters = New-Object System.Data.SqlClient.SqlParameterCollection;

        Foreach ($key in $parameters.Keys) {
            $cmd.Parameters.Add($key, $parameters[$key]);
        }
    }

    Return [Quest.AM.Configuration.Common.SqlHelper]::ExecuteDataTable($cmd);
}
#endregion

#region Cmdlets for Login

<#
 .Synopsis
  Get an Login object from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get an Login object from Archive Manager database. This function supports different parameters like Id, name, Uid, Sid, etc.

 .PARAMETER id
  The login Id. If the Id specified, it will fetch the login by the Id and other parameters will be ignored.

 .PARAMETER name
  The Login name. It will be used with Domain name or display name together. Only when the Id is not specified this parameter will be used.

 .PARAMETER domain
  The domain name of the Login. It musts use with parameter name together.

 .PARAMETER displayName
  The Login display Name. It musts use with parameter name together. Only when the domain name is not specified this parameter will be used.

 .PARAMETER uid
  The Login uid. Only when the parameter id and name not specified this parameter will be used.

 .PARAMETER sid
  The Login sid. Only when the parameter id, name and uid not specified this parameter will be used.

 .EXAMPLE
   Get an Login object by Id 1.
   Get-Login -id 1;
   Get-Login 1;

 .EXAMPLE
   Get an Login object by name and domain.
   Get-Login -name 'Admin' -domain 'DEFAULT';

 .EXAMPLE
   Get an Login object by name and displayname.
   Get-Login -name 'Admin' -displayName 'Archive Manager Administrator';

 .EXAMPLE
   Get an Login object by uid.
   Get-Login -uid '3595958e-8ba6-425c-aec6-7bfad0ed1f1a';

 .EXAMPLE
   Get an Login object by sid.
   Get-Login -sid 'SID-3595958e-8ba6-425c-aec6-7bfad0ed1f1a';
#>
Function Get-Login {
    PARAM(
    [Parameter(Position=0)]
    [int]
    $id = 0,
    [string]
    $name = [string]::Empty,
    [string]
    $domain = [string]::Empty,
    [string]
    $displayName = [string]::Empty,
    [string]
    $uid = [string]::Empty,
    [string]
    $sid = [string]::Empty
    )

    If ($id -gt 0) {
        Write-Host "Get Login by Login Id: $id`n" -ForegroundColor Green;

        Return [Quest.AM.BusinessLayer.Login]::FindLoginId($id);
    }

    If (![string]::IsNullOrEmpty($name)) {
        If (![string]::IsNullOrEmpty($domain)) {
            Write-Host "Get Login by Login Name: $name and Domain Name: $domain.`n" -ForegroundColor Green;

            Return [Quest.AM.BusinessLayer.Login]::FindLoginByNameAndDomainName($domain, $name);
        }

        If (![string]::IsNullOrEmpty($displayName))
        {
            Write-Host "Get Login by Login Name: $name and Display Name: $displayName.`n" -ForegroundColor Green;

            Return [Quest.AM.BusinessLayer.Login]::FindLoginNameDisplayName($name, $displayName);
        }

        Write-Host "Please provide domain name or display name with login name.`n" -ForegroundColor Red;
        Return $null;
    }

    If (![string]::IsNullOrEmpty($uid)) {
        Write-Host "Get Login by Uid: $uid`n" -ForegroundColor Green;

        Return [Quest.AM.BusinessLayer.Login]::FindUid($uid);
    }

    If (![string]::IsNullOrEmpty($sid)) {
        Write-Host "Get Login by Sid: $sid`n" -ForegroundColor Green;

        Return [Quest.AM.BusinessLayer.Login]::FindSid($sid);
    }

    Write-Host 'No parameter specified. Return nothing.`n' -ForegroundColor Red;
    Return $null;
}

<#
 .Synopsis
  Get the Logins count from Archive Manager database.

 .DESCRIPTION
  Get the Logins count from Archive Manager database. This function will return the number of Logins in Archive Manager database.

 .EXAMPLE
   Get Logins count in Archive Manager database.
   Get-LoginsCount;
#>
Function Get-LoginsCount {
    Write-Host "Get Logins Count`n" -ForegroundColor Green;

    Return [Quest.AM.BusinessLayer.Login]::LoginsCount();
}

<#
 .Synopsis
  Get all of the Logins from Archive Manager database.

 .DESCRIPTION
  Get all of the Logins from Archive Manager database. This function will return all of the Login objects in Archive Manager database.

 .PARAMETER name
  The login name. Can use whildchar '*'.

 .PARAMETER displayName
  The login display name. Can use whildchar '*'.

 .PARAMETER securityRole
  The login securityRole, has 4 valid values: Administrator, Manager, Resource, User.

 .PARAMETER emailAddress
  The login email Address. Can use whildchar '*'.

 .PARAMETER domainName
  The login's Domain name. It will fetch all Logins which are related to a specified login domain name. Can use whildchar '*'.

 .PARAMETER active
  Fetch all logins by it's active property. Bt default is NULL, represent all login without care about the property.

 .PARAMETER groupId
  Fetch all logins belong to a specified group.

 .EXAMPLE
   Get Logins in Archive Manager database.
   Get-Logins;

 .EXAMPLE
   Get Logins which start with name 'jshi' in Archive Manager database.
   Get-Logins -name 'jshi*';
   Get-Logins 'jshi*';

 .EXAMPLE
   Get Logins which belong to security role 'User' in Archive Manager database.
   Get-Logins -securityRole 'User';

 .EXAMPLE
   Get Logins which belong to group 1 in Archive Manager database.
   Get-Logins -groupId 1;

 .EXAMPLE
   Get Logins which start with 'jshi', and belong to domain 'DEFAULT', and email address start with 'jshi' and inactive in Archive Manager database.
   Get-Logins -name 'jshi*' -domainName 'DEFAULT' -emailAddress 'jshi*' -active $false;
   Get-Logins 'jshi*' -domainName 'DEFAULT' -emailAddress 'jshi*' -active $false;
#>
Function Get-Logins {
    PARAM(
    [PARAMETER(Position=0)]
    [string]
    $name = [string]::Empty,
    [string]
    $displayName = [string]::Empty,
    [ValidateSet("Administrator", "Manager", "Resource", "User", "")]
    [string]
    $securityRole = [string]::Empty,
    [string]
    $emailAddress = [string]::Empty,
    [string]
    $domainName = [string]::Empty,
    [nullable[bool]]
    $active = $null,
    [int]
    $groupId = 0
    )

    $criteria = New-Object Quest.AM.BusinessLayer.Queries.LoginsSearchQueryObject;

    If (![string]::IsNullOrEmpty($name)) {
        $criteria.LoginName = $name;
    }
    If (![string]::IsNullOrEmpty($displayName)) {
        $criteria.DisplayName = $displayName;
    }
    If (![string]::IsNullOrEmpty($securityRole)) {
        $criteria.SecurityRole = $securityRole;
    }
    If (![string]::IsNullOrEmpty($emailAddress)) {
        $criteria.EmailAddress = $emailAddress;
    }
    If (![string]::IsNullOrEmpty($domainName)) {
        $criteria.DomainName = $domainName;
    }
    If ($active -ne $null) {
        $criteria.Active = $active;
    }
    If ($groupId -gt 0) {
        $criteria.GroupId = $groupId;
    }

    Write-Host "Get all of the Logins from Archive Manager by criteria: $criteria`n" -ForegroundColor Green;

    $login = New-Object Quest.AM.BusinessLayer.Login;
    Return $login.SelectAll($criteria);
}

#endregion

#region Cmdlets for MailBox

<#
 .Synopsis
  Get an MailBox object from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get an MailBox object from Archive Manager database. This function supports different parameters like Id, name, Uid, legacyDN, etc.

 .PARAMETER id
  The MailBox Id. If the Id specified, it will fetch the MailBox by the Id and other parameters will be ignored.

 .PARAMETER name
  The MailBox name. Only when the Id is not specified this parameter will be used.

 .PARAMETER uid
  The MailBox uid. Only when the parameter id and name not specified this parameter will be used.

 .PARAMETER legacyDN
  The Login Legacy Exchange DN. Only when the parameter id, name and uid not specified this parameter will be used.

 .EXAMPLE
   Get an MailBox object by Id 1.
   Get-MailBox -id 1;
   Get-MailBox 1;

 .EXAMPLE
   Get an MailBox object by name.
   Get-MailBox -name 'LidiaH';

 .EXAMPLE
   Get an MailBox object by uid.
   Get-MailBox -uid '3595958e-8ba6-425c-aec6-7bfad0ed1f1a';

 .EXAMPLE
   Get an MailBox object by legacy Exchange DN.
   Get-MailBox -legacyDN '/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=c82539136d6b4af19cc18ca19b598dd3-LidiaH';
#>
Function Get-MailBox {
    PARAM(
    [PARAMETER(Position=0)]
    [int]
    $id = 0,
    [string]
    $name = [string]::Empty,
    [string]
    $uid = [string]::Empty,
    [string]
    $legacyDN = [string]::Empty
    )

    $mbx = New-Object Quest.AM.BusinessLayer.MailBox;
    If ($id -gt 0) {
        Write-Host "Get MailBox by MailBox Id: $id`n" -ForegroundColor Green;

        If ($mbx.Select($id)) {
            Return $mbx;
        }
        Else { 
            Return $null;
        }
    }

    If (![string]::IsNullOrEmpty($name)) {
        Write-Host "Get MailBox by MailBox Name: $name`n" -ForegroundColor Green;

        If ($mbx.SelectByName($name, $true)) {
            Return $mbx;
        }
        Else {
            Return $null;
        }
    }

    If (![string]::IsNullOrEmpty($uid)) {
        Write-Host "Get MailBox by MailBox Uid: $uid`n" -ForegroundColor Green;

        If ($mbx.SelectByUID($uid)) {
            Return $mbx;
        }
        Else {
            Return $null;
        }
    }

    If (![string]::IsNullOrEmpty($legacyDN)) {
        Write-Host "Get Box by MailBox legacy Exchang DN: $legacyDN`n" -ForegroundColor Green;

        If ($mbx.SelectByLegacyDN($legacyDN)) {
            Return $mbx;
        }
        Else {
            Return $null;
        }
    }

    Write-Host "No parameter specified. Return nothing.`n" -ForegroundColor Red;
    Return $null;
}

<#
 .Synopsis
  Get an MailBoxes Count from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get the MailBoxes Count from Archive Manager database. This function supports parameter loginId.

 .PARAMETER loginId
  The login Id. If the Id specified, it will fetch the MailBoxes count owned by the login Id.

 .EXAMPLE
   Get all of the mailboxes count from Archive Manager database.
   Get-MailBoxesCount;

 .EXAMPLE
   Get MailBoxes count which are owned by a specified login Id.
   Get-MailBoxesCount -loginId 1;
   Get-MailBoxesCount 1;
#>
Function Get-MailBoxesCount {
    PARAM(
    [PARAMETER(Position=0)]
    [int]
    $loginId = 0
    )

    If ($loginId -gt 0) {
        Write-Host "Get MailBox Count by login Id: $loginId`n" -ForegroundColor Green;

        $mbxIds = [Quest.AM.BusinessLayer.MailBox]::GetAllowedMailBoxIDs($loginId);
        Return $mbxIds.Length;
    }
    Else {
        Write-Host "Get MailBox Count" -ForegroundColor Green;

        Return [Quest.AM.BusinessLayer.MailBox]::MailboxesCount();
    }
}

<#
 .Synopsis
  Get MailBoxes objects from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get MailBoxes objectes from Archive Manager database. This function supports different parameters like loginId, notLoginId, name, loginDomainName, etc.

 .PARAMETER loginId
  The login Id. If the Id specified, it will fetch all MailBoxes owned by the loginId.

 .PARAMETER notLoginId
  The login Id. If the Id specified, it will fetch all MailBoxes which are not owned by the loginId.

 .PARAMETER name
  The MailBox name. Can use whildchar '*'.

 .PARAMETER loginDomainName
  The MailBox's login Domain name. It will fetch all MailBoxes which are related to a specified login domain name. Can use whildchar '*'.

 .PARAMETER mailServerId
  The MailBox's mailServer Id.

 .PARAMETER enableStoreManager
  Fetch all MailBoxes by it's EnableStoreManager property. Bt default is NULL, represent all mailboxes without care about the property.

 .PARAMETER deleted
  Fetch all MailBoxes by it's Deleted property. Bt default is NULL, represent all mailboxes without care about the property.

 .PARAMETER messageId
  Fetch all MailBoxes which related with the specified messageId. If this parameter is specified, all other parameters will be ignored.

 .EXAMPLE
   Get all of the MailBoxes in Archive Manager database.
   Get-MailBoxes;

 .EXAMPLE
   Get all of the MailBoxes whcih EnableStoreManager is True in Archive Manager database.
   Get-MailBoxes -enableStoreManager $true;

 .EXAMPLE
   Get all of the MailBoxes which are belong to Login 1, and MailServer Id 1 and are not deleted.
   Get-MailBoxes -loginId 1 -mailserverId 1 -Deleted $false;

 .EXAMPLE
   Get all of the MailBoxes which name starts with 'jshi'.
   Get-MailBoxes -name jshi*;

 .EXAMPLE
   Get all of the MailBoxes with specified a message Id 1.
   Get-MailBoxes -messageId 1;
#>
Function Get-MailBoxes {
    PARAM(
    [PARAMETER(Position=0)]
    [int]
    $loginId = 0,
    [int]
    $notLoginId = 0,
    [string]
    $name = [string]::Empty,
    [string]
    $loginDomainName = [string]::Empty,
    [int]
    $mailserverId = 0,
    [nullable[bool]]
    $enableStoreManager = $null,
    [nullable[bool]]
    $deleted = $null,
    [int]
    $messageId = 0
    )

    If (!($messageId -gt 0)) {
        $criteria = New-Object Quest.AM.BusinessLayer.Queries.MailboxesSearchQueryObject
        If ($loginId -gt 0) {
            $criteria.LoginId = $loginId;
        }
        If ($notLoginId -gt 0) {
            $criteria.NotLoginId = $notLoginId;
        }
        If (![string]::IsNullOrEmpty($name)) {
            $criteria.Name = $name;
        }
        If (![string]::IsNullOrEmpty($loginDomainName)) {
            $criteria.LoginDomainName = $loginDomainName;
        }
        If ($mailserverId -gt 0) {
            $criteria.MailServerId = $mailserverId;
        }
        If ($enableStoreManager -ne $null) {
            $criteria.EnableStoreManager = $enableStoreManager;
        }
        If ($deleted -ne $null) {
            $criteria.Deleted = $deleted;
        }

        Write-Host "Get MailBoxes by search criteria: $criteria`n" -ForegroundColor Green;

        $mbx = New-Object Quest.AM.BusinessLayer.MailBox;
        Return $mbx.SelectAll($critera);
    }
    Else {
        Write-Host "Get MailBoxes by message Id: $messageId`n" -ForegroundColor Green;

        $script = "SELECT [MailBoxID] FROM [dbo].[MailBoxMessage] WITH (NOLOCK) WHERE MessageID = $messageId";

        $dt = Execute-SQLDataTable $script;
        If ($dt -ne $null) {
            $mbxes = New-Object System.Collections.Generic.List[Quest.AM.BusinessLayer.MailBox];
            Foreach ($d in $dt) {
                $mbxes.Add([Quest.AM.BusinessLayer.MailBox]::FromMailBoxId([int]$d[0]));
            }

            Return $mbxes;
        }
        Else {
            Write-Host "No mailboxes related with message Id: $messageId. Return nothing.`n" -ForegroundColor Red;

            Return $null;
        }
    }
}

#endregion

#region Cmdlets for Message

<#
 .Synopsis
  Get a Message object from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get a Message object from Archive Manager database. This function supports 2 parameters: id and checkSum. 
  It returns an Quest.AM.BusinessLayer.Email2 object. This function will not do the permission check.

 .PARAMETER id
  The message Id. If the Id specified, it will fetch the message by the Id and other parameters will be ignored.

 .PARAMETER checksum
  The message Checksum. Only when the Id is not specified this parameter will be used.

 .EXAMPLE
   Get a Message object by Id 1.
   Get-Message -id 1;
   Get-Message 1;

 .EXAMPLE
   Get a Message object by checksum.
   Get-MailBox -checksum 'EA14C129-9AC7-29CD-854A-FD8461291FAF';
#>
Function Get-Message {
    PARAM(
    [PARAMETER(Position=0)]
    [int]
    $id = 0,
    [Guid]
    $checksum = [Guid]::Empty
    )

    $ds = $null;
    If ($id -gt 0) {
        Write-Host "Get message by Id: $id`n" -ForegroundColor Green;

        $ds = [Quest.AM.BusinessLayer.Message]::Select_Full_NoLoginCheck($id);
    }
    If ($checksum -ne [Guid]::Empty) {
        Write-Host "Get message by checksum: $checksum`n" -ForegroundColor Green;

        $ds = [Quest.AM.BusinessLayer.Message]::Select_NoLoginCheck($checksum);
    }

    If ($ds -ne $null) {
        Return New-Object Quest.AM.BusinessLayer.Email2 -ArgumentList @($ds);
    }

    Write-Host "No valid parameter specified. Return nothing.`n" -ForegroundColor Red;
    Return $null;
}

<#
 .Synopsis
  Get Messages Count from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get Messages count from Archive Manager database. This function supports 2 parameters: mailboxId and folderId. 

 .PARAMETER mailboxId
  The mailbox Id. If the Id specified, it will fetch the message count under the specified mailbox and other parameters will be ignored.

 .PARAMETER folderId
  The folder Id. It will fetch the message count under the specified folder.

 .EXAMPLE
   Get Messages count under mailbox Id 1.
   Get-MessagesCount -mailboxId 1;
   Get-MessagesCount 1;

 .EXAMPLE
   Get Messages count under folder Id 1.
   Get-MessagesCount -folderId 1;
#>
Function Get-MessagesCount {
    PARAM(
    [PARAMETER(Position=0)]
    [int]
    $mailboxId = 0,
    [int]
    $folderId = 0
    )

    If ($mailboxId -gt 0) {
        Write-Host "Get Messages count under mailbox Id: $mailboxId `n" -ForegroundColor Green;

        Return [Quest.AM.BusinessLayer.MailBox]::MessageCount($mailboxId);
    }
    If ($folderId -gt 0) {
        Write-Host "Get Messages count under folder Id: $folderId `n" -ForegroundColor Green;

        Return [Quest.AM.BusinessLayer.Folder]::MessageCount($folderId);
    }

    Write-Host "No valid parameter specified. Return nothing." -ForegroundColor Red;
    Return $null;
}

<#
 .Synopsis
  Get a list of Messages objects from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get a list of Messages objects from Archive Manager database. This function supports some parameters like loginId, subject, to, from, etc. 
  It returns a list of Quest.AM.Public.SDK.v1.Models.MessageSummary object.

 .PARAMETER loginId
  The login which used to perform the query. Id as the alias.

 .PARAMETER subject
  The message subject. Can be used with whildchar '*'. It will not do the message content search but only for the subject.

 .PARAMETER to
  The message to recipients. Split by ';'.

 .PARAMETER subject
  The message from. Split by ';'.

 .PARAMETER toOrFrom
  The message toOrFrom recipients. Split by ';'.

 .PARAMETER hasAttachments
  Whether the messages have attachments or not. By default is null which represents the query does not care about the attachments.

 .PARAMETER attachmentNames
  The specified attachment names. Only when the parameter hasAttachments is true the parameter will be applied. Split by ';'.

 .PARAMETER attachmentTypes
  The specified attachment types. Only when the parameter hasAttachments is true the parameter will be applied. Spliy by ';'.

 .PARAMETER searchAllEmails
  Whether search all emails or not. By default is $false.

 .EXAMPLE
   Get messages under login Id 2.
   Get-Messages -loginId 2;
   Get-Messages -id 2;
   Get-Messages 2;

 .EXAMPLE
   Get Messages which subject contains 'ESM' under login Id 2.
   Get-Messages -loginId 2 -subject '*ESM*';
   Get-Messages -id 2 -subject '*ESM*';
   Get-Messages 2 -subject '*ESM*';

 .EXAMPLE
   Get Messages which subject contains 'ESM', has attachment under login Id 2.
   Get-Messages -loginId 2 -subject 'ESM' -hasAttachments $true;
   Get-Messages -id 2 -subject 'ESM' -hasAttachments $true;
   Get-Messages 2 -subject 'ESM' -hasAttachments $true;

 .EXAMPLE
   Get Messages which subject contains 'ESM', has attachment types pdf & txt in Archive Manager database.
   Get-Messages -loginId 2 -subject 'ESM' -hasAttachments $true -attachmentTypes '.pdf;.txt' -searchAllEmails $true;
   Get-Messages -id 2 -subject 'ESM' -hasAttachments $true -attachmentTypes '.pdf;.txt' -searchAllEmails $true;
   Get-Messages 2 -subject 'ESM' -hasAttachments $true -attachmentTypes '.pdf;.txt' -searchAllEmails $true;
#>
Function Get-Messages {
    PARAM(
    [Parameter(Mandatory=$true,
               Position=0)]
    [Alias('id')]
    [int]
    $loginId,
    [string]
    $subject = [string]::Empty,
    [string]
    $to = [string]::Empty,
    [string]
    $from = [string]::Empty,
    [string]
    $toOrFrom = [string]::Empty,
    [nullable[bool]]
    $hasAttachments = $null,
    [string]
    $attachmentNames = [string]::Empty,
    [string]
    $attachmentTypes = [string]::Empty,
    [bool]
    $searchAllEmails = $false
    )

    If (!($loginId -gt 0)) {
        Throw "The login Id $loginId is invalid.`n";
    }

    $criteria = New-Object Quest.AM.BusinessLayer.Queries.MessagesSearchQueryObject;

    If (![string]::IsNullOrEmpty($subject)) {
        $criteria.FullTextSearchCriteria = New-Object Quest.AM.Public.SDK.v1.Models.Queries.FullTextSearchCriteria;
        $criteria.FullTextSearchCriteria.Query = $subject;
        $criteria.FullTextSearchCriteria.Scope = [Quest.AM.Public.SDK.v1.Models.Queries.FullTextSearchScope]::Subject;
    }
    If (![string]::IsNullOrEmpty($to)) {
        $criteria.To = $to;
    }
    If (![string]::IsNullOrEmpty($from)) {
        $criteria.From = $from;
    }
    If (![string]::IsNullOrEmpty($toOrFrom)) {
        $criteria.ToOrFrom = $toOrFrom;
    }
    If ($hasAttachments -ne $null) {
        $criteria.HasAttachment = $hasAttachments;

        If ($hasAttachments) {
            If (![string]::IsNullOrEmpty($attachmentName)) {
                $criteria.AttachmentName = $attachmentName;
            }
            If (![string]::IsNullOrEmpty($attachmentType)) {
                $criteria.AttachmentType = $attachmentType;
            }
        }
    }
    $criteria.SearchAllEmails = $searchAllEmails;

    Write-Host "Get messages by criteria: $criteria`n" -ForegroundColor Green;

    $page = New-Object Quest.AM.Public.SDK.v1.Models.Queries.Page;
    $page.Limit = 0;
    [string]$key = [string]::Empty;
    [int]$total = 0;
    $results = [Quest.AM.BusinessLayer.Messages]::Search($loginId, $criteria, $null, $page, $null, [ref][string]$key, [ref][int]$total, $false);

    Write-Host "$($results.Items.Count) messages satisified the criteria.`n" -ForegroundColor Green;

    Return $results.Items;
}

#endregion

#region Cmdlets for Group, MailServer, LoginDomain, Attachments

<#
 .Synopsis
  Get a list of Group objects from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get a list of Group objects from Archive Manager database. This function supports 3 parameters: name, type and domain Name. 
  It returns an list of Quest.AM.BusinessLayer.Group object. This function will not do the permission check.

 .PARAMETER name
  The group name. Can use whildchar '*'.

 .PARAMETER type
  The group type. Can use whildchar '*'. There are 5 group types in Archive Manager: 
  Active Directory, AfterMail, Azure AD, Exchange 5.5 and GroupWise.

 .PARAMETER domainName
  The group related domain name. Can use whildchar '*'.

 .EXAMPLE
   Get groups which name starts with 'jshi'.
   Get-Groups -name 'jshi';
   Get-Groups 'jshi*';

 .EXAMPLE
   Get groups which name starts with 'jshi' and domain name is 'DEFAULT'.
   Get-Groups -name 'jshi*' -domainName 'DEFAULT';
   Get-Groups 'jshi*' -domainName 'DEFAULT';
#>
Function Get-Groups {
    PARAM(
    [PARAMETER(Position=0)]
    [string]
    $name = [string]::Empty,
    [string]
    $type = [string]::Empty,
    [string]
    $domainName = [string]::Empty
    )

    $criteria = New-Object Quest.AM.BusinessLayer.Queries.GroupsSearchQueryObject;

    If (![string]::IsNullOrEmpty($name)) {
        $criteria.GroupName = $name;
    }
    If (![string]::IsNullOrEmpty($type)) {
        $criteria.GroupType = $type;
    }
    If (![string]::IsNullOrEmpty($domainName)) {
        $criteria.LoginDomainName = $domainName;
    }

    Write-Host "Get groups by criteria: $criteria" -ForegroundColor Green;

    $groups = New-Object Quest.AM.BusinessLayer.Group;
    Return $groups.SelectAll($criteria);
}

<#
 .Synopsis
  Get a list of MailServer objects from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get a list of MailServer objects from Archive Manager database. This function supports 3 parameters: name, type and tenantId. 
  It returns an list of Quest.AM.BusinessLayer.MailServer object. This function will not do the permission check.

 .PARAMETER name
  The mail server name. Can use whildchar '*'.

 .PARAMETER type
  The mail server type. There are 2 mail server types in Archive Manager: 
  Exchange, GroupWise.

 .PARAMETER tenantId.
  The mail server's tenant Id. Only O365 mail servers have tenant Id.

 .EXAMPLE
   Get mail servers which name starts with 'jshi'.
   Get-MailServers -name 'jshi';
   Get-MailServers 'jshi*';

 .EXAMPLE
   Get mail servers which name starts with 'jshi' and type is 'Exchange'.
   Get-Groups -name 'jshi*' -type 'Exchange';
   Get-Groups 'jshi*' -type 'Excahnge';
#>
Function Get-MailServers {
    PARAM(
    [PARAMETER(Position=0)]
    [string]
    $name = [string]::Empty,
    [ValidateSet("Exchange", "GroupWise", "")]
    [string]
    $type = [string]::Empty,
    [int]
    $tenantId = 0
    )

    $criteria = New-Object Quest.AM.BusinessLayer.Queries.MailServersSearchQueryObject;

    If (![string]::IsNullOrEmpty($name)) {
        $criteria.MailServerName = $name;
    }
    If (![string]::IsNullOrEmpty($type)) {
        $criteria.MailServerType = $type;
    }
    If ($tenantId -gt 0) {
        $criteria.TenantId = $tenantId;
    }

    Write-Host "Get Mail Servers by criteria: $criteria`n" -ForegroundColor Green;

    $mailServers = New-Object Quest.AM.BusinessLayer.MailServer;
    Return $mailServers.SelectAll($criteria);
}

<#
 .Synopsis
  Get a list of LoginDomain objects from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get a list of LoginDomain objects from Archive Manager database. This function supports 4 parameters: name, biosName, ldapType and tenantId. 
  It returns an list of Quest.AM.BusinessLayer.LoginDomain object. This function will not do the permission check.

 .PARAMETER name
  The LoginDomain name. Can use whildchar '*'.

 .PARAMETER biosName
  The LoginDomain Net Bios name. Can use whildchar '*'. 

 .PARAMETER ldapType
  The LoginDomain LdapType. There are 5 Ldap types in Archive Manager: 
  Active Directory, Azure AD, Domino, NDS and NT4 Domain.

 .PARAMETER tenantId.
  The LoginDomain's tenant Id. Only O365 LoginDomain have tenant Id.

 .EXAMPLE
   Get LoginDomain which name starts with 'jshi'.
   Get-LoginDomains -name 'jshi';
   Get-LoginDomains 'jshi*';

 .EXAMPLE
   Get LoginDomains which name starts with 'jshi' and ldaptype is 'Azure AD'.
   Get-LoginDomains -name 'jshi*' -ldaptype 'Azure AD';
   Get-LoginDomains 'jshi*' -ldaptype 'LoginDomains';
#>
Function Get-LoginDomains {
    PARAM(
    [PARAMETER(Position=0)]
    [string]
    $name = [string]::Empty,
    [string]
    $biosName = [string]::Empty,
    [ValidateSet("Active Directory", "Azure AD", "Domino", "NDS", "NT4 Domain", "")]
    [string]
    $ldapType = [string]::Empty,
    [int]
    $tenantId = 0
    )

    $criteria = New-Object Quest.AM.BusinessLayer.Queries.LoginDomainsSearchQueryObject;

    If (![string]::IsNullOrEmpty($name)) {
        $criteria.DomainName = $name;
    }
    If (![string]::IsNullOrEmpty($biosName)) {
        $criteria.NetBiosName = $biosName;
    }
    If (![string]::IsNullOrEmpty($ldapType)) {
        $criteria.LdapType = $ldapType;
    }
    If ($tenantId -gt 0) {
        $criteria.TenantId = $tenantId;
    }

    Write-Host "Get Login Domains by criteria: $criteria`n" -ForegroundColor Green;

    $domains = New-Object Quest.AM.BusinessLayer.LoginDomain;
    Return $domains.SelectAll($criteria);
}

<#
 .Synopsis
  Get a list of Attachment objects from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get a list of Attachment objects from Archive Manager database. This function supports 3 parameters: name, type and domain Name. 
  It returns an list of Quest.AM.BusinessLayer.Group object. This function will not do the permission check.

 .PARAMETER loginId
  The login which used to perform the query. Id as the alias.

 .PARAMETER name
  The attachment name. Can be used with whildchar '*'. It will not do the attachment content search but only for the name.

 .PARAMETER type
  The specified attachment types. Spliy by ';'.

 .PARAMETER searchAllEmails
  Whether search all emails or not. By default is $false.

 .EXAMPLE
   Get attachments under login Id 2.
   Get-Attachments -loginId 2;
   Get-Attachments -id 2;
   Get-Attachments 2;

 .EXAMPLE
   Get Attachments which name contains 'ESM' under login Id 2.
   Get-Attachments -loginId 2 -name 'ESM';
   Get-Attachments -id 2 -name 'ESM';
   Get-Attachments 2 -name 'ESM';
#>
Function Get-Attachments {
    PARAM(
    [Parameter(Mandatory=$true,
               Position=0)]
    [Alias("id")]
    [int]
    $loginId = 0,
    [string]
    $name = [string]::Empty,
    [string]
    $type = [string]::Empty,
    [bool]
    $searchAllEmails = $false
    )

    $criteria = New-Object Quest.AM.BusinessLayer.Queries.AttachmentsSearchQueryObject;

    If ($loginId -gt 0) {
        $criteria.LoginId = $loginId;
    }
    If (![string]::IsNullOrEmpty($name)) {
        $criteria.AttachmentName = $name;
    }
    If (![string]::IsNullOrEmpty($type)) {
        $criteria.AttachmentType = $type;
    }
    $criteria.SearchAllEmails = $searchAllEmails;
    
    Write-Host "Get attachments by criteria: $criteria`n" -ForegroundColor Green;

    $page = New-Object Quest.AM.Public.SDK.v1.Models.Queries.Page;
    $page.Limit = 0;
    [string]$key = [string]::Empty;
    [int]$total = 0;
    $attachments = [Quest.AM.BusinessLayer.Attachments]::Search($loginId, $criteria, $null, $page, $null, [ref][string]$key, [ref][int]$total, $false);
    
    Write-Host "$($attachments.Items.Count) attachments satisified the criteria.`n" -ForegroundColor Green;

    Return $attachments.Items;
}
#endregion

#region Cmdlets for Folders

<#
 .Synopsis
  Get a list of Folder objects from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get a list of Folder objects from Archive Manager database. This function supports 2 parameters: folderId, messageId. 
  It returns an list of Quest.AM.BusinessLayer.Folder object. This function will not do the permission check.

 .PARAMETER folderId
  The Folder Id. id as alias.

 .PARAMETER messageId
  The message Id. When this parameter specified,all the folders linked with this message will be returned unless the parameter folderId specified. 

 .EXAMPLE
   Get Folder with folder Id 1.
   Get-Folders -folderId 1;
   Get-Folders -id 1;
   Get-Folders 1

 .EXAMPLE
   Get Folders which related with message 1.
   Get-Folders -messageId 1;
#>
Function Get-Folders {
    PARAM(
    [PARAMETER(Position=0)]
    [Alias("id")]
    [int]
    $folderId = 0,
    [int]
    $messageId = 0)

    $folderIds = New-Object System.Collections.Generic.List[int];
    If ($messageId -gt 0) {
        $ds = [Quest.AM.BusinessLayer.MessageFolder]::SelectByMessageID($messageId);
        If (($ds -ne $null) -and ($ds.Tables[0] -ne $null)) {
            Foreach($dr in $ds.Tables[0].Rows) {
                $folderIds.Add([int]$dr["FolderID"]);
            }
        }
    }

    If ($folderId -gt 0) {
        If (($folderIds.Count -gt 0) -And (!$folderIds.Contains($folderId))) {
            Throw "Folder Id $folderId is not related with message: $messageId";
        }

        Write-Host "Get folder by folder Id: $folderId`n" -ForegroundColor Green;

        $folder = [Quest.AM.BusinessLayer.Folder]::Select($folderId);
        Return @($folder);
    }

    If ($folderIds.Count -gt 0) {
        Write-Host "Get folders by message Id: $messageId`n" -ForegroundColor Green;
        
        $folders = New-Object System.Collections.Generic.List[Quest.AM.BusinessLayer.Folder];
        Foreach($id in $folderIds) {
            $f = [Quest.AM.BusinessLayer.Folder]::Select($id);
            $folders.Add($f);
        }

        Return $folders;
    }

    Write-Host "No valid parameters specified. Return nothing.`n" -ForegroundColor Red;

    Return $null;
} 
#endregion

#region Cmdlets for configs

<#
 .Synopsis
  Get the value of a config from Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get the value of a specified config from Archive Manager database. This function supports 1 parameter: config. 
  It returns the value of the config.

 .PARAMETER config
  The config name. Like: Directory Connector Import Users Without MailBoxes, Directory Connector Add Email Address To MailBox

 .EXAMPLE
   Get value of config "Directory Connector Add Email Address To MailBox".
   Get-Config -config "Directory Connector Add Email Address To MailBox";
   Get-Config "Directory Connector Add Email Address To MailBox";
#>
Function Get-Config {
    PARAM(
    [PARAMETER(Mandatory=$true, 
    Position=0)]
    [string]
    $config
    )

    Write-Host "Get value of Config: $config" -ForegroundColor Green;    
    Return [Quest.AM.BusinessLayer.Config]::GetValue($config);
}

<#
 .Synopsis
  Set the value of a specified config in Archive Manager database. It needs Archive Manager be installed in this server.

 .DESCRIPTION
  Get the value of a specified config in from Archive Manager database. This function supports 2 parameters: config, value. 

 .PARAMETER config
  The config name. Like: Directory Connector Import Users Without MailBoxes, Directory Connector Add Email Address To MailBox.

 .PARAMETER value
  The value of specified config. 

 .EXAMPLE
   Set config "Directory Connector Add Email Address To MailBox" to true.
   Set-Config -config "Directory Connector Add Email Address To MailBox" -value $true;
   Set-Config "Directory Connector Add Email Address To MailBox" -value $true;
#>
Function Set-Config {
    PARAM(
    [PARAMETER(Mandatory=$true, 
    Position=0)]
    [string]
    $config,
    [PARAMETER(Mandatory=$true)]
    [object]
    $value
    )

    Write-Host "Set value '$value' to config: $config" -ForegroundColor Green;
    [Quest.AM.BusinessLayer.Config]::SetValue($config, $value);
}

#endregion

#region Send message

<#
.Synopsis
   Send message through O365, On-Premise Exchange, GroupWise.
.DESCRIPTION
   Send message through O365, On-Premise Exchange, GroupWise. You have to specify From, To, Subject, etc.
.PARAMETER environment
   The environment. 3 values are available: O365, On-Premise, GroupWise. By default it is On-Premise.
.PARAMETER smtpServer
   The SmtpServer used to send message. For O365, you don't need to specify it. For On-Premise/GroupWise in AT, you don't need to specify it either.
.PARAMETER port
   The port for sending message. Default is 25. For O365, the default value is 587.
.PARAMETER useSsl
   When sending message, whether use Ssl or not. For O365, the default value is true. For others, the default value is false.
.PARAMETER from
   Specifies the address from which the mail is sent. Enter a name (optional) and email address, such as Name <someone@example.com>. The parameter is required.
.PARAMETER to
   Specifies the addresses to which the mail is sent. Enter names (optional) and the email address, such as Name <someone@example.com>. The parameter is required.
.PARAMETER cc
   Specifies the email addresses to which a carbon copy (CC) of the email message is sent. Enter names (optional) and the email address, such as Name <someone@example.com>.
.PARAMETER bcc
   Specifies the email addresses that receive a copy of the mail but are not listed as recipients of the message. Enter names (optional) and the email address, such as Name <someone@example.com>.
.PARAMETER subject
   Specifies the subject of the email message. This parameter is required.
.PARAMETER body
   Specifies the body of the email message.
.PARAMETER attachment
   Specifies the path and file name of file to be attached to the email message.
.PARAMETER priority
   Specifies the priority of the email message. The acceptable values for this parameter are: Normal, High, Low. Default is Normal.
.PARAMETER cred
   Specifies a user account that has permission to perform this action. The default is the current user.
.EXAMPLE
   Send a message from one user to another in O365
   Send-Message -environment O365 -from "qamat@archivemgr.onmicrosoft.com" -to "Admin@archivemgr.onmicrosoft.com" -subject "Aha, test";
.EXAMPLE
   Send a message from one user to another in O365 with attachment
   Send-Message -environment O365 -from "qamat@archivemgr.onmicrosoft.com" -to "Admin@archivemgr.onmicrosoft.com" -subject "Aha, test" -attachments "C:\Attachments\ESM.wlog";
#>
Function Send-Message {
    PARAM(
    [PARAMETER(Position=0)]
    [ValidateSet("O365", "On-Premise", "GroupWise")]
    [string]
    $environment = "On-Premise",
    [string]
    $smtpServer = [string]::Empty,
    [int]
    $port = 25,
    [bool]
    $useSsl = $false,
    [Parameter(Mandatory=$true)]
    [string]
    $from = [string]::Empty,
    [Parameter(Mandatory=$true)]
    [string]
    $to = [string]::Empty,
    [string]
    $cc = [string]::Empty,
    [string]
    $bcc = [string]::Empty,
    [Parameter(Mandatory=$true)]
    [string]
    $subject,
    [string]
    $body = [string]::Empty,
    [string]
    $attachment = [string]::Empty,
    [System.Net.Mail.MailPriority]
    $priority = [System.Net.Mail.MailPriority]::Normal,
    [System.Management.Automation.PSCredential]
    $cred = $null
    )

    Switch($environment.ToUpper()) {
        "O365" {
            $port = 587;
            $useSsl = $true;

            If ($cred -eq $null) {
                $pwd = ConvertTo-SecureString "Pa`$`$word" -AsPlainText -Force;
                $cred = New-Object System.Management.Automation.PSCredential("qamat@archivemgr.onmicrosoft.com", $pwd);
            }
            If ([string]::IsNullOrEmpty($smtpServer)) {
                $smtpServer = "smtp.office365.com";
            }
            If ([string]::IsNullOrEmpty($from)) {
                $from = "qamat@archivemgr.onmicrosoft.com";
            }
            break;
        }

        "ON-PREMISE" {
            If ([string]::IsNullOrEmpty($smtpServer)) {
                $smtpServer = "DC.am.com";
            }
            break;
        }

        "GROUPWISE" {
            If ([string]::IsNullOrEmpty($smtpServer)) {
                Throw "Smtp server must be specified.";
            }
            break;
        }
    }

    $cmd = "Send-MailMessage -SmtpServer '$smtpServer' -Port $port -Subject '$subject' -From '$from' -To '$to' -Priority $priority";
    If ($useSsl) {
        $cmd += " -UseSsl";
    }
    If (![string]::IsNullOrEmpty($cc)) {
        $cmd += " -CC '$cc'";
    }
    If (![string]::IsNullOrEmpty($bcc)) {
        $cmd += " -Bcc '$bcc'";
    }
    If (![string]::IsNullOrEmpty($attachments)) {
        $cmd += " -Attachments '$attachments'";
    }
    If (![string]::IsNullOrEmpty($body)) {
        $cmd += " -Body '$body'";
    }
    If ($cred -ne $null) {
        $cmd += " -Credential `$cred";
    }

    Write-Host "Sending message by cmdlet: $cmd `n" -ForegroundColor Green;
    Invoke-Expression -Command $cmd;
}

#endregion