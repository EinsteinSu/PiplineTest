    Given "Archive Manager installed" {
        # Do nothing
    }
    And "AD Connector configured" {
        # Do nothing
    }

    AfterEachScenario {
        Write-Host "Stop Archive Manager service." -ForegroundColor Gray;
        Stop-Service -Name "Archive Active Directory Connector Service";
    }
    
    #Scenario: ADC service sync AD objects (User | MailBox | Group) into Archive Manager database
    Given "ADC service has not been executed before" {        
        $adc = Get-Service -Name "Archive Active Directory Connector Service";
        $adc.Status | Should -Be "Stopped";
    }
    When "ADC service was executed" {
        Start-Service -Name "Archive Active Directory Connector Service";
        Sleep 100; # wait the service to finish
    }
    Then "Number of Logins should be greater than 1" {
        $lgnCount = Get-LoginsCount;
        $lgnCount | Should -BeGreaterThan 1;
    }
    And "Number of MailBoxes should be equal or greater than 1" {
        $mbxCount = Get-MailBoxesCount;
        $mbxCount | Should -BeGreaterOrEqual 1;
    }
    And "Number of Groups should be equal or greater than 1" {
        $groups = Get-Groups;
        $groups.Count | Should -BeGreaterOrEqual 1;
    }
    And "User (?<user>.+) should be synced" {
        PARAM($user)
        $login = Get-Logins -emailAddress $user;
        $login | Should -Not -BeNullOrEmpty;
        $login.PrimaryMailBoxID | Should -BeGreaterThan 0;
    }

    #Scenario: ADC service sync Logins without mailboxes into Archive Manager database after config "Directory Connector Import Users Without MailBoxes"
    Given "ADC service is stopped" {
        $adc = Get-Service -Name "Archive Active Directory Connector Service";
        $adc.Status | Should -Be "Stopped";
    }
    And "value of config `"Directory Connector Import Users Without MailBoxes`" is `"false`"" {
        $v = Get-Config "Directory Connector Import Users Without MailBoxes";
        If ($v -ne $null) {
            $v | Should -Be $false;
        }
    }
    When "Set the value of confi `"Directory Connector Import Users Without MailBoxes`" to `"true`"" {
        Set-Config "Directory Connector Import Users Without MailBoxes" -value $true;
    }
    And "ADC service was executed" {
        Start-Service -Name "Archive Active Directory Connector Service";
        Sleep 100; # wait the service to finish
    }
    Then "Number of Logins without mailboxes should be greater than 1" {
        $logins = Get-Logins;
        $loginsWithoutMbx = $logins | %{ If($_.PrimaryMailBoxID -eq -1) { Return $_; }}
        
        $loginsWithoutMbx.Count | Should -BeGreaterThan 1;
    }