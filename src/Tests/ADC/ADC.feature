Feature: ADC service can sync users, mailboxes, groups into Archive Manager

Background: Archive Manager installed and configured.
    Given Archive Manager installed
    And AD Connector configured

@ADC @On-Premise
Scenario: ADC service sync AD objects (User | MailBox | Group) into Archive Manager database
    Given ADC service has not been executed before
    When ADC service was executed
    Then Number of Logins should be greater than 1
    And Number of MailBoxes should be equal or greater than 1
    And Number of Groups should be equal or greater than 1
    And User pattif@m365b021248.onmicrosoft.com should be synced

@ADC @On-Premise @Config
Scenario: ADC service sync Logins without mailboxes into Archive Manager database after config "Directory Connector Import Users Without MailBoxes"
    Given ADC service is stopped
    And value of config "Directory Connector Import Users Without MailBoxes" is "false"
    When Set the value of confi "Directory Connector Import Users Without MailBoxes" to "true"
    And ADC service was executed
    Then Number of Logins without mailboxes should be greater than 1