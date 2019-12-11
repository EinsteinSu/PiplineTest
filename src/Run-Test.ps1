PARAM(
    [ValidateSet("ADC", "ESM", "MAPIDL", "FTI", "FTS", "GDC", "GSM", "Retention", "Alert", "")]
    [string]
    $component = [string]::Empty,
    [string]
    $tags = [string]::Empty)

$basePath = "C:\AT";
$global:testPath = [string]::Empty;
$global:testResult = [string]::Empty;

# Pre-req: Installing the required modules.
Function Import-Modules {
    Write-Host "Preparing to run automation test, installing modules first...`n" -ForegroundColor Green;

    Write-Host "Importing Pester module!" -ForegroundColor Green;
    Import-module -Name Pester -Force;

    Write-Host "Importing AM module!" -ForegroundColor Green;
    $amModule = [System.IO.Path]::Combine($basePath, "Modules\AM\AM.PSD1");
    Import-Module $amModule -Force;
}

# Pre-req: Preparing test path, test result path.
Function Initial-TestPathes {
    $global:testPath = [System.IO.Path]::Combine($basePath, "Tests");
    If (![string]::IsNullOrEmpty($component)) {
        $global:testPath = [System.IO.Path]::Combine($global:testPath, $component);
    }

    Write-Host "Running tests under $global:testPath." -ForegroundColor Green;

    $global:testResult = [System.IO.Path]::Combine($global:testPath, "Results");
    If (!(Test-Path $global:testResult)) {
        Write-Host "Creating Results folder..." -ForegroundColor Green;
        mkdir $global:testResult;
    }

    $global:testResult = [System.IO.Path]::Combine($global:testResult, [string]::Format("{0}{1}", $component, "Result.xml"));
}

Import-Modules;
Initial-TestPathes;

# Execute test cases
If ([string]::IsNullOrEmpty($tags)) {
    Write-Host "Running tests under $global:testPath... Test result will be located: $global:testResult" -ForegroundColor Green;
    Invoke-Gherkin -Path $global:testPath -OutputFile $global:testResult -OutputFormat NUnitXml;
}
Else {
    Write-Host "Running tests under $global:testPath witg tags: $tags... Test result will be located: $global:testResult" -ForegroundColor Green;
    Invoke-Gherkin -Path $global:testPath -Tag $tags -OutputFile $global:testResult -OutputFormat NUnitXml;
}

Return $global:testResult;