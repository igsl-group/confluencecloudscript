<#
	.SYNOPSIS 
	Add page templates in Confluence Cloud. 
	
	.PARAMETER Csv
		CSV input file. Use this SQL to generate it: 
		SELECT 
			TEMPLATENAME, CONTENT 
		INTO OUTFILE 
			'[Path of CSV file]'
		FIELDS 
			TERMINATED BY ','
			ENCLOSED BY '"'
		LINES 
			TERMINATED BY '\n'
		FROM 
			PAGETEMPLATES;

	.PARAMETER CsvHasHeader
		Specify this switch if the CSV contains header row.
		Make sure header row contains these columns: TEMPLATENAME, CONTENT

	.PARAMETER Domain
		Confluence cloud domain, e.g. kcwong.atlassian.net
		
	.PARAMETER Email
		Email of Confluence cloud user with write access to all spaces, e.g. kc.wong@igsl-group.com
		
	.PARAMETER ApiToken
		Confluence cloud API token.
		See: https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/
#>
Param (
	[Parameter(Mandatory)][string] $Csv,
	[switch] $CsvHasHeader,
	[Parameter(Mandatory)][string] $Domain,
	[Parameter(Mandatory)][string] $Email,
	[Parameter(Mandatory)][string] $ApiToken
)

function GetAuthHeader {
	Param (
		[string] $Email,
		[string] $ApiToken
	)
	[hashtable] $Headers = @{
		"Content-Type" = "application/json"
	}
	$Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($Email + ":" + $ApiToken))
	$Headers.Authorization = "Basic " + $Auth
	$Headers
}

# Call Invoke-WebRequest without throwing exception on 4xx/5xx 
function WebRequest {
	Param (
		[string] $Uri,
		[string] $Method,
		[hashtable] $Headers,
		[object] $Body
	)
	$Response = $null
	try {
		$script:ProgressPreference = 'SilentlyContinue'    # Subsequent calls do not display UI.
		$Response = Invoke-WebRequest -Method $Method -Header $Headers -Uri $Uri -Body $Body
	} catch {
		$Response = @{}
		$Response.StatusCode = $_.Exception.Response.StatusCode.value__
		$Response.content = $_.Exception.Message
	} finally {
		$script:ProgressPreference = 'Continue'            # Subsequent calls do display UI.
	}
	$Response
}

function ProcessNull {
	Param (
		[string] $Data
	)
	$Result = $Null
	if ($Data) {
		# Null gets turned into different things depending on how the CSV is generated
		if ($Data -eq '\N' -or $Data -eq 'NULL' -or $Data -eq '') {
			$Result = $Null
		} else {
			$Result = $Data
		}
	}
	$Result
}

function CreatePageTemplate {
	Param (
		[hashtable] $Headers,
		[string] $Name,
		[string] $Content
	)
	$Uri = 'https://' + $Domain + '/wiki/rest/api/template'
	$Body = @{
		'name' = $Name;
		'templateType' = 'page';
		'body' = @{
			'storage' = @{
				'value' = $Content;
				'representation' = 'view'
			}
		};
	}
	$Response = WebRequest $Uri 'POST' $Headers ($Body | ConvertTo-Json)
	switch ($Response.StatusCode) {
		200 {
			Write-Host ('Page template ' + $Name + ' added')
			break
		}
		default {
			throw 'Unable to add page template: ' + $Response.StatusCode + ': ' + $Response.Content
		}
	}
}

# Auth Header
$Headers = GetAuthHeader $Email $ApiToken

# Read space shortcut data from CSV
if ($CsvHasHeader) {
	$Data = Import-Csv -Path $Csv
} else {
	$Data = Import-Csv -Path $Csv -Header @('TEMPLATENAME', 'CONTENT')
}

foreach ($Line in $Data) { 
	$Name = ProcessNull $Line.TEMPLATENAME
	$Content = ProcessNull $Line.CONTENT
	try {
		CreatePageTemplate $Headers $Name $Content
	} catch {
		Write-Host $PSItem
	}
}