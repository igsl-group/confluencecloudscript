<#
	.SYNOPSIS 
	Set space shortcuts in Confluence Cloud. 
	NOTE: Existing space shortcuts will be overwritten.

	.PARAMETER Csv
		CSV input file. Use this SQL to generate it: 
		SELECT 
			link.SPACE_KEY, link.CUSTOM_TITLE, s.SPACEKEY, c.TITLE, link.HARDCODED_URL, link.POSITION 
		INTO OUTFILE 
			'[Path of CSV file]'
		FIELDS 
			TERMINATED BY ','
			ENCLOSED BY '"'
		LINES 
			TERMINATED BY '\n'
		FROM 
			AO_187CCC_SIDEBAR_LINK link
			LEFT JOIN CONTENT c 
				ON c.CONTENTID = link.DEST_PAGE_ID
			LEFT JOIN SPACES s 
				ON s.SPACEID = c.SPACEID 
		WHERE 
			link.CATEGORY = 'QUICK' 
		ORDER BY 
			SPACE_KEY, POSITION DESC;

	.PARAMETER CsvHasHeader
		Specify this switch if the CSV contains header row.
		Make sure header row contains these columns: SPACE_KEY, CUSTOM_TITLE, SPACEKEY, TITLE, HARDCODED_URL, POSITION

	.PARAMETER Domain
		Confluence cloud domain, e.g. kcwong.atlassian.net
		
	.PARAMETER Email
		Email of Confluence cloud user with write access to all spaces, e.g. kc.wong@igsl-group.com
		
	.PARAMETER ApiToken
		Confluence cloud API token.
		See: https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/
#>
Param(
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

function GetSpaceId {
	Param(
		[hashtable] $Headers,
		[string] $SpaceKey
	)
	$Uri = 'https://' + $Domain + '/wiki/api/v2/spaces?keys=' + [uri]::EscapeDataString(${SpaceKey})
	$Response = WebRequest $Uri 'GET' $Headers
	switch ($Response.StatusCode) {
		200 {
			$Json = ConvertFrom-Json $Response.Content
			if ($Json.results.Count -eq 1) {
				$Json.results[0].id
			} else {
				throw 'Unable to locate space id for ' + ${SpaceKey}
			}
			break
		}
		default {
			throw 'Unable to locate space id for ' + ${SpaceKey} + ': ' + $Response.StatusCode + ': ' + $Response.Content
		}
	}
}

function GetPageURL {
	Param (
		[hashtable] $Headers,
		[string] $SpaceId,
		[string] $PageTitle
	)
	$Uri = 'https://' + $Domain + '/wiki/api/v2/pages?space-id=' + [uri]::EscapeDataString(${SpaceId}) + '&title=' + [uri]::EscapeDataString(${PageTitle})
	$Response = WebRequest $Uri 'GET' $Headers
	switch ($Response.StatusCode) {
		200 {
			$Json = ConvertFrom-Json $Response.Content
			if ($Json.results.Count -eq 1) {
				'https://' + $Domain + '/wiki' + $Json.results[0]._links.webui
			} else {
				throw 'Unable to locate page URL for ' + ${SpaceId} + ', ' + ${PageTitle}
			}
			break
		}
		default {
			throw 'Unable to locate page URL for ' + ${SpaceId} + ', ' + ${PageTitle} + ': ' + $Response.StatusCode + ': ' + $Response.Content
		}
	}
}

function SetSpaceShortcuts {
	Param (
		[hashtable] $Headers,
		[string] $Key,
		[System.Collections.ArrayList] $Items
	)
	try {
		# Test target space key
		$Check = GetSpaceId $Headers $Key
		# Set shortcuts
		$Uri = 'https://' + $Domain + '/wiki/rest/ia/1.0/link/batch'
		$Body = @{
			'spaceKey' = ${Key};
			'quickLinks' = [System.Collections.ArrayList]::new();
		}
		foreach ($Item in $Items) {
			[void] $Body.quickLinks.Add($Item)
		}
		$Response = WebRequest $Uri 'POST' $Headers ($Body | ConvertTo-Json)
		switch ($Response.StatusCode) {
			200 {
				Write-Host ('Space ' + $Key + ': ' + $Items.Count + ' shortcut(s) applied')
				break
			}
			default {
				throw 'Unable to add space shortcuts: ' + $Response.StatusCode + ': ' + $Response.Content
			}
		}
	} catch {
		throw ('Unable to resolve container space key: ' + $Key)
	}
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

# Auth Header
$Headers = GetAuthHeader $Email $ApiToken

# Read space shortcut data from CSV
if ($CsvHasHeader) {
	$Data = Import-Csv -Path $Csv
} else {
	$Data = Import-Csv -Path $Csv -Header @('SPACE_KEY','CUSTOM_TITLE','SPACEKEY','TITLE','HARDCODED_URL','POSITION')
}

$Items = [System.Collections.ArrayList]::new()
$PrevKey = $Null
foreach ($Line in $Data) { 
    # Space containing shortcut
	$Key = ProcessNull $Line.SPACE_KEY
	# Web link
	$CustomTitle = ProcessNull $Line.CUSTOM_TITLE
	$Url = ProcessNull $Line.HARDCODED_URL
	# Page link
	$SpaceKey = ProcessNull $Line.SPACEKEY
	$Title = ProcessNull $Line.TITLE
	if ($PrevKey -ne $Null -and $PrevKey -ne $Key) {
		# Create space shortcuts in cloud
		try {
			SetSpaceShortcuts $Headers $PrevKey $Items
		} catch {
			Write-Host $PSItem
		}
		$Items.Clear()
	}
	$NewItem = $Null
	$PrevKey = $Key
	if ($Url -eq $Null) {
		# Page, find new page IDs on Cloud
		try {
			$SpaceId = GetSpaceId $Headers $SpaceKey
			$Url = GetPageURL $Headers $SpaceId $Title
			$NewItem = @{
				'title' = $CustomTitle; 
				'url' = $Url;
				'id' = $Null;
			}
		} catch {
			Write-Host ('Unable to resolve target page space key: ' + $SpaceKey)
		}
	} else {
		# Custom URL
		$NewItem = @{
			'title' = $Null; 
			'url' = $Url;
			'id' = $Null;
		}
	}
	if ($NewItem) {
		[void] $Items.Add($NewItem)
	}
}
try {
	SetSpaceShortcuts $Headers $PrevKey $Items
} catch {
	Write-Host $PSItem
}