<#
.SYNOPSIS
Gets the user data by first 3 digit of telephone number.

.DESCRIPTION
Uses Active Directory to get members data (name, email telephone). Use parameters for
decrease the search area, and get exact list in CSV file format.

.PARAMETER Number
The telephone number first 3 digit. e.g.: [Number]-0000-000 
from 123-0000-000 where number is 123

.PARAMETER OutDir
The directory where the CSV should be saved. This parameter
is required*

.EXAMPLE
.\Get-MembersByPhone.ps1 -Number 644 -OutFile out.csv

Downloads the members data in the current directory into out.csv.
(filtered by phone number which is 644)

.EXAMPLE
.\Get-MembersByPhone.ps1 -OutFile out.csv -Number 918 -Unit Academic

Downloads the "Academic" unit members data in the current directory into out.csv.
(filtered by phone number which is 918)


.NOTES
- Author: Peter Bartha, 2015.03.28.

#>

# Parameter validation
# There are two required parameter: Number and OutFile and one optional (Unit) with default value
param (
	[Parameter(Mandatory=$true)][string] $Number,
	[Parameter(Mandatory=$true)][string] $OutFile,
	[string] $Unit = ""
)

# Checking number format: 3 characters, each of these chars in [0-9] interval
if (!($Number -match "^[0-9]{3}$")) {
	throw "Telephone number format is not correct! It must be 3 characters, including only numbers in 0-9 interval.";
}

# Loading AD Provider if not exist in actual session
if (!(Test-Path 'AD:\')) {
	try {
		Import-Module ActiveDirectory
	} catch [Exception] {
		Write-Error("Active Directory provider cannot be loaded.");
		return;
	}
}

# Get unit "DistinguishedName" by Unit parameter
# Use Get-ADOrganizationalUnit with some filtering option to optimize request
if ($Unit.length -gt 0) {
	$unitObject = Get-ADOrganizationalUnit -Filter { Name -eq $Unit } -Properties Name, DistinguishedName | Select name, distinguishedName | Sort-Object name -Unique; 

	if ($unitObject -eq $null) {
		throw "'"+ $Unit +"' organization unit not exists.";
	} else {
		$unitBase = $unitObject.distinguishedName;
	}
}

# User list filtering by parameters
# Use Get-ADUser with some filtering option to optimize request
# Returns with important properties only
$numberPattern = "$Number*";
if ($unitBase -eq $null) {
	$selectedEmployees = @(Get-ADUser -Filter {(telephoneNumber -like $numberPattern)} -Properties DisplayName, EmailAddress, TelephoneNumber -SearchScope Subtree);
} else {
	$selectedEmployees = @(Get-ADUser -Filter {(telephoneNumber -like $numberPattern)} -Properties DisplayName, EmailAddress, TelephoneNumber -SearchBase $unitBase -SearchScope Subtree);
}

# Create a list with useful user data (for writing into CSV)
$filteredData = @($selectedEmployees | select -Property `
	@{n = "Name";  e = {$_.DisplayName};},
	@{n = "Email"; e = {$_.EmailAddress};},
	@{n = "Phone"; e = {$_.TelephoneNumber};} | Sort-Object Name);

# Write previous list to CSV with specified "Name;Email;Phone" header
try {
	@($filteredData | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | % {$_ -replace '"', ''}) | Out-File -FilePath $OutFile;
} catch [Exception] {
	Write-Error("Cannot write to " + $OutFile);
	return;
}

# Clean up
Clear-Variable unitObject -ErrorAction SilentlyContinue
Clear-Variable selectedEmployees -ErrorAction SilentlyContinue
Clear-Variable filteredData -ErrorAction SilentlyContinue
