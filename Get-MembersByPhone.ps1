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


# PARAMÉTEREK MEGADÁSA, EGYSZERŰ VALIDÁLÁSA
# 	paraméterek ellenőrzése, Number és OutFile kötelezően megadandó paraméterek (lista elejére valóak)
# 	amennyiben a kötelező paraméterek nincsenek kitöltve a [Parameter(Mandatory=$true)] miatt a program bekéri azt
#	a Unit nem kötelező, ezt üres string-el inicializálom, ha nincs megadva (később keresési változó lesz, kell hogy létezzen)
param (
	[Parameter(Mandatory=$true)][string] $Number,
	[Parameter(Mandatory=$true)][string] $OutFile,
	[string] $Unit = ""
)


# 	Number paraméter formátumának ellenőrzése, 3 karakter lehet, mind a 3 szám [0-9] intervallumban
# 	legegyszerűbben regex-el ellenőrizhető a forma, "^" jel a sor elejére illeszt, a "$" pedig a végére,
# 	ezzel megadva a pontos egyezést, belül: [0-9] intervallumból várunk számokat, mégpedig 3-at (jele: {3})

if (!($Number -match "^[0-9]{3}$")) {
	throw "Telephone number format is not correct! It must be 3 characters, including only numbers in 0-9 interval.";
}


# AD PROVIDER BETÖLTÉSE
# 	amennyiben még nem létezik az AD provider, akkor valószínűleg nem volt még betöltve, ezt ellenőrizem
# 	ha nem volt, akkor megpróbálom betölteni, ha viszont nem sikerül akkor hibával tér vissza a script,
#	a betöltés az Import-Module ActiveDirectory végzi
#
#	Ha hiba van akkor kiírom az üzenetet, illetve új sorban pontosan a hiba okát $_.Exception.ToString()-el

if (!(Test-Path 'AD:\')) {
	try {
		Import-Module ActiveDirectory
	} catch [Exception] {
		Write-Error("Active Directory provider cannot be loaded.");
		return;
	}
}


# SZERVEZETI EGYSÉG VALIDÁLÁSA
# 	levizsgálom, hogy ha ki van töltve a Unit paraméter, ha igen akkor leellenőrzöm,
# 	hogy az AD-ban szerepel-e, ha nem akkor hibával térek vissza (ekkor a unitObject értéke null)
#	a vizsgálat helyének keresésére a Get-ADOrganizationalUnit függvény használom fel
#	ha a unitObject null, akkor nincs találat, azaz hibával térek vissza, minden más esetben viszont,
#	kiválasztom a (létező és megtalált) unitObject distinguishedName-ét, amit a Get-ADUser SearchBase-nél használok fel,
#	mivel egy if-en belül van az egész, ezért ha nincs találat, akkor a unitBase értéke is null lesz (nem definiált)
#	ezt fogom kihasználni az ellenőrzéskor

# 	Unit értéke üres string ha nem adjuk meg, ezért a hosszát vizsgálom
if ($Unit.length -gt 0) {
	$unitObject = Get-ADOrganizationalUnit -Filter { Name -eq $Unit } -Properties Name, DistinguishedName | Select name, distinguishedName | Sort-Object name -Unique; 

	if ($unitObject -eq $null) {
		throw "'"+ $Unit +"' organization unit not exists.";
	} else {
		$unitBase = $unitObject.distinguishedName;
	}
}

# FELHASZNÁKÓK SZŰRÉSE A PARAMÉTEREK ALAPJÁN
#	a beadott paraméterek valadálása után, segítségükkel kiszűröm mely felhasználó(k) szerepelnek a végső listában,
#	ehhez a Get-ADUser függvényt használom fel, ezen végze, a szűrést
#	ha a visszatérési érték nem tömb, akkor a @() tömbképző elemmel azzá konvertálom, egyébként marad egyszerű tömb,
#	igyekeztem a szerveren végezni a nagyobb lekérdezést, a properties-el teljesítmény miatt tovább szűrtem,
#	a telefonszámra a numberPattern váltózóval szűrök rá a Filter-ben, hiszen itt már biztosan 3 számjegy van benne,
#	az elágazás megnézi, hogy a unitBase definiált-e, ha nem akkor a -SearchBase kapcsoló nélkül kérdezem le az
# 	adatokat, egyébként pedig a unitBase tartalmazza a megadott Unit OU-ját, így abból fogok kiindulni a keresésékor,
#	azért volt szükség külön lekérdezésre, mert ha a SearchBase utáni változó üres string, akkor a függvény hibával tért vissza,
#	a -SearchScope Subtree kapcsoló az adott Base-ben és a gyerekeiben (ill. annak a gyerekeiben...) keres,
#	a Properties kapcsoló segítségével kiveszem azon mezőket amik a végeredmény szempontjából érdekesek

$numberPattern = "$Number*";
if ($unitBase -eq $null) {
	$selectedEmployees = @(Get-ADUser -Filter {(telephoneNumber -like $numberPattern)} -Properties DisplayName, EmailAddress, TelephoneNumber -SearchScope Subtree);
} else {
	$selectedEmployees = @(Get-ADUser -Filter {(telephoneNumber -like $numberPattern)} -Properties DisplayName, EmailAddress, TelephoneNumber -SearchBase $unitBase -SearchScope Subtree);
}


# 	végül kiszűröm a tényleges eredményt az OutFile változóba megadott fájlhoz,
#	itt már a feladatban megadott oszlopneveket használom az eredmény előállításakor,
#	a Property csatolóval elérem, hogy a $_.(mezőneve) változók ki legyenek töltve,
#	ennek az értékét már csak át kell adnom a hozzátartozó mezőnek egy expression (e)-vel,
#	a végén rendezem az objektumot Name, azaz név szerint ahogy a feladat kéri és így
#	tárolom le a filteredData változóban ami már csak a hasznos és formázott adatot tárolja

$filteredData = @($selectedEmployees | select -Property `
	@{n = "Name";  e = {$_.DisplayName};},
	@{n = "Email"; e = {$_.EmailAddress};},
	@{n = "Phone"; e = {$_.TelephoneNumber};} | Sort-Object Name);


# CSV FÁJLBA ÍRÁS
# 	a szűrést követően az eredményt a feladatban elvárt módon CSV fájlban rögzítem,
#	sorrendben és formailag helyesen a mezők: Name;Email;Phone
#	a filteredData egy PSCustomObject, a ConvertTo-Csv csatoló átkonvertálom CSV formátumúra,
#	hogy ne írja tele a fájlt a típus kommentekkel, ezért használtam a NoTypeInformation csatolót,
#	az elválasztó karakter pedig a feladatban megadott ; karakter,
#	végül kiszedem a már formázott adatból a " karaktereket, hogy minél inkább hasonlítson a megadott kimeneti formátumra,
#	az Out-File pipe, FilePath csatolójával megadom, hogy az OutFile változóban specifikált helyre/névre mentse el a CSV-t,
#
#	ha bármiféle hiba történt, pl. nem lehet írni a fájlt mert használatban van vagy nincs jogosultságunk írni,
#	akkor hibával térek vissza

try {
	@($filteredData | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | % {$_ -replace '"', ''}) | Out-File -FilePath $OutFile;
} catch [Exception] {
	Write-Error("Cannot write to " + $OutFile);
	return;
}


# TISZTÍTÁS
#	törlöm a nagyobb tartalmú változók értékeit, itt már úgysincs rájuk szükség
Clear-Variable unitObject -ErrorAction SilentlyContinue
Clear-Variable selectedEmployees -ErrorAction SilentlyContinue
Clear-Variable filteredData -ErrorAction SilentlyContinue