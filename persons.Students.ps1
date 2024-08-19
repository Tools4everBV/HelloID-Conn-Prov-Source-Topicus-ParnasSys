$config = $configuration | ConvertFrom-Json

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Set debug logging
switch ($($c.isDebug)) {
    $true { $VerbosePreference = 'Continue' }
    $false { $VerbosePreference = 'SilentlyContinue' }
}

#region functions
function Get-ParnasSysLeerlingen {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]
        $Brin,

        [Parameter(Mandatory)]
        [string]
        $WebServiceUri,

        [Parameter(Mandatory)]
        [string]
        $SupplierName,

        [Parameter(Mandatory)]
        [string]
        $SupplierKey,

        [Parameter(Mandatory)]
        [string]
        $SchoolYear,

        [Parameter()]
        [string]
        $proxy
    )
    try {
        $headers = @{
            'Content-Type' = 'text/xml; charset=utf-8'
            SOAPaction     = "`"getLeerlingen`""
        }
        # Fix the '&' char in the supplierName
        $supplierNameEncoded = [System.Net.WebUtility]::HtmlEncode($SupplierName)

        $xml = [xml]('<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema">
                <soap:Body>
                    <getLeerlingen xmlns="http://www.topicus.nl/parnassys">
                        <leveranciernaam xmlns="">{0}</leveranciernaam>
                        <leveranciersleutel xmlns="">{1}</leveranciersleutel>
                        <brinnummer xmlns="">{2}</brinnummer>
                        <schooljaar xmlns="">{3}</schooljaar>
                    </getLeerlingen>
                </soap:Body>
            </soap:Envelope>
        ' -f $supplierNameEncoded, $SupplierKey, $Brin, $SchoolYear)

        $splatWebRequestParameters = @{
            Uri             = $webServiceUri
            Method          = 'Post'
            Headers         = $headers
            UseBasicParsing = $true
            ContentType     = 'text/xml'
            Body            = $xml.InnerXml
        }
        if (-not  [string]::IsNullOrEmpty($Proxy)) {
            $splatWebRequestParameters['Proxy'] = $Proxy
        }

        $result = Invoke-WebRequest @splatWebRequestParameters

        [xml] $parnasSysDataxml = $result.Content
        $leerLingenResponseNode = $parnasSysDataxml.FirstChild.FirstChild.FirstChild
        $returnNode = $leerLingenResponseNode.item('return')
        Write-Output $returnNode
    }
    catch {
        Write-Verbose "Could not get Students for Brin: [$brin] [$SchoolYear]" -Verbose
        Write-Verbose "Error Details: [$($_.ErrorDetails.message)]"
        Write-Verbose "Exception Message: [$($_.Exception.Message)]" 
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function ConvertTo-ReturnXmlToLeerlingenlist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Xml.xmlelement]
        $ReturnNode
    )

    $leerlingenObject = [System.Collections.ArrayList]::new()

    $adressenNode = $returnNode.adressen
    $groepenNode = $returnNode.groepen
    $schooljarenNode = $returnNode.schooljaren
    $leerlingenNode = $returnNode.leerlingen
    $inschrijvingtypesNode = $returnNode.inschrijvingtypes
    $nodePath = "leerling"
    $leerlingNodeList = $leerlingenNode.SelectNodes($nodePath)

    foreach ($leerlingNode in $leerlingNodeList) {
        
        $contracts = [System.Collections.ArrayList]::new();

        $nodePath = "adres[id=`'" + $leerlingNode.leerlingAdres + "`']"
        $adresNode = $adressenNode.SelectSingleNode($nodePath)
        $adres = @{ id = $leerlingNode.leerlingAdres }
        if ($null -ne $adresNode) {
            if ( $adresNode.geheimadres -eq "false") {
                $telefoon = @{}
                if ($adresNode.telefoon.geheim -eq "false") {
                    $telefoon = @{
                        id     = $adresNode.telefoon.id
                        nummer = $adresNode.telefoon.nummer
                    }
                }
                $adres = @{
                    id         = $adresNode.id
                    gemeente   = $adresNode.gemeente
                    plaats     = $adresNode.plaats
                    straat     = $adresNode.straat
                    huisnummer = $adresNode.huisnummer
                    postcode   = $adresNode.postcode
                    land       = $adresNode.land
                    telefoon   = $telefoon;
                }
            }
        }

        $schooltype = "SO"

        $nodePath = "groepsindelingen/groepsindeling"
        $groepsIndelingNodeList = $leerlingNode.SelectNodes($nodePath)
		
        $aangemeld = 'false'
		
        foreach ($groepsIndelingNode in $groepsIndelingNodeList ) {
            $nodePath = "groep[id=`'" + $groepsIndelingNode.groep + "`']"
            $groepNode = $groepenNode.SelectSingleNode($nodePath)
            $groep = @{id = $groepsIndelingNode.groep }
            if ($null -ne $groepNode) {
                $nodePath = "schooljaar[id=`'" + $groepNode.schooljaar + "`']"
                $groepschooljaarNode = $schooljarenNode.SelectSingleNode($nodePath)
                $groepschooljaar = @{id = $groepNode.schooljaar }
                if ($null -ne $groepschooljaarNode) {
                    $groepschooljaar = @{
                        id   = $groepschooljaarNode.id
                        naam = $groepschooljaarNode.naam
                    }
                }
                $groep = @{
                    id         = $groepNode.id
                    naam       = $groepNode.naam
                    lokaal     = $groepNode.lokaal
                    schooljaar = $groepschooljaar
                }
                
            }
			
            if ($groep.naam -eq "aangemeld") {
                $aangemeld = $true
            }

            $nodePathOpleidinggegevens = "opleidinggegevens"
            $opleidinggegevensNode = $groepsIndelingNode.SelectNodes($nodePathOpleidinggegevens)
            
            if (-not([string]::IsNullOrEmpty($opleidinggegevensNode.opleiding)) -and -not([string]::IsNullOrEmpty($opleidinggegevensNode.opleiding.type))) {
                $schooltype = $opleidinggegevensNode.opleiding.type

            }
            elseif (-not([string]::IsNullOrEmpty($opleidinggegevensNode.voortgezet)) -and $opleidinggegevensNode.voortgezet -eq 'true') {
                $schooltype = "VSO"

            }

            $nodePath = "schooljaar[id=`'" + $groepsIndelingNode.schooljaar + "`']"
            $schooljaarNode = $schooljarenNode.SelectSingleNode($nodePath)
            $schooljaar = @{id = $groepsIndelingNode.schooljaar }
            if ($null -ne $schooljaarNode) {
                $schooljaar = @{
                    id   = $schooljaarNode.id
                    naam = $schooljaarNode.naam
                }
            }
            $contract = @{
                ContractType     = "groep"
                id               = $groepsIndelingNode.id
                vanafDatum       = ([System.DateTimeOffset]$groepsIndelingNode.vanafDatum).ToString("yyyy-MM-dd")
                totDatum         = ([System.DateTimeOffset]$groepsIndelingNode.totDatum).ToString("yyyy-MM-dd")
                groep            = $groep
                schooljaar       = $schooljaar
                leerjaar         = $groepsIndelingNode.leerjaar
                bekostigd        = $groepsIndelingNode.bekostigd
                inschrijvingType = @{} #dummy voor de mapping
                dienstverband    = @{} #dummy voor mapping
                Brin             = $Brin
                Voortgezet       = $opleidinggegevensNode.voortgezet
                SchoolType       = $opleidinggegevensNode.opleiding.type
            }
            $contracts += $contract;
        }
        $nodePath = "inschrijvingen/inschrijving"
        $inschrijvingNodeList = $leerlingNode.SelectNodes($nodePath)

        foreach ($inschrijvingNode in $inschrijvingNodeList ) {
            $nodePath = "inschrijvingtype[id=`'" + $inschrijvingNode.inschrijvingType + "`']"
            $inschrijvingtypeNode = $inschrijvingtypesNode.SelectSingleNode($nodePath)
            $inschrijvingType = @{ id = $inschrijvingNode.inschrijvingType }
            if ($null -ne $inschrijvingtypeNode) {
                $inschrijvingType = @{
                    id           = $inschrijvingtypeNode.id
                    omschrijving = $inschrijvingtypeNode.omschrijving
                    code         = $inschrijvingtypeNode.code
                }
            }

            $datumInschrijving = "";
            if (-not([string]::IsNullOrEmpty($inschrijvingNode.datumInschrijving))) {
                $datumInschrijving = ([System.DateTimeOffset]$inschrijvingNode.datumInschrijving).ToString("yyyy-MM-dd")
            }

            $datumUitschrijving = "";
            if (-not([string]::IsNullOrEmpty($inschrijvingNode.datumUitschrijving))) {
                $datumUitschrijving = ([System.DateTimeOffset]$inschrijvingNode.datumUitschrijving).ToString("yyyy-MM-dd")
            }
			
            $contract = @{
                ContractType      = "inschrijving"
                id                = $inschrijvingNode.id
                datumInschrijving = ([System.DateTimeOffset]$inschrijvingNode.datumInschrijving).ToString("yyyy-MM-dd")
                vanafDatum        = ([System.DateTimeOffset]$inschrijvingNode.datumInschrijving).ToString("yyyy-MM-dd")  #copy to ease the mapping
                totDatum          = $datumUitschrijving
                inschrijvingType  = $inschrijvingType
                groep             = @{} #dummy for mapping
                dienstverband     = @{} #dummy for mapping
                Brin              = $Brin
                SchoolType        = $schooltype
            }
            $contracts += $contract
        }

        $geboortedatum = "";
        if (-not([string]::IsNullOrEmpty($leerlingNode.geboortedatum))) {
            $geboortedatum = ([System.DateTimeOffset]$leerlingNode.geboortedatum).ToString("yyyy-MM-dd")
        }

        $leerlingObject = @{
            PersonType             = "leerling"
            Brin                   = $Brin
            ExternalId             = [string]($Brin + "_" + $leerlingNode.leerlingNummer)
            DisplayName            = $leerlingNode.roepNaam + " " + $leerlingNode.achternaam

            achternaam             = $leerlingNode.achternaam
            achternaamOfficieel    = $leerlingNode.achternaamOfficieel
            adres                  = $adres
            Contracts              = $contracts
            datumAanmelding        = $leerlingNode.datumAanmelding
            geboortedatum          = $geboortedatum
            geboortedatumOnzeker   = $leerlingNode.geboortedatumOnzeker
            geboorteplaats         = $leerlingNode.geboorteplaats
            geslacht               = $leerlingNode.geslacht
            id                     = $leerlingNode.id
            leerlingNummer         = $leerlingNode.leerlingNummer
            roepnaam               = $leerlingNode.roepNaam
            tussenvoegsel          = $leerlingNode.tussenvoegsel
            tussenvoegselOfficieel = $leerlingNode.tussenvoegselOfficieel
            voornamen              = $leerlingNode.voornamen
            telefoonWerk           = @{} #dummy for mapping
            SchoolType             = $schooltype
            Aangemeld              = $aangemeld

        }
        $null = $leerlingenObject.add($leerlingObject);
    }
    Write-Output $leerlingenObject
}
#endregion functions

# start of main loop of script execution
$personList = [System.Collections.Generic.List[object]]::new()

$brinNumbers = [array]$config.brinIdentifiers.split(',') | ForEach-Object { $_.trim(' ') }

foreach ($Brin in $brinNumbers) {
    # If no school year specified, getting the current Year.
    $schoolYear = $config.schoolYear
    if ([string]::IsNullOrEmpty($config.schoolYear)) {
        if ((Get-Date).Month -lt 8) {
            $schoolYear = (Get-Date).AddYears(-1).ToString("yyyy") + " / " + (Get-Date).ToString("yyyy")
        }
        else {
            $schoolYear = (Get-Date).ToString("yyyy") + " / " + (Get-Date).AddYears(1).ToString("yyyy")
        }
    }

    Write-Verbose "ParnasSys import students getting data for brin [$Brin] and year [$SchoolYear]"

    $splatParnasSys = @{
        Brin          = $Brin
        WebServiceUri = $config.webServiceUri
        SupplierName  = $config.supplierName
        SupplierKey   = $config.supplierKey
        schoolYear    = $schoolYear
    }

    if (-not [string]::IsNullOrEmpty($config.Proxy)) {
        Write-Verbose "Added Proxy Address to webrequest $($config.Proxy)"
        $splatParnasSys['Proxy'] = $config.Proxy
    }

    $leerLingen , $leerLingenReturnNode = @()   # Needed for the second Brin
    $leerLingenReturnNode = Get-ParnasSysLeerlingen @splatParnasSys

    $leerLingen = ConvertTo-ReturnXmlToLeerlingenlist -ReturnNode $leerLingenReturnNode

    $array = @()
    $array = $array + $leerLingen
    $leerLingen = $array

        
    if (($leerLingen | measure-object).count -gt 0) {
        $personList.AddRange($leerLingen)
    }
    Write-Verbose "Students found [$(($leerLingen | measure-object).Count)] for Brin [$Brin] and year [$schoolyear]" -verbose

    # Get previous schoolyear
    if ($config.includeLastYear) {
        $schoolYearSplitted = $schoolYear -split '/'
        $startYear = $schoolYearSplitted[0].Trim()
        $endYear = $schoolYearSplitted[1].Trim()
        $previousStart = $startYear - 1
        $previousEnd = $endYear - 1
        $schoolYearPrevious = "$previousStart / $previousEnd"

        Write-Verbose "ParnasSys import students getting data for brin [$Brin] and year [$schoolYearPrevious]"

        $splatParnasSys = @{
            Brin          = $Brin
            WebServiceUri = $config.webServiceUri
            SupplierName  = $config.supplierName
            SupplierKey   = $config.supplierKey
            schoolYear    = $schoolYearPrevious
        }

        if (-not [string]::IsNullOrEmpty($config.Proxy)) {
            Write-Verbose "Added Proxy Address to webrequest $($config.Proxy)"
            $splatParnasSys['Proxy'] = $config.Proxy
        }
        $leerLingenPreviousYear , $leerLingenPreviousYearReturnNode = @()  # Needed for the second Brin
        $leerLingenPreviousYearReturnNode = Get-ParnasSysleerLingen @splatParnasSys
        $leerLingenPreviousYear = ConvertTo-ReturnXmlToleerLingenlist -ReturnNode $leerLingenPreviousYearReturnNode
    
        $array = @()
        $array = $array + $leerLingenPreviousYear
        $leerLingenPreviousYear = $array

        if ( ($leerLingenPreviousYear | measure-object).count -gt 0) {
            $leerLingenPreviousYear = $leerLingenPreviousYear | Where-Object { $_.id -notin $personList.id }
        }
        
        if (($leerLingenPreviousYear | measure-object).count -eq 1 ) {
            $personList.add($leerLingenPreviousYear)
        }
        elseif (($leerLingenPreviousYear | measure-object).count -gt 0) {
            $personList.AddRange($leerLingenPreviousYear)
        }
        Write-Verbose "Students found [$(($leerLingenPreviousYear | measure-object).count)] for Brin [$Brin] and year [$schoolYearPrevious]" -Verbose
    }
}
Write-Verbose "Total Students found [$(($personList | measure-object).Count)]" -Verbose
Write-Output $personList | ConvertTo-Json -Depth 10
<#
foreach ($person in $personList)
{
    Write-Output $person | ConvertTo-Json -Depth 10
}
#>