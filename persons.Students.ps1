$config = $configuration | ConvertFrom-Json

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
        $supplierNameEncoded = [System.Web.HttpUtility]::HtmlEncode($SupplierName)

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
    } catch {
        Write-Verbose "Could not get Students for Brin: [$brin]" -Verbose
        Write-Verbose "Error Details: [$($_.ErrorDetails.message)]" -Verbose
        Write-Verbose "Exception Message: [$($_.Exception.Message)]" -Verbose
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

        $nodePath = "groepsindelingen/groepsindeling"
        $groepsIndelingNodeList = $leerlingNode.SelectNodes($nodePath)
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
                    lokaal     = $groepNode.code
                    schooljaar = $groepschooljaar
                }
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
                vanafDatum       = $groepsIndelingNode.vanafDatum
                totDatum         = $groepsIndelingNode.totDatum
                groep            = $groep
                schooljaar       = $schooljaar
                leerjaar         = $groepsIndelingNode.leerjaar
                bekostigd        = $groepsIndelingNode.bekostigd
                inschrijvingType = @{} #dummy voor de mapping
                dienstverband    = @{} #dummy voor mapping
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
            $contract = @{
                ContractType      = "inschrijving"
                id                = $inschrijvingNode.id
                datumInschrijving = $inschrijvingNode.datumInschrijving
                vanafDatum        = $inschrijvingNode.datumInschrijving  #copy to ease the mapping
                inschrijvingType  = $inschrijvingType
                groep             = @{} #dummy for mapping
                dienstverband     = @{} #dummy for mapping
            }
            $contracts += $contract
        }

        $leerlingObject = @{
            PersonType             = "leerling"
            Brin                   = $Brin
            ExternalId             = [string]($Brin + "_" + $leerlingNode.id)
            DisplayName            = $leerlingNode.roepNaam + $leerlingNode.achternaam

            achternaam             = $leerlingNode.achternaam
            achternaamOfficieel    = $leerlingNode.achternaamOfficieel
            adres                  = $adres
            Contracts              = $contracts
            datumAanmelding        = $leerlingNode.datumAanmelding
            geboortedatum          = $leerlingNode.geboortedatum
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

        }
        $null = $leerlingenObject.add($leerlingObject);
    }
    Write-Output $leerlingenObject
}
#endregion functions

# start of main loop of script execution
$personList = [System.Collections.Generic.List[object]]::new()

$birnNumbers = [array]$config.brinIdentifiers.split(',') | ForEach-Object { $_.trim(' ') }
foreach ($Brin in $birnNumbers) {
    Write-Verbose "ParnasSys import students processing brin [$Brin]" -Verbose

    # If no school year specified, getting the current Year.
    $schoolYear = $config.schoolYear
    if ([string]::IsNullOrEmpty($config.schoolYear)) {
        if ((Get-Date).Month -lt 8) {
            $schoolYear = (Get-Date).AddYears(-1).ToString("yyyy") + " / " + (Get-Date).ToString("yyyy")
        } else {
            $schoolYear = (Get-Date).ToString("yyyy") + " / " + (Get-Date).AddYears(1).ToString("yyyy")
        }
    }
    Write-Verbose "ParnasSys import students getting data of shoolyear $SchoolYear" -Verbose
    $spaltParnasSys = @{
        Brin          = $Brin
        WebServiceUri = $config.webServiceUri
        SupplierName  = $config.supplierName
        SupplierKey   = $config.supplierKey
        schoolYear    = $schoolYear
    }

    if (-not [string]::IsNullOrEmpty($config.Proxy)) {
        Write-Verbose "Added Proxy Address to webrequest $($config.Proxy)" -Verbose
        $spaltParnasSys['Proxy'] = $config.Proxy
    }

    $leerLingen , $leerLingenReturnNode = $null   # Needed for the second Brin
    $leerLingenReturnNode = Get-ParnasSysLeerlingen @spaltParnasSys

    $leerLingen = ConvertTo-ReturnXmlToLeerlingenlist -ReturnNode $leerLingenReturnNode
    if ( $leerLingen.count -gt 0) {
        $personList.AddRange($leerLingen)
    }
    Write-Verbose "Students found [$($leerLingen.Count)] for Brin [$Brin]" -Verbose
}
Write-Verbose "Total Students found [$($personList.Count)]" -Verbose
Write-Output $personList | ConvertTo-Json -Depth 10
