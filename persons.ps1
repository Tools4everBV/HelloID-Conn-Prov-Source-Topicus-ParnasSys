$config = $configuration | ConvertFrom-Json

function Get-ParnasSysMedewerkers {
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
    Write-Verbose -Verbose "ParnasSys import Employees getting data of shoolyear $SchoolYear"
    try {
        $headers = @{
            'Content-Type' = 'text/xml; charset=utf-8'
            SOAPaction     = "`"getMedewerkers`""
        }
        $supplierNameEncoded = [System.Net.WebUtility]::HtmlEncode($SupplierName)

        $xml = [xml]( '<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema">
                <soap:Body>
                    <getMedewerkers xmlns="http://www.topicus.nl/parnassys">
                        <leveranciernaam xmlns="">{0}</leveranciernaam>
                        <leveranciersleutel xmlns="">{1}</leveranciersleutel>
                        <brinnummer xmlns="">{2}</brinnummer>
                        <schooljaar xmlns="">{3}</schooljaar>
                    </getMedewerkers>
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

        Write-Verbose  'ParnasSys import medewerkers Invoking webRequest start ' -Verbose
        $result = Invoke-WebRequest @splatWebRequestParameters
        Write-Verbose  'ParnasSys import medewerkers Invoking webRequest finished' -Verbose

        [xml] $parnasSysDataxml = $result.Content
        # envelope/body/getmedewerkersresponse/return/leerlingen
        $medewerkersResponseNode = $parnasSysDataxml.FirstChild.FirstChild.FirstChild
        $returnNode = $medewerkersResponseNode.item('return')
        Write-Output $returnNode

    } catch {
        Write-Verbose "Could not get employees for Brin: [$brin]" -Verbose
        Write-Verbose "Error Details: [$($_.ErrorDetails.message)]" -Verbose
        Write-Verbose "Exception Message: [$($_.Exception.Message)]" -Verbose
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function ConvertTo-ReturnxmlToMedewerkerslist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Xml.xmlelement]
        $ReturnNode
    )

    $medewerkersObject = [System.Collections.ArrayList]::new()

    $adressenNode = $returnNode.item("adressen")
    $medewerkersNode = $returnNode.item("medewerkers")
    $dienstverbandenNode = $returnNode.item("dienstverbanden")

    $nodePath = "medewerker"
    $medewerkerNodeList = $medewerkersNode.SelectNodes($nodePath)

    foreach ($medewerkerNode in $medewerkerNodeList) {

        $contracts = [System.Collections.Generic.list[object]]::new()

        $mobile = @{}
        if ($null -ne $medewerkerNode.item("mobile")) {
            if ($medewerkerNode.item("mobile").item("geheim").FirstChild.Value -eq "false") {
                $mobile = @{
                    id     = $medewerkerNode.item("mobile").item("id").FirstChild.Value
                    nummer = $medewerkerNode.item("mobile").item("nummer").FirstChild.Value
                }
            }
        }

        $telefoonWerk = @{}
        if ($null -ne $medewerkerNode.item("telefoonWerk")) {
            if ($medewerkerNode.item("telefoonWerk").item("geheim").FirstChild.Value -eq "false") {
                $telefoonWerk = @{
                    id     = $medewerkerNode.item("telefoonWerk").item("id").FirstChild.Value
                    nummer = $medewerkerNode.item("telefoonWerk").item("nummer").FirstChild.Value
                }
            }
        }

        $nodePath = "dienstverbanden/dienstverband"
        $medewerker_dienstverbandNodeList = $medewerkerNode.SelectNodes($nodePath)
        foreach ($medewerker_dienstverbandNode in $medewerker_dienstverbandNodeList ) {
            $nodePath = "dienstverband[id=`'" + $medewerker_dienstverbandNode.FirstChild.Value + "`']"
            $dienstverbandNode = $dienstverbandenNode.SelectSingleNode($nodePath)
            $dienstverband = @{id = $dienstverbandNode.FirstChild.Value }
            if ($null -ne $dienstverbandNode) {
                $dienstverband = @{
                    id        = $dienstverbandNode.Item("id").FirstChild.Value
                    afkorting = $dienstverbandNode.Item("afkorting").FirstChild.Value
                    naam      = $dienstverbandNode.Item("afkorting").FirstChild.Value
                }
            }

            $contract = @{
                ContractType     = "dienstverband"
                dienstverband    = $dienstverband
                groep            = @{} #dummy voor mapping
                inschrijvingType = @{} #dummy voor de mapping
            }
            $contracts += $contract

        }

        $medewerkerObject = @{
            PersonType           = "medewerker"
            Brin                 = $Brin
            ExternalId           = [string] ($Brin + "_" + $medewerkerNode.item("id").FirstChild.Value)
            DisplayName          = $medewerkerNode.item("roepNaam").FirstChild.Value + $medewerkerNode.item("achternaam").FirstChild.Value

            aanhef               = $medewerkerNode.item("aanhef").FirstChild.Value
            achternaam           = $medewerkerNode.item("lastName").FirstChild.Value
            adres                = $adres
            Contracts            = $contracts
            email                = $medewerkerNode.item("email").FirstChild.Value
            mobile               = $mobile
            roepnaam             = $medewerkerNode.item("roepNaam").FirstChild.Value
            voornamen            = $medewerkerNode.item("firstName").FirstChild.Value #there is no voornamen in medewerker
            telefoonWerk         = $telefoonWerk
            voornaam             = $medewerkerNode.item("firstName").FirstChild.Value
            geboortedatum        = $medewerkerNode.item("geboortedatum").FirstChild.Value
            geboortedatumOnzeker = $medewerkerNode.item("geboortedatumOnzeker").FirstChild.Value
            geboorteland         = $medewerkerNode.item("geboorteland").FirstChild.Value
        }
        $dummyIndex = $medewerkersObject.add($medewerkerObject)
    }
    return $medewerkersObject
}

# start of main loop of script execution
$personList = [System.Collections.ArrayList]::new()

$brinNumbers = [array]$config.brinIdentifiers.split(',') | ForEach-Object { $_.trim(' ') }
foreach ($Brin in $brinNumbers) {
    Write-Verbose -Verbose "ParnasSys import Employees looping Brins ($Brin)"
    $schoolYear = $config.schoolYear
    if ([string]::IsNullOrEmpty($config.schoolYear)) {
        if ((Get-Date).Month -lt 8) {
            $schoolYear = (Get-Date).AddYears(-1).ToString("yyyy") + " / " + (Get-Date).ToString("yyyy")
        } else {
            $schoolYear = (Get-Date).ToString("yyyy") + " / " + (Get-Date).AddYears(1).ToString("yyyy")
        }
    }
    $splatParnasSys = @{
        Brin          = $Brin
        WebServiceUri = $config.webServiceUri
        SupplierName  = $config.supplierName
        SupplierKey   = $config.supplierKey
        schoolYear    = $schoolYear
    }

    if (-not [string]::IsNullOrEmpty($config.Proxy)) {
        Write-Verbose "Added Proxy Address to webrequest $($config.Proxy)" -Verbose
        $splatParnasSys['Proxy'] = $config.Proxy
    }
    $medewerkers , $medewerkersReturnNode = $null  # Needed for the second Brin

    $medewerkersReturnNode = Get-ParnasSysMedewerkers @splatParnasSys

    $medewerkers = ConvertTo-ReturnxmlToMedewerkerslist -returnNode $medewerkersReturnNode
    if ( $medewerkers.count -gt 0) {
        $dummyIndex = $personList.AddRange($medewerkers)
    }
    Write-Verbose "Employees found [$($medewerkers.Count)] for Brin [$Brin]" -Verbose
}

Write-Verbose -Verbose "ParnasSys import employees succesfull"
Write-Verbose "Total Employees found [$($personList.Count)]" -Verbose
Write-Output $personList | ConvertTo-Json -Depth 10


