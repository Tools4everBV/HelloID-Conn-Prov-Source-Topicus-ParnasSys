# Connection definition
$ConnectorSettings = @{
    webServiceUri = "";
    supplierName = "";
    supplierKey = ""
    proxy = "";
    # Data collection options
    brinIdentifiers = @('');
}

function Get-ParnasSysLeerlingen{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]
        $Brin,

        [Parameter(Mandatory=$true)]
        [string]
        $WebServiceUri,

        [Parameter(Mandatory=$true)]
        [string]
        $SupplierName,

        [Parameter(Mandatory=$true)]
        [string]
        $SupplierKey,

        [Parameter(Mandatory=$true)]
        [string]
        $SchoolYear,

        [Parameter(Mandatory=$false)]

        [string]
        $proxy = ""
    )
    Write-Verbose -Verbose "ParnasSys import leerlingen getting data of shoolyear $SchoolYear";
    try{
        $headers = @{
            'Content-Type' = "text/xml; charset=utf-8"
            SOAPaction = "`"getLeerlingen`""
        }
        # $supplierNameEncoded =[System.Web.HttpUtility]::HtmlEncode($SupplierName)
        $supplierNameEncoded = $SupplierName
        $body = "<?xml version=`"1.0`" encoding=`"utf-8`"?><soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`"><soap:Body><getLeerlingen xmlns=`"http://www.topicus.nl/parnassys`"><leveranciernaam xmlns=`"`">$SupplierNameEncoded</leveranciernaam><leveranciersleutel xmlns=`"`">$SupplierKey</leveranciersleutel><brinnummer xmlns=`"`">$Brin</brinnummer><schooljaar xmlns=`"`">$SchoolYear</schooljaar></getLeerlingen></soap:Body></soap:Envelope>"     

        if ($Proxy -ne "")
        {
            $splatWebRequestParameters = @{
                Uri = $webServiceUri
                Method = 'Post'
                Headers = $headers
                Proxy = $proxy
                UseBasicParsing = $true
                Body = $body;
            }
        }
        else {
            $splatWebRequestParameters = @{
                Uri = $webServiceUri
                Method = 'Post'
                Headers = $headers
                UseBasicParsing = $true
                Body = $body;
            }
        }

        Write-Verbose -Verbose "ParnasSys import leerlingen Invoking webRequest start "; 
        $result = Invoke-WebRequest @splatWebRequestParameters  
        Write-Verbose -Verbose "ParnasSys import leerlingen Invoking webRequest finished";
        [xml] $parnasSysDataxml = $result.Content 
        # envelope/body/getleerlingenresponse/return/leerlingen
        $leerLingenResponseNode= $parnasSysDataxml.FirstChild.FirstChild.FirstChild
        $returnNode = $leerLingenResponseNode.item("return")
    }catch{
        throw $_
    }
    return $returnNode;
}


function Convert_Returnxml_to_Leerlingenlist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.xmlelement]
        $ReturnNode
    )

    $leerlingenObject =  [System.Collections.ArrayList]::new()

    $adressenNode = $returnNode.item("adressen")
    $groepenNode = $returnNode.Item("groepen")
    $schooljarenNode = $returnNode.Item("schooljaren")
    $leerlingenNode = $returnNode.item("leerlingen")
    $inschrijvingtypesNode = $returnNode.item("inschrijvingtypes")
    $nodePath = "leerling" 
    $leerlingNodeList = $leerlingenNode.SelectNodes($nodePath)

    foreach($leerlingNode in $leerlingNodeList){ 
        $contracts = [System.Collections.ArrayList]::new();

        $nodePath = "adres[id=`'" + $leerlingNode.item("leerlingAdres").FirstChild.Value  + "`']"
        $adresNode = $adressenNode.SelectSingleNode($nodePath) 
        $adres = @{ id =  $leerlingNode.item("leerlingAdres").FirstChild.Value }
        if ($null -ne $adresNode){  

            if ( $adresNode.item("geheimadres").FirstChild.Value -eq "false"){

                $telefoon = @{}
                if ($adresNode.item("telefoon").item("geheim").FirstChild.Value -eq "false")
                {
                    $telefoon = @{
                        id = $adresNode.item("telefoon").item("id").FirstChild.Value;
                        nummer = $adresNode.item("telefoon").item("nummer").FirstChild.Value;
                    }
                }
                $adres = @{
                    id =  $adresNode.item("id").FirstChild.Value;
                    gemeente =  $adresNode.item("gemeente").FirstChild.Value;
                    plaats = $adresNode.item("plaats").FirstChild.Value;
                    straat = $adresNode.item("straat").FirstChild.Value;
                    huisnummer = $adresNode.item("huisnummer").FirstChild.Value;
                    postcode = $adresNode.item("postcode").FirstChild.Value;
                    land = $adresNode.item("land").FirstChild.Value;
                    telefoon = $telefoon;
                }
            }
        }

        $nodePath = "groepsindelingen/groepsindeling" 
        $groepsIndelingNodeList = $leerlingNode.SelectNodes($nodePath)

        foreach($groepsIndelingNode in $groepsIndelingNodeList )
        {
            $nodePath = "groep[id=`'" + $groepsIndelingNode.Item("groep").FirstChild.Value + "`']"
            $groepNode = $groepenNode.SelectSingleNode($nodePath) 
            $groep = @{id = $groepsIndelingNode.Item("groep").FirstChild.Value}
            if ($null -ne $groepNode)
            {
                $nodePath = "schooljaar[id=`'" + $groepNode.item("schooljaar").FirstChild.Value + "`']"
                $groepschooljaarNode = $schooljarenNode.SelectSingleNode($nodePath) 
                $groepschooljaar = @{id = $groepNode.item("schooljaar").FirstChild.Value }
                if ($null -ne  $groepschooljaarNode)
                {
                    $groepschooljaar = @{
                        id =  $groepschooljaarNode.item("id").FirstChild.Value
                        naam =  $groepschooljaarNode.item("naam").FirstChild.Value
                    }
                }
                $groep = @{
                    id = $groepNode.item("id").FirstChild.Value;
                    naam = $groepNode.item("naam").FirstChild.Value;
                    lokaal = $groepNode.item("lokaal").FirstChild.Value;
                    schooljaar = $groepschooljaar
                }
            }

            $nodePath = "schooljaar[id=`'" + $groepsIndelingNode.Item("schooljaar").FirstChild.Value + "`']"
            $schooljaarNode = $schooljarenNode.SelectSingleNode($nodePath) 
            $schooljaar = @{id =  $groepsIndelingNode.Item("schooljaar").FirstChild.Value}
            if ($null -ne  $schooljaarNode)
            {
                $schooljaar = @{
                    id =  $schooljaarNode.item("id").FirstChild.Value
                    naam =  $schooljaarNode.item("naam").FirstChild.Value
                }
            }
            $contract = @{
                ContractType = "groep"
                id = $groepsIndelingNode.Item("id").FirstChild.Value
                vanafDatum =  $groepsIndelingNode.Item("vanafDatum").FirstChild.Value
                totDatum = $groepsIndelingNode.Item("totDatum").FirstChild.Value
                groep = $groep
                schooljaar = $schooljaar
                leerjaar = $groepsIndelingNode.Item("leerjaar").FirstChild.Value
                bekostigd = $groepsIndelingNode.Item("bekostigd").FirstChild.Value
                inschrijvingType = @{} #dummy voor de mapping
                dienstverband = @{} #dummy voor mapping
            }
            $contracts += $contract;
        }
        $nodePath = "inschrijvingen/inschrijving" 
        $inschrijvingNodeList = $leerlingNode.SelectNodes($nodePath)

        foreach($inschrijvingNode in $inschrijvingNodeList )
        {
            $nodePath = "inschrijvingtype[id=`'" +$inschrijvingNode.Item("inschrijvingType").FirstChild.Value + "`']" 
            $inschrijvingtypeNode = $inschrijvingtypesNode.SelectSingleNode($nodePath)
            $inschrijvingType = @{ id = $inschrijvingNode.Item("inschrijvingType").FirstChild.Value }
            if ($null -ne $inschrijvingtypeNode)
            {
                $inschrijvingType = @{
                    id = $inschrijvingtypeNode.Item("id").FirstChild.Value 
                    omschrijving = $inschrijvingtypeNode.Item("omschrijving").FirstChild.Value 
                    code =  $inschrijvingtypeNode.Item("code").FirstChild.Value                   
                }
            }
            $contract = @{
                ContractType = "inschrijving"
                id = $inschrijvingNode.Item("id").FirstChild.Value
                datumInschrijving =  $inschrijvingNode.Item("datumInschrijving").FirstChild.Value 
                vanafDatum = $inschrijvingNode.Item("datumInschrijving").FirstChild.Value  #copy to ease the mapping
                inschrijvingType = $inschrijvingType
                groep = @{} #dummy for mapping
                dienstverband = @{} #dummy for mapping
            }
            $contracts += $contract
        }

        $leerlingObject = @{
            PersonType = "leerling"
            Brin = $Brin
            ExternalId = [string] ($Brin + "_" + $leerlingNode.item("id").FirstChild.Value)
            DisplayName = $leerlingNode.item("roepNaam").FirstChild.Value +  $leerlingNode.item("achternaam").FirstChild.Value;

            achternaam = $leerlingNode.item("achternaam").FirstChild.Value;
            achternaamOfficieel = $leerlingNode.item("achternaamOfficieel").FirstChild.Value;
            adres = $adres 
            Contracts = $contracts   
            datumAanmelding = $leerlingNode.item("datumAanmelding").FirstChild.Value;
            geboortedatum = $leerlingNode.item("geboortedatum").FirstChild.Value;
            geboortedatumOnzeker = $leerlingNode.item("geboortedatumOnzeker").FirstChild.Value;
            geboorteplaats = $leerlingNode.item("geboorteplaats").FirstChild.Value;
            geslacht = $leerlingNode.item("geslacht").FirstChild.Value;          
            id = $leerlingNode.item("id").FirstChild.Value;
            leerlingNummer = $leerlingNode.item("leerlingNummer").FirstChild.Value;        
            roepnaam = $leerlingNode.item("roepNaam").FirstChild.Value;
            tussenvoegsel = $leerlingNode.item("tussenvoegsel").FirstChild.Value;
            tussenvoegselOfficieel = $leerlingNode.item("tussenvoegselOfficieel").FirstChild.Value;
            voornamen = $leerlingNode.item("voornamen").FirstChild.Value;
            telefoonWerk = @{} #dummy for mapping

        }
        $null = $leerlingenObject.add($leerlingObject);
    }
    return $leerlingenObject
}


# start of main loop of script execution
$personList = [System.Collections.ArrayList]::new()
foreach ($Brin in $Connectorsettings.brinIdentifiers)
{
    Write-Verbose -Verbose "ParnasSys import leerlingen looping Brins ($Brin)";

    if ((Get-Date).Month -lt 8) {
        $schoolYear = (Get-Date).AddYears(-1).ToString("yyyy") + " / " + (Get-Date).ToString("yyyy")
    } else {
        $schoolYear = (Get-Date).ToString("yyyy") + " / " + (Get-Date).AddYears(1).ToString("yyyy")
    }
    $leerLingenReturnNode = Get-ParnasSysLeerlingen `
                            -Brin $Brin `
                            -WebServiceUri  $Connectorsettings.webServiceUri `
                            -SupplierName   $Connectorsettings.supplierName `
                            -SupplierKey    $Connectorsettings.supplierKey `
                            -schoolYear     $schoolYear `
                            -Proxy $Connectorsettings.Proxy

    $leerLingen = Convert_Returnxml_to_Leerlingenlist -returnNode $leerLingenReturnNode;
    $dummyIndex = $personList.AddRange($leerLingen);
}

Write-Verbose -Verbose "ParnasSys import leerlingen succesfull";
$i=0
foreach ($person in $personList)
{
    Write-Output $person | ConvertTo-json -Depth 10
    $i = $i + 1;
    if ($i -gt 20000 )
    {break;}
}
