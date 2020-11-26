# Connection definition
$ConnectorSettings = @{
    webServiceUri = "";
    supplierName = "";
    supplierKey = ""
    proxy = "";
    # Data collection options
    brinIdentifiers = @('');
}

function Get-ParnasSysMedewerkers{
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
    Write-Verbose -Verbose "ParnasSys import medewerkers getting data of shoolyear $SchoolYear";
    try{
        $headers = @{
            'Content-Type' = "text/xml; charset=utf-8"
            SOAPaction = "`"getMedewerkers`""
        }
        # $supplierNameEncoded =[System.Web.HttpUtility]::HtmlEncode($SupplierName)
        $supplierNameEncoded = $SupplierName
        $body = "<?xml version=`"1.0`" encoding=`"utf-8`"?><soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`"><soap:Body><getMedewerkers xmlns=`"http://www.topicus.nl/parnassys`"><leveranciernaam xmlns=`"`">$SupplierNameEncoded</leveranciernaam><leveranciersleutel xmlns=`"`">$SupplierKey</leveranciersleutel><brinnummer xmlns=`"`">$Brin</brinnummer><schooljaar xmlns=`"`">$SchoolYear</schooljaar></getMedewerkers></soap:Body></soap:Envelope>"

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

        Write-Verbose -Verbose "ParnasSys import medewerkers Invoking webRequest start ";
        $result = Invoke-WebRequest @splatWebRequestParameters  
        Write-Verbose -Verbose "ParnasSys import medewerkers Invoking webRequest finished";
        [xml] $parnasSysDataxml = $result.Content 
        # envelope/body/getleerlingenresponse/return/leerlingen
        $medewerkersResponseNode= $parnasSysDataxml.FirstChild.FirstChild.FirstChild
        $returnNode = $medewerkersResponseNode.item("return")

    }catch{
        throw $_
    }
    return $returnNode;
}

function Convert_Returnxml_to_Medewerkerlist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.xmlelement]
        $ReturnNode
    )

    $medewerkersObject =  [System.Collections.ArrayList]::new()

    $adressenNode = $returnNode.item("adressen")   
    $medewerkersNode = $returnNode.item("medewerkers")
    $dienstverbandenNode = $returnNode.item("dienstverbanden")

    $nodePath = "medewerker" 
    $medewerkerNodeList = $medewerkersNode.SelectNodes($nodePath)

    foreach($medewerkerNode in $medewerkerNodeList){

        $contracts = [System.Collections.ArrayList]::new();

        $nodePath = "adres[id=`'" + $medewerkerNode.item("adres").FirstChild.Value  + "`']"
        $adresNode = $adressenNode.SelectSingleNode($nodePath) 
        $adres = @{ id =  $medewerkerNode.item("adres").FirstChild.Value }
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

        $mobile = @{}
        if ($null -ne $medewerkerNode.item("mobile")){
            if ($medewerkerNode.item("mobile").item("geheim").FirstChild.Value -eq "false") {
                $mobile = @{
                    id = $medewerkerNode.item("mobile").item("id").FirstChild.Value;
                    nummer = $medewerkerNode.item("mobile").item("nummer").FirstChild.Value;
                }
            }
        }

        $telefoonWerk = @{}
        if ($null -ne $medewerkerNode.item("telefoonWerk")){
            if ($medewerkerNode.item("telefoonWerk").item("geheim").FirstChild.Value -eq "false") {
                $telefoonWerk = @{
                    id = $medewerkerNode.item("telefoonWerk").item("id").FirstChild.Value;
                    nummer = $medewerkerNode.item("telefoonWerk").item("nummer").FirstChild.Value;
                }
            }
        } 

        $nodePath = "dienstverbanden/dienstverband" 
        $medewerker_dienstverbandNodeList = $medewerkerNode.SelectNodes($nodePath)
        foreach ($medewerker_dienstverbandNode in $medewerker_dienstverbandNodeList )
        {
            $nodePath = "dienstverband[id=`'" + $medewerker_dienstverbandNode.FirstChild.Value + "`']"
            $dienstverbandNode = $dienstverbandenNode.SelectSingleNode($nodePath)
            $dienstverband = @{id =  $dienstverbandNode.FirstChild.Value}
            if ($null -ne $dienstverbandNode) {
                $dienstverband = @{
                    id = $dienstverbandNode.Item("id").FirstChild.Value
                    afkorting = $dienstverbandNode.Item("afkorting").FirstChild.Value
                    naam = $dienstverbandNode.Item("afkorting").FirstChild.Value
                }
            }

            $contract  = @{
                ContractType = "dienstverband"
                dienstverband = $dienstverband;
                groep = @{} #dummy voor mapping
                inschrijvingType = @{} #dummy voor de mapping
            }
            $contracts += $contract

        }

        $medewerkerObject = @{
            PersonType = "medewerker"
            Brin = $Brin                    
            ExternalId = [string] ($Brin + "_" + $medewerkerNode.item("id").FirstChild.Value)
            DisplayName = $medewerkerNode.item("roepNaam").FirstChild.Value +  $medewerkerNode.item("achternaam").FirstChild.Value;
         
            aanhef = $medewerkerNode.item("aanhef").FirstChild.Value;
            achternaam = $medewerkerNode.item("lastName").FirstChild.Value;
            adres = $adres 
            Contracts = $contracts            
            email = $medewerkerNode.item("email").FirstChild.Value;
            mobile = $mobile
            roepnaam = $medewerkerNode.item("roepNaam").FirstChild.Value;
            voornamen = $medewerkerNode.item("firstName").FirstChild.Value;   #there is no voornamen in medewerker
            telefoonWerk = $telefoonWerk
            voornaam = $medewerkerNode.item("firstName").FirstChild.Value;
            geboortedatum = $medewerkerNode.item("geboortedatum").FirstChild.Value;
            geboortedatumOnzeker = $medewerkerNode.item("geboortedatumOnzeker").FirstChild.Value;
            geboorteplaats = $medewerkerNode.item("geboorteplaats").FirstChild.Value;
            geboorteland = $medewerkerNode.item("geboorteland").FirstChild.Value;
        }
       $dummyIndex = $medewerkersObject.add($medewerkerObject);
    }
    return $medewerkersObject
}

# start of main loop of script execution
$personList = [System.Collections.ArrayList]::new()
foreach ($Brin in $Connectorsettings.brinIdentifiers)
{
    Write-Verbose -Verbose "ParnasSys import medewerkers looping Brins ($Brin)";

    if ((Get-Date).Month -lt 8) {
        $schoolYear = (Get-Date).AddYears(-1).ToString("yyyy") + " / " + (Get-Date).ToString("yyyy")
    } else {
        $schoolYear = (Get-Date).ToString("yyyy") + " / " + (Get-Date).AddYears(1).ToString("yyyy")
    }
    $medewerkersReturnNode = Get-ParnasSysMedewerkers `
                            -Brin $Brin `
                            -WebServiceUri  $Connectorsettings.webServiceUri `
                            -SupplierName   $Connectorsettings.supplierName `
                            -SupplierKey    $Connectorsettings.supplierKey `
                            -schoolYear     $schoolYear `
                            -Proxy $Connectorsettings.Proxy 

    $medewerkers = Convert_Returnxml_to_Medewerkerlist -returnNode $medewerkersReturnNode;
    $dummyIndex = $personList.AddRange($medewerkers);

}

Write-Verbose -Verbose "ParnasSys import medewerkers succesfull";
$i=0
foreach ($person in $personList)
{
    Write-Output $person | ConvertTo-json -Depth 10
    $i = $i + 1;
    if ($i -gt 20000 )
    {break;}
}
