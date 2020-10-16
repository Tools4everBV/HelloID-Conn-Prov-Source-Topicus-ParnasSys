# HelloID-Conn-Prov-Source-Topicus-ParnasSys

Version 0.9


## Connection settings
The connections settings are currently hardcoded at the start of the script. Modify them inside the Persons.ps1 
webServiceUri 
    URL to the webservice. Example: https://parnassys.net or https://acceptatie.parnassys.net/bao/services/cxf/v3/generic
supplierName
    The supplier name (account) used for the connection example: "Identity &amp; Access Manager (IAM)"
supplierKey
    The associated password
brinIdentifiers   
    an array of brin numbers representing the schools/organizations from wich to collect data.  The script loops though all schools in this list 



##Specific remarks regarding this Connector and script
Currently it exist of 1 script Persons.ps1. The department script in HelloID is not used and can be left empty
It collects information about both "Leerlingen" and "medewerkers".  The "PersonType" field is added to denote the type.






 




