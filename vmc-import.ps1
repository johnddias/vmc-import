param(
    [Parameter(Mandatory=$true)][string]$server,
    [Parameter(Mandatory=$true)][string]$file,
    [bool]$ignoreSSL,
    [string]$username,
    [string]$password,
    [string]$group = 'VMC Migration Candidates'
)

function trustAllCerts ()
{
if (!("trustallcertspolicy" -as [type])) {
    ### Ignore TLS/SSL errors    
    add-type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}
}

function fetchResParents ($resID)
{
    $result = @{}
    $uri = 'https://'+$server+'/suite-api/api/resources/' + $resID + '/properties'
    $resp = Invoke-RestMethod -Uri $uri -Credential $credentials -DisableKeepAlive
    for ($i=0; $i -le $resp.'resource-property'.property.Length; $i++)
    {
        if ($resp.'resource-property'.property[$i].name -like 'summary|parent*')
        {
           $result.Add($resp.'resource-property'.property[$i].name,$resp.'resource-property'.property[$i].'#text')
        } 
    }
$result
}


if (($username -and $password)) {
$secure_password = ConvertTo-SecureString -String $password -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential($username, $secure_password)
}
else { $credentials = Get-Credential }

if ($ignoreSSL) {trustAllCerts}


# Fetching vRealize Operations List of Virtual Machines
$uri = "https://"+$server+"/suite-api/api/resources?resourceKind=virtualmachine"
#TODO - Error handling for REST call
$res = Invoke-RestMethod -Uri $uri -Credential $credentials -DisableKeepAlive
$vmsVROps = @{}
foreach ($resource in $res.resources.resource) 
{
   $vmsVROps.Add($resource.identifier, $resource.resourceKey.name)
}

#Fetching VMC Assessment List of Virtual Machines
$objExcel = New-Object -ComObject Excel.Application
#TODO file path input from user
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item('VMC_MIGRATION')
$numRows = ($sheet.UsedRange.Rows).count
$vmsVMC = @()
for ($i=2; $i -le $numRows; $i++) 
{
  $vmsVMC += $sheet.Cells.Item($i,1).text
}

# Searching for vRealize Operations resource ID of VMs to be added to the custom group
# Also checking for VMs that may have duplicate names and cross-referencing with Datacenter and Cluster name
$vmResIDs = @{}
for ($i=2; $i -le $numRows; $i++)
{
  $vm = $vmsVROps.GetEnumerator() | Where-Object {$_.value -eq $sheet.Cells.Item($i,1).text}
  if ($vm.Count -eq 0) 
  {
    Write-Host 'WARNING: No matching VM found for ' $sheet.Cells.Item($i,1).text
  }
  elseif ($vm.Count -gt 1)
  {
    Write-Host "Resolving VM duplicate name" $vm[0].value "with" $vm.Count "duplicates" 
# Getting a list of all VMs with matching names for further validation
    $uri = "https://"+$server+"/suite-api/api/resources?name="+$vm[0].value
# TODO - add error handling for REST call
    $matchingResources = Invoke-RestMethod -Uri $uri -Credential $credentials -DisableKeepAlive
    for ($k=0; $k -le $matchingResources.resources.resource.Length; $k++)
    {
        $parents = fetchResParents $matchingResources.resources.resource[$k].identifier
        if (($parents.'summary|parentDatacenter'-eq $sheet.Cells.Item($i,2).text) -and ($parents.'summary|parentCluster' -eq $sheet.Cells.Item($i,3).text))
        {
            $vmResIDs.Add($matchingResources.resources.resource[$k].identifier,$matchingResources.resources.resource[$k].resourceKey.name)
            break
        }
    }
  }
  else
  {
    Write-Host 'VM ' $sheet.Cells.Item($i,1).text ' is resource ID ' $vm.key
    $vmResIDs.add($vm.key,$vm.value)
  }
}

#Creating Group Type


#Creating Group
$payload = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ops:group xmlns:ops="http://webservice.vmware.com/vRealizeOpsMgr/1.0/" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<ops:resourceKey>
<ops:name>$($group)</ops:name>
<ops:adapterKindKey>Container</ops:adapterKindKey>
<ops:resourceKindKey>Environment</ops:resourceKindKey>
<ops:resourceIdentifiers/>
</ops:resourceKey>
<ops:membershipDefinition>
<ops:includedResources>$($vmResIDs.keys)</ops:includedResources>
</ops:membershipDefinition>
</ops:group>
"@

$headers = @{"X-vRealizeOps-API-use-unsupported" = "true"}
$headers.add("Content-Type", "application/xml")

$uri = "https://"+$server+"/suite-api/internal/resources/groups"
$response = Invoke-RestMethod -Method Post -Uri $uri -Credential $credentials -DisableKeepAlive -Body $payload -Headers $headers

$workbook.close()
$objExcel.Quit()
