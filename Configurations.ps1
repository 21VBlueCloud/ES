<#
.DESCRIPTION
Helper functions for configurations.
#>

function Get-CloudConfigurations {
    param (
        [ValidateSet("Public", "Gallatin", "BlackForest", "TrailBlazer", "PathFinder", "FairfaxGccProd", "USNat", "USSec")]
        [string] $AzureCloud
    )

    $config = [PSCustomObject]@{
        AzureCloud = $AzureCloud
        AzureEnvironment = ""
        KeyVaultSecretFormat = ""
        ManagementVaultAadTenantId = ""
        ManagementVaultSubscriptionId = ""
        ManagementVaultName = ""
        ManagementVaultAadDomain = ""
        UseTorusSecretSafe = $false
    }
    
    switch ($AzureCloud.ToLowerInvariant()) {
        trailblazer {
            $config.AzureEnvironment = 'AzureUSGovernment'
            $config.KeyVaultSecretFormat = 'https://{0}.vault.usgovcloudapi.net:443/secrets/{1}'
            $config.ManagementVaultSubscriptionId = '5aeeb09f-4865-4037-b9fe-b58b864e7f0e'
            $config.ManagementVaultName = 'griffinazmgtkv-tb'
            $config.ManagementVaultAadTenantId = 'a972b02b-2e02-4381-9dda-f8c703e9d5b9'
            $config.ManagementVaultAadDomain = "gsgotrs04.prod.outlook.com"
            $config.UseTorusSecretSafe = $false
        }
        pathfinder {
            $config.AzureEnvironment = 'AzureUSGovernment'
            $config.KeyVaultSecretFormat = 'https://{0}.vault.usgovcloudapi.net:443/secrets/{1}'
            $config.ManagementVaultSubscriptionId = '9bf04ed9-981f-4ba8-965b-eea0d513008b'
            $config.ManagementVaultName = 'griffinazcertpathfinder'
            $config.ManagementVaultAadTenantId = 'a972b02b-2e02-4381-9dda-f8c703e9d5b9'
            $config.ManagementVaultAadDomain = "gsgotrs04.prod.outlook.com"
            $config.UseTorusSecretSafe = $false
        }
        fairfaxgccprod {
            $config.AzureEnvironment = 'AzureUSGovernment'
            $config.KeyVaultSecretFormat = 'https://{0}.vault.usgovcloudapi.net:443/secrets/{1}'
            $config.ManagementVaultSubscriptionId = '22abfc6f-f2b6-4e5b-bd9c-c3cc9dd07303'
            $config.ManagementVaultName = 'griffingccapacityintrnal'
            $config.ManagementVaultAadTenantId = 'a972b02b-2e02-4381-9dda-f8c703e9d5b9'
            $config.ManagementVaultAadDomain = "gsgotrs04.prod.outlook.com"
            $config.UseTorusSecretSafe = $false
        }
        blackforest {
            $config.AzureEnvironment = 'AzureGermanCloud'
            $config.KeyVaultSecretFormat = 'https://{0}.vault.microsoftazure.de:443/secrets/{1}'
            $config.ManagementVaultSubscriptionId = '4e4eb789-fa5b-4d5e-b18a-f8087677f588'
            $config.ManagementVaultName = 'griffinazuremgtkeyvault'
            $config.ManagementVaultAadTenantId = '47700c02-dab5-4be3-958d-f90aff42622e'
            $config.ManagementVaultAadDomain = "deme.gbl"
            $config.UseTorusSecretSafe = $false
        }
        gallatin {
            $config.AzureEnvironment = 'AzureChinaCloud'
            $config.KeyVaultSecretFormat = 'https://{0}.vault.azure.cn:443/secrets/{1}'
            $config.ManagementVaultSubscriptionId = '175341fb-af76-4bd4-8b75-e0139e308e00'
            $config.ManagementVaultName = 'griffinazuremgmtkv'
            $config.ManagementVaultAadTenantId = 'a55a4d5b-9241-49b1-b4ff-befa8db00269'
            $config.ManagementVaultAadDomain = "cme.gbl"
            $config.UseTorusSecretSafe = $true
        }
        usnat {
            $config.AzureEnvironment = 'USNat'
            $config.KeyVaultSecretFormat = 'hhttps://{0}.vault.cloudapi.eaglex.ic.gov/secrets/{1}'
            $config.UseTorusSecretSafe = $true
        }
        ussec {
            $config.AzureEnvironment = 'USSec'
            $config.KeyVaultSecretFormat = 'hhttps://{0}.vault.cloudapi.microsoft.scloud/secrets/{1}'
            $config.UseTorusSecretSafe = $true
        }
        Default {
            # Use microsoft.com tenant for testing
            $config.AzureEnvironment = 'AzureCloud'
            $config.KeyVaultSecretFormat = 'https://{0}.vault.azure.net:443/secrets/{1}'
            $config.ManagementVaultAadTenantId = '72f988bf-86f1-41af-91ab-2d7cd011db47'
            $config.ManagementVaultSubscriptionId = '805448f1-03de-4362-a4b3-b54df1ba0d27'
            $config.ManagementVaultName = 'Ev2OnboardingTest'
            $config.ManagementVaultAadDomain = "microsoft.com"
            $config.UseTorusSecretSafe = $false
        }
    }

    return $config
}

function Get-AadTenantConfigurations {
    param (
        [Parameter(Mandatory, ParameterSetName = "Id")]
        [System.Guid] $Id,

        [Parameter(Mandatory, ParameterSetName = "Name")]
        [ValidateSet("Microsoft", "prdtrs01", "deutrs02", "gsgotrs04", "chntrs07", "trs08", "trs09", "DEME", "CME", "ChinaProd")]
        [string] $FriendlyName
    )

    if ($PSBoundParameters.ContainsKey("FriendlyName")) {
        $Id = Get-AadTenantIdFromFriendlyName -FriendlyName $FriendlyName
    }

    $config = [PSCustomObject]@{
        FriendlyName = ""
        AzureEnvironment = ""
        DomainName = ""
        UpnFormat = ""
        Id = $Id
    }
    switch ($Id.ToString().ToLowerInvariant()) {
        "47700c02-dab5-4be3-958d-f90aff42622e" {
            $config.FriendlyName = "DEME"
            $config.AzureEnvironment = 'AzureGermanCloud'
            $config.DomainName = "deme.gbl"
            $config.UpnFormat = "{0}@deme.gbl"
        }
        "1f1ca554-f396-40d9-8bcf-b57544a58a00" {
            $config.FriendlyName = "deutrs02"
            $config.AzureEnvironment = 'AzureGermanCloud'
            $config.DomainName = "deutrs02.prod.outlook.com"
            $config.UpnFormat = "{0}_debug@deutrs02.prod.outlook.com"
        }
        "a972b02b-2e02-4381-9dda-f8c703e9d5b9" {
            $config.FriendlyName = "gsgotrs04"
            $config.AzureEnvironment = 'AzureUSGovernment'
            $config.DomainName = "gsgotrs04.prod.outlook.com"
            $config.UpnFormat = "{0}_debug@gsgotrs04.prod.outlook.com"
        }
        "a55a4d5b-9241-49b1-b4ff-befa8db00269" {
            $config.FriendlyName = "CME"
            $config.AzureEnvironment = 'AzureChinaCloud'
            $config.DomainName = "cme.gbl"
            $config.UpnFormat = "{0}@cme.gbl"
        }
        "d294a672-1e15-47e3-b224-84ff4f6f24d5" {
            $config.FriendlyName = "chntrs07"
            $config.AzureEnvironment = 'AzureChinaCloud'
            $config.DomainName = "chntrs07.prod.partner.outlook.cn"
            $config.UpnFormat = "{0}_debug@chntrs07.prod.partner.outlook.cn"
        }
        "135f51f9-ca0a-4ef7-b907-429b76c6a053" {
            $config.FriendlyName = "chinaprod"
            $config.AzureEnvironment = 'AzureChinaCloud'
            $config.DomainName = "msftcosi.partner.onmschina.cn"
        }
        "cdc5aeea-15c5-4db6-b079-fcadd2505dc2" {
            $config.FriendlyName = "prdtrs01"
            $config.AzureEnvironment = 'AzureCloud'
            $config.DomainName = "prdtrs01.prod.outlook.com"
            $config.UpnFormat = "{0}_debug@prdtrs01.prod.outlook.com"
        }
        "72f988bf-86f1-41af-91ab-2d7cd011db47" {
            $config.FriendlyName = "Microsoft"
            $config.AzureEnvironment = 'AzureCloud'
            $config.DomainName = "microsoft.com"
            $config.UpnFormat = "{0}@microsoft.com"
        }
        "dc12cfcb-7c57-4e3e-92b9-6ea4cbc258e9" {
            $config.FriendlyName = "trs08"
            $config.AzureEnvironment = 'USNat'
            $config.DomainName = "trs08.prod.exo.eaglex.ic.gov"
            $config.UpnFormat = "{0}_debug@trs08.prod.exo.eaglex.ic.gov"
        }
        "ec044cc1-d8de-4194-a9b6-f5abc1c30081"{
            $config.FriendlyName = "trs09"
            $config.AzureEnvironment = 'USSec'
            $config.DomainName = "TRS09.prod.exo.microsoft.scloud"
            $config.UpnFormat = "{0}_debug@TRS09.prod.exo.microsoft.scloud"
        }
        Default {
            throw "Unknown AAD tenant ID: '$Id'"
        }
    }

    return $config
}

function Get-AadTenantIdFromFriendlyName {
    param (
        [ValidateSet("Microsoft", "prdtrs01", "deutrs02", "gsgotrs04", "chntrs07", "trs08", "trs09", "DEME", "CME", "ChinaProd")]
        [string] $FriendlyName
    )

    switch ($FriendlyName.ToLowerInvariant()) {
        "microsoft" {
            return "72f988bf-86f1-41af-91ab-2d7cd011db47"
        }
        "prdtrs01" {
            return "cdc5aeea-15c5-4db6-b079-fcadd2505dc2"
        }
        "deutrs02" {
            return "1f1ca554-f396-40d9-8bcf-b57544a58a00"
        }
        "gsgotrs04" {
            return "a972b02b-2e02-4381-9dda-f8c703e9d5b9"
        }
        "chntrs07" {
            return "d294a672-1e15-47e3-b224-84ff4f6f24d5"
        }
        "trs08" {
            return "dc12cfcb-7c57-4e3e-92b9-6ea4cbc258e9"
        }
        "trs09" {
            return "ec044cc1-d8de-4194-a9b6-f5abc1c30081"
        }
        "deme" {
            return "47700c02-dab5-4be3-958d-f90aff42622e"
        }
        "cme" {
            return "a55a4d5b-9241-49b1-b4ff-befa8db00269"
        }
        "chinaprod" {
            return "135f51f9-ca0a-4ef7-b907-429b76c6a053"
        }
    }

    throw "Unknow tenant: '$FriendlyName'"
}