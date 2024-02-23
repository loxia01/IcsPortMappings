using namespace System.Security.Principal
using namespace System.Management.Automation

function Enable-IcsPortMapping
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position=0)]
        [string]$ConnectionName,
        
        [Parameter(Position=1)]
        [string[]]$Name,
        
        [switch]$PassThru
    )

    begin
    {
        if (-not ([WindowsPrincipal][WindowsIdentity]::GetCurrent()).IsInRole([WindowsBuiltInRole]'Administrator'))
        {
            $exception = "This function requires administrator rights."
            $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'AdminPrivilegeRequired', 18, $null))
        }

        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare

        $connections = @($netShare.EnumEveryConnection)
        $connectionsProps = $connections | ForEach-Object { $netShare.NetConnectionProps.Invoke($_) } | Where-Object Status -NE $null
            

        if ($connectionsProps.Name -notcontains $ConnectionName)
        {
            $exception = New-Object PSArgumentException "Cannot find a network connection with name '$($_.Value)'."
            $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'ConnectionNotFound', 13, $null))
        }
        else { $ConnectionName = $connectionsProps | Where-Object Name -EQ $ConnectionName | Select-Object -ExpandProperty Name }
            
        $connection = $connections | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $ConnectionName}
        $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)
        $portMappings = @($connectionConfig.EnumPortMappings(0))
            
        if ($Name)
        {
            $pmNames = foreach ($pmName in $Name)
            {
                if (-not ($portMappings.Properties.Name -eq $pmName))
                {
                    $exception = New-Object PSArgumentException "Cannot find a port mapping with name '${pmName}'."
                    $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'PortMappingNotFound', 13, $null))
                }
                else { $portMappings.Properties | Where-Object Name -EQ $pmName | Select-Object -ExpandProperty Name }
            }
        }
    }
    process
    {
        foreach ($pmName in $pmNames)
        {
            $pm = $portMappings | Where-Object {$_.Properties.Name -eq $pmName}
            $pm.Enable()
        }

        $portMappings = @($connectionConfig.EnumPortMappings(0))
        $output = $pmNames | ForEach-Object {
            [pscustomobject]@{
                PortMappingName = $_
                        Enabled = $portMappings.Properties | Where-Object Name -EQ $_ | Select-Object -ExpandProperty Enabled
                TargetIPAddress = $portMappings.Properties | Where-Object Name -EQ $_ | Select-Object -ExpandProperty TargetIPAddress
            }
        }
    }
    end
    {
        if ($PassThru)
        {
            $output | Sort-Object PortMappingName
        }
    }
}
