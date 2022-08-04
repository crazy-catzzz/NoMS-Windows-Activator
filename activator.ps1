[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $ProductKey
)

# Function prototypes
function HWID-GetKey {}
function KMS38-GetKey {}
function Get-BuildNumber {}

# Main program
function main {
    Push-Location -Path $PSScriptRoot

    $Build = Get-BuildNumber # We need the Build Number to determine which keys to set

    if ($ProductKey.Length -eq 0) {
        $SkuId = Get-SKU

        $ProductKey = HWID-GetKey -SkuId $SkuId -Build $Build # Try to use HWID
        if ($ProductKey.Length -eq 0) { # If HWID failed, use KMS38
            $ProductKey = KMS38-GetKey -SkuId $SkuId -Build $Build
        }

        Write-Host $ProductKey # Debug
    }
}

# Function definitions
function HWID-GetKey {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [int]
        $SkuId,
        [Parameter(Mandatory = $true, Position = 1)]
        [int]
        $Build
    ) # Set SKU ID and Build Number as parameters

    process {
        $ProductKey = $null # Clear Product Key variable

        $ProductKeys = @{ # HWID Product Keys array, each element corresponds to a certain SKU ID
            4   = 'XGVPP-NMH47-7TTHJ-W3FW7-8HV2C'
            27  = '3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT'
            48  = 'VK7JG-NPHTM-C97JM-9MPGT-3V66T'
            49  = '2B87N-8KFHP-DKV6R-Y2C8J-PKCKT'
            98  = '4CPRK-NM3K3-X6XXQ-RXX86-WXCHW'
            99  = 'N2434-X9D7W-8PF6X-8DV9T-8TYMD'
            100 = 'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT'
            101 = 'YTMG3-N6DKC-DKB77-7M9GH-8HVX7'
            121 = 'YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY'
            122 = '84NGF-MHBT6-FXBX8-QWJK7-DRR8H'
            161 = 'DXG7C-N36C4-C4HTG-X4T3X-2YV77'
            162 = 'WYPNQ-8C467-V2W6J-TX4WX-WT2RQ'
            164 = '8PTT6-RNW4C-6V7J2-C2D3X-MHBPB'
            165 = 'GJTYN-HDMQY-FRR76-HVGC7-QPF8P'
            175 = 'NJCF7-PW8QT-3324D-688JX-2YV66'
            188 = 'XQQYW-NFFMW-XJPBH-K8732-CKFFD'
            191 = 'QPM6N-7J2WJ-P88HH-P3YRH-YY74H'
            203 = 'KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W'
        }

        switch ($Build) {
            {$_ -eq 10240} {
                $ProductKeys = @{
                    125 = 'FWN7H-PF93Q-4GGP8-M8RF3-MDWWW'
                    126 = '8V8WN-3GXBH-2TCMG-XHRX3-9766K'
                }
            }
            {$_ -eq 14393} {
                $ProductKeys = @{
                    125 = 'NK96Y-D9CD8-W44CQ-R8YTK-DYJWX'
                    126 = '2DBW3-N2PJG-MVHW3-G7TDK-9HKR4'
                }
            }
            {$_ -eq 17763} {
                $ProductKeys = @{
                    125 = '43TBQ-NH92J-XKTM7-KT3KK-P39PB'
                    126 = 'M33WV-NHY3C-R7FPM-BQGPT-239PG'
                }
            }
        }

        if ($null -ne $ProductKeys[$SkuId]) { # Check if a key associated with the SKU ID exists
            $ProductKey = $ProductKeys[$SkuId] # Change the Product Key variable to the HWID Product Key
        }

        $ProductKey

    }
}

function KMS38-GetKey {

}

function Get-BuildNumber {
    process {
        [int](Get-CimInstance -Query 'SELECT BuildNumber FROM Win32_OperatingSystem').BuildNumber
    }
}

main