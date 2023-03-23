# Define the file path for the list of host names
$hostsFile = Join-Path $PSScriptRoot "hosts.txt"

# Define the output file path
$outputFile = Join-Path $PSScriptRoot "output.csv"

# Define the VLAN to floor mappings
$vlanToFloors = @{
    "164" = "Basement"
    "165" = "Basement"
    "166" = "First"
    "167" = "First"
    "168" = "First"
    "169" = "First"
    "170" = "Second"
    "171" = "Second"
    "173" = "Second"
    "174" = "Second"
    "175" = "Third"
    "177" = "Third"
    "178" = "Third"
    "179" = "Third"
    "180" = "Fourth"
    "181" = "Fourth"
    "182" = "Fifth"
    "183" = "Fifth"
}

# Define a function to get the VLAN for a given IP address
function Get-VlanForIpAddress($ipAddress) {
    $pingResult = Test-Connection -Count 1 -Quiet $ipAddress
    if (!$pingResult) {
        Write-Warning "Could not ping $ipAddress"
        return ""
    }

    $arpResult = arp -a $ipAddress | Select-Object -First 1
    if (!$arpResult) {
        Write-Warning "Could not find MAC address for $ipAddress"
        return ""
    }

    $vlanHex = $arpResult.Split()[3]
    $vlanDecimal = [Convert]::ToInt32($vlanHex, 16).ToString()

    if (!$vlanToFloors.ContainsKey($vlanDecimal)) {
        Write-Warning "Unknown VLAN $vlanDecimal for $ipAddress"
        return ""
    }

    return $vlanToFloors[$vlanDecimal]
}

# Read the list of host names from the file
$hostNames = Get-Content $hostsFile | Where-Object { $_ -match '\S' }

# Initialize a counter for the progress bar
$counter = 0

# Loop through the host names and get the IP address and VLAN for each one
$results = foreach ($hostName in $hostNames) {
    $counter++
    Write-Progress -Activity "Getting information for $hostName" -Status "Processing $counter of $($hostNames.Count)" -PercentComplete (($counter / $hostNames.Count) * 100)

    $ipAddress = (Resolve-DnsName -Name $hostName -ErrorAction SilentlyContinue).IPAddress
    if (!$ipAddress) {
        Write-Warning "Could not resolve IP address for $hostName"
        continue
    }

    $vlan = Get-VlanForIpAddress $ipAddress
    if (!$vlan) {
        continue
    }

    [PSCustomObject]@{
        HostName = $hostName
        IpAddress = $ipAddress
        Vlan = $vlan
        Floor = ($vlanToFloors.GetEnumerator() | Where-Object { $_.Value -eq $vlan }).Name
    }
}

# Write the results to the output file
$results | Export-Csv -Path $outputFile -NoTypeInformation
