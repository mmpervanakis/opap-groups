

# Install the Import-Excel module
#Install-Module -Name ImportExcel -Scope CurrentUser

# Import the Excel file
$excelData = Import-Excel -Path "C:\Users\mmper\Desktop\opap groups\ipam_104_export.xlsx"

# Extract relevant columns
$relevantData = $excelData | Select-Object IP, description, hostname | Where-Object {$_.hostname -ne $null}

$relevantData = $excelData | Select-Object IP, description, @{Name='hostname'; Expression={($_.hostname -replace '-\d+$','')}} | Where-Object {$_.hostname -ne $null}
$relevantData;

# Create a concatenated list of IPs based on description

$concatenatedIPs = $relevantData | Group-Object hostname | ForEach-Object {
    [PSCustomObject]@{
        host_group = $_.Name
        ConcatenatedIPs = ($_.Group.IP -join ",")
        ConcatenatedDescriptions = $_.description
    }
}

# Output the concatenated list of IPs
$concatenatedIPs | Export-Csv -Path 'output.csv' -NoTypeInformation
