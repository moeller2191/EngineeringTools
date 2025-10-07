# Simple database test script
$dbPath = "C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"

# Load the SQLite assembly
Add-Type -Path "C:\Scripts\EngineeringTools\packages\System.Data.SQLite.Core.1.0.116\lib\net46\System.Data.SQLite.dll" -ErrorAction SilentlyContinue

try {
    $connection = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$dbPath")
    $connection.Open()
    
    Write-Host "Connected to database successfully"
    
    # Test the query for IK3NC-0000
    $query = @"
SELECT m.*, 
       COALESCE(x.XmlStatus, 'No XML') as XmlStatus,
       COALESCE(x.HighestRelease, 0) as HighestRelease,
       COALESCE(x.ComponentCount, 0) as ComponentCount
FROM MrpPriorityList m
LEFT JOIN XMLIndex x ON m.PartNumber = x.PartNumber AND m.Revision = x.Revision
WHERE m.JobNumber LIKE '%IK3NC-0000%'
"@
    
    $command = New-Object System.Data.SQLite.SQLiteCommand($query, $connection)
    $reader = $command.ExecuteReader()
    
    $found = $false
    while ($reader.Read()) {
        $found = $true
        Write-Host "Found job: $($reader['JobNumber'])"
        Write-Host "Part Number: $($reader['PartNumber'])"
        Write-Host "Revision: $($reader['Revision'])"
        Write-Host "Description: $($reader['Description'])"
        Write-Host "XML Status: $($reader['XmlStatus'])"
        Write-Host "---"
    }
    
    if (-not $found) {
        Write-Host "No jobs found for IK3NC-0000"
        
        # Check total count
        $countQuery = "SELECT COUNT(*) FROM MrpPriorityList"
        $countCommand = New-Object System.Data.SQLite.SQLiteCommand($countQuery, $connection)
        $count = $countCommand.ExecuteScalar()
        Write-Host "Total records in MrpPriorityList: $count"
    }
    
    $reader.Close()
    $connection.Close()
    
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    Write-Host "Unable to test database directly. The application fixes should still work."
}