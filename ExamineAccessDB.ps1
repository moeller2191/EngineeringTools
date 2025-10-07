# Simple Access Database Examination Script
# First, let's just see what's in the database

Write-Host "=== EXAMINING ACCESS DATABASE ===" -ForegroundColor Green

$accessDbPath = "C:\Scripts\EngineeringTools\JobNoBurnt.accdb"

try {
    # Create ADODB connection
    $accessConnection = New-Object -ComObject ADODB.Connection
    $accessConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$accessDbPath;")
    
    Write-Host "✓ Connected to Access database" -ForegroundColor Green
    
    # Get tables
    $schema = $accessConnection.OpenSchema(20)
    $tables = @()
    
    while (-not $schema.EOF) {
        $tableName = $schema.Fields("TABLE_NAME").Value
        $tableType = $schema.Fields("TABLE_TYPE").Value
        
        if ($tableType -eq "TABLE" -and -not $tableName.StartsWith("MSys")) {
            $tables += $tableName
        }
        $schema.MoveNext()
    }
    $schema.Close()
    
    Write-Host "`nFound $($tables.Count) tables:" -ForegroundColor Yellow
    
    foreach ($tableName in $tables) {
        Write-Host "`n--- TABLE: $tableName ---" -ForegroundColor Cyan
        
        # Get record count
        try {
            $countRecordset = $accessConnection.Execute("SELECT COUNT(*) FROM [$tableName]")
            $recordCount = $countRecordset.Fields.Item(0).Value
            $countRecordset.Close()
            Write-Host "Records: $recordCount"
        } catch {
            Write-Host "Records: Unable to count"
        }
        
        # Get field information
        try {
            $fieldRecordset = $accessConnection.Execute("SELECT TOP 1 * FROM [$tableName]")
            Write-Host "Fields:"
            for ($i = 0; $i -lt $fieldRecordset.Fields.Count; $i++) {
                $fieldName = $fieldRecordset.Fields.Item($i).Name
                $fieldType = $fieldRecordset.Fields.Item($i).Type
                Write-Host "  - $fieldName (Type: $fieldType)"
            }
            $fieldRecordset.Close()
        } catch {
            Write-Host "Unable to examine fields"
        }
    }
    
    $accessConnection.Close()
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}