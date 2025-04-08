# Snippet from Yossi to test for Basic tables in workspace, even a table without ingestion since table visibility logic might hide tables.

# Login to Azure

Connect-AzAccount -UseDeviceAuthentication

# Set workspace param

    $rg = "resource-group"

    $ws = "workapce-name"

# Get the Log Analytics workspace

    $workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $rg -Name $ws

# Get the tables in the workspace

    $tables = Get-AzOperationalInsightsTable -ResourceGroupName $rg -WorkspaceName $ws

# Display the table names and their plans

  $tables | ForEach-Object {

      [PSCustomObject]@{

          TableName = $_.Name

          Plan      = $_.Plan

      }

  } | Format-Table -AutoSize