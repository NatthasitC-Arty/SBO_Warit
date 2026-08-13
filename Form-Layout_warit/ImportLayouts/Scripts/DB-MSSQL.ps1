# ============================================================
# DB Plugin: Microsoft SQL Server
# Dot-sourced by main scripts via: . "$PSScriptRoot\DB-MSSQL.ps1"
# ============================================================

$DB_PARAM    = "@"              # parameter prefix
$DB_NOW      = "GETDATE()"      # current timestamp
$DB_ISNUM    = "ISNUMERIC"      # numeric check

function New-DBConnection {
    param(
        [string]$Server,
        [string]$Database,
        [string]$User,
        [string]$Password,
        [int]$Timeout = 10
    )
    $cs = "Server=$Server;Database=$Database;User ID=$User;Password=$Password;Connection Timeout=$Timeout;"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    return $conn
}

function Add-BlobParam {
    param($Command, [string]$Name, [byte[]]$Bytes)
    $p = $Command.Parameters.Add($Name, [System.Data.SqlDbType]::Image)
    $p.Value = $Bytes
}

function Add-DBParam {
    param($Command, [string]$Name, $Value)
    [void]$Command.Parameters.AddWithValue($Name, $Value)
}

function Convert-DBSql { param([string]$Sql) return $Sql }
