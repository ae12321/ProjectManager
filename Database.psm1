class Database {
    hidden $conn
    hidden $tx = $null

    Database([string]$databasePath) {
        $dllPath = Join-Path (Split-Path -Parent $PSCommandPath) 'System.Data.SQLite.dll'
        Add-Type -Path $dllPath
        $c = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$databasePath;Version=3;")
        try { $c.Open() }
        catch {
            $c.Dispose()
            throw [InvalidOperationException]::new("データベースを開けませんでした: $databasePath", $_.Exception)
        }
        $this.conn = $c
    }

    [void] Dispose() {
        if ($null -ne $this.tx) {
            try { $this.tx.Rollback() } catch {}
            try { $this.tx.Dispose() } catch {}
            $this.tx = $null
        }
        if ($null -ne $this.conn) {
            try { $this.conn.Close() } catch {}
            try { $this.conn.Dispose() } catch {}
            $this.conn = $null
        }
    }

    [void] BeginTransaction() {
        if ($null -ne $this.tx) { throw [InvalidOperationException]::new("既にトランザクションが開始されています。")}
        if ($null -eq $this.conn -or $this.conn.State -ne [System.Data.ConnectionState]::Open) { throw [InvalidOperationException]::new("データベース接続がオープン状態ではありません。") }
        $this.tx = $this.conn.BeginTransaction()
    }

    [void] Commit() {
        if ($null -eq $this.tx) { throw [InvalidOperationException]::new("コミットできるトランザクションがありません。") }
        try { $this.tx.Commit() }
        finally { $this.tx.Dispose(); $this.tx = $null }
    }

    [void] Rollback() {
        if ($null -eq $this.tx) { return }
        try { $this.tx.Rollback() }
        finally { $this.tx.Dispose(); $this.tx = $null }
    }

    hidden [void] BindParams($cmd, [hashtable]$params) {
        foreach ($key in $params.Keys) {
            $p = $cmd.CreateParameter()
            $p.ParameterName = $key
            $p.Value = $params[$key]
            $cmd.Parameters.Add($p) | Out-Null
        }
    }

    [void] executeNonQuery([string]$sql) { $this.executeNonQuery($sql, @{}) }
    [void] executeNonQuery([string]$sql, [hashtable]$params) {
        $cmd = $this.conn.CreateCommand()
        try {
            $cmd.CommandText = $sql
            if ($null -ne $this.tx) { $cmd.Transaction = $this.tx }
            $this.BindParams($cmd, $params)
            $cmd.ExecuteNonQuery() | Out-Null
        } finally {
            $cmd.Dispose()
        }
    }

    [object[]] executeQuery([string]$sql) { return $this.executeQuery($sql, @{}) }
    [object[]] executeQuery([string]$sql, [hashtable]$params) {
        $cmd = $this.conn.CreateCommand()
        try {
            $cmd.CommandText = $sql
            if ($null -ne $this.tx) { $cmd.Transaction = $this.tx }
            $this.BindParams($cmd, $params)

            $reader = $cmd.ExecuteReader()
            # ドライバ差異や将来の差し替えに備えてガード
            if ($null -eq $reader) { throw [InvalidOperationException]::new("ExecuteReader が null を返しました。SQL: $sql") }
            try {
                $results = [System.Collections.Generic.List[object]]::new()
                while ($reader.Read()) {
                    $row = @{}
                    for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                        $row[$reader.GetName($i)] = $reader.GetValue($i)
                    }
                    $results.Add($row)
                }
                return $results.ToArray()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $cmd.Dispose()
        }
    }
}
