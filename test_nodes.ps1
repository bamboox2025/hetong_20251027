$nodes = @("香港", "日本", "新加坡", "台湾")
foreach ($node in $nodes) {
    Write-Host "正在测试 $node 节点..." -ForegroundColor Yellow
    # 假设您有节点切换 API，或手动切换后运行
    $web = New-Object System.Net.WebClient
    $web.Proxy = New-Object System.Net.WebProxy("http://127.0.0.1:7890", $true)
    try {
        $null = $web.DownloadString("https://api.github.com")
        Write-Host "$node : 可达" -ForegroundColor Green
    } catch {
        Write-Host "$node : 封锁" -ForegroundColor Red
    }
    Start-Sleep -Seconds 2
}