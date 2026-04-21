# Script tạo máy chủ web ảo (Local Server) siêu nhẹ để Chrome lưu quyền Microphone
$port = 8000
$url = "http://localhost:$port/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($url)
$listener.Start()

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "May chu dang chay tai: $url" -ForegroundColor Green
Write-Host "Ban chi can bam [Cho Phep] Microphone (Allow) 1 lan duy nhat tu bay gio!"  -ForegroundColor Yellow
Write-Host "De tat may chu nay, hay an phim Ctrl + C" -ForegroundColor Red
Write-Host "==========================================" -ForegroundColor Cyan

# Tu dong mo trinh duyet
Start-Process $url

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $requestUrl = $context.Request.Url.LocalPath
        if ($requestUrl -eq "/") { $requestUrl = "/index.html" }
        
        # Chuyen doi duong dan url thanh duong dan file tren may
        $requestUrl = $requestUrl.TrimStart("/")
        $filePath = Join-Path $PWD $requestUrl

        if (Test-Path $filePath -PathType Leaf) {
            $response = $context.Response
            $content = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentLength64 = $content.Length
            
            if ($filePath -match "\.html$") { $response.ContentType = "text/html; charset=utf-8" }
            elseif ($filePath -match "\.js$") { $response.ContentType = "application/javascript; charset=utf-8" }
            elseif ($filePath -match "\.css$") { $response.ContentType = "text/css; charset=utf-8" }
            
            $response.OutputStream.Write($content, 0, $content.Length)
            $response.Close()
        } else {
            $context.Response.StatusCode = 404
            $context.Response.Close()
        }
    }
} catch {
    # Bo qua loi khi stop
} finally {
    $listener.Stop()
}
