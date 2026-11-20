param(
    [ValidateRange(1024, 65535)]
    [int]$Port = 3000
)

$ErrorActionPreference = 'Stop'
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
$htmlPath = Join-Path -Path $PSScriptRoot -ChildPath 'index.html'
$corePath = Join-Path -Path $PSScriptRoot -ChildPath 'dashboard-core.js'
$utf8 = [System.Text.Encoding]::UTF8

function Send-HttpResponse {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.Sockets.NetworkStream]$Stream,
        [Parameter(Mandatory = $true)]
        [int]$StatusCode,
        [Parameter(Mandatory = $true)]
        [string]$Reason,
        [Parameter(Mandatory = $true)]
        [byte[]]$Body,
        [Parameter(Mandatory = $true)]
        [string]$ContentType,
        [string]$ExtraHeaders = ''
    )

    $headerText = "HTTP/1.1 $StatusCode $Reason`r`n" +
        "Content-Type: $ContentType`r`n" +
        "Content-Length: $($Body.Length)`r`n" +
        "Cache-Control: no-store`r`n" +
        "Connection: close`r`n" +
        $ExtraHeaders +
        "`r`n"
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($headerText)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    if ($Body.Length -gt 0) {
        $Stream.Write($Body, 0, $Body.Length)
    }
    $Stream.Flush()
}

function New-TextBody {
    param([string]$Text)
    return $utf8.GetBytes($Text)
}

try {
    if (-not (Test-Path -LiteralPath $htmlPath -PathType Leaf)) {
        throw "No se encontro el archivo de la aplicacion: $htmlPath"
    }

    $listener.Start()
    Write-Host "Dashboard disponible en http://localhost:$Port/"
    Write-Host 'Servidor ligado solo a loopback. Presiona Ctrl+C para detenerlo.'

    while ($true) {
        $client = $null
        $stream = $null
        $reader = $null

        try {
            $client = $listener.AcceptTcpClient()
            $client.ReceiveTimeout = 10000
            $client.SendTimeout = 10000
            $stream = $client.GetStream()
            $stream.ReadTimeout = 10000
            $stream.WriteTimeout = 10000
            $reader = [System.IO.StreamReader]::new(
                $stream,
                [System.Text.Encoding]::ASCII,
                $false,
                1024,
                $true
            )

            $requestLine = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($requestLine)) {
                $body = New-TextBody 'Solicitud HTTP no valida.'
                Send-HttpResponse -Stream $stream -StatusCode 400 -Reason 'Bad Request' `
                    -Body $body -ContentType 'text/plain; charset=utf-8'
                continue
            }

            $headerSize = [System.Text.Encoding]::ASCII.GetByteCount($requestLine) + 2
            while ($true) {
                $headerLine = $reader.ReadLine()
                if ($null -eq $headerLine) {
                    break
                }
                $headerSize += [System.Text.Encoding]::ASCII.GetByteCount($headerLine) + 2
                if ($headerSize -gt 16384) {
                    throw 'Los encabezados HTTP exceden el limite permitido.'
                }
                if ($headerLine.Length -eq 0) {
                    break
                }
            }

            $requestParts = $requestLine -split '\s+'
            if ($requestParts.Length -lt 2) {
                $body = New-TextBody 'Solicitud HTTP no valida.'
                Send-HttpResponse -Stream $stream -StatusCode 400 -Reason 'Bad Request' `
                    -Body $body -ContentType 'text/plain; charset=utf-8'
                continue
            }

            $method = $requestParts[0].ToUpperInvariant()
            $requestTarget = $requestParts[1]
            $rawPath = ($requestTarget -split '\?', 2)[0]
            $path = [System.Uri]::UnescapeDataString($rawPath)

            if ($method -ne 'GET') {
                $body = New-TextBody 'Metodo no permitido.'
                Send-HttpResponse -Stream $stream -StatusCode 405 -Reason 'Method Not Allowed' `
                    -Body $body -ContentType 'text/plain; charset=utf-8' `
                    -ExtraHeaders "Allow: GET`r`n"
                Write-Host "$method $path -> 405"
                continue
            }

            $filePath = $null
            $contentType = $null
            if ($path -eq '/' -or $path -eq '/index.html') {
                $filePath = $htmlPath
                $contentType = 'text/html; charset=utf-8'
            }
            elseif ($path -eq '/dashboard-core.js') {
                $filePath = $corePath
                $contentType = 'application/javascript; charset=utf-8'
            }
            else {
                $body = New-TextBody 'No encontrado.'
                Send-HttpResponse -Stream $stream -StatusCode 404 -Reason 'Not Found' `
                    -Body $body -ContentType 'text/plain; charset=utf-8'
                Write-Host "$method $path -> 404"
                continue
            }

            if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                $body = New-TextBody 'Recurso local no encontrado.'
                Send-HttpResponse -Stream $stream -StatusCode 404 -Reason 'Not Found' `
                    -Body $body -ContentType 'text/plain; charset=utf-8'
                Write-Host "$method $path -> 404"
                continue
            }

            $body = [System.IO.File]::ReadAllBytes($filePath)
            Send-HttpResponse -Stream $stream -StatusCode 200 -Reason 'OK' `
                -Body $body -ContentType $contentType
            Write-Host "$method $path -> 200"
        }
        catch {
            Write-Warning "Error atendiendo una solicitud loopback: $($_.Exception.Message)"
            if ($null -ne $stream -and $stream.CanWrite) {
                try {
                    $body = New-TextBody 'Error interno del servidor local.'
                    Send-HttpResponse -Stream $stream -StatusCode 500 -Reason 'Internal Server Error' `
                        -Body $body -ContentType 'text/plain; charset=utf-8'
                }
                catch {
                    Write-Warning "No se pudo responder al cliente: $($_.Exception.Message)"
                }
            }
        }
        finally {
            if ($null -ne $reader) {
                $reader.Dispose()
            }
            if ($null -ne $stream) {
                $stream.Dispose()
            }
            if ($null -ne $client) {
                $client.Dispose()
            }
        }
    }
}
finally {
    if ($null -ne $listener) {
        $listener.Stop()
    }
}
