# Verifica conexão em porta específica
# Testado em Windows 7 e 10


$client = New-Object System.Net.Sockets.TcpClient
$async = $client.BeginConnect($ip, $port, $null, $null)
$wait = $async.AsyncWaitHandle.WaitOne($timeout, $false)

 if(!$wait -or !$client.Connected){
      Write-Host "Failed to connect to $ip on port $port."
      $result = "Testing connection to $ip on port $port...`r`n"
      $result += "**Failed** to connect."
} else {
      Write-Host "Connection to $ip on port $port successful."
      $result = "Testing connection to $ip on port $port...`r`n"
      $result += "**Successful** connection."
}
$client.Close()
