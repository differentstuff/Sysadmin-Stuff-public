# Source: https://blog.devolutions.net/2025/03/using-rdp-without-leaving-traces-the-mstsc-public-mode/

cmdkey /list | ? { $_ -Match "TERMSRV/" } | % { $_ -Replace ".*: " } | % { cmdkey /delete:$_ }
Remove-Item -Path "$Env:LocalAppData\Microsoft\Terminal Server Client\Cache" -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path "$Env:LocalAppData\Microsoft\Terminal Server Client\Cache" -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Terminal Server Client\Default" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Terminal Server Client\Servers" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Terminal Server Client\LocalDevices" -Recurse -Force -ErrorAction SilentlyContinue