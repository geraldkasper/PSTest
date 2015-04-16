
#lower executionpolicy so you can run scripts locally
Set-ExecutionPolicy remotesigned

#rename computer and restart
Rename-Computer -NewName “Ex2013LAB-AIO” -Restart

#check for InterFace number (in my case there is only 1 nic assigned. The rest of the configuration uses the interfaceindex number gathered here)
get-netadapter -physical

#rename netadapter to something you can remember
get-netadapter -interfaceindex 13 | rename-netadapter -newname “Public”

#bind new ipv4 address to interface (you must understand cidr notation, prefixlength of 24 means 24 bits are masked for the network 255.255.255.0)
New-NetIPAddress -IPAddress 192.168.1.10 -defaultgateway 192.168.1.1 -prefixlength 24 -interfaceindex 13

#set dns server
Set-DnsClientServerAddress -InterfaceIndex 13 -ServerAddresses 192.168.1.3

#configure dns client settings
Set-DNSClient -InterfaceIndex 13 -ConnectionSpecificSuffix “lab.local” -RegisterThisConnectionsAddress $true -UseSuffixWhenRegistering $true

#disable LMHOST (system wide setting)
Invoke-CimMethod -ClassName Win32_NetworkAdapterConfiguration -MethodName EnableWINS -Arguments @{DNSEnabledForWINSResolution = $false; WINSEnableLMHostsLookup = $false}

#disable netbios over TCP/IP for specific adapter
Get-CimInstance win32_networkadapterconfiguration -Filter 'servicename = "netsvc"' | Invoke-CimMethod -MethodName settcpipnetbios -Arguments @{TcpipNetbiosOptions = 2}

#rename filesystem volume
set-volume -driveletter c -newfilesystemlabel System

#assing driveletter Z to dvd drive
Get-CimInstance Win32_Volume -Filter 'drivetype = 5' | Set-CimInstance -Arguments @{driveletter = "Z:”}

#install MBX/CAS features
install-windowsfeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation -IncludeManagementTools -Restart

#enable RDP
get-CimInstance “Win32_TerminalServiceSetting” -Namespace root\cimv2\terminalservices | Invoke-CimMethod -MethodName setallowtsconnections -Arguments @{AllowTSConnections = 1; ModifyFirewallException = 1}

#set RDP to only accept NLA
get-CimInstance “Win32_TSGeneralSetting” -Namespace root\cimv2\terminalservices -Filter ‘TerminalName = “RDP-Tcp”‘ | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{UserAuthenticationRequired = 1}