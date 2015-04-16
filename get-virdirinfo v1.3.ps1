#========================================================================================================
# Created on:   12/16/2012 09:30 PM
# Last Update:	02/08/2014 10:37 AM
# Created by:   Michael Van Horenbeeck
# Website:      http://michaelvh.wordpress.com
# Filename:     get-virdirinfo.ps1
# Version:      1.2
# Credits:      Thomas Torggler for some great ideas and remarks (@Torggler)
#
# Version History:  1.3     Added -Filter parameter, improved error handling in general, added warnings
#                	1.2     Added -ADPropertiesOnly parameter to speed up the script
#                   1.1		Fixed count error when only a single server exists
#					1.0		Initial Version
#========================================================================================================

<#
.Synopsis
   This script will create an HTML-report which will gather the URL-information from different virtual directories over different Exchange Servers (currently only Exchange 2010/Exchange 2013)
.DESCRIPTION
   This script will create an HTML-report which will gather the URL-information from different virtual directories over different Exchange Servers (currently only Exchange 2010/Exchange 2013)
.EXAMPLE
   . .\get-virdirinfo.ps1
   Get-VirDirInfo -filepath c:\reports

   This command will create the report in the following directory: C:\Reports
#>

function Get-VirDirInfo
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        #Specify the report file path
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
                   [ValidateNotNullOrEmpty()]
                   $filepath,
        
        #query AD instead of the IIS metabase
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
                   [Alias("ADPropertiesOnly")]
                   [Switch]
                   $ADProperties,
        
        #specify the computername to connect to. Defaults to the local host.
        [Parameter(Mandatory=$false,
                   Position=1)]
                   [ValidateNotNull()]
                   [ValidateNotNullOrEmpty()]
                   [string]
                   $ComputerName = $env:COMPUTERNAME,

        #optional filter to only query certain CAS servers. Default filter is wildcard.
        [Parameter(Mandatory=$false,
                   Position=2)]
                   [ValidateNotNull()]
                   [ValidateNotNullOrEmpty()]
                   [string]
                   $Filter="*"

    )

    Begin
    {
        #Open a Remote PS session if none already exists.
	    if (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange") {
            try {
                $sExch = New-PSSession -ConfigurationName Microsoft.Exchange -Name ExchMgmt -ConnectionUri http://$ComputerName/PowerShell/ -Authentication Kerberos 
	            $null = Import-PSSession $sExch
            } catch {
                Write-Warning "Could not connect to Exchange."
                Break
            }
        }
    }
    Process
    {
    	try{
            $servers = @(Get-ExchangeServer | ?{$_.ServerRole -like "*ClientAccess*" -and (($_.AdminDisplayVersion -like "*15*") -or ($_.AdminDisplayVersion -like "*14*") -and ($_.Name -like $Filter))} | Select-Object Name)
        }
        catch{
            Write-Warning "An error occured: could not connect to one or more Exchange servers".
            Break
        }

        #HTML headers
        $html += "<html>"
            $html += "<head>"
                $html += "<style type='text/css'>"
		    $html += "body {font-family:verdana;font-size:10pt}"
		    $html += "H1 {font-family:verdana;font-size:12pt}"
                    $html += "table {border:1px solid #000000;font-family:verdana; font-size:10pt;cellspacing:1;cellpadding:0}"
		    $html += "tr.color {background-color:#00A2E8;color:#FFFFFF;font-weight:bold}"
                $html += "</style>"
            $html += "</head>"
        $html += "<body>"

        #Report Legend
        $html += "Get-VirDirInfo.ps1<br/>"
        $html += "<b>Report generated on: </b>"+(get-date).DateTime
        
        #Add warning that the script pulled only the ADProperties
        if($ADProperties){
            $html += "<br/><b>Warning: The script was run using the -ADPropertiesOnly switch and might not show all information</b>"
        }
        $html += "<br/><br/>"
        
        #Autodiscover
		$html += "<h1>Autodiscover</h1>"
		$html += "<table border='1'>"
		$html += "<tr class='color'>"
		$html += "<td>Server</td><td>Internal Uri</td><td>InternalURL</td><td>ExternalUrl</td><td>Auth. (Int.)</td><td>Auth. (Ext.)</td>"
		$html += "</tr>"
		$i=0
		foreach($server in $servers){
            $i++
            Write-Progress -Activity "Getting Autodiscover URL information" -Status "Progress:"-PercentComplete (($i / $servers.count)*100)
			$autodresult = Get-ClientAccessServer -Identity $server.name | Select Name,AutodiscoverServiceInternalUri
            if($ADProperties){
                $autodvirdirresult = Get-AutodiscoverVirtualDirectory -Server $server.name -ADPropertiesOnly | Select InternalUrl,ExternalUrl,InternalAuthenticationMethods,ExternalAuthenticationMethods
            }
            else{
                $autodvirdirresult = Get-AutodiscoverVirtualDirectory -Server $server.name | Select InternalUrl,ExternalUrl,InternalAuthenticationMethods,ExternalAuthenticationMethods
            }
            
			$autodhtml += "<tr>"
			$autodhtml += "<td>"+$autodresult.Name+"</td>"
            $autodhtml += "<td>"+$autodresult.AutodiscoverServiceInternalUri+"</td>"
			$autodhtml += "<td>"+$autodvirdirresult.InternalURL.absoluteURI+"</td>"
			$autodhtml += "<td>"+$autodvirdirresult.ExternalURL.absoluteURI+"</td>"
            $autodhtml += "<td>"+$autodvirdirresult.InternalAuthenticationMethods+"</td>"
			$autodhtml += "<td>"+$autodvirdirresult.ExternalAuthenticationMethods+"</td>"
			$autodhtml += "</tr>"
			
			Clear-Variable -Name autodresult,autodvirdirresult
			
		}
		$html += $autodhtml
		$html += "</table>"

        #Outlook Web App (OWA)
        $html += "<br/><br/>"
		$html += "<h1>Outlook Web App (OWA):</h1>"
		$html += "<table border='1'>"
		$html += "<tr class='color'>"
		$html += "<td>Server</td><td>Name</td><td>InternalURL</td><td>ExternalUrl</td>"
		$html += "</tr>"
		$i=0
		foreach($server in $servers){
            $i++
            Write-Progress -Activity "Getting OWA virtual directory information" -Status "Progress:"-PercentComplete (($i / $servers.count)*100)
            if($ADProperties){
                $owaresult = Get-OWAVirtualDirectory -server $server.name -AdPropertiesOnly | Select Name,Server,InternalUrl,ExternalUrl
            }
            else{
                $owaresult = Get-OWAVirtualDirectory -server $server.name | Select Name,Server,InternalUrl,ExternalUrl
            }
			
			$owahtml += "<tr>"
			$owahtml += "<td>"+$owaresult.Server+"</td>"
			$owahtml += "<td>"+$owaresult.Name+"</td>"
			$owahtml += "<td>"+$owaresult.InternalURL.absoluteURI+"</td>"
			$owahtml += "<td>"+$owaresult.ExternalURL.absoluteURI+"</td>"
			$owahtml += "</tr>"
			
			Clear-Variable -Name owaresult
			
		}
		$html += $owahtml
		$html += "</table>"

        #Exchange Control Panel (ECP)
        $html += "<br/><br/>"
        $html += "<h1>Exchange Control Panel (ECP):</h1>"
		$html += "<table border='1'>"
		$html += "<tr class='color'>"
		$html += "<td>Server</td><td>Name</td><td>InternalURL</td><td>ExternalUrl</td>"
		$html += "</tr>"
		$i=0
		foreach($server in $servers){
            $i++
            Write-Progress -Activity "Getting ECP virtual directory information" -Status "Progress:"-PercentComplete (($i / $servers.count)*100)
            if($ADProperties){
			    $ecpresult = Get-ECPVirtualDirectory -server $server.name -ADPropertiesOnly | Select Name,Server,InternalUrl,ExternalUrl
            }
            else{
                $ecpresult = Get-ECPVirtualDirectory -server $server.name | Select Name,Server,InternalUrl,ExternalUrl
            }

			$ecphtml += "<tr.color>"
			$ecphtml += "<td>"+$ecpresult.Server+"</td>"
			$ecphtml += "<td>"+$ecpresult.Name+"</td>"
			$ecphtml += "<td>"+$ecpresult.InternalURL.absoluteURI+"</td>"
			$ecphtml += "<td>"+$ecpresult.ExternalURL.absoluteURI+"</td>"
			$ecphtml += "</tr>"

            Clear-Variable -Name ecpresult
		}
		$html += $ecphtml
		$html += "</table>"
   		
        #Outlook Anywhere
        $html += "<br/><br/>"
        $html += "<h1>Outlook Anywhere:</h1>"
		$html += "<table border='1'>"
		$html += "<tr class='color'>"
		$html += "<td>Server</td><td>Internal Hostname</td><td>External Hostname</td><td>Auth.(Int.)</td><td>Auth. (Ext.)</td><td>Auth. IIS</td>"
		$html += "</tr>"
		$i=0
		foreach($server in $servers){
            $i++
            Write-Progress -Activity "Getting Outlook Anywhere Information" -Status "Progress:"-PercentComplete (($i / $servers.count)*100)
            if($ADProperties){
			    $oaresult = Get-OutlookAnywhere -server $server.name -ADPropertiesOnly | Select Name,Server,InternalHostname,ExternalHostname,ExternalClientAuthenticationMethod,InternalClientAuthenticationMethod,IISAuthenticationMethods
            }
            else{
                $oaresult = Get-OutlookAnywhere -server $server.name | Select Name,Server,InternalHostname,ExternalHostname,ExternalClientAuthenticationMethod,InternalClientAuthenticationMethod,IISAuthenticationMethods
            }

			$oahtml += "<tr.color>"
			$oahtml += "<td>"+$oaresult.Server+"</td>"
			$oahtml += "<td>"+$oaresult.InternalHostname+"</td>"
			$oahtml += "<td>"+$oaresult.ExternalHostname+"</td>"
            $oahtml += "<td>"+$oaresult.InternalClientAuthenticationMethod+"</td>"
			$oahtml += "<td>"+$oaresult.ExternalClientAuthenticationMethod+"</td>"
            $oahtml += "<td>"+$oaresult.IISAuthenticationMethods+"</td>"
			$oahtml += "</tr>"

            Clear-Variable oaresult
		}
		$html += $oahtml
		$html += "</table>"    


        #Offline Address Book (OAB)
        $html += "<br/><br/>"
        $html += "<h1>Offline Address Book (OAB):</h1>"
		$html += "<table border='1'>"
		$html += "<tr class='color'>"
		$html += "<td>Server</td><td>OABs</td><td>Internal URL</td><td>External Url</td><td>Auth.(Int.)</td><td>Auth. (Ext.)</td>"
		$html += "</tr>"
		$i=0
		foreach($server in $servers){
            $i++
            Write-Progress -Activity "Getting OAB Information" -Status "Progress:"-PercentComplete (($i / $servers.count)*100)
            if($ADProperties){
                $oabresult = Get-OABVirtualDirectory -server $server.name -ADPropertiesOnly | Select Server,InternalUrl,ExternalUrl,ExternalAuthenticationMethods,InternalAuthenticationMethods,OfflineAddressBooks
            }
            else{
                $oabresult = Get-OABVirtualDirectory -server $server.name | Select Server,InternalUrl,ExternalUrl,ExternalAuthenticationMethods,InternalAuthenticationMethods,OfflineAddressBooks
            }
			

			$oabhtml += "<tr.color>"
			$oabhtml += "<td>"+$oabresult.Server+"</td>"
            $oabhtml += "<td>"+$oabresult.OfflineAddressBooks+"</td>"
			$oabhtml += "<td>"+$oabresult.InternalURL.absoluteURI+"</td>"
			$oabhtml += "<td>"+$oabresult.ExternalURL.absoluteURI+"</td>"
            $oabhtml += "<td>"+$oabresult.InternalAuthenticationMethods+"</td>"
			$oabhtml += "<td>"+$oabresult.ExternalAuthenticationMethods+"</td>"
			$oabhtml += "</tr>"

            Clear-Variable oabresult
		}
		$html += $oabhtml
		$html += "</table>"

        #ActiveSync (EAS)
        $html += "<br/><br/>"
        $html += "<h1>ActiveSync (EAS):</h1>"
		$html += "<table border='1'>"
		$html += "<tr class='color'>"
		$html += "<td>Server</td><td>Internal URL</td><td>External Url</td><td>Auth. (Ext.)</td>"
		$html += "</tr>"
		$i=0
		foreach($server in $servers){
            $i++
            Write-Progress -Activity "Getting ActiveSync Information" -Status "Progress:"-PercentComplete (($i / $servers.count)*100)
            if($ADProperties){
                $easresult = Get-ActiveSyncVirtualDirectory -server $server.name -ADPropertiesOnly | Select Server,InternalUrl,ExternalUrl,ExternalAuthenticationMethods,InternalAuthenticationMethods
            }
            else{
                $easresult = Get-ActiveSyncVirtualDirectory -server $server.name | Select Server,InternalUrl,ExternalUrl,ExternalAuthenticationMethods,InternalAuthenticationMethods
            }
			

			$eashtml += "<tr.color>"
			$eashtml += "<td>"+$easresult.Server+"</td>"
			$eashtml += "<td>"+$easresult.InternalURL.absoluteUri+"</td>"
			$eashtml += "<td>"+$easresult.ExternalURL.absoluteUri+"</td>"
			$eashtml += "<td>"+$easresult.ExternalAuthenticationMethods+"</td>"
			$eashtml += "</tr>"

            Clear-Variable easresult
		}
		$html += $eashtml
		$html += "</table>"

		$html | Out-File $filepath"\virdirinfo_"$(get-date -Format d-MM-yyyy_HH\hmm\mss\s)".html"
    }
    End
    {
		Get-PSSession | ?{$_.ComputerName -like "$server"} | Remove-PSSession
		Clear-Variable Owahtml
		Clear-Variable Owaresult
		Clear-Variable html
    }
}