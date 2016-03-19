clear
$currentLocation = Split-Path -Parent $MyInvocation.MyCommand.Definition

# load required assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Web.Security")
[System.Reflection.Assembly]::LoadWithPartialName("System.Web.HttpUtility")
[System.Reflection.Assembly]::LoadFile("$currentLocation\AE.Net.Mail.dll")

Add-Type -AssemblyName System.Web;

Function Write-ToFile {
    Param($Message = "")    
    "<br/><strong>" +  [datetime]::Now.ToString() + "</strong> " + $Message | Out-File ($currentLocation + "\sms.html") -Append
}

Function Html-ToText { 
    param([System.String] $html) 

    # remove line breaks, replace with spaces 
    #$html = $html -replace "(`r|`n|`t)", " " 
    # write-verbose "removed line breaks: `n`n$html`n" 

    # remove invisible content 
    @('head', 'style', 'script', 'object', 'embed', 'applet', 'noframes', 'noscript', 'noembed') | % { 
    $html = $html -replace "<$_[^>]*?>.*?</$_>", "" 
    } 
    # write-verbose "removed invisible blocks: `n`n$html`n" 

    # Condense extra whitespace 
    $html = $html -replace "( )+", " " 
    # write-verbose "condensed whitespace: `n`n$html`n" 

    # Add line breaks 
    @('div','p','blockquote','h[1-9]') | % { $html = $html -replace "</?$_[^>]*?>.*?</$_>", ("`n" + '$0' )}  
    
    # Add line breaks for self-closing tags 
    @('div','p','blockquote','h[1-9]','br') | % { $html = $html -replace "<$_[^>]*?/>", ('$0' + "`n")}  
    # write-verbose "added line breaks: `n`n$html`n" 

    #strip tags  
    $html = $html -replace "<[^>]*?>", "" 
    # write-verbose "removed tags: `n`n$html`n" 

    # replace common entities 
    @(  
    @("&amp;bull;", " * "), 
    @("&amp;lsaquo;", "<"), 
    @("&amp;rsaquo;", ">"), 
    @("&amp;(rsquo|lsquo);", "'"), 
    @("&amp;(quot|ldquo|rdquo);", '"'), 
    @("&amp;trade;", "(tm)"), 
    @("&amp;frasl;", "/"), 
    @("&amp;(quot|#34|#034|#x22);", '"'), 
    @('&amp;(amp|#38|#038|#x26);', "&amp;"), 
    @("&amp;(lt|#60|#060|#x3c);", "<"), 
    @("&amp;(gt|#62|#062|#x3e);", ">"), 
    @('&amp;(copy|#169);', "(c)"), 
    @("&amp;(reg|#174);", "(r)"), 
    @("&amp;nbsp;", " "), 
    @("&amp;(.{2,6});", "") 
    ) | % { $html = $html -replace $_[0], $_[1] } 
    # write-verbose "replaced entities: `n`n$html`n" 

    return $html  
}

Function Send-SMS_BySMSTorrent {
    Param(
        [string]$Sender = "SENDER", 
        
        [Parameter(Position = 1, Mandatory = $true)]
        [array]$Recipients = @(), 
        
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$Body,
        
        [Parameter(Position = 3, Mandatory = $false)]
        [string]$Account="vestude",
        
        [Parameter(Position = 4, Mandatory = $false)]
        [string]$Pass = "vestude",
        
        [Parameter(Position = 5, Mandatory = $false)]
        [ScriptBlock]$OnSuccess = {},
        
        [Parameter(Position = 6, Mandatory = $false)]
        [ScriptBlock]$OnFail = {}
    )
    
    [System.Net.WebClient]$client = New-Object System.Net.WebClient
    $client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR1.0.3705;)")
    $client.Headers.Add("content-type","application/x-www-form-urlencoded")
    $client.Headers.Add("Host","api.smstorrent.net")
    
    [string]$baseurl  = "http://api.smstorrent.net/http/"
    
    $Recipients | ForEach{
        $Recipient = $_
        If($Recipient.Length -gt 0){
            $Parameters = @{"user"=$Account;"cypher"=$Pass; "sender"=$Sender; "resp_type"="data"; "message"=$Body; "recipient"=$Recipient}
   
            [string]$postData = [string]::Empty

            $Parameters.Keys | ForEach{
                $postData += $_ + "=" + [System.Web.HttpUtility]::UrlEncode($Parameters[$_]) + "&"
            }
            $postData = $postData.TrimEnd("&")

            # Post the message to the SMS Gateway
            Try {
                $response = $client.UploadData($baseurl, "POST", [System.Text.Encoding]::ASCII.GetBytes($postData))
                $s = [System.Text.Encoding]::ASCII.GetString($response)
                & $OnSuccess
            } Catch {
                & $OnFail
            } 
            
            Write-ToFile ("" + $s); 
        }else{
             & $OnFail
        }        
    }    
    $client.Dispose()
}

Function Send-SMS_BySMSGator {
    Param(
        [string]$Sender = "SENDER", 
        
        [Parameter(Position = 1, Mandatory = $true)]
        [array]$Recipients = @(), 
        
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$Body,
        
        [Parameter(Position = 3, Mandatory = $false)]
        [string]$Account="vestude",
        
        [Parameter(Position = 4, Mandatory = $false)]
        [string]$Pass = "vestude",
        
        [Parameter(Position = 5, Mandatory = $false)]
        [ScriptBlock]$OnSuccess = {},
        
        [Parameter(Position = 6, Mandatory = $false)]
        [ScriptBlock]$OnFail = {}
    )
    
    [string]$baseurl  = "http://smsgator.com/bulksms"
    
    $Recipients | ForEach{
        $Recipient = $_
        If($Recipient.Length -gt 0){
            $Parameters = @{"email"=$Account;"password"=$Pass; "type"="0"; "dlr"="0"; "destination"=$Recipient; "sender"=$Sender; "message"=$Body;}
            
            # Post the message to the SMS Gateway
            Try {
                $response = Invoke-RestMethod $baseurl -Body $Parameters
                Write-ToFile ("" + $response); 
                & $OnSuccess -ArgumentList $response
            } Catch {
                Write-ToFile ("" + $response); 
                & $OnFail -ArgumentList $response
            }
            
        }else{
             & $OnFail
        }        
    }
}

#The entry point into the program
Function Run-Main {
    
    $config =  Get-Content "$currentLocation\config.json" | ConvertFrom-Json;
    
    
    [string]$server         = $config.ImapServer
    [string]$username       = $config.Username
    [string]$password       = $config.Password
    [string]$folder         = $config.Folder
    [string]$port           = $config.Port
    $isSSL                  = $config.IsSSL
    $emailSender            = $config.EmailSender
    $contentReg             = [regex]$config.Regex
    $replacement            = $config.RegexReplacement
    $extraLine              = [regex]"[\n\r]+"
    $smsRecipients          = $config.Recipients
    $smsSenderId            = $config.SmsSenderId
    $preferredProvider      = $config.PreferedProvider;
    $smsUsername            = ($config.SMSProviders | Select -ExpandProperty $preferredProvider).SmsGatewayUser
    $smsPassword            = ($config.SMSProviders | Select -ExpandProperty $preferredProvider).SmsGatewayPassword
        
    $securePassword         = $password | ConvertTo-SecureString -AsPlainText -Force;
    $MyCredential           = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securePassword
    
    Invoke-GmailSession -Credential $MyCredential -ScriptBlock {
        Param($session)
                
        #$inbox = $session | Get-Mailbox
        $headings = $session | Get-Mailbox | Get-Message -From $emailSender -Prefetch
        
        $headings | % {   
            
                # Convert the body html to plain text.
                $body = Html-ToText -html $_.Body
                                
                # Get only the portion of the body that is needed
                $thisMatch = $contentReg.Match($body);
                If($thisMatch.Success -eq $true){
                    $body = $extraLine.Replace($thisMatch.Result($replacement), "`n");
                    If ($body.Length -gt 0) {
                        
                        # Attempt to send the SMS
                        if($preferredProvider -eq "SMSTorrent"){
                            Send-SMS_BySMSTorrent -Sender $smsSenderId -Recipients $smsRecipients -Body $body -Account $smsUsername -Pass $smsPassword -OnSuccess {
                                
                                # Mark the email as read.
                                Update-Message -Message $_ -Read -Session $session
                            } -OnFail {
                                Write-ToFile ( "Something went wrong");
                            }
                        } else{
                            Send-SMS_BySMSGator -Sender $smsSenderId -Recipients $smsRecipients -Body $body -Account $smsUsername -Pass $smsPassword -OnSuccess {
                                
                                # Mark the email as read.
                                Update-Message -Message $_ -Read -Session $session
                            } -OnFail {
                                Write-ToFile ( "Something went wrong");
                            }#>
                        }
                    }
                }
            }
    }
}

Run-Main