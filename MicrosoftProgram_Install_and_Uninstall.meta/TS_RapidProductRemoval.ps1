#################################################################################
# Copyright (c) 2011, Microsoft Corporation. All rights reserved.
#
# You may use this code and information and create derivative works of it,
# provided that the following conditions are met:
# 1. This code and information and any derivative works may only be used for
# troubleshooting a) Windows and b) products for Windows, in either case using
# the Windows Troubleshooting Platform
# 2. Any copies of this code and information
# and any derivative works must retain the above copyright notice, this list of
# conditions and the following disclaimer.
# 3. THIS CODE AND INFORMATION IS PROVIDED `AS IS'' WITHOUT WARRANTY OF ANY 
# KIND, WHETHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. IF THIS
# CODE AND INFORMATION IS USED OR MODIFIED, THE ENTIRE RISK OF USE OR RESULTS IN
# CONNECTION WITH THE USE OF THIS CODE AND INFORMATION REMAINS WITH THE USER.
#################################################################################

param ($Action)

. .\MSIMATSFN.ps1

Import-LocalizedData -BindingVariable LocalizedStrings -FileName Strings
$RootCauseID = "RC_RapidProductRemoval"

$DiagID = "DIAG_msidiagDiagnostic"
$MarkName = "fixit_" + $DiagID + "_"
$MSDTs = @()
Get-WmiObject win32_process | where {($_.Name -eq "msdt.exe") -and ($_.getowner().user -eq $env:username)} | %{$MSDTs += $_}
if($MSDTs -ne $null)
{
    $ProcessID = $MSDTs[0].ProcessID
}
else
{
    $MATSWiz = @()
    Get-WmiObject win32_process | where {($_.Name -eq "MATSWiz.exe") -and ($_.getowner().user -eq $env:username)} | %{$MATSWiz += $_}
    $ProcessID = $MATSWiz[0].ProcessID
}
$MarkName += [string]$ProcessID

$alreadyDetected =  -not(Mark $MarkName)
if($alreadyDetected)
{
}
else
{
    #FirstRun
    if ($Action -eq "Uninstall")
    {
        #############################################################
        ############### UNINSTALL LOGIC #############################
        #############################################################

        $MasterHashs = ProductListingBuild #Calls dialog to list products
        $IID_ProducttoRemove = $LocalizedStrings.WindowsInstaller_IID_Not_Listed #default
        $IID_ProducttoRemove = Get-DiagInput -Id "IID_ProductRemoval" -Choice $MasterHashs -Parameter @{"IID_ProductRemoval_Dialog"=($LocalizedStrings.WindowsInstaller_IID_If_you_do_not_see_your_program_select_Not_Listed);"IID_ProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_Select_the_program_you_want_to_uninstall)} 


        if ($IID_ProducttoRemove -ne $LocalizedStrings.WindowsInstaller_IID_Not_Listed)
        {
            $Friendly=[MakeStringTest]::GetMSIProductInformation($IID_ProducttoRemove ,"ProductName")
			if($Friendly -eq "Unknown")
			{
				$Friendly=$LocalizedStrings.WindowsInstaller_IID_Name_no_available #If ProductName unknown
			}

            $RootCauseDetected = $true #___________Hand off to RS_WindowsInstaller to fix issue ____________
        }
        else
        {
            #Selected a product that is not listed in the Product List. So we will take a GUID
            $CorrectLength="false"
            while ($CorrectLength -eq "false")
            {
                [string]$IID_ProductRemovalReturn = Get-DiagInput -Id "IID_ManualProductRemoval" 
                if ($IID_ProductRemovalReturn.Length -eq 38)
                { 
                    $CorrectLength="true"
                    $IID_ProducttoRemove=$IID_ProductRemovalReturn
                    $Friendly=[MakeStringTest]::GetMSIProductInformation($IID_ProducttoRemove ,"ProductName")
					if($Friendly -eq "Unknown")
					{
						$Friendly=$LocalizedStrings.WindowsInstaller_IID_Name_no_available #If ProductName unknown
					}

                    $RootCauseDetected = $true #Hand off to RS to fix issue 
                }
                else
                {
                    $Manual_TryAgain = Get-DiagInput -Id "IID_Incorrect_GUID" 
                }
            }
        }
    
        if ($Friendly.length -eq 0)
        {
            $Friendly=$LocalizedStrings.WindowsInstaller_IID_Name_no_available #If ProductName not available
        }
                
        $IID_Install_Type_Return = Get-DiagInput -Id "IID_Install_Type" -Parameter @{"ProductCode"=$IID_ProducttoRemove;"RS_RapidProductRemoval_Dialog_SubTitle"=($LocalizedStrings.WindowsInstaller_IID_Click_cancel_to_exit_the_troubleshooter) ;"ProductName"=$Friendly;"RS_RapidProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_Uninstall_and_cleanup)}
        if ($IID_Install_Type_Return -eq ("True"))
        {
            update-diagrootcause -Id "RC_RapidProductRemoval" -detected $True -Parameter @{"ProductCode"=$IID_ProducttoRemove;"RS_RapidProductRemoval_Dialog_SubTitle"=($LocalizedStrings.WindowsInstaller_IID_Click_cancel_to_exit_the_troubleshooter) ;"ProductName"=$Friendly;"RS_RapidProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_Uninstall_and_cleanup)} 
        }
        else
        {
            update-diagrootcause -Id "RC_RapidProductRemoval" -detected $false  -Parameter @{"ProductCode"=$IID_ProducttoRemove;"RS_RapidProductRemoval_Dialog_SubTitle"=($LocalizedStrings.WindowsInstaller_IID_Click_cancel_to_exit_the_troubleshooter) ;"ProductName"=$Friendly;"RS_RapidProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_Uninstall_and_cleanup)}
        }
    }
    else
    {
        #############################################################
	    ############### INSTALL LOGIC ###############################
	    #############################################################

        $MasterHashs = ProductListingBuild
        $IID_ProducttoRemove = Get-DiagInput -Id "IID_ProductRemoval" -Choice $MasterHashs -Parameter @{"IID_ProductRemoval_Dialog"=($LocalizedStrings.WindowsInstaller_IID_If_you_do_not_see_your_program_select_Not_Listed);"IID_ProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_Select_the_program_you_are_trying_to_install)} 
        $Friendly=[MakeStringTest]::GetMSIProductInformation($IID_ProducttoRemove ,"ProductName")
    	if($Friendly -eq "Unknown")
		{
			$Friendly=$LocalizedStrings.WindowsInstaller_IID_Name_no_available #If ProductName unknown
		}

        if ($IID_ProducttoRemove -ne ($LocalizedStrings.WindowsInstaller_IID_Not_Listed))
        {
            $IID_Install_Type_Return = Get-DiagInput -Id "IID_Install_Type" -Parameter @{"ProductCode"=$IID_ProducttoRemove;"RS_RapidProductRemoval_Dialog_SubTitle"=($LocalizedStrings.WindowsInstaller_IID_Do_you_want_to_uninstall_this_program);"ProductName"=$Friendly;"RS_RapidProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_This_program_may_be_preventing_you_from_installing)} 
         
            if ($IID_Install_Type_Return -eq ("True"))
            {
                update-diagrootcause -Id "RC_RapidProductRemoval" -detected $true -Parameter @{"ProductCode"=$IID_ProducttoRemove;"RS_RapidProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_This_program_may_be_preventing_you_from_installing);"ProductName"=$Friendly;"RS_RapidProductRemoval_Dialog_SubTitle"=""} 
            }
            else
            {
                update-diagrootcause -Id "RC_RapidProductRemoval" -detected $false -Parameter @{"ProductCode"=$IID_ProducttoRemove;"RS_RapidProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_This_program_may_be_preventing_you_from_installing);"ProductName"=$Friendly;"RS_RapidProductRemoval_Dialog_SubTitle"=""} 
            }
        } 
        else
		{
             update-diagrootcause -Id "RC_RapidProductRemoval" -detected $false -Parameter @{"ProductCode"=$IID_ProducttoRemove;"RS_RapidProductRemoval_Dialog_Title"=($LocalizedStrings.WindowsInstaller_IID_This_program_may_be_preventing_you_from_installing);"ProductName"=$Friendly;"RS_RapidProductRemoval_Dialog_SubTitle"=""}  
        }
    }
}
# SIG # Begin signature block
# MIIaxwYJKoZIhvcNAQcCoIIauDCCGrQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+VrITd+/QVX76c6PYBi8xK1t
# naugghWCMIIEwzCCA6ugAwIBAgITMwAAAI1AOzwvrcZ7uAAAAAAAjTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUxMDA3MTgxNDA0
# WhcNMTcwMTA2MTgxNDA0WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjE0OEMtQzRCOS0yMDY2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAmlIcRUD3lFDw
# EK27FJ0L7nJcRvoCDIV9j+iH+phgZsI1T1UGjG7uVQs4h9Mb2JBiK5aScts3g08g
# LnHe5pCo2MWegFpCGe1iEOirtCDXABUGrJ9ZiiCGKfsEXsNPVsY++5zSMY/12JF6
# l7bpGQJelF6hlPrL2XDrOQZtg2MCfZB0FJXRsFR/IxkLriSCe6GrmTNlgTubfcrW
# YFtYRXYJ+SNfM4YCFuygR4Tajdl8gKdJPgtbSls1CS6q+XUgCxusXijAemjvDnd9
# HoFms2vuNAhkbfL1h8TF5Sgt3ZkOqgtrvw3c3d7hTiE2XhnokXxz3g7KwVRzwapf
# YhWha1xNPwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFO+w/NAncAPOnsWMxNE+EIfv
# sAbaMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBACJRuAWCXkNgGnPI5wLB03Kk2xYLTFzuC0uL6Pa4kQ6OlZ5O
# lJ8D1ISf29/eec92Gw/0oSi4XmgLE6yICiYYJUI+0esJjKj79iMyb3+Qo3lwgLBD
# ZPqg4rUXT32fZFHURjPJ0DmCV8JLhIllwTHp1tjrnpQwTq0oqeO8rbwrU0hvsgSv
# esdvxnu54IzsJnSWOsmpE3RCC0fTrnHIsU382SoqUnB5AlPK64q3ImJcQz0Ns5O5
# yUsWy0ef50ExRbA9tvBlvCZtVfCwLH/6U7MLoJcSkwpB7+VPlklyOxDBM8TXvRcE
# rgdbRaiScYxbKZGkb8JwJxLKt3KO4D8JT5MxVxYwggTsMIID1KADAgECAhMzAAAB
# Cix5rtd5e6asAAEAAAEKMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBMB4XDTE1MDYwNDE3NDI0NVoXDTE2MDkwNDE3NDI0NVowgYMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# HjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJL8bza74QO5KNZG0aJhuqVG+2MWPi75R9LH7O3HmbEm
# UXW92swPBhQRpGwZnsBfTVSJ5E1Q2I3NoWGldxOaHKftDXT3p1Z56Cj3U9KxemPg
# 9ZSXt+zZR/hsPfMliLO8CsUEp458hUh2HGFGqhnEemKLwcI1qvtYb8VjC5NJMIEb
# e99/fE+0R21feByvtveWE1LvudFNOeVz3khOPBSqlw05zItR4VzRO/COZ+owYKlN
# Wp1DvdsjusAP10sQnZxN8FGihKrknKc91qPvChhIqPqxTqWYDku/8BTzAMiwSNZb
# /jjXiREtBbpDAk8iAJYlrX01boRoqyAYOCj+HKIQsaUCAwEAAaOCAWAwggFcMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBSJ/gox6ibN5m3HkZG5lIyiGGE3
# NDBRBgNVHREESjBIpEYwRDENMAsGA1UECxMETU9QUjEzMDEGA1UEBRMqMzE1OTUr
# MDQwNzkzNTAtMTZmYS00YzYwLWI2YmYtOWQyYjFjZDA1OTg0MB8GA1UdIwQYMBaA
# FMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8w
# OC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMx
# LTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCmqFOR3zsB/mFdBlrrZvAM2PfZ
# hNMAUQ4Q0aTRFyjnjDM4K9hDxgOLdeszkvSp4mf9AtulHU5DRV0bSePgTxbwfo/w
# iBHKgq2k+6apX/WXYMh7xL98m2ntH4LB8c2OeEti9dcNHNdTEtaWUu81vRmOoECT
# oQqlLRacwkZ0COvb9NilSTZUEhFVA7N7FvtH/vto/MBFXOI/Enkzou+Cxd5AGQfu
# FcUKm1kFQanQl56BngNb/ErjGi4FrFBHL4z6edgeIPgF+ylrGBT6cgS3C6eaZOwR
# XU9FSY0pGi370LYJU180lOAWxLnqczXoV+/h6xbDGMcGszvPYYTitkSJlKOGMIIF
# vDCCA6SgAwIBAgIKYTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
# EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
# MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBC
# mXZTbD4b1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1VwqJyq4gSfTw
# aKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNEf9eB2/O98jakyVxF3K+tPeAoaJcap6Vy
# c1bxF5Tk/TWUcqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2L1TdFDBZ
# +NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049oDI9kM2hOAaFXE5WgigqBTK3S9dP
# Y+fSLWLxRT3nrAgA9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzBrAlf
# A9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFMsR6MrS
# tBZYAck3LjMWFrlMmgofMAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+mPLzYLTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTAfBgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnk
# pDBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEE
# SDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+
# fyZGr+tvQLEytWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAhXaw9L0y6
# oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6WcIC36C1DEVs0t40rSvHDnqA2iA6VW
# 4LiKS1fylUKc8fPv7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WRTsgb
# 0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZEolHN+emjWFbdmwJFRC9f9Nqu
# 1IIybvyklRPk62nnqaIsvsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
# NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxPbQ4TTj18KUicctHzbMrB
# 7HCjV5JXfZSNoBtIA1r3z6NnCnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDord
# EN5k9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7B2BHZWBATrBC7E7t
# s3Z52Ao0CW0cgDEf4g5U3eWh++VHEK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jsh
# rg1cnPCiroZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONxPdcAfmJH0c6I
# ybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3UAjWwz0wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBK8wggSr
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAABCix5rtd5e6as
# AAEAAAEKMAkGBSsOAwIaBQCggcgwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFJeF
# CLeau8K7R2UuPrQZwE9bKvQSMGgGCisGAQQBgjcCAQwxWjBYoDaANABUAFMAXwBS
# AGEAcABpAGQAUAByAG8AZAB1AGMAdABSAGUAbQBvAHYAYQBsAC4AcABzADGhHoAc
# aHR0cDovL3N1cHBvcnQubWljcm9zb2Z0LmNvbTANBgkqhkiG9w0BAQEFAASCAQAC
# WpId7IlHXFoFInrwnFq6p7kqDwR0zyMPJwHxVe7Y4f05XKMuCVEOUERi9E+MK5f3
# jEYhJtaGUdpwU/ux+OtmznUcX512njJF4u+VNRNAhRrjWQ13Ci46jGHCXcz9LVWs
# nzSlCAXI4jO3TanW4Ok7eGNfupJtpuWKUaYErh5HhW1x4hhoOdPcZVRz6xitqR4L
# 4s/hEnCxrJxVP+e+mmIJNyN0iMxajqi/PQbkKGds/lwxX2QyA9rjwL82towdiWwy
# CyZt5UsgMWvF7hiC6aU175w9vxEXPg/0wDljY6YYVmaJtjJv726cqu5QRqmUyRm7
# 3JAThWYBhA3Yj7VmTQWJoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGO
# MHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMT
# GE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAI1AOzwvrcZ7uAAAAAAAjTAJ
# BgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0B
# CQUxDxcNMTUxMjE3MDg1MzA3WjAjBgkqhkiG9w0BCQQxFgQU9ZH8FJv0/Nuo9+SR
# n/HN6lnMm/swDQYJKoZIhvcNAQEFBQAEggEAeR1U26J/FDPgPY5PLyI/qTPBqnMZ
# tccJej5okXIIv/5SXFpPaf4LtP1uEJvJxw+StUu33ofSqcnx8kXA/EbX2n7jbd49
# 6vpv7os1/vhYdGEWPK7FJabYlPkgH1bwUk6jwZPPmywwTx7ayav7j+zXZ0MdO8jN
# 9eUuTvZAu7Yw/JD9ynpge9n1ORx1X4pbcVLg43EKRDt8k5T8OVe0oiYjoHY1I1hF
# 6O/y1p4/2HwuCT1yDejTe/3GwUGgLp3OvuucntBAq5q2FM+/T71OVIpnKrsUSAlZ
# 4a/ruUNe2wAb4g7c9jAjgiWDLr7icmnGibyAP7CfKU0pwCQxUZS1MVkGyQ==
# SIG # End signature block
