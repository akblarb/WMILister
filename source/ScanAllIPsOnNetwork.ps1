###Credit to Andrew Ghobrial for building the logic for the function "scanPort"

Function ConvertIPToBin ($IPToBin){
    $temp = $IPToBin -split '\.' | ForEach-Object {
        [System.Convert]::ToString($_,2).PadLeft(8,'0')
    }
    Return $temp -join ''
    
}

Function ConvertIPBinToDec ($BinToIP)
{
    $temp = $BinToIP -split '\.' | ForEach-Object {
        #Write-Host "here:" $_
        [System.Convert]::ToByte($_,2)
    }
    Return $temp -join '.'
}

Function FindCIDR ($SubnetToCount)
{
    $tempCount = 0
    ForEach ($char in [char[]]$SubnetToCount){
        If ($char -eq '1') {$tempCount +=1}
        #Write-Host $tempCount ":" $char
    }
    Return $tempCount
}

Function FindNetAddr ($ipToExam, $foundCIDR)
{
    $tempCount = $foundCIDR
    $tempOctet = 8
    $temp = ""
    ForEach ($char in [char[]]$ipToExam){
        If ($tempCount -gt 0){
            #Write-Host $tempCount ":" $char
            $temp = -join($temp, $char)
            $tempCount -=1
            $tempOctet -=1
            If ($tempOctet -eq 0)
			{
                $temp = -join($temp, ".")
                $tempOctet = 8
            }
        } Else {
            #Write-Host $tempCount ":" 0
            $temp = -join($temp, '0')
            $tempOctet -=1
            If ($tempOctet -eq 0)
			{
                $temp = -join($temp, ".")
                $tempOctet = 8
            }
        }
    }
    Return $temp.TrimEnd(".")
}

#Might Kill this function
Function countOctets ($bitsOnly)
{
    $tempCount = $foundCIDR
    $tempOctet = 8
    $temp = ""
    ForEach ($char in [char[]]$bitsOnly)
	{
        #Write-Host $tempCount ":" $char
        $temp = -join($temp, $char)
        $tempCount -=1
        $tempOctet -=1
        If ($tempOctet -eq 0)
		{
            $temp = -join($temp, ".")
            $tempOctet = 8
        }
    }
    $tempArray = $temp.TrimEnd(".").ToCharArray()
    [array]::Reverse($tempArray)
    $temp = $tempArray -join ''
    $countOctets = [Math]::Truncate(($temp.ToCharArray() | Where-Object {$_ -eq '0'} | Measure-Object).Count/8)
    #Return $countOctets
    Return $temp, $countOctets
    #Return $countOctets
}

#$range must be specified like 1..255 or 1..16
Function MakeList ($min, $max, $fullOctets)
{
    #$IPtoscan = 192.168.2
    $range = [int]$min..[int]$max
    $countRange = $range.Count
    $IPs = forEach ($r in $range)
	{
        $ip = "{0}{1}" -F $fullOctets,$r
        #Write-Host $ip
        Write-Progress "Building List" $ip -PercentComplete (($r/$countRange)*100)
		Write-Host "IP:" $ip
        
        
    }
    #write-host Found online hosts with port $ports open: -foregroundcolor "Yellow"
    
}
Function MakeListNew ($min, $max)
{
	for ($i=[int]$min; $i -le [int]$max; $i++)
	{
		Write-Host "look at that" $i
	}
}


Function GetMax ($IPLastPart)
{
    ##Need to convert to all ones
    $ipOnes = ""
    ForEach ($char in [char[]]$IPLastPart){
        if ($char -eq "0") {
            $ipOnes = -join($ipOnes, "1")
        } else {
            $ipOnes = -join($ipOnes, ".")
        }
        #Write-Host $char
        #Write-Host $ipOnes
    }
    $ArrayOctets = $IPLastPart.Split(".")
    return $ipOnes
}

Function GetMaxLeadZero ($IPLastPart)
{
    ##Need to convert to all ones
    $ipOnes = ""
    ForEach ($char in [char[]]$IPLastPart){
        if ($char -eq "0") {
            $ipOnes = -join($ipOnes, "1")
        } elseif ($char -eq "1"){
            $ipOnes = -join($ipOnes, "0")
        } else {
            $ipOnes = -join($ipOnes, ".")
        }
        #Write-Host $char
        #Write-Host $ipOnes
    }
    $ArrayOctets = $IPLastPart.Split(".")
    return $ipOnes
}

Function startCounting ($NetAddress, $subnetBin)
{
    $tempCount = 8
    $ntadd = $NetAddress -split '\.'
    #Write-Host $ntadd[0]
    #Write-Host $ntadd[1]
    #Write-Host $ntadd[2]
    #Write-Host $ntadd[3]
    
    ForEach ($char in [char[]]$subnetBin){
        #Write-Host $tempCount ":" $char
        $temp = -join($temp, $char)
        $tempCount -=1
        $tempOctet -=1
        If ($tempCount -eq 0){
            $temp = -join($temp, ".")
            $tempCount = 8
        }
    }
    $subnetBinDots = $(ConvertIPBinToDec $(GetMaxLeadZero($temp.TrimEnd("."))))
    
    #Write-Host "subBinNet:" $subnetBinDots
    #Write-Host $(ConvertIPBinToDec $subnetBinDots)
    
    <#
        010.000.000.000
        255.255.000.000
        [System.Convert]::ToByte($_,2)
    #>
    
    ##$currAddrPart = ""
	
    $tmpCount = 0
    $Subnt = $subnetBinDots -split '\.' | ForEach-Object {
		#if ($tmpCount -ne 0) {$currAddrPart = $currAddrPart + [string]([int]$ntadd[$tmpCount-1]) + "."}
		#Write-Host $currAddrPart
        If ($_ -eq 255)
		{
            #Write-Host "some" $ntadd[$tmpCount]
            #Write-Host "[InvertedSub" $_"] count:" $tmpCount ":" $NetAddress "range255:" $ntadd[$tmpCount]"255"
			$addrRange = $addrRange + $ntadd[$tmpCount]+"-"+"255"
			If ($tmpCount -ne 3) {$addrRange = $addrRange + "."}
			#MakeList $ntadd[$tmpCount] 255 $currAddrPart
        } Else {
            #Write-Host "[InvertedSub" $_"]: " 255 "-" $ntadd[$tmpCount] "="(255-$($ntadd[$tmpCount]))
            #Write-Host (($ntadd[$tmpCount] -as [int])+($_ -as [int]))
            #Write-Host "[InvertedSub" $_"] count:" $tmpCount ":" $NetAddress "range:" $ntadd[$tmpCount](($ntadd[$tmpCount] -as [int])+($_ -as [int]))
			$addrRange = $addrRange + $ntadd[$tmpCount]+"-"+(($ntadd[$tmpCount] -as [int])+($_ -as [int]))
			If ($tmpCount -ne 3) {$addrRange = $addrRange + "."}
			#MakeList $ntadd[$tmpCount] (($ntadd[$tmpCount] -as [int])+($_ -as [int])) $currAddrPart
        }
        $tmpCount += 1
		#Write-Host "octets so far:" $tmpCount
    }
	#for ($i=[int]$min; $i -le [int]$max; $i++) 
	#Write-Host "range"$addrRange
	$countIP=0
	$arrAddrRange = $addrRange -split "\."
	$firstOctetMinMax = $arrAddrRange[0] -split "-"
	$secondOctetMinMax = $arrAddrRange[1] -split "-"
	$thirdOctetMinMax = $arrAddrRange[2] -split "-"
	$fourthOctetMinMax = $arrAddrRange[3] -split "-"
	
	for ($frst=[int]$firstOctetMinMax[0]; $frst -le [int]$firstOctetMinMax[1]; $frst++)
	{
		#Write-Host "1st - "$frst
		for ($scnd=[int]$secondOctetMinMax[0]; $scnd -le [int]$secondOctetMinMax[1]; $scnd++)
		{
			#Write-Host "2nd - "$frst"."$scnd"."
			for ($thrd=[int]$thirdOctetMinMax[0]; $thrd -le [int]$thirdOctetMinMax[1]; $thrd++)
			{
				#Write-Host "3rd - "$frst"."$scnd"."$thrd"."
				for ($frth=[int]$fourthOctetMinMax[0]; $frth -le [int]$fourthOctetMinMax[1]; $frth++)
				{
					$countIP +=1
					$ip = [string]$frst + "." + [string]$scnd + "." + [string]$thrd + "." + [string]$frth
					#Write-Host "4th - "$frst"."$scnd"."$thrd"."$frth           "- Working on IP #"$countIP "of" $TotalIPs
					Write-Progress "Scanning for online computers" $ip" - working on IP "$countIP" / "$TotalIPs" - With timeout of: "$optTimeout"ms" -PercentComplete (($countIP/$TotalIPs)*100)
					Add-Content $logIPsAll $ip  -Encoding ASCII
					scanPort $ip $optTimeout
				}
			}
		}
	}
	
	

                        
        #
}


Function scanPort ($ip, $timeout_ms)
{
	$ports = 135
	#$timeout_ms = 100
	forEach ($port in $ports)
	{
		Add-Content $logIPsTested $ip  -Encoding ASCII
		$ErrorActionPreference = 'SilentlyContinue'
		$socket = new-object System.Net.Sockets.TcpClient
		$connect = $socket.BeginConnect($ip, $port, $null, $null)
		$tryconnect = Measure-Command { $success = $connect.AsyncWaitHandle.WaitOne($timeout_ms, $true) } 
		$tryconnect | Out-Null

		If ($socket.Connected)
		{
			#$ip 
			$socket.Close()
			$socket.Dispose()
			$socket = $null
			Add-Content $logIPsOnline $ip  -Encoding ASCII
			Write-Host "IP Found: "$ip
		}
		
		$ErrorActionPreference = 'Continue'
	}
}


Function scanCurrNetAdapts()
{
	$countActiveAdapters = (Get-WmiObject Win32_NetworkAdapterConfiguration | ? {$_.IPEnabled -and $_.IPAddress[0] -ne "127.0.0.1"} | measure).count
	Write-Host "Total number of active Network Adapters: "$countActiveAdapters
	ForEach ($Item in (Get-WmiObject Win32_NetworkAdapterConfiguration | ? {$_.IPEnabled -and $_.IPAddress[0] -ne "127.0.0.1"}))
	{
		Write-Host "---Working On "$Item.IPAddress
		$IPAddress = $Item.IpAddress[0]
		$SubnetMask = $Item.IPSubnet[0]
		$Gateway = $Item.DefaultIPGateway[0]
		
		$ipBin = ConvertIPToBin $IPAddress
		$subnetBin = ConvertIPToBin $SubnetMask
		
		#Write-Host "Binary IP:" $ipBin
		Write-Host "---Binary Subnet Mask:" $subnetBin
		
		$CIDR = FindCIDR $subnetBin
		$netAddrBin = FindNetAddr $ipBin $CIDR
		
		#Write-Host "Binary Network Address:" $netAddrBin
		
		$NetAddress = ConvertIPBinToDec $netAddrBin
		$TotalIPs = [math]::pow(2, (32-$CIDR))
		Write-Host "IP:" $IPAddress -ForegroundColor "Green"
		Write-Host "Sub:" $SubnetMask "CIDR:" $CIDR -ForegroundColor "Green"
		Write-host "Network Address:" $NetAddress"/"$CIDR -ForegroundColor "Green"
		Write-Host "Total IPs:" $TotalIPs -ForegroundColor "Green"
		Write-Host "Gate:" $Gateway -ForegroundColor "Green"
		
		#Write-Host "Now Do Someting inside of this foreach loop to properly scan/List all possible IP addresses" -ForegroundColor "Blue"
		
		
		$NetAddressBits = $netAddrBin -replace '[.]',''
		#Write-Host "Network Address JustTheBits:       " $NetAddressBits
		
		$NetAddFirstHalf = $NetAddressBits.Substring(0,$CIDR)
		
		#Write-Host "Network Address JustThe_FIRST_Bits:" $NetAddFirstHalf
		
		$NetAddLastHalf = $NetAddressBits.Substring($CIDR)
		#Write-Host "Network Address JustThe_LAST_Bits:" $NetAddLastHalf
		
		
		$NetAddLastHalfZeroes, $fullOctets = countOctets $NetAddLastHalf
		$NetAddLastHalfOnes = GetMax $NetAddLastHalfZeroes
		#need to take $NetAddLastHalf and invert from all zeros to all ones.
		#Then seperate into octets and convert back to decimal.
		#after that I can loop through counting for each possible IP.
		
		#Write-Host "LastOctets:" $fullOctets
		#Write-Host "Last bits of IP that are the range LOW: " $NetAddLastHalfZeroes -ForegroundColor "Red"
		#Write-Host "Last bits of IP that are the range HIGH:" $NetAddLastHalfOnes -ForegroundColor "Red"
		#Write-Host "After counting octets that have full 0-255 range (var fullOctets), start making lists of possible IPs" -ForegroundColor "Blue"
		#Write-Host "Still need to account for the partial octets and how to list those" -ForegroundColor "Blue"
		
		
		startCounting $NetAddress $subnetBin
	}
	scanWMI
}

Function convertCIDRtoBin($locCIDR)
{
	$binNetMask = ""
	For ($i=1; $i -le 32; $i++)
	{
		If ($i -le $locCIDR)
		{
			$binNetMask = -join($binNetMask, "1")
		} Else {
			$binNetMask = -join($binNetMask, "0")
		}
	}
	Return $binNetMask
}

Function scanCustomCIDR($custIPAdd, $custCIDR)
{
	$ipBin = ConvertIPToBin $custIPAdd
	##$subnetBin = ConvertIPToBin $SubnetMask
	
	Write-Host "Binary IP:" $ipBin
	#Write-Host "Binary Subnet Mask:" $subnetBin
	
	$CIDR = $custCIDR
	$subnetBin = convertCIDRtoBin $CIDR
	Write-Host "--Binary Subnet Mask:" $subnetBin
	$netAddrBin = FindNetAddr $ipBin $CIDR
	
	Write-Host "Binary Network Address:" $netAddrBin
	
	$NetAddress = ConvertIPBinToDec $netAddrBin
	$TotalIPs = [math]::pow(2, (32-$CIDR))
	Write-Host "IP:" $custIPAdd -ForegroundColor "Green"
	##Write-Host "Sub:" $SubnetMask "CIDR:" $CIDR -ForegroundColor "Green"
	Write-host "Network Address:" $NetAddress"/"$CIDR -ForegroundColor "Green"
	Write-Host "Total IPs:" $TotalIPs -ForegroundColor "Green"
	##Write-Host "Gate:" $Gateway -ForegroundColor "Green"
	startCounting $NetAddress $subnetBin
	scanWMI
}

Function scanWMI()
{
	#alter command as needed.  Either use # to comment out active line or add switches /L (log only) or /F (force clean)
	ForEach ($line in Get-Content $logIPsOnline)
	{
		Write-Host "Running WMILister on IP $($line)" -ForegroundColor "Blue"
		#cscript //nologo WMILister.vbs /s:$($line)
		cscript //nologo WMILister.vbs /s:$($line) $switchesWMIList
		#cscript //nologo WMILister.vbs /s:$($line) /F
		
	}
	exit
}


#MAIN

$logIPsOnline = "Online_IPs.txt"
Out-File -FilePath $logIPsOnline -Encoding ASCII

$logIPsAll = "All_IPs.txt"
Out-File -FilePath $logIPsAll -Encoding ASCII

$logIPsTested = "Tested_IPs.txt"
Out-File -FilePath $logIPsTested -Encoding ASCII


Function DefaultOptions
{
	$global:optCleanLogOnly = "*"
	$global:optCleanPromptEach = " "
	$global:optCleanForce = " "
	
	$global:switchesWMIList = "/L"
	
	$global:optScanCurrSubnet = "*"
	$global:optScanCustomCIDR = " "
	
	$global:optTimeout = 500
	
	$global:optIPAddr = ""
	$global:optCIDR = ""
	$global:netAndCIDR = ""
}

Function PrintOptions()
{
	clear
	Write-Host "This utility is intended allow scanning of an entire network
for WMI Persistent Threats by using WMILister.
You can download WMILister from:"
	Write-Host "    -- https://www.xednaps.com/download/wmilister/
" -ForegroundColor "Green"
	Write-Host "!!!This tool comes without warranty!!!"
	Write-Host "!!!Use at own risk or at the advisement of an expert from forum.eset.com!!!
"
	Write-Host "Cleaning Options
	($optCleanLogOnly) 1. Log Only - No cleaning
	($optCleanPromptEach) 2. Enable Cleaning - Prompt for each	
	($optCleanForce) 3. Force Cleaning - No prompts
"

	Write-Host "Scan Options
	($optScanCurrSubnet) A. Use active network adapters to scan current subnet/s
	($optScanCustomCIDR) B. Specify custom CIDR Range $netAndCIDR"

	Write-Host "
	Z. Change timeout [$optTimeout ms]
	
	0. Exit
"
	
	
	
	
}

Function callPrompt()
{
	Write-Host "To start a scan of connected subnet, simply press [ENTER]
" -ForegroundColor "Green"
	$choice = Read-Host -prompt "Enter an option to change, or press enter to start scanning"
	SetOptions $choice
}


Function SetOptions($choice)
{
	If (($choice -ge 1) -and ($choice -le 3))
	{
		If ($choice -eq 1) 
		{
			$switchesWMIList = "/L"
			
			$optCleanLogOnly = "*"
			$optCleanPromptEach = " "
			$optCleanForce = " "
		}
		ElseIf ($choice -eq 2) 
		{
			$switchesWMIList = ""
			
			$optCleanLogOnly = " "
			$optCleanPromptEach = "*"
			$optCleanForce = " "
		}
		ElseIf ($choice -eq 3) 
		{
			$switchesWMIList = "/F"
			
			$optCleanLogOnly = " "
			$optCleanPromptEach = " "
			$optCleanForce = "*"
		}
	}
	ElseIf (($choice.ToUpper() -eq "A") -or ($choice.ToUpper() -eq "B"))
	{
		If ($choice.ToUpper() -eq "A")
		{
			$optScanCurrSubnet = "*"
			$optScanCustomCIDR = " "
		}
		If ($choice.ToUpper() -eq "B")
		{
			$optScanCurrSubnet = " "
			$optScanCustomCIDR = "*"
			$optIPAddr = Read-Host -prompt "    Enter IP or Network address"
			$optCIDR = Read-Host -prompt "    Enter CIDR notation"
			$netAndCIDR = "[$optIPAddr/$optCIDR]"
		}
	}
	ElseIf ($choice -eq 0)
	{
		exit
	}
	ElseIf ($choice.ToUpper() -eq "Z")
	{
		$optTimeout = Read-Host -prompt "    Enter timeout in miliseconds"
		If ($optTimeout -eq "") {$optTimeout = 300}
	}
	ElseIf ($choice -eq "")
	{
		If ($optScanCurrSubnet -eq "*") {scanCurrNetAdapts}
		ElseIf ($optScanCustomCIDR -eq "*") {scanCustomCIDR $optIPAddr $optCIDR}
	}
	printOptions
	callPrompt
}


# Main
DefaultOptions
PrintOptions
callPrompt









#scanPort("10.0.0.16")
#scanPort("10.0.0.17")
