Public Function Win32_NetworkAdapterConfiguration(strComputer,strDomain,strUser,strPassword,strHost)

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        
		Win32_NetworkAdapterConfiguration = "SystemName;" _
		& "instance" & ";" _
		& "objItem.MACAddress" & ";" _
		& "objItem.Description" & ";" _
		& "objItem.DHCPEnabled"& ";" _
		& "strIPAddress" & ";" _
		& "strIPSubnet" & ";" _
		& "objItem.IPConnectionMetric" & ";" _
		& "objItem.DHCPLeaseExpires" & ";" _
		& "objItem.DHCPLeaseObtained"& ";" _
		& "objItem.DHCPServer"& ";" _
		& "objItem.DNSDomain"& ";" _
		& "objItem.IPEnabled"& ";" _
		& "strDefaultIPGateway"& ";" _
		& "strGatewayCostMetric"& ";" _
		& "objItem.IPFilterSecurityEnabled"& ";" _
		& "objItem.IPPortSecurityEnabled"& ";" _
		& "strDNSDomainSuffixSearchOrder"& ";" _
		& "objItem.DNSEnabledForWINSResolution"& ";" _
		& "objItem.DNSHostName"      & ";" _
		& "strDNSServerSearchOrder"& ";" _
		& "objItem.DomainDNSRegistrationEnabled"& ";" _
		& "objItem.ForwardBufferMemory"& ";" _
		& "objItem.FullDNSRegistrationEnabled"& ";" _
		& "strGatewayCostMetric"& ";" _
		& "objItem.IGMPLevel"& ";" _
		& "objItem.Index"& ";" _
		& "strIPSecPermitIPProtocols"& ";" _
		& "strIPSecPermitTCPPorts"& ";" _
		& "strIPSecPermitUDPPorts"& ";" _
		& "objItem.IPUseZeroBroadcast"& ";" _
		& "objItem.IPXAddress"& ";" _
		& "objItem.IPXEnabled"& ";" _
		& "strIPXFrameType"& ";" _
		& "strIPXNetworkNumber"& ";" _
		& "objItem.IPXVirtualNetNumber"& ";" _
		& "objItem.KeepAliveInterval"& ";" _
		& "objItem.KeepAliveTime"& ";" _
		& "objItem.MTU"& ";" _
		& "objItem.NumForwardPackets"& ";" _
		& "objItem.PMTUBHDetectEnabled"& ";" _
		& "objItem.PMTUDiscoveryEnabled"& ";" _
		& "objItem.ServiceName"& ";" _
		& "objItem.SettingID"& ";" _
		& "objItem.TcpipNetbiosOptions"& ";" _
		& "objItem.TcpMaxConnectRetransmissions"& ";" _
		& "objItem.TcpMaxDataRetransmissions"& ";" _
		& "objItem.TcpNumConnections"& ";" _
		& "objItem.TcpUseRFC1122UrgentPointer"& ";" _
		& "objItem.TcpWindowSize"& ";" _
		& "objItem.WINSEnableLMHostsLookup"& ";" _
		& "objItem.WINSHostLookupFile"& ";" _
		& "objItem.WINSPrimaryServer"& ";" _
		& "objItem.WINSScopeID"& ";" _
		& "objItem.WINSSecondaryServer"& ";" _
		& "objItem.ArpAlwaysSourceRoute"& ";" _
		& "objItem.ArpUseEtherSNAP"& ";" _
		& "objItem.DatabasePath"& ";" _
		& "objItem.DeadGWDetectEnabled"& ";" _
		& "objItem.DefaultTOS"& ";" _
		& "objItem.DefaultTTL" 
        Exit Function
End if

  ' create new file
Set fpn = objFSO.OpenTextFile(CurrDir & results_path &"\Networks"&".csv", ForAppending,True)

On Error Resume Next
Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_NetworkAdapterConfiguration module on:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)        
		Win32_NetworkAdapterConfiguration = "N/A;" _
			& "N/A" & ";" _
			& "N/A" & ";" _
			& "N/A" & ";" _
			& "N/A"& ";" _
			& "N/A" & ";" _
			& "N/A" & ";" _
			& "N/A" & ";" _
			& "N/A" & ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";" _
			& "N/A"& ";"  & vbcrlf
        Exit Function
End if

Set colItems  = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True",,48)

i = 0

For Each objItem in colItems 

		i = i +1
		Wscript.Echo "-----------------------------------"
         Wscript.Echo "NetworkAdapterConfiguration instance: " & i
         Wscript.Echo "-----------------------------------"
         
        strDefaultIPGateway = GetMultiString_FromArray(objitem.DefaultIPGateway, ", ")
         Wscript.Echo "MACAddress                  : " & vbtab & objItem.MACAddress
         Wscript.Echo "Description                 : " & vbtab & objItem.Description
         Wscript.Echo "DHCPEnabled                 : " & vbtab & objItem.DHCPEnabled

         strIPAddress=GetMultiString_FromArray(objitem.IPAddress, ", ")
         Wscript.Echo "IPAddress                   : " & vbtab & strIPAddress
         strIPSubnet = GetMultiString_FromArray(objitem.IPSubnet, ", ")
         Wscript.Echo "IPSubnet                    : " & vbtab & strIPSubnet
         Wscript.Echo "IPConnectionMetric          : " & vbtab & objItem.IPConnectionMetric
         Wscript.Echo "DHCPLeaseExpires            : " & vbtab & objItem.DHCPLeaseExpires
         Wscript.Echo "DHCPLeaseObtained           : " & vbtab & objItem.DHCPLeaseObtained
         Wscript.Echo "DHCPServer                  : " & vbtab & objItem.DHCPServer
         Wscript.Echo "DNSDomain                   : " & vbtab & objItem.DNSDomain
         Wscript.Echo "IPEnabled                   : " & vbtab & objItem.IPEnabled
         Wscript.Echo "DefaultIPGateway            : " & vbtab & strDefaultIPGateway
         Wscript.Echo "GatewayCostMetric           : " & vbtab & strGatewayCostMetric
         Wscript.Echo "IPFilterSecurityEnabled     : " & vbtab & objItem.IPFilterSecurityEnabled
         Wscript.Echo "IPPortSecurityEnabled       : " & vbtab & objItem.IPPortSecurityEnabled

         strDNSDomainSuffixSearchOrder = GetMultiString_FromArray(objitem.DNSDomainSuffixSearchOrder, ", ")
         Wscript.Echo "DNSDomainSuffixSearchOrder  : " & vbtab & strDNSDomainSuffixSearchOrder
         Wscript.Echo "DNSEnabledForWINSResolution : " & vbtab & objItem.DNSEnabledForWINSResolution
         Wscript.Echo "DNSHostName                 : " & vbtab & objItem.DNSHostName
         
        strDNSServerSearchOrder = GetMultiString_FromArray(objitem.DNSServerSearchOrder, ", ")
         Wscript.Echo "DNSServerSearchOrder        : " & vbtab & strDNSServerSearchOrder
         Wscript.Echo "DomainDNSRegistrationEnabled: " & vbtab & objItem.DomainDNSRegistrationEnabled
         Wscript.Echo "ForwardBufferMemory         : " & vbtab & objItem.ForwardBufferMemory
         Wscript.Echo "FullDNSRegistrationEnabled  : " & vbtab & objItem.FullDNSRegistrationEnabled

         strGatewayCostMetric = GetMultiString_FromArray(objitem.GatewayCostMetric, ", ")
         Wscript.Echo "IGMPLevel                   : " & vbtab & objItem.IGMPLevel
         Wscript.Echo "Index                       : " & vbtab & objItem.Index

         strIPSecPermitIPProtocols = GetMultiString_FromArray(objitem.IPSecPermitIPProtocols, ", ")
         Wscript.Echo "IPSecPermitIPProtocols      : " & vbtab & strIPSecPermitIPProtocols

         strIPSecPermitTCPPorts =GetMultiString_FromArray(objitem.IPSecPermitTCPPorts, ", ")
         Wscript.Echo "IPSecPermitTCPPorts         : " & vbtab & strIPSecPermitTCPPorts

         strIPSecPermitUDPPorts = GetMultiString_FromArray(objitem.IPSecPermitUDPPorts, ", ")
         Wscript.Echo "IPSecPermitUDPPorts         : " & vbtab & strIPSecPermitUDPPorts

         Wscript.Echo "IPUseZeroBroadcast          : " & vbtab & objItem.IPUseZeroBroadcast
         Wscript.Echo "IPXAddress                  : " & vbtab & objItem.IPXAddress
         Wscript.Echo "IPXEnabled                  : " & vbtab & objItem.IPXEnabled

         strIPXFrameType=GetMultiString_FromArray(objitem.IPXFrameType, ", ")
         Wscript.Echo "IPXFrameType                : " & vbtab & strIPXFrameType

         strIPXNetworkNumber=GetMultiString_FromArray(objitem.IPXNetworkNumber, ", ")
         Wscript.Echo "IPXNetworkNumber            : " & vbtab & strIPXNetworkNumber
         Wscript.Echo "IPXVirtualNetNumber         : " & vbtab _
                 & objItem.IPXVirtualNetNumber
         Wscript.Echo "KeepAliveInterval           : " & vbtab _
                 & objItem.KeepAliveInterval
         Wscript.Echo "KeepAliveTime               : " & vbtab & objItem.KeepAliveTime
         Wscript.Echo "MTU                         : " & vbtab & objItem.MTU
         Wscript.Echo "NumForwardPackets           : " & vbtab & objItem.NumForwardPackets
         Wscript.Echo "PMTUBHDetectEnabled         : " & vbtab & objItem.PMTUBHDetectEnabled
         Wscript.Echo "PMTUDiscoveryEnabled        : " & vbtab & objItem.PMTUDiscoveryEnabled
         Wscript.Echo "ServiceName                 : " & vbtab & objItem.ServiceName
         Wscript.Echo "SettingID                   : " & vbtab & objItem.SettingID
         Wscript.Echo "TcpipNetbiosOptions         : " & vbtab & objItem.TcpipNetbiosOptions
         Wscript.Echo "TcpMaxConnectRetransmissions: " & vbtab & objItem.TcpMaxConnectRetransmissions
         Wscript.Echo "TcpMaxDataRetransmissions   : " & vbtab & objItem.TcpMaxDataRetransmissions
         Wscript.Echo "TcpNumConnections           : " & vbtab & objItem.TcpNumConnections
         Wscript.Echo "TcpUseRFC1122UrgentPointer  : " & vbtab & objItem.TcpUseRFC1122UrgentPointer
         Wscript.Echo "TcpWindowSize               : " & vbtab & objItem.TcpWindowSize
         Wscript.Echo "WINSEnableLMHostsLookup     : " & vbtab & objItem.WINSEnableLMHostsLookup
         Wscript.Echo "WINSHostLookupFile          : " & vbtab & objItem.WINSHostLookupFile
         Wscript.Echo "WINSPrimaryServer           : " & vbtab & objItem.WINSPrimaryServer
         Wscript.Echo "WINSScopeID                 : " & vbtab & objItem.WINSScopeID
         Wscript.Echo "WINSSecondaryServer         : " & vbtab & objItem.WINSSecondaryServer
         Wscript.Echo "ArpAlwaysSourceRoute        : " & vbtab & objItem.ArpAlwaysSourceRoute
         Wscript.Echo "ArpUseEtherSNAP             : " & vbtab & objItem.ArpUseEtherSNAP
         Wscript.Echo "DatabasePath                : " & vbtab & objItem.DatabasePath
         Wscript.Echo "DeadGWDetectEnabled         : " & vbtab & objItem.DeadGWDetectEnabled
         Wscript.Echo "DefaultTOS                  : " & vbtab & objItem.DefaultTOS
         Wscript.Echo "DefaultTTL                  : " & vbtab & objItem.DefaultTTL 
	
		
    fpn.Write ("" & strHost & ";" _
		& i & ";" _
		& objItem.MACAddress & ";" _
		& objItem.Description & ";" _
		& objItem.DHCPEnabled& ";" _
		& strIPAddress & ";" _
		& strIPSubnet & ";" _
		& objItem.IPConnectionMetric & ";" _
		& objItem.DHCPLeaseExpires& ";" _
		& objItem.DHCPLeaseObtained& ";" _
		& objItem.DHCPServer& ";" _
		& objItem.DNSDomain& ";" _
		& objItem.IPEnabled& ";" _
		& strDefaultIPGateway& ";" _
		& strGatewayCostMetric& ";" _
		& objItem.IPFilterSecurityEnabled& ";" _
		& objItem.IPPortSecurityEnabled& ";" _
		& strDNSDomainSuffixSearchOrder& ";" _
		& objItem.DNSEnabledForWINSResolution& ";" _
		& objItem.DNSHostName      & ";" _
		& strDNSServerSearchOrder& ";" _
		& objItem.DomainDNSRegistrationEnabled& ";" _
		& objItem.ForwardBufferMemory& ";" _
		& objItem.FullDNSRegistrationEnabled& ";" _
		& strGatewayCostMetric& ";" _
		& objItem.IGMPLevel& ";" _
		& objItem.Index& ";" _
		& strIPSecPermitIPProtocols& ";" _
		& strIPSecPermitTCPPorts& ";" _
		& strIPSecPermitUDPPorts& ";" _
		& objItem.IPUseZeroBroadcast& ";" _
		& objItem.IPXAddress& ";" _
		& objItem.IPXEnabled& ";" _
		& strIPXFrameType& ";" _
		& strIPXNetworkNumber& ";" _
		& objItem.IPXVirtualNetNumber& ";" _
		& objItem.KeepAliveInterval& ";" _
		& objItem.KeepAliveTime& ";" _
		& objItem.MTU& ";" _
		& objItem.NumForwardPackets& ";" _
		& objItem.PMTUBHDetectEnabled& ";" _
		& objItem.PMTUDiscoveryEnabled& ";" _
		& objItem.ServiceName& ";" _
		& objItem.SettingID& ";" _
		& objItem.TcpipNetbiosOptions& ";" _
		& objItem.TcpMaxConnectRetransmissions& ";" _
		& objItem.TcpMaxDataRetransmissions& ";" _
		& objItem.TcpNumConnections& ";" _
		& objItem.TcpUseRFC1122UrgentPointer& ";" _
		& objItem.TcpWindowSize& ";" _
		& objItem.WINSEnableLMHostsLookup& ";" _
		& objItem.WINSHostLookupFile& ";" _
		& objItem.WINSPrimaryServer& ";" _
		& objItem.WINSScopeID& ";" _
		& objItem.WINSSecondaryServer& ";" _
		& objItem.ArpAlwaysSourceRoute& ";" _
		& objItem.ArpUseEtherSNAP& ";" _
		& objItem.DatabasePath& ";" _
		& objItem.DeadGWDetectEnabled& ";" _
		& objItem.DefaultTOS& ";" _
		& objItem.DefaultTTL & ";" _
		& vbcrlf)

Next
	

temp_v = ""

Win32_NetworkAdapterConfiguration = ""
fpn.Close
Wscript.Echo "-----------------------------------"

End Function

Function GetMultiString_FromArray( ArrayString, Seprator)
     If IsNull ( ArrayString ) Then
         StrMultiArray = ArrayString
     else
         StrMultiArray = Join( ArrayString, Seprator )
    end if
    GetMultiString_FromArray = StrMultiArray
    
End Function

