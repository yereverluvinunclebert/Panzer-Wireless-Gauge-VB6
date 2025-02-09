Attribute VB_Name = "modWireless"
Option Explicit

'Private Enum DOT11_PHY_TYPE
'    dot11_phy_type_unknown = 0
'    dot11_phy_type_any = 0
'    dot11_phy_type_fhss = 1
'    dot11_phy_type_dsss = 2
'    dot11_phy_type_irbaseband = 3
'    dot11_phy_type_ofdm = 4
'    dot11_phy_type_hrdsss = 5
'    dot11_phy_type_erp = 6
'    dot11_phy_type_ht = 7
'    dot11_phy_type_IHV_start = &H80000000
'    dot11_phy_type_IHV_end = &HFFFFFFFF
'End Enum

'Private Enum DOT11_BSS_TYPE
'    dot11_BSS_type_infrastructure = 1
'    dot11_BSS_type_independent = 2
'    DOT11_BSS_TYPE_ANY = 3
'End Enum

Private Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Type WLAN_INTERFACE_INFO
    ifGuid As GUID
    InterfaceDescription(511) As Byte
    IsState As Long
End Type


Private Type DOT11_SSID
    uSSIDLength As Long
    ucSSID(31) As Byte
End Type

'Private Type WLAN_RATE_SET
'    uRateSetLength As Long
'    usRateSet(125) As Integer
'End Type
'
'Private Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type

Private Type WLAN_AVAILABLE_NETWORK
    strProfileName(511) As Byte
    dot11Ssid As DOT11_SSID
    dot11BssType As Long
    uNumberOfBssids As Long
    bNetworkConnectable As Long
    wlanNotConnectableReason As Long
    uNumberOfPhyTypes As Long
    dot11PhyTypes(7) As Long
    bMorePhyTypes As Long
    wlanSignalQuality As Long
    bSecurityEnabled As Long
    dot11DefaultAuthAlgorithm As Long
    dot11DefaultCipherAlgorithm As Long
    dwFlags As Long
    dwreserved As Long
End Type

Private Type AVAILABLE_NETWORK
    dot11Ssid As DOT11_SSID
    dot11BssType As Long
    uNumberOfBssids As Long
    bNetworkConnectable As Long
    wlanNotConnectableReason As Long
    uNumberOfPhyTypes As Long
    dot11PhyTypes(7) As Long
    bMorePhyTypes As Long
    wlanSignalQuality As Long
    bSecurityEnabled As Long
    dot11DefaultAuthAlgorithm As Long
    dot11DefaultCipherAlgorithm As Long
    dwFlags As Long
    dwreserved As Long
End Type


'Private Type WLAN_BSS_ENTRY
'    dot11Ssid As DOT11_SSID
'    uPhyId As Long
'    dot11Bssid(7) As Byte
'    dot11BssType As DOT11_BSS_TYPE
'    dot11BssPhyType As DOT11_PHY_TYPE
'    lRssi As Long
'    uLinkQuality As Long
'    bInRegDomain As Long 'Boolean
'    usBeaconPeriod As Long
'    ullTimestamp As FILETIME
'    ullHostTimestamp As FILETIME
'    usCapabilityInformation As Long
'    ulChCenterFrequency As Long
'    wlanRateSet As WLAN_RATE_SET
'    ulIeOffset As Long
'    ulIeSize As Long
'End Type

Private Type WLAN_INTERFACE_INFO_LIST
    dwNumberofItems As Long
    dwIndex As Long
    InterfaceInfo As WLAN_INTERFACE_INFO
End Type

Private Type WLAN_AVAILABLE_NETWORK_LIST
    dwNumberofItems As Long
    dwIndex As Long
    Network As WLAN_AVAILABLE_NETWORK
End Type

Private Type WLAN_BSS_LIST
    dwTotalSize As Long
    dwNumberofItems As Long
    wlanBssEntries As Long
End Type

Private Declare Function WlanOpenHandle Lib "wlanapi.dll" (ByVal dwClientVersion As Long, ByVal pdwReserved As Long, ByRef pdwNegotiaitedVersion As Long, ByRef phClientHandle As Long) As Long
Private Declare Function WlanEnumInterfaces Lib "wlanapi.dll" (ByVal hClientHandle As Long, ByVal pReserved As Long, ppInterfaceList As Long) As Long
Private Declare Function WlanGetAvailableNetworkList Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, ByVal dwFlags As Long, ByVal pReserved As Long, ppAvailableNetworkList As Long) As Long
Private Declare Function WlanScan Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, pDot11Ssid As Long, pIeData As Long, reserved As Long) As Long
Private Declare Function WlanGetNetworkBssList Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGui As GUID, ByVal pDot11Ssid As Long, ByVal dot11BssType As Long, ByVal bSecurityEnabled As Long, ByVal pReserved As Long, ppWlanBssList As Long) As Long
Private Declare Function WlanGetProfile Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, ByVal strProfileName As Long, ByVal pReserved As Long, pstrProfileXml As Long, pdwFlags As Long, pdwGrantedAccess As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub WlanFreeMemory Lib "wlanapi.dll" (ByVal pMemory As Long)
'Private Declare Function lstrlenW& Lib "kernel32.dll" (ByVal lpszSrc&)

Private udtList As WLAN_INTERFACE_INFO_LIST
Private udtBSSList As WLAN_BSS_LIST
Private ConIndex As Integer
Private lHandle As Long
Private lVersion As Long
Private Connected As String
Private bBuffer() As Byte

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)





' ----------------------------------------------------------------
' Procedure Name: ByteToStr
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter bArray (Byte):
' Return Type: String
' Author: beededea
' Date: 30/05/2024
' ----------------------------------------------------------------
Private Function ByteToStr(bArray() As Byte) As String
    On Error GoTo ByteToStr_Error
    Dim lPntr As Long
    Dim bTmp() As Byte
    On Error GoTo ByteErr
    ReDim bTmp(UBound(bArray) * 2 + 1)
    For lPntr = 0 To UBound(bArray)
        bTmp(lPntr * 2) = bArray(lPntr)
    Next lPntr
    Let ByteToStr = bTmp
    Exit Function
ByteErr:
    ByteToStr = ""
    
    On Error GoTo 0
    Exit Function

ByteToStr_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ByteToStr, line " & Erl & "."

End Function



' ----------------------------------------------------------------
' Procedure Name: ScanWireless
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter thisArray (String):
' Parameter thisWirelessPercentArray (Integer):
' Parameter thisWirelessRSSIArray (Integer):
' Parameter lCount (Integer):
' Author: beededea
' Date: 30/05/2024
' ----------------------------------------------------------------
Public Sub ScanWireless(ByRef thisArray() As String, ByRef thisWirelessPercentArray() As Integer, ByRef thisWirelessRSSIArray() As Integer, ByRef lCount As Integer)

    Dim lRet As Long: lRet = 0
    Dim lList As Long: lList = 0
    Dim lAvailable As Long: lAvailable = 0
    Dim lStart As Long: lStart = 0
    Dim sLen As Long: sLen = 0
    Dim lBSS As Long: lBSS = 0
    Dim sSSID As String: sSSID = vbNullString
    Dim lPtr As Long: lPtr = 0
    Dim dwFlags As Long: dwFlags = 0
    Dim dwreserved As Long: dwreserved = 0
    Dim quality As Integer: quality = 0
    Dim dbm As Integer: dbm = 0
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    Dim XMLBuffer(1023) As Byte
    Dim bSSID() As Byte
    Dim udtAvailableList As WLAN_AVAILABLE_NETWORK_LIST
    Dim udtNetwork As WLAN_AVAILABLE_NETWORK
    Dim Network As AVAILABLE_NETWORK
    
    Dim strDBG As String: strDBG = vbNullString


    On Error GoTo ScanWireless_Error
    
    ConIndex = -1

    ReDim bBuffer(0)
    
    ' requests a scan for available networks on the indicated interface.
    If lHandle Then
        strDBG = "lhandlescan"
    
        lRet = WlanScan(lHandle, udtList.InterfaceInfo.ifGuid, ByVal 0&, ByVal 0&, ByVal 0&)
        Screen.MousePointer = vbHourglass
        'Wait for scan to finish, don't like this
        Sleep 500
        Screen.MousePointer = vbDefault
    Else
    
        strDBG = "scan"
        
        'Get adapter handle and find WLAN interfaces
        lRet = WlanOpenHandle(2&, 0&, lVersion, lHandle)
        
        'NOTE: This code currently only processes the first wireless adapter
        
        On Error GoTo l_trap_error2
        
        lRet = WlanEnumInterfaces(ByVal lHandle, 0&, lList)
        
        On Error GoTo ScanWireless_Error
        
        CopyMemory udtList, ByVal lList, Len(udtList)
        Debug.Print udtList.dwNumberofItems, "WiFi Adapter found!"
    End If
    
    strDBG = "udtList"

    If udtList.dwNumberofItems > 0 Then
    
        strDBG = "udtList+"
        
        lRet = WlanGetAvailableNetworkList(lHandle, udtList.InterfaceInfo.ifGuid, 2&, 0&, lAvailable)
        
        If lRet <> 0 Then GoTo l_trap_error2 '  -2144067582 - error
        
        CopyMemory udtAvailableList, ByVal lAvailable, LenB(udtAvailableList)
        
        On Error GoTo ScanWireless_Error

        lCount = 0
        lStart = lAvailable + 8
        
        strDBG = "bBuffer"
        
        If udtAvailableList.dwNumberofItems > 0 Then

            ReDim bBuffer(Len(Network) * udtAvailableList.dwNumberofItems - 1)
            Do 'Create new abbreviated buffer
            
                strDBG = "CopyMemory"
    
                CopyMemory udtNetwork, ByVal lStart, Len(udtNetwork)
                lCount = lCount + 1
                lStart = lStart + Len(udtNetwork)
                CopyMemory bBuffer(lPtr), udtNetwork.dot11Ssid.uSSIDLength, Len(Network)
                lPtr = lPtr + Len(Network)
            Loop Until lCount = udtAvailableList.dwNumberofItems
            
            strDBG = "WlanFreeMemory"
    
            WlanFreeMemory lAvailable
            WlanFreeMemory lList
            'Create new list from new buffer
            lStart = VarPtr(bBuffer(0))
            lCount = 0
            
            strDBG = "ReDim3"
            
            ReDim thisArray(udtAvailableList.dwNumberofItems) As String
            ReDim thisWirelessPercentArray(udtAvailableList.dwNumberofItems) As Integer
            ReDim thisWirelessRSSIArray(udtAvailableList.dwNumberofItems) As Integer
            
            Do
            
                strDBG = "Do"
            
                CopyMemory Network, ByVal lStart, Len(Network)
                sLen = Network.dot11Ssid.uSSIDLength
                If sLen = 0 Then
                
                    strDBG = "0"
                
                    sSSID = "(Unknown)"
                Else
                    strDBG = "else"
                    
                    ReDim bSSID(sLen - 1)
                    CopyMemory bSSID(0), Network.dot11Ssid.ucSSID(0), sLen
                    sSSID = ByteToStr(bSSID)
                End If
                Debug.Print "SSID "; sSSID, "Signal "; Network.wlanSignalQuality
                
                strDBG = "sSSID"
    
                sSSID = Left$(sSSID & Space$(25), 25)
    
                thisArray(lCount) = sSSID
                quality = Network.wlanSignalQuality
                thisWirelessPercentArray(lCount) = quality
                
                strDBG = "quality"
                
                If (quality <= 0) Then
                    dbm = -100
                ElseIf quality >= 100 Then
                    dbm = -50
                Else
                    dbm = (quality / 2) - 100
                End If
                
                strDBG = "thisWirelessRSSIArray"
                
                thisWirelessRSSIArray(lCount) = dbm
                
                If (Network.dwFlags And 1) = 1 Then
                    ConIndex = lCount
                End If
                
                strDBG = "lCount"
    
                lCount = lCount + 1
                lStart = lStart + Len(Network)
            Loop Until lCount = udtAvailableList.dwNumberofItems
'        Else
'            MsgBox "No Wireless Adapters Found"
        End If
    End If
    
    If ConIndex > -1 Then 'Display connected network
    
        strDBG = "Connected"

        Connected = Trim(Left$(thisArray(ConIndex), 25))
        overlayWidget.thisWirelessNo = ConIndex
        Debug.Print sSSID
        Debug.Print quality
        Debug.Print dbm
    End If
    
    On Error GoTo 0
    Exit Sub
    
l_trap_error2:
    
    answer = vbYes
    answerMsg = "ERROR - Wireless Adapter unavailable or disabled. Handled within subroutine ScanWireless by event l_trap_error2"
    If frmMessage.IsVisible = False Then answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Wireless Adapter Error", True, "scanwireless2")
    
    lCount = 0
    Exit Sub

ScanWireless_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ScanWireless, line " & Erl & "." & " error point = " & strDBG
    
    If strDBG = "bBuffer" Then
        MsgBox "Len(Network) " & Len(Network) & " udtAvailableList.dwNumberofItems " & udtAvailableList.dwNumberofItems - 1
    End If

End Sub







