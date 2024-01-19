Attribute VB_Name = "modWireless"
Option Explicit

Private Const WLAN_NOTIFICATION_SOURCE_MOST As Long = &H7F

Private Enum DOT11_PHY_TYPE
    dot11_phy_type_unknown = 0
    dot11_phy_type_any = 0
    dot11_phy_type_fhss = 1
    dot11_phy_type_dsss = 2
    dot11_phy_type_irbaseband = 3
    dot11_phy_type_ofdm = 4
    dot11_phy_type_hrdsss = 5
    dot11_phy_type_erp = 6
    dot11_phy_type_ht = 7
    dot11_phy_type_IHV_start = &H80000000
    dot11_phy_type_IHV_end = &HFFFFFFFF
End Enum

Private Enum DOT11_BSS_TYPE
    dot11_BSS_type_infrastructure = 1
    dot11_BSS_type_independent = 2
    dot11_BSS_type_any = 3
End Enum

Private Enum DOT11_AUTH_ALGORITHM
    DOT11_AUTH_ALGO_80211_OPEN = 1
    DOT11_AUTH_ALGO_80211_SHARED_KEY = 2
    DOT11_AUTH_ALGO_WPA = 3
    DOT11_AUTH_ALGO_WPA_PSK = 4
    DOT11_AUTH_ALGO_WPA_NONE = 5
    DOT11_AUTH_ALGO_RSNA = 6
    DOT11_AUTH_ALGO_RSNA_PSK = 7
    DOT11_AUTH_ALGO_IHV_START = &H80000000
    DOT11_AUTH_ALGO_IHV_END = &HFFFFFFFF
End Enum

Private Enum DOT11_CIPHER_ALGORITHM
    DOT11_CIPHER_ALGO_NONE = &H0
    DOT11_CIPHER_ALGO_WEP40 = &H1
    DOT11_CIPHER_ALGO_TKIP = &H2
    DOT11_CIPHER_ALGO_CCMP = &H4
    DOT11_CIPHER_ALGO_WEP104 = &H5
    DOT11_CIPHER_ALGO_WPA_USE_GROUP = &H100
    DOT11_CIPHER_ALGO_RSN_USE_GROUP = &H100
    DOT11_CIPHER_ALGO_WEP = &H101
    DOT11_CIPHER_ALGO_IHV_START = &H80000000
    DOT11_CIPHER_ALGO_IHV_END = &HFFFFFFFF
End Enum

Private Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Type WLAN_INTERFACE_INFO
    ifGuid As GUID
    InterfaceDescription(255) As Byte
    IsState As Long
End Type

Private Type DOT11_SSID
    uSSIDLength As Long
    ucSSID(31) As Byte
End Type

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

Private Type WLAN_CONNECTION_PARAMETERS
    ConnectionMode As Long
    Profile As Long
    pDot11Ssid As Long
    pDesiredBssidList As Long
    dot11BssType As Long
    dwFlags As Long
End Type

Private Type WLAN_BSS_LIST
    dwTotalSize As Long
    dwNumberofItems As Long
    wlanBssEntries As Long
End Type

Private Declare Function WlanOpenHandle Lib "wlanapi.dll" (ByVal dwClientVersion As Long, ByVal pdwReserved As Long, ByRef pdwNegotiaitedVersion As Long, ByRef phClientHandle As Long) As Long
Private Declare Function WlanEnumInterfaces Lib "wlanapi.dll" (ByVal hClientHandle As Long, ByVal pReserved As Long, ppInterfaceList As Long) As Long
Private Declare Function WlanGetAvailableNetworkList Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, ByVal dwFlags As Long, ByVal pReserved As Long, ppAvailableNetworkList As Long) As Long
Private Declare Function WlanConnect Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, pConnectionParameters As WLAN_CONNECTION_PARAMETERS, ByVal reserved As Long) As Long
Private Declare Function WlanScan Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, pDot11Ssid As Long, pIeData As Long, reserved As Long) As Long
Private Declare Function WlanDisconnect Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, ByVal pReserved As Long) As Long
Private Declare Function WlanGetNetworkBssList Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGui As GUID, ByVal pDot11Ssid As Long, ByVal dot11BssType As Long, ByVal bSecurityEnabled As Long, ByVal pReserved As Long, ppWlanBssList As WLAN_BSS_LIST) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub WlanFreeMemory Lib "wlanapi.dll" (ByVal pMemory As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private udtList As WLAN_INTERFACE_INFO_LIST
Private ConIndex As Long
Private lHandle As Long
Private lVersion As Long
Private Connected As String
Private bBuffer() As Byte
Private Function ByteToStr(bArray() As Byte) As String
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
End Function

Private Sub DebugPrintByte(sDescr As String, bArray() As Byte)
    Dim lPtr As Long
    Debug.Print sDescr & ":"
    If GetbSize(bArray) = 0 Then Exit Sub
    For lPtr = 0 To UBound(bArray)
        Debug.Print Right$("0" & Hex$(bArray(lPtr)), 2) & " ";
        If (lPtr + 1) Mod 16 = 0 Then Debug.Print
    Next lPtr
    Debug.Print
End Sub

Private Function GetbSize(bArray() As Byte) As Long
    On Error GoTo GetSizeErr
    GetbSize = UBound(bArray) + 1
    Exit Function
GetSizeErr:
    GetbSize = 0
End Function

Private Sub GetInfo(ByVal Index As Long)
    Dim Network As AVAILABLE_NETWORK
    Dim lStart As Long
    Dim sLen As Long
    Dim bSSID() As Byte
    Dim sSSID As String
    Dim lPtr As Long
    Dim Msg As String
    lStart = VarPtr(bBuffer(0)) + Index * Len(Network)
    CopyMemory Network, ByVal lStart, Len(Network)
    sLen = Network.dot11Ssid.uSSIDLength
    If sLen = 0 Then
        sSSID = "(Unknown)"
    Else
        ReDim bSSID(sLen - 1)
        CopyMemory bSSID(0), Network.dot11Ssid.ucSSID(0), sLen
        sSSID = ByteToStr(bSSID)
    End If
    Msg = "Signal Strength: " & CStr(Network.wlanSignalQuality)
    If Network.dwFlags And 1 Then
        Msg = Msg & vbCrLf & "Connected"
    Else
        Msg = Msg & vbCrLf & "Not Connected"
    End If
    Select Case Network.dot11BssType
        Case DOT11_BSS_TYPE.dot11_BSS_type_infrastructure
            Msg = Msg & vbCrLf & "BSS: Infrastructure"
        Case DOT11_BSS_TYPE.dot11_BSS_type_independent
            Msg = Msg & vbCrLf & "BSS: Peer to Peer"
    End Select
    If Network.bNetworkConnectable <> 0 Then
        Msg = Msg & vbCrLf & "Connectable"
    Else
        Msg = Msg & vbCrLf & "Not Connectable"
    End If
    For lPtr = 0 To UBound(Network.dot11PhyTypes)
        Select Case Network.dot11PhyTypes(lPtr)
            Case DOT11_PHY_TYPE.dot11_phy_type_ht
                Msg = Msg & vbCrLf & "802.11n"
            Case DOT11_PHY_TYPE.dot11_phy_type_erp
                Msg = Msg & vbCrLf & "802.11g"
            Case DOT11_PHY_TYPE.dot11_phy_type_ofdm
                Msg = Msg & vbCrLf & "802.11a"
        End Select
    Next lPtr
    If Network.bSecurityEnabled Then Msg = Msg & vbCrLf & "Security Enabled"
    Select Case Network.dot11DefaultAuthAlgorithm
        Case DOT11_AUTH_ALGORITHM.DOT11_AUTH_ALGO_80211_OPEN
            Msg = Msg & vbCrLf & "Auth Algorithm: Open"
        Case DOT11_AUTH_ALGORITHM.DOT11_AUTH_ALGO_80211_SHARED_KEY
            Msg = Msg & vbCrLf & "Auth Algorithm: Shared Key"
        Case DOT11_AUTH_ALGORITHM.DOT11_AUTH_ALGO_WPA
            Msg = Msg & vbCrLf & "Auth Algorithm: WPA"
        Case DOT11_AUTH_ALGORITHM.DOT11_AUTH_ALGO_RSNA
            Msg = Msg & vbCrLf & "Auth Algorithm: RSNA"
        Case DOT11_AUTH_ALGORITHM.DOT11_AUTH_ALGO_RSNA_PSK
            Msg = Msg & vbCrLf & "Auth Algorithm: RSNA with Pre-shared Keys"
        Case DOT11_AUTH_ALGORITHM.DOT11_AUTH_ALGO_WPA_PSK
            Msg = Msg & vbCrLf & "Auth Algorithm: WPA with Pre-shared Keys"
        Case DOT11_AUTH_ALGORITHM.DOT11_AUTH_ALGO_80211_SHARED_KEY
            Msg = Msg & vbCrLf & "Auth Algorithm: WEP"
    End Select
    Select Case Network.dot11DefaultCipherAlgorithm
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_CCMP
            Msg = Msg & vbCrLf & "Cypher Algorithm: AES - CCMP"
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_NONE
            Msg = Msg & vbCrLf & "Cypher Algorithm: None"
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_RSN_USE_GROUP
            Msg = Msg & vbCrLf & "Cypher Algorithm: RSN - Use Group Key"
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_TKIP
            Msg = Msg & vbCrLf & "Cypher Algorithm: TKIP"
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_WEP
            Msg = Msg & vbCrLf & "Cypher Algorithm: WEP"
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_WEP104
            Msg = Msg & vbCrLf & "Cypher Algorithm: WEP - 104 Bit Key"
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_WEP40
            Msg = Msg & vbCrLf & "Cypher Algorithm: WEP - 40 Bit Key"
        Case DOT11_CIPHER_ALGORITHM.DOT11_CIPHER_ALGO_WPA_USE_GROUP
            Msg = Msg & vbCrLf & "Cypher Algorithm: WPA - Use Group Key"
    End Select
    MsgBox Msg, , sSSID
End Sub


Public Sub ScanWireless(ByRef thisArray() As String, ByRef thisWirelessPercentArray() As Integer, ByRef lCount As Integer)
    Dim lRet As Long
    Dim lList As Long
    Dim lAvailable As Long
    Dim lStart As Long
    'Dim lCount As Long
    Dim sLen As Long
    Dim bSSID() As Byte
    Dim sSSID As String
    Dim udtAvailableList As WLAN_AVAILABLE_NETWORK_LIST
    Dim udtNetwork As WLAN_AVAILABLE_NETWORK
    Dim Network As AVAILABLE_NETWORK
    Dim lPtr As Long
    
    ConIndex = -1
    ReDim bBuffer(0)
    If lHandle Then
        lRet = WlanScan(lHandle, udtList.InterfaceInfo.ifGuid, ByVal 0&, ByVal 0&, ByVal 0&)
        Screen.MousePointer = vbHourglass
        'Wait for scan to finish (4 seconds)
        'Sleep 4000
        Screen.MousePointer = vbDefault
    Else 'Get adapter handle and find WLAN interfaces
        lRet = WlanOpenHandle(2&, 0&, lVersion, lHandle)
        'NOTE: This code currently only processes the first wireless adapter
        lRet = WlanEnumInterfaces(ByVal lHandle, 0&, lList)
        CopyMemory udtList, ByVal lList, Len(udtList)
        Debug.Print udtList.dwNumberofItems, "WiFi Adapter found!"
    End If
    If udtList.dwNumberofItems > 0 Then
        lRet = WlanGetAvailableNetworkList(lHandle, udtList.InterfaceInfo.ifGuid, 2&, 0&, lAvailable)
        CopyMemory udtAvailableList, ByVal lAvailable, LenB(udtAvailableList)
        lCount = 0
        lStart = lAvailable + 8
        'lblStatus.Caption = CStr(udtAvailableList.dwNumberofItems) & " Networks Found!"
        ReDim bBuffer(Len(Network) * udtAvailableList.dwNumberofItems - 1)
        Do 'Create new abbreviated buffer
            CopyMemory udtNetwork, ByVal lStart, Len(udtNetwork)
            lCount = lCount + 1
            lStart = lStart + Len(udtNetwork)
            CopyMemory bBuffer(lPtr), udtNetwork.dot11Ssid.uSSIDLength, Len(Network)
            lPtr = lPtr + Len(Network)
        Loop Until lCount = udtAvailableList.dwNumberofItems
        WlanFreeMemory lAvailable
        WlanFreeMemory lList
        'Create new list from new buffer
        lStart = VarPtr(bBuffer(0))
        lCount = 0
        
        ReDim thisArray(udtAvailableList.dwNumberofItems) As String
        ReDim thisWirelessPercentArray(udtAvailableList.dwNumberofItems) As Integer
        Do
            CopyMemory Network, ByVal lStart, Len(Network)
            sLen = Network.dot11Ssid.uSSIDLength
            If sLen = 0 Then
                sSSID = "(Unknown)"
            Else
                ReDim bSSID(sLen - 1)
                CopyMemory bSSID(0), Network.dot11Ssid.ucSSID(0), sLen
                sSSID = ByteToStr(bSSID)
            End If
            Debug.Print "SSID "; sSSID, "Signal "; Network.wlanSignalQuality
            thisArray(lCount) = sSSID
            thisWirelessPercentArray(lCount) = Network.wlanSignalQuality
            
            If (Network.dwFlags And 1) = 1 Then
                ConIndex = lCount
            End If
            lCount = lCount + 1
            lStart = lStart + Len(Network)
        Loop Until lCount = udtAvailableList.dwNumberofItems
    Else
        MsgBox "No Wireless Adapters Found"
    End If

    If ConIndex > -1 Then 'Display connected network
        Connected = Trim(Left$(thisArray(ConIndex), 25))
    End If
End Sub



'Private Sub 'lstSSID_Click()
'    Call GetInfo('lstSSID.ListIndex)
'End Sub



