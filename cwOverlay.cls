VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlay
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfGauge
'---------------------------------------------------------------------------------------

Option Explicit

Private wirelessPercent As Long
Private wirelessRSSI As Long

Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

'these Public Properties are set in the Parent (here, cfGauge)
Public cCollOverlayImages As cCollection
Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public WithEvents tmrSampler As cTimer
Attribute tmrSampler.VB_VarHelpID = -1
Public WithEvents tmrAnimator As cTimer
Attribute tmrAnimator.VB_VarHelpID = -1

Private oldSDeg As Double
Private newSDeg As Double
Private animationQuanta As Double
Private animationAmount As Double
Private animationDivisions As Integer
Private animationInterval As Integer
Private animateCCW As Boolean

Private SDeg As Double

Private mTicking As Boolean             ' property Ticking
Private mPointerAnimate As Boolean    ' property PointerAnimate
Private mLocked As Boolean              ' property Locked
Private mHidden As Boolean              ' property Hidden
Private mThisWireless As String            ' property thisWireless
Private mThisWirelessNo As Integer          ' property thisWireless

Private mThisOpacity As String          ' property ThisOpacity
Private mSamplingInterval As String     ' property SamplingInterval




'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   : initialise vars
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()

   On Error GoTo initialisePrivateVars_Error


   ' pvtCurrentDateToDisplay = vbNullString

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize for cwOverlay
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Constructor (Initialise) set the inital values of private vars, set up the timers
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Dim intervalValue As Integer: intervalValue = 0
    
    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars

    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    Set tmrSampler = Nothing
    Set tmrSampler = New_c.Timer(100, mTicking) ' first interval is 1/10th of a second, don't change it, triggers an immediate run. Correct sampling interval read later
    
    animationDivisions = 12
    animationInterval = 40
    
    Set tmrAnimator = Nothing
    Set tmrAnimator = New_c.Timer(animationInterval, False)
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : Ticking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Ticking() As Boolean
   On Error GoTo TickingGet_Error

  Ticking = mTicking

   On Error GoTo 0
   Exit Property

TickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Ticking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Ticking(ByVal newValue As Boolean)
    'Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo TickingLet_Error

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    
    If mTicking Then
'        If mPointerAnimate Then
'            'intervalValue = 100
'        Else
'            'intervalValue = 500
'        End If

        tmrSampler.Enabled = True
    Else
        tmrSampler.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : PointerAnimate
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get PointerAnimate() As Boolean
   On Error GoTo PointerAnimateGet_Error

    PointerAnimate = mPointerAnimate

   On Error GoTo 0
   Exit Property

PointerAnimateGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property PointerAnimate of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : PointerAnimate
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let PointerAnimate(ByVal newValue As Boolean)
    Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo PointerAnimateLet_Error

    If mPointerAnimate <> newValue Then mPointerAnimate = newValue Else Exit Property
    
   On Error GoTo 0
   Exit Property

PointerAnimateLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property PointerAnimate of Class Module cwOverlay"
End Property ' the earth PointerAnimate flag
'---------------------------------------------------------------------------------------
' Property : thisOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisOpacity() As String
   On Error GoTo thisOpacityGet_Error

    thisOpacity = mThisOpacity

   On Error GoTo 0
   Exit Property

thisOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisOpacity
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisOpacity(ByVal newValue As String)
    On Error GoTo thisOpacityLet_Error

    If mThisOpacity <> newValue Then mThisOpacity = newValue Else Exit Property
        
    On Error GoTo 0
    Exit Property

thisOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : thisWireless
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisWireless() As String
   On Error GoTo thisWirelessGet_Error

    thisWireless = mThisWireless

   On Error GoTo 0
   Exit Property

thisWirelessGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWireless of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisWireless
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisWireless(ByVal newValue As String)
    On Error GoTo thisWirelessLet_Error

    If mThisWireless <> newValue Then mThisWireless = newValue Else Exit Property
    
    Call obtainwirelesss
    
    On Error GoTo 0
    Exit Property

thisWirelessLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWireless of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : thisWireless
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisWirelessNo() As Integer
   On Error GoTo thisWirelessNoGet_Error

    thisWirelessNo = mThisWirelessNo

   On Error GoTo 0
   Exit Property

thisWirelessNoGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWirelessNo of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisWirelessNo
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisWirelessNo(ByVal newValue As Integer)
    On Error GoTo thisWirelessNoLet_Error

    If mThisWirelessNo <> newValue Then mThisWirelessNo = newValue Else Exit Property
            
    'given a Wireless number select the Wireless string
    If mThisWirelessNo < gblWirelessCount Then
        mThisWireless = gblWirelessSSIDArray(mThisWirelessNo)
        If mThisWirelessNo = 10 Then
            mThisWirelessNo = 10
        End If
    End If
    
    Call obtainwirelesss
    
    On Error GoTo 0
    Exit Property

thisWirelessNoLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWirelessNo of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : samplingInterval
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get samplingInterval() As String
   On Error GoTo samplingIntervalGet_Error

    samplingInterval = mSamplingInterval

   On Error GoTo 0
   Exit Property

samplingIntervalGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : samplingInterval
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let samplingInterval(ByVal newValue As String)
    On Error GoTo samplingIntervalLet_Error

    If mSamplingInterval <> newValue Then mSamplingInterval = newValue Else Exit Property
        
    If tmrSampler.Interval = 100 Then Call tmrSampler_Timer ' allows first early run
    
    tmrSampler.Interval = Val(mSamplingInterval) * 1000
    If mPointerAnimate = True Then
        If tmrSampler.Interval = 1000 Then tmrSampler.Interval = 1200 ' take into account the animator
    End If
    
    On Error GoTo 0
    Exit Property

samplingIntervalLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property



'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property used to store/set the lock state to lock the widget in place, the actual locking is done elsewhere
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwoverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property used to store/set the lock state to lock the widget in place, the actual locking is done elsewhere
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwoverlay"
End Property




''---------------------------------------------------------------------------------------
'' Procedure : MyOpacity
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   : property (by value) used to determine the opacity of elements in the overlay
''---------------------------------------------------------------------------------------
''
'Public Property Let MyOpacity(ByVal newValue As Double)
'   On Error GoTo MyOpacityLet_Error
'
'   If mMyOpacity <> newValue Then mMyOpacity = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'MyOpacityLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
'End Property
'
''---------------------------------------------------------------------------------------
'' Procedure : MyOpacity
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   : property (by value) used to determine the opacity of elements in the overlay
''---------------------------------------------------------------------------------------
''
'Public Property Get MyOpacity() As Double
'   On Error GoTo MyOpacityGet_Error
'
'   MyOpacity = mMyOpacity
'
'   On Error GoTo 0
'   Exit Property
'
'MyOpacityGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
'End Property



'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value)  the set the hidden value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

   Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwoverlay"
End Property

    


'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   : property to determine (by value)  the set the hidden value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        W.Visible = False
        
        W.Refresh
    Else
        W.Visible = True
        
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwoverlay"
End Property


'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Overlay-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtWirelessPercent As String: txtWirelessPercent = vbNullString
    Dim txtWirelessRSSI As String: txtWirelessRSSI = vbNullString
    Dim textDrawn As Boolean:   textDrawn = False
    Dim someOpacity As Double: someOpacity = 0
    Dim cream As String: cream = vbNullString
    Dim leftmost As Integer: leftmost = 0
    Dim leftmostMax As Integer: leftmostMax = 0
    Dim leftmostCurr As Integer: leftmostCurr = 0
    
    Dim scaleString As String: scaleString = vbNullString
        
    cream = RGB(235, 217, 179) ' cannot create a constant from an RGB value
   
    On Error GoTo W_Paint_Error
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 18
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    someOpacity = (Val(mThisOpacity) / 100) * 0.75
    
    txtWirelessPercent = CStr(Int(wirelessPercent))
    txtWirelessRSSI = wirelessRSSI

    For I = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfGauge by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(I)
      Key = cCollOverlayImages.KeyByIndex(I)
      
      If IsArray(Arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case Key
                 
                 Case "redlamptrue"
                    
                    If wirelessPercent < 80 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                      
                 Case "greenlamptrue"
                    
                    If wirelessPercent >= 80 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "secondshadow", "secondhand"
                                        
                    'set some Font and ForeColor Properties and draw some Text on the Widget
                    W.FontName = "arial"
                    W.ForeColor = cream
                    W.FontSize = 16
                    W.FontBold = True
                    W.SelectFontSettingsInto CC
                    W.Refresh
                    
                    leftmostMax = 265
                    leftmostCurr = 455
                    
                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText leftmostMax, 103, 200, 200, txtWirelessPercent, , , 2, True, , someOpacity
                    CC.DrawText leftmostCurr, 300, 200, 200, txtWirelessRSSI, , , 2, True, , someOpacity
                    textDrawn = True
                    
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    CC.RotateDrawingsDeg SDeg
                    CC.Paint Arr(2), Pat 'render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
                 
                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                    CC.Paint Arr(2), Pat 'render the Pattern
                   
             End Select
             
             CC.Restore
      End If
     Next

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 
 
'---------------------------------------------------------------------------------------
' Procedure : tmrAnimator_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : animate timer, once every few (40) milliseconds
'---------------------------------------------------------------------------------------
'
Private Sub tmrAnimator_Timer()
    
    On Error GoTo tmrAnimator_Timer_Error
    '
    ' move the pointer using SDeg
    '
    animationAmount = animationAmount + animationQuanta

    If animateCCW = True Then
        If SDeg > newSDeg Then
            SDeg = oldSDeg - animationAmount
        Else
            tmrAnimator.Enabled = False ' when the animation is complete then stop the animator
            animationAmount = 0
            Exit Sub
        End If
    Else
        If SDeg < newSDeg Then
            SDeg = oldSDeg + animationAmount
        Else
            tmrAnimator.Enabled = False
            animationAmount = 0
            Exit Sub
        End If
    End If

    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

tmrAnimator_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimator_Timer of Class Module cwOverlay"
End Sub

  

    


'---------------------------------------------------------------------------------------
' Procedure : tmrSampler_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrSampler_Timer()
    
    Dim degDiff As Double: degDiff = 0
    
    On Error GoTo tmrSampler_Timer_Error
    
   If mPointerAnimate = True Then
        If tmrAnimator.Enabled = False Then
        
            ' oldSDeg stored
            oldSDeg = SDeg
            
            If gblWirelessCount = 0 Then Exit Sub
            
            ' get the list of Wireless and the count
            Call getgblWirelessArray(gblWirelessSSIDArray(), gblWirelessPercentArray(), gblWirelessRSSIArray(), gblWirelessCount)
                    
            If mThisWirelessNo <> "9999" Then
                wirelessPercent = fWireless_Usage_Percent
                wirelessRSSI = fWireless_RSSI
            Else
                wirelessPercent = 0
                wirelessRSSI = 0
            End If
                       
            ' calculate the new signal strength amount in degrees
            newSDeg = (wirelessPercent * 3) + 30
            
            If oldSDeg = newSDeg Then Exit Sub
            
            ' calculate the difference
            degDiff = newSDeg - oldSDeg
            
            ' if the result is negative animate CCW
            If degDiff < 0 Then
                animateCCW = True
            Else
                animateCCW = False
            End If
            
            Select Case degDiff
                Case Is <= 5
                    animationDivisions = 4
                Case Is <= 10
                    animationDivisions = 8
                Case Is > 10
                    animationDivisions = 16
                Case Is > 20
                    animationDivisions = 24
            End Select
            
            ' divide that by 24 (one second duration)
            animationQuanta = Abs(degDiff / animationDivisions)
            
            tmrAnimator.Enabled = True
            
        End If
    Else
        Call obtainwirelesss
    End If
    
    W.Parent.Refresh

    tmrSampler.Enabled = True
            
   On Error GoTo 0
   Exit Sub

tmrSampler_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampler_Timer of Class Module cwOverlay"
End Sub






' ----------------------------------------------------------------
' Procedure Name: obtainwirelesss
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 14/01/2024
' ----------------------------------------------------------------
Private Sub obtainwirelesss()
    On Error GoTo obtainwirelesss_Error

        ' get the list of Wirelesss and the count
        If gblWirelessCount > 0 Then Call getgblWirelessArray(gblWirelessSSIDArray(), gblWirelessPercentArray(), gblWirelessRSSIArray(), gblWirelessCount)
        
        'Debug.Print gblWirelessSSIDArray(0)
        
        If mThisWirelessNo <> "9999" And gblWirelessCount > 0 Then
            wirelessPercent = fWireless_Usage_Percent
            wirelessRSSI = fWireless_RSSI
        Else
            wirelessPercent = 0
            wirelessRSSI = 0
        End If
                
        SDeg = (wirelessPercent * 3) + 30
                
        W.Refresh
    
    On Error GoTo 0
    Exit Sub

obtainwirelesss_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainwirelesss, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: fWireless_Usage_Percent
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Return Type: Long
' Author: https://www.tek-tips.com/userinfo.cfm?member=Hypetia
' Date: 09/01/2024
' ----------------------------------------------------------------
Private Function fWireless_Usage_Percent() As Long
    Dim newT As Integer
    On Error GoTo fWireless_Usage_Percent_Error
    
    If SafeArrayGetDim(gblWirelessPercentArray) <> 0 Then
        newT = gblWirelessPercentArray(overlayWidget.thisWirelessNo)
    Else
        newT = 0
    End If

    fWireless_Usage_Percent = newT
     
    On Error GoTo 0
    Exit Function

fWireless_Usage_Percent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fWireless_Usage_Percent, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: fWireless_RSSI
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Return Type: Long
' Author:
' Date: 09/01/2024
' ----------------------------------------------------------------
Private Function fWireless_RSSI() As Long
    Dim newT As Integer
    
    On Error GoTo fWireless_RSSI_Error
    
    If SafeArrayGetDim(gblWirelessRSSIArray) <> 0 Then
        newT = gblWirelessRSSIArray(overlayWidget.thisWirelessNo)
    Else
        newT = 0
    End If

    fWireless_RSSI = newT
     
    On Error GoTo 0
    Exit Function

fWireless_RSSI_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fWireless_RSSI, line " & Erl & "."
 
End Function


