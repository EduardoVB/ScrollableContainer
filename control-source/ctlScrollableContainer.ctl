VERSION 5.00
Begin VB.UserControl ScrollableContainer 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   PropertyPages   =   "ctlScrollableContainer.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlScrollableContainer.ctx":002D
   Begin VB.Timer tmrVScrollValue 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2448
      Top             =   3072
   End
   Begin VB.Timer tmrHScrollInit 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2060
      Top             =   3072
   End
   Begin VB.Timer tmrVScrollInit 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   3072
   End
   Begin VB.Timer tmrCheckFocus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1296
      Top             =   3072
   End
   Begin VB.Timer tmrMoveLeft 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   504
      Top             =   3072
   End
   Begin VB.Timer tmrDontIncreaseMax 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   900
      Top             =   3072
   End
   Begin VB.Timer tmrMoveTop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   108
      Top             =   3060
   End
End
Attribute VB_Name = "ScrollableContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBSSubclass

Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WM_SETREDRAW As Long = &HB&
' Redraw window:
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_INVALIDATE = &H1
Private Const RDW_UPDATENOW = &H100

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long

Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOACTIVATE = &H10&
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20

Private Const BF_Left = &H1
Private Const BF_TOP = &H2
Private Const BF_Right = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_Left Or BF_TOP Or BF_Right Or BF_BOTTOM)

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const WM_NCPAINT = &H85

Private Type BORDERSTYLE_DATA
    Flags As Long
    Width As Long
End Type

Private WithEvents mScroll As cScrollBars
Attribute mScroll.VB_VarHelpID = -1

' Public enums
Public Enum vbExScrollBarVisibilityConstants
    vxScrollBarHide = 0
    vxScrollBarShow = 1
    vxScrollBarAuto = 2
End Enum

Public Enum vbExExtendedBorderStyleConstants
    vxEBSNone = 0
    vxEBSFlat1Pix = 1
    vxEBSFlat2Pix = 2
    vxEBSSunken2Pix = 3
    vxEBSRaised2Pix = 4
    vxEBSEtched2Pix = 5
    vxEBSSunkenOuter1Pix = 6
    vxEBSSunkenInner1Pix = 7
    vxEBSRaisedOuter1Pix = 8
    vxEBSRaisedInner1Pix = 9
End Enum

Public Event VScrollChange()
Attribute VScrollChange.VB_Description = "Generated when VScrollValue changes."
Attribute VScrollChange.VB_MemberFlags = "200"
Public Event HScrollChange()
Attribute HScrollChange.VB_Description = "Generated when HScrollValue changes."


' Persistable properties
Private mBackColor As Long
Private mBorderStyle As vbExExtendedBorderStyleConstants
Private mBorderColor As Long
Private mBottomFreeSpace As Single
Private mRightFreeSpace As Single
Private mVScrollBar As vbExScrollBarVisibilityConstants
Private mHScrollBar As vbExScrollBarVisibilityConstants
Private mAutoScrollOnFocus As Boolean

' Non persistable properties
Private mVirtualHeight As Single
Private mVScrollValue As Single
Private mVirtualWidth As Single
Private mHScrollValue As Single
Private mAddingControls As Boolean

' Variables for vertical handling
Private mMoveTop As Single
Private mTempVScrollValue As Long
Private mTempVScrollMax As Long
Private mTempVirtualHeight As Long
' Variables for horizontal handling
Private mMoveLeft As Single
Private mTempHScrollValue As Long
Private mTempHScrollMax As Long
Private mTempVirtualWidth As Long
' Other variables
Private mNoScroll As Boolean
Private mAddingControls_v As Single
Private mAddingControls_h As Single
Private mFocusHwndList As Collection
Private mUserControlHwnd As Long
Private mUpdating As Boolean

Private Const cDefaultBorderColor As Long = vbWindowFrame
Private Const cDefaultBorderStyle = vxEBSFlat1Pix
Private Const cDefaultBottomFreeSpace As Long = 300 ' twips
Private mScrollBarHeight As Long
Private mScrollBarWidth As Long
Private mUserMode As Boolean
Private mShown As Boolean
Private mTopScrollBound As Single
Private mSubclassed As Boolean

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    Unsubclass
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    If iMsg = WM_NCPAINT Then
        Dim iWindowRect As RECT
        Dim iDC As Long
        Dim iBrush As Long
        Dim iRc As RECT
        Dim iColor As Long
        Dim iBs As BORDERSTYLE_DATA
        
        If (mBorderStyle = vxEBSNone) Or (mBorderStyle = vxEBSFlat1Pix) And (mBorderColor = vbWindowFrame) Or (mBorderStyle = vxEBSSunken2Pix) Then
            IBSSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        ElseIf (mBorderStyle = vxEBSFlat1Pix) Or (mBorderStyle = vxEBSFlat2Pix) Then
            iBs = GetBorderStyleData(mBorderStyle)
            
            iDC = GetWindowDC(hWnd)
            GetWindowRect hWnd, iWindowRect
            iWindowRect.Right = iWindowRect.Right - iWindowRect.Left
            iWindowRect.Bottom = iWindowRect.Bottom - iWindowRect.Top
            iWindowRect.Left = 0
            iWindowRect.Top = 0
            
            TranslateColor mBorderColor, 0&, iColor
            iBrush = CreateSolidBrush(iColor)
            
            iRc = iWindowRect
            iRc.Bottom = iRc.Top + iBs.Width
            FillRect iDC, iRc, iBrush
            
            iRc = iWindowRect
            iRc.Top = iRc.Bottom - iBs.Width
            FillRect iDC, iRc, iBrush
            
            iRc = iWindowRect
            iRc.Right = iRc.Left + iBs.Width
            FillRect iDC, iRc, iBrush
            
            iRc = iWindowRect
            iRc.Left = iRc.Right - iBs.Width
            FillRect iDC, iRc, iBrush
            
            DeleteObject iBrush
            
            ReleaseDC hWnd, iDC
            IBSSubclass_WindowProc = 0
        Else
            iBs = GetBorderStyleData(mBorderStyle)
            
            iDC = GetWindowDC(hWnd)
            GetWindowRect hWnd, iWindowRect
            iWindowRect.Right = iWindowRect.Right - iWindowRect.Left
            iWindowRect.Bottom = iWindowRect.Bottom - iWindowRect.Top
            iWindowRect.Left = 0
            iWindowRect.Top = 0
            
            Call DrawEdge(iDC, iWindowRect, iBs.Flags, BF_RECT)
            
            ReleaseDC hWnd, iDC
            IBSSubclass_WindowProc = 0
        End If
    End If
End Function

Private Sub tmrHScrollInit_Timer()
    tmrHScrollInit.Enabled = False
    pHScrollValue = Val(tmrHScrollInit.Tag)
    tmrHScrollInit.Tag = ""
End Sub

Private Sub tmrVScrollInit_Timer()
    tmrVScrollInit.Enabled = False
    pVScrollValue = Val(tmrVScrollInit.Tag)
    tmrVScrollInit.Tag = ""
End Sub

Private Sub tmrVScrollValue_Timer()
    tmrVScrollValue.Enabled = False
    mScroll.Value(efnSBIVertical) = Val(tmrVScrollValue.Tag)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "UserMode" Then mUserMode = Ambient.UserMode
End Sub

Private Sub tmrDontIncreaseMax_Timer()
    tmrDontIncreaseMax.Enabled = False
End Sub

Private Sub tmrMoveTop_Timer()
    Dim iCtl As Control
    Dim iLng As Long
    
    tmrMoveTop.Enabled = False
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        If TypeName(iCtl) = "Line" Then
            iCtl.Y1 = iCtl.Y1 - mMoveTop
            iCtl.Y2 = iCtl.Y2 - mMoveTop
        Else
            iCtl.Top = iCtl.Top - mMoveTop
        End If
    Next
    On Error GoTo 0
    
    iLng = mVScrollValue \ Screen.TwipsPerPixelY
    If mScroll.Value(efnSBIVertical) <> iLng Then
        mScroll.Value(efnSBIVertical) = iLng
    End If
    If (mMoveTop <> 0) And (Not mAddingControls_v) And (Not mUpdating) Then RaiseEvent VScrollChange
    mMoveTop = 0
End Sub

Private Sub tmrMoveLeft_Timer()
    Dim iCtl As Control
    Dim iLng As Long
    
    tmrMoveLeft.Enabled = False
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        If TypeName(iCtl) = "Line" Then
            iCtl.X1 = iCtl.X1 - mMoveLeft
            iCtl.X2 = iCtl.X2 - mMoveLeft
        Else
            iCtl.Left = iCtl.Left - mMoveLeft
        End If
    Next
    On Error GoTo 0

    iLng = mHScrollValue \ Screen.TwipsPerPixelX
    If mScroll.Value(efnSBIHorizontal) <> iLng Then
        mScroll.Value(efnSBIHorizontal) = iLng
    End If
    If (mMoveLeft <> 0) And (Not mAddingControls_h) And (Not mUpdating) Then RaiseEvent HScrollChange
    mMoveLeft = 0
End Sub

Private Sub AdjustVirtualSpaceToControls()
    Dim c As Long
    Dim iVal As Single
    Dim iVH As Single
    Dim iHW As Single
    
    CreateScrollBars
    
    iVH = mVirtualHeight
    On Error Resume Next
    For c = UserControl.ContainedControls.Count To 1 Step -1
        If TypeName(UserControl.ContainedControls(c - 1)) = "Line" Then
            If UserControl.ContainedControls(c - 1).Y2 > UserControl.ContainedControls(c - 1).Y1 Then
                iVal = UserControl.ContainedControls(c - 1).Y2
            Else
                iVal = UserControl.ContainedControls(c - 1).Y1
            End If
        Else
            iVal = UserControl.ContainedControls(c - 1).Top + UserControl.ContainedControls(c - 1).Height
        End If
        If (iVal + mBottomFreeSpace) > iVH Then
            iVH = (iVal + mBottomFreeSpace)
        End If
    Next c
    On Error GoTo 0
    
    If iVH < UserControl.ScaleHeight Then
        iVH = UserControl.ScaleHeight
    End If
    If iVH > mVirtualHeight Then
        pVirtualHeight = iVH
    End If

    iHW = mVirtualWidth
    On Error Resume Next
    For c = UserControl.ContainedControls.Count To 1 Step -1
        If TypeName(UserControl.ContainedControls(c - 1)) = "Line" Then
            If UserControl.ContainedControls(c - 1).X2 > UserControl.ContainedControls(c - 1).X1 Then
                iVal = UserControl.ContainedControls(c - 1).X2
            Else
                iVal = UserControl.ContainedControls(c - 1).X1
            End If
        Else
            iVal = UserControl.ContainedControls(c - 1).Left + UserControl.ContainedControls(c - 1).Width
        End If
        If (iVal + mRightFreeSpace) > iHW Then
            iHW = (iVal + mRightFreeSpace)
        End If
    Next c
    On Error GoTo 0
    
    If iHW < UserControl.ScaleWidth Then
        iHW = UserControl.ScaleWidth
    End If
    If iHW > mVirtualWidth Then
        pVirtualWidth = iHW
    End If

End Sub

Private Sub mScroll_Change(eBar As efnScrollBarsIdentificationConstants)
    mScroll_Scroll eBar
End Sub

Private Sub mScroll_Scroll(eBar As efnScrollBarsIdentificationConstants)
    Dim iLng As Long
    
    If eBar = efnSBIVertical Then
        pVScrollValue = Screen.TwipsPerPixelY * mScroll.Value(eBar)
        If Not Ambient.UserMode Then
            If mScroll.Value(efnSBIVertical) = mScroll.Max(efnSBIVertical) Then
                If Not tmrDontIncreaseMax.Enabled Then
                    iLng = mScroll.Max(efnSBIVertical) * 1.1
                    If iLng = mScroll.Max(efnSBIVertical) Then
                        iLng = iLng + 10
                    End If
                    mScroll.Max(efnSBIVertical) = mScroll.Max(efnSBIVertical) * 1.1
                    tmrDontIncreaseMax.Enabled = True
                End If
            End If
        End If
    ElseIf eBar = efnSBIHorizontal Then
        pHScrollValue = Screen.TwipsPerPixelX * mScroll.Value(eBar)
        If Not Ambient.UserMode Then
            If mScroll.Value(efnSBIHorizontal) = mScroll.Max(efnSBIHorizontal) Then
                If Not tmrDontIncreaseMax.Enabled Then
                    iLng = mScroll.Max(efnSBIHorizontal) * 1.1
                    If iLng = mScroll.Max(efnSBIHorizontal) Then
                        iLng = iLng + 10
                    End If
                    mScroll.Max(efnSBIHorizontal) = mScroll.Max(efnSBIHorizontal) * 1.1
                    tmrDontIncreaseMax.Enabled = True
                End If
            End If
        End If
    End If
End Sub


Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the Windows handle of the control."
    hWnd = UserControl.hWnd
End Property


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color"
    BackColor = mBackColor
End Property

Public Property Let BackColor(nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
        If UserControl.BackColor <> mBackColor Then
            PropertyChanged "BackColor"
            UserControl.BackColor = mBackColor
        End If
    End If
End Property


Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of the border when it is set to a flat style."
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(nValue As OLE_COLOR)
    If nValue <> mBorderColor Then
        mBorderColor = nValue
        PropertyChanged "BorderColor"
        Call SetWindowPos(UserControl.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_FRAMECHANGED)
    End If
End Property


Private Sub UserControl_Initialize()
    mScrollBarHeight = ScaleY(GetSystemMetrics(SM_CYHSCROLL), vbPixels, vbTwips)
    mScrollBarWidth = ScaleY(GetSystemMetrics(SM_CXVSCROLL), vbPixels, vbTwips)
End Sub

Private Sub UserControl_Show()
    If mTempVScrollValue <> 0 Then
        mNoScroll = True
        mScroll.Max(efnSBIVertical) = mTempVScrollMax
        mScroll.Value(efnSBIVertical) = mTempVScrollValue
        mVScrollValue = Screen.TwipsPerPixelY * mScroll.Value(efnSBIVertical)
        mNoScroll = False
        mVirtualHeight = mTempVirtualHeight
        'mScroll.Value(efnSBIVertical) = 0
        tmrMoveTop_Timer
        mTempVScrollValue = 0
        mTempVScrollMax = 0
        mTempVirtualHeight = 0
    End If
    If mTempHScrollValue <> 0 Then
        mNoScroll = True
        mScroll.Max(efnSBIHorizontal) = mTempHScrollMax
        mScroll.Value(efnSBIHorizontal) = mTempHScrollValue
        mHScrollValue = Screen.TwipsPerPixelX * mScroll.Value(efnSBIHorizontal)
        mNoScroll = False
        mVirtualWidth = mTempVirtualWidth
'        mScroll.Value(efnSBIHorizontal) = 0
        tmrMoveLeft_Timer
        mTempHScrollValue = 0
        mTempHScrollMax = 0
        mTempVirtualWidth = 0
    End If
    mVirtualHeight = 0
    mVirtualWidth = 0
    AdjustVirtualSpaceToControls
    If UserControl.Ambient.UserMode Then
        tmrCheckFocus.Enabled = mAutoScrollOnFocus
        If mAutoScrollOnFocus Then BuildFocusList
    End If
    mShown = True
End Sub

Private Sub UserControl_InitProperties()
    BackColor = Ambient.BackColor
    mVScrollBar = vxScrollBarAuto
    mHScrollBar = vxScrollBarAuto
    mAutoScrollOnFocus = True
    mBorderColor = cDefaultBorderColor
    mBorderStyle = cDefaultBorderStyle
    mBottomFreeSpace = cDefaultBottomFreeSpace
    SetBorderStyle
    On Error Resume Next
    mUserMode = Ambient.UserMode
    mUserControlHwnd = UserControl.hWnd
    On Error GoTo 0
    If mUserControlHwnd <> 0 Then
        AttachMessage Me, mUserControlHwnd, WM_NCPAINT
        mSubclassed = True
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    mTempVScrollValue = PropBag.ReadProperty("SavedVScrollValue", 0)
    mTempVScrollMax = PropBag.ReadProperty("SavedVScrollMax", 0)
    mTempVirtualHeight = PropBag.ReadProperty("SavedVirtualHeight", 0)
    mTempHScrollValue = PropBag.ReadProperty("SavedHScrollValue", 0)
    mTempHScrollMax = PropBag.ReadProperty("SavedHScrollMax", 0)
    mTempVirtualWidth = PropBag.ReadProperty("SavedVirtualWidth", 0)
    mBottomFreeSpace = PropBag.ReadProperty("BottomFreeSpace", cDefaultBottomFreeSpace)
    mRightFreeSpace = PropBag.ReadProperty("RightFreeSpace", 0)
    mVScrollBar = PropBag.ReadProperty("VScrollBar", vxScrollBarAuto)
    mHScrollBar = PropBag.ReadProperty("HScrollBar", vxScrollBarAuto)
    mAutoScrollOnFocus = PropBag.ReadProperty("AutoScrollOnFocus", True)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", cDefaultBorderStyle)
    mBorderColor = PropBag.ReadProperty("BorderColor", cDefaultBorderColor)
    mVScrollValue = PropBag.ReadProperty("VScrollValue", 0)
    mHScrollValue = PropBag.ReadProperty("HScrollValue", 0)
    mTopScrollBound = PropBag.ReadProperty("TopScrollBound", 0)
    
    SetBorderStyle
    CreateScrollBars
    On Error Resume Next
    mUserMode = Ambient.UserMode
    mUserControlHwnd = UserControl.hWnd
    On Error GoTo 0
    If mUserControlHwnd <> 0 Then
        AttachMessage Me, mUserControlHwnd, WM_NCPAINT
        mSubclassed = True
    End If
End Sub

Private Sub CreateScrollBars()
    If mScroll Is Nothing Then
        Set mScroll = New cScrollBars
        mScroll.Create UserControl.hWnd
    End If
End Sub

Private Sub UserControl_Terminate()
    Unsubclass
    Set mFocusHwndList = Nothing
    tmrCheckFocus.Enabled = False
    tmrDontIncreaseMax.Enabled = False
    tmrMoveLeft.Enabled = False
    tmrMoveTop.Enabled = False
End Sub

Private Sub Unsubclass()
    If mSubclassed Then
        DetachMessage Me, mUserControlHwnd, WM_NCPAINT
        mSubclassed = False
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", mBackColor, vbButtonFace
    PropBag.WriteProperty "SavedVScrollValue", mScroll.Value(efnSBIVertical), 0
    PropBag.WriteProperty "SavedVScrollMax", mScroll.Max(efnSBIVertical), 0
    PropBag.WriteProperty "SavedVirtualHeight", mVirtualHeight, 0
    PropBag.WriteProperty "SavedHScrollValue", mScroll.Value(efnSBIHorizontal), 0
    PropBag.WriteProperty "SavedHScrollMax", mScroll.Max(efnSBIHorizontal), 0
    PropBag.WriteProperty "SavedVirtualWidth", mVirtualWidth, 0
    PropBag.WriteProperty "BottomFreeSpace", mBottomFreeSpace, cDefaultBottomFreeSpace
    PropBag.WriteProperty "RightFreeSpace", mRightFreeSpace, 0
    PropBag.WriteProperty "VScrollBar", mVScrollBar, vxScrollBarAuto
    PropBag.WriteProperty "HScrollBar", mHScrollBar, vxScrollBarAuto
    PropBag.WriteProperty "AutoScrollOnFocus", mAutoScrollOnFocus, True
    PropBag.WriteProperty "BorderStyle", mBorderStyle, cDefaultBorderStyle
    PropBag.WriteProperty "BorderColor", mBorderColor, cDefaultBorderColor
    PropBag.WriteProperty "VScrollValue", mVScrollValue, 0
    PropBag.WriteProperty "HScrollValue", mHScrollValue, 0
    PropBag.WriteProperty "TopScrollBound", mTopScrollBound, 0
End Sub

Private Sub UserControl_Resize()
    If (UserControl.ScaleX(UserControl.Width, vbTwips, vbPixels) < 75) Then
        UserControl.Width = 75 * Screen.TwipsPerPixelX
    End If
    If (UserControl.ScaleY(UserControl.Height, vbTwips, vbPixels) < 75) Then
        UserControl.Height = 75 * Screen.TwipsPerPixelY
    End If
    SetWindowRedraw UserControl.hWnd, False
    CreateScrollBars
    Update
    SetWindowRedraw UserControl.hWnd, True
End Sub


Public Property Get BottomFreeSpace() As Single
Attribute BottomFreeSpace.VB_Description = "Returns/sets a value that determines the free space that will be left at the bottom of the virtual space."
    BottomFreeSpace = FixRoundingError(ToContainerSizeY(mBottomFreeSpace, vbTwips))
End Property

Public Property Let BottomFreeSpace(nValue As Single)
    Dim iValue As Single
    
    If nValue < 0 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    
    iValue = FromContainerSizeY(nValue, vbTwips)
    If iValue <> mBottomFreeSpace Then
        mBottomFreeSpace = iValue
        PropertyChanged "BottomFreeSpace"
        AdjustVirtualSpaceToControls
    End If
End Property


Public Property Get RightFreeSpace() As Single
Attribute RightFreeSpace.VB_Description = "Returns/sets a value that determines the free space that will be left at the right of the virtual space."
    RightFreeSpace = FixRoundingError(ToContainerSizeX(mRightFreeSpace, vbTwips))
End Property

Public Property Let RightFreeSpace(nValue As Single)
    Dim iValue As Single
    
    If nValue < 0 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    
    iValue = FromContainerSizeX(nValue, vbTwips)
    If iValue <> mRightFreeSpace Then
        mRightFreeSpace = iValue
        PropertyChanged "RightFreeSpace"
        AdjustVirtualSpaceToControls
    End If
End Property


Public Property Get VScrollValue() As Single
Attribute VScrollValue.VB_Description = "Returns or sets a value that idicates the vertical scroll actual position."
    VScrollValue = FixRoundingError(ToContainerSizeY(mVScrollValue, vbTwips))
End Property

Public Property Let VScrollValue(nValue As Single)
    pVScrollValue = FromContainerSizeY(nValue, vbTwips)
    tmrMoveTop_Timer
    PropertyChanged "VScrollValue"
End Property

Private Property Let pVScrollValue(nValue As Single)
    If Not mShown Then
        tmrVScrollInit.Tag = nValue
        tmrVScrollInit.Enabled = True
        Exit Property
    End If
    If mNoScroll Then Exit Property
'    If nValue < 0 Then
'        RaiseError 380, TypeName(Me) ' invalid property value
'        Exit Property
'    End If
    If nValue <> mVScrollValue Then
        If nValue > (mVirtualHeight - UserControl.ScaleHeight) Then
            If Ambient.UserMode Then
                nValue = mVirtualHeight - UserControl.ScaleHeight
            Else
                pVirtualHeight = nValue + UserControl.ScaleHeight
            End If
        End If
        If mTopScrollBound <> 0 Then
            If nValue > (mTopScrollBound - UserControl.ScaleHeight) Then
                tmrVScrollValue.Tag = (mTopScrollBound - UserControl.ScaleHeight) / Screen.TwipsPerPixelY
                tmrVScrollValue.Enabled = True
            End If
        End If
        If Not tmrVScrollValue.Enabled Then
            mMoveTop = mMoveTop + nValue - mVScrollValue
            mVScrollValue = nValue
            tmrMoveTop.Enabled = True
        End If
    End If
End Property


Public Property Get HScrollValue() As Single
Attribute HScrollValue.VB_Description = "Returns or sets a value that idicates the horizontal scroll actual position."
    HScrollValue = FixRoundingError(ToContainerSizeX(mHScrollValue, vbTwips))
End Property

Public Property Let HScrollValue(nValue As Single)
    pHScrollValue = FromContainerSizeY(nValue, vbTwips)
    tmrMoveLeft_Timer
    PropertyChanged "HScrollValue"
End Property

Private Property Let pHScrollValue(nValue As Single)
    If Not mShown Then
        tmrHScrollInit.Tag = nValue
        tmrHScrollInit.Enabled = True
        Exit Property
    End If
    If mNoScroll Then Exit Property
'    If nValue < 0 Then
'        RaiseError 380, TypeName(Me) ' invalid property value
'        Exit Property
'    End If
    
    If nValue <> mHScrollValue Then
        If nValue > (mVirtualWidth - UserControl.ScaleWidth) Then
            If Ambient.UserMode Then
                nValue = mVirtualWidth - UserControl.ScaleWidth
            Else
                pVirtualWidth = nValue + UserControl.ScaleWidth
            End If
        End If
        mMoveLeft = mMoveLeft + nValue - mHScrollValue
        mHScrollValue = nValue
        tmrMoveLeft.Enabled = True
    End If
End Property


Public Property Get VirtualHeight() As Single
Attribute VirtualHeight.VB_Description = "Returns or sets a value that determines the height of the virtual space where the controls are located."
Attribute VirtualHeight.VB_MemberFlags = "200"
    On Error GoTo ErrExit
    
    VirtualHeight = FixRoundingError(ToContainerSizeY(mVirtualHeight, vbTwips))
    
    If Not Ambient.UserMode Then
        If Abs(VirtualHeight + mScrollBarHeight + GetBorderStyleData(mBorderStyle).Width * Screen.TwipsPerPixelY * 2 - UserControl.Extender.Height) < Screen.TwipsPerPixelY Then
            VirtualHeight = UserControl.Extender.Height
        End If
    End If
    
ErrExit:
End Property

Public Property Let VirtualHeight(ByVal nValue As Single)
    On Error GoTo ErrExit
    
    If nValue < UserControl.Extender.Height Then
        nValue = UserControl.Extender.Height
    End If
    If Not Ambient.UserMode Then
        If Abs(nValue + mScrollBarHeight + GetBorderStyleData(mBorderStyle).Width * Screen.TwipsPerPixelY * 2 - UserControl.Extender.Height) < Screen.TwipsPerPixelY Then
            nValue = UserControl.Extender.Height - mScrollBarHeight
        End If
    End If
    If nValue < 0 Then nValue = 0
    
    pVirtualHeight = FromContainerSizeY(nValue, vbTwips)
    PropertyChanged "VirtualHeight"
ErrExit:
End Property

Private Property Let pVirtualHeight(nValue As Single)
    If nValue <> mVirtualHeight Then
        Dim iVisible As Boolean
        
        iVisible = mScroll.Visible(efnSBIVertical)
        mVirtualHeight = nValue
        
        If mVirtualHeight < UserControl.ScaleHeight Then
            mVirtualHeight = UserControl.ScaleHeight
        End If
        If (mVirtualHeight > UserControl.ScaleHeight) Then
            mScroll.LargeChange(efnSBIVertical) = UserControl.ScaleHeight * 0.9 \ Screen.TwipsPerPixelY
            mScroll.SmallChange(efnSBIVertical) = mScroll.LargeChange(efnSBIVertical) / 10
            mScroll.Max(efnSBIVertical) = (mVirtualHeight - UserControl.ScaleHeight) \ Screen.TwipsPerPixelY
            mScroll.Visible(efnSBIVertical) = (mVScrollBar <> vxScrollBarHide)
            mScroll.Enabled(efnSBIVertical) = True
        Else
            If Ambient.UserMode Then
                mScroll.Visible(efnSBIVertical) = (mVScrollBar = vxScrollBarShow)
                mScroll.Enabled(efnSBIVertical) = False
            Else
                mScroll.LargeChange(efnSBIVertical) = UserControl.ScaleHeight \ Screen.TwipsPerPixelY \ 2
                mScroll.SmallChange(efnSBIVertical) = mScroll.LargeChange(efnSBIVertical) / 10
                mScroll.Max(efnSBIVertical) = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
                mScroll.Visible(efnSBIVertical) = (mVScrollBar <> vxScrollBarHide)
                mScroll.Enabled(efnSBIVertical) = True
            End If
        End If
        mScroll.Value(efnSBIVertical) = mVScrollValue \ Screen.TwipsPerPixelY
        If mScroll.Visible(efnSBIVertical) <> iVisible Then
            If mVirtualWidth <> 0 Then
                mVirtualWidth = mVirtualWidth - 1
                pVirtualWidth = mVirtualWidth + 1
            End If
        End If
        If mTopScrollBound > nValue Then
            mTopScrollBound = 0
        End If
    End If
End Property


Public Property Get VirtualWidth() As Single
Attribute VirtualWidth.VB_Description = "Returns or sets a value that determines the width of the virtual space where the controls are located."
    On Error GoTo ErrExit
    
    VirtualWidth = FixRoundingError(ToContainerSizeX(mVirtualWidth, vbTwips))
    If Not Ambient.UserMode Then
        If Abs(VirtualWidth + mScrollBarWidth + GetBorderStyleData(mBorderStyle).Width * Screen.TwipsPerPixelX * 2 - UserControl.Extender.Width) < Screen.TwipsPerPixelX Then
            VirtualWidth = UserControl.Extender.Width
        End If
    End If
    
ErrExit:
End Property

Public Property Let VirtualWidth(nValue As Single)
    On Error GoTo ErrExit
    
    If nValue < UserControl.Extender.Width Then
        nValue = UserControl.Extender.Width
    End If
    If Not Ambient.UserMode Then
        If Abs(nValue + mScrollBarWidth + GetBorderStyleData(mBorderStyle).Width * Screen.TwipsPerPixelX * 2 - UserControl.Extender.Width) < Screen.TwipsPerPixelX Then
            nValue = UserControl.Extender.Width - mScrollBarWidth
        End If
    End If
    If nValue < 0 Then nValue = 0
    
    pVirtualWidth = FromContainerSizeY(nValue, vbTwips)
    PropertyChanged "VirtualWidth"
ErrExit:
End Property

Private Property Let pVirtualWidth(nValue As Single)
    If nValue <> mVirtualWidth Then
        Dim iVisible As Boolean
        
        iVisible = mScroll.Visible(efnSBIHorizontal)
        mVirtualWidth = nValue
        
        If mVirtualWidth < UserControl.ScaleWidth Then
            mVirtualWidth = UserControl.ScaleWidth
        End If
        If (mVirtualWidth > UserControl.ScaleWidth) Then
            mScroll.LargeChange(efnSBIHorizontal) = UserControl.ScaleWidth * 0.9 \ Screen.TwipsPerPixelX
            mScroll.SmallChange(efnSBIHorizontal) = mScroll.LargeChange(efnSBIHorizontal) / 10
            mScroll.Max(efnSBIHorizontal) = (mVirtualWidth - UserControl.ScaleWidth) \ Screen.TwipsPerPixelX
            mScroll.Visible(efnSBIHorizontal) = (mHScrollBar <> vxScrollBarHide)
            mScroll.Enabled(efnSBIHorizontal) = True
        Else
            If Ambient.UserMode Then
                mScroll.Visible(efnSBIHorizontal) = (mHScrollBar = vxScrollBarShow)
                mScroll.Enabled(efnSBIHorizontal) = False
            Else
                mScroll.LargeChange(efnSBIHorizontal) = UserControl.ScaleWidth \ Screen.TwipsPerPixelX \ 2
                mScroll.SmallChange(efnSBIHorizontal) = mScroll.LargeChange(efnSBIHorizontal) / 10
                mScroll.Max(efnSBIHorizontal) = UserControl.ScaleWidth \ Screen.TwipsPerPixelX
                mScroll.Visible(efnSBIHorizontal) = (mHScrollBar <> vxScrollBarHide)
                mScroll.Enabled(efnSBIHorizontal) = True
            End If
         End If
         mScroll.Value(efnSBIHorizontal) = mHScrollValue \ Screen.TwipsPerPixelX
         If mScroll.Visible(efnSBIHorizontal) <> iVisible Then
            If mVirtualHeight <> 0 Then
                mVirtualHeight = mVirtualHeight - 1
                pVirtualHeight = mVirtualHeight + 1
            End If
         End If
    End If
End Property


Private Function ContainerScaleMode() As ScaleModeConstants
    ContainerScaleMode = vbTwips
    On Error Resume Next
    ContainerScaleMode = UserControl.Extender.Container.ScaleMode
End Function

Private Function FromContainerSizeY(nValue As Variant, Optional nToScale As ScaleModeConstants = vbTwips) As Single
    FromContainerSizeY = UserControl.ScaleY(nValue, ContainerScaleMode, nToScale)
End Function

Private Function ToContainerSizeY(nValue As Variant, Optional nFromScale As ScaleModeConstants = vbTwips) As Single
    ToContainerSizeY = UserControl.ScaleY(nValue, nFromScale, ContainerScaleMode)
End Function


Private Function FromContainerSizeX(nValue As Variant, Optional nToScale As ScaleModeConstants = vbTwips) As Single
    FromContainerSizeX = UserControl.ScaleX(nValue, ContainerScaleMode, nToScale)
End Function

Private Function ToContainerSizeX(nValue As Variant, Optional nFromScale As ScaleModeConstants = vbTwips) As Single
    ToContainerSizeX = UserControl.ScaleY(nValue, nFromScale, ContainerScaleMode)
End Function

Private Function FixRoundingError(nNumber As Single, Optional nDecimals As Long = 3) As Single
    Dim iNum As Single
    
    iNum = Round(nNumber * 10 ^ nDecimals) / 10 ^ nDecimals
    
    If iNum = Int(iNum) Then
        FixRoundingError = iNum
    Else
        If (ContainerScaleMode = vbTwips) Or (ContainerScaleMode = vbPixels) Then
            FixRoundingError = Round(nNumber)
        Else
            FixRoundingError = nNumber
        End If
    End If
End Function

Public Property Get VScrollBar() As vbExScrollBarVisibilityConstants
Attribute VScrollBar.VB_Description = "Returns or sets a value that determines the vertical scrollbar visibility at run time."
    VScrollBar = mVScrollBar
End Property

Public Property Let VScrollBar(nValue As vbExScrollBarVisibilityConstants)
    If (nValue < 0) Or (nValue > 2) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mVScrollBar Then
        mVScrollBar = nValue
        PropertyChanged "VScrollBar"
    End If
End Property


Public Property Get HScrollBar() As vbExScrollBarVisibilityConstants
Attribute HScrollBar.VB_Description = "Returns or sets a value that determines the horizontal scrollbar visibility at run time."
    HScrollBar = mHScrollBar
End Property

Public Property Let HScrollBar(nValue As vbExScrollBarVisibilityConstants)
    If (nValue < 0) Or (nValue > 2) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mHScrollBar Then
        mHScrollBar = nValue
        PropertyChanged "HScrollBar"
    End If
End Property


Public Property Get VScrollMax() As Single
Attribute VScrollMax.VB_Description = "Returns a value that idicates the maximum value that VScrollValue can take."
    VScrollMax = FixRoundingError(ToContainerSizeY(mScroll.Max(efnSBIVertical), vbPixels))
End Property

Public Property Get HScrollMax() As Single
Attribute HScrollMax.VB_Description = "Returns a value that idicates the maximum value that HScrollValue can take."
    HScrollMax = FixRoundingError(ToContainerSizeX(mScroll.Max(efnSBIHorizontal), vbPixels))
End Property


Public Property Get BorderStyle() As vbExExtendedBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets a value that determines how the border of the control looks like."
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(nValue As vbExExtendedBorderStyleConstants)
    If nValue <> mBorderStyle Then
        mBorderStyle = nValue
        PropertyChanged "BorderStyle"
        SetBorderStyle
        Call SetWindowPos(UserControl.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_FRAMECHANGED)
    End If
End Property


Public Property Get AutoScrollOnFocus() As Boolean
Attribute AutoScrollOnFocus.VB_Description = "Returns/sets a value that determines if when a contained control out of view gets the focus, if it will automatically scroll to show that control."
    AutoScrollOnFocus = mAutoScrollOnFocus
End Property

Public Property Let AutoScrollOnFocus(nValue As Boolean)
    If nValue <> mAutoScrollOnFocus Then
        mAutoScrollOnFocus = nValue
        PropertyChanged "AutoScrollOnFocus"
        If UserControl.Ambient.UserMode Then
            tmrCheckFocus.Enabled = mAutoScrollOnFocus
            If mAutoScrollOnFocus Then BuildFocusList
        End If
    End If
End Property

Private Sub BuildFocusList()
    Dim iCtl As Control
    Dim iHwnd As Long
    
    Set mFocusHwndList = New Collection
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        iHwnd = 0
        iHwnd = iCtl.hWnd
        If iHwnd <> 0 Then
            mFocusHwndList.Add iHwnd, CStr(iHwnd)
        End If
    Next
End Sub

Private Sub tmrCheckFocus_Timer()
    Dim iCtl As Control
    Dim iHwnd As Long
    Dim iHwnd2 As Long
    Static sLastHwnd As Long
    
    On Error GoTo TheExit
    Set iCtl = Screen.ActiveControl
    iHwnd = GetFocus
    
    iHwnd2 = 0
    iHwnd2 = mFocusHwndList(CStr(iHwnd))
    If iHwnd2 <> 0 Then
        If iHwnd <> sLastHwnd Then
            Set iCtl = GetControlByHwnd(iHwnd)
            sLastHwnd = iHwnd
            If Not iCtl Is Nothing Then
                EnsureControlVisible iCtl
            End If
        End If
    End If
    
TheExit:
End Sub

Private Function GetControlByHwnd(nHwnd As Long) As Object
    Dim iCtl As Control
    Dim iHwnd As Long
    
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        iHwnd = 0
        iHwnd = iCtl.hWnd
        If iHwnd <> 0 Then
            If iHwnd = nHwnd Then
                Set GetControlByHwnd = iCtl
                Exit Function
            End If
        End If
    Next
End Function
    

Public Sub Update()
Attribute Update.VB_Description = "Updates the virtual space dimensions."
    Dim V As Single
    Dim h As Single
    
    mUpdating = True
    V = VScrollValue
    h = HScrollValue
    VScrollValue = 0
    HScrollValue = 0
    mVirtualHeight = 0
    mVirtualWidth = 0
    AdjustVirtualSpaceToControls
    If V <> 0 Then
        VScrollValue = V
    End If
    If h <> 0 Then
        HScrollValue = h
    End If
    mUpdating = False
End Sub


Public Property Get AddingControls() As Boolean
Attribute AddingControls.VB_Description = "Use this property when adding controls to the container at run time."
Attribute AddingControls.VB_MemberFlags = "400"
    AddingControls = mAddingControls
End Property

Public Property Let AddingControls(nValue As Boolean)
    If nValue <> mAddingControls Then
        If nValue Then mAddingControls = True
        If nValue Then
            mAddingControls_v = VScrollValue
            mAddingControls_h = HScrollValue
            VScrollValue = 0
            HScrollValue = 0
        Else
            mVirtualHeight = 0
            mVirtualWidth = 0
            AdjustVirtualSpaceToControls
            If mAddingControls_v <> 0 Then
                VScrollValue = mAddingControls_v
            End If
            If mAddingControls_h <> 0 Then
                HScrollValue = mAddingControls_h
            End If
            Update
            mAddingControls = False
        End If
    End If
End Property

Public Sub EnsureControlVisible(nControl As Object)
Attribute EnsureControlVisible.VB_Description = "Ensures that the control referenced in the nControl parameter is visible on the container."
    Dim iSW As Single
    Dim iSH As Single
    Dim iVal As Single
    Dim iCtl As Control
    Dim iFound As Boolean
    
    For Each iCtl In UserControl.ContainedControls
        If iCtl Is nControl Then
            iFound = True
            Exit For
        End If
    Next
    
    If Not iFound Then
         RaiseError 1390, TypeName(Me), "The contained controls collection could not be found."
         Exit Sub
    End If
    
    If mScroll.Visible(efnSBIHorizontal) Then
        iSW = ToContainerSizeX(UserControl.ScaleWidth, vbTwips)
        If iCtl.Left + iCtl.Width > iSW Then
            HScrollValue = HScrollValue + iCtl.Left + iCtl.Width + ToContainerSizeX(60, vbTwips) - iSW
        ElseIf iCtl.Left < 0 Then
            iVal = HScrollValue + iCtl.Left - ToContainerSizeX(60, vbTwips)
            If iVal < 0 Then iVal = 0
            HScrollValue = iVal
        End If
    End If
    
    If mScroll.Visible(efnSBIVertical) Then
        iSH = ToContainerSizeY(UserControl.ScaleHeight, vbTwips)
        If iCtl.Top + iCtl.Height > iSH Then
            VScrollValue = VScrollValue + iCtl.Top + iCtl.Height + ToContainerSizeY(60, vbTwips) - iSH
        ElseIf iCtl.Top < 0 Then
            iVal = VScrollValue + iCtl.Top - ToContainerSizeY(60, vbTwips)
            If iVal < 0 Then iVal = 0
            VScrollValue = iVal
        End If
    End If
End Sub

Private Function GetBorderStyleData(nBs As vbExExtendedBorderStyleConstants) As BORDERSTYLE_DATA
    Dim iRet As BORDERSTYLE_DATA
    
    Select Case nBs
        Case vxEBSNone
            iRet.Flags = 0
            iRet.Width = 0
        Case vxEBSFlat1Pix
            iRet.Flags = -1
            iRet.Width = 1
        Case vxEBSSunkenOuter1Pix
            iRet.Flags = BDR_SUNKENOUTER
            iRet.Width = 1
        Case vxEBSSunkenInner1Pix
            iRet.Flags = BDR_SUNKENINNER
            iRet.Width = 1
        Case vxEBSRaisedOuter1Pix
            iRet.Flags = BDR_RAISEDOUTER
            iRet.Width = 1
        Case vxEBSRaisedInner1Pix
            iRet.Flags = BDR_RAISEDINNER
            iRet.Width = 1
        Case vxEBSFlat2Pix
            iRet.Flags = -1
            iRet.Width = 2
        Case vxEBSSunken2Pix
            iRet.Flags = BDR_SUNKENOUTER Or BDR_SUNKENINNER
            iRet.Width = 2
        Case vxEBSRaised2Pix
            iRet.Flags = BDR_RAISEDOUTER Or BDR_RAISEDINNER
            iRet.Width = 2
        Case vxEBSEtched2Pix
            iRet.Flags = BDR_SUNKENOUTER Or BDR_RAISEDINNER
            iRet.Width = 2
    End Select
    
    GetBorderStyleData = iRet
End Function

Private Sub SetBorderStyle()
    Dim iBs As BORDERSTYLE_DATA
    
    iBs = GetBorderStyleData(mBorderStyle)
    If iBs.Width = 0 Then
        UserControl.Appearance = 1
        UserControl.BorderStyle = 0
    ElseIf iBs.Width = 1 Then
        UserControl.Appearance = 0
        UserControl.BorderStyle = 1
    Else
        UserControl.Appearance = 1
        UserControl.BorderStyle = 1
    End If
    UserControl.BackColor = mBackColor
End Sub

Private Sub SetWindowRedraw(nHwnd As Long, nRedraw As Boolean, Optional nForce As Boolean)
    
    If Not nRedraw Then
        If IsWindowVisible(nHwnd) = 0 Then Exit Sub
    End If
    
    Static sHwnds() As Long
    Static sCalls() As Long
    Dim c As Long
    Dim t As Long
    Dim i As Long
   
    i = 0
    On Error Resume Next
    Err.Clear
    t = UBound(sHwnds)
    If Err.Number = 9 Then
        ReDim sHwnds(0)
        ReDim sCalls(0)
        t = 0
    Else
        For c = 1 To t
            If sHwnds(c) = nHwnd Then
                i = c
                Exit For
            End If
        Next c
    End If
    On Error GoTo 0
    If (i = 0) Then
        If nRedraw Then Exit Sub
        ReDim Preserve sHwnds(t + 1)
        sHwnds(t + 1) = nHwnd
        ReDim Preserve sCalls(t + 1)
        sCalls(t + 1) = 1
        i = 1
    Else
        If nRedraw Then
            sCalls(i) = sCalls(i) - 1
            If sCalls(i) < 0 Then sCalls(i) = 0
        Else
            sCalls(i) = sCalls(i) + 1
        End If
    End If
    If nRedraw And nForce Then
        SendMessageLong nHwnd, WM_SETREDRAW, True, 0&
        RedrawWindow nHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
        sCalls(i) = 0
    Else
        Select Case sCalls(i)
            Case 1
                SendMessageLong nHwnd, WM_SETREDRAW, False, 0&
            Case 0
                SendMessageLong nHwnd, WM_SETREDRAW, True, 0&
                RedrawWindow nHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
        End Select
    End If
End Sub

Private Sub RaiseError(ByVal Number As Long, Optional ByVal Source As Variant, Optional ByVal Description As Variant, Optional ByVal HelpFile As Variant, Optional ByVal HelpContext As Variant)
    If InIDE Then
        On Error Resume Next
        Err.Raise Number, Source, Description, HelpFile, HelpContext
        MsgBox "Error " & Err.Number & ". " & Err.Description, vbCritical
    Else
        Err.Raise Number, Source, Description, HelpFile, HelpContext
    End If
End Sub

Private Function InIDE() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Err.Clear
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
    End If
    InIDE = (sValue = 1)
End Function

Public Property Let TopScrollBound(ByVal nValue As Single)
    If nValue >= VirtualHeight Then nValue = 0
    If nValue <> mTopScrollBound Then
        mTopScrollBound = FromContainerSizeY(nValue, vbTwips)
        PropertyChanged "TopScrollBound"
    End If
End Property

Public Property Get TopScrollBound() As Single
    If mTopScrollBound = 0 Then
        TopScrollBound = FixRoundingError(ToContainerSizeY(mVirtualHeight, vbTwips))
    Else
        TopScrollBound = FixRoundingError(ToContainerSizeY(mTopScrollBound, vbTwips))
    End If
End Property
