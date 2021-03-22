VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--- for MST subclassing (1) https://www.vbforums.com/showthread.php?872819
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)

Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private m_pSubclass         As IUnknown
'--- End for MST subclassing (1)

' ===========================================================================
' Name:     cScrollBars
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     24 December 1998
' Requires: SSUBTMR.DLL
'
' ---------------------------------------------------------------------------
' Copyright ® 1998 Steve McMahon (steve@vbaccelerator.com)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
'
' Description:
' A class which can add scroll bars to VB Forms, Picture Boxes and
' UserControls.
' Features:
'  * True API scroll bars, which don't flash or draw badly like
'    the VB ones
'  * Scroll bar values are long integers, i.e. >2 billion values
'  * Set Flat or Encarta scroll bar modes if your COMCTL32.DLL version
'    supports it (>4.72)
'
' Updates:
' 2003-07-02
'  * Added Mouse Wheel Support.  Thanks to Chris Eastwood for
'    the suggestion and starter code.
'    Visit his site at http://vbcodelibrary.co.uk/
'  * Scroll bar now goes to bottom when SB_BOTTOM fired
'    (e.g. right click on scroll bar with mouse)
'  * New ScrollClick events to enable focus
'  * Removed a large quantity of redundant declares which
'    had found their way into this class somehow...
' ===========================================================================


' ---------------------------------------------------------------------
' vbAccelerator Software License
' Version 1.0
' Copyright (c) 2002 vbAccelerator.com
'
' Redistribution and use in source and binary forms, with or
' without modification, are permitted provided that the following
' conditions are met:
'
' 1. Redistributions of source code must retain the above copyright
'    notice, this list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in
'    the documentation and/or other materials provided with the distribution.
'
' 3. The end-user documentation included with the redistribution, if any,
'    must include the following acknowledgment:
'
'  "This product includes software developed by vbAccelerator (http://vbaccelerator.com/)."
'
' Alternately, this acknowledgment may appear in the software itself, if
' and wherever such third-party acknowledgments normally appear.
'
' 4. The name "vbAccelerator" must not be used to endorse or promote products
'    derived from this software without prior written permission. For written
'    permission, please contact vbAccelerator through steve@vbaccelerator.com.
'
' 5. Products derived from this software may not be called "vbAccelerator",
'    nor may "vbAccelerator" appear in their name, without prior written
'    permission of vbAccelerator.
'
' THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES,
' INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
' AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
' VBACCELERATOR OR ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
' INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
' BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
' USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
' THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' ---------------------------------------------------------------------

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Long, ByVal fuWinIni As Long) As Long

'private declare function InitializeFlatSB(hWnd) as long
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lHWnd As Long) As Long

' Scroll bar:
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, LPSCROLLINFO As SCROLLINFO) As Long

Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_Left = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_Right = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
   
'   Private Const SBS_SIZEGRIP = &H10&
   
Private Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

' Non-client messages:
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_NCMBUTTONDOWN = &HA7

' Hit test codes for scroll bars:
Private Const HTHSCROLL = 6
Private Const HTVSCROLL = 7

' Scroll bar messages:
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114
Private Const WM_MOUSEWHEEL = &H20A

' Mouse wheel stuff:
Private Const WHEEL_DELTA = 120
Private Const SPI_GETWHEELSCROLLLINES = &H68

' Old school Wheel Mouse is not supported in this class.
' NT3.51 or Win95 only
'// Class name for MSWHEEL.EXE's invisible window
'// use FindWindow to get hWnd to MSWHEEL


'// Apps need to call RegisterWindowMessage using the #defines
'// below to get the message numbers for:
'// 1) the message that can be sent to the MSWHEEL window to
'//    query if wheel support is active (MSH_WHEELSUPPORT)>
'// 2) the message to query for the number of scroll lines
'//    (MSH_SCROLL_LINES)
'//
'// To send a message to MSWheel window, use FindWindow with the #defines
'// for CLASS and TITLE above.  If FindWindow fails to find the MSWHEEL
'// window or the return from SendMessage is false, then Wheel support
'// is not currently available.

' Flat scroll bars:
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&

Private Const FSB_FLAT_MODE = 2&
Private Const FSB_ENCARTA_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hWnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long

Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hWnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hWnd As Long) As Long

' Initialisation state:
Private m_bInitialised As Boolean

' Orientation
Public Enum efnScrollBarsOrientationConstants
    efnSBOHorizontal
    efnSBOVertical
    efnSBOBoth
End Enum

' Bars:
Public Enum efnScrollBarsIdentificationConstants
    efnSBIHorizontal = SB_HORZ
    efnSBIVertical = SB_VERT
End Enum

' Style
Public Enum efnScrollBarsStyleConstants
    efnSBSRegular = FSB_REGULAR_MODE
    efnSBSEncarta = FSB_ENCARTA_MODE
    efnSBSFlat = FSB_FLAT_MODE
End Enum


Private m_eOrientation As efnScrollBarsOrientationConstants
Private m_eStyle As efnScrollBarsStyleConstants

' Can we have flat scroll bars?
Private m_bNoFlatScrollBars As Boolean

' hWnd we're adding scroll bars too:
Private m_hWnd As Long

' Small change amount
Private m_lSmallChangeHorz As Long
Private m_lSmallChangeVert As Long
' Enabled:
Private m_bEnabledHorz As Boolean
Private m_bEnabledVert As Boolean
' Visible
Private m_bVisibleHorz As Boolean
Private m_bVisibleVert As Boolean

' Number of lines to scroll for each wheel click:
Private m_lWheelScrollLines  As Long

Public Event ScrollClick(eBar As efnScrollBarsIdentificationConstants, eButton As MouseButtonConstants)
Public Event Scroll(eBar As efnScrollBarsIdentificationConstants)
Public Event Change(eBar As efnScrollBarsIdentificationConstants)
Public Event MouseWheel(eBar As efnScrollBarsIdentificationConstants, lAmount As Long)

Public Property Get Visible(ByVal eBar As efnScrollBarsIdentificationConstants) As Boolean
    If (eBar = efnSBIHorizontal) Then
        Visible = m_bVisibleHorz
    Else
        Visible = m_bVisibleVert
    End If
End Property

Public Property Let Visible(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal bState As Boolean)
    If (eBar = efnSBIHorizontal) Then
        m_bVisibleHorz = bState
    Else
        m_bVisibleVert = bState
    End If
    If (m_bNoFlatScrollBars) Then
        ShowScrollBar m_hWnd, eBar, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
    End If
End Property

Public Property Get Orientation() As efnScrollBarsOrientationConstants
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eOrientation As efnScrollBarsOrientationConstants)
    m_eOrientation = eOrientation
    pSetOrientation
End Property

Private Sub pSetOrientation()
    ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = efnSBOBoth) Or (m_eOrientation = efnSBOHorizontal))
    ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = efnSBOBoth) Or (m_eOrientation = efnSBOVertical))
End Sub

Private Sub pGetSI(ByVal eBar As efnScrollBarsIdentificationConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim lO As Long
    
    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, lO, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, lO, tSI
    End If
End Sub

Private Sub pLetSI(ByVal eBar As efnScrollBarsIdentificationConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim lO As Long
        
    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, lO, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, lO, tSI, True
    End If
End Sub

Public Property Get Style() As efnScrollBarsStyleConstants
    Style = m_eStyle
End Property

Public Property Let Style(ByVal eStyle As efnScrollBarsStyleConstants)
    Dim lR As Long
    
    If (eStyle <> efnSBSRegular) Then
        If (m_bNoFlatScrollBars) Then
            ' can't do it..
            Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
            Exit Property
        End If
    End If
    If (m_eOrientation = efnSBOHorizontal) Or (m_eOrientation = efnSBOBoth) Then
        lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
    End If
    If (m_eOrientation = efnSBOVertical) Or (m_eOrientation = efnSBOBoth) Then
        lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
    End If
    m_eStyle = eStyle
End Property

Public Property Get SmallChange(ByVal eBar As efnScrollBarsIdentificationConstants) As Long
    If (eBar = efnSBIHorizontal) Then
        SmallChange = m_lSmallChangeHorz
    Else
        SmallChange = m_lSmallChangeVert
    End If
End Property

Public Property Let SmallChange(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal lSmallChange As Long)
    If (eBar = efnSBIHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Public Property Get Enabled(ByVal eBar As efnScrollBarsIdentificationConstants) As Boolean
    If (eBar = efnSBIHorizontal) Then
        Enabled = m_bEnabledHorz
    Else
        Enabled = m_bEnabledVert
    End If
End Property

Public Property Let Enabled(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal bEnabled As Boolean)
    Dim lO As Long
    Dim lF As Long
         
    lO = eBar
    If (bEnabled) Then
        lF = ESB_ENABLE_BOTH
    Else
        lF = ESB_DISABLE_BOTH
    End If
    If (m_bNoFlatScrollBars) Then
        EnableScrollBar m_hWnd, lO, lF
    Else
        FlatSB_EnableScrollBar m_hWnd, lO, lF
    End If
End Property

Public Property Get Min(ByVal eBar As efnScrollBarsIdentificationConstants) As Long
    Dim tSI As SCROLLINFO
    
    pGetSI eBar, tSI, SIF_RANGE
    Min = tSI.nMin
End Property

Public Property Get Max(ByVal eBar As efnScrollBarsIdentificationConstants) As Long
    Dim tSI As SCROLLINFO
    
    pGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    Max = tSI.nMax - tSI.nPage
End Property

Public Property Get Value(ByVal eBar As efnScrollBarsIdentificationConstants) As Long
    Dim tSI As SCROLLINFO
    
    pGetSI eBar, tSI, SIF_POS
    Value = tSI.nPos
End Property

Public Property Get LargeChange(ByVal eBar As efnScrollBarsIdentificationConstants) As Long
    Dim tSI As SCROLLINFO
    
    pGetSI eBar, tSI, SIF_PAGE
    LargeChange = tSI.nPage
End Property

Public Property Let Min(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal iMin As Long)
    Dim tSI As SCROLLINFO
    
    tSI.nMin = iMin
    tSI.nMax = Max(eBar) + LargeChange(eBar)
    pLetSI eBar, tSI, SIF_RANGE
End Property

Public Property Let Max(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal iMax As Long)
    Dim tSI As SCROLLINFO
    
    tSI.nMax = iMax + LargeChange(eBar)
    tSI.nMin = Min(eBar)
    pLetSI eBar, tSI, SIF_RANGE
    pRaiseEvent eBar, False
End Property

Public Property Let Value(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal iValue As Long)
    Dim tSI As SCROLLINFO
    
    If (iValue <> Value(eBar)) Then
        tSI.nPos = iValue
        pLetSI eBar, tSI, SIF_POS
        pRaiseEvent eBar, False
    End If
End Property

Public Property Let LargeChange(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal iLargeChange As Long)
    Dim tSI As SCROLLINFO
    
   pGetSI eBar, tSI, SIF_ALL
   tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
   tSI.nPage = iLargeChange
   pLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property

Public Property Get CanBeFlat() As Boolean
   CanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub pCreateScrollBar()
    Dim lR As Long

   ' Just checks for flag scroll bars...
   On Error Resume Next
   lR = InitialiseFlatSB(m_hWnd)
   If (Err.Number <> 0) Then
       'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
       ' Means we have version prior to 4.71
       ' We get standard scroll bars.
       m_bNoFlatScrollBars = True
   Else
      Style = m_eStyle
   End If
End Sub

Public Sub Create(ByVal hWndA As Long)
    pClearUp
    m_hWnd = hWndA
    pCreateScrollBar
    pAttachMessages
End Sub

Private Sub pClearUp()
    If m_hWnd <> 0 Then
        On Error Resume Next
        ' Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
           UninitializeFlatSB m_hWnd
        End If
        
        pvUnsubclass
    End If
    m_hWnd = 0
    m_bInitialised = False
End Sub

Private Sub pAttachMessages()
    If (m_hWnd <> 0) Then
        pvSubclass
        SystemParametersInfo SPI_GETWHEELSCROLLLINES, _
              0, m_lWheelScrollLines, 0
        If (m_lWheelScrollLines <= 0) Then
           m_lWheelScrollLines = 3
        End If
        m_bInitialised = True
    End If
End Sub

Private Sub Class_Initialize()
    m_lSmallChangeHorz = 1
    m_lSmallChangeVert = 1
    m_eStyle = efnSBSRegular
    m_eOrientation = efnSBOBoth
End Sub

Private Sub Class_Terminate()
    pClearUp
End Sub

Private Function pRaiseEvent(ByVal eBar As efnScrollBarsIdentificationConstants, ByVal bScroll As Boolean)
    Static s_lLastValue(0 To 1) As Long
    
    If (Value(eBar) <> s_lLastValue(eBar)) Then
        If (bScroll) Then
            RaiseEvent Scroll(eBar)
        Else
            RaiseEvent Change(eBar)
        End If
        s_lLastValue(eBar) = Value(eBar)
    End If
End Function

'--- for MST subclassing (2)
'Autor: wqweto http://www.vbforums.com/showthread.php?872819
'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================
Private Sub pvSubclass()
    If m_hWnd <> 0 Then
        Set m_pSubclass = InitSubclassingThunk(m_hWnd, InitAddressOfMethod().SubclassProc(0, 0, 0, 0, 0))
    End If
End Sub

Private Sub pvUnsubclass()
    Set m_pSubclass = Nothing
End Sub

Private Function InitAddressOfMethod() As cScrollBars
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(Me), 5, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitSubclassingThunk(ByVal hWnd As Long, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwEDMAV1aLdCQUg8YIgz4AdC+L+oHH/BEzAIvCBQQRMwCri8IFQBEzAKuLwgVQETMAq4vCBXgRMwCruQkAAADzpYHC/BEzAFJqFP9SEFqL+IvCq7gBAAAAq4tEJAyri3QkFKWlg+8UagBX/3IM/3cI/1IYi0QkGIk4Xl+4MBIzAC1wEDMAwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRiLClL/cQz/cgj/URyLVCQEiwpS/1EUM8DCBACQVYvsi1UYiwqLQSyFwHQ1Uv/QWoP4AXdUg/gAdQmBfQwDAgAAdEaLClL/UTBahcB1O4sKUmrw/3Ek/1EoWqkAAAAIdShSM8BQUI1EJARQjUQkBFD/dRT/dRD/dQz/dQj/cgz/UhBZWFqFyXURiwr/dRT/dRD/dQz/dQj/USBdwhgADx8A" ' 29.3.2019 13:04:54
    Const THUNK_SIZE    As Long = 448
    Dim hThunk          As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(Me)
    aParams(1) = pfnCallback
    hThunk = GetProp(pvGetGlobalHwnd(), "InitSubclassingThunk")
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        Call SetProp(pvGetGlobalHwnd(), "InitSubclassingThunk", hThunk)
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

Private Function pvGetGlobalHwnd() As Long
    pvGetGlobalHwnd = FindWindowEx(0, 0, "STATIC", App.hInstance & ":" & App.ThreadID & ":MST Global Data")
    If pvGetGlobalHwnd = 0 Then
        pvGetGlobalHwnd = CreateWindowEx(0, "STATIC", App.hInstance & ":" & App.ThreadID & ":MST Global Data", _
            0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0)
    End If
End Function

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    #If hWnd And wParam And lParam And Handled Then '--- touch args
    #End If
    
    Dim lScrollCode As Long
    Dim tSI As SCROLLINFO
    Dim lV As Long, lSC As Long
    Dim eBar As efnScrollBarsIdentificationConstants
    Dim zDelta As Long
    Dim lDelta As Long
    Dim wMKeyFlags As Long

    Select Case wMsg
    Case WM_MOUSEWHEEL
        ' Low-word of wParam indicates whether virtual keys
        ' are down
        wMKeyFlags = wParam And &HFFFF&
        ' High order word is the distance the wheel has been rotated,
        ' in multiples of WHEEL_DELTA:
        If (wParam And &H8000000) Then
            ' Towards the user:
            zDelta = &H8000& - (wParam And &H7FFF0000) \ &H10000
        Else
            ' Away from the user:
            zDelta = -((wParam And &H7FFF0000) \ &H10000)
        End If
        lDelta = (zDelta \ WHEEL_DELTA) * SmallChange(efnSBIVertical) * m_lWheelScrollLines
        eBar = efnSBIVertical
        RaiseEvent MouseWheel(eBar, lDelta)
        If Not (lDelta = 0) Then
            Value(eBar) = Value(eBar) + lDelta
            SubclassProc = 1
        End If
    
    Case WM_VSCROLL, WM_HSCROLL
        If (wMsg = WM_HSCROLL) Then
            eBar = efnSBIHorizontal
        Else
        eBar = efnSBIVertical
        End If
        lScrollCode = (wParam And &HFFFF&)
        Select Case lScrollCode
        Case SB_THUMBTRACK
            ' Is vertical/horizontal?
            pGetSI eBar, tSI, SIF_TRACKPOS
            Value(eBar) = tSI.nTrackPos
            pRaiseEvent eBar, True
          
        Case SB_Left, SB_TOP
            Value(eBar) = Min(eBar)
            pRaiseEvent eBar, False
          
        Case SB_Right, SB_BOTTOM
            Value(eBar) = Max(eBar)
            pRaiseEvent eBar, False
           
        Case SB_LINELEFT, SB_LINEUP
            'Debug.Print "Line"
            lV = Value(eBar)
            If (eBar = efnSBIHorizontal) Then
                lSC = m_lSmallChangeHorz
            Else
                lSC = m_lSmallChangeVert
            End If
            If (lV - lSC < Min(eBar)) Then
                Value(eBar) = Min(eBar)
            Else
                Value(eBar) = lV - lSC
            End If
            pRaiseEvent eBar, False
          
       Case SB_LINERIGHT, SB_LINEDOWN
             'Debug.Print "Line"
            lV = Value(eBar)
            If (eBar = efnSBIHorizontal) Then
                lSC = m_lSmallChangeHorz
            Else
                lSC = m_lSmallChangeVert
            End If
            If (lV + lSC > Max(eBar)) Then
                Value(eBar) = Max(eBar)
            Else
                Value(eBar) = lV + lSC
            End If
            pRaiseEvent eBar, False
           
       Case SB_PAGELEFT, SB_PAGEUP
            Value(eBar) = Value(eBar) - LargeChange(eBar)
            pRaiseEvent eBar, False
          
       Case SB_PAGERIGHT, SB_PAGEDOWN
            Value(eBar) = Value(eBar) + LargeChange(eBar)
            pRaiseEvent eBar, False
          
       Case SB_ENDSCROLL
            pRaiseEvent eBar, False
          
       End Select
       
    Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN
        Dim eBtn As MouseButtonConstants
        eBtn = IIf(wMsg = WM_NCLBUTTONDOWN, vbLeftButton, vbRightButton)
        If wParam = HTVSCROLL Then
            RaiseEvent ScrollClick(efnSBIHorizontal, eBtn)
        ElseIf wParam = HTHSCROLL Then
            RaiseEvent ScrollClick(efnSBIVertical, eBtn)
        End If
    End Select
    
'    If Not mUserMode Then
'        Handled = True
'        SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
'    End If
End Function
'--- End for MST subclassing (2)






