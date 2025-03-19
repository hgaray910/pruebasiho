Attribute VB_Name = "modDiseñoTabs"

Option Explicit
' *********************************************************************************
'  Types Declarations...
' *********************************************************************************

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GRADIENT_RECT
    UPPERLEFT  As Long       '  //--The GRADIENT_RECT structure specifies the index of two vertices in the pVertex array.
    LOWERRIGHT As Long       '      These two vertices form the upper-left and lower-right boundaries of a rectangle.
End Type

Private Type TRIVERTEX
    X       As Long
    Y       As Long
    Red     As Integer       '   //--The TRIVERTEX structure contains color information and position information.
    Green   As Integer
    Blue    As Integer
    Alpha   As Integer
End Type

Private Type RGB
    R As Integer
    G As Integer             '  //--Selects a red, green, blue (RGB) color based on the arguments supplied
    B As Integer
End Type


' *********************************************************************************
'  API Declarations...
' *********************************************************************************

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long) As Long

' *********************************************************************************
'  Const Declarations...
' *********************************************************************************

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_PAINT    As Long = &HF
Private Const WM_DESTROY  As Long = &H2
                         

' *********************************************************************************
'  Enums Declarations...
' *********************************************************************************

Public Enum TabStyle
       cSolidColor = 0
End Enum

Public Enum Direction
       cHorizontal = 0
       cVertical = 1
End Enum

'------------------------------------
Private DestDC      As Long
Private MaskDC      As Long
Private MemDC       As Long
Private OrigDC      As Long
Private MaskPic     As Long              'Temporary DC
Private MemPic      As Long
Private TempPic     As Long
Private OrigPic     As Long
Private TempDC      As Long
'-------------------------------------

Private origBrush As Long
Private TempBrush As Long
Private origColor As Long 'BackColor

Private oldWndProc As Long   '<----Mem RefPoint

Private Function GetLngColor(Color As Long) As Long
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

Public Sub SetStyle(ByVal hwnd As Long, ByRef Style As TabStyle)

           SetProp hwnd, "MyStyle", Style

End Sub

Private Function GetStyleParams(ByVal hwnd As Long) As TabStyle
        
           GetStyleParams = GetProp(hwnd, "MyStyle")
    
End Function


Private Sub SetHookInstance(ByVal hwnd As Long)

           SetProp hwnd, "Hooked", True

End Sub

Private Function CheckHookInstance(ByVal hwnd As Long) As Boolean
        
           CheckHookInstance = GetProp(hwnd, "Hooked")
    
End Function

Public Sub SetSolidColor(ByVal hwnd As Long, ByVal Color As Long)
           
           SetProp hwnd, "MySolidColor", GetLngColor(Color)
 
End Sub

Private Sub GetSolidColor(ByVal hwnd As Long)

     TempBrush = CreateSolidBrush(GetProp(hwnd, "MySolidColor"))
    
End Sub

Public Sub SSTabSubclass(ByVal hwnd As Long)


'//Check if This Window is Already Subclassed!!

If Not CheckHookInstance(hwnd) And Not RunningInVB Then

    
    SetHookInstance hwnd '//--Tells Control is going to be subclassed
    oldWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf oldSSTabProc)

End If


End Sub


Public Function oldSSTabProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     
       oldSSTabProc = NewSSTabProc(hwnd, uMsg, wParam, lParam)

End Function

Private Function NewSSTabProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     On Error Resume Next

    Dim m_ItemRect As RECT
    Dim m_Width    As Long
    Dim m_Height   As Long

'========================================================================================================
'SUBCLASSING THE SSTab..........

    If wMsg = WM_PAINT Then
    
        '---------------------------------------------------
        
        '//--- Get the SSTab's dimensions
        DestDC = GetDC(hwnd)

        GetWindowRect hwnd, m_ItemRect
                m_Width = m_ItemRect.Right - m_ItemRect.Left
                m_Height = m_ItemRect.Bottom - m_ItemRect.Top
       '---------------------------------------------------

        
      '---------------------------------------------------
        
        '//--- Select The Parameters (SSTab New Style)
        Select Case GetStyleParams(hwnd)
             
         Case cSolidColor
              GetSolidColor hwnd
        Case Else
               Debug.Print "Invalid Style"
         End Select
       '---------------------------------------------------
              
       '----------------------------------------------------------------------------
       '//--- To Work With a Cleaner and Less Flicker Screen Create the Temporary DC
       CreateNewDCWorkArea m_Width, m_Height
       '----------------------------------------------------------------------------
               
       '---------------------------------------------------------------------------
        Call SelectBitmap '//-- Selected Image
       '---------------------------------------------------------------------------
       
       '---------------------------------------------------------------------------
        CallWindowProc oldWndProc, hwnd, wMsg, OrigDC, lParam ' PAINT SSTab in TEMPORARY DC
       '---------------------------------------------------------------------------
        
        '---------------------------------------------------------------------------
        Call CreateBackMask(m_Width, m_Height)  '//-- A Mask For RasterOperations
        '---------------------------------------------------------------------------
        
        
        
        'The PatBlt function paints the given rectangle using the brush that is currently
        'selected into the specified device context.
        'The brush color and the surface color(s) are combined by using the given raster operation.

        '-----------------------------------------------------------------------------------------------------
        origBrush = SelectObject(TempDC, TempBrush)
        
        PatBlt TempDC, 0, 0, m_Width, m_Height, vbPatCopy
        
        SelectObject TempDC, origBrush
        '------------------------------------------------------------------------------------------------------
        
        Call DOBitBlt(m_Width, m_Height) '//--- Do RasterOperations
        Call CleanDCs                    '//--- Free Memory <--Prevent Leaks
             
        
        '-----------------------------------------------------------------------
        SetBkColor DestDC, origColor
        ReleaseDC hwnd, DestDC '//-- Free The DC FROM GetDC API ..AND RETURN THE COLOR BACK TO NORMAL
        ValidateRect hwnd, 0
        '-----------------------------------------------------------------------
       
    ElseIf wMsg = WM_DESTROY Then
        DeleteObject TempBrush
        SetWindowLong hwnd, GWL_WNDPROC, oldWndProc
        NewSSTabProc = CallWindowProc(oldWndProc, hwnd, wMsg, wParam, lParam)
    Else '//-- Other Message I Don't Care ;)
        NewSSTabProc = CallWindowProc(oldWndProc, hwnd, wMsg, wParam, lParam)
    End If
    
      
      
      

End Function

'=======================================================================================================================
' SELECT THE CURRENT IMAGE
'=======================================================================================================================

Private Sub SelectBitmap()
Dim cHandle As Long

       cHandle = SelectObject(MaskDC, MaskPic)
       DeleteObject cHandle
       cHandle = SelectObject(MemDC, MemPic)
       DeleteObject cHandle
       cHandle = SelectObject(TempDC, TempPic)
       DeleteObject cHandle
       cHandle = SelectObject(OrigDC, OrigPic)
       DeleteObject cHandle
       
End Sub

'=======================================================================================================================
' CREATE A MASK COLOR BACKGROUND
'=======================================================================================================================

Private Sub CreateBackMask(ByVal m_Width As Long, ByVal m_Height As Long)
        
        origColor = SetBkColor(DestDC, GetSysColor(15))
        SetBkColor OrigDC, GetSysColor(15)
        BitBlt MaskDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcCopy
       
End Sub


'=======================================================================================================================
' CREATE THE NEW TEMP WORK AREA
'=======================================================================================================================

Private Sub CreateNewDCWorkArea(ByVal m_Width As Long, ByVal m_Height As Long)
        
        MaskDC = CreateCompatibleDC(DestDC)
        MaskPic = CreateBitmap(m_Width, m_Height, 1, 1, ByVal 0&)
        MemDC = CreateCompatibleDC(DestDC)
        MemPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        TempDC = CreateCompatibleDC(DestDC)
        TempPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        OrigDC = CreateCompatibleDC(DestDC)
        OrigPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)

End Sub


'=======================================================================================================================
' BITBLT  RasterOperations
'=======================================================================================================================

Private Sub DOBitBlt(ByVal m_Width As Long, ByVal m_Height As Long)
        
        BitBlt MemDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbSrcCopy
        BitBlt MemDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcPaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbMergePaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, MemDC, 0, 0, vbSrcAnd
        BitBlt DestDC, 0, 0, m_Width, m_Height, TempDC, 0, 0, vbSrcCopy

End Sub

'=======================================================================================================================
' CLEAN UP MEMORY
'=======================================================================================================================

Private Sub CleanDCs()
        
        DeleteDC TempDC
        DeleteObject TempPic
        DeleteDC MaskDC
        DeleteObject MaskPic
        DeleteDC MemDC
        DeleteObject MemPic
        DeleteDC OrigDC
        DeleteObject OrigPic
        DeleteObject TempBrush

End Sub

Function RunningInVB() As Boolean
    'Returns whether we are running in vb(true), or compiled (false)
     
        Static counter As Variant
        If IsEmpty(counter) Then
            counter = 1
            Debug.Assert RunningInVB() Or True
            counter = counter - 1
        ElseIf counter = 1 Then
            counter = 0
        End If
        RunningInVB = counter
     
End Function
