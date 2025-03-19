VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmCostoCargos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   4950
      TabIndex        =   6
      Top             =   6370
      Width           =   720
      Begin MyCommandButton.MyButton cmdSave 
         Height          =   600
         Left            =   60
         TabIndex        =   7
         ToolTipText     =   "Enviar transferencia"
         Top             =   130
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   16777215
         Picture         =   "frmCostoCargos.frx":0000
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmCostoCargos.frx":0984
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
   End
   Begin HSFlatControls.MyCombo cboEmpresaContable 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      Style           =   1
      Enabled         =   -1  'True
      Text            =   "MyCombo1"
      Sorted          =   0   'False
      List            =   ""
      ItemData        =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HSFlatControls.MyTabHeader MyTabHeader1 
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   741
      Tabs            =   4
      TabCurrent      =   0
      TabWidth        =   2650
      Caption         =   $"frmCostoCargos.frx":1308
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1320
      TabIndex        =   9
      Top             =   7560
      Width           =   1095
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Siguiente"
         Default         =   -1  'True
         Height          =   255
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Otros conceptos de cargo"
      TabPicture(0)   =   "frmCostoCargos.frx":1342
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdOtrosConceptos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paquetes"
      TabPicture(1)   =   "frmCostoCargos.frx":135E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdPaquetes"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Estudios"
      TabPicture(2)   =   "frmCostoCargos.frx":137A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdEstudios"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Exámenes"
      TabPicture(3)   =   "frmCostoCargos.frx":1396
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdExamenes"
      Tab(3).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid grdOtrosConceptos 
         Height          =   5010
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Notas urgentes"
         Top             =   570
         Width           =   10305
         _cx             =   18177
         _cy             =   8837
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCostoCargos.frx":13B2
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPaquetes 
         Height          =   5010
         Left            =   -74880
         TabIndex        =   3
         ToolTipText     =   "Notas urgentes"
         Top             =   570
         Width           =   10305
         _cx             =   18177
         _cy             =   8837
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCostoCargos.frx":143F
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdEstudios 
         Height          =   5010
         Left            =   -74880
         TabIndex        =   4
         ToolTipText     =   "Notas urgentes"
         Top             =   570
         Width           =   10305
         _cx             =   18177
         _cy             =   8837
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCostoCargos.frx":14CC
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdExamenes 
         Height          =   5010
         Left            =   -74880
         TabIndex        =   5
         ToolTipText     =   "Notas urgentes"
         Top             =   570
         Width           =   10305
         _cx             =   18177
         _cy             =   8837
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCostoCargos.frx":1559
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Empresa contable"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   180
      Width           =   2055
   End
End
Attribute VB_Name = "frmCostoCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lstrModulo As String

Private Sub cboEmpresaContable_Click()
    If cboEmpresaContable.ListIndex > -1 Then pCargaCostos
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    pGuardarDatos
End Sub

Private Sub cmdSiguiente_Click()
    SendKeys vbTab
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    Me.Caption = IIf(lstrModulo = "PV", "Costos por conceptos de cargo y paquetes", IIf(lstrModulo = "IM", "Costos de estudios", IIf(lstrModulo = "LA", "Costos de exámenes", "")))
    SSTab1.TabVisible(0) = IIf(lstrModulo = "PV", True, False)
    MyTabHeader1.TabEnabled(0) = IIf(lstrModulo = "PV", True, False)
    SSTab1.TabVisible(1) = IIf(lstrModulo = "PV", True, False)
    MyTabHeader1.TabEnabled(1) = IIf(lstrModulo = "PV", True, False)
    SSTab1.TabVisible(2) = IIf(lstrModulo = "IM", True, False)
    MyTabHeader1.TabEnabled(2) = IIf(lstrModulo = "IM", True, False)
    SSTab1.TabVisible(3) = IIf(lstrModulo = "LA", True, False)
    MyTabHeader1.TabEnabled(3) = IIf(lstrModulo = "LA", True, False)
    
    pCargaCombos
    SSTab1_Click -1
    If cgstrModulo <> "SI" Then cboEmpresaContable.Enabled = False
    
    
    'Color de Tab
    SetStyle SSTab1.hWnd, 0
    SetSolidColor SSTab1.hWnd, 16777215
    SSTabSubclass SSTab1.hWnd
    
    MyTabHeader1.TabCurrent = SSTab1.Tab
    
    
    
End Sub

Private Sub pCargaCombos()
    Dim rs As ADODB.Recordset
    
    Set rs = frsRegresaRs("SELECT * FROM CNEMPRESACONTABLE ORDER BY vchNombre")
    If Not rs.EOF Then
        pLlenarCboRs_new cboEmpresaContable, rs, 0, 1
    End If
    rs.Close
    cboEmpresaContable.ListIndex = fintLocalizaCbo_new(cboEmpresaContable, CStr(vgintClaveEmpresaContable))
End Sub

Private Sub pCargaCostos()
    Dim rs As ADODB.Recordset
    Dim intRow As Integer
    Dim lngCveEmpresa As Long
    
    lngCveEmpresa = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
    grdOtrosConceptos.Rows = 1
    grdPaquetes.Rows = 1
    grdEstudios.Rows = 1
    grdExamenes.Rows = 1
    Set rs = frsEjecuta_SP(lngCveEmpresa & "|" & lstrModulo, "sp_PVSelCostoCargos")
    Do Until rs.EOF
        If rs!Tipo = "OC" Then
            grdOtrosConceptos.AddItem ""
            intRow = grdOtrosConceptos.Rows - 1
            grdOtrosConceptos.TextMatrix(intRow, 1) = rs!Clave
            grdOtrosConceptos.TextMatrix(intRow, 2) = Trim(rs!descripcion)
            grdOtrosConceptos.TextMatrix(intRow, 3) = IIf(IsNull(rs!costo), "", FormatCurrency(rs!costo, 2))
        End If
        If rs!Tipo = "PA" Then
            grdPaquetes.AddItem ""
            intRow = grdPaquetes.Rows - 1
            grdPaquetes.TextMatrix(intRow, 1) = rs!Clave
            grdPaquetes.TextMatrix(intRow, 2) = Trim(rs!descripcion)
            grdPaquetes.TextMatrix(intRow, 3) = IIf(IsNull(rs!costo), "", FormatCurrency(rs!costo, 2))
        End If
        If rs!Tipo = "ES" Then
            grdEstudios.AddItem ""
            intRow = grdEstudios.Rows - 1
            grdEstudios.TextMatrix(intRow, 1) = rs!Clave
            grdEstudios.TextMatrix(intRow, 2) = Trim(rs!descripcion)
            grdEstudios.TextMatrix(intRow, 3) = IIf(IsNull(rs!costo), "", FormatCurrency(rs!costo, 2))
        End If
        If rs!Tipo = "EX" Then
            grdExamenes.AddItem ""
            intRow = grdExamenes.Rows - 1
            grdExamenes.TextMatrix(intRow, 1) = rs!Clave
            grdExamenes.TextMatrix(intRow, 2) = Trim(rs!descripcion)
            grdExamenes.TextMatrix(intRow, 3) = IIf(IsNull(rs!costo), "", FormatCurrency(rs!costo, 2))
            grdExamenes.TextMatrix(intRow, 5) = "EX"
        End If
        If rs!Tipo = "GE" Then
            grdExamenes.AddItem ""
            intRow = grdExamenes.Rows - 1
            grdExamenes.TextMatrix(intRow, 1) = rs!Clave
            grdExamenes.TextMatrix(intRow, 2) = Trim(rs!descripcion)
            grdExamenes.TextMatrix(intRow, 3) = IIf(IsNull(rs!costo), "", FormatCurrency(rs!costo, 2))
            grdExamenes.TextMatrix(intRow, 5) = "GE"
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub pDecimal(ByRef KeyAscii As Integer, strText As String)
    If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
    If Chr(KeyAscii) = "." Then
        If InStr(1, strText, ".") > 0 Then KeyAscii = 0
    Else
        If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
    End If
End Sub

Private Sub pGuardarDatos()
    Dim lngIndex As Long
    Dim strParametros As String
    Dim strBorrar As String
    Dim lngCveEmpresa As Long
    Dim llngPersonaGraba As Long
    
    If cboEmpresaContable.ListIndex > -1 Then
        lngCveEmpresa = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
        'Otros conceptos
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If llngPersonaGraba <> 0 Then
            EntornoSiho.ConeccionSiho.BeginTrans
            For lngIndex = 1 To grdOtrosConceptos.Rows - 1
                If grdOtrosConceptos.TextMatrix(lngIndex, 4) = "SI" Then
                    strBorrar = IIf(grdOtrosConceptos.TextMatrix(lngIndex, 3) = "", "1", "0")
                    strParametros = lngCveEmpresa & "|OC|" & grdOtrosConceptos.TextMatrix(lngIndex, 1) & "|" & Format(grdOtrosConceptos.TextMatrix(lngIndex, 3), "0.00") & "|" & strBorrar
                    frsEjecuta_SP strParametros, "sp_PVActualizaCosto", True
                    grdOtrosConceptos.TextMatrix(lngIndex, 4) = ""
                    frsEjecuta_SP lngCveEmpresa & "|" & grdOtrosConceptos.TextMatrix(lngIndex, 1) & "|OC", "sp_GNUpdActualizaListasPrecio"
                    pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "COSTOS OTROS CONCEPTOS", grdOtrosConceptos.TextMatrix(lngIndex, 1) & "-" & Format(grdOtrosConceptos.TextMatrix(lngIndex, 3), "0.00")
                End If
            Next
            'Paquetes
            For lngIndex = 1 To grdPaquetes.Rows - 1
                If grdPaquetes.TextMatrix(lngIndex, 4) = "SI" Then
                    strBorrar = IIf(grdPaquetes.TextMatrix(lngIndex, 3) = "", "1", "0")
                    strParametros = lngCveEmpresa & "|PA|" & grdPaquetes.TextMatrix(lngIndex, 1) & "|" & Format(grdPaquetes.TextMatrix(lngIndex, 3), "0.00") & "|" & strBorrar
                    frsEjecuta_SP strParametros, "sp_PVActualizaCosto", True
                    grdPaquetes.TextMatrix(lngIndex, 4) = ""
                    frsEjecuta_SP lngCveEmpresa & "|" & grdPaquetes.TextMatrix(lngIndex, 1) & "|PA", "sp_GNUpdActualizaListasPrecio"
                    pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "COSTOS PAQUETES", grdPaquetes.TextMatrix(lngIndex, 1) & "-" & Format(grdPaquetes.TextMatrix(lngIndex, 3), "0.00")
                End If
            Next
            'Estudios
            For lngIndex = 1 To grdEstudios.Rows - 1
                If grdEstudios.TextMatrix(lngIndex, 4) = "SI" Then
                    strBorrar = IIf(grdEstudios.TextMatrix(lngIndex, 3) = "", "1", "0")
                    strParametros = lngCveEmpresa & "|ES|" & grdEstudios.TextMatrix(lngIndex, 1) & "|" & Format(grdEstudios.TextMatrix(lngIndex, 3), "0.00") & "|" & strBorrar
                    frsEjecuta_SP strParametros, "sp_PVActualizaCosto", True
                    grdEstudios.TextMatrix(lngIndex, 4) = ""
                    frsEjecuta_SP lngCveEmpresa & "|" & grdEstudios.TextMatrix(lngIndex, 1) & "|ES", "sp_GNUpdActualizaListasPrecio"
                    pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "COSTOS ESTUDIOS", grdEstudios.TextMatrix(lngIndex, 1) & "-" & Format(grdEstudios.TextMatrix(lngIndex, 3), "0.00")
                End If
            Next
            'Examanes
            For lngIndex = 1 To grdExamenes.Rows - 1
                If grdExamenes.TextMatrix(lngIndex, 4) = "SI" Then
                    strBorrar = IIf(grdExamenes.TextMatrix(lngIndex, 3) = "", "1", "0")
                    strParametros = lngCveEmpresa & "|" & grdExamenes.TextMatrix(lngIndex, 5) & "|" & grdExamenes.TextMatrix(lngIndex, 1) & "|" & Format(grdExamenes.TextMatrix(lngIndex, 3), "0.00") & "|" & strBorrar
                    frsEjecuta_SP strParametros, "sp_PVActualizaCosto", True
                    grdExamenes.TextMatrix(lngIndex, 4) = ""
                    frsEjecuta_SP lngCveEmpresa & "|" & grdExamenes.TextMatrix(lngIndex, 1) & "|" & grdExamenes.TextMatrix(lngIndex, 5), "sp_GNUpdActualizaListasPrecio"
                    pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "COSTOS EXAMENES", grdExamenes.TextMatrix(lngIndex, 1) & "-" & Format(grdExamenes.TextMatrix(lngIndex, 3), "0.00")
                End If
            Next
            EntornoSiho.ConeccionSiho.CommitTrans
            MsgBox SIHOMsg(358), vbInformation, "Mensaje"
        End If
    End If
End Sub

Private Sub grdEstudios_GotFocus()
    If grdEstudios.Rows > 1 Then
        If grdEstudios.Row < 1 Then
            grdEstudios.Row = 1
            grdEstudios.Col = 3
        End If
    End If
End Sub

Private Sub grdExamenes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If IsNumeric(grdExamenes.TextMatrix(Row, Col)) Then
        grdExamenes.TextMatrix(Row, Col) = FormatCurrency(grdExamenes.TextMatrix(Row, Col), 2)
    Else
        grdExamenes.TextMatrix(Row, Col) = ""
    End If
End Sub

Private Sub grdExamenes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 Then Cancel = True
End Sub

Private Sub grdExamenes_ChangeEdit()
    grdExamenes.TextMatrix(grdExamenes.Row, 4) = "SI"
End Sub

Private Sub grdExamenes_GotFocus()
    If grdExamenes.Rows > 1 Then
        If grdExamenes.Row < 1 Then
            grdExamenes.Row = 1
            grdExamenes.Col = 3
        End If
    End If
End Sub

Private Sub grdExamenes_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    pDecimal KeyAscii, grdExamenes.EditText
End Sub

Private Sub grdOtrosConceptos_GotFocus()
    If grdOtrosConceptos.Rows > 1 Then
        If grdOtrosConceptos.Row < 1 Then
            grdOtrosConceptos.Row = 1
            grdOtrosConceptos.Col = 3
        End If
    End If
End Sub

Private Sub grdPaquetes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If IsNumeric(grdPaquetes.TextMatrix(Row, Col)) Then
        grdPaquetes.TextMatrix(Row, Col) = FormatCurrency(grdPaquetes.TextMatrix(Row, Col), 2)
    Else
        grdPaquetes.TextMatrix(Row, Col) = ""
    End If
End Sub

Private Sub grdPaquetes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 Then Cancel = True
End Sub

Private Sub grdPaquetes_ChangeEdit()
    grdPaquetes.TextMatrix(grdPaquetes.Row, 4) = "SI"
End Sub

Private Sub grdPaquetes_GotFocus()
    If grdPaquetes.Rows > 1 Then
        If grdPaquetes.Row < 1 Then
            grdPaquetes.Row = 1
            grdPaquetes.Col = 3
        End If
    End If
End Sub

Private Sub grdPaquetes_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    pDecimal KeyAscii, grdPaquetes.EditText
End Sub

Private Sub grdOtrosConceptos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If IsNumeric(grdOtrosConceptos.TextMatrix(Row, Col)) Then
        grdOtrosConceptos.TextMatrix(Row, Col) = FormatCurrency(grdOtrosConceptos.TextMatrix(Row, Col), 2)
    Else
        grdOtrosConceptos.TextMatrix(Row, Col) = ""
    End If
End Sub

Private Sub grdOtrosConceptos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 Then Cancel = True
End Sub

Private Sub grdOtrosConceptos_ChangeEdit()
    grdOtrosConceptos.TextMatrix(grdOtrosConceptos.Row, 4) = "SI"
End Sub

Private Sub grdOtrosConceptos_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    pDecimal KeyAscii, grdOtrosConceptos.EditText
End Sub

Private Sub grdEstudios_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If IsNumeric(grdEstudios.TextMatrix(Row, Col)) Then
        grdEstudios.TextMatrix(Row, Col) = Format(grdEstudios.TextMatrix(Row, Col), "0.00")
    Else
        grdEstudios.TextMatrix(Row, Col) = ""
    End If
End Sub

Private Sub grdEstudios_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 Then Cancel = True
End Sub

Private Sub grdEstudios_ChangeEdit()
    grdEstudios.TextMatrix(grdEstudios.Row, 4) = "SI"
End Sub

Private Sub grdEstudios_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    pDecimal KeyAscii, grdEstudios.EditText
End Sub

Private Sub MyTabHeader1_Click(Index As Integer)
    grdOtrosConceptos.Enabled = IIf(MyTabHeader1.TabCurrent = 0, True, False)
    grdPaquetes.Enabled = IIf(MyTabHeader1.TabCurrent = 1, True, False)
    grdEstudios.Enabled = IIf(MyTabHeader1.TabCurrent = 2, True, False)
    grdExamenes.Enabled = IIf(MyTabHeader1.TabCurrent = 3, True, False)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    grdOtrosConceptos.Enabled = IIf(SSTab1.Tab = 0, True, False)
    grdPaquetes.Enabled = IIf(SSTab1.Tab = 1, True, False)
    grdEstudios.Enabled = IIf(SSTab1.Tab = 2, True, False)
    grdExamenes.Enabled = IIf(SSTab1.Tab = 3, True, False)
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
