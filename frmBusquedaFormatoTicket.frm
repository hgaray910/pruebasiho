VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmBusquedaFormatoTicket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de formatos de ticket"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid grdBusqueda 
      Align           =   2  'Align Bottom
      Height          =   5625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7110
      _cx             =   12541
      _cy             =   9922
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      BackColorSel    =   -2147483645
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   14737632
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBusquedaFormatoTicket.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   0
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
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
Attribute VB_Name = "frmBusquedaFormatoTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vgblnBuscar As Boolean '[  Indica si se seleccionó un formato o se dio escape  ]

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim vlrsFormatos As New ADODB.recordSet
    Dim vlstrSentencia As String
    
    vgblnBuscar = False
    vlstrSentencia = "Select pvFormatoTicket.INTCVEFORMATOTICKET, pvFormatoTicket.VCHDESCRIPCION From pvFormatoTicket"
    Set vlrsFormatos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    With grdBusqueda
        .Clear
        .Rows = 1
        .TextMatrix(0, 0) = "Clave"
        .TextMatrix(0, 1) = "Descripción"
        Do While Not vlrsFormatos.EOF
            .Rows = grdBusqueda.Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, 0) = vlrsFormatos!intCveFormatoTicket
            .TextMatrix(.Row, 1) = vlrsFormatos!VCHDESCRIPCION
            vlrsFormatos.MoveNext
        Loop
        .Row = 1
    End With
End Sub

Private Sub grdBusqueda_DblClick()
    vgblnBuscar = True
    Me.Visible = False
End Sub

Private Sub grdBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdBusqueda_DblClick
    End If
End Sub
