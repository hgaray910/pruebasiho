VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmReqPendientesGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisiciones pendientes"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VSFlex7LCtl.VSFlexGrid grdPendientes 
      Height          =   4000
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      _cx             =   16748
      _cy             =   7056
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
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
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmReqPendientesGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Me.Hide
    
End Sub
Private Sub pCarga()
    Dim vlstrSentencia As String
    Dim rsRequisiciones As New ADODB.Recordset
    Dim rsGrupoValidaReqYAuto As New ADODB.Recordset
    Dim i As Integer
    grdPendientes.Rows = 1
    
    vlstrSentencia = "SELECT intmovpaciente, chrtipopaciente " & _
                                         " FROM PVDETALLEFACTURACONSOLID " & _
                                         " WHERE PVDETALLEFACTURACONSOLID.intCveGrupo = " & frmFacturacion.txtMovimientoPaciente.Text
    Set rsGrupoValidaReqYAuto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    rsGrupoValidaReqYAuto.MoveFirst
    
    Do Until rsGrupoValidaReqYAuto.EOF
        vgstrParametrosSP = Str(rsGrupoValidaReqYAuto!INTMOVPACIENTE) & "|" & rsGrupoValidaReqYAuto!CHRTIPOPACIENTE
        Set rsRequisiciones = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelRequisicionesPendietes")
        
        If rsRequisiciones.RecordCount <> 0 Then
            For i = 1 To rsRequisiciones.RecordCount ''*
                grdPendientes.AddItem ""
                grdPendientes.TextMatrix(grdPendientes.Rows - 1, 0) = rsRequisiciones!almacen
                grdPendientes.TextMatrix(grdPendientes.Rows - 1, 1) = rsRequisiciones!numNumRequisicion
                grdPendientes.TextMatrix(grdPendientes.Rows - 1, 2) = rsRequisiciones!dtmFechaRequisicion
                grdPendientes.TextMatrix(grdPendientes.Rows - 1, 3) = rsRequisiciones!SMICVEDEPTOREQUISDESC
                grdPendientes.TextMatrix(grdPendientes.Rows - 1, 4) = rsRequisiciones!Nombre
                rsRequisiciones.MoveNext
            Next
         End If
         
        rsGrupoValidaReqYAuto.MoveNext
    Loop
            
    rsGrupoValidaReqYAuto.Close
    rsRequisiciones.Close
    
    
End Sub

Private Sub Form_Activate()

    grdPendientes.SetFocus
    Form_Load
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    grdPendientes.ColWidth(0) = 2250
    grdPendientes.ColWidth(1) = 1750
    grdPendientes.ColWidth(2) = 1250
    grdPendientes.ColWidth(3) = 3500
    grdPendientes.ColWidth(4) = 4000
    
    grdPendientes.ColAlignment(0) = flexAlignLeftCenter
    grdPendientes.ColAlignment(1) = flexAlignLeftCenter
    grdPendientes.ColAlignment(2) = flexAlignLeftCenter
    grdPendientes.ColAlignment(3) = flexAlignLeftCenter
    grdPendientes.ColAlignment(4) = flexAlignLeftCenter
    grdPendientes.TextMatrix(0, 0) = "Almacén"
    grdPendientes.TextMatrix(0, 1) = "Número de requisición"
    grdPendientes.TextMatrix(0, 2) = "Fecha"
    grdPendientes.TextMatrix(0, 3) = "Departamento que solicitó"
    grdPendientes.TextMatrix(0, 4) = "Paciente"
    
    
    pCarga
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        
    End If
End Sub



Private Sub grdPendientes_Click()

End Sub
