VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRequisicionesPendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSi 
      Caption         =   "Sí"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2200
      Width           =   1095
   End
   Begin VSFlex7LCtl.VSFlexGrid grdRequisiciones 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Requisiciones pendientes"
      Top             =   500
      Width           =   5415
      _cx             =   9551
      _cy             =   2566
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoResize      =   0   'False
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
   Begin VB.Label lblMensajeNoContinua 
      Caption         =   "No se puede cerrar la cuenta, existen requisiciones pendientes de surtir para este paciente."
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   80
      Width           =   5490
   End
   Begin VB.Label lblContinuar 
      AutoSize        =   -1  'True
      Caption         =   "¿Desea continuar?"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2030
      Width           =   1350
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      Caption         =   "Existen requisiciones pendientes de surtir para este paciente."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   4305
   End
End
Attribute VB_Name = "frmRequisicionesPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lblnPermiteContinuar As Boolean
Public lblnContinuarCerrarCuenta As Boolean
Private Sub cmdAceptar_Click()
    lblnContinuarCerrarCuenta = False
    Me.Hide
End Sub

Private Sub cmdNo_Click()
    lblnContinuarCerrarCuenta = False
    Me.Hide
End Sub

Private Sub cmdSi_Click()
    lblnContinuarCerrarCuenta = True
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = 27 Then
        Me.Hide
    End If
    KeyAscii = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Me.Icon = frmMenuPrincipal.Icon
    pConfiguraGrid
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub
Private Sub pConfiguraGrid()
On Error GoTo NotificaError
    With grdRequisiciones
        .Clear
        .Cols = 4
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Número|Fecha|Estado"
        .ColWidth(0) = 100
        .ColWidth(1) = 1120
        .ColWidth(2) = 1830
        .ColWidth(3) = 2030
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrid"))
End Sub

Public Sub pMostrarRequisiciones(llngCuenta As Long, lstrTipo As String, lblnContinuar As Boolean)
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    lblnContinuarCerrarCuenta = False
    lblnPermiteContinuar = lblnContinuar
        
    vgstrParametrosSP = Str(llngCuenta) & "|" & lstrTipo
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelRequisicionesPendietes")
    grdRequisiciones.Rows = 1
    Do Until rs.EOF
        grdRequisiciones.AddItem ""
        grdRequisiciones.TextMatrix(grdRequisiciones.Rows - 1, 1) = rs!numnumRequisicion
        grdRequisiciones.TextMatrix(grdRequisiciones.Rows - 1, 2) = Format(rs!dtmFechaRequisicion, "dd/MMM/yyyy") & " " & Format(rs!dtmHoraRequisicion, "hh:mm")
        grdRequisiciones.TextMatrix(grdRequisiciones.Rows - 1, 3) = rs!vchEstatusRequis
        rs.MoveNext
    Loop
    rs.Close
    
    cmdAceptar.Visible = Not lblnPermiteContinuar
    lblMensajeNoContinua.Visible = Not lblnPermiteContinuar
    lblMensaje.Visible = lblnPermiteContinuar
    lblContinuar.Visible = lblnPermiteContinuar
    cmdSi.Visible = lblnPermiteContinuar
    cmdNo.Visible = lblnPermiteContinuar
    
    Me.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMostrarRequisiciones"))
End Sub
