VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBusquedaGrupos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de grupos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   5520
      TabIndex        =   21
      Top             =   1680
      Width           =   735
      Begin VB.CommandButton cmdBuscaGrupos 
         Height          =   495
         Left            =   90
         Picture         =   "frmBusquedaGrupos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Búsqueda de pacientes"
         Top             =   320
         UseMaskColor    =   -1  'True
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   3000
      TabIndex        =   17
      Top             =   1680
      Width           =   2415
      Begin VB.CheckBox chkFacturados 
         Caption         =   "Incluye Facturados"
         Height          =   375
         Left            =   160
         TabIndex        =   20
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtFolioFacturaB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Folio factura"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   870
      End
   End
   Begin VB.Frame fraOpcional 
      Caption         =   "Tipo de paciente"
      Height          =   680
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   6135
      Begin VB.TextBox txtCtaPacB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4920
         TabIndex        =   15
         Top             =   220
         Width           =   1095
      End
      Begin VB.OptionButton optTipoPacB 
         Caption         =   "Internos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   280
         Width           =   975
      End
      Begin VB.OptionButton optTipoPacB 
         Caption         =   "Externos"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   280
         Width           =   975
      End
      Begin VB.OptionButton optTipoPacB 
         Caption         =   "Ambos"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   12
         Top             =   280
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta del paciente"
         Height          =   195
         Left            =   3240
         TabIndex        =   16
         Top             =   285
         Width           =   1425
      End
   End
   Begin VB.Frame Frame10 
      Height          =   930
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   5775
      End
      Begin VB.OptionButton optTipoGrupoB 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   210
         Width           =   1095
      End
      Begin VB.OptionButton optTipoGrupoB 
         Caption         =   "Particular"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton optTipoGrupoB 
         Caption         =   "Todos"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   0
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1050
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
      Begin VB.CheckBox chkRangoFechasB 
         Caption         =   "Rango de fechas"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpFechaInicialB 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   280
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   21299201
         CurrentDate     =   38147
      End
      Begin MSComCtl2.DTPicker dtpFechaFinalB 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   640
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   21299201
         CurrentDate     =   38147
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   340
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Fecha final"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   700
         Width           =   780
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid MSFGResultado 
      Height          =   3255
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   6135
      _cx             =   10821
      _cy             =   5741
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   0
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
Attribute VB_Name = "frmBusquedaGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vgintCveGrupo As Integer
Public vgblnFormaBusqueda As Boolean
Private vgintOrden As Integer

Private Sub chkRangoFechasB_Click()
    If chkRangoFechasB.Value Then
        dtpFechaInicialB.Enabled = True
        dtpFechaFinalB.Enabled = True
        dtpFechaInicialB.SetFocus
    Else
        dtpFechaInicialB.Enabled = False
        dtpFechaFinalB.Enabled = False
    End If
End Sub

Private Sub cmdBuscaGrupos_Click()
    Dim strParametros As String
    Dim rsGrupos As New ADODB.Recordset
    
    strParametros = IIf(optTipoGrupoB(0).Value, IIf(cboEmpresa.ListIndex = 0, "-2", cboEmpresa.ItemData(cboEmpresa.ListIndex)), IIf(optTipoGrupoB(1).Value, "-1", "-3")) & "|" & _
                    IIf(chkRangoFechasB.Value, fstrFechaSQL(dtpFechaInicialB.Value, "00:00:00"), "") & "|" & _
                    fstrFechaSQL(dtpFechaFinalB.Value, "23:59:59") & "|" & _
                    IIf(txtFolioFacturaB.Text <> "", "1", IIf(chkFacturados.Value, "1", "0")) & "|" & _
                    IIf(optTipoPacB(2).Value, "-1", txtCtaPacB.Text) & "|" & _
                    IIf(optTipoPacB(0).Value, "I", "E") & "|" & _
                    IIf(txtFolioFacturaB.Text = "", "-1", txtFolioFacturaB.Text) & "|" & _
                    vgintClaveEmpresaContable
    Set rsGrupos = frsEjecuta_SP(strParametros, "Sp_Pvselconsultagrupos")
    MSFGResultado.Clear
    If rsGrupos.RecordCount > 0 Then
        pLlenaVsfGrid MSFGResultado, rsGrupos
        pConfiguraGridResultadoBusqueda False
    Else
        pConfiguraGridResultadoBusqueda True
    End If

End Sub

Private Sub pConfiguraGridResultadoBusqueda(pblnInicializa As Boolean)
    With MSFGResultado
        .Redraw = False
        .Cols = 4
        If pblnInicializa Then .Rows = 2
        .FixedCols = 0
        .FixedRows = 1
        .ColWidth(0) = 800
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .TextMatrix(0, 0) = "Clave"
        .TextMatrix(0, 1) = "Creación"
        .TextMatrix(0, 2) = "Factura"
        .TextMatrix(0, 3) = "Empresa"
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .Redraw = True
    End With
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = 27 Then
        Me.Hide
        vgintCveGrupo = -1
    End If
End Sub

Private Sub Form_Load()
    Dim rsEmpresas As New ADODB.Recordset
    Dim StrSQL As String
    
    Me.Icon = frmMenuPrincipal.Icon
    '----------------------------------------------------------------
    '|  Valida que se haya establecido un concepto de liquidación
    '|  para que se puedan crear grupos de particulares.
    '----------------------------------------------------------------
    StrSQL = "Select COUNT(*) Co From PVCONCEPTOPAGOempresa Where BITCONCEPTOLIQUIDACION = 1 and intcveempresa = " & vgintClaveEmpresaContable
    If frsRegresaRs(StrSQL, adLockReadOnly, adOpenForwardOnly)!CO = 0 Then
        optTipoGrupoB(1).Enabled = False
        optTipoGrupoB(0).Value = True
        optTipoGrupoB(1).Enabled = False
        optTipoGrupoB(2).Enabled = False
    End If
    '|  Carga combo con las empresas que tengan grupos de facturas
    StrSQL = "SELECT distinct CcEmpresa.INTCVEEMPRESA as Clave " & _
             "     , CcEmpresa.VCHDESCRIPCION as Nombre " & _
             "  FROM PvFacturacionConsolidada " & _
             "       INNER JOIN CcEmpresa ON (PvFacturacionConsolidada.INTCVEEMPRESA = CcEmpresa.INTCVEEMPRESA) " & _
             " WHERE PvFacturacionConsolidada.chrFolioFactura is null"
    Set rsEmpresas = frsRegresaRs(StrSQL, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rsEmpresas, 0, 1, 3
    '| Inicializa fechas
    dtpFechaFinalB.Value = fdtmServerFecha
    dtpFechaInicialB.Value = fdtmServerFecha
    
    If vgblnFormaBusqueda Then
        MSFGResultado.Top = 2790
        MSFGResultado.Height = 2595
        fraOpcional.Enabled = True
    Else
        MSFGResultado.Top = 1660
        MSFGResultado.Height = 3675
        fraOpcional.Enabled = False
    End If
    cboEmpresa.ListIndex = 0
    
    'pLimpiaMshFGrid MSFGResultado
    pConfiguraGridResultadoBusqueda True

    pConfiguraGridResultadoBusqueda True
    vgintOrden = 2
    vgintCveGrupo = -1
    rsEmpresas.Close
    
    
End Sub

Private Sub MSFGResultado_DblClick()
    'Evalúa si escogió un grupo
    If MSFGResultado.Row > 0 Then
        Me.Hide
        If MSFGResultado.TextMatrix(MSFGResultado.Row, 1) <> "" Then
            vgintCveGrupo = MSFGResultado.TextMatrix(MSFGResultado.Row, 0)
        End If
        
    End If
End Sub

Private Sub optTipoGrupoB_Click(Index As Integer)
    Select Case Index
        Case 0
            cboEmpresa.Enabled = True
            pEnfocaCbo cboEmpresa
        Case 1, 2
            cboEmpresa.Enabled = False
    End Select
End Sub

Private Sub optTipoPacB_Click(Index As Integer)
    Select Case Index
        Case 0, 1
            txtCtaPacB.Enabled = True
            txtCtaPacB.SetFocus
        Case 2
            txtCtaPacB.Text = ""
            txtCtaPacB.Enabled = False
    End Select
End Sub

Private Sub txtCtaPacB_GotFocus()
    pEnfocaTextBox txtCtaPacB
End Sub

Private Sub txtCtaPacB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If RTrim(txtCtaPacB.Text) = "" Then
            With FrmBusquedaPacientes
                If optTipoPacB(1).Value Then 'Externos
                    .txtBusqueda.CausesValidation = False
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSoloActivos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optTodos.Enabled = True
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                ElseIf optTipoPacB(0).Value Then
                    .txtBusqueda.CausesValidation = False
                    .vgstrTipoPaciente = "I"  'Internos
                    .vgblnPideClave = False
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                
                txtCtaPacB.Text = .flngRegresaPaciente()
                
                If txtCtaPacB.Text <> -1 Then
                    txtCtaPacB_KeyDown vbKeyReturn, 0
                Else
                    txtCtaPacB.Text = ""
                End If
                .txtBusqueda.CausesValidation = True
            End With
        End If
    End If
End Sub

Private Sub txtCtaPacB_KeyPress(KeyAscii As Integer)
    '|  Si la tecla presionada no es un número
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        '|  Valida si la tecla presionada es una "E" o "P" para cambiar el tipo de grupo
        Select Case UCase(Chr(KeyAscii))
            Case "E"
                optTipoPacB(1).Value = True
            Case "I"
                optTipoGrupoB(0).Value = True
            Case "A"
                optTipoGrupoB(2).Value = True
        End Select
        KeyAscii = 7
    End If
End Sub

Private Sub txtCtaPacB_LostFocus()
    If txtCtaPacB.Text = "" Then optTipoPacB(2).Value = True
End Sub

Private Sub txtFolioFacturaB_GotFocus()
    pEnfocaTextBox txtFolioFacturaB
End Sub

Private Sub txtFolioFacturaB_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
