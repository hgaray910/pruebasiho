VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPacientesAtendidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pacientes atendidos"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmDetallado 
      Height          =   750
      Left            =   4320
      TabIndex        =   38
      Top             =   3720
      Width           =   2860
      Begin VB.CheckBox ChkDetallado 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         ToolTipText     =   "Reporte por paciente detallado"
         Top             =   160
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChkSinFacturar 
         Caption         =   "Sin facturar"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Mostrar cargos sin facturar"
         Top             =   440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox ChkFacturados 
         Caption         =   "Facturados"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "Mostrar cargos facturados"
         Top             =   160
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de fechas de ingreso"
      Height          =   750
      Left            =   60
      TabIndex        =   34
      Top             =   3720
      Width           =   4170
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   795
         TabIndex        =   15
         ToolTipText     =   "Fecha de inicio"
         Top             =   285
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   315
         Left            =   2775
         TabIndex        =   16
         ToolTipText     =   "Fecha de fin"
         Top             =   285
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2115
         TabIndex        =   36
         Top             =   345
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   345
         Width           =   465
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Index           =   1
      Left            =   60
      TabIndex        =   30
      Top             =   -15
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa contable"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame frmOrdenamiento 
      Height          =   525
      Left            =   2790
      TabIndex        =   29
      Top             =   2655
      Width           =   4395
      Begin VB.OptionButton optOrdenFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   1390
         TabIndex        =   10
         Top             =   220
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton optOrdenNombre 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   2480
         TabIndex        =   11
         Top             =   220
         Width           =   915
      End
      Begin VB.OptionButton optOrdenFactura 
         Caption         =   "Factura"
         Height          =   195
         Left            =   3450
         TabIndex        =   12
         Top             =   220
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ordenamiento"
         Height          =   195
         Left            =   75
         TabIndex        =   32
         Top             =   220
         Width           =   990
      End
   End
   Begin VB.Frame frmNivel 
      Height          =   525
      Left            =   2790
      TabIndex        =   28
      Top             =   3165
      Width           =   4395
      Begin VB.OptionButton optAgrupado 
         Caption         =   "Cargo"
         Height          =   195
         Index           =   1
         Left            =   2480
         TabIndex        =   14
         Top             =   210
         Width           =   1650
      End
      Begin VB.OptionButton optAgrupado 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   0
         Left            =   1390
         TabIndex        =   13
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desplegar a nivel"
         Height          =   195
         Left            =   75
         TabIndex        =   33
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.Frame frmReporte 
      Caption         =   "Tipo de reporte"
      Height          =   1035
      Left            =   60
      TabIndex        =   27
      Top             =   2655
      Width           =   2685
      Begin VB.OptionButton optReporte 
         Caption         =   "Por paciente "
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   650
         Width           =   1245
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Por concepto de facturación"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   2490
      End
   End
   Begin VB.Frame Frame6 
      Height          =   750
      Left            =   3060
      TabIndex        =   26
      Top             =   4560
      Width           =   1140
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRPTPacientesAtendidos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRPTPacientesAtendidos.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1995
      Left            =   60
      TabIndex        =   22
      Top             =   645
      Width           =   7125
      Begin VB.OptionButton optInternoExterno 
         Caption         =   "Interno"
         Height          =   195
         Index           =   0
         Left            =   2685
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optInternoExterno 
         Caption         =   "Externo"
         Height          =   195
         Index           =   1
         Left            =   3630
         TabIndex        =   5
         Top             =   1200
         Width           =   960
      End
      Begin VB.OptionButton optInternoExterno 
         Caption         =   "Todos"
         Height          =   195
         Index           =   2
         Left            =   1785
         TabIndex        =   3
         Top             =   1200
         Width           =   945
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1785
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Departamento que realizó los cargos"
         Top             =   260
         Width           =   5145
      End
      Begin VB.TextBox txtNumPaciente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1500
         Width           =   855
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1785
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Tipo de paciente o empresa"
         Top             =   700
         Width           =   5145
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   270
         Left            =   135
         TabIndex        =   37
         Top             =   1182
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Departamento cargó"
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label lblNombrePaciente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2670
         TabIndex        =   7
         Top             =   1500
         Width           =   4230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de expediente"
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   1560
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   760
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmPacientesAtendidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmReporteGanancias                                    -
'-------------------------------------------------------------------------------------
'| Objetivo: Sacar un reporte de pacientes externos atendidos por departamento
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 17/Jul/2003
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : hoy
'| Fecha última modificación: 01/Jul/2003
'-------------------------------------------------------------------------------------
Option Explicit
Private vgrptReporte As CRAXDRT.Report
Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboEmpresa.SetFocus
End Sub
Private Sub cboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then optInternoExterno(2).SetFocus
End Sub
Private Sub cboHospital_Click()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset

    If cboHospital.ListIndex <> -1 Then
        cboDepartamento.Clear
        vgstrParametrosSP = "-1|1|*|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldepartamento")
        If rs.RecordCount <> 0 Then
            pLlenarCboRs cboDepartamento, rs, 0, 1
        End If
        cboDepartamento.AddItem "<TODOS>", 0
        cboDepartamento.ItemData(cboDepartamento.newIndex) = 0
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub
Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboDepartamento.SetFocus
End Sub
Private Sub chkDetallado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdVistaPreliminar.SetFocus
End Sub
Private Sub chkFacturados_Click()
    If Me.chkFacturados.Value = 0 Then
       If Me.ChkSinFacturar.Value = 0 Then
          chkFacturados.Value = 1
      Else
         If Me.optOrdenFactura.Value = True Then
         Me.optOrdenFecha.Value = True
         End If
         Me.optOrdenFactura.Enabled = False
      End If
    Else
      Me.optOrdenFactura.Enabled = True
    End If
End Sub
Private Sub ChkFacturados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ChkSinFacturar.SetFocus
End Sub
Private Sub ChkSinFacturar_Click()
   If Me.ChkSinFacturar.Value = False Then
       If Me.chkFacturados.Value = False Then
          ChkSinFacturar.Value = 1
       End If
    End If
End Sub
Private Sub ChkSinFacturar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then chkDetallado.SetFocus
End Sub
Private Sub cmdImprimir_Click()
  If CDate(Me.mskFin.Text) < CDate(Me.mskInicio.Text) Then
    MsgBox SIHOMsg(64), vbExclamation, "Mensaje"
    Me.mskInicio.SetFocus
  Else
    pImprime "I"
  End If
End Sub
Private Sub cmdVistaPreliminar_Click()
 If CDate(Me.mskFin.Text) < CDate(Me.mskInicio.Text) Then
    MsgBox SIHOMsg(64), vbExclamation, "Mensaje"
    Me.mskInicio.SetFocus
 Else
   pImprime "P"
 End If
 End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlintTipoArticulo As Integer
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 360
    Case "SE"
         lngNumOpcion = 1537
    End Select
    
    pCargaHospital lngNumOpcion
    
    'Empresas
    vlstrSentencia = "select intCveEmpresa, vchDescripcion from ccEmpresa "
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rs, 0, 1
    rs.Close
    
    'Tipos de Paciente
    vlstrSentencia = "Select tnyCveTipoPaciente, vchDescripcion from adTipoPaciente order by tnyCveTipoPaciente desc"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        With cboEmpresa
            .AddItem rs!vchDescripcion, 0
            .ItemData(.newIndex) = rs!tnyCveTipoPaciente * -1
        End With
        rs.MoveNext
    Loop
    rs.Close
    cboEmpresa.AddItem "<TODOS>", 0
    cboEmpresa.ItemData(0) = 0
    cboEmpresa.ListIndex = 0
    
    'Fechas
    dtmfecha = fdtmServerFecha
    mskInicio.Text = dtmfecha
    mskFin.Text = dtmfecha
    
    optInternoExterno(2) = True
    optReporte(0) = True
    optAgrupado(0) = True
End Sub

Private Sub mskFin_GotFocus()
    pSelMkTexto mskFin
End Sub
Private Sub mskFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If chkFacturados.Enabled Then
            chkFacturados.SetFocus
        Else
            cmdVistaPreliminar.SetFocus
        End If
    End If
End Sub
Private Sub mskFin_LostFocus()
If Not IsDate(mskFin.Text) Then
       MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
       Me.mskFin.SetFocus
End If

End Sub

Private Sub mskInicio_GotFocus()
    pSelMkTexto mskInicio
End Sub

Private Sub mskInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaMkTexto mskFin
End Sub
Private Sub mskInicio_LostFocus()
If Not IsDate(mskInicio.Text) Then
       MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
       Me.mskInicio.SetFocus
End If
End Sub

Private Sub optAgrupado_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mskInicio.SetFocus
End Sub
Private Sub optInternoExterno_Click(Index As Integer)
    txtNumPaciente.Text = ""
    Me.lblNombrePaciente.Caption = ""
    txtNumPaciente.Enabled = Index <> 2
End Sub
Private Sub optInternoExterno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If txtNumPaciente.Enabled = True Then
          txtNumPaciente.SetFocus
       Else
          optReporte(0).SetFocus
       End If
     End If
End Sub
Private Sub optOrdenFactura_Click()
    'cboDepartamento.SetFocus
End Sub
Private Sub optOrdenFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mskInicio.SetFocus
End Sub
Private Sub optOrdenFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mskInicio.SetFocus
End Sub
Private Sub optOrdenFecha_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'cboDepartamento.SetFocus
End Sub
Private Sub optOrdenNombre_Click()
    'cboDepartamento.SetFocus
End Sub
Private Sub optOrdenNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mskInicio.SetFocus
End Sub
Private Sub optReporte_Click(Index As Integer)
    frmNivel.Enabled = IIf(optReporte(0) = True, True, False)
    optAgrupado(0).Enabled = IIf(optReporte(0) = True, True, False)
    optAgrupado(1).Enabled = IIf(optReporte(0) = True, True, False)
    frmOrdenamiento.Enabled = IIf(optReporte(0) = True, False, True)
    optOrdenFecha.Enabled = IIf(optReporte(0) = True, False, True)
    optOrdenNombre.Enabled = IIf(optReporte(0) = True, False, True)
    chkDetallado.Value = IIf(optReporte(0) = True, 0, 1)
    chkFacturados.Value = 1
    optOrdenFactura.Enabled = IIf(optReporte(0) = True, False, True)
    ChkSinFacturar.Value = IIf(optReporte(0) = True, 0, 1)
    FrmDetallado.Enabled = IIf(optReporte(0) = True, False, True)
    chkDetallado.Enabled = IIf(optReporte(0) = True, False, True)
    chkFacturados.Enabled = IIf(optReporte(0) = True, False, True)
    ChkSinFacturar.Enabled = IIf(optReporte(0) = True, False, True)
    'If optReporte(1) = True Then optOrdenFecha.SetFocus
End Sub
Private Sub optReporte_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optReporte(0).Value = True Then
           optAgrupado(0).SetFocus
        Else
           optOrdenFecha.SetFocus
        End If
    End If
End Sub

Private Sub txtNumPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vllngNumeroPaciente As Long
       If KeyCode = vbKeyReturn Then
        If RTrim(txtNumPaciente.Text) = "" Then
            With FrmBusquedaPacientes
                .vgstrforma = "frmPacientesAtendidos"
                .vgstrTipoPaciente = IIf(Me.optInternoExterno(0).Value = True, "I", "E")
                .Caption = .Caption & IIf(Me.optInternoExterno(0).Value = True, " internos", " externos")
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "C"
                .optSoloActivos.Enabled = True
                .optSinFacturar.Enabled = True
                .optTodos.Enabled = True
                .optTodos.Value = True
                .vgStrOtrosCampos = ", CCempresa.vchDescripcion as Empresa, " & _
                " (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else '  Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
                " From ExPacienteDomicilio " & _
                " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
                " And GnDomicilio.intCveTipoDomicilio = 1 " & _
                " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
                " ExPaciente.dtmFechaNacimiento as ""Fecha Nac."", " & _
                " (Select GnTelefono.vchTelefono " & _
                " From ExPacienteTelefono " & _
                " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
                " And GnTelefono.intCveTipoTelefono = 1 " & _
                " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Telefono "
                .vgstrTamanoCampo = "800,3400,3000,4100,990,980"
                
                vllngNumeroPaciente = .flngRegresaPaciente()
                
                If vllngNumeroPaciente <> -1 Then
                    txtNumPaciente.Text = vllngNumeroPaciente
                    txtNumPaciente_KeyDown vbKeyReturn, 0
                Else
                    lblNombrePaciente.Caption = ""
                End If
            End With
        Else
        
        
           vlstrSentencia = "SELECT rtrim(VCHAPELLIDOPATERNO)||' '||rtrim(VCHAPELLIDOMATERNO)||' '||rtrim(VCHNOMBRE) as Nombre " & _
                    " from expaciente " & _
                    " Where expaciente.INTNUMPACIENTE = " & txtNumPaciente.Text
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            If rs.RecordCount <> 0 Then
                lblNombrePaciente.Caption = rs!Nombre
                optReporte(0).SetFocus
            Else
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
            End If
            rs.Close
        End If
    End If
End Sub
Private Sub txtNumPaciente_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub
Private Sub txtNumPaciente_LostFocus()
    If Trim(txtNumPaciente.Text) = "" Then lblNombrePaciente.Caption = ""
End Sub
Sub pImprime(pstrDestino As String)
    Dim rsReporte As ADODB.Recordset
    Dim vlstrInternoExterno As String
    Dim vlstrParametros As String
    Dim alstrParametros(1) As String
    
    If optInternoExterno(0) Then
        vlstrInternoExterno = "I"
    ElseIf optInternoExterno(1) Then
        vlstrInternoExterno = "E"
    ElseIf optInternoExterno(2) Then
        vlstrInternoExterno = "T"
    End If
    
    If optReporte(0) = True Then
        pInstanciaReporte vgrptReporte, "rptFacturasPacientes.rpt"
        vlstrParametros = 4 & "|" & fstrFechaSQL(mskInicio.Text, "00:00:00") & "|" & fstrFechaSQL(mskFin.Text, "23:59:59") & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & CLng(Val(txtNumPaciente.Text)) & "|" & chkFacturados.Value & "|" & vlstrInternoExterno & "|" & IIf(optAgrupado(1).Value, 1, 0) & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
    ElseIf optReporte(1) = True Then
        If chkDetallado.Value Then
            pInstanciaReporte vgrptReporte, "rptPacientesAtendidosDet.rpt"
            vlstrParametros = IIf(optOrdenFecha.Value, 1, IIf(optOrdenNombre.Value, 2, 3)) & "|" & fstrFechaSQL(mskInicio.Text, "00:00:00") & "|" & fstrFechaSQL(mskFin.Text, "23:59:59") & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & CLng(Val(txtNumPaciente.Text)) & "|" & chkFacturados.Value & "|" & ChkSinFacturar.Value & "|" & vlstrInternoExterno & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
        Else
            pInstanciaReporte vgrptReporte, "rptPacientesAtendidos.rpt"
            vlstrParametros = IIf(optOrdenFecha.Value, 1, IIf(optOrdenNombre.Value, 2, 3)) & "|" & fstrFechaSQL(mskInicio.Text, "00:00:00") & "|" & fstrFechaSQL(mskFin.Text, "23:59:59") & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & CLng(Val(txtNumPaciente.Text)) & "|" & chkFacturados.Value & "|" & ChkSinFacturar.Value & "|" & vlstrInternoExterno & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
        End If
    End If

    If optReporte(1) Then
        If chkDetallado.Value Then
           Set rsReporte = frsEjecuta_SP(vlstrParametros, "Sp_PvrptpacientesatendidosDet", , , , True)
        Else
           Set rsReporte = frsEjecuta_SP(vlstrParametros, "Sp_Pvrptpacientesatendidoscon", , , , True)
        End If
    Else
        Set rsReporte = frsEjecuta_SP(vlstrParametros, "Sp_Pvrptpacientesatendidos", , , , True)
    End If

    If rsReporte.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital; " & cboHospital.List(cboHospital.ListIndex)
        alstrParametros(1) = "paramFechas;" & "DEL " & UCase(Format(mskInicio.Text, "dd/MMM/yyyy")) & " AL " & UCase(Format(mskFin.Text, "dd/MMM/yyyy"))
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsReporte, pstrDestino, "Pacientes atendidos"
    End If
    rsReporte.Close

End Sub
Private Sub pCargaHospital(lngNumOpcion As Long)
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP("-1", "Sp_Gnselempresascontable")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboHospital, rs, 1, 0
        cboHospital.ListIndex = flngLocalizaCbo(cboHospital, Str(vgintClaveEmpresaContable))
    End If
    
    cboHospital.Enabled = fblnRevisaPermiso(vglngNumeroLogin, lngNumOpcion, "C")
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaHospital"))
    Unload Me
End Sub
