VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmFilRptDescuentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos por tipo de paciente"
   ClientHeight    =   4410
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1305
      Left            =   4995
      TabIndex        =   23
      Top             =   2280
      Width           =   2190
      Begin VB.OptionButton OptConcentrado 
         Caption         =   "Concentrado"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Reporte concetrado"
         Top             =   780
         Width           =   1960
      End
      Begin VB.OptionButton OptConcepto 
         Caption         =   "Detallado por concepto"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Detallado por concepto"
         Top             =   480
         Width           =   1960
      End
      Begin VB.OptionButton OptCargo 
         Caption         =   "Detallado por cargo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Detallado por cargo"
         Top             =   180
         Value           =   -1  'True
         Width           =   2000
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   21
      Top             =   0
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1530
      Left            =   60
      TabIndex        =   19
      Top             =   705
      Width           =   7125
      Begin VB.ComboBox cboTipoConvenio 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   5460
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         ItemData        =   "FrmFilRptDescuentos.frx":0000
         Left            =   1470
         List            =   "FrmFilRptDescuentos.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Departamento"
         Top             =   240
         Width           =   5460
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         ItemData        =   "FrmFilRptDescuentos.frx":0004
         Left            =   1470
         List            =   "FrmFilRptDescuentos.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Tipos de pacientes y empresas"
         Top             =   660
         Width           =   5460
      End
      Begin VB.Label lblTipoConvenio 
         Caption         =   "Tipo de convenio"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de fechas"
      Height          =   1305
      Left            =   1800
      TabIndex        =   16
      Top             =   2280
      Width           =   3180
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   315
         Left            =   1305
         TabIndex        =   7
         ToolTipText     =   "Fecha inicial"
         Top             =   375
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107806721
         CurrentDate     =   38794
      End
      Begin MSComCtl2.DTPicker DtpFechaFin 
         Height          =   315
         Left            =   1305
         TabIndex        =   8
         ToolTipText     =   "Fecha final"
         Top             =   795
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107806721
         CurrentDate     =   38794
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha final"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   855
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicial"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   435
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2520
      TabIndex        =   15
      Top             =   3720
      Width           =   2175
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   1065
         Picture         =   "FrmFilRptDescuentos.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   555
         Picture         =   "FrmFilRptDescuentos.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vista previa"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame FraPaciente 
      Caption         =   "Paciente"
      Height          =   1305
      Left            =   60
      TabIndex        =   14
      Top             =   2280
      Width           =   1650
      Begin VB.OptionButton optPaciente 
         Caption         =   "Todos"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Ambos"
         Top             =   280
         Width           =   975
      End
      Begin VB.OptionButton optPaciente 
         Caption         =   "Externos"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Pacientes externos"
         Top             =   880
         Width           =   975
      End
      Begin VB.OptionButton optPaciente 
         Caption         =   "Internos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Pacientes internos"
         Top             =   585
         Width           =   975
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de paciente"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1020
      Width           =   1200
   End
End
Attribute VB_Name = "FrmFilRptDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vgrptReporte As CRAXDRT.Report
Dim vlstrSentencia As String

Private Sub cboDepartamento_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_KeyPress"))
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboDepartamento.SetFocus
End Sub


Private Sub cboTipoConvenio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
            If optPaciente(0).Value = True Then
                optPaciente(0).SetFocus
            Else
                If optPaciente(1).Value = True Then
                    optPaciente(1).SetFocus
                Else
                    optPaciente(2).SetFocus
                End If
            End If
        End If
End Sub

Private Sub cboTipoPaciente_Click()
    If cboTipoPaciente Like "PARTICULAR" Or cboTipoPaciente.ListIndex > 0 Then
            cboTipoPaciente_KeyDown vbKeyReturn, 0
    End If
End Sub

Private Sub cboTipoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cboTipoPaciente Like "PARTICULAR" Then
            If cboTipoConvenio.Enabled = True Then
                cboTipoConvenio.Enabled = False
                lblTipoConvenio.Enabled = False
                cboTipoConvenio.ListIndex = 0
            End If
        Else
            cboTipoConvenio.Enabled = True
            lblTipoConvenio.Enabled = True
            cboTipoConvenio.SetFocus
        End If
    
        
        If cboTipoConvenio.Enabled = False Then
            If optPaciente(0).Value = True Then
                optPaciente(0).SetFocus
            Else
                If optPaciente(1).Value = True Then
                    optPaciente(1).SetFocus
                Else
                        optPaciente(2).SetFocus
                End If
            End If
        Else
            cboTipoConvenio.SetFocus
        End If
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
      Unload Me
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError
        pImprime "I"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprimir_Click"))
      Unload Me
End Sub

Private Sub cmdVistaPreliminar_Click()
    On Error GoTo NotificaError
        pImprime "P"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVistaPreliminar_Click"))
      Unload Me
End Sub

Private Sub dtpFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError 'Manejo del error
        If KeyCode = vbKeyReturn Then cmdVistaPreliminar.SetFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":DtpFechaFin_KeyDown"))
End Sub

Private Sub dtpFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError 'Manejo del error
        If KeyCode = vbKeyReturn Then
            If Not IsDate(DtpFechaInicio.Value) Then
                MsgBox SIHOMsg(29), vbCritical, "Mensaje"
                DtpFechaInicio.SetFocus
            Else
                dtpFechaFin.SetFocus
            End If
        End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":DtpFechaInicio_KeyDown"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError 'Manejo del error
        If KeyAscii = 27 Then Unload Me
        Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
      Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError 'Manejo del error
    
    Dim dtmfecha As Date
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim rsDepartamento As New ADODB.Recordset
    Dim rsConvenio As New ADODB.Recordset
    
    Dim lngNumOpcion As Long
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 328
    Case "SE"
         lngNumOpcion = 2002
    End Select
    
    'Departamentos
    Set rsDepartamento = frsEjecuta_SP("-1|1|*|" & vgintClaveEmpresaContable, "sp_GnSelDepartamento")
    If rsDepartamento.RecordCount > 0 Then
        pLlenarCboRs cboDepartamento, rsDepartamento, 0, 1, 3
        cboDepartamento.ListIndex = 0
        cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
    End If
    cboDepartamento.Enabled = fblnRevisaPermiso(vglngNumeroLogin, lngNumOpcion, "C")
    
    ' Pacientes en negativo y empresas en positivo
    Set rsTipoPaciente = frsEjecuta_SP("2", "sp_GnSelTipoPacienteEmpresa")
    pLlenarCboRs cboTipoPaciente, rsTipoPaciente, 1, 0, 3
    cboTipoPaciente.ListIndex = 0
    
    ' Tipos de Convenios
    Set rsConvenio = frsEjecuta_SP("", "SP_GNSELTIPOCONVENIO")
    pLlenarCboRs cboTipoConvenio, rsConvenio, 0, 1, 3
    cboTipoConvenio.ListIndex = 0
    
    pCargaHospital lngNumOpcion
    
    dtmfecha = fdtmServerFecha
    DtpFechaInicio.Value = dtmfecha
    dtpFechaFin.Value = dtmfecha
    optPaciente(2).Value = True
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
      Unload Me
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


Private Sub OptPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then DtpFechaInicio.SetFocus
End Sub

Private Sub pImprime(vlStrOpcion As String)
    On Error GoTo NotificaError
    
        Dim rsReporte As ADODB.Recordset
        Dim vlstrParametro As String
        Dim alstrParametros(1) As String
        
        ' Fecha inicial
        vlstrParametro = fstrFechaSQL(DtpFechaInicio.Value, "00:00:00")
        ' Fecha final
        vlstrParametro = vlstrParametro & "|" & fstrFechaSQL(dtpFechaFin.Value, "23:59:59")
        ' Pacientes, en el combo cboTipoPaciente son negativos
        vlstrParametro = IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) < 0, vlstrParametro & "|" & (cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)) * -1, vlstrParametro & "|0")
        ' Empresas, en el combo cboTipoPaciente son positivos
        vlstrParametro = IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) > 0, vlstrParametro & "|" & (cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)), vlstrParametro & "|0")
        ' Paciente Interno, Externo o Ambos
        vlstrParametro = IIf(optPaciente(0).Value, vlstrParametro & "|I", IIf(optPaciente(1).Value, vlstrParametro & "|E", vlstrParametro & "|A"))
        ' Empresa contable
        vlstrParametro = vlstrParametro & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
        
        
        
        If Not OptConcentrado.Value Then
            If cboTipoConvenio.Enabled = False And Not OptConcentrado.Value Then
                vlstrParametro = vlstrParametro & "|0"
            Else
            ' Tipo de Convenio
            vlstrParametro = vlstrParametro & "|" & Str(cboTipoConvenio.ItemData(cboTipoConvenio.ListIndex))
            End If
        Else
                
        End If
        vlstrParametro = vlstrParametro & "|" & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))
        
        pInstanciaReporte vgrptReporte, IIf(OptConcentrado.Value, "rptDescuentosTipoPacienteConcentrado.rpt", "rptDescuentosTipoPaciente.rpt")
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital; " & Trim(cboHospital.List(cboHospital.ListIndex))
        alstrParametros(1) = "Detallado;" & IIf(optCargo.Value, "True", "False") & ";BOOLEAN"
        
        pCargaParameterFields alstrParametros, vgrptReporte
        
        Set rsReporte = frsEjecuta_SP(vlstrParametro, IIf(OptConcentrado.Value, "SP_PVRPTDESCUENTOSCONCENTRADO", "SP_PVRPTDESCUENTOSXTIPOPAC"))
        If rsReporte.EOF Then
            MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        Else
            pImprimeReporte vgrptReporte, rsReporte, vlStrOpcion, "Descuentos por tipo de paciente"
        End If
        rsReporte.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime_Click"))
      Unload Me
End Sub
