VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRptConciliacionEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliación de servicios entre empresas"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTipoPaciente 
      Caption         =   "Pacientes"
      Height          =   1080
      Left            =   255
      TabIndex        =   28
      Top             =   1700
      Width           =   7680
      Begin VB.TextBox txtCuenta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   1710
         TabIndex        =   8
         Text            =   "txtCuenta"
         ToolTipText     =   "Cuenta del paciente"
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optExterno 
         Caption         =   "Externos"
         Height          =   255
         Left            =   4110
         TabIndex        =   7
         Top             =   280
         Width           =   975
      End
      Begin VB.OptionButton optInterno 
         Caption         =   "Internos"
         Height          =   255
         Left            =   2910
         TabIndex        =   6
         Top             =   280
         Width           =   975
      End
      Begin VB.OptionButton optAmbos 
         Caption         =   "Ambos"
         Height          =   255
         Left            =   1710
         TabIndex        =   5
         Top             =   280
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   150
         TabIndex        =   31
         Top             =   320
         Width           =   1350
      End
      Begin VB.Label lblNombrePaciente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNombrePaciente"
         Height          =   315
         Left            =   2910
         TabIndex        =   30
         Top             =   600
         Width           =   4555
      End
      Begin VB.Label Label4 
         Caption         =   "Número de cuenta"
         Height          =   255
         Left            =   150
         TabIndex        =   29
         Top             =   650
         Width           =   1350
      End
   End
   Begin VB.Frame fraRangoFechas 
      Caption         =   "Rango de fechas"
      Height          =   1600
      Left            =   255
      TabIndex        =   25
      Top             =   2870
      Width           =   2500
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   300
         Left            =   1005
         TabIndex        =   9
         ToolTipText     =   "Fecha inicial"
         Top             =   450
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   300
         Left            =   1005
         TabIndex        =   10
         ToolTipText     =   "Fecha final"
         Top             =   950
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Left            =   300
         TabIndex        =   27
         Top             =   1000
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Left            =   300
         TabIndex        =   26
         Top             =   500
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3520
      TabIndex        =   24
      Top             =   4600
      Width           =   1140
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   75
         Picture         =   "frmRptConciliacionEmpresas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         Picture         =   "frmRptConciliacionEmpresas.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cboClasificacion 
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Clasificación del servicio"
      Top             =   930
      Width           =   6000
   End
   Begin VB.ComboBox cboConceptoFacturacion 
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Concepto de facturación del servicio"
      Top             =   1320
      Width           =   6000
   End
   Begin VB.CheckBox chkDetallado 
      Caption         =   "Detallado"
      Height          =   255
      Left            =   255
      TabIndex        =   16
      Top             =   4580
      Width           =   1095
   End
   Begin VB.ComboBox cboProveedor 
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Empresa que realiza servicios a otra empresa"
      Top             =   525
      Width           =   6000
   End
   Begin VB.ComboBox cboCliente 
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Empresa que puede solicitar servicios a otra empresa"
      Top             =   120
      Width           =   6000
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Ordenar por"
      Height          =   1600
      Left            =   5430
      TabIndex        =   19
      Top             =   2870
      Width           =   2500
      Begin VB.OptionButton optClasificacion 
         Caption         =   "Clasificación"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   750
         Width           =   1335
      End
      Begin VB.OptionButton optConceptoFactura 
         Caption         =   "Concepto facturación"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   1150
         Width           =   1935
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   350
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraAgrupacion 
      Caption         =   "Agrupar por"
      Height          =   1600
      Left            =   2840
      TabIndex        =   0
      Top             =   2870
      Width           =   2500
      Begin VB.OptionButton optAConceptoFactura 
         Caption         =   "Concepto facturación"
         Height          =   255
         Left            =   300
         TabIndex        =   12
         Top             =   1000
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optAClasificacion 
         Caption         =   "Clasificación"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   500
         Width           =   1215
      End
   End
   Begin VB.Label lblClasificacion 
      Caption         =   "Clasificación"
      Height          =   255
      Left            =   255
      TabIndex        =   23
      Top             =   980
      Width           =   1605
   End
   Begin VB.Label Label3 
      Caption         =   "Concepto facturación"
      Height          =   255
      Left            =   255
      TabIndex        =   22
      Top             =   1370
      Width           =   1605
   End
   Begin VB.Label lblProveedor 
      Caption         =   "Empresa proveedor"
      Height          =   255
      Left            =   255
      TabIndex        =   21
      Top             =   570
      Width           =   1605
   End
   Begin VB.Label lblCliente 
      Caption         =   "Empresa cliente"
      Height          =   255
      Left            =   255
      TabIndex        =   20
      Top             =   165
      Width           =   1605
   End
End
Attribute VB_Name = "frmRptConciliacionEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vgrptReporte As CRAXDRT.Report
Dim rs As New ADODB.Recordset

Private Sub cmdImprimir_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            SendKeys vbTab
        Case vbKeyEscape
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                Unload Me
            End If
    End Select
    
End Sub

Private Sub Form_Load()

    Me.Icon = frmMenuPrincipal.Icon
    
    txtCuenta.Text = ""
    lblNombrePaciente.Caption = ""
    
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = fdtmServerFecha
    mskFechaInicio.Mask = "##/##/####"

    mskFechaFin.Mask = ""
    mskFechaFin.Text = fdtmServerFecha
    mskFechaFin.Mask = "##/##/####"
    
    optConceptoFactura.Enabled = False
    
    'Empresas cliente
    Set rs = frsEjecuta_SP("", "sp_PvSelEmpresaCliente")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboCliente, rs, 0, 1
        cboCliente.ListIndex = 0
    Else
        'No existen empresas configuradas como clientes
        MsgBox SIHOMsg(962), vbInformation, "Mensaje"
        Exit Sub
    End If
    rs.Close
    
    'Empresa proveedor
    Set rs = frsEjecuta_SP("-1", "sp_GnSelEmpresascontable")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboProveedor, rs, 1, 0
        cboProveedor.ListIndex = fintLocalizaCbo(cboProveedor, CStr(vgintClaveEmpresaContable))
    End If
    rs.Close

    'Clasificaciones
    Set rs = frsEjecuta_SP(CStr(cboProveedor.ItemData(cboProveedor.ListIndex)), "sp_PvSelClasificacionesEstExa")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboClasificacion, rs, 0, 1, 3
        cboClasificacion.ListIndex = 0
    Else
        'No existen clasificaciones activas
        MsgBox SIHOMsg(963), vbInformation, "Mensaje"
        Unload Me
    End If
    rs.Close
    
    'Concepto de facturación
    vgstrParametrosSP = "0|1|0|" & CStr(cboProveedor.ItemData(cboProveedor.ListIndex))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelConceptoFacturacion")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboConceptoFacturacion, rs, 0, 1, 3
        cboConceptoFacturacion.ListIndex = 0
    Else
        'No existen conceptos de facturación activos.
        MsgBox SIHOMsg(756), vbInformation, "Mensaje"
        Unload Me
    End If
    rs.Close
       
End Sub

Private Sub pImprime(lstrDestino As String)
    On Error GoTo NotificaError

    Dim lstrFechaInicio As String
    Dim lstrFechaFin As String
    Dim rsReporte As New ADODB.Recordset
    Dim alstrParametros(11) As String
      
    If cboCliente.ListIndex = -1 Then
        'No existen empresas configuradas como clientes
        MsgBox SIHOMsg(962), vbInformation, "Mensaje"
        Exit Sub
    End If
    If CDate(mskFechaInicio.Text) > fdtmServerFecha Then
        '¡La fecha debe ser menor o igual a la del sistema!
        MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaInicio.SetFocus
        Exit Sub
    ElseIf CDate(mskFechaFin.Text) > fdtmServerFecha Then
        '¡La fecha debe ser menor o igual a la del sistema!
        MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaFin.SetFocus
        Exit Sub
    ElseIf CDate(mskFechaInicio.Text) > CDate(mskFechaFin.Text) Then
        '¡Rango de fechas no válido!
        MsgBox SIHOMsg(64), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaInicio.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11

    If chkDetallado.Value = 0 Then
        pInstanciaReporte vgrptReporte, "rptConciliacionServicios.rpt"
    Else
        pInstanciaReporte vgrptReporte, "rptConciliacionServiciosDet.rpt"
    End If
    vgrptReporte.DiscardSavedData
    
    lstrFechaInicio = mskFechaInicio.Text
    lstrFechaFin = mskFechaFin.Text
    
    vgstrParametrosSP = CStr(cboCliente.ItemData(cboCliente.ListIndex)) & _
                        "|" & CStr(cboProveedor.ItemData(cboProveedor.ListIndex)) & _
                        "|" & lstrFechaInicio & _
                        "|" & lstrFechaFin & _
                        "|" & cboClasificacion.ItemData(cboClasificacion.ListIndex) & _
                        "|" & cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex) & _
                        "|" & IIf(optAmbos.Value, "A", IIf(optInterno.Value, "I", "E")) & _
                        "|" & Val(txtCuenta.Text) & _
                        "|" & IIf(optAClasificacion.Value, 0, 1) & _
                        "|" & IIf(optFecha.Value, 0, IIf(optClasificacion.Value, 1, 2))

    
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptConciliancionEmpresas")
    If rsReporte.RecordCount > 0 Then
        alstrParametros(0) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(1) = "Cliente;" & Trim(cboCliente.Text)
        alstrParametros(2) = "Proveedor;" & Trim(cboProveedor.Text)
        alstrParametros(3) = "Clasificacion;" & Trim(cboClasificacion)
        alstrParametros(4) = "ConceptoF;" & Trim(cboConceptoFacturacion.Text)
        alstrParametros(5) = "Cuenta;" & IIf(Val(txtCuenta.Text) > 0, (txtCuenta.Text), "<TODAS>")
        alstrParametros(6) = "Tipo paciente;" & IIf(optAmbos.Value, "", IIf(optInterno.Value, "INTERNO", "EXTERNO"))
        alstrParametros(7) = "Detallado;" & IIf(chkDetallado.Value, 1, 0)
        alstrParametros(8) = "FechaInicio;" & CDate(mskFechaInicio.Text) & ";DATE"
        alstrParametros(9) = "FechaFin;" & CDate(mskFechaFin.Text) & ";DATE"
        alstrParametros(10) = "Grupo;" & IIf(optAClasificacion.Value, "Concepto facturación", "Clasificación")
        
        pCargaParameterFields alstrParametros, vgrptReporte

        pImprimeReporte vgrptReporte, rsReporte, lstrDestino
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    rsReporte.Close

    Me.MousePointer = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
End Sub

Private Sub mskFechaFin_GotFocus()
    pSelMkTexto mskFechaFin
End Sub

Private Sub mskFechaFin_LostFocus()
    
    If Not IsDate(mskFechaFin.Text) Then
        mskFechaFin.Mask = ""
        mskFechaFin.Text = fdtmServerFecha
        mskFechaFin.Mask = "##/##/####"
    End If
    
End Sub

Private Sub mskFechaInicio_GotFocus()
    pSelMkTexto mskFechaInicio
End Sub

Private Sub mskFechaInicio_LostFocus()

    If Not IsDate(mskFechaInicio.Text) Then
        mskFechaInicio.Mask = ""
        mskFechaInicio.Text = fdtmServerFecha
        mskFechaInicio.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optAClasificacion_Click()
    pFiltros
End Sub

Private Sub optAConceptoFactura_Click()
    pFiltros
End Sub

Private Sub optAmbos_Click()
    txtCuenta.Enabled = False
    txtCuenta.Text = ""
    lblNombrePaciente.Caption = ""
End Sub
    
Private Sub optATipoCargo_Click()
    pFiltros
End Sub

Private Sub pFiltros()
    
    optClasificacion.Enabled = True
    optConceptoFactura.Enabled = True
    optFecha.Value = True
    
    If optAClasificacion.Value Then
        optClasificacion.Enabled = False
    ElseIf optAConceptoFactura Then
        optConceptoFactura.Enabled = False
    End If
    
End Sub

Private Sub optExterno_Click()
    txtCuenta.Enabled = True
End Sub

Private Sub optInterno_Click()
    txtCuenta.Enabled = True
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
    
        If Trim(txtCuenta.Text) = "" Then
            
            With FrmBusquedaPacientes
                If optInterno.Value Then
                    
                    .vgstrMovCve = "M"
                    .vgstrTipoPaciente = "I"
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgblnPideClave = True
                    .optTodos.Value = True
                    .optTodos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"

                ElseIf optExterno.Value Then

                    .vgstrMovCve = "M"
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgIntMaxRecords = 100
                    .vgblnPideClave = True
                    .optTodos.Value = True
                    .optTodos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                    
                End If
        
                txtCuenta.Text = .flngRegresaPaciente()
                If txtCuenta.Text <> -1 Then
                    txtCuenta_KeyDown vbKeyReturn, 0
                Else
                    txtCuenta.Text = ""
                End If
            End With
            
        Else
            
            vgstrParametrosSP = Val(txtCuenta.Text) & "|" & IIf(optInterno.Value, "I", "E") & "|" & cboCliente.ItemData(cboCliente.ListIndex)
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldatospaciente")
            If rs.RecordCount > 0 Then lblNombrePaciente.Caption = rs!Nombre
            rs.Close
            
        End If
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
End Sub
