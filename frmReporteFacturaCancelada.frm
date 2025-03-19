VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteFacturaCancelada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas canceladas en el departamento"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Rango de fechas de cancelación"
      Height          =   810
      Left            =   90
      TabIndex        =   21
      Top             =   3000
      Width           =   7125
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   315
         Left            =   1605
         TabIndex        =   10
         ToolTipText     =   "Fecha inicial"
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   315
         Left            =   4065
         TabIndex        =   11
         ToolTipText     =   "Fecha final"
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
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
         Left            =   3240
         TabIndex        =   23
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   270
         TabIndex        =   22
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   90
      TabIndex        =   19
      Top             =   -15
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   1605
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   5325
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   90
      TabIndex        =   14
      Top             =   660
      Width           =   7125
      Begin VB.OptionButton optSocios 
         Caption         =   "Socios"
         Height          =   200
         Left            =   3840
         TabIndex        =   8
         Top             =   1515
         Width           =   1785
      End
      Begin VB.ComboBox CboTipo 
         Height          =   315
         Left            =   1605
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Departamento"
         Top             =   660
         Width           =   5325
      End
      Begin VB.OptionButton optGrupos 
         Caption         =   "Grupos de cuentas"
         Height          =   200
         Left            =   3840
         TabIndex        =   7
         Top             =   1305
         Width           =   1785
      End
      Begin VB.OptionButton optAmbos 
         Caption         =   "Todos"
         Height          =   200
         Left            =   1605
         TabIndex        =   3
         Top             =   1110
         Width           =   960
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Clientes"
         Height          =   200
         Left            =   3840
         TabIndex        =   6
         Top             =   1110
         Width           =   1305
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1605
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Departamento"
         Top             =   225
         Width           =   5325
      End
      Begin VB.OptionButton optExternos 
         Caption         =   "Pacientes externos"
         Height          =   200
         Left            =   1605
         TabIndex        =   5
         Top             =   1515
         Width           =   1755
      End
      Begin VB.OptionButton optInternos 
         Caption         =   "Pacientes internos"
         Height          =   200
         Left            =   1605
         TabIndex        =   4
         Top             =   1305
         Width           =   1725
      End
      Begin VB.ComboBox cboProcedencia 
         Height          =   315
         Left            =   1605
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Selección del tipo de paciente"
         Top             =   1875
         Width           =   5325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   270
         TabIndex        =   24
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Facturada"
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblDepartamentoGrid 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   270
         TabIndex        =   15
         Top             =   1935
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3082
      TabIndex        =   16
      Top             =   3840
      Width           =   1140
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteFacturaCancelada.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   540
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteFacturaCancelada.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmReporteFacturaCancelada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------------
' Reporte de las facturas canceladas en un rango de fechas del departamento
' Fecha de programación: Martes 16 de Abril de 2002
'--------------------------------------------------------------------------------------
'Ultimas modificaciones, especificar:
'Fecha:
'Descripción del cambio:
'--------------------------------------------------------------------------------------
Option Explicit

Dim vlstrX As String
Private vgrptReporte As CRAXDRT.Report
Public vglngNumeroOpcion As Long


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
        cboDepartamento.ItemData(cboDepartamento.NewIndex) = -1
        cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, Str(vgintNumeroDepartamento))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub CboTipo_Click()

        optAmbos.Value = 1
        optAmbos.Enabled = IIf(CboTipo.ListIndex = 3, False, True)
        optInternos.Enabled = IIf(CboTipo.ListIndex = 3, False, True)
        optExternos.Enabled = IIf(CboTipo.ListIndex = 3, False, True)
        optCliente.Enabled = IIf(CboTipo.ListIndex = 3, False, True)
        optGrupos.Enabled = IIf(CboTipo.ListIndex = 3, False, True)
        optSocios.Enabled = IIf(CboTipo.ListIndex = 3 Or CboTipo.ListIndex = 2, False, True)
        cboProcedencia.ListIndex = 0
        cboProcedencia.Enabled = IIf(CboTipo.ListIndex = 3, False, True)
        
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub


Private Sub Form_Activate()

    cboDepartamento.Enabled = fblnRevisaPermiso(vglngNumeroLogin, vglngNumeroOpcion, "C")

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon

    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 343
    Case "SE"
         lngNumOpcion = 1535
    End Select
    
    pCargaHospital lngNumOpcion

    pInstanciaReporte vgrptReporte, "rptFacturasCanceladas.rpt"
    
    
    Set rs = frsEjecuta_SP("2", "sp_GnSelTipoPacienteEmpresa")
    pLlenarCboRs cboProcedencia, rs, 1, 0, 3
    cboProcedencia.ListIndex = 0
    
    CboTipo.AddItem "<TODOS>", 0
    CboTipo.AddItem "FACTURAS CANCELADAS", 1
    CboTipo.AddItem "REFACTURACIONES", 2
    CboTipo.AddItem "FOLIOS CANCELADOS", 3
    CboTipo.ListIndex = 0
    
    optAmbos.Value = True
   
    dtmfecha = fdtmServerFecha
   
    mskFecIni.Mask = ""
    mskFecIni.Text = dtmfecha
    mskFecIni.Mask = "##/##/####"
    
    mskFecFin.Mask = ""
    mskFecFin.Text = dtmfecha
    mskFecFin.Mask = "##/##/####"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub


Private Sub mskFecFin_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFecFin

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_GotFocus"))
End Sub


Private Sub mskFecFin_LostFocus()
    On Error GoTo NotificaError

    If Trim(mskFecFin.ClipText) = "" Then
        mskFecFin.Mask = ""
        mskFecFin.Text = fdtmServerFecha
        mskFecFin.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecFin.Text) Then
            mskFecFin.Mask = ""
            mskFecFin.Text = fdtmServerFecha
            mskFecFin.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_LostFocus"))
End Sub

Private Sub mskFecIni_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFecIni

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_GotFocus"))
End Sub


Private Sub mskFecIni_LostFocus()
    On Error GoTo NotificaError
    
    If Trim(mskFecIni.ClipText) = "" Then
        mskFecIni.Mask = ""
        mskFecIni.Text = fdtmServerFecha
        mskFecIni.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecIni.Text) Then
            mskFecIni.Mask = ""
            mskFecIni.Text = fdtmServerFecha
            mskFecIni.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_LostFocus"))
End Sub

Private Sub optAmbos_Click()

    cboProcedencia.Enabled = True

End Sub


Private Sub optCliente_Click()

    cboProcedencia.ListIndex = 0
    cboProcedencia.Enabled = Not optCliente.Value

End Sub


Private Sub optExternos_Click()

    cboProcedencia.Enabled = True

End Sub


Private Sub optGrupos_Click()
    
    cboProcedencia.ListIndex = 0
    cboProcedencia.Enabled = Not optGrupos.Value

End Sub

Private Sub optInternos_Click()

    cboProcedencia.Enabled = True
    
End Sub


Sub pImprime(pstrDestino As String)
    On Error GoTo NotificaError
    
    Dim strTipoPaciente As String
    Dim strtipopacientereporte As String
    Dim rs As New ADODB.Recordset
    Dim alstrParametros(5) As String
    
    If cboDepartamento.ListIndex = -1 Then
        'Seleccione el departamento.
        MsgBox SIHOMsg(242), vbOKOnly + vbInformation, "Mensaje"
    Else
        
        strTipoPaciente = IIf(optInternos.Value, "I", strTipoPaciente)
        strTipoPaciente = IIf(optExternos.Value, "E", strTipoPaciente)
        strTipoPaciente = IIf(optCliente.Value, "C", strTipoPaciente)
        strTipoPaciente = IIf(optGrupos.Value, "G", strTipoPaciente)
        strTipoPaciente = IIf(optSocios.Value, "S", strTipoPaciente)
        strTipoPaciente = IIf(optAmbos.Value, "*", strTipoPaciente)
        
        strtipopacientereporte = IIf(optInternos.Value, "PACIENTES INTERNOS", strtipopacientereporte)
        strtipopacientereporte = IIf(optExternos.Value, "PACIENTES EXTERNOS", strtipopacientereporte)
        strtipopacientereporte = IIf(optCliente.Value, "CLIENTES", strtipopacientereporte)
        strtipopacientereporte = IIf(optGrupos.Value, "GRUPO DE CUENTAS", strtipopacientereporte)
        strtipopacientereporte = IIf(optSocios.Value, "SOCIOS", strtipopacientereporte)
        strtipopacientereporte = IIf(optAmbos.Value, "TODOS", strtipopacientereporte)
        
        vgstrParametrosSP = _
        fstrFechaSQL(mskFecIni.Text) & _
        "|" & fstrFechaSQL(mskFecFin.Text) & _
        "|" & IIf(cboDepartamento.ItemData(cboDepartamento.ListIndex) = 0, -1, cboDepartamento.ItemData(cboDepartamento.ListIndex)) & _
        "|" & strTipoPaciente & _
        "|" & Str(cboProcedencia.ItemData(cboProcedencia.ListIndex)) & _
        "|" & Str(cboHospital.ItemData(cboHospital.ListIndex)) & _
        "|" & CboTipo.ListIndex
        
        
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptFacturaCancelada")
        If rs.RecordCount <> 0 Then
            If CboTipo.ListIndex = 3 Then
                pInstanciaReporte vgrptReporte, "rptFoliosCancelados.rpt"
            Else
                pInstanciaReporte vgrptReporte, "rptFacturasCanceladas.rpt"
            End If
            vgrptReporte.DiscardSavedData
            alstrParametros(0) = "NombreHospital;" & cboHospital.List(cboHospital.ListIndex)
            alstrParametros(1) = "FechaInicio;" & UCase(Format(mskFecIni.Text, "dd/mmm/yyyy"))
            alstrParametros(2) = "FechaFin;" & UCase(Format(mskFecFin.Text, "dd/mmm/yyyy"))
            alstrParametros(3) = "Departamento;" & cboDepartamento.List(cboDepartamento.ListIndex)
            alstrParametros(4) = "Facturada;" & strtipopacientereporte
            alstrParametros(5) = "Procedencia;" & IIf(cboProcedencia.List(cboProcedencia.ListIndex) = "<TODOS>", "TODOS", cboProcedencia.List(cboProcedencia.ListIndex))
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rs, pstrDestino, "Facturas canceladas en el departamento"
        Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rs.Close
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
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

Private Sub optSocios_Click()
    cboProcedencia.ListIndex = 0
    cboProcedencia.Enabled = False

End Sub
