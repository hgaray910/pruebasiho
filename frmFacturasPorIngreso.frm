VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFacturasPorIngreso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas por lugar de ingreso al hospital"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   780
      Left            =   6780
      TabIndex        =   20
      Top             =   1830
      Width           =   405
   End
   Begin VB.Frame Frame4 
      Height          =   780
      Left            =   5205
      TabIndex        =   19
      Top             =   1830
      Width           =   405
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   60
      TabIndex        =   16
      Top             =   1830
      Width           =   5100
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         ToolTipText     =   "Fecha de inicio"
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   315
         Left            =   3570
         TabIndex        =   4
         ToolTipText     =   "Fecha de fin"
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   3030
         TabIndex        =   18
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rango de fechas"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   300
         Width           =   1230
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   14
      Top             =   -45
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Siguiente"
         Default         =   -1  'True
         Height          =   255
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   7125
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1485
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   645
         Width           =   5445
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   5445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   705
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame6 
      Height          =   780
      Left            =   5625
      TabIndex        =   8
      Top             =   1830
      Width           =   1140
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFacturasPorIngreso.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir"
         Top             =   195
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFacturasPorIngreso.frx":03CD
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Vista previa"
         Top             =   195
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFacturasPorIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private vgrptReporte As CRAXDRT.Report

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
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    pImprime "P"
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSiguiente_Click()
    SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 385
    Case "SE"
         lngNumOpcion = 2004
    End Select
    
    pCargaHospital lngNumOpcion
    
    Me.Icon = frmMenuPrincipal.Icon
    pInstanciaReporte vgrptReporte, "rptFacturasPorLugarIngreso.rpt"
    
    'Empresas
    strSQL = "select intCveEmpresa, vchDescripcion from ccEmpresa "
    Set rs = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rs, 0, 1
    rs.Close
    
    'Tipos de Paciente
    strSQL = "Select tnyCveTipoPaciente, vchDescripcion from adTipoPaciente order by tnyCveTipoPaciente desc"
    Set rs = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        With cboEmpresa
            .AddItem rs!vchDescripcion, 0
            .ItemData(.NewIndex) = rs!tnyCveTipoPaciente * -1
        End With
        rs.MoveNext
    Loop
    rs.Close
    cboEmpresa.AddItem "<TODOS>", 0
    cboEmpresa.ItemData(0) = 0
    cboEmpresa.ListIndex = 0
    
    dtmfecha = fdtmServerFecha
    
    mskInicio.Mask = ""
    mskInicio.Text = dtmfecha
    mskInicio.Mask = "##/##/####"
    
    mskFin.Mask = ""
    mskFin.Text = dtmfecha
    mskFin.Mask = "##/##/####"

End Sub

Private Sub mskFin_GotFocus()
    pSelMkTexto mskFin
End Sub

Private Sub mskFin_Validate(Cancel As Boolean)
    If Not IsDate(mskFin.Text) Then
        mskFin.Text = fdtmServerFecha
    End If
End Sub

Private Sub mskInicio_GotFocus()
    pSelMkTexto mskInicio
End Sub


Private Sub pImprime(strDestino As String)
    Dim alstrParametros(4) As String
    Dim rsReporte As ADODB.Recordset
    Dim strPar As String
    
    alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
    alstrParametros(1) = "FechaInicio;" & CDate(mskInicio.Text) & ";DATE"
    alstrParametros(2) = "FechaFin;" & CDate(mskFin.Text) & ";DATE"
    alstrParametros(3) = "TipoPacienteEmpresa;" & cboEmpresa.Text
    alstrParametros(4) = "Departamento;" & cboDepartamento.Text
    strPar = cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & Format(CDate(mskInicio.Text), "YYYY-MM-DD") & "|" & Format(CDate(mskFin.Text), "YYYY-MM-DD") & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
    
    Set rsReporte = frsEjecuta_SP(strPar, "sp_PVRptFacturasPorIngreso")
    If rsReporte.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        vgrptReporte.DiscardSavedData
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsReporte, strDestino, "Facturas por lugar de ingreso del paciente"
    End If
    rsReporte.Close
    mskInicio.SetFocus
End Sub

Private Sub mskInicio_Validate(Cancel As Boolean)
    If Not IsDate(mskInicio.Text) Then
        mskInicio.Text = DateSerial(Year(fdtmServerFecha), Month(fdtmServerFecha), 1)
    End If
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

