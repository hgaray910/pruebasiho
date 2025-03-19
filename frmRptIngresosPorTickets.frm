VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptIngresosPorTickets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos por tickets"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3052
      TabIndex        =   15
      Top             =   2985
      Width           =   1170
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   585
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptIngresosPorTickets.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptIngresosPorTickets.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox chkDetallado 
      Caption         =   "Detallado"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   3060
      Width           =   1695
   End
   Begin VB.Frame fraAgrupacion 
      Caption         =   "Agrupar por"
      Height          =   1455
      Left            =   4290
      TabIndex        =   14
      Top             =   1485
      Width           =   2895
      Begin VB.OptionButton optTicket 
         Caption         =   "Ticket"
         Height          =   255
         Left            =   165
         TabIndex        =   19
         Top             =   1120
         Width           =   1215
      End
      Begin VB.OptionButton optVenta 
         Caption         =   "Venta"
         Height          =   255
         Left            =   165
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDepartamento 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   550
         Width           =   1455
      End
   End
   Begin VB.Frame fraDepartamento 
      Height          =   855
      Left            =   60
      TabIndex        =   13
      Top             =   630
      Width           =   7125
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   5730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   330
         Width           =   1005
      End
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Rango de fechas"
      Height          =   1455
      Left            =   60
      TabIndex        =   10
      Top             =   1485
      Width           =   4185
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         ToolTipText     =   "Fecha inicial"
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   2700
         TabIndex        =   3
         ToolTipText     =   "Fecha final"
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   405
         Width           =   465
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   390
         Width           =   420
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   -45
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   5730
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRptIngresosPorTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPreview_Click()
    pImprime "P"
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Dim rs As New ADODB.Recordset
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 2137
    Case "SE"
         lngNumOpcion = 2138
    End Select
    
    pCargaHospital lngNumOpcion
    
    dtmfecha = fdtmServerFecha
    
    mskFechaIni.Mask = ""
    mskFechaIni.Text = dtmfecha
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"
End Sub


Private Sub mskFechaFin_GotFocus()
    pSelMkTexto mskFechaFin
End Sub

Private Sub mskFechaIni_GotFocus()
    pSelMkTexto mskFechaIni
End Sub

Private Sub pImprime(strDestino As String)
    Dim rsReporte As New ADODB.Recordset
    Dim rptReporte As CRAXDRT.Report
    Dim alstrParametros(5) As String
    Dim intSeleccion As Integer
        
        If optFecha.Value = True Then intSeleccion = 1
        If optDepartamento.Value = True Then intSeleccion = 2
        If optVenta.Value = True Then intSeleccion = 3
        If optTicket.Value = True Then intSeleccion = 4
        
    vgstrParametrosSP = fstrFechaSQL(mskFechaIni.Text, "00:00:00") & "|" & fstrFechaSQL(mskFechaFin.Text, "23:59:59") & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & intSeleccion & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptIngresosPorTicket")
    
    If rsReporte.RecordCount <> 0 Then
        pInstanciaReporte rptReporte, "rptIngresosPorTicket.rpt"
        rptReporte.DiscardSavedData

        alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex)) & ";TRUE"
        alstrParametros(1) = "FechaIni;" & CDate(mskFechaIni.Text) & ";DATE"
        alstrParametros(2) = "FechaFin;" & CDate(mskFechaFin.Text) & ";DATE"
        alstrParametros(3) = "Departamento;" & cboDepartamento.Text
        alstrParametros(4) = "Detallado;" & IIf(fblLicenciaIEPS, IIf(chkDetallado.Value = 1, 1, 2), IIf(chkDetallado.Value = 1, 3, 4))
        alstrParametros(5) = "Grupo;" & intSeleccion
        
        pCargaParameterFields alstrParametros, rptReporte
        pImprimeReporte rptReporte, rsReporte, strDestino

    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
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
        cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me

End Sub


