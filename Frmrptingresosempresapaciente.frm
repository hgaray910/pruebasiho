VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmrptingresosempresapaciente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos por empresa referida"
   ClientHeight    =   5025
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fraagrupar 
      Caption         =   "Agrupar por"
      Height          =   855
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.OptionButton Opttipo 
         Caption         =   "Tipo de paciente"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Tipo de paciente"
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Opttipo 
         Caption         =   "Tipo de convenio"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Tipo de convenio"
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Opttipo 
         Caption         =   "Empresa convenio"
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   3
         ToolTipText     =   "Empresa convenio"
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Opttipo 
         Caption         =   "Empresa referida"
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   4
         ToolTipText     =   "Empresa referida"
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3247
      TabIndex        =   12
      Top             =   4080
      Width           =   1170
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Frmrptingresosempresapaciente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   585
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Frmrptingresosempresapaciente.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Rango de fechas"
      Height          =   855
      Left            =   1545
      TabIndex        =   9
      Top             =   3240
      Width           =   4575
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   315
         Left            =   855
         TabIndex        =   5
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
         TabIndex        =   6
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
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   390
         Width           =   420
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   405
         Width           =   465
      End
   End
   Begin VB.Frame Frafiltros 
      Height          =   1935
      Left            =   80
      TabIndex        =   13
      Top             =   1080
      Width           =   7455
      Begin VB.ComboBox cbotipopaciente 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Selección del tipo de paciente"
         Top             =   360
         Width           =   5415
      End
      Begin VB.ComboBox cbotipoconvenio 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Selección del tipo de convenio"
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox cboempresaconvenio 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Selección de la empresa"
         Top             =   1080
         Width           =   5415
      End
      Begin VB.ComboBox cboempresapaciente 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Selección de la empresa referida"
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   420
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de convenio"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa convenio"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1140
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa referida"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Selección de la empresa referida"
         Top             =   1500
         Width           =   1185
      End
   End
End
Attribute VB_Name = "Frmrptingresosempresapaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboempresaconvenio_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

    If KeyCode = 13 Then
        cboEmpresaPaciente.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboempresaconvenio_KeyDown"))
    Unload Me

End Sub

Private Sub cboEmpresaPaciente_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

    If KeyCode = 13 Then
        mskFechaIni.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpresaPaciente_KeyDown"))
    Unload Me
    
End Sub

Private Sub cboTipoConvenio_Click()
Dim vlstrsql As String
Dim rsEmpresas As New ADODB.Recordset

On Error GoTo NotificaError


'COMBO EMPRESA CONVENIO
'                vlstrsql = "SELECT EMPRESAS.INTCVEEMPRESA," & _
'                        "EMPRESAS.VCHDESCRIPCION" & _
'                        " From" & _
'                        "(SELECT CcEmpresa.intCveEmpresa, CcEmpresa.vchDescripcion FROM CcEmpresa" & _
'                        " Where (tnyCveTipoConvenio =" & cbotipoconvenio.ItemData(cbotipoconvenio.ListIndex) & ") And BITACTIVO = 1" & _
'                        " Union" & _
'                        " SELECT CcEmpresa.intCveEmpresa, CcEmpresa.vchDescripcion FROM CcEmpresa" & _
'                        " WHERE BITACTIVO= 1 AND CCEMPRESA.BITMOSTRAREMPRESAPACIENTE=1 AND CCEMPRESA.tnyCveTipoConvenio<>" & cbotipoconvenio.ItemData(cbotipoconvenio.ListIndex) & ")EMPRESAS"
                cboempresaconvenio.Clear
                vlstrsql = "SELECT CcEmpresa.intCveEmpresa, CcEmpresa.vchDescripcion FROM CcEmpresa WHERE (tnyCveTipoConvenio = " & cboTipoConvenio.ItemData(cboTipoConvenio.ListIndex) & ") AND BITACTIVO = 1"
                
                
                Set rsEmpresas = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
                If rsEmpresas.RecordCount > 0 Then
                    Call pLlenarCboRs(cboempresaconvenio, rsEmpresas, 0, 1)
                    cboempresaconvenio.ListIndex = 0
                    cboempresaconvenio.Enabled = True
                End If
                rsEmpresas.Close
                
                cboempresaconvenio.AddItem "<TODOS>", 0
                cboempresaconvenio.ItemData(cboempresaconvenio.newIndex) = -1
                cboempresaconvenio.ListIndex = 0
                
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoConvenio_Click"))
    Unload Me
    
End Sub

Private Sub cboTipoConvenio_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

    If KeyCode = 13 Then
        cboempresaconvenio.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoConvenio_KeyDown"))
    Unload Me

End Sub

Private Sub cboTipoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    
 On Error GoTo NotificaError
 
    If KeyCode = 13 Then
        cboTipoConvenio.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoPaciente_KeyDown"))
    Unload Me

End Sub

Private Sub cmdPreview_Click()

On Error GoTo NotificaError

    pImprime "P"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPreview_Click"))
    Unload Me

End Sub

Private Sub cmdPrint_Click()
    
On Error GoTo NotificaError

    pImprime "I"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
    Unload Me
    
End Sub

Private Sub Form_Activate()

On Error GoTo NotificaError

    llenacombos
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
    
End Sub

Private Sub pImprime(strDestino As String)
    
    Dim rsReporte As New ADODB.Recordset
    Dim rptReporte As CRAXDRT.Report
    Dim alstrParametros(5) As String
        
    On Error GoTo NotificaError
    
    vgstrParametrosSP = fstrFechaSQL(mskFechaIni.Text, "00:00:00") & "|" & fstrFechaSQL(mskFechaFin.Text, "23:59:59") & "|" & CStr(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)) & _
    "|" & CStr(cboTipoConvenio.ItemData(cboTipoConvenio.ListIndex)) & "|" & CStr(cboempresaconvenio.ItemData(cboempresaconvenio.ListIndex)) & _
    "|" & CStr(cboEmpresaPaciente.ItemData(cboEmpresaPaciente.ListIndex)) & "|" & IIf(optTipo(0).Value, 0, IIf(optTipo(1).Value, 1, IIf(optTipo(2).Value, 2, 3))) & _
    "|" & CStr(vgintClaveEmpresaContable)
    
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "SP_PVRPTINGRESOSPOREMPRESAPACI")
    
    If rsReporte.RecordCount <> 0 Then
        pInstanciaReporte rptReporte, "rptingresosporempresapaci.rpt"
        rptReporte.DiscardSavedData

        alstrParametros(0) = "NombreEmpresa;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(1) = "FechaIni;" & CDate(mskFechaIni.Text) & ";DATE"
        alstrParametros(2) = "FechaFin;" & CDate(mskFechaFin.Text) & ";DATE"
        'alstrParametros(3) = "Departamento;" & cboDepartamento.Text
        'alstrParametros(4) = "Detallado;" & IIf(chkDetallado.Value = 1, 1, 0)
        'alstrParametros(5) = "Grupo;" & IIf(optFecha.Value, 1, 0)
        
        pCargaParameterFields alstrParametros, rptReporte
        pImprimeReporte rptReporte, rsReporte, strDestino, "Ingresos por empresa referida"

    
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    rsReporte.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
    Unload Me
    
End Sub

Private Sub llenacombos()

Dim rsTipoPaciente As New ADODB.Recordset
Dim rsTiposConvenio As New ADODB.Recordset
Dim rsEmpresas As New ADODB.Recordset
Dim rsEmpresaspaciente As New ADODB.Recordset
Dim vlstrsql As String

On Error GoTo NotificaError
                
                'COMBO TIPO PACIENTE
                vlstrsql = "Select tnyCveTipoPaciente, vchDescripcion " & _
                               "FROM AdTipoPaciente "
                
                Set rsTipoPaciente = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
                If rsTipoPaciente.RecordCount > 0 Then
                    pLlenarCboRs cboTipoPaciente, rsTipoPaciente, 0, 1
                End If
                rsTipoPaciente.Close
                cboTipoPaciente.AddItem "<TODOS>", 0
                cboTipoPaciente.ItemData(cboTipoPaciente.newIndex) = -1
                cboTipoPaciente.ListIndex = 0
                
                'COMBO TIPO CONVENIO
                vlstrsql = "select tnyCveTipoConvenio, vchDescripcion from CcTipoConvenio"
                Set rsTiposConvenio = frsRegresaRs(vlstrsql)
                If rsTiposConvenio.RecordCount <> 0 Then
                    pLlenarCboRs cboTipoConvenio, rsTiposConvenio, 0, 1
                
                End If
                rsTiposConvenio.Close
                cboTipoConvenio.AddItem "<TODOS>", 0
                cboTipoConvenio.ItemData(cboTipoConvenio.newIndex) = -1
                cboTipoConvenio.ListIndex = 0
                
                
                
                'COMBO EMPRESAS PACIENTE
                vlstrsql = "SELECT CcEmpresa.intCveEmpresa, CcEmpresa.vchDescripcion FROM CcEmpresa WHERE   BITACTIVO = 1 AND BITMOSTRAREMPRESAPACIENTE=1"
                Set rsEmpresaspaciente = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
                If rsEmpresaspaciente.RecordCount > 0 Then
                    Call pLlenarCboRs(cboEmpresaPaciente, rsEmpresaspaciente, 0, 1)
                    cboEmpresaPaciente.Enabled = True
                End If
                rsEmpresaspaciente.Close
                cboEmpresaPaciente.AddItem "<TODOS>", 0
                cboEmpresaPaciente.ItemData(cboEmpresaPaciente.newIndex) = -1
                cboEmpresaPaciente.ListIndex = 0
                
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":llenacombos"))
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    dtmfecha = fdtmServerFecha
    
    mskFechaIni.Mask = ""
    mskFechaIni.Text = dtmfecha
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
    
End Sub

Private Sub mskFechaFin_GotFocus()

On Error GoTo NotificaError

    pSelMkTexto mskFechaFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_GotFocus"))

End Sub

Private Sub mskFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

If KeyCode = 13 Then
    cmdPreview.SetFocus
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_KeyDown"))
    
End Sub

Private Sub mskFechaIni_GotFocus()

On Error GoTo NotificaError

    pSelMkTexto mskFechaIni

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaIni_GotFocus"))
    
End Sub

Private Sub mskFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

    If KeyCode = 13 Then
        mskFechaFin.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaIni_KeyDown"))

End Sub

Private Sub Opttipo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

    cboTipoPaciente.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Opttipo_KeyDown"))

End Sub
