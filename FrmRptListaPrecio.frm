VERSION 5.00
Begin VB.Form FrmRptListaPrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de precios"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   120
      TabIndex        =   19
      Top             =   -30
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6090
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   120
      TabIndex        =   10
      Top             =   615
      Width           =   7125
      Begin VB.OptionButton OptPaciente 
         Caption         =   "Urgencias"
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   5265
      End
      Begin VB.OptionButton OptPaciente 
         Caption         =   "Interno"
         Height          =   195
         Index           =   0
         Left            =   1725
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton OptPaciente 
         Caption         =   "Externo"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   5265
      End
      Begin VB.ComboBox cboConceptoFac 
         Height          =   315
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   5265
      End
      Begin VB.TextBox txtDirigido 
         Height          =   315
         Left            =   1725
         TabIndex        =   8
         Top             =   2040
         Width           =   5265
      End
      Begin VB.ComboBox cboEmpresa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   5265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Concepto de factura"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dirigido a"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   2100
         Width           =   660
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3105
      TabIndex        =   9
      Top             =   3360
      Width           =   1140
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmRptListaPrecio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmRptListaPrecio.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmRptListaPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vgrptReporte As CRAXDRT.Report
Dim vlstrsql As String
Dim rsTemp As New ADODB.Recordset
Dim vlCveEmp As Integer

Private Sub cboConceptoFac_KeyDown(KeyCode As Integer, shift As Integer)
   If KeyCode = vbKeyReturn Then pEnfocaTextBox txtDirigido
End Sub

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, shift As Integer)
   If KeyCode = vbKeyReturn Then pEnfocaCbo cboConceptoFac
End Sub

Private Sub cboEmpresa_Click()
   txtDirigido.Text = cboEmpresa.List(cboEmpresa.ListIndex)
   vlCveEmp = cboEmpresa.ItemData(cboEmpresa.ListIndex)
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
        cboDepartamento.ItemData(cboDepartamento.NewIndex) = 0
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, shift As Integer)
    If KeyCode = vbKeyReturn Then optPaciente(0).SetFocus
End Sub

Private Sub cboTipoPaciente_Click()
    
    vlstrsql = "Select bitUtilizaConvenio from AdTipoPaciente Where Chrtipo = 'CO' AND tnyCveTipoPaciente = " & _
    cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)
    Set rsTemp = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
        cboEmpresa.Enabled = IIf(rsTemp.RecordCount = 0, False, True)
        vlCveEmp = IIf(rsTemp.RecordCount = 0, 0, cboEmpresa.ItemData(cboEmpresa.ListIndex))
        txtDirigido.Text = IIf(rsTemp.RecordCount = 0, "", cboEmpresa.List(cboEmpresa.ListIndex))
    rsTemp.Close
    
End Sub

Private Sub cboEmpresa_KeyDown(KeyCode As Integer, shift As Integer)
   vlCveEmp = cboEmpresa.ItemData(cboEmpresa.ListIndex)
   If KeyCode = vbKeyReturn Then pEnfocaCbo cboDepartamento
End Sub

Private Sub cboTipoPaciente_KeyDown(KeyCode As Integer, shift As Integer)
   vlCveEmp = cboEmpresa.ItemData(cboEmpresa.ListIndex)
   cboTipoPaciente_Click
   If KeyCode = vbKeyReturn And cboEmpresa.Enabled Then
      pEnfocaCbo cboEmpresa
   ElseIf KeyCode = vbKeyReturn And Not cboEmpresa.Enabled Then
      pEnfocaCbo cboDepartamento
   Else
      vlCveEmp = 0
   End If
End Sub

Private Sub cmdPreview_Click()
   pImprime "P"
End Sub
Private Sub pImprime(Impresora As String)
    Dim vlrsRptListaPrecio As New ADODB.Recordset
    Dim alstrParametros(1) As String
    
    frsEjecuta_SP Trim(Str(vlCveEmp)) & "|" & _
         Trim(Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))) & "|" & _
         Trim(Str(cboConceptoFac.ItemData(cboConceptoFac.ListIndex))) & "|" & _
         Trim(Str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex))) & "|" & _
         IIf(optPaciente(0).Value, "I", IIf(optPaciente(1).Value, "E", "U")) & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex)), "sp_pvselrptListaPrecio"
    
   Set vlrsRptListaPrecio = frsRegresaRs("SELECT * FROM PVRPTLISTAPRECIO")
   If vlrsRptListaPrecio.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
   Else
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
        alstrParametros(1) = "Dirigido;" & Trim(txtDirigido.Text)
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, vlrsRptListaPrecio, Impresora, "Lista de precios"
   End If
   vlrsRptListaPrecio.Close
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim lngNumOpcion As Long

    Me.Icon = frmMenuPrincipal.Icon
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 346
    Case "SE"
         lngNumOpcion = 1536
    End Select
    
    pCargaHospital lngNumOpcion
   
    vlCveEmp = 0
    
    pInstanciaReporte vgrptReporte, "rptListaPrecio.rpt"
    
    vlstrsql = "Select * from pvConceptoFacturacion"
    Set rsTemp = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboConceptoFac, rsTemp, 0, 1, 3
    rsTemp.Close
    cboConceptoFac.ListIndex = 0
    
    vlstrsql = "Select * from ccempresa"
    Set rsTemp = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rsTemp, 0, 1, 3
    rsTemp.Close
    cboEmpresa.ListIndex = 0
    
    vlstrsql = "Select * from adtipoPaciente "
    Set rsTemp = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboTipoPaciente, rsTemp, 0, 1
    rsTemp.Close
    cboTipoPaciente.ListIndex = 0
   
End Sub

Private Sub OptPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then pEnfocaCbo cboTipoPaciente
End Sub

Private Sub txtDirigido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPreview.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
