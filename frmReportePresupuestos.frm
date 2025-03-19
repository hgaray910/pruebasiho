VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmReportePresupuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   585
      Left            =   4365
      TabIndex        =   18
      Top             =   1905
      Width           =   2820
   End
   Begin VB.Frame Frame8 
      Height          =   705
      Left            =   6375
      TabIndex        =   20
      Top             =   2460
      Width           =   810
   End
   Begin VB.Frame Frame7 
      Height          =   705
      Left            =   4365
      TabIndex        =   19
      Top             =   2460
      Width           =   810
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   15
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
         TabIndex        =   16
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenado por"
      Height          =   1260
      Left            =   60
      TabIndex        =   14
      Top             =   1905
      Width           =   1860
      Begin VB.OptionButton optOrden 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "Nombre"
         Top             =   795
         Width           =   1080
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   3
         ToolTipText     =   "Fecha"
         Top             =   435
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   5220
      TabIndex        =   13
      Top             =   2460
      Width           =   1125
      Begin VB.CommandButton cmdVistaPrevia 
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000014&
         Picture         =   "frmReportePresupuestos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Vista previa del reporte"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   555
         MaskColor       =   &H80000014&
         Picture         =   "frmReportePresupuestos.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Permite imprimir el reporte"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   60
      TabIndex        =   12
      Top             =   675
      Width           =   7125
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Tipo de paciente"
         Top             =   300
         Width           =   5505
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Tipo de paciente"
         Top             =   675
         Width           =   5505
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   105
         TabIndex        =   21
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   735
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Rango de fechas"
      Height          =   1260
      Left            =   1965
      TabIndex        =   7
      Top             =   1905
      Width           =   2355
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   "Fecha de inicio"
         Top             =   315
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
         Left            =   840
         TabIndex        =   6
         ToolTipText     =   "Fecha de fin"
         Top             =   705
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   375
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   765
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmReportePresupuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim vlstrsql As String
Private vgrptReporte As CRAXDRT.Report

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cboTipo.SetFocus
    End If

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
        cboDepartamento.ItemData(cboDepartamento.NewIndex) = -1
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cboDepartamento.SetFocus
    End If

End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      optOrden(0).SetFocus
   End If
End Sub

Private Sub cmdImprimir_Click()

    pImprime "I"
    
End Sub

Private Sub pImprime(vlstrDestino As String)

  Dim rsReporte As New ADODB.Recordset
  Dim vlstrx As String
  Dim alstrParametros(0) As String
  
  If IsDate(mskInicio) And IsDate(mskFin) Then
    
    vlstrx = fstrFechaSQL(mskInicio.Text, "00:00:00", True)
    vlstrx = vlstrx & "|" & fstrFechaSQL(mskFin.Text, "23:59:59", True)
    vlstrx = vlstrx & "|" & cboTipo.ItemData(cboTipo.ListIndex)
    vlstrx = vlstrx & "|" & IIf(optOrden(0).Value, 0, 1)
    vlstrx = vlstrx & "|" & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))
    vlstrx = vlstrx & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
    
    Set rsReporte = frsEjecuta_SP(vlstrx, "sp_PvRptPresupuesto")
    
    If rsReporte.RecordCount > 0 Then
      vgrptReporte.DiscardSavedData
      alstrParametros(0) = "NombreHospital; " & Trim(vgstrNombreHospitalCH)
      pCargaParameterFields alstrParametros, vgrptReporte
      
      pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Presupuestos"
    Else
      'No existe información con esos parámetros.
      MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    If rsReporte.State <> adStateClosed Then rsReporte.Close
  Else
    MsgBox SIHOMsg(29), vbCritical, "Mensaje"
    pEnfocaMkTexto mskInicio
  End If


End Sub

Private Sub cmdVistaPrevia_Click()
    pImprime "P"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()

   Dim lngNumOpcion As Long
   Dim dtmfecha As Date
   

   Me.Icon = frmMenuPrincipal.Icon
   
   Select Case cgstrModulo
   Case "PV"
        lngNumOpcion = 369
   Case "SE"
        lngNumOpcion = 2007
   End Select
   
   pCargaHospital lngNumOpcion
   
   pInstanciaReporte vgrptReporte, "rptPresupuestos.rpt"
   
   dtmfecha = fdtmServerFecha
   
   mskInicio.Text = dtmfecha
   mskFin.Text = dtmfecha
   vlstrsql = "SELECT tnyCveTipoPaciente * -1 , vchDescripcion From AdTipoPaciente order by vchDescripcion"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount <> 0 Then
        pLlenarCboRs cboTipo, rs, 0, 1
   End If
   rs.Close
   
   vlstrsql = "SELECT intCveEmpresa, vchDescripcion From CCempresa order by vchDescripcion"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   With rs
      Do While Not .EOF
            cboTipo.AddItem Trim(!vchDescripcion), cboTipo.ListCount
            cboTipo.ItemData(cboTipo.NewIndex) = !intCveEmpresa
         .MoveNext
      Loop
      rs.Close
   End With
   cboTipo.AddItem "<TODOS>", 0
   cboTipo.ItemData(cboTipo.NewIndex) = 0
   cboTipo.ListIndex = 0

End Sub

Private Sub mskFin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      cmdVistaPrevia.SetFocus
   End If
End Sub

Private Sub mskInicio_GotFocus()
   pEnfocaMkTexto mskInicio
End Sub

Private Sub mskInicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      pEnfocaMkTexto mskFin
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


Private Sub optOrden_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        mskInicio.SetFocus
    End If

End Sub
