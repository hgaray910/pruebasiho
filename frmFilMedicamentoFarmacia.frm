VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFilMedicamentoFarmacia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de cargos facturados por procedencia"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   90
      TabIndex        =   26
      Top             =   15
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
         TabIndex        =   27
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   720
      Left            =   3127
      TabIndex        =   25
      Top             =   4530
      Width           =   1110
      Begin VB.CommandButton cmdVistaPrevia 
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000014&
         Picture         =   "frmFilMedicamentoFarmacia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Permite imprimir el presupuesto"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   555
         MaskColor       =   &H80000014&
         Picture         =   "frmFilMedicamentoFarmacia.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Permite imprimir el presupuesto"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3690
      Left            =   90
      TabIndex        =   15
      Top             =   750
      Width           =   7125
      Begin VB.ComboBox cboTipoCargo 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Tipo de cargo"
         Top             =   2475
         Width           =   5340
      End
      Begin VB.ComboBox cboConcepto 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Concepto de facturación"
         Top             =   2040
         Width           =   5340
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Departamento"
         Top             =   1605
         Width           =   5340
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Paciente"
         Height          =   900
         Left            =   30
         TabIndex        =   17
         Top             =   555
         Width           =   7005
         Begin VB.TextBox txtPaciente 
            Height          =   315
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   500
            Width           =   4155
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1140
            TabIndex        =   18
            Top             =   200
            Width           =   2925
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Externo"
               Height          =   195
               Index           =   1
               Left            =   1065
               TabIndex        =   4
               ToolTipText     =   "Pacientes externos"
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Interno"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   3
               ToolTipText     =   "Pacientes internos"
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Ambos"
               Height          =   195
               Index           =   2
               Left            =   1965
               TabIndex        =   5
               ToolTipText     =   "Ambos tipos de pacientes"
               Top             =   0
               Value           =   -1  'True
               Width           =   795
            End
         End
         Begin VB.TextBox txtMovimientoPaciente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   6
            ToolTipText     =   "Número de cuenta del paciente"
            Top             =   500
            Width           =   1110
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "Todos"
            Height          =   225
            Left            =   195
            TabIndex        =   2
            ToolTipText     =   "Todos los pacientes"
            Top             =   200
            Value           =   1  'Checked
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número de cuenta"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   570
            Width           =   1320
         End
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Tipo de paciente"
         Top             =   240
         Width           =   5340
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   1590
         TabIndex        =   11
         ToolTipText     =   "Fecha de inicio"
         Top             =   2925
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
         Left            =   3540
         TabIndex        =   12
         ToolTipText     =   "Fecha de fin"
         Top             =   2925
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Caption         =   "Rango de fechas de factura"
         Height          =   555
         Left            =   180
         TabIndex        =   24
         Top             =   2925
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   3015
         TabIndex        =   23
         Top             =   2985
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cargo"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   2535
         Width           =   990
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto de facturación"
         Height          =   390
         Left            =   195
         TabIndex        =   21
         Top             =   1995
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Departamento que factura"
         Height          =   405
         Left            =   195
         TabIndex        =   20
         Top             =   1530
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   300
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmFilMedicamentoFarmacia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
' Programa para relacion de cargos por departamento
' Fecha de programación: Marzo, 2003
'-----------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset
Dim vlstrSentencia  As String
Dim vgstrtipoReporte As String
Private vgrptReporte As CRAXDRT.Report

Private Sub cboConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

   If KeyCode = vbKeyReturn Then
      pEnfocaCbo cboTipoCargo
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboConcepto_KeyDown"))
    Unload Me
End Sub

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
   
   If KeyCode = vbKeyReturn Then
      pEnfocaCbo cboConcepto
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_KeyDown"))
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
        cboDepartamento.ItemData(cboDepartamento.NewIndex) = -1
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cboTipo.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_KeyDown"))
    Unload Me
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
   
   If KeyCode = vbKeyReturn Then
      chkTodos.SetFocus
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipo_KeyDown"))
    Unload Me
End Sub

Private Sub cboTipoCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

   If KeyCode = vbKeyReturn Then
        mskInicio.SetFocus
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoCargo_KeyDown"))
    Unload Me
End Sub

Private Sub chkTodos_Click()
    On Error GoTo NotificaError

   txtMovimientoPaciente.Text = ""
   txtPaciente.Text = ""
   If chkTodos.Value = 1 Then
      txtMovimientoPaciente.Enabled = False
   ElseIf Not OptTipoPaciente(2).Value Then
      txtMovimientoPaciente.Enabled = True
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkTodos_Click"))
    Unload Me
End Sub

Private Sub chkTodos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

   If KeyAscii = vbKeyReturn Then
      pEnfocaCbo cboDepartamento
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkTodos_KeyPress"))
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

Private Sub pImprime(vlstrDestino As String)
    
    On Error GoTo NotificaError

    Dim vldtmFechaInicio As Date
    Dim vldtmFechaFin As Date
    Dim vlstrTipoCargo As String
    Dim vlstrTituloReporte As String
    Dim vlstrSentencia As String
    Dim vlrsResultado As New ADODB.Recordset
    Dim alstrParametros(1) As String
    Dim vlstrPar As String
    

    If IsDate(mskInicio) And IsDate(mskFin) Then
        vgstrtipoReporte = "P"
        If txtMovimientoPaciente = "" And (chkTodos.Value = 0 And OptTipoPaciente(2).Value = False) Then
            'Debes seleccionar un paciente.
            MsgBox SIHOMsg(27), vbOKOnly + vbInformation, "Mensaje"
        Else
            vlstrTituloReporte = "RELACION DE "
            Select Case cboTipoCargo.ItemData(cboTipoCargo.ListIndex)
                Case -1
                    vlstrTipoCargo = "TO"
                    vlstrTituloReporte = vlstrTituloReporte + "CARGOS "
                Case 1
                    vlstrTipoCargo = "OC"
                    vlstrTituloReporte = vlstrTituloReporte + "OTROS CONCEPTOS DE CARGO "
                Case 2
                    vlstrTipoCargo = "EX"
                    vlstrTituloReporte = vlstrTituloReporte + "EXAMENES "
                Case 3
                    vlstrTipoCargo = "ES"
                    vlstrTituloReporte = vlstrTituloReporte + "ESTUDIOS "
                Case 4
                    vlstrTipoCargo = "AR"
                    vlstrTituloReporte = vlstrTituloReporte + "ARTICULOS "
            End Select
            vlstrPar = ""
            vlstrPar = vlstrPar & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|"
            vlstrPar = vlstrPar & vlstrTipoCargo & "|"
            vlstrPar = vlstrPar & fstrFechaSQL(mskInicio, "00:00:00", True) & "|"
            vlstrPar = vlstrPar & fstrFechaSQL(mskFin, "23:59:59", True) & "|"
            vlstrPar = vlstrPar & cboConcepto.ItemData(cboConcepto.ListIndex) & "|"
            vlstrPar = vlstrPar & cboTipo.ItemData(cboTipo.ListIndex) & "|"
            vlstrPar = vlstrPar & IIf(OptTipoPaciente(0).Value, "I", IIf(OptTipoPaciente(1).Value, "E", "A")) & "|"
            vlstrPar = vlstrPar & IIf(Val(txtMovimientoPaciente) = 0, -1, Val(txtMovimientoPaciente)) & "|"
            vlstrPar = vlstrPar & Trim(vgstrNombreDepartamento) & "|"
            vlstrPar = vlstrPar & Trim(cboHospital.ItemData(cboHospital.ListIndex))
            
            EntornoSIHO.ConeccionSIHO.CommandTimeout = 0
            
            Set vlrsResultado = frsEjecuta_SP(vlstrPar, "Sp_Ccrptrelacioncargoempresa", , , , True)
            
            If vlrsResultado.RecordCount > 0 Then
                vgrptReporte.DiscardSavedData
                vlstrTituloReporte = vlstrTituloReporte & "PROPORCIONADOS POR " & IIf(cboDepartamento.ItemData(cboDepartamento.ListIndex) = -1, "TODOS LOS DEPARTAMENTOS DEL ", cboDepartamento.List(cboDepartamento.ListIndex) + " DEL ")
                vlstrTituloReporte = vlstrTituloReporte & Trim(vgstrNombreHospitalCH)
                vlstrTituloReporte = vlstrTituloReporte & " A LOS AFILIADOS DE " + IIf(cboTipo.ItemData(cboTipo.ListIndex) = 0, "TODAS LAS EMPRESAS", cboTipo.List(cboTipo.ListIndex))
                vlstrTituloReporte = vlstrTituloReporte & " DEL " & UCase(Format(mskInicio.Text, "dd/mmm/yyyy")) + " AL " + UCase(Format(mskFin.Text, "dd/mmm/yyyy"))
                alstrParametros(0) = "TituloReporte;" + vlstrTituloReporte
                alstrParametros(1) = "NombreHospital;" + cboHospital.List(cboHospital.ListIndex)
                
                pCargaParameterFields alstrParametros, vgrptReporte
                pImprimeReporte vgrptReporte, vlrsResultado, vlstrDestino, "Reporte de "
            Else
                MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
            End If
            vlrsResultado.Close
        End If
    Else
        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
        pEnfocaMkTexto mskInicio
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
    Unload Me
End Sub

Private Sub pTraeNomPaciente()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset

    txtPaciente.Text = ""
    
    vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & "0" & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintClaveEmpresaContable
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")
    If rs.RecordCount <> 0 Then
        txtPaciente.Text = rs!Nombre
        pEnfocaCbo cboDepartamento
    Else
        pEnfocaTextBox txtMovimientoPaciente
    End If
    rs.Close
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pTraeNomPaciente"))
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    On Error GoTo NotificaError

    pImprime "P"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVistaPrevia_Click"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

   If KeyAscii = vbKeyEscape Then
      Unload Me
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
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

Private Sub Form_Load()
   On Error GoTo NotificaError
   
   Dim dtmfecha As Date
   Dim lngNumOpcion As Long
   
   dtmfecha = fdtmServerFecha
   
   Me.Icon = frmMenuPrincipal.Icon
   
   Select Case cgstrModulo
   Case "PV"
        lngNumOpcion = 370
   Case "CC"
        lngNumOpcion = 638
   Case "SE"
        lngNumOpcion = 1999
   End Select
   
   pCargaHospital lngNumOpcion
   
   vlstrSentencia = "select smiCveConcepto, chrDescripcion from pvConceptoFacturacion"
   Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount <> 0 Then
        pLlenarCboRs cboConcepto, rs, 0, 1
   End If
   rs.Close
   cboConcepto.AddItem "<TODOS>", 0
   cboConcepto.ItemData(0) = -1
   cboConcepto.ListIndex = 0
   mskInicio.Text = dtmfecha
   mskFin.Text = dtmfecha
   
   cboTipoCargo.AddItem "OTROS CONCEPTOS"
   cboTipoCargo.ItemData(cboTipoCargo.NewIndex) = 1
   cboTipoCargo.AddItem "EXAMENES"
   cboTipoCargo.ItemData(cboTipoCargo.NewIndex) = 2
   cboTipoCargo.AddItem "ESTUDIOS"
   cboTipoCargo.ItemData(cboTipoCargo.NewIndex) = 3
   cboTipoCargo.AddItem "ARTICULOS"
   cboTipoCargo.ItemData(cboTipoCargo.NewIndex) = 4
   cboTipoCargo.AddItem "<TODOS>", 0
   cboTipoCargo.ItemData(cboTipoCargo.NewIndex) = -1
   cboTipoCargo.ListIndex = 0
   vlstrSentencia = "SELECT tnyCveTipoPaciente, vchDescripcion From AdTipoPaciente order by vchDescripcion"
   Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount <> 0 Then
        pLlenarCboRs cboTipo, rs, 0, 1
   End If
   rs.Close
   
   vlstrSentencia = "SELECT intCveEmpresa, vchDescripcion From CCempresa order by vchDescripcion"
   Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount <> 0 Then
        With rs
           Do While Not .EOF
                 cboTipo.AddItem Trim(!vchDescripcion), cboTipo.ListCount
                 cboTipo.ItemData(cboTipo.NewIndex) = -1 * !intcveempresa
              .MoveNext
           Loop
           
        End With
   End If
   rs.Close
   cboTipo.AddItem "<TODOS>", 0
   cboTipo.ItemData(0) = 0
   cboTipo.ListIndex = 0
   pInstanciaReporte vgrptReporte, "rptrelacionmedicamentocargo.rpt"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub mskFin_GotFocus()

    pSelMkTexto mskFin

End Sub

Private Sub mskFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

   If KeyCode = vbKeyReturn Then
      cmdVistaPrevia.SetFocus
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFin_KeyDown"))
    Unload Me
End Sub


Private Sub mskInicio_GotFocus()

    pSelMkTexto mskInicio

End Sub

Private Sub mskInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

   If KeyCode = vbKeyReturn Then
      pEnfocaMkTexto mskFin
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskInicio_KeyDown"))
    Unload Me
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError

   txtMovimientoPaciente.Text = ""
   txtPaciente.Text = ""
   If Index = 2 Then
      txtMovimientoPaciente.Enabled = False
   Else
      txtMovimientoPaciente.Enabled = True
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPaciente_Click"))
    Unload Me
End Sub

Private Sub optTipoPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError

   If KeyAscii = vbKeyReturn Then
      If Index = 2 Then
         pEnfocaCbo cboDepartamento
      Else
         pEnfocaTextBox txtMovimientoPaciente
      End If
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_KeyPress"))
    Unload Me
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

   If KeyCode = vbKeyReturn Then
        If chkTodos.Value = 0 Then
            If Trim(txtMovimientoPaciente.Text) = "" Then
               With FrmBusquedaPacientes
                  If OptTipoPaciente(1).Value Then 'Externos
                     .vgstrTipoPaciente = "E"
                     .Caption = .Caption & " externos"
                     .vgblnPideClave = False
                     .vgIntMaxRecords = 100
                     .vgstrMovCve = "M"
                     .optSoloActivos.Enabled = True
                     .optSinFacturar.Enabled = True
                     .optTodos.Enabled = True
                     .optTodos.Value = True
                     .vgStrOtrosCampos = ", CCempresa.vchDescripcion as Empresa, " & _
                     " (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else ' Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
                    " From ExPacienteDomicilio " & _
                    " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
                    " And GnDomicilio.intCveTipoDomicilio = 1 " & _
                    " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
                     " ExPaciente.dtmFechaNacimiento ""Fecha Nac."", " & _
                     " (Select GnTelefono.vchTelefono " & _
                    " From ExPacienteTelefono " & _
                    " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
                    " And GnTelefono.intCveTipoTelefono = 1 " & _
                    " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Telefono "
                     .vgstrTamanoCampo = "800,3400,2800,4100,990,980"
                  Else
                     .vgstrTipoPaciente = "I"  'Internos
                     .vgblnPideClave = False
                     .Caption = .Caption & " internos"
                     .vgIntMaxRecords = 100
                     .vgstrMovCve = "M"
                     .optSinFacturar.Value = True
                     .optSinFacturar.Enabled = True
                     .optSoloActivos.Enabled = True
                     .optTodos.Enabled = True
                     .optTodos.Value = True
                     .vgStrOtrosCampos = ", CCempresa.vchDescripcion as Empresa, " & _
                     " (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else ' Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
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
                     .vgstrTamanoCampo = "950,3400,2800,4100,990,980"
                  End If
                  
                  txtMovimientoPaciente.Text = .flngRegresaPaciente()
               
                  If txtMovimientoPaciente <> -1 Then
                     txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                  Else
                     txtMovimientoPaciente.Text = ""
                  End If
               End With
            End If
            pTraeNomPaciente
            pEnfocaCbo cboDepartamento
      End If
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            OptTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
            OptTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyPress"))
    Unload Me
End Sub
