VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRptIngresoPaquetes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos por paquetes"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmBotonera 
      Height          =   735
      Left            =   3120
      TabIndex        =   25
      Top             =   4080
      Width           =   1140
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         Picture         =   "frmRptIngresoPaquetes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVista 
         Height          =   495
         Left            =   75
         Picture         =   "frmRptIngresoPaquetes.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rango de fechas de factura"
      Height          =   700
      Left            =   130
      TabIndex        =   22
      Top             =   3285
      Width           =   4200
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         ToolTipText     =   "Fecha final"
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   315
         Left            =   720
         TabIndex        =   10
         ToolTipText     =   "Fecha inicial"
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   255
         Left            =   2190
         TabIndex        =   24
         Top             =   285
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   320
         Width           =   465
      End
   End
   Begin VB.CheckBox chkPresentacion 
      Caption         =   "Detallado"
      Height          =   255
      Left            =   6282
      TabIndex        =   12
      ToolTipText     =   "Reporte detallado"
      Top             =   3765
      Value           =   1  'Checked
      Width           =   968
   End
   Begin VB.Frame Frame1 
      Height          =   2520
      Left            =   130
      TabIndex        =   16
      Top             =   720
      Width           =   7120
      Begin VB.ComboBox cboMedico 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Selección del médico tratante"
         Top             =   600
         Width           =   5400
      End
      Begin VB.TextBox txtNumeroCuenta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   8
         ToolTipText     =   "Número de cuenta"
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         Height          =   465
         Left            =   1590
         TabIndex        =   17
         Top             =   1350
         Width           =   3200
         Begin VB.OptionButton optPaciente 
            Caption         =   "Externos"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   7
            ToolTipText     =   "Sólo externos"
            Top             =   175
            Width           =   930
         End
         Begin VB.OptionButton optPaciente 
            Caption         =   "Internos"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   6
            ToolTipText     =   "Sólo internos"
            Top             =   175
            Width           =   915
         End
         Begin VB.OptionButton optPaciente 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Todos los tipos de ingreso"
            Top             =   175
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Selección del tipo de paciente"
         Top             =   230
         Width           =   5400
      End
      Begin VB.ComboBox cboPaquete 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Selección del paquete"
         Top             =   1005
         Width           =   5400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Médico tratante"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label lblNombrePaciente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2655
         TabIndex        =   9
         ToolTipText     =   "Nombre del paciente"
         Top             =   1920
         Width           =   4305
      End
      Begin VB.Label lblNumeroCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Número de cuenta"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de ingreso"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paquete"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1065
         Width           =   600
      End
   End
   Begin VB.Frame Frame7 
      Height          =   675
      Left            =   130
      TabIndex        =   0
      Top             =   0
      Width           =   7120
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Selección de la empresa"
         Top             =   230
         Width           =   6060
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   290
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRptIngresoPaquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset   'Varios usos
Dim ldtmFecha As Date           'Fecha actual
Dim vlstrHoraInicio As String
Dim vlstrHoraFin As String

Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError
        If fblnValidos() Then pImprime "I"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdImprimir_Click"))
End Sub

Private Sub pBuscaPaciente()
    On Error GoTo NotificaError

    If Trim(txtNumeroCuenta.Text) = "" Then
        With FrmBusquedaPacientes
            If optPaciente(2).Value Then    'Externos
                .vgstrTipoPaciente = "E"
                .Caption = .Caption & " externos"
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSoloActivos.Enabled = True
                .optSinFacturar.Enabled = True
                .optTodos.Enabled = True
                '.optSinFacturar.Value = True
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
            
            txtNumeroCuenta.Text = .flngRegresaPaciente()
               
            If Val(txtNumeroCuenta.Text) <> -1 Then
                pCargaNombrePaciente
            Else
                txtNumeroCuenta.Text = ""
            End If
        End With
    Else
        pCargaNombrePaciente
    End If
    
    If Trim(lblNombrePaciente.Caption) <> "" Then
        mskFechaInicio.SetFocus
    Else
        pEnfocaTextBox txtNumeroCuenta
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pBuscaPaciente"))
End Sub

Private Sub pCargaNombrePaciente()
    On Error GoTo NotificaError

    vgstrParametrosSP = txtNumeroCuenta.Text & "|" & "0" & "|" & IIf(optPaciente(1).Value, "I", "E") & "|" & vgintClaveEmpresaContable
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")
    If rs.RecordCount <> 0 Then
        lblNombrePaciente.Caption = rs!Nombre
    Else
        MsgBox SIHOMsg(355), vbOKOnly + vbInformation, "Mensaje" 'No se encontró la información del paciente.
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaNombrePaciente"))
End Sub

Private Sub cmdVista_Click()
    On Error GoTo NotificaError
        pImprime "P"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdVista_Click"))
End Sub

Private Function fblnValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnValidos = True
    If fstrFechaSQL(mskFechaInicio.Text, "00:00:00") > fstrFechaSQL(mskFechaFin.Text, "23:59:59") Then
        fblnValidos = False
        MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje" '¡Rango de fechas no válido!
        mskFechaInicio.SetFocus
    End If
    
    If fblnValidos And Val(txtNumeroCuenta.Text) <> 0 And Trim(lblNombrePaciente.Caption) = "" Then
        fblnValidos = False
        MsgBox SIHOMsg(27), vbOKOnly + vbInformation, "Mensaje" 'Debe seleccionar un paciente.
        txtNumeroCuenta.SetFocus
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnValidos"))
End Function

Private Sub pImprime(strDestino As String)
    On Error GoTo NotificaError
    Dim rptReporte As CRAXDRT.Report
    Dim alstrParametros(8) As String
    Dim lngCveCargo As Long
    Dim strTipoPaciente As String
    
    If fblnValidos() Then
        FrmBotonera.Enabled = False
        strTipoPaciente = ""
        strTipoPaciente = IIf(optPaciente(0).Value, "*", IIf(optPaciente(1).Value, "I", IIf(optPaciente(2).Value, "E", "T")))
    
        vgstrParametrosSP = ""
        vgstrParametrosSP = CStr(cboHospital.ItemData(cboHospital.ListIndex)) _
            & "|" & CStr(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)) _
            & "|" & CStr(cboPaquete.ItemData(cboPaquete.ListIndex)) _
            & "|" & strTipoPaciente _
            & "|" & IIf(Trim(lblNombrePaciente.Caption) = "", "0", txtNumeroCuenta.Text) _
            & "|" & Format(mskFechaInicio, "yyyy-mm-dd") & " 00:00:00" _
            & "|" & Format(mskFechaFin, "yyyy-mm-dd") & " 23:59:59" _
            & "|" & CStr(cboMedico.ItemData(cboMedico.ListIndex))
    
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINGRESOSPAQUETES")
        If rs.RecordCount <> 0 Then
            pInstanciaReporte rptReporte, IIf(chkPresentacion, "rptIngresoPaquetes.rpt", "rptIngresoPaquetesConcentrado.rpt")
            rptReporte.DiscardSavedData
            alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
            alstrParametros(2) = "TipoPaciente;" & Trim(cboTipoPaciente.List(cboTipoPaciente.ListIndex))
            alstrParametros(3) = "Paciente;" & IIf(strTipoPaciente = "*", "<TODOS>", IIf(strTipoPaciente = "I", IIf(Trim(lblNombrePaciente.Caption) = "", "<INTERNOS>", lblNombrePaciente.Caption), IIf(strTipoPaciente = "E", IIf(Trim(lblNombrePaciente.Caption) = "", "<EXTERNOS>", lblNombrePaciente.Caption), "")))
            alstrParametros(4) = "FechaInicio;" & UCase(Format(mskFechaInicio.Text, "dd/mmm/yyyy"))
            alstrParametros(5) = "FechaFin;" & UCase(Format(mskFechaFin.Text, "dd/mmm/yyyy"))
            alstrParametros(6) = "Paquete;" & Trim(cboPaquete.List(cboPaquete.ListIndex))
            
            pCargaParameterFields alstrParametros, rptReporte
            pImprimeReporte rptReporte, rs, strDestino, "Ingresos por paquetes"
        Else
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje" 'No existe información con esos parámetro
        End If
        rs.Close
        FrmBotonera.Enabled = True
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtNumeroCuenta" Then
            pBuscaPaciente
        Else
            SendKeys vbTab
        End If
    Else
        If KeyCode = vbKeyEscape Then
            Unload Me
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyDown"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim lngNumOpcion As Long
    Dim rsPaquetes As New ADODB.Recordset
    Dim rsMedicos As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 338
    Case "SE"
         lngNumOpcion = 1361
    End Select
    
    pCargaHospital lngNumOpcion
    pCargaTipoPaciente  'Tipos de paciente
    
    Set rsPaquetes = frsRegresaRs("SELECT intNumPaquete, TRIM(chrDescripcion) FROM PVPAQUETE WHERE bitActivo = 1 ")
    If rsPaquetes.RecordCount > 0 Then
        pLlenarCboRs cboPaquete, rsPaquetes, 0, 1
        rsPaquetes.Close
    End If
    cboPaquete.AddItem "<TODOS>", 0
    cboPaquete.ItemData(cboPaquete.newIndex) = 0
    cboPaquete.ListIndex = 0
    
    Set rsMedicos = frsEjecuta_SP("-1|1", "SP_EXSELMEDICO")
    If rsMedicos.RecordCount > 0 Then
        pLlenarCboRs cboMedico, rsMedicos, 0, 1
        rsMedicos.Close
    End If
    cboMedico.AddItem "<TODOS>", 0
    cboMedico.ItemData(cboMedico.newIndex) = 0
    cboMedico.ListIndex = 0

    optPaciente(0).Value = True
    OptPaciente_Click 0
    
    ldtmFecha = fdtmServerFecha
    
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = ldtmFecha
    mskFechaInicio.Mask = "##/##/####"
    mskFechaFin.Mask = ""
    mskFechaFin.Text = ldtmFecha
    mskFechaFin.Mask = "##/##/####"
    
    vlstrHoraInicio = "00:00:00"
    vlstrHoraFin = "23:59:59"
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pCargaTipoPaciente()
    On Error GoTo NotificaError

    vgstrParametrosSP = "1|*"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELTIPOPACIENTE")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboTipoPaciente, rs, 0, 1
    End If
    cboTipoPaciente.AddItem "<TODOS>", 0
    cboTipoPaciente.ItemData(cboTipoPaciente.newIndex) = 0
    cboTipoPaciente.ListIndex = 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaTipoPaciente"))
End Sub

Private Sub mskFechaFin_GotFocus()
    On Error GoTo NotificaError
        pSelMkTexto mskFechaFin
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaFin_LostFocus()
    On Error GoTo NotificaError
        If Not IsDate(Format(mskFechaFin.Text, "dd/mm/yyyy")) Then
            mskFechaFin.Mask = ""
            mskFechaFin.Text = ldtmFecha
            mskFechaFin.Mask = "##/##/####"
        Else
            mskFechaFin.Text = Format(mskFechaFin.Text, "dd/mm/yyyy")
        End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_LostFocus"))
End Sub

Private Sub mskFechaInicio_GotFocus()
    On Error GoTo NotificaError
        pSelMkTexto mskFechaInicio
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_GotFocus"))
End Sub

Private Sub mskFechaInicio_LostFocus()
    On Error GoTo NotificaError
        If Not IsDate(Format(mskFechaInicio.Text, "dd/mm/yyyy")) Then
            mskFechaInicio.Mask = ""
            mskFechaInicio.Text = ldtmFecha
            mskFechaInicio.Mask = "##/##/####"
        Else
            mskFechaInicio.Text = Format(mskFechaInicio.Text, "dd/mm/yyyy")
        End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_LostFocus"))
End Sub

Private Sub OptPaciente_Click(Index As Integer)
    On Error GoTo NotificaError

    lblNumeroCuenta.Enabled = Index <> 0
    txtNumeroCuenta.Enabled = Index <> 0
    lblNombrePaciente.Enabled = Index <> 0
    txtNumeroCuenta.Text = ""

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optPaciente_Click"))
End Sub

Private Sub txtNumeroCuenta_Change()
    On Error GoTo NotificaError
        lblNombrePaciente.Caption = ""
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCuenta_Change"))
End Sub

Private Sub txtNumeroCuenta_GotFocus()
    On Error GoTo NotificaError
        pSelTextBox txtNumeroCuenta
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCuenta_GotFocus"))
End Sub

Private Sub txtNumeroCuenta_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            optPaciente(1).Value = UCase(Chr(KeyAscii)) = "I"
            optPaciente(2).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCuenta_KeyPress"))
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
