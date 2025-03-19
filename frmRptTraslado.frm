VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRptTraslado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslado de cargos"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Height          =   750
      Left            =   6240
      TabIndex        =   32
      Top             =   3405
      Width           =   1155
   End
   Begin VB.Frame Frame7 
      Height          =   750
      Left            =   4005
      TabIndex        =   31
      Top             =   3405
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   105
      TabIndex        =   29
      Top             =   -30
      Width           =   7290
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CheckBox chkDetallado 
      Caption         =   "Detallado"
      Height          =   225
      Left            =   135
      TabIndex        =   28
      ToolTipText     =   "Mostrar la información detallada o concentrada"
      Top             =   4260
      Value           =   1  'Checked
      Width           =   1110
   End
   Begin VB.Frame Frame6 
      Height          =   750
      Left            =   5115
      TabIndex        =   27
      Top             =   3405
      Width           =   1095
      Begin VB.CommandButton cmdPrevia 
         Height          =   495
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptTraslado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vista previa"
         Top             =   180
         Width           =   495
      End
      Begin VB.CommandButton cmdImprime 
         Height          =   495
         Left            =   540
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptTraslado.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Rango de fechas de traslado"
      Height          =   750
      Left            =   120
      TabIndex        =   24
      Top             =   3405
      Width           =   3855
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   315
         Left            =   720
         TabIndex        =   11
         ToolTipText     =   "Fecha inicial"
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         ToolTipText     =   "Fecha final"
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   165
         TabIndex        =   26
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2010
         TabIndex        =   25
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   120
      TabIndex        =   15
      Top             =   645
      Width           =   7275
      Begin VB.TextBox txtNumCtaDestino 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         ToolTipText     =   "Número de cuenta origen"
         Top             =   2190
         Width           =   885
      End
      Begin VB.Frame Frame3 
         Height          =   405
         Left            =   1620
         TabIndex        =   22
         Top             =   1725
         Width           =   2865
         Begin VB.OptionButton optTipoDestino 
            Caption         =   "Externo"
            Height          =   195
            Index           =   2
            Left            =   1890
            TabIndex        =   9
            Top             =   150
            Width           =   900
         End
         Begin VB.OptionButton optTipoDestino 
            Caption         =   "Interno"
            Height          =   195
            Index           =   1
            Left            =   990
            TabIndex        =   8
            Top             =   150
            Width           =   900
         End
         Begin VB.OptionButton optTipoDestino 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.ComboBox cboPersona 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Seleccione la persona que realizó el traslado"
         Top             =   585
         Width           =   5490
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Seleccione el departamento"
         Top             =   210
         Width           =   5490
      End
      Begin VB.Frame Frame2 
         Height          =   405
         Left            =   1620
         TabIndex        =   20
         Top             =   915
         Width           =   2865
         Begin VB.OptionButton optTipoOrigen 
            Caption         =   "Externo"
            Height          =   195
            Index           =   2
            Left            =   1875
            TabIndex        =   5
            Top             =   150
            Width           =   855
         End
         Begin VB.OptionButton optTipoOrigen 
            Caption         =   "Interno"
            Height          =   195
            Index           =   1
            Left            =   975
            TabIndex        =   4
            Top             =   150
            Width           =   855
         End
         Begin VB.OptionButton optTipoOrigen 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   3
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.TextBox txtNumCtaOrigen 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         ToolTipText     =   "Número de cuenta origen"
         Top             =   1380
         Width           =   885
      End
      Begin VB.Label lblPacienteDestino 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   23
         Top             =   2190
         Width           =   4590
      End
      Begin VB.Label lblCuentaDestino 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta destino"
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   2250
         Width           =   1065
      End
      Begin VB.Label lblPacienteOrigen 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   19
         Top             =   1380
         Width           =   4590
      End
      Begin VB.Label lblCuentaOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta origen"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Persona"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   645
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   270
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmRptTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboHospital_Click()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset

    If cboHospital.ListIndex <> -1 Then

        vgstrParametrosSP = "-1|-1|*|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_GnSelDepartamento")
        
        pLlenarCboRs cboDepartamento, rs, 0, 1
        
        cboDepartamento.AddItem "<TODOS>", 0
        cboDepartamento.ItemData(cboDepartamento.NewIndex) = -1
        cboDepartamento.ListIndex = 0
        
        vgstrParametrosSP = "-1|-1|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_GnSelEmpleado")
        
        pLlenarCboRs cboPersona, rs, 0, 1
        
        cboPersona.AddItem "<TODOS>", 0
        cboPersona.ItemData(cboPersona.NewIndex) = -1
        cboPersona.ListIndex = 0
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboHospital_Click"))
End Sub

Private Sub cmdImprime_Click()
    On Error GoTo NotificaError


    pImprime "I"

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdImprime_Click"))
End Sub

Private Sub cmdPrevia_Click()
    On Error GoTo NotificaError


    pImprime "P"
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrevia_Click"))
End Sub

Private Sub pImprime(strDestino As String)
    On Error GoTo NotificaError

    Dim rsReporte As New ADODB.Recordset
    Dim rptReporte As CRAXDRT.Report
    Dim alstrParametros(3) As String

    Dim strFechaIni As String
    Dim strFechaFin As String
    Dim strCveDepto As String
    Dim strCveEmpleado As String
    Dim strCtaOrigen As String
    Dim strTipoOrigen As String
    Dim strCtaDestino As String
    Dim strTipoDestino As String
    
    If fblnDatosValidos() Then
    
        strFechaIni = fstrFechaSQL(mskFechaIni.Text)
        strFechaFin = fstrFechaSQL(mskFechaFin.Text)
        strCveDepto = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
        strCveEmpleado = CStr(cboPersona.ItemData(cboPersona.ListIndex))
        strCtaOrigen = IIf(Trim(txtNumCtaOrigen.Text) = "", "-1", txtNumCtaOrigen.Text)
        strTipoOrigen = IIf(optTipoOrigen(0).Value, "*", IIf(optTipoOrigen(1).Value, "I", "E"))
        strCtaDestino = IIf(Trim(txtNumCtaDestino.Text) = "", "-1", txtNumCtaDestino.Text)
        strTipoDestino = IIf(optTipoDestino(0).Value, "*", IIf(optTipoDestino(1).Value, "I", "E"))
            
        vgstrParametrosSP = _
        strFechaIni & _
        "|" & strFechaFin & _
        "|" & strCveDepto & _
        "|" & strCveEmpleado & _
        "|" & strCtaOrigen & _
        "|" & strTipoOrigen & _
        "|" & strCtaDestino & _
        "|" & strTipoDestino & _
        "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
            
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptTraslado")
        If rsReporte.RecordCount <> 0 Then
            pInstanciaReporte rptReporte, "rptTraslado.rpt"
            rptReporte.DiscardSavedData

            alstrParametros(0) = "NombreHospital" & ";" & Trim(cboHospital.List(cboHospital.ListIndex)) & ";TRUE"
            alstrParametros(1) = "FechaIni" & ";" & UCase(Format(mskFechaIni.Text, "dd/mmm/yyyy")) & ";TRUE"
            alstrParametros(2) = "FechaFin" & ";" & UCase(Format(mskFechaFin.Text, "dd/mmm/yyyy")) & ";TRUE"
            alstrParametros(3) = "Detallado" & ";" & CStr(chkDetallado.Value) & ";TRUE"
            
            pCargaParameterFields alstrParametros, rptReporte
            pImprimeReporte rptReporte, rsReporte, strDestino, "Traslados de cargos"
        Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rsReporte.Close
    
    End If



    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnDatosValidos = True

    If Val(txtNumCtaOrigen.Text) <> 0 And Trim(lblPacienteOrigen.Caption) = "" Then
        fblnDatosValidos = False
        'No se encontró la información del paciente.
        MsgBox SIHOMsg(355), vbOKOnly + vbInformation, "Mensaje"
        txtNumCtaOrigen.SetFocus
    End If
    If fblnDatosValidos And Val(txtNumCtaDestino.Text) <> 0 And Trim(lblPacienteDestino.Caption) = "" Then
        fblnDatosValidos = False
        'No se encontró la información del paciente.
        MsgBox SIHOMsg(355), vbOKOnly + vbInformation, "Mensaje"
        txtNumCtaDestino.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(mskFechaIni.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        mskFechaIni.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(mskFechaFin.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        mskFechaFin.SetFocus
    End If
    If fblnDatosValidos Then
        If CDate(mskFechaIni.Text) > fdtmServerFecha Then
            fblnDatosValidos = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
            mskFechaIni.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        If CDate(mskFechaFin.Text) > fdtmServerFecha Then
            fblnDatosValidos = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
            mskFechaFin.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        If CDate(mskFechaIni.Text) > CDate(mskFechaFin.Text) Then
            fblnDatosValidos = False
            '¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
            mskFechaIni.SetFocus
        End If
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function


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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    
    Dim rs As New ADODB.Recordset
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date

    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 334
    Case "SE"
         lngNumOpcion = 2009
    End Select
    
    pCargaHospital lngNumOpcion
    
    optTipoOrigen(0).Value = True
    optTipoDestino(0).Value = True
    
    dtmfecha = fdtmServerFecha
    
    mskFechaIni.Mask = ""
    mskFechaIni.Text = dtmfecha
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub mskFechaFin_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFechaFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaIni_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFechaIni

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaIni_GotFocus"))
End Sub

Private Sub optTipoDestino_Click(Index As Integer)
    On Error GoTo NotificaError


    lblCuentaDestino.Enabled = Index <> 0
    txtNumCtaDestino.Enabled = Index <> 0
    
    lblPacienteDestino.Caption = ""
    txtNumCtaDestino.Text = ""
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoDestino_Click"))
End Sub

Private Sub optTipoOrigen_Click(Index As Integer)
    On Error GoTo NotificaError


    lblCuentaOrigen.Enabled = Index <> 0
    txtNumCtaOrigen.Enabled = Index <> 0
        
    lblPacienteOrigen.Caption = ""
    txtNumCtaOrigen.Text = ""

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoOrigen_Click"))
End Sub

Private Sub txtNumCtaDestino_Change()
    On Error GoTo NotificaError


    lblPacienteDestino.Caption = ""

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaDestino_Change"))
End Sub

Private Sub txtNumCtaDestino_GotFocus()
    On Error GoTo NotificaError


    pSelTextBox txtNumCtaDestino

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaDestino_GotFocus"))
End Sub

Private Sub txtNumCtaDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    Dim blnEncontro As Boolean
    
    If KeyCode = vbKeyReturn Then
        blnEncontro = fblnNombrePaciente(lblPacienteDestino, txtNumCtaDestino, txtNumCtaDestino.Text, IIf(optTipoDestino(1).Value, "I", "E"))
        If Not blnEncontro Then
            pEnfocaTextBox txtNumCtaDestino
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaDestino_KeyDown"))
End Sub

Private Sub txtNumCtaDestino_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError


    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaDestino_KeyPress"))
End Sub

Private Sub txtNumCtaOrigen_Change()
    On Error GoTo NotificaError

    
    lblPacienteOrigen.Caption = ""

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaOrigen_Change"))
End Sub

Private Sub txtNumCtaOrigen_GotFocus()
    On Error GoTo NotificaError


    pSelTextBox txtNumCtaOrigen

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaOrigen_GotFocus"))
End Sub

Private Sub txtNumCtaOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    Dim blnEncontro As Boolean
    
    If KeyCode = vbKeyReturn Then
        blnEncontro = fblnNombrePaciente(lblPacienteOrigen, txtNumCtaOrigen, txtNumCtaOrigen.Text, IIf(optTipoOrigen(1).Value, "I", "E"))
        If Not blnEncontro Then
            pEnfocaTextBox txtNumCtaOrigen
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaOrigen_KeyDown"))
End Sub


Private Function fblnNombrePaciente(lblLabel As Label, txtCuenta As TextBox, strCuenta As String, strTipo As String) As Boolean
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim lngCuenta As Long

    fblnNombrePaciente = False
    
    If Trim(strCuenta) = "" Then
        With FrmBusquedaPacientes
            If strTipo = "E" Then
                .vgstrTipoPaciente = "E"
                .Caption = .Caption & " externos"
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSoloActivos.Enabled = True
                .optSinFacturar.Enabled = True
                .optTodos.Enabled = True
                .optSinFacturar.Value = True
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
          
            lngCuenta = .flngRegresaPaciente()
       End With
    Else
        lngCuenta = CLng(strCuenta)
    End If
    
    If lngCuenta <> -1 Then
        vgstrParametrosSP = Str(lngCuenta) & "|" & "0" & "|" & strTipo & "|" & vgintClaveEmpresaContable
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")
        If rs.RecordCount <> 0 Then
            txtCuenta.Text = lngCuenta
            lblLabel.Caption = rs!Nombre
            fblnNombrePaciente = True
        End If
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnNombrePaciente"))
End Function


Private Sub txtNumCtaOrigen_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError


    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumCtaOrigen_KeyPress"))
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

