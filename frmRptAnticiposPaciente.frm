VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptAnticiposPaciente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entradas y salidas de dinero"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmEstado 
      Caption         =   "Estado"
      Height          =   760
      Left            =   120
      TabIndex        =   16
      Top             =   2890
      Width           =   3855
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   130
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Selección de estado"
         Top             =   280
         Width           =   3555
      End
   End
   Begin VB.Frame FrmDepartamentos 
      Caption         =   "Departamentos"
      Height          =   760
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione el departamento"
         Top             =   280
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paciente"
      Height          =   1050
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3855
      Begin VB.ComboBox cboPaciente 
         Height          =   315
         Left            =   130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Selección del paciente"
         Top             =   570
         Width           =   3555
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Todos"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   280
         Width           =   900
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Externo"
         Height          =   210
         Index           =   1
         Left            =   2100
         TabIndex        =   5
         Top             =   280
         Width           =   945
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Interno"
         Height          =   210
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   280
         Width           =   945
      End
   End
   Begin VB.Frame frmFechas 
      Caption         =   "Rango de fechas"
      Height          =   760
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   3855
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   330
         Left            =   600
         TabIndex        =   1
         ToolTipText     =   "Fecha inicial para el reporte"
         Top             =   280
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   330
         Left            =   2400
         TabIndex        =   2
         ToolTipText     =   "Fecha final para el reporte"
         Top             =   285
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Al"
         Height          =   190
         Left            =   2040
         TabIndex        =   14
         Top             =   355
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   190
         Left            =   130
         TabIndex        =   13
         Top             =   355
         Width           =   735
      End
   End
   Begin VB.Frame frmBotonera 
      Height          =   690
      Left            =   1492
      TabIndex        =   10
      Top             =   3700
      Width           =   1110
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   555
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptAnticiposPaciente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir el reporte"
         Top             =   135
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptAnticiposPaciente.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Vista preliminar del reporte"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmRptAnticiposPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vlstrx As String
Dim rs As New ADODB.Recordset
Dim vglngCuentaPaciente As Long
Public vglngNumeroOpcion As Long
Private vgrptReporte As CRAXDRT.Report
Public vlblnTodosDeptos As Boolean


Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError
    pImprime "I"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprimir_Click"))
End Sub

Private Sub cmdVistaPreliminar_Click()
    On Error GoTo NotificaError
    pImprime "P"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVistaPreliminar_Click"))
End Sub

Private Sub pImprime(vlstrDestino As String)
    On Error GoTo NotificaError

    Dim vllngCuenta As Long
    Dim alstrParametros(11) As String
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    
    If fblnVerificaDatos Then
        
        vllngCuenta = 0
        
        If cboPaciente.ListCount <> 0 Then
            vllngCuenta = cboPaciente.ItemData(cboPaciente.ListIndex)
        End If
        
        vlstrx = fstrFechaSQL(txtFechaInicio.Text)
        vlstrx = vlstrx & "|" & fstrFechaSQL(txtFechaFin.Text)
        vlstrx = vlstrx & "|" & vllngCuenta
        vlstrx = vlstrx & "|" & IIf(optTipoPaciente(2).Value, "*", IIf(optTipoPaciente(0).Value, "I", "E"))
        vlstrx = vlstrx & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex)
        vlstrx = vlstrx & "|" & vgintClaveEmpresaContable
        vlstrx = vlstrx & "|" & cboEstado.ListIndex

        Set rsReporte = frsEjecuta_SP(vlstrx, "Sp_PVAnticipoPaciente")
        If rsReporte.RecordCount > 0 Then
        
          pInstanciaReporte vgrptReporte, "rptPVAnticipoPaciente.rpt"
          vgrptReporte.DiscardSavedData
        
          alstrParametros(0) = "p_empresa;" & Trim(vgstrNombreHospitalCH)
          alstrParametros(4) = "p_finicio;" & UCase(Format(txtFechaInicio, "dd/mmm/yyyy"))   'fstrFechaSQL(txtFechaInicio.Text, "")
          alstrParametros(5) = "p_ffin;" & UCase(Format(txtFechaFin, "dd/mmm/yyyy"))    'fstrFechaSQL(txtFechaFin.Text, "")
          If optTipoPaciente(0).Value Then
              alstrParametros(6) = "p_tipopaciente;" & "Internos"
          Else
              If optTipoPaciente(1).Value Then
                  alstrParametros(6) = "p_tipopaciente;" & "Externos"
              Else
                  alstrParametros(6) = "p_tipopaciente;" & "<TODOS>"
              End If
          End If
          alstrParametros(7) = "p_paciente;" & Trim(cboPaciente.List(cboPaciente.ListIndex))
          alstrParametros(11) = "p_tiporpt;" & "ENTRADAS Y SALIDAS DE DINERO"
          
          pCargaParameterFields alstrParametros, vgrptReporte
          pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Entradas y salidas de dinero"
          
        Else
          'No existe información con esos parámetros.
          MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        If rsReporte.State <> adStateClosed Then rsReporte.Close
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If
    If KeyAscii = 13 Then
       SendKeys vbTab
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    
    Set rs = frsEjecuta_SP("-1|1|*|" & vgintClaveEmpresaContable, "Sp_Gnseldepartamento")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboDepartamento, rs, 0, 1
    End If
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ItemData(cboDepartamento.newIndex) = 0
    cboDepartamento.ListIndex = 0
    
    pLlenaCboEstado
    
    pInicializa
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    cboPaciente.Clear
    cboPaciente.AddItem "<TODOS>"
    cboPaciente.ItemData(cboPaciente.newIndex) = 0
    cboPaciente.ListIndex = 0
End Sub

Private Sub optTipoPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If optTipoPaciente(0).Value Then
            '******************************************************
            ' Internos
            '******************************************************
            vgstrEstatusPaciente = "I"
            With FrmBusquedaPacientes
                .vgstrTipoPaciente = "I"
                .vgblnPideClave = True
                .Caption = .Caption & " Internos"
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSinFacturar.Value = True
                .optSinFacturar.Enabled = True
                .optSoloActivos.Enabled = True
                .optTodos.Enabled = False
                .optSoloActivos.Value = True
                .vgStrOtrosCampos = ", ccEmpresa.vchDescripcion as Empresa, " & _
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
                " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Teléfono "
                .vgstrTamanoCampo = "950,3400,4100,990,980"
            End With
            If FrmBusquedaPacientes.vgstrMovCve = "M" Then 'si es por número de cuenta
                DoEvents
                vglngCuentaPaciente = FrmBusquedaPacientes.flngRegresaPaciente() 'número de cuenta del paciente
                If vglngCuentaPaciente = -1 Then
                
                    cboPaciente.Clear
                    cboPaciente.AddItem ""
                    cboPaciente.ItemData(cboPaciente.newIndex) = 0
                    cboPaciente.ListIndex = 0
                    
                    If fblnCanFocus(optTipoPaciente(2)) Then optTipoPaciente(2).SetFocus
                    
                Else
                    pAgregaPacienteInterno vglngCuentaPaciente
                End If
            Else 'se es por clave de paciente
                vglngCvePaciente = FrmBusquedaPacientes.flngRegresaPaciente() 'clave del paciente
                If vglngCvePaciente = -1 Then
                    
                    cboPaciente.Clear
                    cboPaciente.AddItem ""
                    cboPaciente.ItemData(cboPaciente.newIndex) = 0
                    cboPaciente.ListIndex = 0
                    
                    If fblnCanFocus(optTipoPaciente(2)) Then optTipoPaciente(2).SetFocus
                Else
                    pAgregaPacienteInterno vglngCuentaPaciente
                End If
            End If
        ElseIf optTipoPaciente(1).Value Then
            '******************************************************
            ' Externos
            '******************************************************
            vgstrEstatusPaciente = "E"
            With FrmBusquedaPacientes
                .vgstrTipoPaciente = "E"
                .Caption = .Caption & " Externos"
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSoloActivos.Enabled = False
                .optTodos.Value = True
                .vgStrOtrosCampos = ", ccEmpresa.vchDescripcion as Empresa, " & _
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
                " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Teléfono "
                .vgstrTamanoCampo = "950,3400,4100,990,980"
            End With
            vglngCuentaPaciente = FrmBusquedaPacientes.flngRegresaPaciente()
            
            If vglngCuentaPaciente = -1 Then
            
                cboPaciente.Clear
                cboPaciente.AddItem ""
                cboPaciente.ItemData(cboPaciente.newIndex) = 0
                cboPaciente.ListIndex = 0
                
                If fblnCanFocus(optTipoPaciente(2)) Then optTipoPaciente(2).SetFocus
                
            Else
                pAgregaPacienteExterno vglngCuentaPaciente
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_KeyPress"))
End Sub

Private Sub pAgregaPacienteExterno(vllngCuenta As Long)
    On Error GoTo NotificaError

    cboPaciente.Clear

    vlstrx = "" & _
    "select " & _
        "ltrim(rtrim(Externo.chrApePaterno))||' '||ltrim(rtrim(Externo.chrApeMaterno))||' '||ltrim(rtrim(Externo.chrNombre)) Nombre," & _
        "RegistroExterno.intNumCuenta Cuenta " & _
    "From " & _
        "RegistroExterno " & _
        "inner join Externo on RegistroExterno.intNumPaciente=Externo.intNumPaciente " & _
    "Where " & _
        "RegistroExterno.intNumCuenta = " & Str(vllngCuenta)
    
    Set rs = frsRegresaRs(vlstrx)
    If rs.RecordCount <> 0 Then
        cboPaciente.AddItem rs!Nombre, 0
        cboPaciente.ItemData(cboPaciente.newIndex) = rs!cuenta
        cboPaciente.ListIndex = 0
        
        If fblnCanFocus(cboPaciente) Then cboPaciente.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregaPacienteExterno"))
End Sub

Private Sub pAgregaPacienteInterno(vllngCuenta As Long)
    On Error GoTo NotificaError

    cboPaciente.Clear

    vlstrx = "" & _
    "select " & _
        "ltrim(rtrim(AdPaciente.vchApellidoPaterno))||' '||ltrim(rtrim(AdPaciente.vchApellidoMaterno))||' '||ltrim(rtrim(AdPaciente.vchNombre)) Nombre," & _
        "AdAdmision.numNumCuenta Cuenta " & _
    "From " & _
        "AdAdmision " & _
        "inner join AdPaciente on AdAdmision.numCvePaciente=AdPaciente.numCvePaciente " & _
    "Where " & _
        "AdAdmision.numNumCuenta = " & Str(vllngCuenta)
    
    Set rs = frsRegresaRs(vlstrx)
    If rs.RecordCount <> 0 Then
        cboPaciente.AddItem rs!Nombre, 0
        cboPaciente.ItemData(cboPaciente.newIndex) = rs!cuenta
        cboPaciente.ListIndex = 0
        
        If fblnCanFocus(cboPaciente) Then cboPaciente.SetFocus
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregaPacienteInterno"))
End Sub

Private Sub txtFechaFin_GotFocus()
'--------------------------------------------------------------------------
' Procedimiento para que cada vez que tenga el enfoque el control, lo marque
' en azul o seleccionado
'--------------------------------------------------------------------------
    On Error GoTo NotificaError 'Manejo del error
    
    pSelMkTexto txtFechaFin
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaFin_GotFocus"))
End Sub

Private Sub txtFechaFin_LostFocus()
    On Error GoTo NotificaError

    If Trim(txtFechaFin.ClipText) = "" Then
        txtFechaFin.Mask = ""
        txtFechaFin.Text = fdtmServerFecha
        txtFechaFin.Mask = "##/##/####"
    Else
        If Not IsDate(txtFechaFin.Text) Then
            txtFechaFin.Mask = ""
            txtFechaFin.Text = fdtmServerFecha
            txtFechaFin.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFolioIni_KeyPress"))
End Sub

Private Sub txtFechaInicio_GotFocus()
'--------------------------------------------------------------------------
' Procedimiento para que cada vez que tenga el enfoque el control, lo marque
' en azul o seleccionado
'--------------------------------------------------------------------------
    On Error GoTo NotificaError 'Manejo del error
    
    pSelMkTexto txtFechaInicio

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaInicio_GotFocus"))
End Sub

Private Sub txtFechaInicio_LostFocus()
    On Error GoTo NotificaError

    If Trim(txtFechaInicio.ClipText) = "" Then
        txtFechaInicio.Mask = ""
        txtFechaInicio.Text = fdtmServerFecha - 8
        txtFechaInicio.Mask = "##/##/####"
    Else
        If Not IsDate(txtFechaInicio.Text) Then
            txtFechaInicio.Mask = ""
            txtFechaInicio.Text = fdtmServerFecha - 8
            txtFechaInicio.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFolioIni_KeyPress"))
End Sub

Private Sub pInicializa()
    On Error GoTo NotificaError
    
    txtFechaInicio.Mask = ""
    txtFechaInicio.Text = fdtmServerFecha - 8
    txtFechaInicio.Mask = "##/##/####"
    
    txtFechaFin.Mask = ""
    txtFechaFin.Text = fdtmServerFecha
    txtFechaFin.Mask = "##/##/####"
    
    optTipoPaciente(2).Value = True
           
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pInicializa"))
    Unload Me
End Sub

Private Function fblnVerificaDatos() As Boolean
Dim rsVerifica As New ADODB.Recordset

    On Error GoTo NotificaError
    
    fblnVerificaDatos = True
    
    If CDate(txtFechaFin.Text) < CDate(txtFechaInicio.Text) Then
        MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
        If fblnCanFocus(txtFechaInicio) Then
            If fblnCanFocus(txtFechaFin) Then txtFechaFin.SetFocus
        End If
        fblnVerificaDatos = False
    End If
    
    Exit Function

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnVerificaDatos"))
    Unload Me
End Function

Private Function pLlenaCboEstado()
 On Error GoTo NotificaError
    'Se agregan los elementos del combo de estado
    cboEstado.AddItem "<TODOS>", 0
    cboEstado.ItemData(cboEstado.newIndex) = 0
    cboEstado.AddItem "ACTIVO", 1
    cboEstado.ItemData(cboEstado.newIndex) = 1
    cboEstado.AddItem "FACTURADO", 2
    cboEstado.ItemData(cboEstado.newIndex) = 2
    cboEstado.AddItem "CANCELADO", 3
    cboEstado.ItemData(cboEstado.newIndex) = 3
    cboEstado.ListIndex = 0
    
           
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCboEstado"))
    Unload Me
End Function

