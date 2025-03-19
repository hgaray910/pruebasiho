VERSION 5.00
Begin VB.Form frmRptCargosExcedentesEnSA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos excedentes de la suma asegurada"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3300
      TabIndex        =   19
      Top             =   3300
      Width           =   1140
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         Picture         =   "frmRptCargosExcedentesEnSA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   570
         Picture         =   "frmRptCargosExcedentesEnSA.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame fraDatosPaciente 
      Height          =   3300
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   7530
      Begin VB.TextBox txtObservaciones 
         Height          =   870
         Left            =   165
         MaxLength       =   800
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Observaciones a imprimir en el reporte"
         Top             =   2300
         Width           =   7160
      End
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   0
         ToolTipText     =   "Número de cuenta del paciente"
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtPaciente 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Nombre del paciente"
         Top             =   600
         Width           =   5700
      End
      Begin VB.TextBox txtEmpresaPaciente 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Nombre de la empresa del paciente"
         Top             =   960
         Width           =   5700
      End
      Begin VB.TextBox txtTipoPaciente 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Tipo de paciente"
         Top             =   1320
         Width           =   5700
      End
      Begin VB.TextBox txtFechaInicial 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Fecha de inicio de atención"
         Top             =   1680
         Width           =   1710
      End
      Begin VB.TextBox txtFechaFinal 
         Height          =   315
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Fecha final de atención"
         Top             =   1680
         Width           =   1710
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   2895
         TabIndex        =   7
         Top             =   180
         Width           =   1950
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Externo"
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   2
            Top             =   150
            Width           =   855
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Interno"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Top             =   135
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   2080
         Width           =   1100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de cuenta"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   650
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   1010
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   1370
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de atención"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1730
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   3390
         TabIndex        =   13
         Top             =   1730
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmRptCargosExcedentesEnSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rptReporte As CRAXDRT.Report
Dim vlblnLimpiar As Boolean
Dim lngCveEmpresaPaciente As Long


Private Sub pLimpia()
    On Error GoTo NotificaError
        
    txtPaciente.Text = ""
    txtEmpresaPaciente.Text = ""
    txtTipoPaciente.Text = ""
    txtFechaFinal.Text = ""
    txtFechaInicial.Text = ""
    txtObservaciones.Text = ""

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub
Private Sub cmdPreview_Click()
    pImprime "P"
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    pInstanciaReporte rptReporte, "rptCargosExcedentesEnSA.rpt"
    vlblnLimpiar = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub optTipoPaciente_Click(Index As Integer)
pEnfocaTextBox txtMovimientoPaciente
End Sub

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If Asc(UCase(Chr(KeyAscii))) = vbKeyI Then
            optTipoPaciente(0).Value = True
        ElseIf Asc(UCase(Chr(KeyAscii))) = vbKeyE Then
            optTipoPaciente(1).Value = True
        End If
        KeyAscii = 7
    End If

End Sub

Private Sub txtObservaciones_GotFocus()
    pSelTextBox txtObservaciones
End Sub

Private Sub txtMovimientoPaciente_GotFocus()
    
    On Error GoTo NotificaError
    
    pSelTextBox txtMovimientoPaciente
    
    If vlblnLimpiar Then
        pLimpia
    Else
        vlblnLimpiar = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_GotFocus"))
    
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
Dim rsDatosPaciente As New ADODB.Recordset
Dim intAseguradora As Integer
Dim lngCuenta As Long
    
    If KeyCode = vbKeyReturn Then
    
        If Trim(txtMovimientoPaciente.Text) = "" Then
        
            txtPaciente.Text = ""
            txtEmpresaPaciente.Text = ""
            txtTipoPaciente.Text = ""
            txtFechaInicial.Text = ""
            txtFechaFinal.Text = ""
        
            With FrmBusquedaPacientes
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                
                .optTodos.Value = True
                .optSinFacturar.Enabled = False
                .optSoloActivos.Enabled = False
                .optTodos.Enabled = True
                .vgIntMaxRecords = 200
                
                If optTipoPaciente(1).Value Then 'Externos
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " Externos"
                Else
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                    .vgstrTipoPaciente = "I"
                    .Caption = .Caption & " Internos"
                End If
        
                txtMovimientoPaciente.Text = .flngRegresaPaciente()
                If txtMovimientoPaciente <> -1 Then
                    vlblnLimpiar = False
                    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                Else
                    txtMovimientoPaciente.Text = ""
                End If
            End With
            
        Else
        
            vgstrParametrosSP = Val(txtMovimientoPaciente.Text) & "|" & "0" & "|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|" & vgintClaveEmpresaContable
            Set rsDatosPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDATOSPACIENTE")
            
            txtPaciente.Text = ""
            txtEmpresaPaciente.Text = ""
            txtTipoPaciente.Text = ""
            txtFechaInicial.Text = ""
            txtFechaFinal.Text = ""
            
            With rsDatosPaciente
                If .RecordCount <> 0 Then
                
                    intAseguradora = IIf((!bitUtilizaConvenio = 1 And !Aseguradora = 1), 1, 0)
                    
                    If intAseguradora <> 1 Then
                        
                        '¡La cuenta no es de convenio con aseguradora!
                        MsgBox SIHOMsg(1105), vbExclamation, "Mensaje"
                        txtMovimientoPaciente.Text = ""
                        Exit Sub
                        
                    Else
                        
'                        lngCveEmpresaPaciente = IIf(IsNull(!intcveempresa), 0, !intcveempresa)
'                        txtMovimientoPaciente.Text = lngCuenta
                        txtPaciente.Text = !Nombre
                        txtEmpresaPaciente.Text = IIf(IsNull(!Empresa), "", !Empresa)
                        txtTipoPaciente.Text = !Tipo
                        
                        txtFechaInicial.Text = ""
                        If Not IsNull(!Ingreso) Then
                            txtFechaInicial.Text = Format(!Ingreso, "dd/mmm/yyyy hh:mm")
                        End If
                        txtFechaFinal.Text = ""
                        If Not IsNull(!Egreso) Then
                            txtFechaFinal.Text = Format(!Egreso, "dd/mmm/yyyy hh:mm")
                        End If
                        
                        txtObservaciones.Text = ""
                    
                    End If
                    txtObservaciones.SetFocus
                    
                Else
                    
                    '¡La información no existe!
                    MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                    pEnfocaTextBox txtMovimientoPaciente
                    
                End If
                .Close
            End With
            
            
        End If
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
End Sub

Private Sub pImprime(strDestino As String)
On Error GoTo NotificaError
Dim rsReporte As New ADODB.Recordset
Dim rsDatosPaciente As New ADODB.Recordset
Dim rsExcedente As New ADODB.Recordset
Dim alstrParametros(25) As String

' Carga datos del paciente
vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & "|0|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|" & vgintClaveEmpresaContable
Set rsDatosPaciente = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")

If rsDatosPaciente.EOF Then
    'No existe información con esos parámetros.
    MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
Else
   EntornoSIHO.ConeccionSIHO.BeginTrans
    
    vgstrParametrosSP = Val(txtMovimientoPaciente.Text) & "|" & IIf(optTipoPaciente(0).Value, "I", "E")
    Set rsExcedente = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelExcedenteSumaAsegurada")
    
    If Not rsExcedente.EOF Then
        vgstrParametrosSP = Val(txtMovimientoPaciente.Text) & "|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|" & rsExcedente!excedente
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptCargosExcedentesEnSA")
        If rsReporte.RecordCount <> 0 Then
    
            rptReporte.DiscardSavedData
            alstrParametros(0) = "Fecha;" & fdtmServerFecha
            alstrParametros(1) = "Hora;" & fdtmServerHora
            alstrParametros(2) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
            alstrParametros(3) = "DireccionHospital;" & Trim(vgstrDireccionCH) & " " & Trim(vgstrColoniaCH) & " " & Trim(vgstrCiudadCH)
            alstrParametros(4) = "RFC;" & "RFC " & vgstrRfCCH
            alstrParametros(5) = "NúmeroPaciente;" & rsDatosPaciente!NumPaciente
            alstrParametros(6) = "Cuenta;" & txtMovimientoPaciente.Text
            alstrParametros(7) = "Tipo;" & IIf(optTipoPaciente(0).Value, "INTERNO", "EXTERNO")
            alstrParametros(8) = "Nombre;" & IIf(IsNull(rsDatosPaciente!Nombre), "", rsDatosPaciente!Nombre)
            alstrParametros(9) = "Domicilio;" & IIf(IsNull(rsDatosPaciente!Domicilio), "", rsDatosPaciente!Domicilio)
            alstrParametros(10) = "Ciudad;" & IIf(IsNull(rsDatosPaciente!Ciudad), "", rsDatosPaciente!Ciudad)
            alstrParametros(11) = "Estado;" & IIf(IsNull(rsDatosPaciente!Estado), "", rsDatosPaciente!Estado)
            alstrParametros(12) = "FechaNacimiento;"
            If Not IsNull(rsDatosPaciente!FechaNacimiento) Then
                alstrParametros(12) = "FechaNacimiento;" & Format(rsDatosPaciente!FechaNacimiento, "dd/mmm/yyyy")
            End If
            alstrParametros(13) = "Edad;"
            If Not IsNull(rsDatosPaciente!FechaNacimiento) Then
                alstrParametros(13) = "Edad;" & fstrObtieneEdad(rsDatosPaciente!FechaNacimiento, rsDatosPaciente!Ingreso)
            End If
            alstrParametros(14) = "FechaIngreso;"
            If Not IsNull(rsDatosPaciente!Ingreso) Then
                alstrParametros(14) = "FechaIngreso;" & Format(rsDatosPaciente!Ingreso, "dd/mmm/yyyy hh:mm")
            End If
            alstrParametros(15) = "FechaEgreso;"
            If Not IsNull(rsDatosPaciente!Egreso) Then
               alstrParametros(15) = "FechaEgreso;" & Format(rsDatosPaciente!Egreso, "dd/mmm/yyyy hh:mm")
            End If
            alstrParametros(16) = "UltimoCuarto;" & IIf(IsNull(rsDatosPaciente!Cuarto), "", rsDatosPaciente!Cuarto)
            alstrParametros(17) = "Responsable;" & IIf(IsNull(rsDatosPaciente!Responsable), "", rsDatosPaciente!Responsable)
            alstrParametros(18) = "MedicoTratante;" & IIf(IsNull(rsDatosPaciente!Medico), "", rsDatosPaciente!Medico)
            alstrParametros(19) = "TipoPaciente;" & IIf(IsNull(rsDatosPaciente!Tipo), "", rsDatosPaciente!Tipo)
            alstrParametros(20) = "Empresa;" & IIf(IsNull(rsDatosPaciente!Empresa), "", rsDatosPaciente!Empresa)
            alstrParametros(21) = "telefono;" & rsDatosPaciente!TelefonoPaciente
            alstrParametros(22) = "Comentario;" & UCase(txtObservaciones.Text)
            alstrParametros(23) = "diagnóstico;" & IIf(IsNull(rsDatosPaciente!Diagnostico), "", rsDatosPaciente!Diagnostico)
            alstrParametros(24) = "EXCEDENTE;" & IIf(IsNull(rsExcedente!excedente), "0.00", rsExcedente!excedente)
            alstrParametros(25) = "IVA;" & IIf(IsNull(rsExcedente!IVA), "0.00", rsExcedente!IVA)
                    
            pCargaParameterFields alstrParametros, rptReporte
            pImprimeReporte rptReporte, rsReporte, strDestino, "Cargos excedentes de la suma asegurada"
        
        Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rsReporte.Close
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    EntornoSIHO.ConeccionSIHO.CommitTrans
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdPreview.SetFocus
    End If
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtObservaciones_KeyPress"))
End Sub
