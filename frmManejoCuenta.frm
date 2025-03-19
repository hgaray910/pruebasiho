VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmManejoCuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manejo de cuentas de pacientes"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechasFactura 
      Caption         =   "Ingreso/Egreso"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5780
      TabIndex        =   35
      ToolTipText     =   "Modificar fechas de ingreso o egreso"
      Top             =   3220
      Width           =   1830
   End
   Begin MSMask.MaskEdBox mskHoraEgreso 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "hh:mm AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   4
      EndProperty
      Height          =   315
      Left            =   5190
      TabIndex        =   33
      Top             =   4400
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:mm AM/PM"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskHoraIngreso 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "hh:mm AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   5190
      TabIndex        =   32
      Top             =   3960
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:mm AM/PM"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFechaEgreso 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   1930
      TabIndex        =   31
      Top             =   4400
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame frmModificaFechas 
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   3690
      Width           =   7575
      Begin MSMask.MaskEdBox mskFechaIngreso 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1810
         TabIndex        =   30
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         Caption         =   "Hora egreso"
         Height          =   435
         Left            =   3770
         TabIndex        =   29
         Top             =   670
         Width           =   1340
      End
      Begin VB.Label Label12 
         Caption         =   "Hora ingreso"
         Height          =   435
         Left            =   3770
         TabIndex        =   28
         Top             =   250
         Width           =   1340
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha egreso"
         Height          =   435
         Left            =   220
         TabIndex        =   27
         Top             =   670
         Width           =   1440
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha ingreso"
         Height          =   435
         Left            =   220
         TabIndex        =   26
         Top             =   250
         Width           =   1440
      End
   End
   Begin VB.TextBox txtAfiliacion 
      CausesValidation=   0   'False
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1935
      MaxLength       =   14
      TabIndex        =   24
      ToolTipText     =   "Número de afiliación"
      Top             =   2190
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   3050
      Width           =   7575
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar afiliación"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3820
         TabIndex        =   18
         ToolTipText     =   "Actualizar el número de afiliación"
         Top             =   180
         Width           =   1830
      End
      Begin VB.CommandButton cmdFacturada 
         Caption         =   "Poner como facturada"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1900
         TabIndex        =   17
         ToolTipText     =   "Cambiar el estado de facturación de la cuenta"
         Top             =   180
         Width           =   1920
      End
      Begin VB.CommandButton cmdEstatusAdm 
         Caption         =   "Cerrar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   70
         TabIndex        =   16
         ToolTipText     =   "Cerrar o abrir la cuenta"
         Top             =   180
         Width           =   1830
      End
   End
   Begin VB.Frame FrePaciente 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtFechaCierre 
         Height          =   315
         Left            =   5080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2430
         Width           =   2210
      End
      Begin VB.TextBox txtFechaApertura 
         Height          =   315
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2430
         Width           =   1590
      End
      Begin VB.TextBox txtEstatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5080
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Estado actual de la cuenta"
         Top             =   1710
         Width           =   2210
      End
      Begin VB.TextBox txtTipoPaciente 
         Height          =   315
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Procedencia del paciente"
         Top             =   990
         Width           =   5475
      End
      Begin VB.TextBox txtEmpresaPaciente 
         Height          =   315
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Empresa"
         Top             =   1350
         Width           =   5475
      End
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Introduzca el número de cuenta"
         Top             =   255
         Width           =   1590
      End
      Begin VB.TextBox txtPaciente 
         Height          =   315
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Nombre del paciente"
         Top             =   630
         Width           =   5475
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Interno"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   4
         ToolTipText     =   "Paciente hospitalizado"
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Externo"
         Height          =   255
         Index           =   1
         Left            =   5080
         TabIndex        =   3
         ToolTipText     =   "Paciente no hospitalizado"
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtCuarto 
         Height          =   315
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Cuarto actual o último"
         Top             =   1710
         Width           =   1590
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha de cierre"
         Height          =   195
         Left            =   3840
         TabIndex        =   23
         Top             =   2490
         Width           =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha de apertura"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   2490
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número de afiliación"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   2130
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   1770
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1050
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1410
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   690
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de cuenta"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cuarto"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1770
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   3577
      TabIndex        =   34
      Top             =   4915
      Width           =   660
      Begin VB.CommandButton cmdGrabaRegistro 
         Height          =   495
         Left            =   60
         Picture         =   "frmManejoCuenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Guardar el registro"
         Top             =   140
         UseMaskColor    =   -1  'True
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmManejoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmManejoCuenta                                        -
'-------------------------------------------------------------------------------------
'| Objetivo: Cerrar y activar cuentas, así como cambiar estatus de facturada la cuenta
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 12/Oct/2001
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : Hoy
'| ultimas modificaciones, especificar:
'-------------------------------------------------------------------------------------

Option Explicit
Dim vgstrEstatusAdm As String
Dim lintCuentaFacturada As Integer
Dim vgstrEstadoManto As String
Dim vllngCuentaCerrada As Long
Dim vlintNumDiasAbrirExt As Integer
Dim vlintNumDiasAbrirInt As Integer
Dim vldtmFechaIngreso As Date
Dim lblnPermitirCerrarCuenta As Boolean

Private Sub cmdActualizar_Click()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim vllngPersonaGraba As Long
    Dim SQL As String
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
            
            vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & Trim(txtAfiliacion.Text)
            frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDNUMEROAFILIACION"
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "ACTUALIZAR EL NUMERO DE AFILIACION DE LA CUENTA ", Trim(txtMovimientoPaciente.Text))
            
            SQL = "delete from pvTipoPacienteProceso where pvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
                "and pvTipoPacienteProceso.intproceso = " & enmTipoProceso.AbrirCerrarCuentas
            pEjecutaSentencia SQL
            
            SQL = "insert into pvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.AbrirCerrarCuentas & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ")"
            pEjecutaSentencia SQL
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbInformation, "Mensaje"
        pCancelar
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdActualizar_Click"))
    Unload Me
End Sub

Private Sub cmdEstatusAdm_Click()
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim vllngPersonaGraba As Long
    Dim SQL As String
    Dim intDias As Integer
    Dim blnPermisoEscritura As Boolean
    Dim blnPermisoCTotal As Boolean
    Dim vllngVar As Long
    Dim vldtmFechaAntes As Date
    Dim vldtmFechaDespues As Date
    Dim rsCargos As New ADODB.Recordset
    Dim vlmsgCargosProgramados As String
    Dim rsPostergado As ADODB.Recordset
    Dim rsCargosPostergadosSinFact As ADODB.Recordset 'Validar cargos sin facturar, en cuentas postergadas
    Dim rsConceptosSeguroSinFact 'Validar conceptos de seguro sin facturar, en cuentas postergadas
        
    If vllngCuentaCerrada = 1 Then
        intDias = DateDiff("D", vldtmFechaIngreso, fdtmServerFechaHora)
        If fblnRevisaPermiso(vglngNumeroLogin, 1918, "E") Then
        'Si tiene permiso de escritura o control total
        
            If Not fblnRevisaPermiso(vglngNumeroLogin, 1918, "C") Then
            'Si no tiene permiso de control total
                If lintCuentaFacturada = 0 Then
                    If intDias > IIf(OptTipoPaciente(0), vlintNumDiasAbrirInt, vlintNumDiasAbrirExt) Then
                        'No puede abrir cuenta si ya se pasó de los días permitidos
                        MsgBox SIHOMsg(742), vbInformation, "Mensaje"
                        Exit Sub
                    End If
                Else
                    'El usuario no tiene permiso para realizar esta operación.
                    MsgBox SIHOMsg(635), vbInformation, "Mensaje"
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
    Else
        'si existen requisiciones pendientes
        If fblnRequisicionPaciente(Val(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E")) Then
            frmRequisicionesPendientes.pMostrarRequisiciones Val(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E"), lblnPermitirCerrarCuenta
            If Not frmRequisicionesPendientes.lblnContinuarCerrarCuenta Then
                Unload frmRequisicionesPendientes
                Exit Sub
            End If
            Unload frmRequisicionesPendientes
        End If
    End If
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
            
        Set rsPostergado = frsRegresaRs("SELECT BITPOSTERGADA, INTCUENTACERRADA FROM EXPACIENTEINGRESO WHERE INTNUMCUENTA = " & txtMovimientoPaciente.Text, adLockOptimistic, adOpenDynamic)
        If rsPostergado.RecordCount > 0 Then
            If rsPostergado!BITPOSTERGADA = 1 And rsPostergado!INTCUENTACERRADA = 1 Then
                MsgBox "No es posible abrir la cuenta, se encuentra postergada.", vbInformation + vbOKOnly, "Mensaje"
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                Exit Sub
            Else
                Set rsCargosPostergadosSinFact = frsRegresaRs("SELECT pvc.* FROM PVCARGO pvc JOIN PVCARGOPOSTERGADO pvcp ON pvc.INTNUMCARGO = pvcp.INTNUMCARGO WHERE pvc.INTMOVPACIENTE = " & txtMovimientoPaciente.Text & " AND NVL(pvc.CHRFOLIOFACTURA, ' ') = ' '", adLockOptimistic, adOpenDynamic)
                If rsCargosPostergadosSinFact.RecordCount > 0 Then
                    MsgBox "No es posible abrir la cuenta, fue postergada y no se ha facturado completamente.", vbInformation + vbOKOnly, "Mensaje"
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    Exit Sub
                Else
                    Set rsConceptosSeguroSinFact = frsRegresaRs("SELECT pvcap.* FROM PVCONTROLASEGURADORAPOSTERGADO pvcap WHERE pvcap.INTMOVPACIENTE = " & txtMovimientoPaciente.Text, adLockOptimistic, adOpenDynamic)
                    If rsConceptosSeguroSinFact.RecordCount > 0 Then
                        MsgBox "No es posible abrir la cuenta, fue postergada y no se ha facturado completamente.", vbInformation + vbOKOnly, "Mensaje"
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        Exit Sub
                    End If
                End If
            End If
        End If
'        If rsPostergado.ActiveConnection = True Then
'            rsPostergado.Close
'        End If
            
        vlmsgCargosProgramados = ""
        
        If vllngCuentaCerrada = 0 Then
            vldtmFechaAntes = fdtmServerFechaHora
            vllngVar = 1
            frsEjecuta_SP Trim(txtMovimientoPaciente.Text) & "|'" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'|" & fstrFechaSQL(CStr(vldtmFechaAntes), CStr(vldtmFechaAntes)), "FN_PVINSCARGOSPROGRAMADOS", True, vllngVar
            vldtmFechaDespues = fdtmServerFechaHora
            
            Set rsCargos = frsEjecuta_SP(Trim(txtMovimientoPaciente.Text) & "|'" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'|" & fstrFechaSQL(CStr(vldtmFechaAntes), CStr(vldtmFechaAntes)) & "|" & fstrFechaSQL(CStr(vldtmFechaDespues), CStr(vldtmFechaDespues)), "SP_PVSELCARGOSPROGRAMADOSBITA")
            If rsCargos.RecordCount > 0 Then
                rsCargos.MoveFirst
                Do While Not rsCargos.EOF
                    vlmsgCargosProgramados = IIf(Trim(vlmsgCargosProgramados) = "", rsCargos!FechaHora & Chr(9) & rsCargos!Cargo & Chr(9) & rsCargos!Descripcion, vlmsgCargosProgramados & Chr(13) & rsCargos!FechaHora & Chr(9) & rsCargos!Cargo & Chr(9) & rsCargos!Descripcion)
                    rsCargos.MoveNext
                Loop
            End If
            
            pEjecutaSentencia "UPDATE PVCARGOPROGRAMADO SET chrestado = 'S', INTPERSONAFINALIZA = " & vllngPersonaGraba & ", CHRMEDICOENFERMERA = 'E' WHERE intnumcuenta = " & Trim(txtMovimientoPaciente.Text) & " AND chrtipoingreso = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "' AND chrestado = 'A'"
        End If
        
        'vllngCuentaCerrada  1 Cuenta cerrada 0 Cuenta abierta
        vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & "|" & _
                            IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & _
                            IIf(vllngCuentaCerrada = 1, "0", "1") & "|" & _
                            IIf(vllngCuentaCerrada = 0 And OptTipoPaciente(1).Value, fstrFechaSQL(fdtmServerFecha, fdtmServerHora), Null)
          
        frsEjecuta_SP vgstrParametrosSP, "SP_EXUPDCERRARABRIRCUENTA"
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, IIf(vllngCuentaCerrada = 1, "APERTURA DE CUENTA", "CIERRE DE CUENTA"), Trim(txtMovimientoPaciente.Text))
    
        SQL = "delete from pvTipoPacienteProceso where pvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
              "and pvTipoPacienteProceso.intproceso = " & enmTipoProceso.AbrirCerrarCuentas
        pEjecutaSentencia SQL
    
        SQL = "insert into pvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.AbrirCerrarCuentas & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ")"
        pEjecutaSentencia SQL
            
        If vllngCuentaCerrada = 1 Then
            pEjecutaSentencia "delete PVCuentasReabiertas where numNumCuenta = " & Trim(txtMovimientoPaciente.Text) & " and chrTipoPaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'"
            pEjecutaSentencia "insert into PVCuentasReabiertas (numNumCuenta, chrTipoPaciente, intCveEmpleado, intDiasReabrirE, intDiasReabierta) values (" & Trim(txtMovimientoPaciente.Text) & ", '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "', " & vllngPersonaGraba & ", " & IIf(OptTipoPaciente(0), vlintNumDiasAbrirInt, vlintNumDiasAbrirExt) & ", " & intDias & ")"
        End If
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        If Trim(vlmsgCargosProgramados) <> "" Then
            vlmsgCargosProgramados = "Se realizaron los siguientes cargos automáticos por cuidados especiales pendientes de aplicarse." & Chr(13) & vlmsgCargosProgramados
            
            'La información se actualizó satisfactoriaments mnas los cargos automaticos realizados
            MsgBox SIHOMsg(284) & Chr(13) & Chr(13) & vlmsgCargosProgramados, vbInformation, "Mensaje"
        Else
            'La información se actualizó satisfactoriamente.
            MsgBox SIHOMsg(284), vbInformation, "Mensaje"
        End If
        pCancelar
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEstatusAdm_Click"))
    Unload Me
End Sub

Private Sub cmdFacturada_Click()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, 4116, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 4116, "E", True) Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
                vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & IIf(lintCuentaFacturada, "0", "1")
                frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDCUENTAFACTURADA"
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "MARCAR COMO FACTURADA LA CUENTA DE PACIENTE", Trim(txtMovimientoPaciente.Text))
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La información se actualizó satisfactoriamente.
            MsgBox SIHOMsg(284), vbInformation, "Mensaje"
            pCancelar
        End If
    Else
        '¡El usuario no tiene permiso para grabar datos!.
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdFacturada_Click"))
    Unload Me
End Sub

Private Sub cmdFechasFactura_Click()
    Dim rs As New ADODB.Recordset
    Dim vlstrMensaje As String
    
    If OptTipoPaciente(0).Value Then
        vlstrMensaje = "¿Seguro que quiere modificar las fechas de ingreso o egreso?"
        Label10.Height = 195
        Label10.Top = 290
        Label10.Caption = "Fecha de ingreso"
        mskFechaIngreso.ToolTipText = "Fecha de ingreso"
        Label11.Height = 195
        Label11.Top = 710
        Label11.Caption = "Fecha de egreso"
        mskFechaEgreso.ToolTipText = "Fecha de egreso"
        Label12.Height = 195
        Label12.Top = 290
        Label12.Caption = "Hora de ingreso"
        mskHoraIngreso.ToolTipText = "Hora de ingreso"
        Label13.Height = 195
        Label13.Top = 710
        Label13.Caption = "Hora de egreso"
        mskHoraEgreso.ToolTipText = "Hora de egreso"
    Else
        vlstrMensaje = "¿Seguro que quiere modificar las fechas de inicio o fin de atención?"
        Label10.Height = 435
        Label10.Top = 230
        Label10.Caption = "Fecha de inicio de atención"
        mskFechaIngreso.ToolTipText = "Fecha de inicio de atención"
        Label11.Height = 435
        Label11.Top = 670
        Label11.Caption = "Fecha de fin de atención"
        mskFechaEgreso.ToolTipText = "Fecha de fin de atención"
        Label12.Height = 435
        Label12.Top = 230
        Label12.Caption = "Hora de inicio de atención"
        mskHoraIngreso.ToolTipText = "Hora de inicio de atención"
        Label13.Height = 435
        Label13.Top = 670
        Label13.Caption = "Hora de fin de atención"
        mskHoraEgreso.ToolTipText = "Hora de fin de atención"
    End If
    
    If MsgBox(vlstrMensaje, vbYesNo + vbQuestion, "Mensaje") = vbYes Then
        Me.Height = 6090
        mskFechaIngreso.SetFocus
        cmdFechasFactura.Enabled = False
        Frame1.Enabled = False
        cmdGrabaRegistro.Enabled = False
    End If
End Sub

Private Sub cmdGrabaRegistro_Click()
        Dim vlstrSentencia As String
        Dim vllngPersonaGraba As Long
        Dim rsFechasFactura As New ADODB.Recordset
        Dim vlbDatosCorrectos As Boolean
        
    On Error GoTo NotificaError
    vlbDatosCorrectos = True
    If Not IsDate(mskFechaIngreso.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        vlbDatosCorrectos = False
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaIngreso.SetFocus
    End If
    If Not IsDate(mskHoraIngreso.Text) And vlbDatosCorrectos Then
        '¡Hora no válida!, formato de hora hh:mm
        vlbDatosCorrectos = False
        MsgBox SIHOMsg(41), vbOKOnly + vbExclamation, "Mensaje"
        mskHoraIngreso.SetFocus
    End If
    If Not IsDate(mskFechaEgreso.Text) And vlbDatosCorrectos Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        vlbDatosCorrectos = False
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaEgreso.SetFocus
    End If
    If Not IsDate(mskHoraEgreso.Text) And vlbDatosCorrectos Then
        '¡Hora no válida!, formato de hora hh:mm
        vlbDatosCorrectos = False
        MsgBox SIHOMsg(41), vbOKOnly + vbExclamation, "Mensaje"
        mskHoraEgreso.SetFocus
    End If
    If vlbDatosCorrectos Then
        If CDate(mskFechaIngreso.Text) > CDate(mskFechaEgreso.Text) Then
            MsgBox "¡La fecha de ingreso debe de ser menor a la de egreso!.", vbOKOnly + vbExclamation, "Mensaje"
            mskFechaIngreso.SetFocus
            vlbDatosCorrectos = False
        End If
        If CDate(mskFechaEgreso.Text) = CDate(mskFechaIngreso.Text) And CDate(mskHoraIngreso.Text) >= CDate(mskHoraEgreso.Text) And vlbDatosCorrectos Then
            MsgBox "¡La hora de ingreso y egreso son iguales o la de ingreso es menor!.", vbOKOnly + vbExclamation, "Mensaje"
            mskHoraIngreso.SetFocus
            vlbDatosCorrectos = False
        End If
    End If
    
   If vlbDatosCorrectos = True Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & _
                                fstrFechaSQL(mskFechaIngreso.Text, mskHoraIngreso.Text) & "|" & _
                                fstrFechaSQL(mskFechaEgreso.Text, mskHoraEgreso.Text)
            frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDFECHAINGRESOEGRESO"

            Me.Height = 4155
            vlbDatosCorrectos = True
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La información se actualizó satisfactoriamente.
            MsgBox SIHOMsg(284), vbInformation, "Mensaje"
            pCancelar
        End If
    End If
    Frame1.Enabled = True
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdActualizar_Click"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
      
    Me.Icon = frmMenuPrincipal.Icon
      
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.AbrirCerrarCuentas) > 0 Then
        If fintEsInterno(vglngNumeroLogin, enmTipoProceso.AbrirCerrarCuentas) = 1 Then
            OptTipoPaciente(0).Value = True
        Else
            OptTipoPaciente(1).Value = True
        End If
    End If

    txtAfiliacion.Locked = True
    
    pCargaParametrosLocales
    
End Sub

Private Sub pCargaParametrosLocales()
    On Error GoTo NotificaError
        Dim rs As ADODB.Recordset
        Set rs = frsRegresaRs("select intDiasAbrirCuentasInternos, intDiasAbrirCuentasExternos from PVParametro where tnyclaveempresa = " & vgintClaveEmpresaContable, adLockReadOnly, adOpenForwardOnly)
        If Not rs.EOF Then
            vlintNumDiasAbrirExt = IIf(IsNull(rs!intDiasAbrirCuentasExternos), 0, rs!intDiasAbrirCuentasExternos)
            vlintNumDiasAbrirInt = IIf(IsNull(rs!intDiasAbrirCuentasInternos), 0, rs!intDiasAbrirCuentasInternos)
        Else
            vlintNumDiasAbrirExt = 0
            vlintNumDiasAbrirInt = 0
        End If
        
        Set rs = frsSelParametros("PV", vgintClaveEmpresaContable, "BITCERRARCUENTAREQUIPENDIENTES")
        If Not rs.EOF Then
            If IsNull(rs!valor) Then
                lblnPermitirCerrarCuenta = False
            Else
                lblnPermitirCerrarCuenta = IIf(rs!valor = "1", True, False)
            End If
        Else
            lblnPermitirCerrarCuenta = False
        End If
        rs.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaParametrosLocales"))
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    cmdFechasFactura.Enabled = False
    If vgstrEstadoManto = "C" Then
        Cancel = 1
        If MsgBox(SIHOMsg(9), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            vgstrEstadoManto = ""
            Me.Height = 4155
            pCancelar
        Else
            Me.Height = 6045
        End If
    End If
    Frame1.Enabled = True
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub mskFechaEgreso_GotFocus()
    On Error GoTo NotificaError
    pSelMkTexto mskFechaEgreso
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaIngreso_GotFocus"))
    Unload Me
End Sub

Private Sub mskFechaEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mskHoraEgreso.SetFocus
    End If
End Sub

Private Sub mskFechaIngreso_GotFocus()
    On Error GoTo NotificaError
    pSelMkTexto mskFechaIngreso
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaIngreso_GotFocus"))
    Unload Me
End Sub

Private Sub mskFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mskHoraIngreso.SetFocus
    End If
End Sub

Private Sub mskHoraEgreso_GotFocus()
    On Error GoTo NotificaError
    pSelMkTexto mskHoraEgreso
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaIngreso_GotFocus"))
    Unload Me
End Sub

Private Sub mskHoraEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdGrabaRegistro.Enabled Then
            cmdGrabaRegistro.SetFocus
        End If
    End If
End Sub

Private Sub mskHoraIngreso_GotFocus()
    On Error GoTo NotificaError
    pSelMkTexto mskHoraIngreso
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaIngreso_GotFocus"))
    Unload Me
End Sub

Private Sub mskHoraIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mskFechaEgreso.SetFocus
    End If
End Sub

Private Sub mskHoraIngreso_LostFocus()
    If mskHoraEgreso.Text = "" Then
        cmdGrabaRegistro.Enabled = False
    Else
        cmdGrabaRegistro.Enabled = True
    End If
End Sub

Private Sub txtAfiliacion_GotFocus()
    pSelTextBox txtAfiliacion
End Sub

Private Sub txtAfiliacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdActualizar.SetFocus
    End If
End Sub

Private Sub txtAfiliacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If KeyCode = vbKeyReturn Then
        If RTrim(txtMovimientoPaciente.Text) = "" Then
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
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                                        .vgstrTamanoCampo = "800,3400,1700,4100"
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
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                
                txtMovimientoPaciente.Text = .flngRegresaPaciente()
                
                If txtMovimientoPaciente <> -1 Then
                    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                Else
                    txtMovimientoPaciente.Text = ""
                End If
            End With
        Else
            
            vgstrParametrosSP = txtMovimientoPaciente.Text & _
            "|" & "0" & _
            "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & _
            "|" & vgintClaveEmpresaContable
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")
            
            If rs.RecordCount <> 0 Then
                If Not IsNull(rs!Ingreso) Then
                    vldtmFechaIngreso = rs!Ingreso
                    txtFechaApertura.Text = Format(rs!Ingreso, "dd/MM/yyyy HH:mm")
                End If
                If Not IsNull(rs!Egreso) Then
                    txtFechaCierre.Text = Format(rs!Egreso, "dd/MM/yyyy HH:mm")
                End If
                vgstrEstadoManto = "C" 'Cargando
                FrePaciente.Enabled = False
                txtPaciente.Text = rs!Nombre
                txtEmpresaPaciente.Text = IIf(IsNull(rs!empresa), "", rs!empresa)
                txtTipoPaciente.Text = rs!TipoPaciente
                txtCuarto = IIf(IsNull(rs!Cuarto), 0, rs!Cuarto)
                txtAfiliacion.Locked = False
                txtAfiliacion.Text = Trim(IIf(IsNull(rs!Afiliacion), " ", rs!Afiliacion))
                        
                vgstrEstatusAdm = rs!chrEstatusAdmision
                lintCuentaFacturada = rs!Facturada
                vllngCuentaCerrada = rs!CuentaCerrada
                                
                txtEstatus.Text = IIf(vllngCuentaCerrada = 1, "Cerrada", "Abierta") & "/" & IIf(lintCuentaFacturada = 1, "Facturada", "No facturada")
                
                If lintCuentaFacturada = 1 And Trim(vgstrEstatusAdm) <> "A" Then
                    cmdFacturada.Caption = "Poner como no facturada"
                    cmdFacturada.Enabled = True
                Else
                    If lintCuentaFacturada = 0 And Trim(vgstrEstatusAdm) <> "A" Then
                        cmdFacturada.Caption = "Poner como facturada"
                        cmdFacturada.Enabled = True
                    Else
                        cmdFacturada.Enabled = False
                    End If
                End If
                
                cmdEstatusAdm.Enabled = True
                If vllngCuentaCerrada = 1 Then
                    cmdEstatusAdm.Caption = "Abrir cuenta"
                Else
                    cmdEstatusAdm.Caption = "Cerrar cuenta"
                End If
                cmdActualizar.Enabled = True
            Else
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                pCancelar
            End If
        End If
    End If

If txtEstatus.Text = "Abierta/No facturada" Then
    cmdFechasFactura.Enabled = True
End If
If txtEstatus.Text = "Cerrada/No facturada" Then
    cmdFechasFactura.Enabled = True
End If
If txtEstatus.Text = "cerrada/facturada" Then
    cmdFechasFactura.Enabled = False
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub pCancelar()
    On Error GoTo NotificaError
    
    FrePaciente.Enabled = True
    txtPaciente.Text = ""
    txtTipoPaciente.Text = ""
    txtEmpresaPaciente.Text = ""
    txtCuarto.Text = ""
    txtAfiliacion.Text = ""
    txtAfiliacion.Locked = True
    txtEstatus.Text = ""
    txtFechaApertura.Text = ""
    txtFechaCierre.Text = ""
    cmdEstatusAdm.Enabled = False
    cmdFacturada.Enabled = False
    cmdActualizar.Enabled = False
    pEnfocaTextBox txtMovimientoPaciente
    vgstrEstadoManto = ""
    mskFechaIngreso.Mask = ""
    mskFechaIngreso.Text = ""
    mskFechaIngreso.Mask = "##/##/####"
    mskFechaEgreso.Mask = ""
    mskFechaEgreso.Text = ""
    mskFechaEgreso.Mask = "##/##/####"
    mskHoraIngreso.Mask = ""
    mskHoraIngreso.Text = ""
    mskHoraIngreso.Mask = "##:##"
    mskHoraEgreso.Mask = ""
    mskHoraEgreso.Text = ""
    mskHoraEgreso.Mask = "##:##"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pcancelar"))
    Unload Me
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError
    If OptTipoPaciente(0).Value Then
        Label7.Top = 2490
        Label7.Height = 195
        Label7.Caption = "Fecha de ingreso"
        Label9.Top = 2490
        Label9.Height = 195
        Label9.Caption = "Fecha de egreso"
        cmdFechasFactura.Caption = "Ingreso/Egreso"
        cmdFechasFactura.ToolTipText = "Modificar fechas de ingreso o egreso"
    Else
        Label7.Top = 2400
        Label7.Height = 435
        Label7.Caption = "Fecha de inicio de atención"
        Label9.Top = 2400
        Label9.Height = 435
        Label9.Caption = "Fecha de fin de atención"
        cmdFechasFactura.Caption = "Inicio/Fin de atención"
        cmdFechasFactura.ToolTipText = "Modificar fechas de inicio o fin de atención"
    End If
    pEnfocaTextBox txtMovimientoPaciente
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":OptTipoPaciente_Click"))
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_KeyPress"))
    Unload Me
End Sub

