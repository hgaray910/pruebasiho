VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExclusionCargos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exclusión de cargos"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freTrabajando 
      Height          =   1335
      Left            =   3255
      TabIndex        =   26
      Top             =   9405
      Visible         =   0   'False
      Width           =   4560
      Begin MSComCtl2.Animation anmTrabajando2 
         Height          =   765
         Left            =   210
         TabIndex        =   27
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1349
         _Version        =   393216
         FullWidth       =   61
         FullHeight      =   51
      End
      Begin VB.Label Label17 
         Caption         =   "Excluyendo cargos, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1305
         TabIndex        =   28
         Top             =   345
         Width           =   3090
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1290
      Left            =   2250
      TabIndex        =   23
      Top             =   7875
      Visible         =   0   'False
      Width           =   7680
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   360
         Left            =   1035
         TabIndex        =   24
         Top             =   675
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         Caption         =   "Consultando cargos del paciente, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   180
         Width           =   7410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   30
         Top             =   120
         Width           =   7620
      End
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Orden"
      Height          =   1400
      Left            =   75
      TabIndex        =   18
      Top             =   6135
      Width           =   2130
      Begin VB.OptionButton optOrdenCargos 
         Caption         =   "Descripción"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Mostrar cargos ordenados por descripción de cargo"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optOrdenCargos 
         Caption         =   "Departamento"
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Mostrar cargos ordenados por departamento"
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optOrdenCargos 
         Caption         =   "Concepto facturación"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Mostrar cargos ordenados por concepto de facturación"
         Top             =   435
         Width           =   1935
      End
      Begin VB.OptionButton optOrdenCargos 
         Caption         =   "Fecha"
         Height          =   270
         Index           =   0
         Left            =   105
         TabIndex        =   19
         ToolTipText     =   "Mostrar cargos ordenados por fecha"
         Top             =   195
         Value           =   -1  'True
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   75
      TabIndex        =   3
      Top             =   0
      Width           =   7530
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   0
         ToolTipText     =   "Número de cuenta del paciente"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPaciente 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Nombre del paciente"
         Top             =   585
         Width           =   5700
      End
      Begin VB.TextBox txtEmpresaPaciente 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Nombre de la empresa del paciente"
         Top             =   915
         Width           =   5700
      End
      Begin VB.TextBox txtTipoPaciente 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Tipo de paciente"
         Top             =   1260
         Width           =   4035
      End
      Begin VB.TextBox txtFechaInicial 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Fecha de inicio de atención"
         Top             =   1605
         Width           =   1890
      End
      Begin VB.TextBox txtFechaFinal 
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Fecha final de atención"
         Top             =   1605
         Width           =   1890
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   3135
         TabIndex        =   4
         Top             =   180
         Width           =   1950
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Externo"
            Height          =   195
            Index           =   1
            Left            =   990
            TabIndex        =   6
            Top             =   150
            Width           =   855
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Interno"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   135
            Width           =   825
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de cuenta"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   645
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   975
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de atención"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   1665
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   3570
         TabIndex        =   12
         Top             =   1680
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   5603
      TabIndex        =   2
      Top             =   6570
      Width           =   615
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000000&
         Picture         =   "frmExclusionCargos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Ultimo pago"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCargos 
      Height          =   3990
      Left            =   75
      TabIndex        =   1
      Top             =   2100
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   7038
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      GridColor       =   12632256
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Height          =   195
      Left            =   7500
      TabIndex        =   34
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Excluido"
      Height          =   195
      Left            =   7755
      TabIndex        =   33
      Top             =   6120
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Precio modificado"
      Height          =   195
      Left            =   10410
      TabIndex        =   32
      Top             =   6135
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   10110
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   7440
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Height          =   195
      Left            =   7515
      TabIndex        =   31
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080C0FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   8460
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fecha modificada"
      Height          =   195
      Left            =   8760
      TabIndex        =   30
      Top             =   6120
      Width           =   1260
   End
End
Attribute VB_Name = "frmExclusionCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
' Programa marcar un cargo como excluido o viceversa a pacientes internos o externos,
' solo cargos que no han sido facturados.
' Fecha de programación: Miércoles 17 de Enero de 2001
'------------------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
'------------------------------------------------------------------------------------------

Dim vlstrSentencia As String
Dim vlblnLimpiar As Boolean
Dim vgstrEstadoManto As String
Dim vgintEmpresa As Integer         'Clave de la empresa
Dim vgintTipoPaciente As Integer    'Clave del Tipo de Paciente
Dim vlblnEntrando As Boolean
Dim rsPvSelDatosPaciente As New ADODB.Recordset

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    Dim vllngPersonaGraba As Long
    Dim vlstrBloqueo As String
    Dim SQL As String
    Dim X As Long
    Dim lblnExcluido As Boolean
    Dim blnUPDCostoDescuento As Boolean
    
    '-------------------------------------------------------------------
    '   Valida si la cuenta se encuentra bloqueada por trabajo social
    '-------------------------------------------------------------------
    If fblnCuentaBloqueada(Trim(txtMovimientoPaciente.Text), IIf(optTipoPaciente(0).Value, "I", "E")) Then
        'No se puede realizar ésta operación. La cuenta se encuentra bloqueada por trabajo social.
        MsgBox SIHOMsg(662), vbCritical, "Mensaje"
        Exit Sub
    End If
    
    '-------------------------------------
    ' Persona que graba
    '-------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        '--------------------------------
        ' Pongo Letrerito
        '--------------------------------
        freTrabajando.Top = 2500
        freTrabajando.Visible = True
        frmExclusionCargos.Refresh
        lblnExcluido = False
        EntornoSIHO.ConeccionSIHO.BeginTrans
        vlstrBloqueo = fstrBloqueaCuenta(Val(txtMovimientoPaciente.Text), IIf(optTipoPaciente(0).Value, "I", "E"))
        
        blnUPDCostoDescuento = fblnUPDCostoDescuento
        
        If vlstrBloqueo = "L" Then
            For X = 1 To grdCargos.Rows - 1
                If Trim(grdCargos.TextMatrix(X, grdCargos.Cols - 1)) = "*" Then
                    vlstrSentencia = "Update PvCargo set bitExcluido= "
                    vlstrSentencia = vlstrSentencia + IIf(Trim(grdCargos.TextMatrix(X, 9)) = "", "0", "1")
                    vlstrSentencia = vlstrSentencia + ",intNumPaquete = " + IIf(Trim(grdCargos.TextMatrix(X, 9)) = "", "intNumPaquete", "0")
                    vlstrSentencia = vlstrSentencia + ",intCantidadPaquete = " + IIf(Trim(grdCargos.TextMatrix(X, 9)) = "", "intCantidadPaquete", "0")
                    vlstrSentencia = vlstrSentencia + ",intCantidadExtraPaquete = " + IIf(Trim(grdCargos.TextMatrix(X, 9)) = "", "intCantidadExtraPaquete", "0")
                    vlstrSentencia = vlstrSentencia + " where intNumCargo=" + Str(grdCargos.RowData(X))
                    pEjecutaSentencia vlstrSentencia
                    
                    If blnUPDCostoDescuento Then
                        vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) _
                                            & "|" & IIf(optTipoPaciente(0).Value, "'I'", "'E'") _
                                            & "|" & Trim(txtMovimientoPaciente.Text) _
                                            & "|" & IIf(optTipoPaciente(0).Value, "'I'", "'E'") _
                                            & "|" & Trim(Str(vgintTipoPaciente)) _
                                            & "|" & Trim(Str(vgintEmpresa)) _
                                            & "|" & Trim(Str(grdCargos.RowData(X))) _
                                            & "|1" _
                                            & "|0" _
                                            & "|0" _
                                            & "|0" _
                                            & "|0" _
                                            & "|0" _
                                            & "|" & Str(vllngPersonaGraba) _
                                            & "|" & Str(vgintNumeroDepartamento) _
                                            & "|0"
                                            
                        frsEjecuta_SP vgstrParametrosSP, "SP_PvUpdTrasladoCargos", True
                    End If
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "EXCLUSION DE CARGOS", CStr(grdCargos.RowData(X)))
                End If
            Next X
            
            SQL = "delete from pvTipoPacienteProceso where pvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
                "and pvTipoPacienteProceso.intproceso = " & enmTipoProceso.Exclusion
            pEjecutaSentencia SQL
            
            SQL = "insert into pvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.Exclusion & "," & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & ")"
            pEjecutaSentencia SQL
            pLiberaCuenta
            lblnExcluido = True
        Else
            If vlstrBloqueo = "F" Then
                'La cuenta ya ha sido facturada
                MsgBox SIHOMsg(299), vbOKOnly + vbInformation, "Mensaje"
            Else
                If vlstrBloqueo = "O" Then
                    'La cuenta esta siendo usada por otra persona, intente de nuevo.
                    MsgBox SIHOMsg(300), vbOKOnly + vbInformation, "Mensaje"
                End If
            End If
        End If
        '--------------------------------
        ' Quito Letrerito
        '--------------------------------
        freTrabajando.Visible = False
        frmExclusionCargos.Refresh
        If lblnExcluido Then
            EntornoSIHO.ConeccionSIHO.CommitTrans
        Else
            EntornoSIHO.ConeccionSIHO.RollbackTrans
        End If
        If vlstrBloqueo = "L" Then
            pLlenaCargos
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub
Private Function fblnUPDCostoDescuento() As Boolean
'funcion que busca valor del parametro BITCONSERVARCOSTOSDESCUENTOEXCLUSION para saber si se debe de actualizar los costos y decuentos de los cargos al momento de'
'hacer una exclusión de cargos

Dim ObjRs As New ADODB.Recordset
Dim ObjStr As String

fblnUPDCostoDescuento = True

ObjStr = "select vchvalor from siparametro where vchnombre = 'BITCONSERVARCOSTOSDESCUENTOEXCLUSION' and INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
Set ObjRs = frsRegresaRs(ObjStr, adLockOptimistic)

If ObjRs.RecordCount > 0 Then
   If ObjRs!vchvalor = "1" Then
      fblnUPDCostoDescuento = False
   End If
End If

End Function
Private Sub optOrdenCargos_Click(Index As Integer)
    
    If Not vlblnEntrando Then
        pLlenaCargos
    End If
    
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If RTrim(txtMovimientoPaciente.Text) = "" Then
            With FrmBusquedaPacientes
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSinFacturar.Value = True
                .optSinFacturar.Enabled = False
                .optSoloActivos.Enabled = False
                .optTodos.Enabled = False
                
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
            If optTipoPaciente(0).Value Then 'Internos
                Set rsPvSelDatosPaciente = frsEjecuta_SP(Val(txtMovimientoPaciente.Text) & "|0|" & "I" & "|" & vgintClaveEmpresaContable, "SP_PvSelDatosPaciente", False)
            Else
                Set rsPvSelDatosPaciente = frsEjecuta_SP(Val(txtMovimientoPaciente.Text) & "|0|" & "E" & "|" & vgintClaveEmpresaContable, "SP_PvSelDatosPaciente", False)
            End If
            If rsPvSelDatosPaciente.RecordCount <> 0 Then
                If rsPvSelDatosPaciente!bitUtilizaConvenio = 1 Then
                    If rsPvSelDatosPaciente!Facturada = 0 Or rsPvSelDatosPaciente!Facturada = False Then
                        txtPaciente.Text = rsPvSelDatosPaciente!Nombre
                        txtEmpresaPaciente.Text = IIf(IsNull(rsPvSelDatosPaciente!empresa), "", rsPvSelDatosPaciente!empresa)
                        txtTipoPaciente.Text = rsPvSelDatosPaciente!tipo
                        txtFechaInicial = rsPvSelDatosPaciente!Ingreso
                        txtFechaFinal = IIf(IsNull(rsPvSelDatosPaciente!Egreso), "", rsPvSelDatosPaciente!Egreso)
                        vgintTipoPaciente = rsPvSelDatosPaciente!tnyCveTipoPaciente
                        vgintEmpresa = rsPvSelDatosPaciente!intcveempresa
                        
                        pLlenaCargos
                        
                        grdCargos.SetFocus
                        cmdSave.Enabled = True
                        vgstrEstadoManto = "C"
                    Else
                        'La cuenta del paciente está completamente facturada.
                        MsgBox SIHOMsg(597), vbOKOnly + vbInformation, "Mensaje"
                        txtMovimientoPaciente.Text = ""
                    End If
                Else
                    'El paciente no es de tipo "Convenio".
                    MsgBox SIHOMsg(351), vbOKOnly + vbInformation, "Mensaje"
                    txtMovimientoPaciente.Text = ""
                End If
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                txtMovimientoPaciente.Text = ""
            End If
            rsPvSelDatosPaciente.Close
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
End Sub

Public Sub pLlenaCargos()
    Dim vlintcontador As Integer
    Dim rsSeleccionaCargos As New ADODB.Recordset
    
    '-------------------------------------------------------------------
    ' Este Procedure lo usan las pantallas de Cargos directos de Caja y de
    '-------------------------------------------------------------------
    grdCargos.Redraw = False
    grdCargos.Rows = 2
    
    pLimpiaGrid grdCargos

    '------------------------
    ' Progress Bar : "Cargando Datos..."
    '------------------------
    freBarra.Top = 3000
    freBarra.MousePointer = ssHourglass
    freBarra.Visible = True
    freBarra.Refresh
    pgbBarra.Value = 10
    
    
    pConfiguraGridCargos grdCargos
    
    vgstrParametrosSP = txtMovimientoPaciente & "|" & IIf(optTipoPaciente(0), "I", "E") & "|" & 0 & "|" & "-1|C|N|0"
    Set rsSeleccionaCargos = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCARGOSPACIENTE", False)
    
    With rsSeleccionaCargos
    
        If .RecordCount <> 0 Then
        
            fraOrden.Enabled = True
            
            If optOrdenCargos(1).Value Then
                .Sort = "Concepto"
            ElseIf optOrdenCargos(2).Value Then
                    .Sort = "NombreDepartamento"
                ElseIf optOrdenCargos(3).Value Then
                    .Sort = "DescripcionCargo"
                    ElseIf optOrdenCargos(0).Value Then
                        .Sort = "dtmFechaHora"
            End If
            
            Do While Not .EOF
                If grdCargos.RowData(1) <> -1 Then
                     grdCargos.Rows = grdCargos.Rows + 1
                     grdCargos.Row = grdCargos.Rows - 1
                End If
                pgbBarra.Value = (.Bookmark / .RecordCount) * 100
                grdCargos.RowData(grdCargos.Row) = !IntNumCargo
                grdCargos.TextMatrix(grdCargos.Row, 0) = IIf(Trim(!FolioFactura) = "", "", "F")
                grdCargos.TextMatrix(grdCargos.Row, 1) = !IntNumCargo
                grdCargos.TextMatrix(grdCargos.Row, 2) = Format(CStr(!dtmFechahora), "dd/mmm/yyyy HH:mm")
                grdCargos.TextMatrix(grdCargos.Row, 3) = !TipoDocumento
                grdCargos.TextMatrix(grdCargos.Row, 4) = !intFolioDocumento
                grdCargos.TextMatrix(grdCargos.Row, 5) = !chrTipoCargo
                grdCargos.TextMatrix(grdCargos.Row, 6) = IIf(IsNull(!DescripcionCargo), "", Trim(!DescripcionCargo))
                grdCargos.TextMatrix(grdCargos.Row, 7) = !MNYCantidad
                grdCargos.TextMatrix(grdCargos.Row, 8) = FormatCurrency(!mnyPrecio, 2)
                grdCargos.TextMatrix(grdCargos.Row, 9) = IIf(IsNull(!Excluido), "", !Excluido)
                grdCargos.TextMatrix(grdCargos.Row, 10) = Trim(!Concepto)
                grdCargos.TextMatrix(grdCargos.Row, 11) = IIf(IsNull(!HojaConsumo), "", !HojaConsumo)
                grdCargos.TextMatrix(grdCargos.Row, 12) = !NumeroCirugia
                grdCargos.TextMatrix(grdCargos.Row, 13) = Trim(!NombreDepartamento)
                grdCargos.TextMatrix(grdCargos.Row, 14) = Trim(!NombreEmpleado)
                
                        
            If !PrecioManual Then
                grdCargos.Col = 8
                grdCargos.CellBackColor = &HC0FFFF   'fondo amarillo
            End If
            If Not IsNull(!fechamanual) Then
                If !fechamanual = 1 Then
                   grdCargos.Col = 2
                   grdCargos.CellBackColor = &HC0E0FF 'fondo naranja
                End If
            End If
            If !Excluido = "X" Then
                For vlintcontador = 1 To grdCargos.Cols - 1
                    grdCargos.Col = vlintcontador
                    grdCargos.CellForeColor = &HFF0000 'letra azul
                Next
            End If

            .MoveNext
            Loop
        End If
        
        .Close
    End With
    freBarra.Visible = False
    grdCargos.Redraw = True

End Sub

Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    Dim vlbytColumnas As Byte
    
    With ObjGrd
        .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
        
        .Cols = 15
        
        For vlbytColumnas = 1 To .Cols - 1
            .TextMatrix(1, vlbytColumnas) = ""
        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
    End With
End Sub

Private Sub pLiberaCuenta()
    On Error GoTo NotificaError
    
    frsEjecuta_SP IIf(optTipoPaciente(0).Value, "I", "E") & "|" & txtMovimientoPaciente.Text & "|0", "SP_EXUPDCUENTAOCUPADA"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLiberaCuenta"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    vgstrEstadoManto = ""
    vlblnEntrando = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlblnLimpiar = True
    vlblnEntrando = True
    
    optTipoPaciente(0).Value = True
    
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Exclusion) > 0 Then
      If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Exclusion) = 1 Then
        optTipoPaciente(0).Value = True
      Else
        optTipoPaciente(1).Value = True
      End If
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If vgstrEstadoManto = "C" Then
        Cancel = 1
        pEnfocaTextBox txtMovimientoPaciente
    End If
End Sub

Private Sub grdCargos_Click()
    On Error GoTo NotificaError
    
    If grdCargos.RowData(1) <> 0 Then
        If grdCargos.Col = 9 Then
            If Trim(grdCargos.TextMatrix(grdCargos.Row, 9)) = "" Then
                grdCargos.TextMatrix(grdCargos.Row, 9) = "X"
            Else
                grdCargos.TextMatrix(grdCargos.Row, 9) = ""
            End If
            grdCargos.TextMatrix(grdCargos.Row, grdCargos.Cols - 1) = "*"
        End If
    End If
    grdCargos.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_Click"))
End Sub

Private Sub grdCargos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If grdCargos.RowData(1) <> 0 Then
            If grdCargos.Col = 9 Then
                If Trim(grdCargos.TextMatrix(grdCargos.Row, 9)) = "" Then
                    grdCargos.TextMatrix(grdCargos.Row, 9) = "X"
                Else
                    grdCargos.TextMatrix(grdCargos.Row, 9) = ""
                End If
                grdCargos.TextMatrix(grdCargos.Row, grdCargos.Cols - 1) = "*"
            End If
        End If
        grdCargos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_KeyPress"))
End Sub

Private Sub optTipoPaciente_GotFocus(Index As Integer)
    On Error GoTo NotificaError
    
     'txtMovimientoPaciente.SetFocus
     pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_GotFocus"))
End Sub

Private Sub optTipoPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtMovimientoPaciente.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_KeyPress"))
End Sub

Private Sub txtMovimientoPaciente_GotFocus()
    On Error GoTo NotificaError
    
    If vlblnLimpiar Then
        pLimpia
    Else
        vlblnLimpiar = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_GotFocus"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    txtMovimientoPaciente.Text = ""
    txtPaciente.Text = ""
    txtEmpresaPaciente.Text = ""
    txtTipoPaciente.Text = ""
    txtFechaFinal.Text = ""
    txtFechaInicial.Text = ""
    
    grdCargos.Rows = 2
    grdCargos.Clear
    grdCargos.RowData(1) = 0
    pConfiguraGridCargos grdCargos

    cmdSave.Enabled = False
    vgstrEstadoManto = ""
    vgintEmpresa = 0
    vgintTipoPaciente = 0
    
    fraOrden.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub
Private Sub pFormatoPesos()
    On Error GoTo NotificaError
    
    Dim X As Long
    
    For X = 1 To grdCargos.Rows - 1
        grdCargos.TextMatrix(X, 8) = FormatCurrency(Val(grdCargos.TextMatrix(X, 8)), 2)
    Next X

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFormatoPesos"))
End Sub

Private Sub pConfiguraGridCargos(grdNombre As MSHFlexGrid)
    On Error GoTo NotificaError
    
    With grdNombre
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Número cargo|Fecha/Hora|Tipo|Folio|Tipo|Descripción|Cantidad|Precio|Excluido|Concepto|Hoja consumo|Número cirugía|Departamento realizó cargo|Empleado realizó cargo"
        .ColWidth(0) = 100
        .ColWidth(1) = 0        'Número cargo
        .ColWidth(2) = 1500     'Fecha
        .ColWidth(3) = 1100      'Tipo documento
        .ColWidth(4) = 1000     'Folio
        .ColWidth(5) = 600      'Tipo cargo
        .ColWidth(6) = 3500     'Descripcion
        .ColWidth(7) = 750      'Cantidad
        .ColWidth(8) = 1000     'Precio
        .ColWidth(9) = 700      'Excluido
        .ColWidth(10) = 2500    'Concepto
        .ColWidth(11) = 1500    'Hoja consumo
        .ColWidth(12) = 1200    'Numero cirugia
        .ColWidth(13) = 3000    'Departamento que realizó el cargo
        .ColWidth(14) = 3000    'Empleado que realizó el cargo
        .Cols = .Cols + 1
        .ColWidth(.Cols - 1) = 0 'Estatus para saber cual registro se va a guardar
        
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignLeftCenter
        .ColAlignment(11) = flexAlignLeftCenter
        .ColAlignment(12) = flexAlignRightCenter
        .ColAlignment(13) = flexAlignLeftCenter
        .ColAlignment(14) = flexAlignLeftCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignmentFixed(11) = flexAlignCenterCenter
        .ColAlignmentFixed(12) = flexAlignCenterCenter
        .ColAlignmentFixed(13) = flexAlignCenterCenter
        .ColAlignmentFixed(14) = flexAlignCenterCenter
        
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
End Sub

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            optTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
            optTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyPress"))
End Sub

