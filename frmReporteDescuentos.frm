VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReporteDescuentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos asignados"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   90
      TabIndex        =   27
      Top             =   0
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   220
         Width           =   6015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   280
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vigencias"
      Height          =   1185
      Left            =   90
      TabIndex        =   26
      Top             =   4200
      Width           =   7125
      Begin VB.OptionButton optVigencia 
         Caption         =   "Todos "
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Mostrar todos con o sin vigencia"
         Top             =   285
         Width           =   1725
      End
      Begin MSMask.MaskEdBox mskInicioVigencia 
         Height          =   315
         Left            =   2800
         TabIndex        =   18
         ToolTipText     =   "Inicio de vigencia"
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton optVigencia 
         Caption         =   "Vigencia específica"
         Height          =   210
         Index           =   2
         Left            =   2800
         TabIndex        =   17
         ToolTipText     =   "Mostrar los que cumplan con la vigencia descrita"
         Top             =   285
         Width           =   1725
      End
      Begin VB.OptionButton optVigencia 
         Caption         =   "Todos con vigencia"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Mostrar todos los que tienen alguna vigencia"
         Top             =   810
         Width           =   1725
      End
      Begin VB.OptionButton optVigencia 
         Caption         =   "Todos sin vigencia"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Mostrar únicamente los que no tienen vigencia"
         Top             =   540
         Width           =   1725
      End
      Begin MSMask.MaskEdBox mskFinVigencia 
         Height          =   315
         Left            =   4080
         TabIndex        =   19
         ToolTipText     =   "Inicio de vigencia"
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3075
      TabIndex        =   25
      Top             =   5400
      Width           =   1140
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteDescuentos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteDescuentos.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipos de pacientes"
      Height          =   1230
      Left            =   90
      TabIndex        =   24
      Top             =   2940
      Width           =   2490
      Begin VB.OptionButton optPaciente 
         Caption         =   "Urgencias"
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   29
         ToolTipText     =   "Pacientes externos"
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton optPaciente 
         Caption         =   "Externos"
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         ToolTipText     =   "Pacientes externos"
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton optPaciente 
         Caption         =   "Internos"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Pacientes internos"
         Top             =   720
         Width           =   945
      End
      Begin VB.OptionButton optPaciente 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Todos los pacientes"
         Top             =   360
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipos de descuento asignado"
      Height          =   1230
      Left            =   2655
      TabIndex        =   23
      Top             =   2940
      Width           =   4560
      Begin VB.OptionButton optDescuento 
         Caption         =   "Otros conceptos"
         Height          =   195
         Index           =   5
         Left            =   2730
         TabIndex        =   13
         ToolTipText     =   "Otros conceptos"
         Top             =   840
         Width           =   1560
      End
      Begin VB.OptionButton optDescuento 
         Caption         =   "Exámenes"
         Height          =   195
         Index           =   4
         Left            =   2730
         TabIndex        =   11
         ToolTipText     =   "Exámenes"
         Top             =   360
         Width           =   1440
      End
      Begin VB.OptionButton optDescuento 
         Caption         =   "Estudios"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Estudios"
         Top             =   600
         Width           =   1260
      End
      Begin VB.OptionButton optDescuento 
         Caption         =   "Artículos"
         Height          =   195
         Index           =   2
         Left            =   2730
         TabIndex        =   12
         ToolTipText     =   "Artículos"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optDescuento 
         Caption         =   "Conceptos de facturación"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Conceptos de facturación"
         Top             =   840
         Width           =   2235
      End
      Begin VB.OptionButton optDescuento 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Todos los tipos"
         Top             =   360
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame FreConDescuento 
      Height          =   2235
      Left            =   90
      TabIndex        =   22
      Top             =   660
      Width           =   7125
      Begin VB.OptionButton optTipo 
         Caption         =   "Pacientes"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Pacientes con descuentos"
         Top             =   240
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Tipo de paciente"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Tipos de pacientes con descuento"
         Top             =   240
         Width           =   1635
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Empresas"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   3
         ToolTipText     =   "Empresas con descuento"
         Top             =   240
         Width           =   1080
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTipoAsignacion 
         Height          =   1590
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   2805
         _Version        =   393216
         GridColor       =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmReporteDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmReporteDescuentos
'-------------------------------------------------------------------------------------
'| Objetivo: Reportear descuentos
'-------------------------------------------------------------------------------------

Option Explicit
Private vgrptReporte As CRAXDRT.Report
Dim rs As New ADODB.Recordset
Dim vlstrSentencia As String

Private Sub pCargaGuardados()
    On Error GoTo NotificaError
    
    With grdTipoAsignacion
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 2
        .Cols = 4
        .TextMatrix(1, 1) = "-1"
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = "<TODOS>"
    End With
       
    vgstrParametrosSP = Str(cboHospital.ItemData(cboHospital.ListIndex)) & "|" & IIf(optTipo(0).Value, "P", IIf(optTipo(1).Value, "T", "E"))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDESCUENTOSASIGNADOS")
    
    Do While Not rs.EOF
        With grdTipoAsignacion
            If Trim(.TextMatrix(1, 1)) = "" Then
                .Row = 1
            Else
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            .TextMatrix(.Row, 1) = IIf(IsNull(rs!Clave), "", rs!Clave)
            .TextMatrix(.Row, 2) = IIf(rs!TipoPaciente = "A", "T", rs!TipoPaciente)
            .TextMatrix(.Row, 3) = IIf(IsNull(rs!Nombre), "", rs!Nombre)
        End With
        rs.MoveNext
    Loop
    
    With grdTipoAsignacion
        .ColWidth(0) = 100
        .ColWidth(1) = 0        'clave
        .ColWidth(2) = 300      'tipo de paciente
        .ColWidth(3) = 4700     'descripcion del paciente, empresa, tipo de paciente
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": pCargaGuardados"))
    Unload Me
End Sub

Private Sub cboHospital_Click()
    On Error GoTo NotificaError
    Dim intIndex As Integer
    
    intIndex = IIf(optTipo(0).Value, 0, IIf(optTipo(1).Value, 1, 2))
    optTipo_Click intIndex

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        If KeyCode = vbKeyReturn Then optTipo(0).SetFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_KeyDown"))
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo NotificaError
        pImprime "P"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": cmdPreview_Click"))
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo NotificaError
        pImprime "I"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": cmdPrint_Click"))
    Unload Me
End Sub

Private Sub pImprime(vlstrDestino As String)
    On Error GoTo NotificaError
    Dim rsReporte As ADODB.Recordset
    Dim vlstrTipoDescuento As String
    Dim vldblClave As Double
    Dim vlStrTipoPaciente As String
    Dim vlstrTipoCargo As String
    Dim vldblVigencia As Double
    Dim vlstrIniVigencia As String
    Dim vlstrFinVigencia As String
    Dim alstrParametros(0) As String
    Dim vlblnContinuar As Boolean
    
    vlblnContinuar = True
    
    vlstrTipoDescuento = IIf(optTipo(0).Value, "P", IIf(optTipo(1).Value, "T", "E"))
    vldblClave = Val(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 1))
    vlStrTipoPaciente = IIf(Val(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 1)) = -1, IIf(optPaciente(0).Value, "*", IIf(optPaciente(1).Value, "I", IIf(optPaciente(2).Value, "E", "U"))), IIf(Trim(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 2)) = "T", "A", Trim(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 2))))
    vlstrTipoCargo = IIf(optDescuento(0).Value, "*", IIf(optDescuento(1).Value, "CF", IIf(optDescuento(2).Value, "AR", IIf(optDescuento(3).Value, "ES", IIf(optDescuento(4).Value, "EX", IIf(optDescuento(5).Value, "OC", "GE"))))))
    vldblVigencia = IIf(optVigencia(3).Value, -1, IIf(optVigencia(0).Value, 0, IIf(optVigencia(1).Value, 1, 2)))
    If optVigencia(2).Value Then
        'Validar fechas
        If Not IsDate(mskInicioVigencia.Text) Then
            vlblnContinuar = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            mskInicioVigencia.SetFocus
        Else
            If Not IsDate(mskFinVigencia.Text) Then
                vlblnContinuar = False
                '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                mskFinVigencia.SetFocus
            Else
                vlstrIniVigencia = mskInicioVigencia.Text
                vlstrFinVigencia = mskFinVigencia.Text
            End If
        End If
    Else
        vlstrIniVigencia = "01/01/2004"
        vlstrFinVigencia = "01/01/2004"
    End If
        
    If vlblnContinuar Then
    Set rsReporte = frsEjecuta_SP(vlstrTipoDescuento & "|" & vldblClave & "|" & vlStrTipoPaciente & "|" & vlstrTipoCargo & "|" & vldblVigencia & "|" & Format(vlstrIniVigencia, "dd/mm/yyyy") & "|" & Format(vlstrFinVigencia, "dd/mm/yyyy") & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex)), "SP_PVRPTDESCUENTOS")
    If rsReporte.RecordCount = 0 Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital; " & Trim(cboHospital.List(cboHospital.ListIndex))
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Reporte de descuentos asignados"
    End If
    rsReporte.Close
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": pImprime"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = vbKeyEscape Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim lngNumOpcion As Long
   
    Me.Icon = frmMenuPrincipal.Icon
   
    Select Case cgstrModulo
    Case "PV"
        lngNumOpcion = 371
    Case "SE"
        lngNumOpcion = 1538
    End Select
   
    pCargaHospital lngNumOpcion
   
    pInstanciaReporte vgrptReporte, "rptDescuentosAsignados.rpt"
    optTipo_Click 0
   
    optPaciente(0).Value = True
    optDescuento(0).Value = True
    optVigencia(3).Value = True
    optVigencia_Click 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": Form_Load"))
    Unload Me
End Sub

Private Sub grdTipoAsignacion_Click()
    On Error GoTo NotificaError
        pHabilitaTipoPaciente
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": grdTipoAsignacion_Click"))
    Unload Me
End Sub

Private Sub grdTipoAsignacion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If optPaciente(0).Enabled Then
            If optPaciente(0).Value Then
                optPaciente(0).SetFocus
            Else
                If optPaciente(1).Value Then
                    optPaciente(1).SetFocus
                Else
                    If optPaciente(2).Value Then
                        optPaciente(2).SetFocus
                    Else
                        optPaciente(3).SetFocus
                    End If
                End If
            End If
        Else
            pEnfocaOpt optPaciente
        End If
    Else
        pHabilitaTipoPaciente
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": grdTipoAsignacion_KeyDown"))
    Unload Me
End Sub

Private Sub mskFinVigencia_GotFocus()
    On Error GoTo NotificaError
        pSelMkTexto mskFinVigencia
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": mskFinVigencia_GotFocus"))
    Unload Me
End Sub

Private Sub mskFinVigencia_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        If KeyCode = vbKeyReturn Then cmdPreview.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": mskFinVigencia_KeyDown"))
    Unload Me
End Sub

Private Sub mskInicioVigencia_GotFocus()
    On Error GoTo NotificaError
        pSelMkTexto mskInicioVigencia
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": mskInicioVigencia_GotFocus"))
    Unload Me
End Sub

Private Sub mskInicioVigencia_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        If KeyCode = vbKeyReturn Then mskFinVigencia.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": mskInicioVigencia_KeyDown"))
    Unload Me
End Sub

Private Sub optDescuento_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = vbKeyReturn Then pEnfocaOpt optVigencia
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": optDescuento_KeyPress"))
    Unload Me
End Sub

Private Sub OptPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = vbKeyReturn Then pEnfocaOpt optDescuento
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": OptPaciente_KeyPress"))
    Unload Me
End Sub

Private Sub optTipo_Click(Index As Integer)
    On Error GoTo NotificaError
   
    If Index = 0 Then
       FreConDescuento.Caption = "Pacientes con descuentos (I)=Interno (E)=Externo"
    ElseIf Index = 1 Then
       FreConDescuento.Caption = "Tipos de paciente con descuentos (I)=Interno (E)=Externo (U)=Urgencias (T)=Todos"
    Else
       FreConDescuento.Caption = "Empresas con descuentos (I)=Interno (E)=Externo (U)=Urgencias (T)=Todos"
    End If

    pCargaGuardados
    grdTipoAsignacion.Row = 1
    grdTipoAsignacion.Col = 3
    
    pHabilitaTipoPaciente
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": optTipo_Click"))
    Unload Me
End Sub

Private Sub pHabilitaTipoPaciente()
    On Error GoTo NotificaError

    optPaciente(0).Enabled = Trim(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 2)) = ""
    optPaciente(1).Enabled = Trim(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 2)) = ""
    optPaciente(2).Enabled = Trim(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 2)) = ""
    optPaciente(3).Enabled = IIf(optTipo(0).Value, False, Trim(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 2)) = "")
    
    If Trim(grdTipoAsignacion.TextMatrix(grdTipoAsignacion.Row, 2)) = "" Then
        optPaciente(0).Value = 1
    Else
        optPaciente(0).Value = 0
        optPaciente(1).Value = 0
        optPaciente(2).Value = 0
        optPaciente(3).Value = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaTipoPaciente"))
    Unload Me
End Sub

Private Sub optTipo_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = vbKeyReturn Then grdTipoAsignacion.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": optTipo_KeyPress"))
    Unload Me
End Sub

Private Sub optVigencia_Click(Index As Integer)
    On Error GoTo NotificaError
    
    mskInicioVigencia.Enabled = optVigencia(2).Value
    mskFinVigencia.Enabled = optVigencia(2).Value
    
    If Not optVigencia(2).Value Then
        mskInicioVigencia.Mask = ""
        mskInicioVigencia.Text = ""
        mskInicioVigencia.Mask = "##/##/####"
        mskFinVigencia.Mask = ""
        mskFinVigencia.Text = ""
        mskFinVigencia.Mask = "##/##/####"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": optVigencia_Click"))
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

Private Sub optVigencia_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If optVigencia(2).Value Then
            mskInicioVigencia.SetFocus
        Else
            cmdPreview.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": optVigencia_KeyDown"))
    Unload Me
End Sub
