VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptCuentaPendiente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas pendientes de facturar"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDetallado 
      Caption         =   "Detallado"
      Height          =   255
      Left            =   105
      TabIndex        =   12
      Top             =   3750
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Height          =   780
      Left            =   3645
      TabIndex        =   24
      Top             =   2910
      Width           =   4665
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   315
         Left            =   885
         TabIndex        =   10
         ToolTipText     =   "Fecha inicial"
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         ToolTipText     =   "Fecha final"
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   345
         TabIndex        =   26
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2460
         TabIndex        =   25
         Top             =   315
         Width           =   420
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   105
      TabIndex        =   22
      Top             =   0
      Width           =   8205
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   1515
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Height          =   675
      Left            =   3240
      TabIndex        =   18
      Top             =   3720
      Width           =   1080
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptCuentaPendiente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vista previa"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   540
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptCuentaPendiente.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprimir"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   105
      TabIndex        =   17
      Top             =   2910
      Width           =   3495
      Begin VB.OptionButton optRangoFechas 
         Caption         =   "Fecha del cargo"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   15
         Top             =   405
         Width           =   1575
      End
      Begin VB.OptionButton optRangoFechas 
         Caption         =   "Fecha de apertura de la cuenta"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   165
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   120
      TabIndex        =   16
      Top             =   675
      Width           =   8175
      Begin VB.CheckBox chkTickets 
         Caption         =   "Incluir tickets"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         ToolTipText     =   "Incluir tickets"
         Top             =   1065
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Internos egresados"
         Height          =   200
         Index           =   4
         Left            =   5400
         TabIndex        =   6
         Top             =   1065
         Width           =   1815
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Actualmente internos"
         Height          =   200
         Index           =   3
         Left            =   5400
         TabIndex        =   5
         Top             =   772
         Width           =   2055
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Internos"
         Height          =   200
         Index           =   2
         Left            =   4440
         TabIndex        =   4
         Top             =   772
         Width           =   975
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Externos"
         Height          =   200
         Index           =   1
         Left            =   3480
         TabIndex        =   3
         Top             =   772
         Width           =   975
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Todos"
         Height          =   200
         Index           =   0
         Left            =   2640
         TabIndex        =   2
         Top             =   772
         Width           =   855
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Seleccione el tipo de paciente"
         Top             =   255
         Width           =   5355
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Seleccione el departamento"
         Top             =   1440
         Width           =   5355
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pacientes"
         Height          =   200
         Left            =   120
         TabIndex        =   21
         Top             =   772
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento que abre la cuenta de externo y ventas al público"
         Height          =   390
         Left            =   120
         TabIndex        =   19
         Top             =   1380
         Width           =   2415
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmRptCuentaPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------
'| Nombre del proyecto      : prjCaja
'| Nombre del formulario    : frmRptCuentaPendiente
'-------------------------------------------------------------------------
'| Objetivo: Reporte de las cuentas pendientes de facturar por departamento
'-------------------------------------------------------------------------
Option Explicit

Public vllngNumOpcion As Long 'Número de opción, según el módulo donde corra (caja, supervisión y estadísticas)

Private vgrptReporte As CRAXDRT.Report
Dim vlblnpermiso As Boolean

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        optRangoFechas(0).SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_KeyDown"))
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
        cboDepartamento.ItemData(cboDepartamento.newIndex) = 0
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
        cboTipoPaciente.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboHospital_KeyDown"))
End Sub

Private Sub cboTipoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    
    If KeyCode = vbKeyReturn Then
        If optTipoPaciente(0).Value Then
            optTipoPaciente(0).SetFocus
        Else
            If optTipoPaciente(1).Value Then
                optTipoPaciente(1).SetFocus
            Else
                If optTipoPaciente(2).Value Then
                    optTipoPaciente(2).SetFocus
                End If
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipoPaciente_KeyDown"))
End Sub

Private Sub chkDetallado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkDetallado_KeyDown"))
End Sub

Private Sub chkTickets_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkTickets_KeyDown"))
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo NotificaError

    pImprime "P"
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPreview_Click"))
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo NotificaError


    pImprime "I"
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrint_Click"))
End Sub


Private Sub pImprime(vlstrDestino As String)
    On Error GoTo NotificaError


    Dim vlstrCveDepartamento As String
    Dim vlstrCveTipoPaciente As String
    Dim vlstrFechaInicio As String
    Dim vlstrFechaFin As String
    Dim vlStrTipoPaciente As String
    Dim rsReporte As New ADODB.Recordset
    Dim alstrParametros(9) As String
    Dim strTipo As String
     
    If fblnDatosValidos() Then
    
        Me.MousePointer = 11
    
        If chkDetallado.Value Then
            pInstanciaReporte vgrptReporte, "rptCuentaPendiente.rpt"
        Else
            pInstanciaReporte vgrptReporte, "rptCuentaPendienteConcentrado.rpt"
        End If
        vgrptReporte.DiscardSavedData
        
                
        vlstrCveDepartamento = Trim(Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)))
        vlstrCveTipoPaciente = Trim(Str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)))
        vlstrFechaInicio = fstrFechaSQL(mskFechaInicio.Text, , True)
        vlstrFechaFin = fstrFechaSQL(mskFechaFin.Text, , True)
        vlStrTipoPaciente = IIf(optTipoPaciente(0).Value, "*", IIf(optTipoPaciente(1).Value, "E", "I"))
        
        If optTipoPaciente(3).Value = True Or optTipoPaciente(4).Value = True Then
            
            ' Estatus del paciente
                        
            If optTipoPaciente(3).Value = True Then
            
                strTipo = "A"   ' Actualmente interno
                
            Else
                
                strTipo = "E"   ' Interno egresado
                
            End If
                
                vlStrTipoPaciente = "I"
        Else
            
            strTipo = "NA"
            
        End If
        
        
        vgstrParametrosSP = _
        vlstrCveDepartamento _
        & "|" & vlstrCveTipoPaciente _
        & "|" & vlstrFechaInicio _
        & "|" & vlstrFechaFin _
        & "|" & vlStrTipoPaciente _
        & "|" & IIf(optRangoFechas(0).Value, "0", "1") _
        & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex)) _
        & "|" & strTipo _
        & "|" & IIf(chkTickets.Value = 1, 1, 0)
        
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptCuentaPendiente")
        Set rsReporte = frsUltimoRecordset(rsReporte)
        If rsReporte.RecordCount > 0 Then
            alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
            alstrParametros(1) = "FechaInicio;" & CDate(mskFechaInicio.Text) & ";DATE"
            alstrParametros(2) = "FechaFin;" & CDate(mskFechaFin.Text) & ";DATE"
            alstrParametros(3) = "TituloReporte;" & UCase(IIf(cboDepartamento.ItemData(cboDepartamento.ListIndex) = 0, fRegresaParametro("VCHTITULOCTASPENDFACT", "PvParametro", 0), fRegresaParametro("VCHTITULOCTASPENDFACT", "PvParametro", 0) & " DE " & cboDepartamento.List(cboDepartamento.ListIndex)))
            alstrParametros(4) = "CveDepartamento;" & cboDepartamento.ItemData(cboDepartamento.ListIndex)
            alstrParametros(5) = "TipoReporte;" & IIf(optRangoFechas(0).Value, "0", "1") & ";NUMBER"
            alstrParametros(6) = "PacIntExt;" & IIf(optTipoPaciente(0).Value, "TODOS", IIf(optTipoPaciente(1).Value, "EXTERNOS", IIf(optTipoPaciente(2).Value, "INTERNOS", _
            IIf(optTipoPaciente(3).Value, "ACTUALMENTE INTERNOS", "INTERNOS EGRESADOS"))))
            alstrParametros(7) = "TipoDePac;" & cboTipoPaciente
            alstrParametros(8) = "Departamento;" & cboDepartamento.Text
            alstrParametros(9) = "Detallado;" & IIf(chkDetallado.Value = 0, False, True) & ";BOOLEAN"
            
            pCargaParameterFields alstrParametros, vgrptReporte
    
            pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Cuentas pendientes de facturar"
        Else
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rsReporte.Close

        Me.MousePointer = 0
    
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnDatosValidos = True
    
    If Not IsDate(mskFechaInicio.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!
        MsgBox SIHOMsg(254), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaInicio.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(mskFechaFin.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!
        MsgBox SIHOMsg(254), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaFin.SetFocus
    End If
    If fblnDatosValidos Then
        If CDate(mskFechaInicio.Text) > fdtmServerFecha Then
            fblnDatosValidos = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaInicio.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        If CDate(mskFechaFin.Text) > fdtmServerFecha Then
            fblnDatosValidos = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaFin.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        If CDate(mskFechaInicio.Text) > CDate(mskFechaFin.Text) Then
            fblnDatosValidos = False
            '¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaInicio.SetFocus
        End If
    End If
    


    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function

Private Sub Form_Activate()
    On Error GoTo NotificaError
    optTipoPaciente(1).Value = Not vlblnpermiso
    If cboDepartamento.Enabled Then
        cboDepartamento.SetFocus
    Else
        cboDepartamento.ListIndex = 0
        cboTipoPaciente.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Function fintPosiciona(vlintCveDepartamento As Integer) As Integer
    On Error GoTo NotificaError
    Dim vlintContador As Integer
    
    fintPosiciona = 0
    For vlintContador = 1 To cboDepartamento.ListCount - 1
        If cboDepartamento.ItemData(vlintContador) = vlintCveDepartamento Then
            fintPosiciona = vlintContador
        End If
    Next vlintContador

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fintPosiciona"))
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError


    If KeyAscii = 27 Then
        Unload Me
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon
    
    dtmfecha = fdtmServerFecha
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 1839
    Case "SE"
         lngNumOpcion = 2000
    End Select
   
    pCargaHospital lngNumOpcion
    
    vlstrSentencia = _
    "select " & _
        "AdTipoPaciente.vchDescripcion Descripcion," & _
        "AdTipoPaciente.tnyCveTipoPaciente*-1 Clave," & _
        "1 Orden " & _
    "From " & _
        "AdTipoPaciente " & _
    "Union " & _
    "select " & _
        "CcEmpresa.vchDescripcion Descripcion," & _
        "CcEmpresa.intCveEmpresa Clave," & _
        "2 Orden " & _
    "From " & _
        "CcEmpresa " & _
    "Where " & _
        "CcEmpresa.bitActivo = 1 " & _
    "Order By " & _
        "Orden," & _
        "Descripcion"
    pLlenarCboSentencia cboTipoPaciente, vlstrSentencia, 0, 1, "<TODOS>", 0
        
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = dtmfecha
    mskFechaInicio.Mask = "##/##/####"

    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"

    vlblnpermiso = fblnRevisaPermiso(vglngNumeroLogin, vllngNumOpcion, "C", True)
    cboDepartamento.Enabled = vlblnpermiso
    Label1.Enabled = vlblnpermiso
    optTipoPaciente(0).Enabled = vlblnpermiso
    optTipoPaciente(0).Value = vlblnpermiso

    

    Me.Caption = fRegresaParametro("VCHTITULOCTASPENDFACT", "PvParametro", 0)
    

    cboDepartamento.ListIndex = 0
    cboDepartamento.Enabled = False
    Label1.Enabled = False
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

Private Sub mskFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        If Trim(mskFechaFin.ClipText) = "" Then
            mskFechaFin.Mask = ""
            mskFechaFin.Text = fdtmServerFecha
            mskFechaFin.Mask = "##/##/####"
        End If
        
        SendKeys vbTab
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_KeyDown"))
End Sub

Private Sub mskFechaInicio_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFechaInicio


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_GotFocus"))
End Sub

Private Sub mskFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        If Trim(mskFechaInicio.ClipText) = "" Then
            mskFechaInicio.Mask = ""
            mskFechaInicio.Text = fdtmServerFecha
            mskFechaInicio.Mask = "##/##/####"
        End If
        mskFechaFin.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_KeyDown"))
End Sub

Private Sub optRangoFechas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        mskFechaInicio.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optRangoFechas_KeyDown"))

End Sub





Private Sub OptTipoPaciente_Click(Index As Integer)
    
    If optTipoPaciente(2).Value = True Or optTipoPaciente(3).Value = True Or optTipoPaciente(4) = True Then
        
        optTipoPaciente(3).Enabled = True
        optTipoPaciente(4).Enabled = True
       
    Else
            optTipoPaciente(3).Enabled = False
            optTipoPaciente(4).Enabled = False
        
    End If
    If optTipoPaciente(1).Value = True Then
        Label1.Enabled = True
        cboDepartamento.Enabled = True
        cboDepartamento.SetFocus
    Else
        cboDepartamento.ListIndex = 0
        Label1.Enabled = False
        cboDepartamento.Enabled = False
    End If
        
        
End Sub

Private Sub optTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoPaciente_KeyDown"))
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

