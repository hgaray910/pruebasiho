VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporteIngresosTurno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos por turno"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   20
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1080
      Left            =   60
      TabIndex        =   17
      Top             =   630
      Width           =   7125
      Begin VB.ComboBox CboTurnos 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Turno de trabajo"
         Top             =   240
         Width           =   5760
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Departamentos"
         Top             =   615
         Width           =   5760
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   675
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de reporte"
      Height          =   1065
      Left            =   2085
      TabIndex        =   15
      Top             =   1740
      Width           =   1980
      Begin VB.OptionButton OptDetallado 
         Caption         =   "Detallado"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "Muestra el reporte detallado"
         Top             =   675
         Width           =   1470
      End
      Begin VB.OptionButton OptAcumulado 
         Caption         =   "Concentrado"
         Height          =   210
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   "Muestra el reporte concentrado"
         Top             =   420
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupo"
      Height          =   1065
      Left            =   60
      TabIndex        =   13
      Top             =   1740
      Width           =   1980
      Begin VB.OptionButton optGrupoFormaPago 
         Caption         =   "Forma de pago"
         Height          =   210
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Agrupado por forma de pago"
         Top             =   390
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton optGrupoTipoDocumento 
         Caption         =   "Tipo documento"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   "Agrupado por tipo de documento"
         Top             =   645
         Width           =   1470
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3045
      TabIndex        =   12
      Top             =   2880
      Width           =   1140
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteIngresosTurno.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteIngresosTurno.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame freRangos 
      Caption         =   "Rango de reporte"
      Height          =   1065
      Left            =   4095
      TabIndex        =   11
      Top             =   1740
      Width           =   3090
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   285
         Left            =   1605
         TabIndex        =   8
         ToolTipText     =   "Fecha final"
         Top             =   645
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   37351
      End
      Begin MSComCtl2.DTPicker dtpFechaInicio 
         Height          =   285
         Left            =   1605
         TabIndex        =   7
         ToolTipText     =   "Fecha inicial"
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   37351
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   345
         Width           =   465
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   690
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmReporteIngresosTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmReporteIngresosTurno                                     -
'-------------------------------------------------------------------------------------
'| Objetivo: Es el reporte de ingresos de caja por turno
'-------------------------------------------------------------------------------------

Private vgrptReporte As CRAXDRT.Report
Dim rsTemp As New ADODB.Recordset

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaOpt optGrupoFormaPago
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
        CboTurnos.SetFocus
    End If

End Sub

Private Sub CboTurnos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboDepartamento.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub
Sub pImprime(pstrDestino As String)
    On Error GoTo NotificaError
    
    Dim vlstrsql As String
    Dim alstrParametros(13) As String
    Dim vlstrFechaIni As String
    Dim vlstrFechaFin As String
                
    Dim lngNumCorte As Long
    Dim intIntermedios As Integer
    Dim intFacturas As Integer
    Dim intRecibos As Integer
    Dim intTickets As Integer
    Dim intSalidas As Integer
    Dim intSoloCancelados As Integer
    Dim intPagosCredito As Integer
    Dim intHonorarios As Integer
    Dim intFondoFijo As Integer
    Dim intTransferencia As Integer
    Dim intSalidaCajaChica As Integer
    Dim intEntradaCajaChica As Integer
    Dim intTipoOrden As Integer
    Dim lngCveEmpleado As Long
    Dim intAgrupadoEmpleado As Integer
    Dim intCveDepto As Integer
                
    mskHoraIni = "00:00"
    mskHoraFin = "23:59"
    If CboTurnos.ListIndex <> 0 Then
        vlstrsql = " SELECT dtmhorainicio, dtmhorafin "
        vlstrsql = vlstrsql & " FROM NOTURNO"
        vlstrsql = vlstrsql & " WHERE intclave = " & CboTurnos.ItemData(CboTurnos.ListIndex)
        Set rsTemp = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        If rsTemp.RecordCount > 0 Then
            mskHoraIni = Format(rsTemp!dtmhorainicio, "hh:mm")
            mskHoraFin = Format(rsTemp!dtmHoraFin, "hh:mm")
        End If
    End If
    
    lngNumCorte = 0
    intIntermedios = 1
    vlstrFechaIni = fstrFechaSQL(DtpFechaInicio.Value, mskHoraIni & ":00", True)
    vlstrFechaFin = fstrFechaSQL(DtpFechaFin.Value, mskHoraFin & ":00", True)
    intCveDepto = cboDepartamento.ItemData(cboDepartamento.ListIndex)
    intFacturas = 1
    intRecibos = 1
    intTickets = 1
    intSalidas = 1
    intSoloCancelados = 0
    intPagosCredito = 1
    intHonorarios = 1
    intFondoFijo = 1
    intTransferencia = 1
    intSalidaCajaChica = 0
    intEntradaCajaChica = 0
    intTipoOrden = IIf(optGrupoFormaPago.Value, 1, 2)
    lngCveEmpleado = -1
    intAgrupadoEmpleado = 1
    
    vgstrParametrosSP = _
    CStr(lngNumCorte) _
    & "|" & CStr(intIntermedios) _
    & "|" & vlstrFechaIni _
    & "|" & vlstrFechaFin _
    & "|" & CStr(intCveDepto) _
    & "|" & CStr(intFacturas) _
    & "|" & CStr(intRecibos) _
    & "|" & CStr(intTickets) _
    & "|" & CStr(intSalidas) _
    & "|" & CStr(intSoloCancelados) _
    & "|" & CStr(intPagosCredito) _
    & "|" & CStr(intHonorarios) _
    & "|" & CStr(intFondoFijo) _
    & "|" & CStr(intTransferencia) _
    & "|" & CStr(intSalidaCajaChica) _
    & "|" & CStr(intEntradaCajaChica) _
    & "|" & CStr(intTipoOrden) _
    & "|" & CStr(lngCveEmpleado) _
    & "|" & CStr(intAgrupadoEmpleado) _
    & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
    
    Set rsTemp = frsEjecuta_SP(vgstrParametrosSP, "sp_pvselcortemovimiento")

    If rsTemp.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        pInstanciaReporte vgrptReporte, "rptDetalleMovimientos.rpt"
        vgrptReporte.DiscardSavedData
        
        alstrParametros(0) = "FechaInicio;" & Format(vlstrFechaIni, "dd/mmm/yyyy hh:mm")
        alstrParametros(1) = "FechaFin;" & Format(vlstrFechaFin, "dd/mmm/yyyy hh:mm")
        alstrParametros(2) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
        alstrParametros(3) = "NumCorte;" & "0"
        alstrParametros(4) = "Titulo;" & "INGRESOS DE CAJA POR TURNO"
        alstrParametros(5) = "Departamento;" & cboDepartamento.Text
        alstrParametros(6) = "PersonaAbre;" & " "
        alstrParametros(7) = "PersonaCierra;" & " "
        alstrParametros(8) = "Ejercicio;" & "0"
        alstrParametros(9) = "Mes;" & "0"
        alstrParametros(10) = "Folio;" & "0"
        alstrParametros(11) = "Acumulado;" & IIf(OptAcumulado.Value, 1, 0)
        alstrParametros(12) = "TipoCorte;" & "Caja ingresos"
        alstrParametros(13) = "TipoCorteCaracter;" & "P"
        
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsTemp, pstrDestino, "Ingresos de caja por turno"

        pInstanciaReporte vgrptReporte, "rptDetalleMovimientos.rpt"
        vgrptReporte.DiscardSavedData
    End If
    rsTemp.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

Private Sub dtpFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cmdVistaPreliminar.SetFocus
    End If

End Sub

Private Sub DtpFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        DtpFechaFin.SetFocus
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Dim vlstrsentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsTurnos As New ADODB.Recordset
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 2011
    Case "SE"
         lngNumOpcion = 1769
    End Select
    
    pCargaHospital lngNumOpcion
    
    DtpFechaInicio.Year = Year(Date)
    DtpFechaInicio.Month = Month(Date)
    DtpFechaInicio.Day = Day(Date)
    DtpFechaFin.Year = Year(Date)
    DtpFechaFin.Month = Month(Date)
    DtpFechaFin.Day = Day(Date)
    
    dtmfecha = fdtmServerFecha
    
    DtpFechaInicio.Value = dtmfecha
    DtpFechaFin.Value = dtmfecha
    
    'Combo de turnos.
    vlstrsentencia = "SELECT intClave, vchDescripcion From NoTurno Where (bitEstatus = 1)"
    Set rsTurnos = frsRegresaRs(vlstrsentencia, adLockReadOnly, adOpenForwardOnly)
    If rsTurnos.RecordCount > 0 Then
        pLlenarCboRs CboTurnos, rsTurnos, 0, 1
        CboTurnos.ListIndex = 0
    Else
        MsgBox SIHOMsg(13) & Chr(13) & "Dato: " & CboTurnos.ToolTipText, vbExclamation, "Mensaje"
        vgblnExistenDatos = False
        Unload Me
        Exit Sub
    End If
    rsTurnos.Close
    
    CboTurnos.AddItem "<TODOS>", 0
    CboTurnos.ItemData(0) = -1
    CboTurnos.ListIndex = 0
      
End Sub



Private Sub OptAcumulado_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        DtpFechaInicio.SetFocus
    End If

End Sub

Private Sub OptDetallado_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        DtpFechaInicio.SetFocus
    End If

End Sub

Private Sub optGrupoFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If OptAcumulado.Value Then
            OptAcumulado.SetFocus
        Else
            OptDetallado.SetFocus
        End If
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

