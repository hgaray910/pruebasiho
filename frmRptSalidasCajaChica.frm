VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptSalidasCajaChica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salidas de caja chica"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraImprime 
      Height          =   720
      Left            =   2160
      TabIndex        =   24
      Top             =   5040
      Width           =   1155
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   580
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptSalidasCajaChica.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   80
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptSalidasCajaChica.frx":03CD
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vista previa"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame fraFechas 
      Caption         =   " Rango de fechas de búsqueda"
      Height          =   750
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   5475
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
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
         Left            =   3840
         TabIndex        =   7
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
      Begin VB.Label lblFechaFin 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2880
         TabIndex        =   23
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblFechaIni 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   165
         TabIndex        =   22
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame fraBusqueda 
      Height          =   675
      Left            =   120
      TabIndex        =   19
      Top             =   1470
      Width           =   5475
      Begin VB.ComboBox cboBusqueda 
         Height          =   315
         Left            =   1230
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Seleccionar el proveedor"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblProveedor 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraTipo 
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   740
      Width           =   5475
      Begin VB.OptionButton optTipo 
         Caption         =   "Disminución de fondo"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   4
         ToolTipText     =   "Disminución de fondo"
         Top             =   290
         Width           =   1935
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Honorario"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   3
         ToolTipText     =   "Honorario"
         Top             =   290
         Width           =   1095
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Factura"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Factura"
         Top             =   290
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame fraFiltros 
      Height          =   1290
      Left            =   120
      TabIndex        =   25
      Top             =   2930
      Width           =   5475
      Begin VB.CheckBox chkSinDepositar 
         Caption         =   "Sin depositar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         ToolTipText     =   "Mostrar sin depositar"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkReembolsadas 
         Caption         =   "Reembolsadas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         ToolTipText     =   "Mostrar reembolsadas"
         Top             =   900
         Width           =   1455
      End
      Begin VB.CheckBox chkPendientes 
         Caption         =   "Pendientes"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         ToolTipText     =   "Mostrar pendientes"
         Top             =   570
         Width           =   1335
      End
      Begin VB.CheckBox chkDepositadas 
         Caption         =   "Depositadas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         ToolTipText     =   "Mostrar depositadas"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkCanceladas 
         Caption         =   "Canceladas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Mostrar canceladas"
         Top             =   900
         Width           =   1335
      End
      Begin VB.CheckBox chkActivas 
         Caption         =   "Activas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Mostrar activas"
         Top             =   570
         Width           =   1335
      End
      Begin VB.CheckBox chkTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Mostrar todas"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.Frame fraForma 
      Height          =   6135
      Left            =   -360
      TabIndex        =   26
      Top             =   -120
      Width           =   6135
      Begin VB.Frame FraConceptoDescripcion 
         Height          =   720
         Left            =   480
         TabIndex        =   29
         Top             =   4370
         Width           =   5475
         Begin VB.OptionButton optDescripcioSalida 
            Caption         =   "Descripción de la salida"
            Height          =   255
            Left            =   3120
            TabIndex        =   16
            ToolTipText     =   "Mostrará la descripción de la salida"
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optConceptoSalida 
            Caption         =   "Concepto de la salida"
            Height          =   255
            Left            =   600
            TabIndex        =   15
            ToolTipText     =   "Mostrará el concepto de la salida"
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Height          =   675
         Left            =   480
         TabIndex        =   27
         Top             =   170
         Width           =   5475
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Selección del departamento"
            Top             =   240
            Width           =   4005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Departamento "
            Height          =   195
            Left            =   165
            TabIndex        =   28
            Top             =   300
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frmRptSalidasCajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset 'Varios usos
Private vgrptReporte As CRAXDRT.Report

Private Sub cmdPreview_Click()
    pImprime "P"
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub mskFechaFin_GotFocus()
    On Error GoTo NotificaError

    pSelMkTexto mskFechaFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaFin_LostFocus()
    If Not IsDate(mskFechaFin.Text) Then
       MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
       Me.mskFechaFin.SetFocus
    End If
End Sub

Private Sub mskFechaIni_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFechaIni

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaIni_GotFocus"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    Select Case KeyCode
        Case 27
            Unload Me
        Case 13
            SendKeys vbTab
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
    Unload Me
    
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon
    
    pCargaProveedores
    
    dtmfecha = fdtmServerFecha
    
    mskFechaIni.Mask = ""
    mskFechaIni.Text = dtmfecha
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"
    
    cboDepartamento.Clear
    vgstrParametrosSP = "-1|1|*|" & vgintClaveEmpresaContable
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldepartamento")
  
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboDepartamento, rs, 0, 1
    End If
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
    cboDepartamento.ListIndex = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
    
End Sub

Private Sub chkTodas_Click()
    If chkTodas.Value = vbChecked Then
        chkActivas.Value = vbUnchecked
        chkActivas.Enabled = False
        chkCanceladas.Value = vbUnchecked
        chkCanceladas.Enabled = False
        chkDepositadas.Value = vbUnchecked
        chkDepositadas.Enabled = False
        chkPendientes.Value = vbUnchecked
        chkPendientes.Enabled = False
        chkReembolsadas.Value = vbUnchecked
        chkReembolsadas.Enabled = False
        chkSinDepositar.Value = vbUnchecked
        chkSinDepositar.Enabled = False
    Else
        chkActivas.Enabled = True
        chkCanceladas.Enabled = True
        chkDepositadas.Enabled = True
        chkPendientes.Enabled = True
        chkReembolsadas.Enabled = True
        chkSinDepositar.Enabled = True
    End If
End Sub

Private Sub mskFechaIni_LostFocus()
    If Not IsDate(mskFechaIni.Text) Then
       MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
       Me.mskFechaIni.SetFocus
    End If
End Sub

Private Sub optTipo_Click(Index As Integer)
    On Error GoTo NotificaError

    If optTipo(0).Value = True Then
        pCargaProveedores
        cboBusqueda.Enabled = True
        frmRptSalidasCajaChica.Height = 6330
        lblProveedor.Caption = "Proveedor"
        fraBusqueda.Visible = True
        FraConceptoDescripcion.Visible = True
        fraBusqueda.Top = 1470
        fraFechas.Top = 2160
        fraFiltros.Top = 2930
        fraImprime.Top = 5040 '4260
    ElseIf optTipo(1).Value = True Then
        pCargaProveedores
        cboBusqueda.Enabled = True
        frmRptSalidasCajaChica.Height = 6330
        lblProveedor.Caption = "Médico"
        fraBusqueda.Visible = True
        FraConceptoDescripcion.Visible = True
        fraBusqueda.Top = 1470
        fraFechas.Top = 2160
        fraFiltros.Top = 2930
        fraImprime.Top = 5040 '4260
    ElseIf optTipo(2).Value = True Then
        pCargaProveedores
        cboBusqueda.Enabled = False
        frmRptSalidasCajaChica.Height = 4770
        fraBusqueda.Visible = False
        FraConceptoDescripcion.Visible = False
        fraFechas.Top = 1470
        fraFiltros.Top = 2230
        fraImprime.Top = 3520
    End If
    
    dtmfecha = fdtmServerFecha

    mskFechaIni.Mask = ""
    mskFechaIni.Text = dtmfecha
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub pCargaProveedores()
    Dim strSentencia As String
    On Error GoTo NotificaError
    Dim strSentencia2 As String
    Dim vlblnbandera As Boolean
    Dim vlintCondador As Long

    cboBusqueda.Clear
    
    vlintCondador = 9999999
    
    strSentencia2 = "SELECT DISTINCT VCHNOMBRECOMERCIAL, VCHNOMBREPROVEEDOR, COPROVEEDOR.INTCVEPROVEEDOR, CPFACTURACAJACHICA.INTCVEPROVEEDOR AS " & Chr(34) & "CVEPROVEEDOR" & Chr(34) & " FROM COPROVEEDOR FULL JOIN CPFACTURACAJACHICA ON COPROVEEDOR.INTCVEPROVEEDOR = CPFACTURACAJACHICA.INTCVEPROVEEDOR ORDER BY VCHNOMBRECOMERCIAL"

    If optTipo(0).Value = True Then
        
        
        Set rs = frsRegresaRs(strSentencia2, adLockReadOnly, adOpenForwardOnly)
        
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                
                If rs!VCHNOMBREPROVEEDOR <> "" Then
                    cboBusqueda.AddItem rs!VCHNOMBREPROVEEDOR
                    vlblnbandera = True
                ElseIf rs!VCHNOMBRECOMERCIAL <> "" Then
                    cboBusqueda.AddItem rs!VCHNOMBRECOMERCIAL
                    vlblnbandera = True
                Else
                
                End If
                
                If rs!INTCVEPROVEEDOR <> "" Then
                    cboBusqueda.ItemData(cboBusqueda.newIndex) = rs!INTCVEPROVEEDOR
                    vlblnbandera = False
                ElseIf rs!CVEPROVEEDOR <> "" Then
                    cboBusqueda.ItemData(cboBusqueda.newIndex) = rs!CVEPROVEEDOR
                    vlblnbandera = False
                Else
                    If vlblnbandera = True Then
                        cboBusqueda.ItemData(cboBusqueda.newIndex) = vlintCondador
                        vlintCondador = vlintCondador - 1
                    Else
                    
                    End If
                End If
                
                rs.MoveNext
            Loop
        End If
    Else
        strSentencia = "Select Distinct CoProveedor.INTCVEPROVEEDOR Clave, CoProveedor.VCHNOMBRECOMERCIAL Nombre " & _
                                "  From CoProveedor Inner Join HOMedico On (CoProveedor.VCHRFC = HOMedico.VCHRFCMEDICO )  " & _
                                " Order by Nombre"
        Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
       If rs.RecordCount <> 0 Then
            Call pLlenarCboRs(cboBusqueda, rs, 0, 1, 0)
       End If
    End If
    
    cboBusqueda.AddItem "<TODOS>", 0
    cboBusqueda.ItemData(cboBusqueda.newIndex) = -1
    cboBusqueda.ListIndex = 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaProveedores"))
End Sub

Private Sub pImprime(vlstrTipo As String)

On Error GoTo NotificaError

    Dim rsSalidasCajaChica As New ADODB.Recordset
    Dim alstrParametros(4) As String
    Dim strParametros As String
    Dim vlIntProveedor As Long
    Dim vlstrRpt As String
    Dim strSentencia2 As String
    
    If chkTodas.Value = vbUnchecked And chkActivas.Value = vbUnchecked And chkCanceladas.Value = vbUnchecked And chkDepositadas.Value = vbUnchecked And chkPendientes.Value = vbUnchecked And chkReembolsadas.Value = vbUnchecked And chkSinDepositar.Value = vbUnchecked Then
        MsgBox SIHOMsg(1618), vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    Else
    
    End If
    
        
    If Me.cboBusqueda.ItemData(Me.cboBusqueda.ListIndex) <> -1 Then
        
        If optTipo(0).Value Then
            strTipo = "F"
            vlstrRpt = "rptSalidasCajaChica.rpt"
        End If
        
        If optTipo(1).Value Then
            strTipo = "H"
            vlstrRpt = "rptSalidasCajaChica.rpt"
        End If
        
        If optTipo(2).Value Then
            strTipo = "D"
            vlstrRpt = "rptSalidasCajaChicaDisminucion.rpt"
        End If
        
        If optTipo(0).Value = True Or optTipo(1).Value = True Then
            vlIntProveedor = cboBusqueda.ItemData(cboBusqueda.ListIndex)
        ElseIf optTipo(2).Value = True Then
            vlIntProveedor = -1
        End If
        
        If optDescripcioSalida.Value Then   ' JASM 20220203
            strTipoDesc = "D"
        Else
            strTipoDesc = "C"
        End If
        
        strFechaIni = fstrFechaSQL(Format(CDate(mskFechaIni.Text), "dd/mm/yyyy"))
        strFechaFin = fstrFechaSQL(Format(CDate(mskFechaFin.Text), "dd/mm/yyyy"))
        
'        strParametros = CStr("-1") _
'        & "|" & strFechaIni _
'        & "|" & strFechaFin _
'        & "|" & CStr(1) _
'        & "|" & CStr(vlIntProveedor) _
'        & "|" & CStr(vgintNumeroDepartamento) _
'        & "|" & vgintClaveEmpresaContable _
'        & "|" & strTipo _
'        & "|" & Me.cboBusqueda.Text _
'        & "|" & Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex) _
'        & "|" & IIf(chkTodas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkActivas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkDepositadas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkPendientes.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkReembolsadas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkSinDepositar.Value = vbChecked, 1, 0)
        
        strParametros = CStr("-1") _
        & "|" & strFechaIni _
        & "|" & strFechaFin _
        & "|" & CStr(1) _
        & "|" & CStr(vlIntProveedor) _
        & "|" & CStr(-1) _
        & "|" & vgintClaveEmpresaContable _
        & "|" & strTipo _
        & "|" & Me.cboBusqueda.Text _
        & "|" & Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex) _
        & "|" & IIf(chkTodas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkActivas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkDepositadas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkPendientes.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkReembolsadas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkSinDepositar.Value = vbChecked, 1, 0)
    
        pInstanciaReporte vgrptReporte, vlstrRpt
                    
        Set rsSalidasCajaChica = frsEjecuta_SP(strParametros, "SP_CPSELFACTURACAJACHICA")
        
        If rsSalidasCajaChica.RecordCount <> 0 Then
                
                alstrParametros(0) = "lblNombreHospital;" & Trim(vgstrNombreHospitalCH)
                alstrParametros(1) = "FechaIni" & ";" & UCase(Format(mskFechaIni.Text, "dd/mmm/yyyy"))
                alstrParametros(2) = "FechaFin" & ";" & UCase(Format(mskFechaFin.Text, "dd/mmm/yyyy"))
                alstrParametros(3) = "lblDepartamento;" & cboDepartamento.List(cboDepartamento.ListIndex)
                alstrParametros(4) = "lblTipoDesc;" & IIf(optDescripcioSalida.Value = True, "D", "C") ' JASM 20220203
                pCargaParameterFields alstrParametros, vgrptReporte
                pImprimeReporte vgrptReporte, rsSalidasCajaChica, IIf(vlstrTipo = "P", "P", "I"), "Salidas de caja chica"
        
        Else
            'No existe información con esos parámetros
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        rsSalidasCajaChica.Close

        
        
    Else
    
    
    
    
    
    If optTipo(0).Value Then
                strTipo = "F"
            vlstrRpt = "rptSalidasCajaChica.rpt"
        End If
        
        If optTipo(1).Value Then
            strTipo = "H"
            vlstrRpt = "rptSalidasCajaChica.rpt"
        End If
        
        If optTipo(2).Value Then
            strTipo = "D"
            vlstrRpt = "rptSalidasCajaChicaDisminucion.rpt"
        End If
        
        If optTipo(0).Value = True Or optTipo(1).Value = True Then
            vlIntProveedor = cboBusqueda.ItemData(cboBusqueda.ListIndex)
        ElseIf optTipo(2).Value = True Then
            vlIntProveedor = -1
        End If
        
        If optDescripcioSalida.Value Then   ' JASM 20220203
            strTipoDesc = "D"
        Else
            strTipoDesc = "C"
        End If
        
        strFechaIni = fstrFechaSQL(Format(CDate(mskFechaIni.Text), "dd/mm/yyyy"))
        strFechaFin = fstrFechaSQL(Format(CDate(mskFechaFin.Text), "dd/mm/yyyy"))
        
'        strParametros = CStr("-1") _
'        & "|" & strFechaIni _
'        & "|" & strFechaFin _
'        & "|" & CStr(1) _
'        & "|" & CStr(vlIntProveedor) _
'        & "|" & CStr(vgintNumeroDepartamento) _
'        & "|" & vgintClaveEmpresaContable _
'        & "|" & strTipo _
'        & "|" & Me.cboBusqueda.Text _
'        & "|" & Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex) _
'        & "|" & IIf(chkTodas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkActivas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkDepositadas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkPendientes.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkReembolsadas.Value = vbChecked, 1, 0) _
'        & "|" & IIf(chkSinDepositar.Value = vbChecked, 1, 0)

        strParametros = CStr("-1") _
        & "|" & strFechaIni _
        & "|" & strFechaFin _
        & "|" & CStr(1) _
        & "|" & CStr(vlIntProveedor) _
        & "|" & CStr(-1) _
        & "|" & vgintClaveEmpresaContable _
        & "|" & strTipo _
        & "|" & Me.cboBusqueda.Text _
        & "|" & Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex) _
        & "|" & IIf(chkTodas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkActivas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkDepositadas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkPendientes.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkReembolsadas.Value = vbChecked, 1, 0) _
        & "|" & IIf(chkSinDepositar.Value = vbChecked, 1, 0)
    
        pInstanciaReporte vgrptReporte, vlstrRpt
                    
        Set rsSalidasCajaChica = frsEjecuta_SP(strParametros, "SP_CPSELFACTURACAJACHICA")
        
        If rsSalidasCajaChica.RecordCount <> 0 Then
                
                alstrParametros(0) = "lblNombreHospital;" & Trim(vgstrNombreHospitalCH)
                alstrParametros(1) = "FechaIni" & ";" & UCase(Format(mskFechaIni.Text, "dd/mmm/yyyy"))
                alstrParametros(2) = "FechaFin" & ";" & UCase(Format(mskFechaFin.Text, "dd/mmm/yyyy"))
                alstrParametros(3) = "lblDepartamento;" & cboDepartamento.List(cboDepartamento.ListIndex)
                alstrParametros(4) = "lblTipoDesc;" & IIf(optDescripcioSalida.Value = True, "D", "C") ' JASM 20220203
                pCargaParameterFields alstrParametros, vgrptReporte
                pImprimeReporte vgrptReporte, rsSalidasCajaChica, IIf(vlstrTipo = "P", "P", "I"), "Salidas de caja chica"
        
        Else
            'No existe información con esos parámetros
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        rsSalidasCajaChica.Close

    End If
            
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

