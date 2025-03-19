VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptDescuentosAplicados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de descuentos aplicados"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6045
      Left            =   45
      TabIndex        =   14
      Top             =   45
      Width           =   4860
      Begin VB.CheckBox chkDetallado 
         Caption         =   "Detallado"
         Height          =   330
         Left            =   285
         TabIndex        =   23
         Top             =   5550
         Width           =   1845
      End
      Begin VB.Frame Frame7 
         Caption         =   "Estado del descuento"
         Height          =   705
         Left            =   255
         TabIndex        =   22
         Top             =   3735
         Width           =   4290
         Begin VB.OptionButton optEdoDscto 
            Caption         =   "Facturados"
            Height          =   345
            Index           =   0
            Left            =   225
            TabIndex        =   7
            Top             =   270
            Width           =   1215
         End
         Begin VB.OptionButton optEdoDscto 
            Caption         =   "No facturados"
            Height          =   345
            Index           =   1
            Left            =   1635
            TabIndex        =   8
            Top             =   255
            Width           =   1410
         End
         Begin VB.OptionButton optEdoDscto 
            Caption         =   "Todos"
            Height          =   345
            Index           =   2
            Left            =   3165
            TabIndex        =   9
            Top             =   255
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo de paciente"
         Height          =   780
         Left            =   315
         TabIndex        =   21
         Top             =   1860
         Width           =   4275
         Begin VB.OptionButton optAmbos 
            Caption         =   "Ambos"
            Height          =   200
            Left            =   3015
            TabIndex        =   5
            Top             =   405
            Width           =   960
         End
         Begin VB.OptionButton optExternos 
            Caption         =   "Externos"
            Height          =   200
            Left            =   1500
            TabIndex        =   4
            Top             =   405
            Width           =   960
         End
         Begin VB.OptionButton optInternos 
            Caption         =   "Internos"
            Height          =   200
            Left            =   75
            TabIndex        =   3
            Top             =   405
            Width           =   960
         End
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Selección de la empresa"
         Top             =   1245
         Width           =   4170
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   345
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Selección del tipo de paciente"
         Top             =   645
         Width           =   4170
      End
      Begin VB.Frame Frame3 
         Caption         =   "Rango de fechas "
         Height          =   855
         Left            =   255
         TabIndex        =   15
         Top             =   4605
         Width           =   4215
         Begin MSMask.MaskEdBox mskFecIni 
            Height          =   315
            Left            =   720
            TabIndex        =   10
            ToolTipText     =   "Fecha inicial"
            Top             =   360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFecFin 
            Height          =   315
            Left            =   2520
            TabIndex        =   11
            ToolTipText     =   "Fecha final"
            Top             =   360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   2160
            TabIndex        =   17
            Top             =   440
            Width           =   90
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   440
            Width           =   210
         End
      End
      Begin VB.ComboBox CboConceptoFacturacion 
         Height          =   315
         Left            =   330
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Filtro de concepto de facturación"
         Top             =   3315
         Width           =   4170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   405
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label lblConceptoFact 
         Caption         =   "Concepto de facturación"
         Height          =   240
         Left            =   315
         TabIndex        =   18
         Top             =   3075
         Width           =   2250
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   1890
      TabIndex        =   0
      Top             =   6255
      Width           =   1140
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   585
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptDescuentosAplicados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptDescuentosAplicados.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmRptDescuentosAplicados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------
' Reporte descuentos aplicados
' Fecha de programación: 8 de febrero del 2006
'--------------------------------------------------------------------------------------
Dim vlstrx As String
Private vgrptReporte As CRAXDRT.Report
Public vglngNumeroOpcion As Long

Private Sub cboConceptoFacturacion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
            If optEdoDscto(0).Value Then
                optEdoDscto(0).SetFocus
            Else
                If optEdoDscto(1).Value Then
                    optEdoDscto(1).SetFocus
                Else
                    optEdoDscto(2).SetFocus
                End If
            End If

    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboConceptoFacturacion_KeyPress"))
End Sub


Private Sub cboEmpresa_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If optAmbos.Value Then
            optAmbos.SetFocus
        Else
            If optInternos.Value Then
                optInternos.SetFocus
            Else
                optExternos.SetFocus
            End If
        End If

    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpresa_KeyPress"))
End Sub

Private Sub cboTipoPaciente_Click()
    On Error GoTo NotificaError
    
    Dim X As Long
    Dim vlblnTermina As Boolean

    If cboTipoPaciente.ListIndex > 0 Then
        
        vlstrx = "select bitUtilizaConvenio from AdTipoPaciente where tnyCveTipoPaciente=" & Str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex))
        If Not IIf(frsRegresaRs(vlstrx).Fields(0) = 1, True, False) Then
            cboEmpresa.Enabled = False
        Else
            cboEmpresa.Enabled = True
        End If
    
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoPaciente_Click"))
End Sub

Private Sub cboTipoPaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        
        cboTipoPaciente_Click
        
        If cboEmpresa.Enabled Then
            cboEmpresa.SetFocus
        Else
            If optAmbos.Value Then
                optAmbos.SetFocus
            Else
                If optInternos.Value Then
                    optInternos.SetFocus
                Else
                    optExternos.SetFocus
                End If
            End If
        End If

    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoPaciente_KeyPress"))
End Sub

Private Sub chkDetallado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdVistaPreliminar.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkDetallado_KeyPress"))
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
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
    
    Dim rs As New ADODB.Recordset
    pInstanciaReporte vgrptReporte, "rptDescuentosAplicados.rpt"
    ' Tipos de paciente
    vlstrx = " select vchDescripcion Descripcion, tnyCveTipoPaciente Clave From AdTipoPaciente "
    Set rs = frsRegresaRs(vlstrx)
    pLlenarCboRs cboTipoPaciente, rs, 1, 0, 3
    cboTipoPaciente.ListIndex = 0
    
    'Empresas
    vlstrx = "Select vchDescripcion Descripcion, intCveEmpresa Clave From CcEmpresa "
    Set rs = frsRegresaRs(vlstrx)
    pLlenarCboRs cboEmpresa, rs, 1, 0, 3
    cboEmpresa.ListIndex = 0
    
    Call pLlenarcboConceptoFacturacion
    vginttipoorden = "A"
   
    mskFecIni.Mask = ""
    mskFecIni.Text = fdtmServerFecha
    mskFecIni.Mask = "##/##/####"
    
    mskFecFin.Mask = ""
    mskFecFin.Text = fdtmServerFecha
    mskFecFin.Mask = "##/##/####"
    
    optAmbos.Value = True
    optEdoDscto(2).Value = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub


Private Sub mskFecFin_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFecFin

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_GotFocus"))
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        chkDetallado.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_KeyPress"))
End Sub

Private Sub mskFecFin_LostFocus()
    On Error GoTo NotificaError

    If Trim(mskFecFin.ClipText) = "" Then
        mskFecFin.Mask = ""
        mskFecFin.Text = fdtmServerFecha
        mskFecFin.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecFin.Text) Then
            mskFecFin.Mask = ""
            mskFecFin.Text = fdtmServerFecha
            mskFecFin.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_LostFocus"))
End Sub

Private Sub mskFecIni_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFecIni

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_GotFocus"))
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFecFin.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_KeyPress"))
End Sub

Private Sub mskFecIni_LostFocus()
    On Error GoTo NotificaError
    
    If Trim(mskFecIni.ClipText) = "" Then
        mskFecIni.Mask = ""
        mskFecIni.Text = fdtmServerFecha
        mskFecIni.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecIni.Text) Then
            mskFecIni.Mask = ""
            mskFecIni.Text = fdtmServerFecha
            mskFecIni.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_LostFocus"))
End Sub

Private Sub optAmbos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        CboConceptoFacturacion.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optAmbos_KeyPress"))
End Sub

Private Sub optEdoDscto_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFecIni.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optInternos_KeyPress"))

End Sub

Private Sub optExternos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        CboConceptoFacturacion.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optExternos_KeyPress"))
End Sub

Private Sub optInternos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        CboConceptoFacturacion.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optInternos_KeyPress"))
End Sub

Sub pImprime(pstrDestino As String)
    Dim alstrParametros(2) As String
    Dim vlrsPvSelFacturasCanceladas As New ADODB.Recordset
    
    On Error GoTo NotificaError
    
    Set vlrsPvSelFacturasCanceladas = frsEjecuta_SP(IIf(optAmbos.Value, "A", IIf(optInternos.Value, "I", "E")) & "|" & _
                                      IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) = 0, -1, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)) & "|" & _
                                      IIf(cboEmpresa.ItemData(cboEmpresa.ListIndex) = 0, -1, cboEmpresa.ItemData(cboEmpresa.ListIndex)) & "|" & _
                                      IIf(optEdoDscto(0).Value, "F", IIf(optEdoDscto(1).Value, "N", "T")) & "|" & _
                                      IIf(CboConceptoFacturacion.ItemData(CboConceptoFacturacion.ListIndex) = 0, -1, CboConceptoFacturacion.ItemData(CboConceptoFacturacion.ListIndex)) & "|" & _
                                      fstrFechaSQL(mskFecIni.Text, "00:00:00", False) & "|" & _
                                      fstrFechaSQL(mskFecFin.Text, "23:59:59", False), "SP_PVRPTDESCUENTOSAPLICADOS")
    
    If vlrsPvSelFacturasCanceladas.EOF Then
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    Else
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "ImprimeDetalle;" & chkDetallado.Value
        alstrParametros(1) = "RangoFechas;DEL " & Trim(Format(mskFecIni, "DD/MMM/YYYY") & " AL " & Format(mskFecFin, "DD/MMM/YYYY"))
        alstrParametros(2) = "NombreHospital; " & Trim(vgstrNombreHospitalCH)
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, vlrsPvSelFacturasCanceladas, pstrDestino, "Facturas canceladas"
    End If
    
    
    vlrsPvSelFacturasCanceladas.Close
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
End Sub

Private Sub pLlenarcboConceptoFacturacion()
'------------------------------------------------------------------------
' Llena el combo del departamento para poder filtrar las entradas/salidas
'------------------------------------------------------------------------
    On Error GoTo NotificaError
        
    Dim rsConceptoFact As New ADODB.Recordset
    Dim vlstrsql As String
    
    
    vlstrsql = " select smiCveconcepto, chrDescripcion from pvconceptofacturacion where bitactivo = 1"
    Set rsConceptoFact = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    Call pLlenarCboRs(CboConceptoFacturacion, rsConceptoFact, 0, 1, -1)
    
    rsConceptoFact.Close
    
    CboConceptoFacturacion.AddItem "<TODOS>", 0
    CboConceptoFacturacion.ItemData(0) = -1
    CboConceptoFacturacion.ListIndex = 0
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarcboConceptoFacturacion"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub


