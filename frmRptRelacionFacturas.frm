VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptRelacionFacturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   120
      TabIndex        =   34
      Top             =   -15
      Width           =   7365
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   1215
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6000
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3232
      TabIndex        =   27
      Top             =   5400
      Width           =   1140
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   585
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptRelacionFacturas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprimir"
         Top             =   165
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptRelacionFacturas.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Vista previa"
         Top             =   165
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4725
      Left            =   120
      TabIndex        =   22
      Top             =   630
      Width           =   7365
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   1200
         TabIndex        =   36
         Top             =   3960
         Width           =   6015
         Begin VB.CheckBox chkFoliosFiscales 
            Caption         =   "Mostrar folios fiscales"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "Mostrar folio fiscal"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   4230
         TabIndex        =   31
         Top             =   2670
         Width           =   2985
         Begin VB.CheckBox chkFolios 
            Caption         =   "Por folios"
            Height          =   195
            Left            =   225
            TabIndex        =   16
            Top             =   45
            Width           =   1485
         End
         Begin VB.TextBox txtASerie 
            Height          =   315
            Left            =   1185
            MaxLength       =   12
            TabIndex        =   18
            Top             =   720
            Width           =   1305
         End
         Begin VB.TextBox txtDeSerie 
            Height          =   315
            Left            =   1185
            MaxLength       =   12
            TabIndex        =   17
            Top             =   375
            Width           =   1305
         End
         Begin VB.Label lblDesdeFolio 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   510
            TabIndex        =   33
            Top             =   435
            Width           =   465
         End
         Begin VB.Label lblHastaFolio 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   510
            TabIndex        =   32
            Top             =   780
            Width           =   420
         End
      End
      Begin VB.Frame fraFechas 
         Height          =   1215
         Left            =   1215
         TabIndex        =   28
         Top             =   2670
         Width           =   2955
         Begin VB.CheckBox chkFechas 
            Caption         =   "Por fechas"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   45
            Width           =   1905
         End
         Begin MSMask.MaskEdBox mskFecIni 
            Height          =   315
            Left            =   1185
            TabIndex        =   14
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
            Left            =   1185
            TabIndex        =   15
            ToolTipText     =   "Fecha final"
            Top             =   720
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label lblDesdeFecha 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   465
            TabIndex        =   30
            Top             =   420
            Width           =   465
         End
         Begin VB.Label lblHastaFecha 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   465
            TabIndex        =   29
            Top             =   780
            Width           =   420
         End
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1215
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Filtro departamentos"
         Top             =   240
         Width           =   6000
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Selección del tipo de paciente"
         Top             =   1530
         Width           =   6000
      End
      Begin VB.Frame Frame4 
         Caption         =   "Facturada"
         Height          =   870
         Left            =   1215
         TabIndex        =   24
         Top             =   600
         Width           =   6000
         Begin VB.OptionButton optSocios 
            Caption         =   "Socios"
            Height          =   200
            Left            =   4005
            TabIndex        =   7
            Top             =   555
            Width           =   960
         End
         Begin VB.OptionButton optGrupo 
            Caption         =   "Grupo de cuentas"
            Height          =   200
            Left            =   4005
            TabIndex        =   6
            Top             =   330
            Width           =   1620
         End
         Begin VB.OptionButton optCliente 
            Caption         =   "Clientes"
            Height          =   200
            Left            =   2025
            TabIndex        =   5
            Top             =   555
            Width           =   960
         End
         Begin VB.OptionButton optInternos 
            Caption         =   "Pacientes internos"
            Height          =   200
            Left            =   135
            TabIndex        =   3
            Top             =   555
            Width           =   1740
         End
         Begin VB.OptionButton optExternos 
            Caption         =   "Pacientes externos"
            Height          =   200
            Left            =   2025
            TabIndex        =   4
            Top             =   330
            Width           =   1755
         End
         Begin VB.OptionButton optAmbos 
            Caption         =   "Todos"
            Height          =   200
            Left            =   135
            TabIndex        =   2
            Top             =   330
            Width           =   960
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Estado de la factura"
         Height          =   705
         Left            =   1215
         TabIndex        =   23
         Top             =   1890
         Width           =   6000
         Begin VB.OptionButton optEstado 
            Caption         =   "Crédito"
            Height          =   345
            Index           =   3
            Left            =   3750
            TabIndex        =   12
            Top             =   270
            Width           =   990
         End
         Begin VB.OptionButton optEstado 
            Caption         =   "Todos"
            Height          =   345
            Index           =   2
            Left            =   150
            TabIndex        =   9
            Top             =   270
            Width           =   840
         End
         Begin VB.OptionButton optEstado 
            Caption         =   "Pagada"
            Height          =   345
            Index           =   1
            Left            =   2610
            TabIndex        =   11
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton optEstado 
            Caption         =   "Cancelada"
            Height          =   345
            Index           =   0
            Left            =   1305
            TabIndex        =   10
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Label lblDepartamento 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo paciente"
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   1590
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRptRelacionFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------
' Reporte descuentos aplicados
' Fecha de programación: 8 de febrero del 2006
'--------------------------------------------------------------------------------------
Option Explicit

Dim vlstrx As String
Private vgrptReporte As CRAXDRT.Report
Public vglngNumeroOpcion As Long

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
        cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
        cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, Str(vgintNumeroDepartamento))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me

End Sub

Private Sub chkFechas_Click()

    mskFecIni.Enabled = chkFechas.Value = 1
    mskFecFin.Enabled = chkFechas.Value = 1
    lblDesdeFecha.Enabled = chkFechas.Value = 1
    lblHastaFecha.Enabled = chkFechas.Value = 1
    chkFolios.Value = IIf(chkFechas.Value = 1, 0, 1)

End Sub

Private Sub chkFolios_Click()

    txtDeSerie.Enabled = chkFolios.Value = 1
    txtASerie.Enabled = chkFolios.Value = 1
    lblDesdeFolio.Enabled = chkFolios.Value = 1
    lblHastaFolio.Enabled = chkFolios.Value = 1
    
    chkFechas.Value = IIf(chkFolios.Value = 1, 0, 1)

End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub Form_Activate()
    
    cboDepartamento.Enabled = fblnRevisaPermiso(vglngNumeroLogin, vglngNumeroOpcion, "C", True)

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
    Else
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    
    Me.Icon = frmMenuPrincipal.Icon
    
    pInstanciaReporte vgrptReporte, "rptRelacionFacturas.rpt"
    
    dtmfecha = fdtmServerFecha
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 1881
    Case "SE"
         lngNumOpcion = 2003
    End Select
    
    pCargaHospital lngNumOpcion
    
    lblDepartamento.Enabled = cboHospital.Enabled
    cboDepartamento.Enabled = cboHospital.Enabled
    
    ' Tipos de paciente
    Set rs = frsEjecuta_SP("2", "sp_GnSelTipoPacienteEmpresa")
    pLlenarCboRs cboTipoPaciente, rs, 1, 0, 3
    cboTipoPaciente.ListIndex = 0
       
    mskFecIni.Mask = ""
    mskFecIni.Text = dtmfecha
    mskFecIni.Mask = "##/##/####"
    
    mskFecFin.Mask = ""
    mskFecFin.Text = dtmfecha
    mskFecFin.Mask = "##/##/####"
    
    chkFechas.Value = 1
    chkFechas_Click
    chkFolios_Click
    
    optAmbos.Value = True
    optEstado(2).Value = True

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

Private Sub mskFecini_GotFocus()
    On Error GoTo NotificaError
    
    
    pSelMkTexto mskFecIni

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_GotFocus"))
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

Private Sub optAmbos_Click()
    
    cboTipoPaciente.Enabled = True
    optEstado(3).Enabled = True
    
End Sub

Private Sub optCliente_Click()

    cboTipoPaciente.Enabled = Not optCliente.Value
    cboTipoPaciente.ListIndex = 0
    optEstado(3).Enabled = True

End Sub

Private Sub optExternos_Click()
    
    cboTipoPaciente.Enabled = True
    optEstado(3).Enabled = True

End Sub

Private Sub optGrupo_Click()

    cboTipoPaciente.Enabled = True
    optEstado(3).Enabled = True
    
End Sub

Private Sub optInternos_Click()
        
    cboTipoPaciente.Enabled = True
    optEstado(3).Enabled = True

End Sub

Sub pImprime(pstrDestino As String)
    Dim rs As New ADODB.Recordset
    
    Dim alstrParametros(4) As String
    Dim strTipoPaciente As String
    Dim strEstado As String
    Dim strEstadoFiltro As String
    Dim intMostrarCancelacion As Integer
    Dim largo As Integer 'Longitud del folio de factura
    Dim i As Integer, Caracter As String, Cadena As String 'Variables de funcion que extrae letras y numeros del folio de factura
    Dim strNumero As String 'Variable que contiene los numeros de folio inicial
    Dim strNumeroF As String 'Variable que contiene los numeros de folio final
    Dim strLetra As String 'Variable que contiene el identificador del folio inicial
    Dim strLetraF As String 'Variable que contiene el identificador del folio final
    Dim folioinicial, foliofinal As String 'Variables que contienen intfolio
    Dim identificadorInicial As String, IdentificadorFinal As String 'Variables que contienen vchserie
    
    If cboDepartamento.ListIndex = -1 Then
        'Seleccione el departamento.
        MsgBox SIHOMsg(242), vbOKOnly + vbInformation, "Mensaje"
    Else
        strTipoPaciente = IIf(optInternos.Value, "I", strTipoPaciente)
        strTipoPaciente = IIf(optExternos.Value, "E", strTipoPaciente)
        strTipoPaciente = IIf(optCliente.Value, "C", strTipoPaciente)
        strTipoPaciente = IIf(optGrupo.Value, "G", strTipoPaciente)
        strTipoPaciente = IIf(optSocios.Value, "S", strTipoPaciente)
        strTipoPaciente = IIf(optAmbos.Value, "*", strTipoPaciente)
        
        strEstado = IIf(optEstado(0).Value, "C", strEstado)
        strEstado = IIf(optEstado(1).Value, "P", strEstado)
        strEstado = IIf(optEstado(3).Value, "R", strEstado)
        strEstado = IIf(optEstado(2).Value, "*", strEstado)
        
        If strEstado = "*" Or strEstado = "C" Then
            intMostrarCancelacion = 1
        Else
            intMostrarCancelacion = 0
        End If
        
        strEstadoFiltro = IIf(optEstado(0).Value, "<CANCELADAS>", strEstadoFiltro)
        strEstadoFiltro = IIf(optEstado(1).Value, "<PAGADAS>", strEstadoFiltro)
        strEstadoFiltro = IIf(optEstado(3).Value, "<CREDITO>", strEstadoFiltro)
        strEstadoFiltro = IIf(optEstado(2).Value, "<TODAS>", strEstadoFiltro)
        
'Separar el folio de factura en letras y numeros
 '---------------------------------------------------------FolioInicial
    Cadena = Trim(txtDeSerie.Text)
    largo = Len(Cadena)
        For i = 1 To largo
           Caracter = Right(Left(Cadena, i), 1)
             If IsNumeric(Caracter) Then
              strNumero = strNumero & Caracter
             Else
              strLetra = strLetra & Caracter
             End If
        Next
     identificadorInicial = strLetra
     folioinicial = strNumero
 '---------------------------------------------------------FolioFinal
    Cadena = Trim(txtASerie)
    largo = Len(Cadena)
        For i = 1 To largo
          Caracter = Right(Left(Cadena, i), 1)
            If IsNumeric(Caracter) Then
             strNumeroF = strNumeroF & Caracter
            Else
             strLetraF = strLetraF & Caracter
            End If
        Next
    IdentificadorFinal = strLetraF
    foliofinal = strNumeroF
  '--------------------------------------------------------Parametros_SP
vgstrParametrosSP = _
        CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & _
        "|" & strTipoPaciente & _
        "|" & CStr(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)) & _
        "|" & strEstado & _
        "|" & IIf(chkFechas.Value = 1, 1, 0) & _
        "|" & fstrFechaSQL(mskFecIni.Text) & _
        "|" & fstrFechaSQL(mskFecFin.Text) & _
        "|" & IIf(chkFolios.Value = 1, 1, 0) & _
        "|" & identificadorInicial & _
        "|" & folioinicial & _
        "|" & IdentificadorFinal & _
        "|" & foliofinal & _
        "|" & Str(cboHospital.ItemData(cboHospital.ListIndex)) & _
        "|" & IIf(Me.chkFoliosFiscales.Value = vbChecked, 1, 0)
        

        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptRelacionFactura")
        If rs.RecordCount <> 0 Then
            vgrptReporte.DiscardSavedData
            alstrParametros(0) = "NombreHospital;" & cboHospital.List(cboHospital.ListIndex)
            alstrParametros(1) = "Filtro;" & IIf(chkFechas.Value = 1, "DEL " & UCase(Format(mskFecIni.Text, "dd/mmm/yyyy")) & " AL " & UCase(Format(mskFecFin.Text, "dd/mmm/yyyy")), "DE LA " & txtDeSerie.Text & " A LA " & txtASerie.Text)
            alstrParametros(2) = "Departamento;" & cboDepartamento.List(cboDepartamento.ListIndex)
            alstrParametros(3) = "Estado;" & strEstadoFiltro
            alstrParametros(4) = "MostarCancelacion;" & intMostrarCancelacion
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rs, pstrDestino, "Relación de facturas"
        Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rs.Close
    End If
End Sub

Private Sub optSocios_Click()
    
    cboTipoPaciente.Enabled = False
    optEstado(3).Enabled = False

End Sub

Private Sub txtASerie_GotFocus()
    
    pSelTextBox txtASerie
    
End Sub

Private Sub txtASerie_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtDeSerie_GotFocus()
    
    pSelTextBox txtDeSerie

End Sub

Private Sub txtDeSerie_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

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

