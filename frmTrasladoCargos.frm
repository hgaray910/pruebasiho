VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrasladoCargos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traslado de cargos"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   4320
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCaptura 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   3870
      TabIndex        =   49
      Top             =   7700
      Visible         =   0   'False
      Width           =   4400
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   100
         TabIndex        =   55
         Top             =   435
         Width           =   4200
         Begin VB.TextBox txtCapturaDato 
            Height          =   345
            Left            =   1770
            TabIndex        =   48
            Top             =   300
            Width           =   2300
         End
         Begin VB.Label lblTituloAxa 
            Caption         =   "Número de control"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1665
         End
      End
      Begin VB.Frame fraGuardar 
         Height          =   720
         Left            =   1878
         TabIndex        =   54
         Top             =   1320
         Width           =   645
         Begin VB.CommandButton cmdGuardar 
            Height          =   495
            Left            =   75
            Picture         =   "frmTrasladoCargos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Grabar información"
            Top             =   155
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdEsc 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   230
         Left            =   4135
         TabIndex        =   52
         Top             =   60
         Width           =   230
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Información AXA"
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
         Left            =   80
         TabIndex        =   53
         Top             =   60
         Width           =   2775
      End
      Begin VB.Label lblTituloCancelaciones 
         BackColor       =   &H8000000D&
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
         Height          =   315
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   4400
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1290
      Left            =   1710
      TabIndex        =   41
      Top             =   5220
      Visible         =   0   'False
      Width           =   7680
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   360
         Left            =   165
         TabIndex        =   42
         Top             =   675
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         Caption         =   "Trasladando cargos, por favor espere..."
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
         TabIndex        =   43
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
   Begin VB.Frame freCargos 
      Height          =   3465
      Left            =   1013
      TabIndex        =   44
      Top             =   4500
      Width           =   8655
      Begin VB.CommandButton cmdCerrarSeleccion 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   3855
         TabIndex        =   47
         Top             =   3075
         Width           =   1050
      End
      Begin VB.ListBox lstCargos 
         Height          =   2535
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   46
         Top             =   450
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Selección de cargos"
         ForeColor       =   &H80000014&
         Height          =   225
         Left            =   90
         TabIndex        =   45
         Top             =   150
         Width           =   3225
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000002&
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   30
         Top             =   135
         Width           =   8565
      End
   End
   Begin VB.Frame freTrasladar 
      Height          =   1330
      Left            =   3255
      TabIndex        =   40
      Top             =   2880
      Width           =   9120
      Begin VB.OptionButton OptTipoConcepto 
         Caption         =   "No aplicado"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   22
         ToolTipText     =   "Concepto no aplicado"
         Top             =   750
         Width           =   1200
      End
      Begin VB.OptionButton OptTipoConcepto 
         Caption         =   "Aplicado"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   4400
         TabIndex        =   21
         ToolTipText     =   "Concepto aplicado"
         Top             =   750
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox chkSinActualizarPrecio 
         Caption         =   "Trasladar cargos sin actualizar precio ni descuento"
         Height          =   195
         Left            =   165
         TabIndex        =   23
         ToolTipText     =   "Al trasladar no actualizar el precio ni el descuento"
         Top             =   1030
         Width           =   4815
      End
      Begin VB.CheckBox chkCambioConcepto 
         Caption         =   "Cambiar el concepto de facturación de medicamentos"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         ToolTipText     =   "Al trasladar los cargos cambiarlos de concepto de facturación"
         Top             =   780
         Width           =   4150
      End
      Begin VB.CheckBox chkCerrarCuenta 
         Caption         =   "Cerrar cuenta origen (sólo externos)"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         ToolTipText     =   "Cerrar la cuenta del paciente externo para que no se utilice más"
         Top             =   270
         Width           =   2835
      End
      Begin VB.CheckBox chkPorCargo 
         Caption         =   "Trasladar por cargo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   165
         TabIndex        =   19
         ToolTipText     =   "Seleccionar de la cuenta del paciente, los cargos a trasladar"
         Top             =   530
         Width           =   1890
      End
      Begin VB.CommandButton cmdTraslado 
         Caption         =   "Trasladar cuenta"
         Enabled         =   0   'False
         Height          =   450
         Left            =   6720
         TabIndex        =   24
         ToolTipText     =   "Comenzar el proceso de traslado de cargos"
         Top             =   770
         Width           =   2250
      End
   End
   Begin VB.Frame freParametros 
      Caption         =   "Asignar a la cuenta destino"
      Height          =   1330
      Left            =   75
      TabIndex        =   39
      Top             =   2880
      Width           =   3105
      Begin VB.CheckBox chkPagos 
         Caption         =   "Pagos"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   1030
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkSolExamenes 
         Caption         =   "Solicitudes de exámenes"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   780
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkSolEstudios 
         Caption         =   "Solicitudes de estudios"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   530
         Value           =   1  'Checked
         Width           =   2070
      End
      Begin VB.CheckBox chkRequisiciones 
         Caption         =   "Requisiciones"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   270
         Value           =   1  'Checked
         Width           =   1395
      End
   End
   Begin VB.Frame FrePaciente2 
      Caption         =   "Cuenta destino"
      Enabled         =   0   'False
      Height          =   1365
      Left            =   75
      TabIndex        =   33
      Top             =   1480
      Width           =   12300
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1575
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Empresa"
         Top             =   930
         Width           =   10570
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   1575
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Tipo de paciente"
         Top             =   585
         Width           =   7260
      End
      Begin VB.TextBox txtCuarto2 
         Height          =   285
         Left            =   10150
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Cuarto ocupado"
         Top             =   585
         Width           =   1980
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Interno"
         Height          =   255
         Index           =   2
         Left            =   2850
         TabIndex        =   8
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Externo"
         Height          =   255
         Index           =   3
         Left            =   3825
         TabIndex        =   9
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtPaciente2 
         Height          =   285
         Left            =   6275
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Nombre del paciente"
         Top             =   270
         Width           =   5860
      End
      Begin VB.TextBox txtMov2 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Introduzca la cuenta"
         Top             =   270
         Width           =   1095
      End
      Begin VB.ComboBox cboEmpleado 
         Height          =   315
         Left            =   1575
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Empleado"
         Top             =   930
         Width           =   10270
      End
      Begin VB.ComboBox cboMedico 
         Height          =   315
         Left            =   1575
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         ToolTipText     =   "Médico"
         Top             =   930
         Width           =   10270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuarto"
         Height          =   195
         Left            =   9080
         TabIndex        =   38
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número de cuenta"
         Height          =   195
         Left            =   165
         TabIndex        =   37
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   5380
         TabIndex        =   36
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lblRelacion 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   165
         TabIndex        =   35
         Top             =   975
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   165
         TabIndex        =   34
         Top             =   630
         Width           =   1200
      End
   End
   Begin VB.Frame FrePaciente 
      Caption         =   "Cuenta origen"
      Height          =   1365
      Left            =   75
      TabIndex        =   26
      Top             =   80
      Width           =   12300
      Begin VB.TextBox txtCuarto 
         Height          =   285
         Left            =   10150
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Cuarto ocupado"
         Top             =   585
         Width           =   1980
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Externo"
         Height          =   255
         Index           =   1
         Left            =   3825
         TabIndex        =   2
         Top             =   315
         Width           =   975
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Interno"
         Height          =   255
         Index           =   0
         Left            =   2850
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtPaciente 
         Height          =   285
         Left            =   6275
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Nombre del paciente"
         Top             =   270
         Width           =   5860
      End
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Introduzca cuenta del paciente"
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txtTipoPaciente 
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Tipo de paciente"
         Top             =   585
         Width           =   7260
      End
      Begin VB.TextBox txtEmpleadoPaciente 
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Empleado"
         Top             =   900
         Width           =   10560
      End
      Begin VB.TextBox txtMedicoPaciente 
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   58
         ToolTipText     =   "Medico"
         Top             =   900
         Width           =   10560
      End
      Begin VB.TextBox txtEmpresaPaciente 
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "Empresa"
         Top             =   900
         Width           =   10560
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cuarto"
         Height          =   195
         Left            =   9080
         TabIndex        =   32
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de cuenta"
         Height          =   195
         Left            =   165
         TabIndex        =   31
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   5380
         TabIndex        =   30
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lblRelacionOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   165
         TabIndex        =   29
         Top             =   945
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   165
         TabIndex        =   28
         Top             =   630
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmTrasladoCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Modifiqué para eliminar el uso de la variable <gblnCatCentralizados>
' Se eliminó el procedimiento: <pLlenaColEmpresas> y las colecciones <colEmpresasAfiliacion>, <objAfiliacion>

'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmTrasladoCargos                                      -
'-------------------------------------------------------------------------------------
'| Objetivo: Realizar el traslado de cargos de una cuenta a otra
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 28/Ene/2001
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : 29/Ene/2001
'| Fecha última modificación: 29/Ene/2001
'-------------------------------------------------------------------------------------

Option Explicit

Const clngOpcion = 363

'- Claves de la tabla SITIPOINGRESO -'
Const clngCveInternamientoNormal = 1
Const clngCveAmbulatorio = 2
Const clngCvePrepago = 3
Const clngCveInternoFueUrgencias = 4
Const clngCveInternoFueAmbulatorio = 5
Const clngCveRecienNacido = 6
Const clngCveUrgencias = 7
Const clngCveExterno = 8
Const clngCvePrevio = 11
'------------------------------------'

Dim vgstrEstadoManto As String
Dim vgintEmpresa As Integer
Dim vgintTipoPaciente As Integer
Dim vgintEmpresa2 As Integer
Dim vgintTipoPaciente2 As Integer
Dim vlstrCargos As String
Dim vlrsSeleccionaCargos As New ADODB.Recordset

'-- Agregados para caso 7673 --'
Dim vllngNumPaciente As Long                'Variable que indica el número de paciente
Dim vlngCveTipoIngreso As Long              'Variable que indica la clave del tipo de ingreso, claves de la tabla SITIPOINGRESO
'------------------------------'

'---------------- INTERFAZ WS -------------------'
Public vglngCveInterfazWS As Long           'Variable que indica la clave de la interfaz según la empresa convenio
Public vglngCveTipoIngresoAXA As Long       'Variable que indica el tipo de ingreso del paciente para la captura de datos de AXA
Public vgstrContratoAXA As String           'Variable que indica la clave del contrato para AXA, configurado en el catálogo de equivalencias
Public vgstrControlAXA As String            'Variable que indica el número de control para AXA
Public vgstrNumCuartoAXA As String          'Variable que indica el número de cuarto para la interfaz AXA
Public vgstrAutorizaGralAXA As String       'Variable que indica el número de autorización general para la interfaz AXA
Public vgstrAutorizaEspecialAXA As String   'Variable que indica el número de autorización especial para la interfaz AXA
Public vgstrMedicoTratanteAXA As String     'Variable que indica el nombre del médico tratante para la interfaz AXA (INTERNOS)
Public vgstrMedicoEmergenciasAXA As String  'Variable que indica el nombre del médico para emergencias para la interfaz AXA (URGENCIAS)
Public vglngPersonaGrabaAXA As Long         'Variable que indica la clave de la persona que realiza la transacción para la interfaz AXA
'------------------------------------------------'

Dim vlblnNoClickTipoPaciente As Boolean
Dim vlstrNumCuentaDestino As String
Dim vlblnNoLimpiaNumCuentaDestino As Boolean

Dim vllngCveEmpresaOrigen As Long

Dim vllngCveEmpresaDestino As Long
Dim vlblnBanderaGenera As Boolean 'Indica si maneja el genera paciente externo y asi no activa las casillas de seleccion de "Asignar cuenta destino"
Dim vblnBanderaNoAplicado As Boolean 'Indicara si es la cuenta externa a la que solo le pasaran medicamentos no aplicados
Dim vlblnBanderaExterno As Boolean 'Indica si se genero un externo para validar al buscar nueva cuenta si es la misma del interno-externo
Dim vlstrvchvalor As Integer

Private Sub pLlenaCargos()
On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim vlintcontador As Integer
    Dim vlstrDepartamento As String
    
    '---------------------'
    ' Barrita de progreso '
    '---------------------'
    freBarra.Top = 800
    pgbBarra.Value = 0
    lblTextoBarra.Caption = "Consultando cargos, por favor espere..."
    freBarra.Visible = True
    freBarra.Refresh
    
    Set vlrsSeleccionaCargos = frsEjecuta_SP(CLng(txtMovimientoPaciente.Text) & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|2|-1|C|N|0", "SP_PVSELCARGOSPACIENTE")
    vlrsSeleccionaCargos.Sort = "NombreDepartamento,dtmFechaHora,DescripcionCargo"
    
    lstCargos.Visible = False
    lstCargos.Clear
    With vlrsSeleccionaCargos
        vlstrDepartamento = ""
        Do While Not .EOF
            If vlstrDepartamento <> Trim(!nombreDepartamento) Then
                lstCargos.AddItem "---------- " & !nombreDepartamento & " ----------"
                lstCargos.ItemData(lstCargos.newIndex) = -1
                vlstrDepartamento = Trim(!nombreDepartamento)
            End If
            lstCargos.AddItem "(" & !chrTipoCargo & ") - " & Format(!dtmFechahora, "dd/mm/yyyy") & " -- " & Trim(!DescripcionCargo) & " (" & Trim(str(!MNYCantidad)) & ")"
            lstCargos.ItemData(lstCargos.newIndex) = !IntNumCargo
            pgbBarra.Value = .Bookmark / .RecordCount * 100
            .MoveNext
        Loop
        .Close
    End With
    If vlstrvchvalor = 1 And vblnBanderaNoAplicado = True Then
        seleccionarMedicamentos
    End If
    
    freBarra.Visible = False
    lstCargos.Visible = True
    freCargos.Top = 210
    freCargos.Left = 1900
    freCargos.Visible = True
    FrePaciente2.Enabled = False
    freTrasladar.Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCargos"))
    Unload Me
End Sub

Private Sub pLlenaCombos()
On Error GoTo NotificaError
    
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset
      
    vlstrSentencia = "select tnyCveTipoPaciente, vchDescripcion from adTipoPaciente where tnyCveTipoPaciente <> " & flngTipoPacienteSocio & _
                         " AND BITACTIVO = 1 "
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboTipoPaciente, rs, 0, 1, 0, False
        cboTipoPaciente.ListIndex = -1
        rs.Close
    End If
    
    vlstrSentencia = "select intCveEmpresa, vchDescripcion from ccEmpresa where BITACTIVO = 1"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboEmpresa, rs, 0, 1, 0, False
        cboEmpresa.ListIndex = -1
    End If
    rs.Close
    
    'Empleados
    vlstrSentencia = "Select intCveEmpleado, " & _
                             "Trim(vchApellidoPaterno)||' '||Trim(vchApellidoMaterno)||' '||Trim(vchNombre) Nombre " & _
                             "From NoEmpleado " & _
                             "Where bitActivo = 1 "
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboEmpleado, rs, 0, 1, 0, False
        cboEmpleado.ListIndex = -1
    End If
    rs.Close
        
    'Médicos
    Set rs = frsEjecuta_SP("-1|1", "SP_EXSelMedico")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboMedico, rs, 0, 1, 0, False
        cboMedico.ListIndex = -1
    End If
    rs.Close

    cboTipoPaciente.Enabled = False
    
    cboEmpresa.Visible = True
    cboEmpleado.Visible = False
    cboMedico.Visible = False
    lblRelacion.Caption = "Empresa"
    
    cboEmpresa.Enabled = False
    cboEmpleado.Enabled = False
    cboMedico.Enabled = False
    
    txtEmpresaPaciente.Visible = True
    txtEmpleadoPaciente.Visible = False
    txtMedicoPaciente.Visible = False
    lblRelacionOrigen.Caption = "Empresa"
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenaCombos"))
    Unload Me
End Sub

Private Sub cboEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cmdTraslado.Enabled Then
            cmdTraslado.SetFocus
        Else
            txtMov2.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_KeyDown"))
    Unload Me
End Sub

Private Sub cboMedico_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cmdTraslado.Enabled Then
            cmdTraslado.SetFocus
        Else
            txtMov2.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboMedico_KeyDown"))
    Unload Me
End Sub

Private Sub CboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cmdTraslado.Enabled Then
            cmdTraslado.SetFocus
        Else
            txtMov2.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpresa_KeyDown"))
    Unload Me
End Sub

Private Sub cboTipoPaciente_Click()
On Error GoTo NotificaError
    
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset
    
    If cboTipoPaciente.ListIndex > -1 Then  'no hay nada seleccionado
        
        vlstrSentencia = "Select chrTipo From AdTipoPaciente Where tnyCveTipoPaciente = " & Trim(str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)))
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        
        Select Case rs!chrTipo
        
        Case "CO"
        'Convenio
            cboEmpresa.Visible = True
            cboEmpleado.Visible = False
            cboMedico.Visible = False
        
            cboEmpresa.Enabled = True
            cboEmpleado.Enabled = False
            cboMedico.Enabled = False
            
            lblRelacion.Caption = "Empresa"
            
            cboEmpresa.ListIndex = -1
            cboEmpleado.ListIndex = -1
            cboMedico.ListIndex = -1
        
        Case "EM"
        'Empleado o familiar de empleado
            cboEmpresa.Visible = False
            cboEmpleado.Visible = True
            cboMedico.Visible = False
        
            cboEmpresa.Enabled = False
            cboEmpleado.Enabled = True
            cboMedico.Enabled = False
            
            lblRelacion.Caption = "Empleado"
            
            cboEmpresa.ListIndex = -1
            cboEmpleado.ListIndex = -1
            cboMedico.ListIndex = -1
        
        Case "ME"
        'Médico o familiar de médico
            cboEmpresa.Visible = False
            cboEmpleado.Visible = False
            cboMedico.Visible = True
        
            cboEmpresa.Enabled = False
            cboEmpleado.Enabled = False
            cboMedico.Enabled = True
            
            lblRelacion.Caption = "Médico"
            
            cboEmpresa.ListIndex = -1
            cboEmpleado.ListIndex = -1
            cboMedico.ListIndex = -1
        
        Case Else
        
            cboEmpresa.Visible = True
            cboEmpleado.Visible = False
            cboMedico.Visible = False
            
            cboEmpresa.Enabled = False
            cboEmpleado.Enabled = False
            cboMedico.Enabled = False
            
            lblRelacion.Caption = "Empresa"

            cboEmpresa.ListIndex = -1
            cboEmpleado.ListIndex = -1
            cboMedico.ListIndex = -1
            
        End Select
        rs.Close
        
    Else

        cboEmpresa.Visible = True
        cboEmpleado.Visible = False
        cboMedico.Visible = False

        cboEmpresa.Enabled = False
        cboEmpleado.Enabled = False
        cboMedico.Enabled = False

        lblRelacion.Caption = "Empresa"

        cboEmpresa.ListIndex = -1
        cboEmpleado.ListIndex = -1
        cboMedico.ListIndex = -1

    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoPaciente_Click"))
    Unload Me
End Sub

Private Sub cboTipoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cboEmpresa.Enabled And cboEmpresa.Visible Then
            cboEmpresa.SetFocus
        ElseIf cboEmpleado.Enabled And cboEmpleado.Visible Then
                cboEmpleado.SetFocus
        ElseIf cboMedico.Enabled And cboMedico.Visible Then
            cboMedico.SetFocus
        Else
            If cmdTraslado.Enabled Then
                cmdTraslado.SetFocus
            Else
                txtMov2.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub chkCambioConcepto_Click()
    If chkCambioConcepto.Value Then
        OptTipoConcepto(0).Enabled = True
        OptTipoConcepto(1).Enabled = True
    Else
        OptTipoConcepto(0).Enabled = False
        OptTipoConcepto(1).Enabled = False
    End If
    
End Sub

Private Sub chkCambioConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If chkCambioConcepto.Value = 1 Then
            OptTipoConcepto(0).SetFocus
        Else
            chkSinActualizarPrecio.SetFocus
        End If
    End If
End Sub

Private Sub chkCerrarCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If chkPorCargo.Enabled = True Then
            chkPorCargo.SetFocus
        Else
            chkCambioConcepto.SetFocus
        End If
    End If
End Sub

Private Sub chkPorCargo_Click()
On Error GoTo NotificaError
    
    If chkPorCargo.Value = 1 Then
        pLlenaCargos
        txtMov2.Enabled = False
    Else
        If vlstrvchvalor <> 0 Then
            cmdCerrarSeleccion_Click
            lstCargos.Clear
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkPorCargo_Click"))
    Unload Me
End Sub

Private Sub chkPorCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkCambioConcepto.SetFocus
    End If
End Sub

Private Sub cmdCerrarSeleccion_Click()
On Error GoTo NotificaError
    
    FrePaciente2.Enabled = True
    freTrasladar.Enabled = True
    freCargos.Visible = False
    txtMov2.Enabled = True
    If fblnCanFocus(chkPorCargo) Then chkPorCargo.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCerrarSeleccion_Click"))
    Unload Me
End Sub

Private Sub pVerificarListaPrecio()
    Dim vlaryParametros() As String
    Dim vldblPrecio As Double '<-- Se cambió tipo de dato para caso 7365
    Dim vllngCargos As Long
    Dim vllngContador As Long
    Dim vllngNumCargo As Long
    Dim vllngCveEmpresa As Long
    'Caso 19900
    Dim rsAuditoriaCargos As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlblnAuditoriacargos As Boolean
    Dim vlblnNoCambiaPrecio As Boolean

    'Caso 19900 -- verifica si esta activo o no el parámetro de auditoría de cargos
    vlstrSentencia = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITAUDITORIADECARGOS'"
    Set rsAuditoriaCargos = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With rsAuditoriaCargos
        If .RecordCount > 0 Then
            vlblnAuditoriacargos = IIf(IsNull(!VCHVALOR), False, IIf(!VCHVALOR = 0, False, True))
        End If
        .Close
    End With
    
    
    vlstrCargos = ""
    
    If chkPorCargo.Value <> 1 Then pLlenaCargos
    freCargos.Visible = False
    freTrasladar.Enabled = True
    FrePaciente2.Enabled = True
    
    vldblPrecio = 0
    
    vllngCargos = lstCargos.ListCount
    vlrsSeleccionaCargos.Open
    For vllngContador = 0 To vllngCargos - 1
       vlrsSeleccionaCargos.MoveFirst
       
       vldblPrecio = 0
       vllngNumCargo = lstCargos.ItemData(vllngContador)
       
       'Caso 19900 -- Si está activo el parámetro de Auditoría de cargos verifica si el cargo está en la lista
       vlblnNoCambiaPrecio = False
       If vlblnAuditoriacargos = True Then
            vlstrSentencia = "select count(*) total from PVPRECIOSAUDITORIA "
            vlstrSentencia = vlstrSentencia & " where CHRTIPOCARGO = (select chrtipocargo from pvcargo where intnumcargo = " & vllngNumCargo & ")"
            vlstrSentencia = vlstrSentencia & " AND CHRCVECARGO = (select chrcvecargo from pvcargo where intnumcargo = " & vllngNumCargo & ")"
            Set rsAuditoriaCargos = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            With rsAuditoriaCargos
                If !Total > 0 Then
                    vlblnNoCambiaPrecio = True
                End If
                .Close
            End With
       End If
       
       If vlblnNoCambiaPrecio = False Then 'Caso 19900, si el cargo no está en la lista hace el cambio de precio
            'Si es la entrada en la lista del departamento ignorar
            If Mid(lstCargos.List(vllngContador), 1, 11) <> "---------- " Then
                'Si se escogió por cargos checar si esta seleccionado
                 If (chkPorCargo.Value = 1 And lstCargos.Selected(vllngContador)) Or chkPorCargo.Value = 0 Then
                 If cboEmpresa.ListIndex = -1 Or Not cboEmpresa.Enabled Or Not cboEmpresa.Visible Then
                     vllngCveEmpresa = 0
                 Else
                     vllngCveEmpresa = cboEmpresa.ItemData(cboEmpresa.ListIndex)
                 End If
                 'En caso de no existir si revisa todo pero si existe en 0 entonces no revisa todo y regresa 0
                 'una validacion donde cheque que el cargo exista y tenga precio en la lista de precios a la cual se va a trasladar
                 vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) _
                                     & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                     & "|" & Trim(txtMov2.Text) _
                                     & "|" & IIf(OptTipoPaciente(2).Value, "I", "E") _
                                     & "|" & cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) _
                                     & "|" & CStr(vllngCveEmpresa) _
                                     & "|" & Trim(str(vllngNumCargo)) _
                                     & "|" & "1" _
                                     & "|" & IIf(OptTipoPaciente(1).Value And chkCerrarCuenta.Value = 1, "1", "0") _
                                     & "|" & IIf(chkRequisiciones.Value = 1, "1", "0") _
                                     & "|" & IIf(chkSolEstudios.Value = 1, "1", "0") _
                                     & "|" & IIf(chkSolExamenes.Value = 1, "1", "0") _
                                     & "|" & IIf(chkCambioConcepto.Value = 1, "1", "0") _
                                     & "|" & CStr(vgintNumeroDepartamento)
                 pCargaArreglo vlaryParametros, "|" & vbDouble
                 frsEjecuta_SP vgstrParametrosSP, "SP_PVSELCHKTRASLADO", True, , vlaryParametros
                 pObtieneValores vlaryParametros, vldblPrecio
                 If vldblPrecio = 0 Then
                     'vlstrCargos = vlstrCargos & Chr(13) & lstCargos.List(vllngContador)
                     Do While vlrsSeleccionaCargos!IntNumCargo <> lstCargos.ItemData(vllngContador)
                         vlrsSeleccionaCargos.MoveNext
                     Loop
                     If InStr(1, vlstrCargos, vlrsSeleccionaCargos!DescripcionCargo) = 0 Then
                         vlstrCargos = vlstrCargos & Chr(13) & vlrsSeleccionaCargos!DescripcionCargo
                     End If
                 End If
             End If
            End If
       End If 'Caso 19900
    Next
    If chkPorCargo.Value <> 1 Then lstCargos.Clear
    vlrsSeleccionaCargos.Close
End Sub

Private Function fblnValidaEmpresaSiEsConvenio() As Boolean
 Dim vlstrSentencia As String
 Dim rs As New ADODB.Recordset
 
 fblnValidaEmpresaSiEsConvenio = False
 vlstrSentencia = "Select chrTipo From AdTipoPaciente Where tnyCveTipoPaciente = " & Trim(str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)))
 Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
 If rs!chrTipo = "CO" Then
    If cboEmpresa.ListCount = 0 Or cboEmpresa.ListIndex = -1 Then
       fblnValidaEmpresaSiEsConvenio = True
    End If
 End If
 rs.Close
 
End Function

Private Function fblnValidaEmpleado() As Boolean
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset
 
 fblnValidaEmpleado = False
 
 vlstrSentencia = "Select chrTipo From adTipoPaciente where tnyCveTipoPaciente = " & Trim(str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)))
 Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
 If rs!chrTipo = "EM" Then
    If cboEmpleado.ListCount = 0 Or cboEmpleado.ListIndex = -1 Then
       fblnValidaEmpleado = True
    End If
 End If
 rs.Close
 
End Function

Private Function fblnValidaMedico() As Boolean
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset
 
 fblnValidaMedico = False
 
 vlstrSentencia = "Select chrTipo From adTipoPaciente where tnyCveTipoPaciente = " & Trim(str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)))
 Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
 If rs!chrTipo = "ME" Then
    If cboMedico.ListCount = 0 Or cboMedico.ListIndex = -1 Then
       fblnValidaMedico = True
    End If
 End If
 rs.Close
 
End Function

'- CASO 7673: Interfaz AXA -'
Private Sub cmdEsc_Click()
On Error GoTo NotificaError

    FraCaptura.Visible = False
    FraCaptura.Top = 7700
    cmdTraslado.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEsc_Click"))
End Sub

'- CASO 7673: Interfaz AXA -'
Private Sub cmdGuardar_Click()
On Error GoTo NotificaError

    Dim lstrSentencia As String

    If Trim(txtCapturaDato.Text) = "" Then
        MsgBox SIHOMsg(2), vbInformation + vbOKOnly, "Mensaje"
        txtCapturaDato.SetFocus
        Exit Sub
    End If

    '- Guardar el número de control para la interfaz de AXA -'
    lstrSentencia = "UPDATE ExPacienteIngreso SET vchNumAfiliacion = '" & Trim(txtCapturaDato.Text) & "'" & _
                    " WHERE intNumPaciente = " & CStr(vllngNumPaciente) & " AND intNumCuenta = " & Trim(txtMov2.Text)
    pEjecutaSentencia lstrSentencia
    vgstrControlAXA = Trim(txtCapturaDato.Text)
    MsgBox SIHOMsg(284), vbInformation + vbOKOnly, "Mensaje"
    
    FraCaptura.Top = 7700
    FraCaptura.Visible = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGuardar_Click"))
End Sub

Private Sub cmdTraslado_Click()
On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim vllngCargos As Long
    Dim vllngContador As Long
    Dim vllngNumCargo As Long
    Dim vlbolSeleccionado As Boolean
    Dim vllngCveEmpresa As Long
    Dim vllngPersonaGraba As Long
    Dim SQL As String
    Dim vlaryParametros() As String
    Dim vllngPrecio As Long
    Dim vlmsgConceptosSeguro As String
    Dim rsConceptosSeguro As New ADODB.Recordset
    Dim vlintConceptoFac As Integer 'Concepto de factura indica 1:Aplicados , 0:No aplicados
    Dim rsParametro As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim vlngNuevaCuenta As Long
    Dim vblnBandera As Boolean
    Dim rsUltimaCuenta As New ADODB.Recordset
    Dim vlngNumPaciente As Long
    'Caso 19900
    Dim rsAuditoriaCargos As New ADODB.Recordset
    Dim vlblnAuditoriacargos As Boolean
    
    If fblnValidaEmpresaSiEsConvenio Then
       '¡No se puede realizar el traslado, se debe elegir una empresa para el tipo de paciente seleccionado!
        MsgBox SIHOMsg(1182) & vlstrCargos, vbExclamation, "Mensaje"
        cboEmpresa.SetFocus
        Exit Sub
    End If
    
    If fblnValidaEmpleado Then
       '¡No se puede realizar el traslado, se debe elegir un empleado para el tipo de paciente seleccionado!
        MsgBox SIHOMsg(1429) & vlstrCargos, vbExclamation, "Mensaje"
        cboEmpleado.SetFocus
        Exit Sub
    End If
    
    If fblnValidaMedico Then
       '¡No se puede realizar el traslado, se debe elegir un médico para el tipo de paciente seleccionado!
        MsgBox SIHOMsg(1504) & vlstrCargos, vbExclamation, "Mensaje"
        cboMedico.SetFocus
        Exit Sub
    End If
    
    '- CASO 7673: Se valida si la empresa seleccionada está configurada para usarse con alguna interfaz de WS -'
    vglngCveInterfazWS = 0
    If cboEmpresa.ListIndex > -1 And cboEmpresa.Enabled And cboEmpresa.Visible Then
        'Se obtiene la clave de la interfaz a utilizar
        vglngCveInterfazWS = 1
        frsEjecuta_SP cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & vgintClaveEmpresaContable, "FN_GNSELINTERFAZWS", True, vglngCveInterfazWS
        If vglngCveInterfazWS <> 0 Then
            vglngCveInterfazWS = IIf(fblnLicenciaWS(vglngCveInterfazWS) = True, vglngCveInterfazWS, 0)
        End If
    End If
    
    '----------------------------------------------------------------------------------------------------------'
    If vblnBanderaNoAplicado = True And cmdTraslado.Enabled = True And vlstrvchvalor = 1 Then
        If chkPorCargo = 0 Then
            '¡No se puede realizar el traslado, se debe elegir un medicamento!
            MsgBox "¡No se puede realizar el traslado, se debe elegir un medicamento!", vbInformation, "Mensaje"
            Exit Sub
        End If
        If chkCambioConcepto = 0 Then
            '¡No se puede realizar el traslado, se debe cambiar el concepto de facturación!
            MsgBox "¡No se puede realizar el traslado, se debe cambiar el concepto de facturación!", vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
    
        '----------------------------------------------------------------------------------------------------------'
    'Caso 20262 -- verifica si esta activo o no el parámetro de auditoría de cargos para mostrar el mensaje de precio de 1 peso
    vlstrSentencia = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITAUDITORIADECARGOS'"
    Set rsAuditoriaCargos = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With rsAuditoriaCargos
        If .RecordCount > 0 Then
            vlblnAuditoriacargos = IIf(IsNull(!VCHVALOR), False, IIf(!VCHVALOR = 0, False, True))
        Else
            vlblnAuditoriacargos = False
        End If
        .Close
    End With
    
    If vlblnAuditoriacargos = True Then
        'Caso 19900 - Mensaje para no permitir traslado de cargos a cuentas con cargo con precio de un peso
        vlstrSentencia = "select count(*) total from pvcargo where mnyprecio = 1 and intmovpaciente = " & txtMovimientoPaciente.Text & " and chrtipopaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'"
        Set rsAuditoriaCargos = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        With rsAuditoriaCargos
            If !Total > 0 Then
                '¡No se puede realizar el traslado, existen cargos con precio de un peso!
                MsgBox "¡No se puede realizar el traslado, existen cargos con precio de un peso!", vbInformation, "Mensaje"
                .Close
                Exit Sub
            End If
        End With
    End If
    '----------------------------------------------------------------------------------------------------------'
    vlstrCargos = ""
    If chkSinActualizarPrecio.Value = 0 Then
        pVerificarListaPrecio
    End If
    If vlstrCargos <> "" Then
        'No se pudieron trasladar los siguientes cargos, no se encontró precio asignado.
        MsgBox SIHOMsg(1139) & vlstrCargos, vbInformation, "Mensaje"
        Exit Sub
    End If
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    '************************************** AGREGADO PARA INTERFAZ AXA (CASO 7673) **************************************'
    'Se valida si el paciente tiene licencia para el ingreso por WS
    If vglngCveInterfazWS > 0 And ((vlngCveTipoIngreso = clngCveUrgencias) Or (vlngCveTipoIngreso = clngCveInternoFueUrgencias) Or (vlngCveTipoIngreso = clngCveInternamientoNormal)) Then ' Or (vlngCveTipoIngreso = clngCvePrevio) Then
        Dim rsLogInterfaz As ADODB.Recordset
        'Se valida si el paciente ya tiene un registro previo en las transacciones con la interfaz del WS
        vgstrParametrosSP = CStr(vllngNumPaciente) & "|" & Trim(txtMov2.Text) & "|" & CStr(vlngCveTipoIngreso)
        Set rsLogInterfaz = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELLOGINTERFAZAXA")
        If rsLogInterfaz.RecordCount = 0 Then
            pAsignaValoresVariables CLng(txtMov2.Text), vllngPersonaGraba 'Se asignan los valores a las variables de la interfaz de AXA
            
            If Not fblnDatosValidos Then 'Revisar que los datos principales de la interfaz estén capturados
                If FraCaptura.Visible Then txtCapturaDato.SetFocus
                Exit Sub
            End If
            
            frmDatosWSAXA.vgblnTrasladoCargos = True 'Indicar a la forma de datos AXA desde donde se cargará la información
            frmDatosWSAXA.Show vbModal, Me

            vgstrParametrosSP = "''" & "|" & vgintNumeroModulo & "|" & vglngCveTipoIngresoAXA & "|" & Trim(txtMov2.Text) & "|" & vllngNumPaciente & "|" & IIf(OptTipoPaciente(2).Value, "I", "E") & "|||" & IIf(frmDatosWSAXA.vgblnConexionCorrecta = False, "NO", "SI") & "||||" & vllngPersonaGraba & "|" & frmDatosWSAXA.vglngFolioTrans & "|0|"
            frsEjecuta_SP vgstrParametrosSP, "Sp_GnInsLogInterfazAxa"
            
            'Se valida si se realizó una conexión exitosa
            If frmDatosWSAXA.vgblnConexionCorrecta = False Then
                'Si no se realizó una conexión exitosa no permitir realizar el traslado
                MsgBox "No se realizó una conexión con el servicio web de AXA. " & vbNewLine & "No es posible realizar el traslado de cargos a la cuenta " & Trim(txtMov2.Text) & ".", vbInformation, "Mensaje"
                pCancelar 2
                pEnfocaTextBox txtMov2
                Exit Sub
            End If
        End If
    End If
    '********************************************************************************************************************'
    
    If cboEmpresa.ListIndex = -1 Or Not cboEmpresa.Enabled Or Not cboEmpresa.Visible Then
        vllngCveEmpresaDestino = 0
    Else
        vllngCveEmpresaDestino = cboEmpresa.ItemData(cboEmpresa.ListIndex)
    End If
    
    If Trim(txtMovimientoPaciente.Text) = Trim(txtMov2.Text) And _
        OptTipoPaciente(0).Value = OptTipoPaciente(2).Value And _
        OptTipoPaciente(1).Value = OptTipoPaciente(3).Value And _
        vllngCveEmpresaOrigen <> vllngCveEmpresaDestino Then
        vlmsgConceptosSeguro = ""
        vlstrSentencia = "SELECT chrfoliofacturacoaseguro F_Coaseguro " & _
                            ",chrfoliofacturacoaseguroadici F_CoaseguroAdicional " & _
                            ",chrfoliofacturacoaseguromed F_CoaseguroMedico " & _
                            ",chrfoliofacturacopago F_Copago " & _
                            ",chrfoliofacturadeducible F_Deducible " & _
                            ",chrfoliofacturaexcedente F_Excedente " & _
                            ",chrfoliorecibocoaseguro R_Coaseguro " & _
                            ",chrfoliorecibocoaseguroadicion R_CoaseguroAdicional " & _
                            ",chrfoliorecibocopago R_Copago " & _
                            ",chrfoliorecibodeducible R_Deducible " & _
                        " From PVCONTROLASEGURADORA " & _
                        " Where intmovpaciente = " & Trim(txtMovimientoPaciente.Text) & _
                            " AND chrtipopaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'" & _
                            " AND intcveempresa = " & vllngCveEmpresaOrigen
        Set rsConceptosSeguro = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsConceptosSeguro.RecordCount <> 0 Then
            If Trim(rsConceptosSeguro!R_Deducible) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Deducible. Recibo" & "  " & Trim(rsConceptosSeguro!R_Deducible), vlmsgConceptosSeguro & Chr(13) & "Deducible. Recibo" & "  " & Trim(rsConceptosSeguro!R_Deducible))
            End If
            If Trim(rsConceptosSeguro!R_Coaseguro) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Coaseguro. Recibo" & "  " & Trim(rsConceptosSeguro!R_Coaseguro), vlmsgConceptosSeguro & Chr(13) & "Coaseguro. Recibo" & "  " & Trim(rsConceptosSeguro!R_Coaseguro))
            End If
            If Trim(rsConceptosSeguro!R_CoaseguroAdicional) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Coaseguro adicional. Recibo" & "  " & Trim(rsConceptosSeguro!R_CoaseguroAdicional), vlmsgConceptosSeguro & Chr(13) & "Coaseguro adicional. Recibo" & "  " & Trim(rsConceptosSeguro!R_CoaseguroAdicional))
            End If
            If Trim(rsConceptosSeguro!R_Copago) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Copago. Recibo" & "  " & Trim(rsConceptosSeguro!R_Copago), vlmsgConceptosSeguro & Chr(13) & "Copago. Recibo" & "  " & Trim(rsConceptosSeguro!R_Copago))
            End If
            If Trim(rsConceptosSeguro!F_Excedente) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Excedente en suma asegurada. Factura" & "  " & Trim(rsConceptosSeguro!F_Excedente), vlmsgConceptosSeguro & Chr(13) & "Excedente en suma asegurada. Factura" & "  " & Trim(rsConceptosSeguro!F_Excedente))
            End If
            If Trim(rsConceptosSeguro!F_Deducible) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Deducible. Factura" & "  " & Trim(rsConceptosSeguro!F_Deducible), vlmsgConceptosSeguro & Chr(13) & "Deducible. Factura" & "  " & Trim(rsConceptosSeguro!F_Deducible))
            End If
            If Trim(rsConceptosSeguro!F_Coaseguro) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Coaseguro. Factura" & "  " & Trim(rsConceptosSeguro!F_Coaseguro), vlmsgConceptosSeguro & Chr(13) & "Coaseguro. Factura" & "  " & Trim(rsConceptosSeguro!F_Coaseguro))
            End If
            If Trim(rsConceptosSeguro!F_CoaseguroAdicional) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Coaseguro adicional. Factura" & "  " & Trim(rsConceptosSeguro!F_CoaseguroAdicional), vlmsgConceptosSeguro & Chr(13) & "Coaseguro adicional. Factura" & "  " & Trim(rsConceptosSeguro!F_CoaseguroAdicional))
            End If
            If Trim(rsConceptosSeguro!F_Copago) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Copago. Factura" & "  " & Trim(rsConceptosSeguro!F_Copago), vlmsgConceptosSeguro & Chr(13) & "Copago. Factura" & "  " & Trim(rsConceptosSeguro!F_Copago))
            End If
            If Trim(rsConceptosSeguro!F_CoaseguroMedico) <> "" Then
                vlmsgConceptosSeguro = IIf(Trim(vlmsgConceptosSeguro) = "", "Coaseguro médico. Factura" & "  " & Trim(rsConceptosSeguro!F_CoaseguroMedico), vlmsgConceptosSeguro & Chr(13) & "Coaseguro médico. Factura" & "  " & Trim(rsConceptosSeguro!F_CoaseguroMedico))
            End If
        End If
        
        If Trim(vlmsgConceptosSeguro) <> "" Then
            If MsgBox(SIHOMsg(1254) & Chr(13) & Chr(13) & vlmsgConceptosSeguro, vbYesNo + vbQuestion, "Mensaje") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    If chkPorCargo.Value = 1 Then
        'Por si esta vacia la lista
        If lstCargos.ListCount = 0 Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            MsgBox SIHOMsg(420), vbInformation, "Mensaje"
            pCancelar 2
            pCancelar 1
            chkPorCargo.Value = 0
            pEnfocaTextBox txtMovimientoPaciente
            Exit Sub
        End If
        'Barrita de progreso
        freBarra.Top = 800
        pgbBarra.Value = 0
        lblTextoBarra.Caption = "Trasladando cargos, por favor espere..."
        freBarra.Visible = True
        freBarra.Refresh
        vllngCargos = lstCargos.ListCount
    Else
        vllngCargos = 1
    End If
    
    cmdTraslado.Caption = "Procesando..."
    For vllngContador = 0 To vllngCargos - 1
        vllngPrecio = 0
        If chkPorCargo.Value = 1 Then
            vlbolSeleccionado = lstCargos.Selected(vllngContador)
        Else
            vlbolSeleccionado = True
        End If
        
        If vlbolSeleccionado Then
            If chkPorCargo.Value = 1 Then
                vllngNumCargo = lstCargos.ItemData(vllngContador)
            Else
                vllngNumCargo = 0
            End If
            
            If cboEmpresa.ListIndex = -1 Or Not cboEmpresa.Enabled Or Not cboEmpresa.Visible Then
                vllngCveEmpresa = 0
            Else
                vllngCveEmpresa = cboEmpresa.ItemData(cboEmpresa.ListIndex)
            End If
            
            If chkCambioConcepto.Value = 1 Then
                vlintConceptoFac = IIf(OptTipoConcepto(0).Value, "1", "0")
            Else
                vlintConceptoFac = -1
            End If
            
            'En caso de no existir si revisa todo pero si existe en 0 entonces no revisa todo y regresa 0
            'una validacion donde cheque que el cargo exista y tenga precio en la lista de precios a la cual se va a trasladar
            vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) _
                                & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                & "|" & Trim(txtMov2.Text) _
                                & "|" & IIf(OptTipoPaciente(2).Value, "I", "E") _
                                & "|" & cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) _
                                & "|" & CStr(vllngCveEmpresa) _
                                & "|" & Trim(str(vllngNumCargo)) _
                                & "|" & IIf(chkSinActualizarPrecio.Value = 1, "2", "1") _
                                & "|" & IIf(OptTipoPaciente(1).Value And chkCerrarCuenta.Value = 1, "1", "0") _
                                & "|" & IIf(chkRequisiciones.Value = 1, "1", "0") _
                                & "|" & IIf(chkSolEstudios.Value = 1, "1", "0") _
                                & "|" & IIf(chkSolExamenes.Value = 1, "1", "0") _
                                & "|" & IIf(chkCambioConcepto.Value = 1, "1", "0") _
                                & "|" & CStr(vllngPersonaGraba) _
                                & "|" & CStr(vgintNumeroDepartamento) _
                                & "|" & vlintConceptoFac
            frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdTrasladoCargos", True
                                    
             pEjecutaSentencia "update pvcargo set CHRFOLIOFACTURA=null, INTCVECARTA =NULL WHERE INTMOVPACIENTE=" & Trim(txtMov2.Text)
                                    
                                    
            If cboEmpleado.ListIndex = -1 Or Not cboEmpleado.Enabled Or Not cboEmpleado.Visible Then
                pEjecutaSentencia "UPDATE EXPACIENTEINGRESO SET INTCVEEMPLEADORELACIONADO = NULL WHERE INTNUMCUENTA = " & Trim(txtMov2.Text) & " AND INTCVETIPOINGRESO IN (SELECT INTCVETIPOINGRESO FROM SITIPOINGRESO WHERE CHRTIPOINGRESO = '" & IIf(OptTipoPaciente(2).Value, "I", "E") & "') "
            Else
                pEjecutaSentencia "UPDATE EXPACIENTEINGRESO SET INTCVEEMPLEADORELACIONADO = " & cboEmpleado.ItemData(cboEmpleado.ListIndex) & " WHERE INTNUMCUENTA = " & Trim(txtMov2.Text) & " AND INTCVETIPOINGRESO IN (SELECT INTCVETIPOINGRESO FROM SITIPOINGRESO WHERE CHRTIPOINGRESO = '" & IIf(OptTipoPaciente(2).Value, "I", "E") & "') "
            End If
            
            If cboMedico.ListIndex = -1 Or Not cboMedico.Enabled Or Not cboMedico.Visible Then
                pEjecutaSentencia "UPDATE EXPACIENTEINGRESO SET INTCVEMedicoRELACIONADO = NULL WHERE INTNUMCUENTA = " & Trim(txtMov2.Text) & " AND INTCVETIPOINGRESO IN (SELECT INTCVETIPOINGRESO FROM SITIPOINGRESO WHERE CHRTIPOINGRESO = '" & IIf(OptTipoPaciente(2).Value, "I", "E") & "') "
            Else
                pEjecutaSentencia "UPDATE EXPACIENTEINGRESO SET INTCVEMedicoRELACIONADO = " & cboMedico.ItemData(cboMedico.ListIndex) & " WHERE INTNUMCUENTA = " & Trim(txtMov2.Text) & " AND INTCVETIPOINGRESO IN (SELECT INTCVETIPOINGRESO FROM SITIPOINGRESO WHERE CHRTIPOINGRESO = '" & IIf(OptTipoPaciente(2).Value, "I", "E") & "') "
            End If
        End If
        
        
        
        
        If freBarra.Visible Then
            pgbBarra.Value = (vllngContador / vllngCargos) * 100
        End If
    Next
    
    If chkPagos.Value = 1 Then
        vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & Trim(txtMov2.Text) & "|" & IIf(OptTipoPaciente(2).Value, "I", "E")
        frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDTRASLADAPAGOS"
    End If
    
    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "TRASLADO DE CARGOS", Trim(txtMovimientoPaciente.Text) & "-->" & Trim(txtMov2.Text))
    
    cmdTraslado.Caption = "Trasladar cuenta"
    freBarra.Visible = False
    
    SQL = "Delete From PvTipoPacienteProceso Where PvTipoPacienteProceso.intNumeroLogin = " & vglngNumeroLogin & _
          " And PvTipoPacienteProceso.intProceso = " & enmTipoProceso.TrasladoCargos
    pEjecutaSentencia SQL
    
    SQL = "Insert Into PvTipoPacienteProceso (intNumeroLogin, intProceso, chrTipoPaciente) Values(" & vglngNumeroLogin & "," & enmTipoProceso.TrasladoCargos & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ")"
    pEjecutaSentencia SQL
        
    EntornoSIHO.ConeccionSIHO.CommitTrans
      
    'La operación se realizó satisfactoriamente.
    MsgBox SIHOMsg(420), vbInformation, "Mensaje"
        
    Unload Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTraslado_Click"))
    Unload Me
End Sub

Private Sub Form_Activate()
    vgstrNombreForm = Me.Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then
        If FraCaptura.Visible Then
            FraCaptura.Top = 7700
            FraCaptura.Visible = False
        Else
            Unload Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim rsParametro As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    Me.Icon = frmMenuPrincipal.Icon
    
   pLlenaCombos
    
    Set rs = frsRegresaRs("Select bitTrasladaCargos From PvParametro Where tnyClaveEmpresa = " & vgintClaveEmpresaContable)
    chkSinActualizarPrecio.Value = IIf(rs!bitTrasladaCargos = 1, 1, 0)
    rs.Close
    
    
    chkCambioConcepto.Enabled = fblnRevisaPermiso(vglngNumeroLogin, 363, "C")
    
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.TrasladoCargos) > 0 Then
        If fintEsInterno(vglngNumeroLogin, enmTipoProceso.TrasladoCargos) = 1 Then
          OptTipoPaciente(0).Value = True
        Else
          OptTipoPaciente(1).Value = True
        End If
    End If

    FraCaptura.Top = 7700
    FraCaptura.Visible = False 'Agregado para interfaz AXA (CASO 7673)
    vlblnBanderaGenera = False
    vlstrSentencia = "select vchvalor from siparametro where vchnombre = 'BITGENERAPACIENTEEXTERNO' AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable & " AND CHRMODULO = 'PV' "
    Set rsParametro = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsParametro.RecordCount <> 0 Then vlstrvchvalor = IIf(IsNull(rsParametro!VCHVALOR), 0, rsParametro!VCHVALOR)
    rsParametro.Close
    vblnBanderaNoAplicado = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError
    
    If Not FrePaciente.Enabled Then
        pCancelar 2
        pCancelar 1
        chkPorCargo.Value = 0
        chkSinActualizarPrecio.Value = 0
        chkCambioConcepto.Value = 0
        OptTipoConcepto(0).Value = True
        pEnfocaTextBox txtMovimientoPaciente
        Cancel = 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub lstCargos_Click()
On Error GoTo NotificaError
    
    If lstCargos.ItemData(lstCargos.ListIndex) = -1 Then
        lstCargos.Selected(lstCargos.ListIndex) = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstCargos_Click"))
    Unload Me
End Sub
Private Sub lstCargos_ItemCheck(Item As Integer)
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset
Dim vlstrNumCargo As String

    If vlstrvchvalor = 1 And vblnBanderaNoAplicado = True Then
        If Left(lstCargos.List(Item), 4) <> "(AR)" Then
            lstCargos.Selected(lstCargos.ListIndex) = False
        Else
            If lstCargos.Selected(lstCargos.ListIndex) = True Then
                vlstrNumCargo = lstCargos.ItemData(Item)
                vlstrSentencia = "select intnumcargo from pvcargo inner join ivarticulo on ivarticulo.INTIDARTICULO = pvcargo.CHRCVECARGO " & _
                " where pvcargo.INTNUMCARGO = " & vlstrNumCargo & " and ivarticulo.CHRCVEARTMEDICAMEN = 1"
                Set rs = frsRegresaRs(vlstrSentencia)
                If rs.RecordCount > 0 Then
                    If rs!IntNumCargo <> "" Then
                        lstCargos.Selected(lstCargos.ListIndex) = True
                    Else
                        lstCargos.Selected(lstCargos.ListIndex) = False
                    End If
                Else
                    lstCargos.Selected(lstCargos.ListIndex) = False
                End If
                rs.Close
            End If
        End If
    End If
End Sub

Private Sub OptTipoConcepto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkSinActualizarPrecio.SetFocus
    End If
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
On Error GoTo NotificaError
    
    If Index = 0 Or Index = 1 Then
        pEnfocaTextBox txtMovimientoPaciente
    Else
        vlblnNoLimpiaNumCuentaDestino = True
        If Not vlblnNoClickTipoPaciente Then pCancelar 2
        vlblnNoLimpiaNumCuentaDestino = False
        pEnfocaTextBox txtMov2
    End If
    
    If OptTipoPaciente(1).Value And OptTipoPaciente(2).Value Then
        chkCambioConcepto.Value = 1
    Else
        chkCambioConcepto.Value = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPaciente_Click"))
    Unload Me
End Sub

Private Sub txtCapturaDato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtMov2_GotFocus()
    vlstrNumCuentaDestino = txtMov2.Text
End Sub

Private Sub txtMov2_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            OptTipoPaciente(2).Value = UCase(Chr(KeyAscii)) = "I"
            OptTipoPaciente(3).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMov2_KeyPress"))
    Unload Me
End Sub

Private Sub txtMov2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If vlstrvchvalor = 1 And OptTipoPaciente(0).Value = True Then
            Dim rsUltimaCuenta As New ADODB.Recordset
            Dim rsParametro As New ADODB.Recordset
            Dim vlstrSentencia As String
            Dim vlngNumPaciente As Long
            
            vlstrSentencia = "select INTNUMPACIENTE from expacienteingreso where INTNUMCUENTA = " & Trim(txtMovimientoPaciente.Text)
            Set rsParametro = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            If rsParametro.RecordCount <> 0 Then vlngNumPaciente = IIf(IsNull(rsParametro!intnumpaciente), 0, rsParametro!intnumpaciente)
            rsParametro.Close
            
            Set rsUltimaCuenta = frsEjecuta_SP(CStr(Trim(vlngNumPaciente)) & "|" & vgintClaveEmpresaContable & "|" & vgintNumeroDepartamento, "sp_GnSelUltimaCuentaPaciente")
            If Not rsUltimaCuenta.EOF Then
                If (rsUltimaCuenta!Estatus = "A" Or rsUltimaCuenta!Estatus = "P") And (cgstrModulo = "AD" Or rsUltimaCuenta!tipo = "E") Then
                    If Trim(rsUltimaCuenta!cuenta) = Trim(txtMov2.Text) Then
                        vblnBanderaNoAplicado = True
                    Else
                        vblnBanderaNoAplicado = False
                    End If
                End If
            End If
        End If
        pDatosPaciente IIf(OptTipoPaciente(2).Value, "I", "E")
        
'        If Trim(txtMovimientoPaciente.Text) = Trim(txtMov2.Text) And _
'        ( _
'        (OptTipoPaciente(0).Value And OptTipoPaciente(2).Value) Or _
'        (OptTipoPaciente(1).Value And OptTipoPaciente(3).Value) _
'        ) Then
'            chkRequisiciones.Value = 0
'            chkSolEstudios.Value = 0
'            chkSolExamenes.Value = 0
'            chkPagos.Value = 0
'        Else
'            chkRequisiciones.Value = 1
'            chkSolEstudios.Value = 1
'            chkSolExamenes.Value = 1
'            chkPagos.Value = 1
'        End If
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMov2_KeyDown"))
    Unload Me
End Sub

Private Sub pDatosPaciente(strTipoPaciente As String)
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If RTrim(txtMov2.Text) = "" Then
        With FrmBusquedaPacientes
            If OptTipoPaciente(3).Value Then 'Externos
                .vgstrTipoPaciente = "E"
                .Caption = .Caption & " externos"
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSoloActivos.Enabled = True
                .optSinFacturar.Enabled = True
                .optTodos.Enabled = False
                .optSinFacturar.Value = True
                .vgStrOtrosCampos = ", SiTipoIngreso.vchNombre as Tipo, ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                .vgstrTamanoCampo = "800,3400,1500,1750,4100"
            ElseIf OptTipoPaciente(2).Value Then 'Internos
                .vgstrTipoPaciente = "I"
                .vgblnPideClave = False
                .Caption = .Caption & " internos"
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSinFacturar.Value = True
                .optSinFacturar.Enabled = True
                .optSoloActivos.Enabled = True
                .optTodos.Enabled = False
                .vgStrOtrosCampos = ", SiTipoIngreso.vchNombre as Tipo,  ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."",  ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                .vgstrTamanoCampo = "800,3400,2200,1050,1050,4100"
            End If
            
            txtMov2.Text = .flngRegresaPaciente()
            
            If txtMov2 <> -1 Then
                vlstrNumCuentaDestino = txtMov2.Text
                txtMov2_KeyDown vbKeyReturn, 0
            Else
                txtMov2.Text = ""
                vlstrNumCuentaDestino = ""
            End If
        End With
    Else
        If fblnCuentaValida(CDbl(txtMov2.Text), strTipoPaciente, False, rs) Then
            vgstrEstadoManto = "C" 'Cargando
            txtPaciente2.Text = rs!Nombre
                        
            cboTipoPaciente.ListIndex = fintLocalizaCbo(cboTipoPaciente, rs!cveTipoPaciente)
            cboTipoPaciente.Enabled = True
            
            If cboEmpresa.Enabled And cboEmpresa.Visible Then
                cboEmpresa.ListIndex = fintLocalizaCbo(cboEmpresa, rs!cveEmpresa)
            End If
            
            If cboEmpleado.Enabled And cboEmpleado.Visible Then
                If Trim(rs!TipoPaciente) = "EM" Then
                    If Not IsNull(rs!CveExtra) Then
                        cboEmpleado.ListIndex = fintLocalizaCbo(cboEmpleado, rs!CveExtra)
                    End If
                End If
            End If
            
            If cboMedico.Enabled And cboMedico.Visible Then
                If Trim(rs!TipoPaciente) = "ME" Then
                    If Not IsNull(rs!CveExtra) Then
                        cboMedico.ListIndex = fintLocalizaCbo(cboMedico, rs!CveExtra)
                    End If
                End If
            End If
            
            cmdTraslado.Enabled = True
            
            cboTipoPaciente.ListIndex = fintLocalizaCbo(cboTipoPaciente, rs!cveTipoPaciente)
            
            If cboEmpresa.Enabled And cboEmpresa.Visible Then
                cboEmpresa.ListIndex = fintLocalizaCbo(cboEmpresa, rs!cveEmpresa)
            End If
            
            If cboEmpleado.Enabled And cboEmpleado.Visible Then
                If Trim(rs!TipoPaciente) = "EM" Then
                    If Not IsNull(rs!CveExtra) Then
                        cboEmpleado.ListIndex = fintLocalizaCbo(cboEmpleado, rs!CveExtra)
                    End If
                End If
            End If
            
            If cboMedico.Enabled And cboMedico.Visible Then
                If Trim(rs!TipoPaciente) = "ME" Then
                    If Not IsNull(rs!CveExtra) Then
                        cboMedico.ListIndex = fintLocalizaCbo(cboMedico, rs!CveExtra)
                    End If
                End If
            End If
            
            txtCuarto2 = IIf(IsNull(rs!Cuarto), 0, rs!Cuarto)
            cboTipoPaciente.SetFocus
            
            '- Agregado para interfaz AXA (CASO 7673) -'
            vllngNumPaciente = rs!CvePaciente 'Número del paciente
            vlngCveTipoIngreso = flngTipoIngreso(CLng(txtMov2.Text), vllngNumPaciente) 'Clave del tipo de ingreso del paciente
            '------------------------------------------'
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDatosPaciente"))
    Unload Me
End Sub

Private Sub txtMov2_KeyUp(KeyCode As Integer, Shift As Integer)
If txtMov2.Text <> vlstrNumCuentaDestino And vlstrNumCuentaDestino <> "" Then 'hay un cambio se debe de limpiar todo de todo
           vlstrNumCuentaDestino = ""
           vlblnNoLimpiaNumCuentaDestino = True
           pCancelar 2
           vlblnNoLimpiaNumCuentaDestino = False
  End If
End Sub

Private Sub txtMov2_LostFocus()
On Error GoTo NotificaError
   
    If Trim(txtMovimientoPaciente.Text) = Trim(txtMov2.Text) And _
    ( _
    (OptTipoPaciente(0).Value And OptTipoPaciente(2).Value) Or _
    (OptTipoPaciente(1).Value And OptTipoPaciente(3).Value) _
    ) Or vlblnBanderaGenera = True Then
        chkRequisiciones.Value = 0
        chkSolEstudios.Value = 0
        chkSolExamenes.Value = 0
        chkPagos.Value = 0
    Else
        chkRequisiciones.Value = 1
        chkSolEstudios.Value = 1
        chkSolExamenes.Value = 1
        chkPagos.Value = 1
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMov2_LostFocus"))
    Unload Me
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsParametro As New ADODB.Recordset
    Dim rsUltimaCuenta As New ADODB.Recordset
    Dim vlngNumPaciente As Long
    Dim vlngNuevaCuenta As Long
    Dim vlngCuentaDestino As Long
    Dim vintValueA As Integer
    Dim vintValueB As Integer
    vblnBanderaNoAplicado = False
    
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
                    .optTodos.Enabled = False
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", SiTipoIngreso.vchNombre as Tipo, TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1500,1750,4100"
                Else
                    .vgstrTipoPaciente = "I"  'Internos
                    .vgblnPideClave = False
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = False
                    .vgStrOtrosCampos = ", SiTipoIngreso.vchNombre as Tipo, TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha ing."", TO_CHAR(ExPacienteIngreso.dtmFechaHoraEgreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,2200,1050,1050,4100"
                End If
                
                txtMovimientoPaciente.Text = .flngRegresaPaciente()
                
                If txtMovimientoPaciente <> -1 Then
                    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                Else
                    txtMovimientoPaciente.Text = ""
                End If
            End With
        Else
            If fblnCuentaValida(CDbl(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E"), True, rs) Then
                'Cargamos datos de cuenta origen
                vlblnBanderaExterno = False
                vllngCveEmpresaOrigen = rs!cveEmpresa
                vgstrEstadoManto = "C" 'Cargando
                FrePaciente.Enabled = False
                FrePaciente2.Enabled = True
                txtPaciente.Text = rs!Nombre
               
                txtTipoPaciente.Text = rs!tipo
                If cboEmpresa.Enabled Then
                    cboEmpresa.ListIndex = fintLocalizaCbo(cboEmpresa, rs!cveEmpresa)
                End If
                If cboEmpresa.Enabled And cboEmpresa.Visible And cboEmpresa.ListIndex <> -1 Then
                    txtEmpresaPaciente.Text = IIf(IsNull(rs!empresa), "", rs!empresa)
                Else
                    If cboEmpleado.Enabled And cboEmpleado.Visible And cboEmpleado.ListIndex <> -1 Then
                        txtEmpleadoPaciente.Visible = True
                        txtMedicoPaciente.Visible = False
                        txtEmpresaPaciente.Visible = False
                        lblRelacionOrigen.Caption = "Empleado"
                        txtEmpleadoPaciente.Text = Trim(cboEmpleado.Text)
                    ElseIf cboMedico.Enabled And cboMedico.Visible And cboMedico.ListIndex <> -1 Then
                        txtMedicoPaciente.Visible = True
                        txtEmpleadoPaciente.Visible = False
                        txtEmpresaPaciente.Visible = False
                        lblRelacionOrigen.Caption = "Médico"
                        txtMedicoPaciente.Text = Trim(cboMedico.Text)
                    End If
                End If
                txtCuarto = IIf(IsNull(rs!Cuarto), 0, rs!Cuarto)
                
                vlstrSentencia = "select INTNUMPACIENTE from expacienteingreso where INTNUMCUENTA = " & Trim(txtMovimientoPaciente.Text)
                Set rsParametro = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rsParametro.RecordCount <> 0 Then vlngNumPaciente = IIf(IsNull(rsParametro!intnumpaciente), 0, rsParametro!intnumpaciente)
                rsParametro.Close
                    
                If vlstrvchvalor = 1 And OptTipoPaciente(0).Value = True Then
                    vlngNuevaCuenta = 0
                    vlblnBanderaGenera = False 'Variable para cuando se mierda el focus del txtmov2 no active los check list de "Asignar a la cuenta destino"
                    'Verificamos si el parametro esta activo para trasladar los cargos a una cuenta externa HFM
                    vlngNuevaCuenta = 0
                    Set rsUltimaCuenta = frsEjecuta_SP(CStr(Trim(vlngNumPaciente)) & "|" & vgintClaveEmpresaContable & "|" & vgintNumeroDepartamento, "sp_GnSelUltimaCuentaPaciente")
                    If Not rsUltimaCuenta.EOF Then
                        If (rsUltimaCuenta!Estatus = "A" Or rsUltimaCuenta!Estatus = "P") And (cgstrModulo = "AD" Or rsUltimaCuenta!tipo = "E") Then
                            If MsgBox("¿Desea trasladar el medicamento a la cuenta de externo?", vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                                vlngCuentaDestino = rsUltimaCuenta!cuenta
                                vintValueA = 0
                                vintValueB = 1
                                vlblnBanderaGenera = True 'Variable para cuando se mierda el focus del txtmov2 no active los check list de "Asignar a la cuenta destino"
                                vblnBanderaNoAplicado = True
                            Else
                                vlngCuentaDestino = txtMovimientoPaciente.Text
                                vintValueA = OptTipoPaciente(0).Value
                                vintValueB = OptTipoPaciente(1).Value
                                vblnBanderaNoAplicado = False
                            End If
                        Else
                            If MsgBox("¿Desea abrir una cuenta de externo para el medicamento?", vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                                vlngNuevaCuenta = flngcrearExterno(Trim(txtMovimientoPaciente.Text))
                                vlblnBanderaGenera = True 'Variable para cuando se mierda el focus del txtmov2 no active los check list de "Asignar a la cuenta destino"
                                vblnBanderaNoAplicado = True
                            Else
                                vlngCuentaDestino = txtMovimientoPaciente.Text
                                vintValueA = OptTipoPaciente(0).Value
                                vintValueB = OptTipoPaciente(1).Value
                                vblnBanderaNoAplicado = False
                            End If
                            If vlngNuevaCuenta <> 0 Then
                                vlngCuentaDestino = vlngNuevaCuenta
                                vintValueA = 0
                                vintValueB = 1
                            End If
                        End If
                    End If
                Else
                    vlngCuentaDestino = txtMovimientoPaciente.Text
                    vintValueA = OptTipoPaciente(0).Value
                    vintValueB = OptTipoPaciente(1).Value
                    vblnBanderaNoAplicado = False
                End If
                                                            
                    txtMov2.Text = vlngCuentaDestino
                    vlstrNumCuentaDestino = vlngCuentaDestino
                    vlblnNoClickTipoPaciente = True
                    OptTipoPaciente(2).Value = vintValueA
                    OptTipoPaciente(3).Value = vintValueB
                    vlblnNoClickTipoPaciente = False
                    
                    txtPaciente2.Text = rs!Nombre
                    cboTipoPaciente.ListIndex = fintLocalizaCbo(cboTipoPaciente, rs!cveTipoPaciente)
                    cboTipoPaciente.Enabled = True
                    txtCuarto2.Text = txtCuarto.Text
                    txtEmpresaPaciente.Visible = True
                    txtEmpleadoPaciente.Visible = False
                    txtMedicoPaciente.Visible = False
                    lblRelacionOrigen.Caption = "Empresa"
                    
                    If cboEmpresa.Enabled Then
                        cboEmpresa.ListIndex = fintLocalizaCbo(cboEmpresa, rs!cveEmpresa)
                    End If
                    
                    If cboEmpleado.Enabled And cboEmpleado.Visible Then
                        If Trim(rs!TipoPaciente) = "EM" Then
                            If Not IsNull(rs!CveExtra) Then
                                cboEmpleado.ListIndex = fintLocalizaCbo(cboEmpleado, rs!CveExtra)
                            End If
                        End If
                    End If
                    
                    If cboMedico.Enabled And cboMedico.Visible Then
                        If Trim(rs!TipoPaciente) = "ME" Then
                            If Not IsNull(rs!CveExtra) Then
                                cboMedico.ListIndex = fintLocalizaCbo(cboMedico, rs!CveExtra)
                            End If
                        End If
                    End If
                    
                    If cboEmpresa.Enabled And cboEmpresa.Visible And cboEmpresa.ListIndex <> -1 Then
                        txtEmpresaPaciente.Text = IIf(IsNull(rs!empresa), "", rs!empresa)
                    Else
                        If cboEmpleado.Enabled And cboEmpleado.Visible And cboEmpleado.ListIndex <> -1 Then
                            txtEmpleadoPaciente.Visible = True
                            txtMedicoPaciente.Visible = False
                            txtEmpresaPaciente.Visible = False
                            lblRelacionOrigen.Caption = "Empleado"
                            txtEmpleadoPaciente.Text = Trim(cboEmpleado.Text)
                        ElseIf cboMedico.Enabled And cboMedico.Visible And cboMedico.ListIndex <> -1 Then
                            txtMedicoPaciente.Visible = True
                            txtEmpleadoPaciente.Visible = False
                            txtEmpresaPaciente.Visible = False
                            lblRelacionOrigen.Caption = "Médico"
                            txtMedicoPaciente.Text = Trim(cboMedico.Text)
                        End If
                    End If
                                        
                    txtCuarto = IIf(IsNull(rs!Cuarto), 0, rs!Cuarto)
                    txtCuarto2.Text = txtCuarto.Text
                    
                    cmdTraslado.Enabled = True
                    chkPorCargo.Enabled = True
                    txtMov2_KeyDown vbKeyReturn, 0
                    
                    If txtMovimientoPaciente.Text <> txtMov2.Text Then
                        cmdTraslado.SetFocus
                    Else
                        pEnfocaTextBox txtMov2
                    End If
                    If vblnBanderaNoAplicado = True Then
                        vlblnBanderaExterno = True
                        chkCambioConcepto.Value = 1
                        OptTipoConcepto(1).Value = 1
                        chkPorCargo.Value = 1
                        cmdCerrarSeleccion_Click
                        seleccionarMedicamentos
                    End If
            End If
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub pCancelar(vlbytCual As Byte)
On Error GoTo NotificaError
    
    If vlbytCual = 1 Then
        FrePaciente.Enabled = True
        FrePaciente2.Enabled = False
        cmdTraslado.Enabled = False
        chkPorCargo.Enabled = False
        txtMovimientoPaciente.Text = ""
        txtPaciente.Text = ""
        txtTipoPaciente.Text = ""
        txtEmpresaPaciente.Visible = True
        txtEmpleadoPaciente.Visible = False
        txtMedicoPaciente.Visible = False
        txtEmpresaPaciente.Text = ""
        txtEmpleadoPaciente.Text = ""
        txtMedicoPaciente.Text = ""
        lblRelacionOrigen.Caption = "Empresa"
        txtCuarto.Text = ""
    Else
        FrePaciente2.Enabled = True
        cmdTraslado.Enabled = False
        If Not vlblnNoLimpiaNumCuentaDestino Then txtMov2.Text = ""
        txtPaciente2.Text = ""
        txtCuarto2.Text = ""
        
        cboTipoPaciente.ListIndex = -1
        cboEmpresa.ListIndex = -1
        cboEmpleado.ListIndex = -1
        cboMedico.ListIndex = -1
        cboTipoPaciente.Enabled = False
                
        '- Limpiar variables de la cuenta destino -'
        vllngNumPaciente = 0
        vlngCveTipoIngreso = 0
    End If
    
    '- Limpiar variables de la interfaz de AXA -'
    vglngCveInterfazWS = 0
    vglngCveTipoIngresoAXA = 0
    vgstrContratoAXA = ""
    vgstrControlAXA = ""
    vgstrNumCuartoAXA = ""
    vgstrAutorizaGralAXA = ""
    vgstrAutorizaEspecialAXA = ""
    vgstrMedicoTratanteAXA = ""
    vgstrMedicoEmergenciasAXA = ""
    vglngPersonaGrabaAXA = 0
    '-------------------------------------------'

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCancelar"))
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyPress"))
    Unload Me
End Sub

Private Function fblnCuentaValida(lngCuenta As Double, strTipoCuenta As String, blnCuentaOrigen As Boolean, rsDatosPaciente As ADODB.Recordset)
    Dim vlstrSentencia As String
    
On Error GoTo NotificaError
    
    fblnCuentaValida = True
    
    If strTipoCuenta = "I" Then '|  Internos
        vlstrSentencia = "SELECT RTRIM(AdPaciente.vchApellidoPaterno) || ' ' || RTRIM(AdPaciente.vchApellidoMaterno) || ' ' || RTRIM(AdPaciente.vchNombre) AS Nombre, " & _
                         "       ISNULL(AdAdmision.intCveEmpresa, 0) cveEmpresa, " & _
                         "       ISNULL(ccEmpresa.vchDescripcion, '') AS Empresa, " & _
                         "       AdAdmision.tnyCveTipoPaciente cveTipoPaciente, " & _
                         "       AdTipoPaciente.vchDescripcion AS Tipo,  " & _
                         "       AdAdmision.vchNumCuarto Cuarto, " & _
                         "       AdAdmision.bitFacturado Facturada, " & _
                         "       AdAdmision.bitCuentaCerrada Cerrada, " & _
                         "       AdAdmision.numCvePaciente CvePaciente, " & _
                         "       AdAdmision.intCveExtra CveExtra, " & _
                         "       AdTipoPaciente.chrtipo TipoPaciente, " & _
                         "       AdTipoPaciente.BitFamiliar BitFamiliar " & _
                         "FROM   AdAdmision " & _
                         "       INNER JOIN AdPaciente ON (AdAdmision.numCvePaciente = AdPaciente.numCvePaciente) " & _
                         "       INNER JOIN NoDepartamento ON AdAdmision.INTCVEDEPARTAMENTO = NoDepartamento.SMICVEDEPARTAMENTO " & _
                         "       INNER JOIN AdTipoPaciente ON (AdAdmision.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente) " & _
                         "       LEFT OUTER JOIN CcEmpresa ON (AdAdmision.intCveEmpresa = CcEmpresa.intCveEmpresa) " & _
                         "WHERE  AdAdmision.numNumCuenta = " & lngCuenta & " AND NoDepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
    Else '|  Externos
        vlstrSentencia = "SELECT RTRIM(chrApePaterno) || ' ' || RTRIM(chrApeMaterno) || ' ' || RTRIM(chrNombre) AS Nombre, " & _
                         "       ISNULL(RegistroExterno.intClaveEmpresa, 0) cveEmpresa, " & _
                         "       ISNULL(CcEmpresa.vchDescripcion, '') AS Empresa, " & _
                         "       RegistroExterno.tnyCveTipoPaciente AS cveTipoPaciente, " & _
                         "       AdTipoPaciente.vchDescripcion AS Tipo, " & _
                         "       ISNULL(RegistroExterno.vchNumCuarto, '') AS Cuarto, " & _
                         "       RegistroExterno.bitFacturado Facturada, " & _
                         "       RegistroExterno.bitCuentaCerrada Cerrada, " & _
                         "       RegistroExterno.intNumPaciente CvePaciente, " & _
                         "       RegistroExterno.intCveExtra CveExtra, " & _
                         "       AdTipoPaciente.chrtipo TipoPaciente, " & _
                         "       AdTipoPaciente.BitFamiliar BitFamiliar " & _
                         "FROM   RegistroExterno " & _
                         "       INNER JOIN Externo ON (RegistroExterno.intNumPaciente = Externo.intNumPaciente) " & _
                         "       INNER JOIN AdTipoPaciente ON (RegistroExterno.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente) " & _
                         "       INNER JOIN NoDepartamento ON RegistroExterno.INTCVEDEPARTAMENTO = NoDepartamento.SMICVEDEPARTAMENTO " & _
                         "       LEFT OUTER JOIN CcEmpresa ON (RegistroExterno.intClaveEmpresa = CcEmpresa.intCveEmpresa) " & _
                         "WHERE  intNumCuenta = " & lngCuenta & " AND NoDepartamento.tnyClaveEmpresa = " & vgintClaveEmpresaContable
    End If
    Set rsDatosPaciente = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsDatosPaciente.RecordCount <> 0 Then
        If Val(rsDatosPaciente!cveTipoPaciente) = flngTipoPacienteSocio Then
            '|  No se pueden realizar traslados de cargos a pacientes de tipo socio
            MsgBox SIHOMsg(1131), vbCritical, "Mensaje"
            fblnCuentaValida = False
        ElseIf rsDatosPaciente!Cerrada <> 0 Then
            '|  La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
            MsgBox SIHOMsg(596), vbCritical, "Mensaje"
            fblnCuentaValida = False
        
        End If
    Else
        '|  ¡La cuenta no existe!
        MsgBox SIHOMsg(67), vbCritical, "Mensaje"
        fblnCuentaValida = False
    End If
    
    If fblnCuentaValida = False Then
        cmdTraslado.Enabled = False
        pCancelar IIf(blnCuentaOrigen, 1, 2)
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaValida"))
End Function

'- CASO 7673: Busca y asigna los datos del paciente a las variables de la interfaz AXA -'
Private Sub pAsignaValoresVariables(llngCuenta As Long, llngPersonaGraba As Long)
On Error GoTo NotificaError

    Dim rsDatosPaciente As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim lstrSentencia As String
    
    If OptTipoPaciente(2).Value Then  '|  Internos
        lstrSentencia = "SELECT NVL(AdAdmision.intCveEmpresa, 0) cveEmpresa, " & _
                        "       NVL(MedEmer.vchApellidoPaterno, '') || ' ' || NVL(MedEmer.vchApellidoMaterno, '') || ' ' || NVL(MedEmer.vchNombre, '') AS NombreMedicoEmer, " & _
                        "       NVL(MedCargo.vchApellidoPaterno, '') || ' ' || NVL(MedCargo.vchApellidoMaterno, '') || ' ' || NVL(MedCargo.vchNombre, '') AS NombreMedicoCargo, " & _
                        "       AdAdmision.vchNumAfiliacion NumControl, " & _
                        "       AdAdmision.vchAutorizacion Autorizacion " & _
                        " FROM  AdAdmision " & _
                        " INNER JOIN NoDepartamento ON AdAdmision.intCveDepartamento = NoDepartamento.smiCveDepartamento " & _
                        " LEFT  JOIN HoMedico MedEmer ON MedEmer.intCveMedico = AdAdmision.intCveMedicoEmer " & _
                        " LEFT  JOIN HoMedico MedCargo ON MedCargo.intCveMedico = AdAdmision.intCveMedicoCargo " & _
                        " WHERE AdAdmision.numNumCuenta = " & llngCuenta & " AND NoDepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
    ElseIf OptTipoPaciente(3).Value Then '|  Externos
        lstrSentencia = "SELECT NVL(RegistroExterno.intClaveEmpresa, 0) cveEmpresa, " & _
                        "       NVL(HoMedico.vchApellidoPaterno, '') || ' ' || NVL(HoMedico.vchApellidoMaterno, '') || ' ' || NVL(HoMedico.vchNombre, '') AS NombreMedicoEmer, " & _
                        "       ' ' NombreMedicoCargo, " & _
                        "       NVL(RegistroExterno.vchNumAfiliacion, ' ') NumControl, " & _
                        "       NVL(RegistroExterno.vchAutorizacion, ' ') Autorizacion " & _
                        " FROM  RegistroExterno " & _
                        " INNER JOIN NoDepartamento ON RegistroExterno.intCveDepartamento = NoDepartamento.smiCveDepartamento " & _
                        " LEFT  JOIN HoMedico ON HoMedico.intCveMedico = RegistroExterno.intMedico " & _
                        " WHERE RegistroExterno.intNumCuenta = " & llngCuenta & " AND NoDepartamento.tnyClaveEmpresa = " & vgintClaveEmpresaContable
    End If
    Set rsDatosPaciente = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsDatosPaciente.RecordCount <> 0 Then
        ' - Número de contrato AXA - '
        If cboEmpresa.ItemData(cboEmpresa.ListIndex) <> 0 And cboEmpresa.Enabled And cboEmpresa.Visible Then
            Set rs = frsEjecuta_SP(cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & "CONTRATO", "SP_SISELEQUIVALENCIA")
            If rs.RecordCount <> 0 Then
                vgstrContratoAXA = Trim(rs!clave)
            End If
            rs.Close
        End If
        
        ' - Número del cuarto - '
        vgstrNumCuartoAXA = Trim(txtCuarto2.Text)
        
        ' - Número de control (número de nómina AXA) - '
        If Trim(rsDatosPaciente!NumControl) <> "" Then
            vgstrControlAXA = Trim(rsDatosPaciente!NumControl)
        End If
        
        ' - Número de autorización - '
        If Trim(rsDatosPaciente!Autorizacion) <> "" Then
            vgstrAutorizaGralAXA = Trim(rsDatosPaciente!Autorizacion)
        End If
        
        ' - Nombre del médico de emergencia - '
        If Trim(rsDatosPaciente!NombreMedicoEmer) <> "" Then
            vgstrMedicoEmergenciasAXA = Trim(rsDatosPaciente!NombreMedicoEmer)
        End If
    
        ' - Nombre del médico tratante - '
        If Trim(rsDatosPaciente!NombreMedicoCargo) <> "" Then
            vgstrMedicoTratanteAXA = Trim(rsDatosPaciente!NombreMedicoCargo)
        End If
        
        ' - Tipo de ingreso del paciente - '
        vglngCveTipoIngresoAXA = vlngCveTipoIngreso
        
        '- Persona que guarda la transacción -'
        vglngPersonaGrabaAXA = llngPersonaGraba
    End If
    rsDatosPaciente.Close
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsignaValoresVariables"))
    Unload Me
End Sub

'CASO 7673: Valida que los datos para la conexión de la interfaz AXA estén capturados'
Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError
    Dim rsInterfazWS As ADODB.Recordset

    fblnDatosValidos = True
    
    '- Número de proveedor AXA -'
    Set rsInterfazWS = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGINTERFAZWS")
    If fblnDatosValidos And rsInterfazWS!CVEPROVEEDOR = "" Then
        fblnDatosValidos = False
        MsgBox "No se ha configurado la clave de proveedor AXA para el uso de la interfaz con el web service.", vbInformation + vbOKOnly, "Mensaje"
    End If

    '- Número de contrato AXA -'
    If fblnDatosValidos And Trim(vgstrContratoAXA) = "" Then
        fblnDatosValidos = False
        MsgBox "No se ha configurado la clave de contrato para la empresa del paciente.", vbInformation + vbOKOnly, "Mensaje"
    End If

    '- Número de control (número de nómina AXA) -'
    If fblnDatosValidos And Trim(vgstrControlAXA) = "" Then
        fblnDatosValidos = False
        MsgBox "No se ha configurado el número de control del paciente.", vbInformation + vbOKOnly, "Mensaje"
        txtCapturaDato.Text = ""
        FraCaptura.Top = 540
        FraCaptura.Visible = True
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function

'- CASO 7673: Regresa el tipo de ingreso del paciente -'
Private Function flngTipoIngreso(lngCuenta As Long, lngCvePaciente As Long) As Long
    Dim rs As ADODB.Recordset
    
    flngTipoIngreso = 0
    Set rs = frsRegresaRs("SELECT intCveTipoIngreso FROM ExPacienteIngreso WHERE intNumPaciente = " & CStr(vllngNumPaciente) & " AND intNumCuenta = " & CStr(lngCuenta), adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then
        flngTipoIngreso = Trim(rs!intCveTipoIngreso)
    End If
    rs.Close
End Function
Private Function flngcrearExterno(lngCuenta As Long) As Long
    Dim rs As ADODB.Recordset
    Dim vlngNumPaciente As Long
    Dim lngNumPaciente As Long
    Dim VDTMFECHANACIMIENTO As String
    Dim VDTMFECHANACIMIENTOMADRE As String
    Dim VDTMFECHANACIMIENTOPADRE As String
    Dim VDTMFECHANACIMIENTOCONYUGE As String
    Dim VLNGCVEEMPRESA As Long
    Dim VLNGCVETIPOPOLIZA As Long
    Dim vlngCveEmpresaPaciente As Long
    Dim lngnumCuenta As Long
    Dim vlstrSentencia As String
    
    'Traemos el numpaciente para traer sus datos de Expaciente
    Set rs = frsRegresaRs("select INTNUMPACIENTE from EXPACIENTEINGRESO where INTNUMCUENTA = " & lngCuenta, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then
        vlngNumPaciente = Trim(rs!intnumpaciente)
    End If
    rs.Close
    'Creamos un nuevo registro con los mismos datos de  ExPaciente
    Set rs = frsRegresaRs("select * from EXPACIENTE where INTNUMPACIENTE = " & vlngNumPaciente, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then
        Do While Not rs.EOF
        If IsNull(Trim(rs!dtmFechaNacimiento)) Then VDTMFECHANACIMIENTO = "" Else VDTMFECHANACIMIENTO = fstrFechaSQL(rs!dtmFechaNacimiento)
        If IsNull(Trim(rs!dtmFechaNacimientoMadre)) Then VDTMFECHANACIMIENTOMADRE = "" Else VDTMFECHANACIMIENTOMADRE = fstrFechaSQL(rs!dtmFechaNacimientoMadre)
        If IsNull(Trim(rs!dtmFechaNacimientoPadre)) Then VDTMFECHANACIMIENTOPADRE = "" Else VDTMFECHANACIMIENTOPADRE = fstrFechaSQL(rs!dtmFechaNacimientoPadre)
        If IsNull(Trim(rs!dtmFechaNacimientoConyuge)) Then VDTMFECHANACIMIENTOCONYUGE = "" Else VDTMFECHANACIMIENTOCONYUGE = fstrFechaSQL(rs!dtmFechaNacimientoConyuge)
        If IsNull(Trim(rs!intcveempresa)) Then VLNGCVEEMPRESA = 0 Else VLNGCVEEMPRESA = rs!intcveempresa
        If IsNull(Trim(rs!intCveTipoPoliza)) Then VLNGCVETIPOPOLIZA = 0 Else VLNGCVETIPOPOLIZA = rs!intCveTipoPoliza
        If IsNull(Trim(rs!INTCVEEMPRESAPACIENTE)) Then vlngCveEmpresaPaciente = 0 Else vlngCveEmpresaPaciente = rs!INTCVEEMPRESAPACIENTE

        vgstrParametrosSP = rs!intCveTipoPaciente & "|" & VLNGCVEEMPRESA & "|" & _
                            VLNGCVETIPOPOLIZA & "|" & vlngCveEmpresaPaciente & "|" & _
                            rs!intCveReligion & "|" & rs!intCveEstadoCivil & "|" & _
                            rs!INTCVECIUDADNACIMIENTO & "|" & rs!INTCVENACIONALIDAD & "|" & _
                             "|" & rs!vchApellidoPaterno & "|" & _
                            rs!vchApellidoMaterno & "|" & rs!vchNombre & "|" & _
                            rs!chrSexo & "|" & VDTMFECHANACIMIENTO & "|" & _
                            rs!vchRFC & "|" & rs!vchCURP & "|" & _
                            rs!vchOcupacion & "|" & rs!vchCorreoElectronico & "|" & _
                            rs!vchConyugeNombre & "|" & rs!vchConyugeApellidoPaterno & "|" & _
                            rs!vchConyugeApellidoMaterno & "|" & rs!vchNombrePadre & "|" & _
                            rs!vchNombreMadre & "|" & rs!VCHNUMAFILIACION & "|" & _
                            rs!VCHAUTORIZACION & "|" & rs!VCHNUMPOLIZA & "|" & _
                            rs!intPrevio & "|" & vlngNumPaciente & "|" & _
                            rs!intCveOcupacion & "|" & rs!intCveIdioma & "|" & _
                            rs!vchAlergias & "|" & rs!vchFormaNacimiento & "|" & _
                            rs!BITENVIARPROMOCIONES & "|" & VDTMFECHANACIMIENTOPADRE & "|" & _
                            VDTMFECHANACIMIENTOMADRE & "|" & VDTMFECHANACIMIENTOCONYUGE
            rs.MoveNext
        Loop
        lngNumPaciente = 1
        frsEjecuta_SP vgstrParametrosSP, "sp_ExInSPaciente", True, lngNumPaciente
    End If
    rs.Close
    If lngNumPaciente = 0 Then
        'No se pueden guardar los datos
        MsgBox SIHOMsg(33), vbOKOnly + vbExclamation, "Mensaje"
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        Exit Function
    End If

    'Creamos un nuevo registro con los mismos datos de ExPacienteIngreso
    Set rs = frsRegresaRs("select nvl(INTCVETIPOINGRESO,0)INTCVETIPOINGRESO,nvl(INTCVEEMPLEADOINGRESO,0)INTCVEEMPLEADOINGRESO," & _
    "nvl(intCveDepartamento,0)intCveDepartamento,nvl(INTCVEDEPTOINGRESO,0)INTCVEDEPTOINGRESO,nvl(INTCVEPROCEDENCIA,0)INTCVEPROCEDENCIA," & _
    "nvl(INTCVEEMPRESAPACIENTE,0)INTCVEEMPRESAPACIENTE,nvl(intCveTipoPaciente,0)intCveTipoPaciente,nvl(intcveempresa,0)intcveempresa," & _
    "nvl(intCveTipoPoliza,0)intCveTipoPoliza,nvl(INTNUMCUENTAMAMA,0)INTNUMCUENTAMAMA,nvl(intCveMedicoRelacionado,0)intCveMedicoRelacionado," & _
    "nvl(INTCVEEMPLEADORELACIONADO,0)INTCVEEMPLEADORELACIONADO," & _
    "nvl(INTCVEPARENTESCOEMERGENCIA,0)INTCVEPARENTESCOEMERGENCIA,nvl(INTCVEPARENTESCORESPONSABLE,0)INTCVEPARENTESCORESPONSABLE," & _
    "nvl(INTCVEMEDICOEMERGENCIAS,0)INTCVEMEDICOEMERGENCIAS,nvl(INTCVEMEDICOTRATANTE,0)INTCVEMEDICOTRATANTE,nvl(INTCVEDIAGNOSTICOPREVIO,0)INTCVEDIAGNOSTICOPREVIO," & _
    "nvl(INTCVEESTADOSALUD,0)INTCVEESTADOSALUD,nvl(intCvePaquete,0)intCvePaquete,nvl(INTCVECONCEPTOATENCION,0)INTCVECONCEPTOATENCION,nvl(INTCVECUARTO,0)INTCVECUARTO," & _
    "nvl(CHRESTATUS,0)CHRESTATUS,nvl(VCHNUMAFILIACION,'')VCHNUMAFILIACION,nvl(VCHAUTORIZACION,'')VCHAUTORIZACION,nvl(VCHNUMPOLIZA,'')VCHNUMPOLIZA," & _
    "nvl(CHRTIPOATENCION,'')CHRTIPOATENCION,nvl(NUMANTICIPOSUGERIDO,0)NUMANTICIPOSUGERIDO,nvl(VCHNOMBREEMERGENCIA,'')VCHNOMBREEMERGENCIA," & _
    "nvl(VCHDOMICILIOEMERGENCIA,'')VCHDOMICILIOEMERGENCIA,nvl(VCHTELEFONOEMERGENCIA,'')VCHTELEFONOEMERGENCIA,nvl(VCHNOMBRERESPONSABLE,'')VCHNOMBRERESPONSABLE," & _
    "nvl(VCHDOMICILIORESPONSABLE,'')VCHDOMICILIORESPONSABLE,nvl(VCHTELEFONORESPONSABLE,'')VCHTELEFONORESPONSABLE,nvl(VCHLUGARTRABAJORESPONSABLE,'')VCHLUGARTRABAJORESPONSABLE," & _
    "nvl(VCHOBSERVACION,'')VCHOBSERVACION,nvl(INTCUENTAFACTURADA,0)INTCUENTAFACTURADA,nvl(INTCUENTACERRADA,0)INTCUENTACERRADA,nvl(INTCUENTABLOQUEADA,0)INTCUENTABLOQUEADA," & _
    "nvl(INTCUENTAOCUPADA,0)INTCUENTAOCUPADA,nvl(INTORDENINTERNAMIENTO,0)INTORDENINTERNAMIENTO,nvl(VCHDIAGNOSTICOESPECIFICO,'')VCHDIAGNOSTICOESPECIFICO," & _
    "nvl(VCHMOTIVOINGRESO,'')VCHMOTIVOINGRESO,nvl(INTCVEFAMILIAR,0)INTCVEFAMILIAR,nvl(INTCVESOCIO,'')INTCVESOCIO," & _
    "nvl(dtmFechaNacimientoResponsable,'')dtmFechaNacimientoResponsable,nvl(dtmFechaNacimientoEmergencia,'')dtmFechaNacimientoEmergencia,nvl(DTMFECHAHORAINGRESO,'')DTMFECHAHORAINGRESO" & _
    " from EXPACIENTEINGRESO where INTNUMCUENTA = " & lngCuenta, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then
        Do While Not rs.EOF
        
            vgstrParametrosSP = lngNumPaciente & "|" & 8 & "|" & _
                                rs!INTCVEEMPLEADOINGRESO & "|" & vgintNumeroDepartamento & "|" & _
                                0 & "|" & rs!intcveprocedencia & "|" & _
                                rs!INTCVEEMPRESAPACIENTE & "|" & 0 & "|" & _
                                0 & "|" & rs!intCveTipoPaciente & "|" & _
                                rs!intcveempresa & "|" & rs!intCveTipoPoliza & "|" & _
                                rs!intNumCuentaMama & "|" & rs!intCveMedicoRelacionado & "|" & _
                                rs!intCveEmpleadoRelacionado & "|" & rs!intCveParentescoEmergencia & "|" & _
                                rs!intCveParentescoResponsable & "|" & rs!intCveMedicoEmergencias & "|" & _
                                0 & "|" & rs!intCveDiagnosticoPrevio & "|" & _
                                rs!INTCVEESTADOSALUD & "|" & rs!intCvePaquete & "|" & _
                                rs!intCveConceptoAtencion & "|" & 0 & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|"
            vgstrParametrosSP = vgstrParametrosSP & "|" & rs!chrEstatus & "|" & rs!VCHNUMAFILIACION & "|" & _
                                rs!VCHAUTORIZACION & "|" & rs!VCHNUMPOLIZA & "|" & _
                                rs!chrTipoAtencion & "|" & rs!numAnticipoSugerido & "|" & _
                                rs!VCHNOMBREEMERGENCIA & "|" & rs!vchDomicilioEmergencia & "|" & _
                                rs!vchTelefonoEmergencia & "|" & rs!vchNombreResponsable & "|" & _
                                rs!vchDomicilioResponsable & "|" & rs!vchTelefonoResponsable & "|" & _
                                rs!vchLugarTrabajoResponsable & "|" & rs!vchObservacion & "|" & _
                                rs!INTCUENTAFACTURADA & "|" & rs!intcuentacerrada & "|" & _
                                rs!INTCUENTABLOQUEADA & "|" & rs!INTCUENTAOCUPADA & "|" & _
                                0 & "|" & 0 & "|" & _
                                rs!VCHDIAGNOSTICOESPECIFICO & "|" & rs!VCHMOTIVOINGRESO & "|" & _
                                rs!intCveFamiliar & "|" & rs!intcvesocio & "|" & _
                                rs!dtmFechaNacimientoResponsable & "|" & rs!dtmFechaNacimientoEmergencia
            lngnumCuenta = 1
            frsEjecuta_SP vgstrParametrosSP, "sp_ExInSPacienteIngreso", True, lngnumCuenta
            rs.MoveNext
        Loop
    End If
    rs.Close
    If lngnumCuenta = 0 Then
        'No se pueden guardar los datos
        MsgBox SIHOMsg(33), vbOKOnly + vbExclamation, "Mensaje"
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        Exit Function
    End If
    
    MsgBox SIHOMsg(420) & vbNewLine & "El número de cuenta externo relacionado al paciente es: " & lngnumCuenta, vbOKOnly + vbInformation, "Mensaje"
    
    flngcrearExterno = lngnumCuenta
    
End Function

Private Function seleccionarMedicamentos()
Dim vlContador As Integer

    For vlContador = 1 To lstCargos.ListCount
        If vlstrvchvalor = 1 And vblnBanderaNoAplicado = True Then
            If Left(lstCargos.List(vlContador), 4) = "(AR)" Then
                lstCargos.Selected(vlContador) = True
            End If
        End If
    Next
End Function




