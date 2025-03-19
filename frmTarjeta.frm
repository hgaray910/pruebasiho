VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTarjeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarjeta de descuento"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTObj 
      Height          =   7785
      Left            =   -45
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   -360
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   13732
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTarjeta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTitular"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDependiente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraImpresion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraBotonera"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtClavePaciente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmTarjeta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraBusqueda"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtClavePaciente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1935
         MaxLength       =   8
         TabIndex        =   0
         ToolTipText     =   "Número de expediente del paciente"
         Top             =   930
         Width           =   1320
      End
      Begin VB.Frame fraBotonera 
         Height          =   720
         Left            =   3922
         TabIndex        =   55
         Top             =   6840
         Width           =   1140
         Begin VB.CommandButton cmdgrabapaciente 
            Height          =   495
            Left            =   585
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmTarjeta.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Grabar registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   495
            Left            =   60
            Picture         =   "frmTarjeta.frx":037A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Búsqueda de pacientes"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame fraImpresion 
         Height          =   720
         Left            =   5040
         TabIndex        =   54
         Top             =   6840
         Width           =   645
         Begin VB.CommandButton cmdImprime 
            Enabled         =   0   'False
            Height          =   495
            Left            =   75
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmTarjeta.frx":04EC
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Impresión del reporte"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Búsqueda de pacientes con tarjeta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7125
         Left            =   -74805
         TabIndex        =   48
         Top             =   495
         Width           =   8595
         Begin VB.CheckBox chkDepende 
            Caption         =   "Dependientes"
            Height          =   195
            Left            =   5010
            TabIndex        =   32
            ToolTipText     =   "Imprimir tarjetas de dependientes"
            Top             =   6698
            Width           =   1335
         End
         Begin VB.CommandButton cmdImpAgrupado 
            Caption         =   "Imprimir agrupado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   6420
            TabIndex        =   31
            Top             =   6570
            Width           =   2055
         End
         Begin VB.Frame Frame1 
            Caption         =   "Datos de los pacientes"
            Height          =   675
            Left            =   945
            TabIndex        =   52
            Top             =   1920
            Width           =   6690
            Begin VB.OptionButton optDatosPaciente 
               Caption         =   "# Expediente"
               Height          =   195
               Index           =   3
               Left            =   5190
               TabIndex        =   28
               ToolTipText     =   "Búsqueda por número de cuenta"
               Top             =   300
               Width           =   1335
            End
            Begin VB.OptionButton optDatosPaciente 
               Caption         =   "Nombre"
               Height          =   195
               Index           =   2
               Left            =   3525
               TabIndex        =   27
               ToolTipText     =   "Búsqueda por nombre"
               Top             =   300
               Width           =   855
            End
            Begin VB.OptionButton optDatosPaciente 
               Caption         =   "Materno"
               Height          =   195
               Index           =   1
               Left            =   1830
               TabIndex        =   26
               ToolTipText     =   "Búsqueda por apellido materno"
               Top             =   300
               Width           =   885
            End
            Begin VB.OptionButton optDatosPaciente 
               Caption         =   "Paterno"
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   25
               ToolTipText     =   "Búsqueda por apellido paterno"
               Top             =   300
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Búsqueda por fechas"
            Height          =   885
            Left            =   945
            TabIndex        =   49
            Top             =   915
            Width           =   6690
            Begin VB.Frame Frame4 
               Height          =   570
               Left            =   3105
               TabIndex        =   57
               Top             =   180
               Width           =   45
            End
            Begin VB.OptionButton optCualFecha 
               Caption         =   "Fecha vencimiento"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   1365
               TabIndex        =   22
               Top             =   525
               Width           =   1680
            End
            Begin VB.OptionButton optCualFecha 
               Caption         =   "Fecha expedición"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   1365
               TabIndex        =   21
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
            Begin VB.CheckBox chkFecha 
               Caption         =   "Fechas"
               Height          =   195
               Left            =   315
               TabIndex        =   20
               ToolTipText     =   "Búsqueda por rango de fechas"
               Top             =   405
               Width           =   825
            End
            Begin MSMask.MaskEdBox mskFechaIni 
               Height          =   315
               Left            =   3605
               TabIndex        =   23
               ToolTipText     =   "Fecha de inicio de búsqueda"
               Top             =   345
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFechaFin 
               Height          =   315
               Left            =   5220
               TabIndex        =   24
               ToolTipText     =   "Fecha de termino de búsqueda"
               Top             =   345
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin VB.Frame Frame3 
               Height          =   570
               Left            =   1230
               TabIndex        =   56
               Top             =   180
               Width           =   45
            End
            Begin VB.Label lbl2 
               AutoSize        =   -1  'True
               Caption         =   "Al"
               Enabled         =   0   'False
               Height          =   195
               Left            =   4960
               TabIndex        =   51
               Top             =   405
               Width           =   135
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               Caption         =   "Del"
               Enabled         =   0   'False
               Height          =   195
               Left            =   3240
               TabIndex        =   50
               Top             =   405
               Width           =   240
            End
         End
         Begin VB.TextBox txtBusqueda 
            Height          =   315
            Left            =   945
            MaxLength       =   30
            TabIndex        =   29
            ToolTipText     =   "Criterios para la búsqueda de solicitudes"
            Top             =   2670
            Width           =   6690
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdhConsulta 
            Height          =   3435
            Left            =   105
            TabIndex        =   30
            ToolTipText     =   "Busqueda"
            Top             =   3060
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   6059
            _Version        =   393216
            BackColorBkg    =   16777215
            GridColor       =   -2147483633
            AllowBigSelection=   0   'False
            HighLight       =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame Frame7 
            Caption         =   "Procedencia"
            Height          =   615
            Left            =   945
            TabIndex        =   63
            Top             =   255
            Width           =   6690
            Begin VB.CheckBox chkSeleccion 
               Caption         =   "Seleccionar"
               Height          =   195
               Left            =   345
               TabIndex        =   18
               Top             =   270
               Width           =   1275
            End
            Begin VB.ComboBox cboProcede 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   1770
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   19
               ToolTipText     =   "Procedencia del paciente"
               Top             =   210
               Width           =   4710
            End
         End
      End
      Begin VB.Frame fraDependiente 
         Caption         =   "Datos dependientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   330
         TabIndex        =   47
         Top             =   3780
         Width           =   8325
         Begin VB.Frame Frame5 
            Height          =   630
            Left            =   195
            TabIndex        =   58
            Top             =   180
            Width           =   7935
            Begin VB.TextBox txtCvePac 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1260
               MaxLength       =   8
               TabIndex        =   60
               ToolTipText     =   "Número de expediente del paciente"
               Top             =   195
               Width           =   1320
            End
            Begin VB.CheckBox chkImprimeDependiente 
               Caption         =   "Imprimir todos"
               Height          =   195
               Left            =   6525
               TabIndex        =   59
               Top             =   255
               Width           =   1320
            End
            Begin VB.Frame Frame6 
               Height          =   600
               Left            =   6330
               TabIndex        =   62
               Top             =   0
               Width           =   45
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Expediente"
               Height          =   195
               Left            =   195
               TabIndex        =   61
               Top             =   255
               Width           =   795
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPacientes 
            Height          =   1890
            Left            =   195
            TabIndex        =   14
            Top             =   870
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3334
            _Version        =   393216
            BackColorBkg    =   16777215
            GridColor       =   -2147483633
            AllowBigSelection=   0   'False
            HighLight       =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraTitular 
         Caption         =   "Datos del titular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   330
         TabIndex        =   34
         Top             =   525
         Width           =   8325
         Begin VB.CheckBox chkTitular 
            Caption         =   "Imprimir titular"
            Height          =   195
            Left            =   6840
            TabIndex        =   13
            Top             =   2685
            Value           =   1  'Checked
            Width           =   1320
         End
         Begin VB.ComboBox cboProcede 
            Height          =   315
            Index           =   0
            Left            =   1605
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Procedencia del paciente"
            Top             =   2250
            Width           =   4005
         End
         Begin VB.Frame fraSexo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   405
            Left            =   5820
            TabIndex        =   36
            Top             =   300
            Width           =   2235
            Begin VB.OptionButton optFemenino 
               Caption         =   "Femenino"
               Height          =   195
               Left            =   1185
               TabIndex        =   2
               ToolTipText     =   "Sexo femenino"
               Top             =   150
               Width           =   990
            End
            Begin VB.OptionButton optMasculino 
               Caption         =   "Masculino"
               Height          =   195
               Left            =   45
               TabIndex        =   1
               ToolTipText     =   "Sexo masculino"
               Top             =   150
               Width           =   1020
            End
         End
         Begin VB.TextBox txtTelefono 
            Height          =   315
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Teléfono del paciente"
            Top             =   1140
            Width           =   1230
         End
         Begin VB.TextBox txtRfc 
            Height          =   315
            Left            =   5580
            Locked          =   -1  'True
            TabIndex        =   6
            ToolTipText     =   "Registro Federal de Causantes del paciente"
            Top             =   1140
            Width           =   1170
         End
         Begin VB.TextBox txtNombrePac 
            Height          =   315
            Left            =   1605
            Locked          =   -1  'True
            TabIndex        =   3
            ToolTipText     =   "Nombre del paciente"
            Top             =   765
            Width           =   6555
         End
         Begin VB.TextBox txtDirPac 
            Height          =   315
            Left            =   1605
            Locked          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Dirección del paciente"
            Top             =   1500
            Width           =   6555
         End
         Begin VB.TextBox txtTipoPaciente 
            Height          =   315
            Left            =   1605
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Tipo de paciente"
            Top             =   1875
            Width           =   6555
         End
         Begin VB.TextBox txtFechaNac 
            Height          =   315
            Left            =   1605
            Locked          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "Fecha de nacimiento del paciente"
            Top             =   1140
            Width           =   1320
         End
         Begin VB.TextBox txtEdad 
            Height          =   315
            Left            =   7305
            Locked          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Edad actual del paciente"
            Top             =   1140
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskFechaExp 
            Height          =   315
            Left            =   1605
            TabIndex        =   11
            ToolTipText     =   "Fecha de inicio de búsqueda"
            Top             =   2625
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaVen 
            Height          =   315
            Left            =   4380
            TabIndex        =   12
            ToolTipText     =   "Fecha de termino de búsqueda"
            Top             =   2625
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha vencimiento"
            Height          =   195
            Left            =   2925
            TabIndex        =   53
            Top             =   2685
            Width           =   1350
         End
         Begin VB.Label lblTelefono 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            Height          =   195
            Index           =   0
            Left            =   3015
            TabIndex        =   46
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblRfc 
            AutoSize        =   -1  'True
            Caption         =   "R.F.C."
            Height          =   195
            Left            =   5040
            TabIndex        =   45
            Top             =   1200
            Width           =   450
         End
         Begin VB.Label lblDireccion 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   44
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label lblEmpresa 
            AutoSize        =   -1  'True
            Caption         =   "Fecha expedición"
            Height          =   195
            Index           =   13
            Left            =   165
            TabIndex        =   43
            Top             =   2685
            Width           =   1260
         End
         Begin VB.Label lblConvenio 
            AutoSize        =   -1  'True
            Caption         =   "Procedencia"
            Height          =   195
            Index           =   12
            Left            =   165
            TabIndex        =   42
            Top             =   2310
            Width           =   900
         End
         Begin VB.Label lblTipoPaciente 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de paciente"
            Height          =   195
            Index           =   11
            Left            =   165
            TabIndex        =   41
            Top             =   1935
            Width           =   1200
         End
         Begin VB.Label lblFechaNacimiento 
            AutoSize        =   -1  'True
            Caption         =   "Fecha nacimiento"
            Height          =   195
            Index           =   6
            Left            =   165
            TabIndex        =   40
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label lblNombrePaciente 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   39
            Top             =   825
            Width           =   555
         End
         Begin VB.Label lblSexo 
            AutoSize        =   -1  'True
            Caption         =   "Sexo"
            Height          =   195
            Index           =   9
            Left            =   5265
            TabIndex        =   38
            Top             =   465
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Edad"
            Height          =   195
            Index           =   8
            Left            =   6840
            TabIndex        =   37
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Expediente"
            Height          =   195
            Left            =   165
            TabIndex        =   35
            Top             =   465
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Expediente
'| Nombre del Formulario    : frmTarjeta
'-------------------------------------------------------------------------------------
'| Objetivo: Asignar tarjetas de descuento a pacientes externos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Dante Martínez González
'| Autor                    : Dante Martínez González
'| Fecha de Creación        : 23/Enero/2003
'| Modificó                 :
'| Fecha última modificación:
'-------------------------------------------------------------------------------------

Option Explicit

Dim vgstrEstado As String           'Determina el estado de la pantalla
Dim vgstrSql As String              'Contiene las instrucciones SQL
Dim rsDatos As New ADODB.Recordset  'Conexión a las tablas
Dim vglngCont As Long               'Contador estructuras de control

Private Sub pCargaProcedencia()
'Procedimiento que carga las diferentes procedencias de un paciente
'Tipos de paciente diferenciados con signo negativo
'Empresas signo positivo
    On Error GoTo NotificaError


    vgstrSql = "SELECT (tnyCveTipoPaciente * -1) Cve, Rtrim(vchDescripcion) Nombre " & _
        "From AdTipoPaciente " & _
        "Union " & _
        "SELECT intCveEmpresa Cve, Rtrim(vchDescripcion) Nombre " & _
        "From CcEmpresa ORDER BY Nombre"
    
    Set rsDatos = frsRegresaRs(vgstrSql, adLockReadOnly, adOpenForwardOnly)
    
    If rsDatos.RecordCount > 0 Then
        pLlenarCboRs cboProcede(0), rsDatos, 0, 1
        cboProcede(0).ListIndex = 0
        pLlenarCboRs cboProcede(1), rsDatos, 0, 1
        cboProcede(1).ListIndex = 0
    End If
    
    rsDatos.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaProcedencia"))
End Sub

Private Sub pMuestraDatosExt(Clave As Long, Optional Dependiente As Boolean)
'--------------------------------------------------------------------------
' Despliega los datos de los pacientes Externos, unicamente para consulta '
'--------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim rsDatosMostradosExt As New ADODB.Recordset
    'Variables para la cuenta del paciente
    Dim vlstrCta As String
    Dim rsCuenta As New ADODB.Recordset
    Dim vlintConvenio As Integer
    Dim vllngEmpresa As Long
    Dim vlintTipoPaciente As Integer
    
    If Dependiente Then
        'Se valida que el paciente no este asignado en el grid
        If fValidaPac(Clave) Then Exit Sub
    End If

    vlstrSentencia = "SELECT Externo.intNumPaciente, " & _
    "Rtrim(Externo.chrNombre)||' '||Rtrim(Externo.chrApePaterno)||' '||Rtrim(Externo.chrApeMaterno) Nombre, " & _
    "Externo.dtmFechaNac, Externo.chrSexo, TRIM(EXTERNO.CHRCALLE) || ' ' || TRIM(EXTERNO.VCHNUMEROEXTERIOR)||CASE WHEN EXTERNO.VCHNUMEROINTERIOR IS NULL THEN '' ELSE ' Int. '|| Trim(EXTERNO.VchNumeroInterior) END ChrCalle , " & _
    "Externo.chrTelefono, Externo.chrRFC, Rtrim(AdTipoPaciente.vchDescripcion) TipoPaciente, " & _
    "Rtrim(CcTipoConvenio.vchDescripcion) Convenio, " & _
    "Rtrim(CcEmpresa.vchDescripcion) Empresa, " & _
    "isnull(CcEmpresa.intCveEmpresa,0) as CveEmp, " & _
    "isnull(CcTipoConvenio.tnyCveTipoConvenio,0) as CveConvenio, " & _
    "isnull(AdTipoPaciente.tnyCveTipoPaciente,0) CveTipoPac " & _
    "FROM Externo left outer JOIN " & _
    "AdTipoPaciente ON " & _
    "Externo.tnyTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente left outer " & _
    "Join " & _
    "CcTipoConvenio ON " & _
    "Externo.tnyTipoConvenio = CcTipoConvenio.tnyCveTipoConvenio " & _
    "left outer Join " & _
    "CcEmpresa ON " & _
    "Externo.intClaveEmpresa = CcEmpresa.intCveEmpresa " & _
    "Where Externo.intNumPaciente = " & CStr(Clave)
              
    Set rsDatosMostradosExt = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    With rsDatosMostradosExt
         If .RecordCount > 0 Then
            
            If Dependiente = False Then
            
                vgstrEstado = "A"       'Alta de registro
            
                txtClavePaciente.Text = !intNumPaciente
                txtNombrePac.Text = Trim(!Nombre)
            
                'Variables globales
                vglngNumeroPaciente = !intNumPaciente
                vgstrNombrePaciente = !Nombre
                vgstrTipoPaciente = "E"
                
                If Not IsNull(!dtmFechaNac) Then
                    txtFechaNac.Text = Format(!dtmFechaNac, "dd/mmm/yyyy")
                    txtEdad = fstrObtieneEdad(CDate(!dtmFechaNac), Date)
                Else
                    txtFechaNac.Text = ""
                    txtEdad = ""
                End If
                If !chrSexo = "M" Then
                    optMasculino.Value = True
                Else
                    optFemenino.Value = True
                End If
                txtDirPac.Text = IIf(IsNull(!chrCalle), "", RTrim(!chrCalle))
                txtTelefono.Text = IIf(IsNull(!chrTelefono), "", !chrTelefono)
                txtRFC.Text = IIf(IsNull(!chrRFC), "", RTrim(!chrRFC))
                txtTipoPaciente.Text = Trim(!TipoPaciente)
                
                vlintConvenio = !CveConvenio
                vllngEmpresa = !CveEmp
                vlintTipoPaciente = !CveTipoPac
            Else
                'Llenado del grid de pacientes dependientes
                
                If Val(grdPacientes.TextMatrix(grdPacientes.Rows - 1, 1)) <> 0 Then grdPacientes.Rows = grdPacientes.Rows + 1
                
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, 1) = !intNumPaciente
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, 2) = !Nombre
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, 3) = !TipoPaciente
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, 4) = !chrSexo
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, 5) = !dtmFechaNac
            End If
         
         Else
             
             MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
             .Close      'Cierra y se sale
             If Dependiente Then
                 txtCvePac = ""
             Else
                 txtClavePaciente = ""
             End If
             Exit Sub
         
         End If
         .Close
    End With
         
'    'Número de moviemiento del paciente
'    vgstrSql = "Select isnull(intNumCuenta,0) Cuenta From RegistroExterno Where intNumPaciente = " & vglngNumeroPaciente & " And bitFacturado = 0"
'    Set rsDatos = frsRegresaRs(vgstrSql, adLockReadOnly, adOpenForwardOnly)
'    If rsDatos.RecordCount = 0 Then
'
'        '/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
'        vlstrCta = "select intMedicoDefaultPOS Medico from pvParametro"
'        Set rsCuenta = frsRegresaRs(vlstrCta, adLockReadOnly, adOpenForwardOnly)
'        If rsCuenta.RecordCount > 0 Then
'            vglngNumeroCuenta = EntornoSiho.cmdNumCuentaExterno(CLng(Val(txtClavePaciente.Text)), vlintConvenio, vllngEmpresa, vlintTipoPaciente, rsCuenta!Medico)
'        Else
'            MsgBox "No se puede generar el número de cuenta." & Chr(13) & "No existen parametros.", vbExclamation, "Mensaje"
''            pIniciaCaptura
'            Exit Sub
'        End If
'        rsCuenta.Close
'        '/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
'    Else
'        vglngNumeroCuenta = rsDatos!Cuenta
'    End If
'    rsDatos.Close
'
'    txtCuenta = vglngNumeroCuenta
     
    fraTitular.Enabled = True
    fraDependiente.Enabled = True
    cmdgrabapaciente.Enabled = True
    If Dependiente Then
        pEnfocaTextBox txtCvePac
    Else
        cboProcede(0).SetFocus
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraDatosExt"))
End Sub

Private Sub pIniciaCaptura()
'Procedimiento que inicia los campos para la captura
    On Error GoTo NotificaError

    txtClavePaciente.Text = ""
    txtCvePac = ""
'    txtCuenta = ""
    txtNombrePac.Text = ""
    txtFechaNac.Text = ""
    txtEdad = ""
    optMasculino.Value = False
    optFemenino.Value = False
    txtDirPac.Text = ""
    txtTelefono.Text = ""
    txtRFC.Text = ""
    txtTipoPaciente.Text = ""
    txtBusqueda = ""
    mskFechaExp = fdtmServerFecha
    mskFechaVen = DateAdd("yyyy", 1, fdtmServerFecha)
    mskFechaIni = fdtmServerFecha
    mskFechaFin = fdtmServerFecha
    fraTitular.Enabled = False
    fraDependiente.Enabled = False
    cmdImprime.Enabled = False
    cmdgrabapaciente.Enabled = False
    
    vgstrEstado = ""        'Se inicia la variable que determina el estado de la pantalla
'    vllngConsecutivo = -1
        
    pPreparaGridGral grdPacientes, 6, 2, 1, 1
    pConfiguraGrd
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIniciaCaptura"))
End Sub

Private Sub pConfiguraGrd(Optional Cual As Boolean)
'Procedimiento para configurar el grid de pacientes dependientes
    On Error GoTo NotificaError


    If Cual Then
        'Grid de consulta
        With grdhConsulta
            .TextMatrix(0, 1) = "Expediente"
            .TextMatrix(0, 2) = "Nombre"
            .TextMatrix(0, 3) = "Expedición"
            .TextMatrix(0, 4) = "Vencimiento"
            .TextMatrix(0, 5) = "Tipo"
            .TextMatrix(0, 6) = "Titular"
            
            .ColWidth(0) = 100     'fix
            .ColWidth(1) = 900     'expediente
            .ColWidth(2) = 3500     'nombre
            .ColWidth(3) = 1000     'expedición
            .ColWidth(4) = 1000      'vencimiento
            .ColWidth(5) = 1500     'Tipo
            .ColWidth(6) = 0        'cve titular
            
            .ColAlignmentFixed(0) = flexAlignCenterBottom
            .ColAlignmentFixed(1) = flexAlignCenterBottom
            .ColAlignmentFixed(2) = flexAlignCenterBottom
            .ColAlignmentFixed(3) = flexAlignCenterBottom
            .ColAlignmentFixed(4) = flexAlignCenterBottom
            .ColAlignmentFixed(5) = flexAlignCenterBottom
            
            .ColAlignment(1) = flexAlignRightBottom
            .ColAlignment(2) = flexAlignLeftBottom
            .ColAlignment(3) = flexAlignLeftBottom
            .ColAlignment(4) = flexAlignLeftBottom
            .ColAlignment(5) = flexAlignLeftBottom
            
            .Col = 1
            .Row = 1
        End With
        '****************************************
    Else
        'Grid de pacientes
        With grdPacientes
            .TextMatrix(0, 1) = "Expediente"
            .TextMatrix(0, 2) = "Nombre"
            .TextMatrix(0, 3) = "Convenio"
            .TextMatrix(0, 4) = "Sexo"
            .TextMatrix(0, 5) = "Fecha nac."
            
            .ColWidth(0) = 250     'fix
            .ColWidth(1) = 900     'expediente
            .ColWidth(2) = 3500     'nombre
            .ColWidth(3) = 3500     'Convenio
            .ColWidth(4) = 700      'sexo
            .ColWidth(5) = 1000     'fecha
            
            .ColAlignmentFixed(0) = flexAlignCenterBottom
            .ColAlignmentFixed(1) = flexAlignCenterBottom
            .ColAlignmentFixed(2) = flexAlignCenterBottom
            .ColAlignmentFixed(3) = flexAlignCenterBottom
            .ColAlignmentFixed(4) = flexAlignCenterBottom
            .ColAlignmentFixed(5) = flexAlignCenterBottom
            
            .ColAlignment(1) = flexAlignRightBottom
            .ColAlignment(2) = flexAlignLeftBottom
            .ColAlignment(3) = flexAlignLeftBottom
            .ColAlignment(4) = flexAlignLeftBottom
            .ColAlignment(5) = flexAlignRightBottom
            
            .Col = 1
            .Row = 1
        End With
        '****************************************
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrd"))
End Sub

Private Function fValidaPac(Cve As Long) As Boolean
'Valida que el paciente no este asignado en el grid
    On Error GoTo NotificaError

    fValidaPac = False

    With grdPacientes
        For vglngCont = 1 To .Rows - 1
            If Val(.TextMatrix(vglngCont, 1)) = Cve Then
                fValidaPac = True
                MsgBox "¡El paciente ya esta asignado!", vbExclamation, "Mensaje"
                Exit For
            End If
        Next
    End With

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fValidaPac"))
End Function

Private Sub pConsultaTarjeta(Clave As Long)
'Procedimiento para la consulta de los datos de los pacientes con tarjetas
'el parametro Clave contiene la clave del paciente titular de la tarjeta
    On Error GoTo NotificaError

    'Datos del titular
    pMuestraDatosExt Clave
    '*********************
    
    vgstrSql = "Select dtmFechaExpedicion FExp, dtmFechaVencimiento FVen, intCveProcedencia CveProc From PvTarjeta Where intNumPaciente = " & Clave & _
        " And chrTitular = 'T'"
    Set rsDatos = frsRegresaRs(vgstrSql, adLockReadOnly, adOpenForwardOnly)
    With rsDatos
        If .RecordCount > 0 Then
            mskFechaExp = !FExp
            mskFechaVen = !FVen
            cboProcede(0).ListIndex = fintLocalizaCbo(cboProcede(0), !CveProc)
        End If
        .Close
    End With
    
        
    pPreparaGridGral grdPacientes, 6, 2, 1, 1
    pConfiguraGrd
        
    'Datos de los dependientes
    vgstrSql = "SELECT Externo.intNumPaciente, " & _
    "Rtrim(Externo.chrApePaterno)||' '||Rtrim(Externo.chrApeMaterno)||' '||Rtrim(Externo.chrNombre) Nombre, " & _
    "Externo.dtmFechaNac, Externo.chrSexo, Rtrim(AdTipoPaciente.vchDescripcion) TipoPaciente " & _
    "FROM Externo left outer JOIN AdTipoPaciente ON " & _
    "Externo.tnyTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
    "Inner Join PvTarjeta On Externo.intNumPaciente = PvTarjeta.intNumPaciente " & _
    "Where PvTarjeta.intCveTitular = " & Clave & " And chrTitular = 'D'"

    Set rsDatos = frsRegresaRs(vgstrSql, adLockReadOnly, adOpenForwardOnly)
    If rsDatos.RecordCount > 0 Then
        With grdPacientes
            Do While Not rsDatos.EOF
                'Llenado del grid de pacientes dependientes
                
                If Val(.TextMatrix(.Rows - 1, 1)) <> 0 Then .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, 1) = rsDatos!intNumPaciente
                .TextMatrix(.Rows - 1, 2) = rsDatos!Nombre
                .TextMatrix(.Rows - 1, 3) = rsDatos!TipoPaciente
                .TextMatrix(.Rows - 1, 4) = rsDatos!chrSexo
                .TextMatrix(.Rows - 1, 5) = rsDatos!dtmFechaNac
                
                rsDatos.MoveNext
            Loop
        End With
    End If
    rsDatos.Close
    
    cmdImprime.Enabled = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConsultaTarjeta"))
End Sub

Private Sub pBusquedaTarjetas(Optional TipoBusqueda As Byte)
'Procedimiento para la busqueda de pacientes con tarjetas <Titulares o Dependientes>
    'PROCEDIMIENTO QUE BUSCA HOJAS DE CONSULTA DE ACUERDO A UNA SERIE DE CRITERIOS
    'DE ACUERDO A LOS DATOS DE LOS PACIENTES EXTERNOS
    'Tipos de busqueda:
    '0: Por los parametros de búsqueda
    '1: Solo por rango de fechas
    '2: Con los parametros en un rango de fechas
    '3: Con una procedencia
    On Error GoTo NotificaError
    Dim rsTemporal As New ADODB.Recordset
    Dim vlstrsql As String
    Dim vlstrFiltro As String
    Dim vlintCont As Integer
    
    
    pPreparaGridGral grdhConsulta, 7, 2, 1, 1   'se limpia e inicia el grid de búsqueda
    pConfiguraGrd True
    
    vlstrsql = ""
    
    If TipoBusqueda = 0 Or TipoBusqueda = 2 Then        'Busqueda por fechas o por datos de pacientes en un rango de fechas
        If Trim(txtBusqueda.Text) = "" Then Exit Sub
    End If
    
    Me.MousePointer = 11
    
'   If TipoBusqueda <> 1 Then vlstrsql = "Set rowCount 50 "      'Cuando es por fechas se muestran los que sean
    
    vlstrsql = "SELECT Externo.intNumPaciente, "
        If optDatosPaciente(0) Or optDatosPaciente(3) Then
            vlstrsql = vlstrsql + "RTRIM(Externo.chrApePaterno)||' '||RTRIM(Externo.chrApeMaterno)||' '||RTRIM(Externo.chrNombre) AS Paciente, "
        ElseIf optDatosPaciente(1) Then
            vlstrsql = vlstrsql + "RTRIM(Externo.chrApeMaterno)||' '||RTRIM(Externo.chrApePaterno)||' '||RTRIM(Externo.chrNombre) AS Paciente, "
        ElseIf optDatosPaciente(2) Then
            vlstrsql = vlstrsql + "RTRIM(Externo.chrNombre)||' '||RTRIM(Externo.chrApePaterno)||' '||RTRIM(Externo.chrApeMaterno) AS Paciente, "
        End If
        vlstrsql = vlstrsql + "PvTarjeta.dtmFechaExpedicion FExp, PvTarjeta.dtmFechaVencimiento FVen, " & _
        "Case When PvTarjeta.chrTitular = 'T' Then 'TITULAR' Else 'DEPENDIENTE' End Tipo, IsNull(PvTarjeta.intCveTitular,-1) CveTitular " & _
    "From PvTarjeta " & _
        "INNER JOIN Externo ON PvTarjeta.intNumPaciente = Externo.intNumPaciente " & _
    "Where (0 = " & chkSeleccion.Value & " Or PvTarjeta.intCveProcedencia = " & cboProcede(1).ItemData(cboProcede(1).ListIndex) & ") And "

    If optDatosPaciente(0) Then
        vlstrFiltro = "RTRIM(Externo.chrApePaterno)||' '||RTRIM(Externo.chrApeMaterno)||' '||RTRIM(Externo.chrNombre)"
    ElseIf optDatosPaciente(1) Then
        vlstrFiltro = "RTRIM(Externo.chrApeMaterno)||' '||RTRIM(Externo.chrApePaterno)||' '||RTRIM(Externo.chrNombre)"
    ElseIf optDatosPaciente(2) Then
        vlstrFiltro = "RTRIM(Externo.chrNombre)||' '||RTRIM(Externo.chrApePaterno)||' '||RTRIM(Externo.chrApeMaterno)"
    ElseIf optDatosPaciente(3) Then
        vlstrFiltro = "PvTarjeta.intNumPaciente"
    End If
    
    'Se forma la instrucción para la busqueda
    Select Case TipoBusqueda
    Case 0      'Busqueda solo por datos de los pacientes

        vlstrsql = vlstrsql + "(" & vlstrFiltro & " like '" & txtBusqueda.Text & "%') "
        vlstrsql = vlstrsql + "Order By " & vlstrFiltro

    Case 1      'busqueda por rango de fechas
        
        If optCualFecha(0).Value Then
            vlstrsql = vlstrsql + "(PvTarjeta.dtmFechaExpedicion Between " & fstrFechaSQL(mskFechaIni, "00:00:00") & " And " & fstrFechaSQL(mskFechaFin, "23:59:59") & ") "
            vlstrsql = vlstrsql + "Order By PvTarjeta.dtmFechaExpedicion desc "
        Else
            vlstrsql = vlstrsql + "(PvTarjeta.dtmFechaVencimiento Between " & fstrFechaSQL(mskFechaIni, "00:00:00") & " And " & fstrFechaSQL(mskFechaFin, "23:59:59") & ") "
            vlstrsql = vlstrsql + "Order By PvTarjeta.dtmFechaVencimiento desc "
        End If

    Case 2      'Datos de pacientes y rangos de fechas

        vlstrsql = vlstrsql + "((" + vlstrFiltro + " like '" + txtBusqueda.Text + "%') And "
        If optCualFecha(0).Value Then
            vlstrsql = vlstrsql + "(PvTarjeta.dtmFechaExpedicion Between " & fstrFechaSQL(mskFechaIni, "00:00:00") & " And " & fstrFechaSQL(mskFechaFin, "23:59:59") & ")) "
            vlstrsql = vlstrsql + "Order By PvTarjeta.dtmFechaExpedicion desc "
        Else
            vlstrsql = vlstrsql + "(PvTarjeta.dtmFechaVencimiento Between " & fstrFechaSQL(mskFechaIni, "00:00:00") & " And " & fstrFechaSQL(mskFechaFin, "23:59:59") & ")) "
            vlstrsql = vlstrsql + "Order By PvTarjeta.dtmFechaVencimiento desc "
        End If

    End Select
'   vlstrsql = vlstrsql + " Set RowCount 0"
    ' * * * * Se termina de formar la instrucción * * * * * *
    
    Set rsTemporal = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly, IIf(TipoBusqueda <> 1, 50, 1000))
    
'            .TextMatrix(0, 1) = "Expediente"
'            .TextMatrix(0, 2) = "Nombre"
'            .TextMatrix(0, 3) = "Expedición"
'            .TextMatrix(0, 4) = "Vencimiento"
'            .TextMatrix(0, 5) = "Tipo"
'            .TextMatrix(0, 6) = "Titular"
    If rsTemporal.RecordCount > 0 Then
        vlintCont = 1
        rsTemporal.MoveFirst
        With grdhConsulta
            .Visible = False
            .Redraw = False
            Do While Not rsTemporal.EOF
                .TextMatrix(vlintCont, 1) = IIf(IsNull(rsTemporal!intNumPaciente), "", Trim(rsTemporal!intNumPaciente))
                .TextMatrix(vlintCont, 2) = IIf(IsNull(rsTemporal!Paciente), "", Trim(rsTemporal!Paciente))
                .TextMatrix(vlintCont, 3) = IIf(IsNull(rsTemporal!FExp), "", Trim(rsTemporal!FExp))
                .TextMatrix(vlintCont, 4) = IIf(IsNull(rsTemporal!FVen), "", FormatDateTime(rsTemporal!FVen, vbShortDate))
                .TextMatrix(vlintCont, 5) = IIf(IsNull(rsTemporal!Tipo), "", Trim(rsTemporal!Tipo))
                .TextMatrix(vlintCont, 6) = IIf(IsNull(rsTemporal!CveTitular), "", Trim(rsTemporal!CveTitular))
                
                If rsTemporal.Bookmark <> rsTemporal.RecordCount Then .Rows = .Rows + 1
                
                vlintCont = vlintCont + 1
                
                rsTemporal.MoveNext
            Loop
            
            .Redraw = True
            .Visible = True
        End With
    Else
        If TipoBusqueda = 1 Then MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
    End If
    rsTemporal.Close

    Me.MousePointer = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBusquedaHojas"))
End Sub

Private Sub pImprime(TipoRep As Integer)
'Procedimiento para la impresión de tarjetas
'TipoRep indica que reporte es el que se va a generar
    On Error GoTo NotificaError
            
    Dim X As Integer 'Contador
            
    vlY = 0
    
    Printer.ScaleMode = vbCharacters
    Printer.Font.Name = "Arial"
    'Printer.PaperSize = 5
    Printer.Font.Size = 8
    
    
    If TipoRep = 1 Then  'Impresión de los datos de la pantalla principal
        If chkTitular.Value = 1 Then    'Si se quiere imprimir el titular
        
            MsgBox SIHOMsg(343), vbInformation, "Mensaje"
        
            pImpDato 62, Trim(vglngNumeroPaciente), vlY
            vlY = vlY + 8
            Printer.Font.Size = 7
            pImpDato 46, Trim(txtNombrePac), vlY
            Printer.Font.Size = 8
            vlY = vlY + 1
            pImpDato 46, "Venc.", vlY
            pImpDato 50, CStr(FormatDateTime(mskFechaExp, vbShortDate)) + " Al " + CStr(FormatDateTime(mskFechaVen, vbShortDate)), vlY
            vlY = vlY + 4
            
            X = X + 1
            
        End If
        
        With grdPacientes
            If Val(.TextMatrix(1, 1)) <> 0 Then
                For vglngCont = 1 To .Rows - 1
                    If .TextMatrix(vglngCont, 0) = "*" Then
                        
                        X = X + 1
                        
                        pImpDato 62, Trim(.TextMatrix(vglngCont, 1)), vlY
                        vlY = vlY + 8
                        Printer.Font.Size = 7
                        pImpDato 46, Trim(.TextMatrix(vglngCont, 2)), vlY
                        Printer.Font.Size = 8
                        vlY = vlY + 1
                        pImpDato 46, "Venc.", vlY
                        pImpDato 50, CStr(FormatDateTime(mskFechaExp, vbShortDate)) + " Al " + CStr(FormatDateTime(mskFechaVen, vbShortDate)), vlY
                        vlY = vlY + 4
                        
                        If X = 5 Then vlY = vlY + 30
                        
                    End If
                Next
            End If
        End With
    Else
    'Impresión de reporte agrupado
        
        If chkSeleccion.Value = 0 Then
            If chkFecha.Value = 0 Then
                MsgBox "Seleccione un filtro para el reporte", vbExclamation, "Mensaje"
                Exit Sub
            End If
        End If
        
        vgstrSql = "SELECT Externo.intNumPaciente Cve, " & _
        "RTRIM(Externo.chrNombre)||' '||RTRIM(Externo.chrApePaterno)||' '||RTRIM(Externo.chrApeMaterno) AS Paciente, " & _
        "PvTarjeta.dtmFechaExpedicion FExp, PvTarjeta.dtmFechaVencimiento FVen, " & _
        "PvTarjeta.chrTitular Tipo " & _
        "From PvTarjeta " & _
        "INNER JOIN Externo ON PvTarjeta.intNumPaciente = Externo.intNumPaciente " & _
        "Where (0 = " & chkSeleccion.Value & " Or PvTarjeta.intCveProcedencia = " & cboProcede(1).ItemData(cboProcede(1).ListIndex) & ")  "
        
        If chkFecha.Value = 1 Then
            If optCualFecha(0).Value Then
                vgstrSql = vgstrSql + " And (PvTarjeta.dtmFechaExpedicion Between " & fstrFechaSQL(mskFechaIni, "00:00:00") & " And " & fstrFechaSQL(mskFechaFin, "23:59:59") & ") "
            Else
                vgstrSql = vgstrSql + " And (PvTarjeta.dtmFechaVencimiento Between " & fstrFechaSQL(mskFechaIni, "00:00:00") & " And " & fstrFechaSQL(mskFechaFin, "23:59:59") & ") "
            End If
        End If
        
        'Filtro para los dependientes
        If chkDepende.Value = 0 Then vgstrSql = vgstrSql & " And (PvTarjeta.chrTitular = 'T')"
    
        Set rsDatos = frsRegresaRs(vgstrSql, adLockReadOnly, adOpenForwardOnly)
        With rsDatos
            If .RecordCount > 0 Then
        
                MsgBox SIHOMsg(343), vbInformation, "Mensaje"
                
                Do While Not .EOF
                    pImpDato 62, Trim(!Cve), vlY
                    vlY = vlY + 8
                    Printer.Font.Size = 7
                    pImpDato 46, Trim(!Paciente), vlY
                    vlY = vlY + 1
                    Printer.Font.Size = 8
                    pImpDato 46, "Venc.", vlY
                    pImpDato 50, CStr(FormatDateTime(!FExp, vbShortDate)) + " Al " + CStr(FormatDateTime(!FVen, vbShortDate)), vlY
                    vlY = vlY + 4
                    .MoveNext
                Loop
            
            Else
                MsgBox SIHOMsg(236), vbExclamation, "Mensaje"
            End If
            .Close
        End With
    End If
    
    '************
    Printer.EndDoc

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

Private Sub cboProcede_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            pEnfocaMkTexto mskFechaExp
        Else
            chkFecha.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboProcede_KeyDown"))
End Sub

Private Sub chkFecha_Click()
    On Error GoTo NotificaError

    optCualFecha(0).Enabled = chkFecha.Value = 1
    optCualFecha(1).Enabled = chkFecha.Value = 1
    lbl1.Enabled = chkFecha.Value = 1
    lbl2.Enabled = chkFecha.Value = 1
    mskFechaIni.Enabled = chkFecha.Value = 1
    mskFechaFin.Enabled = chkFecha.Value = 1
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkFecha_Click"))
End Sub

Private Sub chkFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If optCualFecha(0).Enabled Then
            optCualFecha(0).SetFocus
        Else
            optDatosPaciente(0).SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkFecha_KeyDown"))
End Sub

Private Sub chkImprimeDependiente_Click()
    On Error GoTo NotificaError
    
    With grdPacientes
        If Val(.TextMatrix(1, 1)) <> 0 Then
            .Redraw = False
            For vglngCont = 1 To .Rows - 1
                .TextMatrix(vglngCont, 0) = IIf(chkImprimeDependiente.Value = 1, "*", "")
                .Row = vglngCont
                .Col = 0
                .CellFontSize = 10
                .CellFontBold = True
            Next
            .Redraw = True
        End If
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkImprimeDependiente_Click"))
End Sub

Private Sub chkSeleccion_Click()
    On Error GoTo NotificaError

    cboProcede(1).Enabled = chkSeleccion.Value = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkSeleccion_Click"))
End Sub

Private Sub chkSeleccion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If cboProcede(1).Enabled Then
            cboProcede(1).SetFocus
        Else
            chkFecha.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkSeleccion_KeyDown"))
End Sub

Private Sub chkTitular_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then txtCvePac.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkTitular_KeyDown"))
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError

    chkFecha.Value = 0
    chkSeleccion.Value = 0
    txtBusqueda = ""
    mskFechaIni = fdtmServerFecha
    mskFechaFin = fdtmServerFecha
    pPreparaGridGral grdhConsulta, 7, 2, 1, 1
    pConfiguraGrd True
    SSTObj.Tab = 1
    If txtBusqueda.Enabled And txtBusqueda.Visible Then
      txtBusqueda.SetFocus
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
End Sub

Private Sub cmdgrabapaciente_Click()
'Se graban los datos
    On Error GoTo NotificaError
    Dim rsTarjeta As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    Dim vlblnAlta As Boolean
    
    If vglngNumeroPaciente < 1 Then Exit Sub
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    EntornoSIHO.ConeccionSIHO.BeginTrans        'Comienza la transacción

    vlblnAlta = False
    ' T I T U L A R
    vgstrSql = "Select * From PvTarjeta Where intNumPaciente = " & vglngNumeroPaciente
    Set rsTarjeta = frsRegresaRs(vgstrSql, adLockOptimistic, adOpenDynamic)
    With rsTarjeta
        
        If .RecordCount = 0 Then
          .AddNew
          vlblnAlta = True
        End If
        '****Se graba el titular******
        !dtmFechaExpedicion = CDate(mskFechaExp)
        !dtmFechaVencimiento = CDate(mskFechaVen)
        !intNumPaciente = vglngNumeroPaciente
        !chrTipoPaciente = "E"
        !chrTitular = "T"
        !intCveProcedencia = cboProcede(0).ItemData(cboProcede(0).ListIndex)
        '**** ***** ***** *****
        .Update
        .Close
    End With
    Call pGuardarLogTransaccion(Me.Name, IIf(vlblnAlta, EnmGrabar, EnmCambiar), vllngPersonaGraba, "TARJETA DE DESCUENTO", CStr(vglngNumeroPaciente))
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'D E P E N D I E N T E S
    If Val(grdPacientes.TextMatrix(1, 1)) <> 0 Then
        For vglngCont = 1 To grdPacientes.Rows - 1
            vgstrSql = "Select * From PvTarjeta Where intNumPaciente = " & Val(grdPacientes.TextMatrix(vglngCont, 1))
            Set rsTarjeta = frsRegresaRs(vgstrSql, adLockOptimistic, adOpenDynamic)
            With rsTarjeta
                
                If .RecordCount = 0 Then .AddNew
                
                '****Se graba el dependiente******
                !dtmFechaExpedicion = CDate(mskFechaExp)
                !dtmFechaVencimiento = CDate(mskFechaVen)
                !intNumPaciente = Val(grdPacientes.TextMatrix(vglngCont, 1))
                !chrTipoPaciente = "E"
                !chrTitular = "D"
                !intCveTitular = vglngNumeroPaciente
                '**** ***** ***** *****
                .Update
                .Close
            End With
        Next
    End If
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    EntornoSIHO.ConeccionSIHO.CommitTrans        'Termina la transacción
    
    MsgBox SIHOMsg(284), vbInformation, "Mensaje"
    
    txtClavePaciente.SetFocus
    cmdImprime.Enabled = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdgrabapaciente_Click"))
End Sub

Private Sub cmdImpAgrupado_Click()
'Procedimiento para imprimir las tarjetas
    On Error GoTo NotificaError

    pImprime 2

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImpAgrupado_Click"))
End Sub

Private Sub cmdImprime_Click()
'Procedimiento para imprimir las tarjetas
    On Error GoTo NotificaError

    pImprime 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprime_Click"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError

    If cboProcede(0).ListCount = 0 Then
        'No existen procedencias de pacientes
        MsgBox SIHOMsg(13) + Chr(13) + "Dato: " + cboProcede(0).ToolTipText, vbExclamation, "Mensaje"
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyEscape Then Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    vgOrden = 1
    vgstrNombreForm = Me.Name
    
    pCargaProcedencia
    pIniciaCaptura
    SSTObj.Tab = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    If SSTObj.Tab <> 0 Then
        Cancel = 1
        SSTObj.Tab = 0
    Else
        If vgstrEstado <> "" Then
            Cancel = 1
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then pIniciaCaptura
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
End Sub

Private Sub grdhConsulta_Click()
    On Error GoTo NotificaError

    pOrdenaGridClick grdhConsulta

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdhConsulta_Click"))
End Sub

Private Sub grdhConsulta_DblClick()
    With grdhConsulta
        If Val(.TextMatrix(.Row, 1)) = 0 Then Exit Sub
        pConsultaTarjeta IIf(.TextMatrix(.Row, 6) = -1, .TextMatrix(.Row, 1), .TextMatrix(.Row, 6))
        SSTObj.Tab = 0
    End With

End Sub

Private Sub grdhConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        With grdhConsulta
            If Val(.TextMatrix(.Row, 1)) = 0 Then Exit Sub
            pConsultaTarjeta IIf(.TextMatrix(.Row, 6) = -1, .TextMatrix(.Row, 1), .TextMatrix(.Row, 6))
            SSTObj.Tab = 0
        End With
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdhConsulta_KeyDown"))
End Sub

Private Sub grdPacientes_Click()
    On Error GoTo NotificaError
    
    With grdPacientes
        If .MouseCol = 0 Then
            .Redraw = False
            .TextMatrix(.Row, 0) = IIf(.TextMatrix(.Row, 0) = "*", "", "*")
            .Col = 0
            .CellFontSize = 10
            .CellFontBold = True
            .Redraw = True
        End If
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPacientes_Click"))
End Sub

Private Sub grdPacientes_DblClick()
'Borrado de pacientes dependientes
    On Error GoTo NotificaError

    With grdPacientes
        If .Rows > 0 Then
            If .MouseRow > .FixedCols - 1 Then
                If (.Rows = 2) And (.TextMatrix(.Row, 1)) = "" Then
                    Exit Sub
                Else
                    If MsgBox(SIHOMsg(6), vbCritical + vbYesNo, "Mensaje") = vbYes Then
                        
                        ' ===== Se elimina el paciente de la tabla de tarjetas =====
                        vgstrSql = "Delete From PvTarjeta Where intNumPaciente = " & Val(.TextMatrix(.Row, 1)) & _
                            " And chrTipoPaciente = 'E'"
                        pEjecutaSentencia vgstrSql
                        '***********************************************************
                        
                        If .Rows > 2 Then
                            pBorrarRegMshFGrd grdPacientes, .Row
                        Else
                            If .Row = 1 Then
                                .Clear
                                .ClearStructure
                                pConfiguraGrd
                            End If
                        End If
                    End If
                    .Refresh
                End If
            End If
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPacientes_DblClick"))
End Sub

Private Sub mskFechaExp_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then pEnfocaMkTexto mskFechaVen

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaExp_KeyDown"))
End Sub

Private Sub mskFechaExp_LostFocus()
    On Error GoTo NotificaError

    If Not IsDate(mskFechaExp) Then mskFechaExp = fdtmServerFecha

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaExp_LostFocus"))
End Sub

Private Sub mskFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Not IsDate(mskFechaFin) Then mskFechaFin = fdtmServerFecha
        
        pBusquedaTarjetas 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_KeyDown"))
End Sub

Private Sub mskFechaFin_LostFocus()
    On Error GoTo NotificaError

    If Not IsDate(mskFechaFin) Then mskFechaFin = fdtmServerFecha
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_LostFocus"))
End Sub

Private Sub mskFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then pEnfocaMkTexto mskFechaFin

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaIni_KeyDown"))
End Sub

Private Sub mskFechaIni_LostFocus()
    On Error GoTo NotificaError

    If Not IsDate(mskFechaIni) Then mskFechaIni = fdtmServerFecha
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaIni_LostFocus"))
End Sub

Private Sub mskFechaVen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then chkTitular.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaVen_KeyDown"))
End Sub

Private Sub mskFechaVen_LostFocus()
    On Error GoTo NotificaError

    If Not IsDate(mskFechaVen) Then mskFechaVen = DateAdd("yyyy", 1, fdtmServerFecha)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaVen_LostFocus"))
End Sub

Private Sub optCualFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If mskFechaIni.Enabled Then pEnfocaMkTexto mskFechaIni
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optCualFecha_KeyDown"))
End Sub

Private Sub SSTObj_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
'    Dim vlTemp As Integer
'
'    vlTemp = SSTObj.Tab
'
'    fraTitular.Enabled = vlTemp = 0
'    fraDependiente.Enabled = vlTemp = 0
'    fraBotonera.Enabled = vlTemp = 0
'    fraImpresion.Enabled = vlTemp = 0
'    fraBusqueda.Enabled = vlTemp = 1


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":SSTObj_Click"))
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then If Val(grdhConsulta.TextMatrix(1, 1)) > 0 Then grdhConsulta.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyDown"))
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyPress"))
End Sub

Private Sub txtBusqueda_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    pBusquedaTarjetas IIf(chkFecha.Value = 1, 2, 0)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyUp"))
End Sub

Private Sub txtClavePaciente_Change()
    Dim i As Integer
  
    vglngNumeroPaciente = 0
    vgstrNombrePaciente = ""
    vgstrTipoPaciente = ""
    
    txtNombrePac.Text = ""
    txtFechaNac.Text = ""
    txtEdad = ""
    optMasculino.Value = False
    optFemenino.Value = False
    txtDirPac.Text = ""
    txtTelefono.Text = ""
    txtRFC.Text = ""
    txtTipoPaciente.Text = ""
    grdPacientes.Row = 1
    If grdPacientes.Rows > 2 Then
      grdPacientes.Rows = 2
      pConfiguraGrd False
    End If
    If grdPacientes.TextMatrix(1, 1) <> "" Then
      For i = 1 To grdPacientes.Cols - 1
        grdPacientes.TextMatrix(1, i) = ""
      Next
    End If
    cmdgrabapaciente.Enabled = False
   
End Sub

Private Sub txtClavePaciente_KeyDown(KeyCode As Integer, Shift As Integer)
'Procedimiento que busca el paciente por número de expediente
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Trim(txtClavePaciente) = "" Then     'Se muestra la búsqueda de pacientes
            '***** se mandan los parametros que requiere la forma de búsqueda de pacientes *****
            With FrmBusquedaPacientes
                .vgstrTipoPaciente = "E"
                .Caption = .Caption & " Externos"
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "C"
                .optSoloActivos.Enabled = False
                .optTodos.Value = True
                .vgStrOtrosCampos = ", (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else ' Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
                " From ExPacienteDomicilio " & _
                " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
                " And GnDomicilio.intCveTipoDomicilio = 1 " & _
                " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
                " ExPaciente.dtmFechaNacimiento as ""Fecha Nac."", " & _
                " (Select GnTelefono.vchTelefono " & _
                " From ExPacienteTelefono " & _
                " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
                " And GnTelefono.intCveTipoTelefono = 1 " & _
                " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Teléfono "
                .vgstrTamanoCampo = "950,3400,4100,990,980"
            End With
            
            vglngNumeroPaciente = FrmBusquedaPacientes.flngRegresaPaciente()
            
            If vglngNumeroPaciente = -1 Then
                '¿Desea dar de alta al paciente?
                If MsgBox(SIHOMsg(258), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    
                    frmAdmisionPaciente.vlblnMostrarTabGenerales = True
                    frmAdmisionPaciente.vlblnMostrarTabInternamiento = True
                    frmAdmisionPaciente.vlblnMostrarTabInternos = False
                    frmAdmisionPaciente.vlblnMostrarTabPrepagos = False
                    frmAdmisionPaciente.vlblnMostrarTabIngresosPrevios = False
                    frmAdmisionPaciente.vlblnMostrarTabEgresados = False
                    frmAdmisionPaciente.vlblnMostrarTabExternos = False
                    frmAdmisionPaciente.vlintPestañaInicial = 0
                    
                    frmAdmisionPaciente.blnAbrirCuenta = False
                    frmAdmisionPaciente.blnActivar = False
                    frmAdmisionPaciente.blnHabilitarAbrirCuenta = False
                    frmAdmisionPaciente.blnHabilitarActivar = False
                    frmAdmisionPaciente.blnHabilitarReporte = False
                    frmAdmisionPaciente.blnConsulta = False
                    frmAdmisionPaciente.vglngExpedienteConsulta = 0
                    frmAdmisionPaciente.blnHonorariosCC = False
                    
                    frmAdmisionPaciente.vllngNumeroOpcionExterno = 352
                    frmAdmisionPaciente.Show vbModal, Me
                    
                    vglngNumeroPaciente = frmAdmisionPaciente.vglngExpediente
                    
                    If vglngNumeroPaciente > 0 Then
                        pMuestraDatosExt vglngNumeroPaciente
                    End If
                    Exit Sub
                End If
            Else
                pMuestraDatosExt vglngNumeroPaciente
            End If
        Else
            pMuestraDatosExt Val(txtClavePaciente)
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClavePaciente_KeyDown"))
End Sub

Private Sub txtClavePaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClavePaciente_KeyPress"))
End Sub

Private Sub txtCvePac_KeyDown(KeyCode As Integer, Shift As Integer)
'Procedimiento que busca el paciente por número de expediente
    On Error GoTo NotificaError
    Dim vlClave As Long

    If KeyCode = vbKeyReturn And vglngNumeroPaciente > 0 Then
        If Trim(txtCvePac) = "" Then     'Se muestra la búsqueda de pacientes
            '***** se mandan los parametros que requiere la forma de búsqueda de pacientes *****
            With FrmBusquedaPacientes
                .vgstrTipoPaciente = "E"
                .Caption = .Caption & " Externos"
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "C"
                .optSoloActivos.Enabled = False
                .optTodos.Value = True
                .vgStrOtrosCampos = ", (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else ' Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
                " From ExPacienteDomicilio " & _
                " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
                " And GnDomicilio.intCveTipoDomicilio = 1 " & _
                " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
                " ExPaciente.dtmFechaNacimiento as ""Fecha Nac."", " & _
                " (Select GnTelefono.vchTelefono " & _
                " From ExPacienteTelefono " & _
                " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
                " And GnTelefono.intCveTipoTelefono = 1 " & _
                " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Teléfono "
                .vgstrTamanoCampo = "950,3400,4100,990,980"
            End With
            
            vlClave = FrmBusquedaPacientes.flngRegresaPaciente()
            
            If vlClave = -1 Then
                '¿Desea dar de alta al paciente?
                If MsgBox(SIHOMsg(258), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                
                    frmAdmisionPaciente.vlblnMostrarTabGenerales = True
                    frmAdmisionPaciente.vlblnMostrarTabInternamiento = True
                    frmAdmisionPaciente.vlblnMostrarTabInternos = False
                    frmAdmisionPaciente.vlblnMostrarTabPrepagos = False
                    frmAdmisionPaciente.vlblnMostrarTabIngresosPrevios = False
                    frmAdmisionPaciente.vlblnMostrarTabEgresados = False
                    frmAdmisionPaciente.vlblnMostrarTabExternos = False
                    frmAdmisionPaciente.vlintPestañaInicial = 0
                    
                    frmAdmisionPaciente.blnAbrirCuenta = False
                    frmAdmisionPaciente.blnActivar = False
                    frmAdmisionPaciente.blnHabilitarAbrirCuenta = False
                    frmAdmisionPaciente.blnHabilitarActivar = False
                    frmAdmisionPaciente.blnHabilitarReporte = False
                    frmAdmisionPaciente.blnConsulta = False
                    frmAdmisionPaciente.vglngExpedienteConsulta = 0
                    frmAdmisionPaciente.blnHonorariosCC = False
                    
                    frmAdmisionPaciente.vllngNumeroOpcionExterno = 352
                    frmAdmisionPaciente.Show vbModal, Me
                    
                    vlClave = frmAdmisionPaciente.vglngExpediente
                    
                    If vlClave > 0 Then
                        'Procedimiento para desplegar los datos del paciente en el grid
                        pMuestraDatosExt vlClave, True
                    End If
                    Exit Sub
                End If
            Else
                'Procedimiento para desplegar los datos del paciente en el grid
                pMuestraDatosExt vlClave, True
            End If
        Else
            'Procedimiento para desplegar los datos del paciente en el grid
            pMuestraDatosExt Val(txtCvePac), True
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClavePaciente_KeyDown"))
End Sub

Private Sub txtCvePac_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCvePac_KeyPress"))
End Sub
