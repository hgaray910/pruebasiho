VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manejo de socios"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBotonera 
      Height          =   735
      Left            =   4158
      TabIndex        =   0
      Top             =   9060
      Width           =   2745
      Begin VB.CommandButton cmdCambioSocio 
         Caption         =   "Titular a dependiente"
         Height          =   495
         Left            =   1560
         TabIndex        =   108
         ToolTipText     =   "Cambiar socio de titular a dependiente"
         Top             =   165
         Width           =   1110
      End
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   495
         Left            =   570
         Picture         =   "frmSocios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Guardar el registro"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   495
         Left            =   75
         Picture         =   "frmSocios.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Búsqueda"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   1065
         Picture         =   "frmSocios.frx":0314
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Imprimir informe del socio"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   9015
      Left            =   120
      TabIndex        =   55
      Top             =   0
      Width           =   10815
      Begin VB.TextBox txtFactorRH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         ToolTipText     =   "Factor RH (+/ -)"
         Top             =   2880
         Width           =   255
      End
      Begin VB.Frame fraEmergencia 
         Caption         =   "En caso de emergencia comunicarse con"
         Height          =   1095
         Left            =   5040
         TabIndex        =   178
         Top             =   2880
         Width           =   5655
         Begin VB.TextBox txtTelefonoEmergencia 
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   21
            ToolTipText     =   "Teléfono para comunicarse en caso de emergencia"
            Top             =   600
            Width           =   3975
         End
         Begin VB.TextBox TxtNombreEmergencia 
            Height          =   315
            Left            =   1440
            MaxLength       =   150
            TabIndex        =   20
            ToolTipText     =   "Nombre para llamar en caso de emergencia"
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label lblTelefonoEmergencia 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono(s)"
            Height          =   195
            Left            =   120
            TabIndex        =   180
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lblNombreEmergencia 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   179
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.ComboBox cboGrupoSanguineo 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Grupo sanguíneo"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtComentarios 
         Height          =   675
         Left            =   2040
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Comentarios"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CheckBox chkBitExtranjero 
         Caption         =   "Extranjero"
         Height          =   255
         Left            =   6480
         TabIndex        =   19
         Top             =   2520
         Width           =   1000
      End
      Begin VB.CommandButton cmdAgregaImagen 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtSerie 
         Height          =   315
         Left            =   720
         MaxLength       =   5
         TabIndex        =   1
         ToolTipText     =   "Número de serie"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCredencial 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Número de credencial"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdEliminaImagen 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10335
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtClaveContabilidad 
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Número de cuenta contable"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtClaveUnica 
         Height          =   315
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Clave única"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtRutaImagen 
         Height          =   315
         Left            =   8880
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "Femenino"
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   16
         ToolTipText     =   "Femenino"
         Top             =   1500
         Width           =   990
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "Masculino"
         Height          =   195
         Index           =   0
         Left            =   6480
         TabIndex        =   15
         ToolTipText     =   "Masculino"
         Top             =   1500
         Width           =   1065
      End
      Begin VB.TextBox txtEdad 
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         ToolTipText     =   "Edad del socio"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   8
         ToolTipText     =   "Nombre del socio"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtApeMaterno 
         Height          =   315
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Apellido materno del socio"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtApePaterno 
         Height          =   315
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Apellido paterno del socio"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtRegistroSBE 
         Height          =   315
         Left            =   2040
         MaxLength       =   150
         TabIndex        =   5
         ToolTipText     =   "Registro SBE"
         Top             =   720
         Width           =   6735
      End
      Begin VB.TextBox txtCorreoElectronico 
         Height          =   315
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   18
         ToolTipText     =   "Correo electrónico del socio"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboEstadoCivil 
         Height          =   315
         ItemData        =   "frmSocios.frx":04B6
         Left            =   6480
         List            =   "frmSocios.frx":04B8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Estado civil del socio"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtCurp 
         Height          =   315
         Left            =   6480
         MaxLength       =   20
         TabIndex        =   14
         ToolTipText     =   "Clave Única de Registro Poblacional del paciente"
         Top             =   1080
         Width           =   2310
      End
      Begin MSComDlg.CommonDialog cdlImagen 
         Left            =   10200
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox mskFechaNac 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         ToolTipText     =   "Fecha de nacimiento del socio"
         Top             =   2160
         Width           =   1120
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskRFC 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Registro Federal de Contribuyentes"
         Top             =   2520
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         PromptChar      =   " "
      End
      Begin TabDlg.SSTab sstOpcion 
         Height          =   4815
         Left            =   120
         TabIndex        =   109
         Top             =   4080
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Domicilios"
         TabPicture(0)   =   "frmSocios.frx":04BA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraGeneral"
         Tab(0).Control(1)=   "Frame2"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Estado"
         TabPicture(1)   =   "frmSocios.frx":04D6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraEstado"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Documentación"
         TabPicture(2)   =   "frmSocios.frx":04F2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frameSociosNFC"
         Tab(2).Control(1)=   "fraDocumentacion"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Dictámenes"
         TabPicture(3)   =   "frmSocios.frx":050E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Dependientes"
         TabPicture(4)   =   "frmSocios.frx":052A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame7"
         Tab(4).ControlCount=   1
         Begin VB.Frame Frame7 
            Caption         =   "Información del dependiente"
            Height          =   4335
            Left            =   -74880
            TabIndex        =   157
            Top             =   360
            Width           =   10335
            Begin VB.ComboBox cboCiudadD 
               Height          =   315
               Left            =   1800
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   71
               ToolTipText     =   "Ciudad origen del dependiente"
               Top             =   1320
               Width           =   3015
            End
            Begin VB.TextBox txtLugarNacD 
               Height          =   315
               Left            =   1800
               MaxLength       =   100
               TabIndex        =   75
               ToolTipText     =   "Lugar de nacimiento"
               Top             =   2040
               Width           =   7935
            End
            Begin VB.TextBox txtNumeroExteriorD 
               Height          =   315
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   67
               ToolTipText     =   "Número exterior"
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txtNumeroInteriorD 
               Height          =   315
               Left            =   4920
               MaxLength       =   10
               TabIndex        =   68
               ToolTipText     =   "Número interior"
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txtFaxD 
               Height          =   315
               Left            =   8040
               MaxLength       =   50
               TabIndex        =   73
               ToolTipText     =   "Fax local"
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox txtTelefonoD 
               Height          =   315
               Left            =   5640
               MaxLength       =   50
               TabIndex        =   72
               ToolTipText     =   "Telefono local"
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox txtPoblacionD 
               Height          =   315
               Left            =   1800
               MaxLength       =   50
               TabIndex        =   74
               ToolTipText     =   "Población"
               Top             =   1680
               Width           =   7935
            End
            Begin VB.TextBox txtCPD 
               Height          =   315
               Left            =   8040
               MaxLength       =   5
               TabIndex        =   69
               ToolTipText     =   "Código Postal"
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txtColoniaD 
               Height          =   315
               Left            =   1800
               LinkTimeout     =   48
               MaxLength       =   100
               TabIndex        =   70
               ToolTipText     =   "Colonia local"
               Top             =   960
               Width           =   7935
            End
            Begin VB.TextBox txtDomicilioD 
               Height          =   315
               Left            =   1800
               MaxLength       =   100
               TabIndex        =   66
               ToolTipText     =   "Domicilio local"
               Top             =   240
               Width           =   7935
            End
            Begin VB.ComboBox cboDerechosD 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   76
               ToolTipText     =   "Derechos del socio"
               Top             =   2760
               Width           =   2895
            End
            Begin VB.Frame Frame5 
               Caption         =   "Hispanidad"
               Height          =   1215
               Left            =   5160
               TabIndex        =   158
               Top             =   2640
               Width           =   4575
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Cónyuge"
                  Height          =   195
                  Index           =   10
                  Left            =   240
                  TabIndex        =   81
                  ToolTipText     =   "Hispanidad"
                  Top             =   840
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Bisnieto de español"
                  Height          =   195
                  Index           =   9
                  Left            =   2160
                  TabIndex        =   80
                  ToolTipText     =   "Hispanidad"
                  Top             =   600
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Viudo(a)"
                  Height          =   195
                  Index           =   11
                  Left            =   2160
                  TabIndex        =   82
                  ToolTipText     =   "Hispanidad"
                  Top             =   840
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Nieto de español"
                  Height          =   195
                  Index           =   8
                  Left            =   240
                  TabIndex        =   79
                  ToolTipText     =   "Hispanidad"
                  Top             =   600
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Hijo de español"
                  Height          =   195
                  Index           =   7
                  Left            =   2160
                  TabIndex        =   78
                  ToolTipText     =   "Hispanidad"
                  Top             =   360
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Español"
                  Height          =   195
                  Index           =   6
                  Left            =   240
                  TabIndex        =   77
                  ToolTipText     =   "Hispanidad"
                  Top             =   360
                  Width           =   1905
               End
            End
            Begin MSMask.MaskEdBox mskFechaIngresoD 
               Height          =   315
               Left            =   1800
               TabIndex        =   83
               ToolTipText     =   "Fecha de ingreso"
               Top             =   3180
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFechaBajaD 
               Height          =   315
               Left            =   1800
               TabIndex        =   84
               ToolTipText     =   "Fecha de baja"
               Top             =   3540
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFechaEmisionCredencialD 
               Height          =   315
               Left            =   1800
               TabIndex        =   85
               ToolTipText     =   "Fecha de emisión de credencial"
               Top             =   3900
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Ciudad"
               Height          =   195
               Left            =   120
               TabIndex        =   185
               Top             =   1380
               Width           =   495
            End
            Begin VB.Label Label21 
               Caption         =   "Lugar de nacimiento"
               Height          =   195
               Left            =   120
               TabIndex        =   184
               Top             =   2100
               Width           =   1515
            End
            Begin VB.Label Label20 
               Caption         =   "Número interior"
               Height          =   255
               Left            =   3720
               TabIndex        =   183
               Top             =   660
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "Número exterior"
               Height          =   255
               Left            =   120
               TabIndex        =   182
               Top             =   660
               Width           =   1215
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Fax"
               Height          =   195
               Left            =   7560
               TabIndex        =   168
               Top             =   1380
               Width           =   255
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Teléfono"
               Height          =   195
               Left            =   4920
               TabIndex        =   167
               Top             =   1380
               Width           =   630
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Población"
               Height          =   195
               Left            =   120
               TabIndex        =   166
               Top             =   1740
               Width           =   705
            End
            Begin VB.Label Label10 
               Caption         =   "Código postal"
               Height          =   195
               Left            =   6840
               TabIndex        =   165
               Top             =   660
               Width           =   1035
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Colonia"
               Height          =   195
               Left            =   120
               TabIndex        =   164
               Top             =   1020
               Width           =   525
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Domicilio"
               Height          =   195
               Left            =   120
               TabIndex        =   163
               Top             =   300
               Width           =   630
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Emisión de credencial"
               Height          =   195
               Left            =   120
               TabIndex        =   162
               Top             =   3960
               Width           =   1545
            End
            Begin VB.Label Label16 
               Caption         =   "Fecha de baja"
               Height          =   195
               Left            =   120
               TabIndex        =   161
               Top             =   3600
               Width           =   1995
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de ingreso"
               Height          =   195
               Left            =   120
               TabIndex        =   160
               Top             =   3240
               Width           =   1230
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Derechos"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   159
               Top             =   2760
               Width           =   690
            End
         End
         Begin VB.Frame fraGeneral 
            Caption         =   "Domicilio local"
            Height          =   2655
            Left            =   -74880
            TabIndex        =   150
            Top             =   420
            Width           =   10335
            Begin VB.TextBox txtPoblacionT 
               Height          =   315
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   30
               ToolTipText     =   "Población"
               Top             =   1800
               Width           =   7935
            End
            Begin VB.TextBox txtNumeroInterior 
               Height          =   315
               Left            =   5160
               MaxLength       =   10
               TabIndex        =   24
               ToolTipText     =   "Número interior"
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtNumeroExterior 
               Height          =   315
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   23
               ToolTipText     =   "Número exterior"
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox cboCiudad 
               Height          =   315
               Left            =   2040
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   27
               ToolTipText     =   "Ciudad origen del Socio"
               Top             =   1440
               Width           =   3015
            End
            Begin VB.TextBox txtLugarNac 
               Height          =   315
               Left            =   2040
               MaxLength       =   100
               TabIndex        =   31
               ToolTipText     =   "Lugar de nacimiento"
               Top             =   2160
               Width           =   7935
            End
            Begin VB.TextBox txtDomicilio 
               Height          =   315
               Left            =   2040
               MaxLength       =   100
               TabIndex        =   22
               ToolTipText     =   "Domicilio local"
               Top             =   360
               Width           =   7935
            End
            Begin VB.TextBox txtColonia 
               Height          =   315
               Left            =   2040
               MaxLength       =   100
               TabIndex        =   26
               ToolTipText     =   "Colonia local"
               Top             =   1080
               Width           =   7935
            End
            Begin VB.TextBox txtCP 
               Height          =   315
               Left            =   8280
               MaxLength       =   5
               TabIndex        =   25
               ToolTipText     =   "Código postal"
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtTelefono 
               Height          =   315
               Left            =   6000
               MaxLength       =   50
               TabIndex        =   28
               ToolTipText     =   "Teléfono local"
               Top             =   1440
               Width           =   1695
            End
            Begin VB.TextBox txtFax 
               Height          =   315
               Left            =   8280
               MaxLength       =   50
               TabIndex        =   29
               ToolTipText     =   "Fax local"
               Top             =   1440
               Width           =   1695
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Clave"
               Height          =   375
               Left            =   720
               TabIndex        =   151
               Top             =   6600
               Width           =   615
            End
            Begin VB.Label lblPoblacionT 
               AutoSize        =   -1  'True
               Caption         =   "Población"
               Height          =   195
               Left            =   240
               TabIndex        =   181
               Top             =   1860
               Width           =   705
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Código postal"
               Height          =   195
               Left            =   7080
               TabIndex        =   174
               Top             =   780
               Width           =   960
            End
            Begin VB.Label Label22 
               Caption         =   "Número exterior"
               Height          =   255
               Left            =   240
               TabIndex        =   173
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label29 
               Caption         =   "Número interior"
               Height          =   255
               Left            =   3960
               TabIndex        =   172
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label lblLugarNac 
               Caption         =   "Lugar de nacimiento"
               Height          =   195
               Left            =   240
               TabIndex        =   169
               Top             =   2220
               Width           =   1515
            End
            Begin VB.Label lblDomicilio 
               AutoSize        =   -1  'True
               Caption         =   "Domicilio"
               Height          =   195
               Left            =   240
               TabIndex        =   156
               Top             =   420
               Width           =   630
            End
            Begin VB.Label lblColonia 
               AutoSize        =   -1  'True
               Caption         =   "Colonia"
               Height          =   195
               Left            =   240
               TabIndex        =   155
               Top             =   1140
               Width           =   525
            End
            Begin VB.Label lblPoblacion 
               AutoSize        =   -1  'True
               Caption         =   "Ciudad"
               Height          =   195
               Left            =   240
               TabIndex        =   154
               Top             =   1500
               Width           =   495
            End
            Begin VB.Label lblTelefono 
               AutoSize        =   -1  'True
               Caption         =   "Teléfono"
               Height          =   195
               Left            =   5160
               TabIndex        =   153
               Top             =   1500
               Width           =   630
            End
            Begin VB.Label lblFax 
               AutoSize        =   -1  'True
               Caption         =   "Fax"
               Height          =   195
               Left            =   7800
               TabIndex        =   152
               Top             =   1500
               Width           =   255
            End
         End
         Begin VB.Frame fraEstado 
            Caption         =   "Datos del socio"
            Height          =   4215
            Left            =   120
            TabIndex        =   138
            Top             =   360
            Width           =   10335
            Begin VB.TextBox txtDependientes 
               Height          =   315
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               ToolTipText     =   "Número de dependientes"
               Top             =   600
               Width           =   2895
            End
            Begin VB.ComboBox cboCuotaSocio 
               Height          =   315
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   40
               ToolTipText     =   "Tipo de cuota"
               Top             =   960
               Width           =   2895
            End
            Begin VB.ComboBox cboFormaPago 
               Height          =   315
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   41
               ToolTipText     =   "Forma de pago"
               Top             =   1320
               Width           =   2895
            End
            Begin VB.Frame frameHispanidad 
               Caption         =   "Hispanidad"
               Height          =   1215
               Left            =   5160
               TabIndex        =   139
               Top             =   300
               Width           =   4575
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Español"
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   45
                  ToolTipText     =   "Hispanidad"
                  Top             =   360
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Hijo de español"
                  Height          =   195
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   46
                  ToolTipText     =   "Hispanidad"
                  Top             =   360
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Nieto de español"
                  Height          =   195
                  Index           =   2
                  Left            =   240
                  TabIndex        =   47
                  ToolTipText     =   "Hispanidad"
                  Top             =   600
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Viudo(a)"
                  Height          =   195
                  Index           =   5
                  Left            =   2160
                  TabIndex        =   50
                  ToolTipText     =   "Hispanidad"
                  Top             =   840
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Bisnieto de español"
                  Height          =   195
                  Index           =   3
                  Left            =   2160
                  TabIndex        =   48
                  ToolTipText     =   "Hispanidad"
                  Top             =   600
                  Width           =   1905
               End
               Begin VB.OptionButton optHispanidad 
                  Caption         =   "Cónyuge"
                  Height          =   195
                  Index           =   4
                  Left            =   240
                  TabIndex        =   49
                  ToolTipText     =   "Hispanidad"
                  Top             =   840
                  Width           =   1905
               End
            End
            Begin VB.ComboBox cboDerechos 
               Height          =   315
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   42
               ToolTipText     =   "Derechos del socio"
               Top             =   1680
               Width           =   2895
            End
            Begin VB.TextBox txtObservaciones 
               Height          =   1155
               Left            =   2040
               MaxLength       =   1000
               TabIndex        =   54
               ToolTipText     =   "Observaciones"
               Top             =   2880
               Width           =   7695
            End
            Begin VB.OptionButton optTipoSocio 
               Caption         =   "Dependiente"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   3120
               TabIndex        =   38
               ToolTipText     =   "Dependiente"
               Top             =   360
               Width           =   1905
            End
            Begin VB.OptionButton optTipoSocio 
               Caption         =   "Titular"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   2040
               TabIndex        =   37
               TabStop         =   0   'False
               ToolTipText     =   "Titular"
               Top             =   360
               Width           =   945
            End
            Begin VB.CheckBox chkAutorizadoPagoCaja 
               Caption         =   "Autorizado para efectuar pagos del ejercicio en caja general"
               Height          =   255
               Left            =   5160
               TabIndex        =   51
               ToolTipText     =   "Autorizacion para realizar pagos en caja"
               Top             =   1740
               Width           =   4695
            End
            Begin MSMask.MaskEdBox mskFechaIngreso 
               Height          =   315
               Left            =   2040
               TabIndex        =   43
               ToolTipText     =   "Fecha de ingreso"
               Top             =   2160
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFechaBaja 
               Height          =   315
               Left            =   2040
               TabIndex        =   44
               ToolTipText     =   "Fecha de baja"
               Top             =   2520
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFechaUltimoPago 
               Height          =   315
               Left            =   7200
               TabIndex        =   52
               ToolTipText     =   "Fecha de último pago"
               Top             =   2160
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFechaEmisionCredencial 
               Height          =   315
               Left            =   7200
               TabIndex        =   53
               ToolTipText     =   "Fecha de emisión de la credencial"
               Top             =   2520
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin VB.Label lblDependientes 
               AutoSize        =   -1  'True
               Caption         =   "Número de dependientes"
               Height          =   195
               Left            =   120
               TabIndex        =   149
               Top             =   660
               Width           =   1785
            End
            Begin VB.Label lblCuotaSocio 
               AutoSize        =   -1  'True
               Caption         =   "Cuota socio"
               Height          =   195
               Left            =   120
               TabIndex        =   148
               Top             =   1020
               Width           =   840
            End
            Begin VB.Label lblFormaPago 
               AutoSize        =   -1  'True
               Caption         =   "Forma de pago"
               Height          =   195
               Left            =   120
               TabIndex        =   147
               Top             =   1380
               Width           =   1065
            End
            Begin VB.Label lblDerechos 
               AutoSize        =   -1  'True
               Caption         =   "Derechos"
               Height          =   195
               Left            =   120
               TabIndex        =   146
               Top             =   1740
               Width           =   690
            End
            Begin VB.Label lblFechaIngreso 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de ingreso"
               Height          =   195
               Left            =   120
               TabIndex        =   145
               Top             =   2220
               Width           =   1230
            End
            Begin VB.Label lblFechaBaja 
               Caption         =   "Fecha de baja"
               Height          =   195
               Left            =   120
               TabIndex        =   144
               Top             =   2520
               Width           =   1995
            End
            Begin VB.Label lblEmisionCredencias 
               AutoSize        =   -1  'True
               Caption         =   "Emisión de credencial"
               Height          =   195
               Left            =   4920
               TabIndex        =   143
               Top             =   2580
               Width           =   1545
            End
            Begin VB.Label lblFechaUltimoPago 
               AutoSize        =   -1  'True
               Caption         =   "Fecha del último pago"
               Height          =   195
               Left            =   4920
               TabIndex        =   142
               Top             =   2220
               Width           =   1560
            End
            Begin VB.Label lblObservaciones 
               AutoSize        =   -1  'True
               Caption         =   "Observaciones"
               Height          =   195
               Left            =   120
               TabIndex        =   141
               Top             =   2880
               Width           =   1065
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo  de socio"
               Height          =   195
               Left            =   120
               TabIndex        =   140
               Top             =   300
               Width           =   1005
            End
         End
         Begin VB.Frame frameSociosNFC 
            Caption         =   "Socios numerarios que firman conocimiento"
            Height          =   1215
            Left            =   -74880
            TabIndex        =   131
            Top             =   420
            Width           =   10335
            Begin VB.CommandButton cmdQuitarSocio1 
               Height          =   315
               Left            =   8040
               MaskColor       =   &H00DCDCDC&
               Picture         =   "frmSocios.frx":0546
               Style           =   1  'Graphical
               TabIndex        =   171
               ToolTipText     =   "Eliminar socio numerario"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   470
            End
            Begin VB.CommandButton cmdQuitarSocio2 
               Height          =   315
               Left            =   8040
               MaskColor       =   &H00DCDCDC&
               Picture         =   "frmSocios.frx":0888
               Style           =   1  'Graphical
               TabIndex        =   170
               ToolTipText     =   "Eliminar socio numerario"
               Top             =   720
               UseMaskColor    =   -1  'True
               Width           =   470
            End
            Begin VB.TextBox txtCveSocioNum2 
               Height          =   285
               Left            =   8160
               TabIndex        =   135
               Top             =   720
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox txtCveSocioNum1 
               Height          =   285
               Left            =   8160
               TabIndex        =   134
               Top             =   360
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox txtNombreSocio2 
               Height          =   315
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   133
               ToolTipText     =   "Socio numerario que firma conocimiento"
               Top             =   720
               Width           =   5655
            End
            Begin VB.CommandButton cmdCambiarSocio2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7485
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmSocios.frx":0BCA
               Style           =   1  'Graphical
               TabIndex        =   57
               ToolTipText     =   "elegir socio numerario"
               Top             =   720
               UseMaskColor    =   -1  'True
               Width           =   470
            End
            Begin VB.TextBox txtNombreSocio1 
               Height          =   315
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   132
               ToolTipText     =   "Socio numerario que firma conocimiento"
               Top             =   360
               Width           =   5655
            End
            Begin VB.CommandButton cmdCambiarSocio1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7485
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmSocios.frx":0D3C
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Elegir socio numerario"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   470
            End
            Begin VB.Label lblNombreS2 
               Caption         =   "Nombre del socio"
               Height          =   195
               Left            =   240
               TabIndex        =   137
               Top             =   780
               Width           =   1275
            End
            Begin VB.Label lblNombreS1 
               Caption         =   "Nombre del socio"
               Height          =   195
               Left            =   240
               TabIndex        =   136
               Top             =   420
               Width           =   1275
            End
         End
         Begin VB.Frame fraDocumentacion 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   125
            Top             =   1620
            Width           =   10335
            Begin VB.TextBox txtAcreditacionHispanidad 
               Height          =   315
               Left            =   3480
               MaxLength       =   4000
               TabIndex        =   63
               ToolTipText     =   "Acreditación de hispanidad"
               Top             =   1920
               Width           =   5895
            End
            Begin VB.CheckBox chkSolicitudInscripcion 
               Caption         =   "Solicitud de inscripción"
               Height          =   255
               Left            =   240
               TabIndex        =   58
               ToolTipText     =   "Solicitud de inscripción"
               Top             =   360
               Width           =   2295
            End
            Begin VB.CheckBox chkActaMatrimonio 
               Caption         =   "Acta de matrimonio"
               Height          =   255
               Left            =   240
               TabIndex        =   60
               ToolTipText     =   "Acta de matrimonio"
               Top             =   840
               Width           =   3015
            End
            Begin VB.CheckBox chkActaNac 
               Caption         =   "Acta de nacimiento titular"
               Height          =   255
               Left            =   240
               TabIndex        =   59
               ToolTipText     =   "Acta de nacimiento del titular"
               Top             =   600
               Width           =   2655
            End
            Begin VB.CheckBox chkActaNacDep 
               Caption         =   "Acta de nacimiento dependientes"
               Height          =   255
               Left            =   240
               TabIndex        =   61
               ToolTipText     =   "Acta de nacimiento dependientes"
               Top             =   1080
               Width           =   2895
            End
            Begin VB.CheckBox chkSolicitudCambio 
               Caption         =   "Solicitud de cambio de registro"
               Height          =   255
               Left            =   240
               TabIndex        =   62
               ToolTipText     =   "Solicitud de cambio de registro"
               Top             =   1320
               Width           =   3015
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdhHispanidad 
               Height          =   975
               Left            =   3480
               TabIndex        =   126
               ToolTipText     =   "Histórico de clasificaciones de hispanidad"
               Top             =   585
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   1720
               _Version        =   393216
               Cols            =   4
               _NumberOfBands  =   1
               _Band(0).Cols   =   4
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdhCredenciales 
               Height          =   975
               Left            =   6840
               TabIndex        =   127
               ToolTipText     =   "Histórico de cambios de credenciales"
               Top             =   585
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   1720
               _Version        =   393216
               Cols            =   4
               _NumberOfBands  =   1
               _Band(0).Cols   =   4
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.Label lblAcreditacionHispanidad 
               AutoSize        =   -1  'True
               Caption         =   "Acreditación de hispanidad"
               Height          =   195
               Left            =   240
               TabIndex        =   130
               Top             =   1980
               Width           =   1920
            End
            Begin VB.Label lblReclasificación 
               Caption         =   "Reclasificaciones"
               Height          =   195
               Left            =   3480
               TabIndex        =   129
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Cambios de credencial"
               Height          =   195
               Left            =   6840
               TabIndex        =   128
               Top             =   360
               Width           =   1605
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Dictámenes"
            Height          =   4095
            Left            =   -74880
            TabIndex        =   116
            Top             =   360
            Width           =   10335
            Begin VB.TextBox txtNombreDictamen 
               Height          =   315
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   118
               ToolTipText     =   "Nombre completo del socio"
               Top             =   360
               Width           =   3975
            End
            Begin VB.TextBox txtObservacionesDictamen 
               Height          =   555
               Left            =   1560
               MaxLength       =   999
               TabIndex        =   65
               ToolTipText     =   "Observaciones del dictamen"
               Top             =   720
               Width           =   8655
            End
            Begin VB.TextBox txtFolioDictamen 
               Height          =   315
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   117
               ToolTipText     =   "Folio del dictamen"
               Top             =   360
               Width           =   735
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdhDependientes 
               Height          =   2415
               Left            =   120
               TabIndex        =   119
               Top             =   1560
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4260
               _Version        =   393216
               Cols            =   4
               _NumberOfBands  =   1
               _Band(0).Cols   =   4
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSMask.MaskEdBox mskFechaDictamen 
               Height          =   315
               Left            =   8640
               TabIndex        =   64
               ToolTipText     =   "Fecha de entrega del dictamen"
               Top             =   360
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Dependientes"
               Height          =   195
               Left            =   240
               TabIndex        =   124
               Top             =   1320
               Width           =   990
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Nombre"
               Height          =   195
               Left            =   2400
               TabIndex        =   123
               Top             =   420
               Width           =   555
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Observaciones"
               Height          =   195
               Left            =   240
               TabIndex        =   122
               Top             =   780
               Width           =   1065
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de entrega"
               Height          =   195
               Left            =   7200
               TabIndex        =   121
               Top             =   420
               Width           =   1260
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Folio"
               Height          =   195
               Left            =   240
               TabIndex        =   120
               Top             =   420
               Width           =   330
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Domicilio de trabajo"
            Height          =   1575
            Left            =   -74880
            TabIndex        =   110
            Top             =   3120
            Width           =   10335
            Begin VB.TextBox txtCPT 
               Height          =   315
               Left            =   8280
               MaxLength       =   5
               TabIndex        =   34
               ToolTipText     =   "Código Postal del trabajo"
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txtProfesion 
               Height          =   315
               Left            =   6600
               MaxLength       =   75
               TabIndex        =   36
               ToolTipText     =   "Profesión"
               Top             =   960
               Width           =   3375
            End
            Begin VB.TextBox txtTelefonoT 
               Height          =   315
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   35
               ToolTipText     =   "Teléfono del trabajo"
               Top             =   960
               Width           =   3615
            End
            Begin VB.TextBox txtLugarTrabajo 
               Height          =   315
               Left            =   2040
               MaxLength       =   100
               TabIndex        =   32
               ToolTipText     =   "Lugar de trabajo"
               Top             =   240
               Width           =   7935
            End
            Begin VB.TextBox txtDomicilioT 
               Height          =   315
               Left            =   2040
               MaxLength       =   100
               TabIndex        =   33
               ToolTipText     =   "Domicilio del trabajo"
               Top             =   600
               Width           =   4575
            End
            Begin VB.Label lblProfesion 
               Caption         =   "Profesión"
               Height          =   195
               Left            =   5760
               TabIndex        =   115
               Top             =   1020
               Width           =   795
            End
            Begin VB.Label lblLugarTrabajo 
               Caption         =   "Lugar de trabajo"
               Height          =   195
               Left            =   240
               TabIndex        =   114
               Top             =   300
               Width           =   1155
            End
            Begin VB.Label lblDomicilioT 
               Caption         =   "Domicilio"
               Height          =   195
               Left            =   240
               TabIndex        =   113
               Top             =   660
               Width           =   915
            End
            Begin VB.Label lblCPT 
               AutoSize        =   -1  'True
               Caption         =   "Código postal"
               Height          =   195
               Left            =   6840
               TabIndex        =   112
               Top             =   660
               Width           =   960
            End
            Begin VB.Label lblTelefonoT 
               Caption         =   "Teléfono"
               Height          =   195
               Left            =   240
               TabIndex        =   111
               Top             =   1020
               Width           =   915
            End
         End
      End
      Begin VB.Label lblComentarios 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   240
         TabIndex        =   177
         Top             =   3300
         Width           =   870
      End
      Begin VB.Label lblRH 
         AutoSize        =   -1  'True
         Caption         =   "RH"
         Height          =   195
         Left            =   3240
         TabIndex        =   176
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label lblGrupoSanguineo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo sanguíneo"
         Height          =   195
         Left            =   240
         TabIndex        =   175
         Top             =   2940
         Width           =   1245
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   195
         Left            =   240
         TabIndex        =   104
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblCredencial 
         AutoSize        =   -1  'True
         Caption         =   "Credencial"
         Height          =   195
         Left            =   1200
         TabIndex        =   103
         Top             =   300
         Width           =   750
      End
      Begin VB.Image picImagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   8880
         Stretch         =   -1  'True
         ToolTipText     =   "Fotografía del socio"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblClaveContabilidad 
         AutoSize        =   -1  'True
         Caption         =   "Clave contabilidad"
         Height          =   195
         Left            =   5760
         TabIndex        =   102
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblClave 
         AutoSize        =   -1  'True
         Caption         =   "Clave única"
         Height          =   195
         Left            =   3000
         TabIndex        =   101
         Top             =   300
         Width           =   840
      End
      Begin VB.Label lblSexo 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Left            =   5040
         TabIndex        =   100
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblEdad 
         AutoSize        =   -1  'True
         Caption         =   "Edad"
         Height          =   195
         Left            =   3240
         TabIndex        =   99
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label lblFechaNac 
         AutoSize        =   -1  'True
         Caption         =   "Fecha nacimiento"
         Height          =   195
         Left            =   240
         TabIndex        =   98
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre completo"
         Height          =   195
         Left            =   240
         TabIndex        =   97
         Top             =   1860
         Width           =   1245
      End
      Begin VB.Label lblMaterno 
         AutoSize        =   -1  'True
         Caption         =   "Apellido materno"
         Height          =   195
         Left            =   240
         TabIndex        =   96
         Top             =   1500
         Width           =   1170
      End
      Begin VB.Label lblPaterno 
         AutoSize        =   -1  'True
         Caption         =   "Apellido paterno"
         Height          =   195
         Left            =   240
         TabIndex        =   95
         Top             =   1140
         Width           =   1140
      End
      Begin VB.Label lblRegistroSBE 
         AutoSize        =   -1  'True
         Caption         =   "Registro SBE"
         Height          =   195
         Left            =   240
         TabIndex        =   94
         Top             =   780
         Width           =   945
      End
      Begin VB.Label lblEstadoCivil 
         AutoSize        =   -1  'True
         Caption         =   "Estado civil"
         Height          =   195
         Left            =   5040
         TabIndex        =   93
         Top             =   1860
         Width           =   810
      End
      Begin VB.Label lblRfc 
         AutoSize        =   -1  'True
         Caption         =   "R.F.C."
         Height          =   195
         Left            =   240
         TabIndex        =   92
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label lblCurp 
         AutoSize        =   -1  'True
         Caption         =   "CURP"
         Height          =   195
         Left            =   5040
         TabIndex        =   91
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label lblCorreoElectronico 
         AutoSize        =   -1  'True
         Caption         =   "Correo electrónico"
         Height          =   195
         Left            =   5040
         TabIndex        =   90
         Top             =   2220
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmSocios
'-------------------------------------------------------------------------------------
'| Objetivo: Permite el registro de la informacion de los socios,
'|           así como la realizacion de cargos y su facturacion
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Víctor Gonzalez
'| Fecha de Creación        :
'-------------------------------------------------------------------------------------
'|------------------------------------------------------------------------------------
'| CASO:      | DESCRIPCIÓN CORTA:                  | PROGRAMADOR:     | FECHA (MM/AA):
'|-------------------------------------------------------------------------------------
'| 19114      | Agregar forma pago Mensual en Tab:  | RRIVERA          | 03 / 2023
'|            | Estado / Forma de Pago              |                  |
'|-------------------------------------------------------------------------------------


Option Explicit

Public Enum enmStatus
    stNuevo = 1
    stedicion = 2
    stEspera = 3
    stConsulta = 5
End Enum
Public stEstado                        As enmStatus
Private vgrptReporte As CRAXDRT.Report

Dim rs As New ADODB.Recordset
Dim rsSocio As New ADODB.Recordset
Dim rsSocioTitular As New ADODB.Recordset
Dim rsCredencial As New ADODB.Recordset
Dim rsHistoricoCredencial As New ADODB.Recordset
Dim rsHispanidad As New ADODB.Recordset
Dim rsDictamenes As New ADODB.Recordset
Dim rsNumerarios As New ADODB.Recordset
Dim rsDependientes As New ADODB.Recordset

Dim vgblnFotoExistente              As Boolean
Dim vlstrsql                        As String
Dim vlStrEstructuraRFC              As String
Dim vllngPersonaGraba               As Long      'Variable que se utiliza para la validacion de la seguridad al grabar
Dim vllngPersonaGraba2              As Long      'Variable que se utiliza para la validacion de la seguridad al grabar
Public vlblnSocioNumerario          As Boolean   'bandera para llamara la pestaña de busqueda para seleccionar socios numerarios
Dim vlblnNuevoSocio                 As Boolean   'se activa cuando se da de alta un nuevo socio
Dim vlblnSerieNueva                 As Boolean   'se activa cuando se agrega una serie nueva
Dim vllngCveSocio                   As Long      'Almacena la clave del socio en consulta
Public vllngEstatusForma            As Integer   'Almacena el estado de la forma
Dim vlblnCambioSerie                As Boolean   'True cuando es un cambio de serie
Dim vlintCuentaContable             As Integer   'Almacena la Cuenta contable que se trajo de la frmBusquedaCuentaContable
Dim vlstrRutaImagen                 As String    'Almacena la ruta de la imagen del socio para posteriormente borrarla
Dim vlintCredencialActual           As Integer
Dim vlchrSerie                      As String
Dim vlchrClaveUnica                 As String    'Almacena la clave unica del socio

Public vlblnMostrarTabDomicilios    As Boolean
Public vlblnMostrarTabEstado        As Boolean
Public vlblnMostrarTabDocumentacion As Boolean
Public vlblnMostrarTabDictamenes    As Boolean
Public vlblnMostrarTabDependientes  As Boolean
Public vlblnNuevoRegistro           As Boolean
Public vlblnDependiente             As Boolean
Public vlblnAsignaTitular           As Boolean
Public vllngTitular                 As Long
Dim vlstrHispanidadAntes As String ' almacena la hispanidad del socio antes de realizarle un cambio de hispanidad
Dim vlstrSentenciaSQL As String
Dim vlblnCargaGrupoSanguineoUnaVEZ As Boolean

'------------------------------------------------------------------------------
'  Función que valida que los datos de la forma esten correctamente llenados
'  y que no se queden vacios campos que son necesarios
'------------------------------------------------------------------------------
Private Function fblnDatosCorrectos() As Boolean
    On Error GoTo NotificaError
    Dim rsCveUnica As ADODB.Recordset ' para confirmar que no se repitan las claves unicas
    fblnDatosCorrectos = True   'se inicializa la bandera
    
    
    ' si se realizaron modificaciones "manuales a la clave unica se debe de validar
    ' quela clave que se coloca no este ya dada de alta en el sistema
    
    vlstrsql = "SELECT * FROM SOSOCIO WHERE VCHCLAVESOCIO='" & txtClaveUnica.Text & "'"
    If Not vlblnNuevoSocio Then
       vlstrsql = vlstrsql & "and intcvesocio <> " & vllngCveSocio
    End If
    
    Set rsCveUnica = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    If Not rsCveUnica.EOF Then  ' si encuentra algo en la base
       fblnDatosCorrectos = False
       'La clave ya existe, la operación no se realizó.
       MsgBox SIHOMsg(649), vbCritical, "Mensaje"
       pEnfocaTextBox txtClaveUnica
       Exit Function
    End If
      
    
    If fblnDatosCorrectos And Trim(txtSerie) = "" Then
       fblnDatosCorrectos = False
      '|  ¡No ha ingresado datos!
       MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtSerie.ToolTipText, vbCritical, "Mensaje"
       pEnfocaTextBox txtSerie
       Exit Function
    End If
    
    If Not vlblnDependiente Then
        If fblnDatosCorrectos And Trim(txtClaveContabilidad) = "" Then
            fblnDatosCorrectos = False
            '|  ¡No ha ingresado datos!
            MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtClaveContabilidad.ToolTipText, vbCritical, "Mensaje"
            pEnfocaTextBox txtClaveContabilidad
            Exit Function
        End If
        
        If fblnDatosCorrectos And Trim(txtRegistroSBE) = "" Then
            fblnDatosCorrectos = False
            '|  ¡No ha ingresado datos!
            MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtRegistroSBE.ToolTipText, vbCritical, "Mensaje"
            pEnfocaTextBox txtRegistroSBE
            Exit Function
        End If
    Else
        If fblnDatosCorrectos And Trim(txtRegistroSBE) = "" Then
            fblnDatosCorrectos = False
            '|  ¡No ha ingresado datos!
            MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtRegistroSBE.ToolTipText, vbCritical, "Mensaje"
            pEnfocaTextBox txtRegistroSBE
            Exit Function
        End If
        
        If fblnDatosCorrectos And Trim(txtRegistroSBE) = Trim(txtApePaterno.Text) & " " & Trim(txtApeMaterno.Text) & " " & Trim(txtNombre.Text) Then
            fblnDatosCorrectos = False
            '|  ¡No ha ingresado datos!
            MsgBox SIHOMsg(1096), vbCritical, "Mensaje"
            pEnfocaTextBox txtRegistroSBE
            Exit Function
        End If
    End If
    
    '------  Datos Generales  ------
    If fblnDatosCorrectos And Trim(txtApePaterno) = "" Then
        fblnDatosCorrectos = False
        '|  ¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtApePaterno.ToolTipText, vbCritical, "Mensaje"
        pEnfocaTextBox txtApePaterno
        Exit Function
    End If
    If fblnDatosCorrectos And Trim(txtApeMaterno) = "" Then
        fblnDatosCorrectos = False
        '|  ¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtApeMaterno.ToolTipText, vbCritical, "Mensaje"
        pEnfocaTextBox txtApeMaterno
        Exit Function
    End If
    
    If fblnDatosCorrectos And Trim(txtNombre) = "" Then
        fblnDatosCorrectos = False
        '|  ¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtNombre.ToolTipText, vbCritical, "Mensaje"
        pEnfocaTextBox txtNombre
        Exit Function
    End If
    
    If fblnDatosCorrectos Then
        If Not IsDate(mskFechaNac.Text) Then
            fblnDatosCorrectos = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            mskFechaNac.SetFocus
            Exit Function
        End If
    End If
    
     If fblnDatosCorrectos Then
        If Trim(cboGrupoSanguineo.Text) <> "" Then
            If Trim(txtFactorRH.Text) = "" Then
                fblnDatosCorrectos = False
                '|  ¡No ha ingresado datos!
              MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & txtFactorRH.ToolTipText, vbCritical, "Mensaje"
               pEnfocaTextBox txtFactorRH
            Exit Function
            End If
        End If
    End If
    
    If vlblnDependiente Then
       If fblnDatosCorrectos And cboCiudadD.ListIndex = -1 Then
        fblnDatosCorrectos = False
        '|  ¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & cboCiudadD.ToolTipText, vbCritical, "Mensaje"
        cboCiudadD.SetFocus
         Exit Function
       End If
       
       If fblnDatosCorrectos And Len(txtCPD.Text) <> 5 And Trim(txtCPD.Text) <> "" Then
          fblnDatosCorrectos = False
          '|  ¡No ha ingresado datos!
          MsgBox SIHOMsg(1181), vbExclamation, "Mensaje"
          pEnfocaTextBox txtCPD
          Exit Function
       End If
    Else
       If fblnDatosCorrectos And cboCiudad.ListIndex = -1 Then
          fblnDatosCorrectos = False
          '|  ¡No ha ingresado datos!
          MsgBox SIHOMsg(2) & Chr(13) & "Dato: " & cboCiudad.ToolTipText, vbCritical, "Mensaje"
          cboCiudad.SetFocus
          sstOpcion.Tab = 0
          Exit Function
       End If
              
       If fblnDatosCorrectos And Len(txtCP.Text) <> 5 And Trim(txtCP.Text) <> "" Then
          fblnDatosCorrectos = False
          '|  ¡No ha ingresado datos!
          MsgBox SIHOMsg(1181), vbExclamation, "Mensaje"
          sstOpcion.Tab = 0
          pEnfocaTextBox txtCP
          Exit Function
       End If
              
       If fblnDatosCorrectos And Len(txtCPT.Text) <> 5 And Trim(txtCPT.Text) <> "" Then
          fblnDatosCorrectos = False
          '|  ¡No ha ingresado datos!
          MsgBox SIHOMsg(1181), vbExclamation, "Mensaje"
          sstOpcion.Tab = 0
          pEnfocaTextBox txtCPT
          Exit Function
       End If
           
       
    End If
    
    If vlblnDependiente Then
        If fblnDatosCorrectos Then
            If Not IsDate(mskFechaIngresoD.Text) Then
                fblnDatosCorrectos = False
                '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                mskFechaIngresoD.SetFocus
                     Exit Function
            End If
        End If
        If fblnDatosCorrectos Then
            If mskFechaBajaD.Enabled Then
                If Not IsDate(mskFechaBajaD.Text) Then
                    If Not mskFechaBajaD.Text = "  /  /    " Then
                        fblnDatosCorrectos = False
                        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                        mskFechaBajaD.SetFocus
                             Exit Function
                    End If
                End If
            End If
        End If
        If fblnDatosCorrectos Then
            If Not IsDate(mskFechaEmisionCredencialD.Text) Then
                If Not mskFechaEmisionCredencialD.Text = "  /  /    " Then
                    fblnDatosCorrectos = False
                    '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                    MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaEmisionCredencialD.SetFocus
                         Exit Function
                End If
            End If
        End If
    Else
        If fblnDatosCorrectos Then
            If Not IsDate(mskFechaIngreso.Text) Then
                fblnDatosCorrectos = False
                '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                sstOpcion.Tab = 1
                MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                mskFechaIngreso.SetFocus
                     Exit Function
            End If
        End If
        If fblnDatosCorrectos Then
            If mskFechaBaja.Enabled Then
                If Not IsDate(mskFechaBaja.Text) Then
                    If Not mskFechaBaja.Text = "  /  /    " Then
                        fblnDatosCorrectos = False
                        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                        sstOpcion.Tab = 1
                        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                        mskFechaBaja.SetFocus
                             Exit Function
                    End If
                End If
            End If
        End If
        If fblnDatosCorrectos Then
            If Not IsDate(mskFechaEmisionCredencial.Text) Then
                If Not mskFechaEmisionCredencial.Text = "  /  /    " Then
                    fblnDatosCorrectos = False
                    '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                    sstOpcion.Tab = 1
                    MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaEmisionCredencial.SetFocus
                         Exit Function
                End If
            End If
        End If
        If fblnDatosCorrectos Then
            If Not IsDate(mskFechaUltimoPago.Text) Then
                If Not mskFechaUltimoPago.Text = "  /  /    " Then
                    fblnDatosCorrectos = False
                    '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                    sstOpcion.Tab = 1
                    MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaUltimoPago.SetFocus
                         Exit Function
                End If
            End If
        End If
        If fblnDatosCorrectos Then
            If Not IsDate(mskFechaDictamen.Text) Then
                If Not mskFechaDictamen.Text = "  /  /    " Then
                    fblnDatosCorrectos = False
                    '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                    sstOpcion.Tab = 3
                    MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaDictamen.SetFocus
                         Exit Function
                End If
            End If
        End If
    End If
    
    
    If Not vlblnDependiente Then
        If txtNombreSocio1 = txtNombreSocio2 And Not txtNombreSocio1.Text = "" And Not txtNombreSocio2.Text = "" Then
            fblnDatosCorrectos = False
            '|  ¡No ha ingresado datos!
            MsgBox SIHOMsg(1095), vbCritical, "Mensaje"
            cmdCambiarSocio1.SetFocus
                 Exit Function
        End If
        If txtNombreSocio1 = txtApePaterno.Text & " " & txtApeMaterno.Text & " " & txtNombre.Text Then
            fblnDatosCorrectos = False
            '|  ¡No ha ingresado datos!
            MsgBox SIHOMsg(1100), vbCritical, "Mensaje"
            cmdCambiarSocio1.SetFocus
                 Exit Function
        End If
        If txtNombreSocio2 = txtApePaterno.Text & " " & txtApeMaterno.Text & " " & txtNombre.Text Then
            fblnDatosCorrectos = False
            '|  ¡No ha ingresado datos!
            MsgBox SIHOMsg(1100), vbCritical, "Mensaje"
            cmdCambiarSocio2.SetFocus
                 Exit Function
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosCorrectos"))
End Function
Private Sub pPosiblesCambiosHispanidad(hispanidad As String) ' activa o descativa las opciones de cambios de hispanidad
    If vlblnDependiente Then ' si esta en la pantalla de dependientes
            Select Case hispanidad
                Case "ET"                    'ET,HT,NT,BT,CT,VT,DE,DH,DN,DB,DC,DV EN LA VENTANA DE DEPENDIENTES
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0  'solo puede cambiar a DC
                     optHispanidad(10).Value = True ' se activa el conyuge por que es el unico al cual puede cambiar
                Case "HT"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0 'solo puede cambiar a DC
                    optHispanidad(10).Value = True ' se activa el conyuge por que es el unico al cual puede cambiar
                Case "NT"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0 'solo puede cambiar a DC
                     optHispanidad(10).Value = True ' se activa el conyuge por que es el unico al cual puede cambiar
                Case "VT"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0 'solo puede cambiar a DC
                    optHispanidad(10).Value = True ' se activa el conyuge por que es el unico al cual puede cambiar
                Case "DE"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0 'no puede cambiar
                Case "DH"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0 'solo puede cambiar a DE
                Case "DN"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0 'solo puede cambiar a DH o DE
                Case "DB"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0 'solo puede cambiar a DN
                Case "DC"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0 'no pude cambiar
            End Select                        '0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,11
                                              'a, b, c, d, e, f, g, h, i, j, k, l
    
    Else ' si esta en la pantalla de titulares
            Select Case hispanidad
                Case "ET"                    'ET,HT,NT,BT,CT,VT,DE,DH,DN,DB,DC,DV EN LA VENTANA DE SOCIOS TITULARES
                    phabilitacambioshispanidad 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 'no puede cambiar
                Case "HT"
                    phabilitacambioshispanidad 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 'solo puede cambiar a ET
                Case "NT"
                    phabilitacambioshispanidad 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0 'solo puede cambiar a HT
                Case "VT"
                    phabilitacambioshispanidad 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0 'no puede cambiar
                Case "DE"
                    phabilitacambioshispanidad 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 'solo puede cambiar a ET
                Case "DH"
                    phabilitacambioshispanidad 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 'solo puede cambiar a HT o ET
                Case "DN"
                    phabilitacambioshispanidad 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0 'solo puede cambiar a HT o ET o NT
                Case "DB" 'no debe aparecer en la consulta de dependientes no hay bisnietos titulares
                 phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 'no puede cambiar
                Case "DC"
                    phabilitacambioshispanidad 1, 1, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0  'solo puede cambiar a ET,HT,NT,VT
                    optHispanidad(0).Value = True ' cambia a titular español como valor por defaul
            End Select                        '0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,11
                                              'a, b, c, d, e, f, g, h, i, j, k, l
    

    End If
End Sub
Private Sub phabilitacambioshispanidad(a As Boolean, b As Boolean, c As Boolean, d As Boolean, e As Boolean, f As Boolean, g As Boolean, h As Boolean, i As Boolean, j As Boolean, k As Boolean, l As Boolean)
    If a Then
        optHispanidad(0).Enabled = True
    Else
        optHispanidad(0).Enabled = False
    End If
    If b Then
        optHispanidad(1).Enabled = True
    Else
        optHispanidad(1).Enabled = False
    End If
    If c Then
        optHispanidad(2).Enabled = True
    Else
        optHispanidad(2).Enabled = False
    End If
    If d Then
        optHispanidad(3).Enabled = True
    Else
        optHispanidad(3).Enabled = False
    End If
    If e Then
        optHispanidad(4).Enabled = True
    Else
        optHispanidad(4).Enabled = False
    End If
    If f Then
        optHispanidad(5).Enabled = True
    Else
        optHispanidad(5).Enabled = False
    End If
    If g Then
        optHispanidad(6).Enabled = True
    Else
        optHispanidad(6).Enabled = False
    End If
    If h Then
        optHispanidad(7).Enabled = True
    Else
        optHispanidad(7).Enabled = False
    End If
    If i Then
        optHispanidad(8).Enabled = True
    Else
        optHispanidad(8).Enabled = False
    End If
    If j Then
        optHispanidad(9).Enabled = True
    Else
        optHispanidad(9).Enabled = False
    End If
    If k Then
        optHispanidad(10).Enabled = True
    Else
        optHispanidad(10).Enabled = False
    End If
    If l Then
        optHispanidad(11).Enabled = True
    Else
        optHispanidad(11).Enabled = False
    End If
End Sub

Private Sub pLlenaEstadoCivil()
    On Error GoTo NotificaError

    Set rs = frsEjecuta_SP("-1", "SP_GNSELESTADOCIVIL")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboEstadoCivil, rs, 0, 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaEstadoCivil"))
    Unload Me
End Sub

Private Sub cboCiudad_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub cboCiudadD_Click()
 If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub cboCuotaSocio_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub cboDerechos_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub cboDerechos_LostFocus()
    If cboDerechos.ListIndex = 0 Then
        mskFechaBaja.Enabled = True
        mskFechaBaja.Mask = ""
        mskFechaBaja.Text = fdtmServerFecha
        mskFechaBaja.Mask = "##/##/####"
    Else
        mskFechaBaja.Enabled = False
        mskFechaBaja.Mask = ""
        mskFechaBaja.Text = "  /  /    "
        mskFechaBaja.Mask = "##/##/####"
    End If
End Sub
Private Sub cboDerechosD_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub cboDerechosD_LostFocus()
    If cboDerechosD.ListIndex = 0 Then
        mskFechaBajaD.Enabled = True
        mskFechaBajaD.Mask = ""
        mskFechaBajaD.Text = fdtmServerFecha
        mskFechaBajaD.Mask = "##/##/####"
    Else
        mskFechaBajaD.Enabled = False
        mskFechaBajaD.Mask = ""
        mskFechaBajaD.Text = "  /  /    "
        mskFechaBajaD.Mask = "##/##/####"
    End If
End Sub
Private Sub cboEstadoCivil_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub cboFormaPago_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub cboGrupoSanguineo_Click()
    
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
    
    If cboGrupoSanguineo.Text = "" Then
       txtFactorRH.Text = ""
       txtFactorRH.Enabled = False
    Else
       txtFactorRH.Enabled = True
    End If

End Sub
Private Sub chkActaMatrimonio_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub chkActaNac_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub chkActaNacDep_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub chkAutorizadoPagoCaja_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub chkBitExtranjero_Click()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub chkSolicitudCambio_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub chkSolicitudInscripcion_GotFocus()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub cmdAgregaImagen_Click()
    With cdlImagen
        '|  Pone el título
        .DialogTitle = "Abrir archivo"
        '|  Establece la lista de archivos que se van a poder seleccionar
        .Filter = "All Picture Files|*.jpg;*.bmp|Bitmaps (*.bmp; *.dib)|*.bmp|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|Metafiles (*.wmf; *.emf)|*.wmf;*.emf"
        '|  Pone el tipo de archivos por defecto
        .FilterIndex = 1
        '|  Establece que los archivos existan
        .Flags = cdlOFNFileMustExist
        '|  Dispara un error si el cuadro de diálogo es cancelado
        .CancelError = True
        '|  Habilita el manejador del error para cachar si se cancela
        On Error Resume Next
        '|  Muestra el cuadro de diálogo
        .ShowOpen
        If Err Then
            '|  Si se cancela el cuadro de diálogo
            Exit Sub
        End If
        '|  Muestra el mensaje
        Set picImagen.Picture = LoadPicture(.FileName, vbLPLarge, vbLPColor)
        txtRutaImagen.Text = .FileName
        If Not vlblnNuevoSocio Then
            pPonEstado stedicion
            'vllngEstatusForma = stedicion
        End If

    End With
End Sub
Private Sub cmdBuscar_Click()
    cmdBuscar.Enabled = False
    cmdCambioSocio.Enabled = False
     
    With frmSociosBusqueda
        .vgblnAsignaTitular = vlblnAsignaTitular
        .vgblnDependiente = vlblnDependiente
        .Show vbModal, Me
        If Not .vgblnEscape Then
            pLlenaInformacionSocio
        Else
         txtSerie.SetFocus
        End If
        Unload frmSociosBusqueda
        
        cmdBuscar.Enabled = True
        cmdCambioSocio.Enabled = True
    
    DoEvents
    End With
End Sub

Private Sub cmdCambiarSocio1_Click()
    With frmSociosBusqueda
        .vglngClaveSocio = 0
        .vgstrNombreSocio = ""
        .Show vbModal, Me
        If .vgstrNombreSocio <> "" Then
            txtCveSocioNum1.Text = .vglngClaveSocio
            txtNombreSocio1.Text = .vgstrNombreSocio
        End If

        If vlblnDependiente Then
            sstOpcion.Tab = 5
        Else
            sstOpcion.Tab = 2
        End If
    End With
    Unload frmSociosBusqueda
   
End Sub

Private Sub cmdCambiarSocio2_Click()

    With frmSociosBusqueda
        .vglngClaveSocio = 0
        .vgstrNombreSocio = ""
        .Show vbModal, Me
        If .vgstrNombreSocio <> "" Then
            txtCveSocioNum2.Text = .vglngClaveSocio
            txtNombreSocio2.Text = .vgstrNombreSocio
        End If

        If vlblnDependiente Then
            sstOpcion.Tab = 5
        Else
            sstOpcion.Tab = 2
        End If
    End With
    Unload frmSociosBusqueda



End Sub

Private Sub cmdCambioSocio_Click()
    cmdBuscar.Enabled = False
    cmdCambioSocio.Enabled = False
    With frmSociosBusqueda
        .vlblnCambioTipoSocio = True
        .vgblnAsignaTitular = vlblnDependiente
        .vgblnDependiente = vlblnDependiente
        .Show vbModal, Me
        If Not .vgblnEscape Then
            .vlblnCambioTipoSocio = False
            pLlenaInformacionSocio
            frmSocios.vlblnAsignaTitular = False
            pGeneraCLaveUnica
            txtRegistroSBE.Text = ""
            pEnfocaTextBox txtRegistroSBE
        Else
            txtSerie.SetFocus
        End If
        Unload frmSociosBusqueda
    
    End With
    cmdCambioSocio.Enabled = True
    cmdBuscar.Enabled = True
     

End Sub
Private Sub pCargaGrupoSanguineo()
On Error GoTo NotificaError
'Se cargan los grupos sanguineos
Dim vlstrSentencia As String
Dim rsDatos As New ADODB.Recordset
    cboGrupoSanguineo.Clear
    
    vlstrSentencia = "SELECT 0 cve, chrGrupo FROM BSTIPOSANGUINEO WHERE BITSTATUS = 1 GROUP BY chrGrupo"
    Set rsDatos = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsDatos.RecordCount > 0 Then
        With cboGrupoSanguineo
            .AddItem "", 0
            Do While Not rsDatos.EOF
                .AddItem Trim(rsDatos!chrGrupo), .ListCount
                .ItemData(.newIndex) = rsDatos!Cve
                rsDatos.MoveNext
            Loop
            .ListIndex = 0
        End With
    Else
        MsgBox SIHOMsg(1183), vbOKOnly + vbExclamation, "Mensaje"
        Unload Me
    End If
    rsDatos.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaGrupoSanguineo"))
    Unload Me
End Sub
Private Sub cmdEliminaImagen_Click()
    txtRutaImagen.Text = ""
    Set picImagen.Picture = LoadPicture("", vbLPLarge, vbLPColor)
    pPonEstado stedicion
    vllngEstatusForma = stedicion
End Sub

Private Sub cmdGrabarRegistro_Click()
    On Error GoTo NotificaError
    
    Dim stmImagen As New ADODB.Stream
    Dim vlblnCambioCredencial As Boolean
    Dim vlstrSerieAnterior As String
    Dim vlintCredencialAnterior As Integer
    Dim vlstrTemp As String
    Dim rsCredencialesTemp As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim i As Integer
    Dim vlchrSerie As String
    Dim vllngGraba1 As Long
    Dim vllngGraba2 As Long
    Dim rs2aVal As New ADODB.Recordset
    Dim vlstr2v As String

    If vlblnDependiente Then
        If Not fblnRevisaPermiso(vglngNumeroLogin, 2415, "E") Then Exit Sub
   Else
        If Not fblnRevisaPermiso(vglngNumeroLogin, 2414, "E") Then Exit Sub

    End If
    
    If Not fblnDatosCorrectos() Then Exit Sub
    
    vllngPersonaGraba2 = 0
    vllngPersonaGraba = 0
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    vlstr2v = "select VCHSENTENCIA from siparametro where vchnombre like 'BITUTILIZASOCIOS'"
    Set rs2aVal = frsRegresaRs(vlstr2v, adLockOptimistic, adOpenForwardOnly)
    If rs2aVal.RecordCount > 0 Then
        If IsNull(rs2aVal!vchSentencia) Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            rs2aVal!vchSentencia = 1
            rs2aVal.Update
            EntornoSIHO.ConeccionSIHO.CommitTrans
        End If
        
        rs2aVal.Requery
        If rs2aVal!vchSentencia = 1 Then
    
            Do While vllngPersonaGraba2 <> vllngPersonaGraba
                vllngPersonaGraba2 = flngPersonaGraba(vgintNumeroDepartamento)
                If vllngPersonaGraba2 = 0 Then Exit Sub
                If vllngPersonaGraba2 <> vllngPersonaGraba Then
                    Exit Do
                Else
                    MsgBox SIHOMsg(1097), vbOKOnly + vbInformation, "Mensaje"
                    vllngPersonaGraba2 = 0
                End If
            Loop
    
        End If
    End If
    rs2aVal.Close
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    With rsSocio
        If stEstado = stNuevo Then
            .AddNew
        End If
        
        If vlblnDependiente Then
            vllngCveSocio = !intcvesocio
            !VCHCLAVESOCIO = Trim(txtClaveUnica.Text)
            !vchregistrosbe = Trim(txtRegistroSBE.Text)
            !vchNombre = Trim(txtNombre.Text)
            !vchApellidoPaterno = Trim(txtApePaterno.Text)
            !vchApellidoMaterno = Trim(txtApeMaterno.Text)
            !vchRFC = Trim(CStr(mskRFC.Text))
            !vchCURP = Trim(CStr(txtCurp.Text))
            !vchPoblacion = Trim(txtPoblacionD.Text)
            !bitExtranjero = IIf(chkBitExtranjero.Value, 1, 0)
                        
            !VCHNOMBREEMERGENCIA = Trim(TxtNombreEmergencia.Text)
            !VCHTELEMERGENCIA = Trim(txtTelefonoEmergencia.Text)
            !VCHCOMENTARIOS = Trim(txtComentarios.Text)
            !CHRFACTORRH = Trim(txtFactorRH)
            !chrGruposanguineo = cboGrupoSanguineo.Text
            !vchLuarNacimiento = Trim(txtLugarNacD.Text)
                        
            If mskFechaNac = "  /  /    " Then
                !dtmFechaNacimiento = Null
            Else
                !dtmFechaNacimiento = CDate(mskFechaNac)
            End If
            !chrSexo = IIf(optSexo(0), "M", "F")
            !intCveEstadoCivil = cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex)
            !vchCorreoElectronico = Trim(txtCorreoElectronico.Text)
            !chrTipoSocio = IIf(vlblnDependiente, "D", "T")
            If mskFechaIngresoD = "  /  /    " Then
                !dtmfechaingreso = Null
            Else
                !dtmfechaingreso = CDate(mskFechaIngresoD)
            End If
            
            If mskFechaBajaD = "  /  /    " Then
                !dtmfechaBaja = Null
            Else
                !dtmfechaBaja = CDate(mskFechaBajaD)
            End If
            
            If mskFechaEmisionCredencialD = "  /  /    " Then
                !dtmFechaCredencial = Null
            Else
                !dtmFechaCredencial = CDate(mskFechaEmisionCredencialD)
            End If
            !intDerechos = cboDerechosD.ListIndex
            
            'If !chrHispanidad = IIf(optHispanidad(6), "ES", IIf(optHispanidad(7), "HE", IIf(optHispanidad(8), "NE", IIf(optHispanidad(9), "BE", IIf(optHispanidad(10), "CE", IIf(optHispanidad(11), "VE", "")))))) Or IsNull(!chrHispanidad) Then
             '   !chrHispanidad = IIf(optHispanidad(6), "ES", IIf(optHispanidad(7), "HE", IIf(optHispanidad(8), "NE", IIf(optHispanidad(9), "BE", IIf(optHispanidad(10), "CE", IIf(optHispanidad(11), "VE", ""))))))
            'Else
            If Not vlblnNuevoSocio Then
                pCambioHispanidad !intcvesocio, vlstrHispanidadAntes, IIf(optHispanidad(6), "ES", IIf(optHispanidad(7), "HE", IIf(optHispanidad(8), "NE", IIf(optHispanidad(9), "BE", IIf(optHispanidad(10), "CE", IIf(optHispanidad(11), "VE", ""))))))
            End If
                !chrHispanidad = IIf(optHispanidad(6), "ES", IIf(optHispanidad(7), "HE", IIf(optHispanidad(8), "NE", IIf(optHispanidad(9), "BE", IIf(optHispanidad(10), "CE", IIf(optHispanidad(11), "VE", ""))))))
            
            'End If
            
            If Trim(txtRutaImagen.Text) <> "" Then
                stmImagen.Type = adTypeBinary
                '|  Se carga el archivo en el objeto stream, para agregarlo al recordset.
                stmImagen.Open
                stmImagen.LoadFromFile txtRutaImagen.Text
                '|  se llena el campo del recordset con el objeto stream.
                !blbFoto = stmImagen.Read
            Else
                !blbFoto = Null
                vgblnFotoExistente = False
            End If
            
            If !vchseriecredencial = Trim(txtSerie.Text) Or IsNull(!vchseriecredencial) Then
                !vchseriecredencial = Trim(txtSerie.Text)
                vlblnCambioCredencial = False
            Else
                vlstrSerieAnterior = !vchseriecredencial
                !vchseriecredencial = Trim(txtSerie.Text)
                vlblnCambioCredencial = True
            End If
            
            If txtCredencial = "" Then
                !intnumeroCredencial = Null
            Else
                If !intnumeroCredencial = CInt(txtCredencial) Or IsNull(!intnumeroCredencial) Then
                    !intnumeroCredencial = CInt(txtCredencial)
                    vlblnCambioCredencial = False
                Else
                    vlintCredencialAnterior = !intnumeroCredencial
                    !intnumeroCredencial = CInt(txtCredencial)
                    vlblnCambioCredencial = True
                End If
            End If
            
            !intnumerocuentacontable = Trim(vlintCuentaContable)
            
            If vlblnCambioCredencial = True Then
              vlstrSentenciaSQL = "insert into SOHISTORICOCREDENCIAL values(null," & _
                                 !intcvesocio & ",'" & vlstrSerieAnterior & "'," & vlintCredencialAnterior & ",'" & !vchseriecredencial & "'," & _
                                 !intnumeroCredencial & "," & fstrFechaSQL(CStr(Date)) & ")"
              
              pEjecutaSentencia vlstrSentenciaSQL
               
         '      rsHistoricoCredencial.AddNew
'                    rsHistoricoCredencial!intcvesocio = !intcvesocio
'                    rsHistoricoCredencial!vchseriecredencialanterior = vlstrSerieAnterior
'                    rsHistoricoCredencial!intnumerocredencialanterior = vlintCredencialAnterior
'                    rsHistoricoCredencial!vchSeriecredencialactual = !vchserieCredencial
'                    rsHistoricoCredencial!intnumerocredencialactual = !intnumeroCredencial
'                    rsHistoricoCredencial!dtmFechacambiocredencial = Date
'                    rsHistoricoCredencial.Update
'                    rsHistoricoCredencial.Requery
            End If
            
            .Update
            .Requery
            If vlblnNuevoSocio Then
'            .MoveLast
            vllngCveSocio = frsRegresaRs("SELECT max(intcvesocio) from sosocio", adLockOptimistic).Fields(0)
'            .MoveFirst
            End If
            
        Else
            vllngCveSocio = !intcvesocio
            !VCHCLAVESOCIO = Trim(txtClaveUnica.Text)
            !vchregistrosbe = Trim(txtRegistroSBE.Text)
            !vchNombre = Trim(txtNombre.Text)
            !vchApellidoPaterno = Trim(txtApePaterno.Text)
            !vchApellidoMaterno = Trim(txtApeMaterno.Text)
            !vchRFC = Trim(mskRFC.Text)
            !vchCURP = Trim(txtCurp.Text)
            
            !VCHNOMBREEMERGENCIA = Trim(TxtNombreEmergencia.Text)
            !VCHTELEMERGENCIA = Trim(txtTelefonoEmergencia.Text)
            !VCHCOMENTARIOS = Trim(txtComentarios.Text)
            !CHRFACTORRH = Trim(txtFactorRH)
            !chrGruposanguineo = cboGrupoSanguineo.Text
            !vchPoblacion = Trim(txtPoblacionT.Text)
            
            If mskFechaNac = "  /  /    " Then
                !dtmFechaNacimiento = Null
            Else
                !dtmFechaNacimiento = CDate(mskFechaNac)
            End If
            !chrSexo = IIf(optSexo(0), "M", "F")
            !intCveEstadoCivil = cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex)
            !vchCorreoElectronico = Trim(txtCorreoElectronico.Text)
            !vchLuarNacimiento = Trim(txtLugarNac.Text)
            !vchLugarTrabajo = Trim(txtLugarTrabajo.Text)
            !vchProfesion = Trim(txtProfesion.Text)
            !bitExtranjero = IIf(chkBitExtranjero.Value, 1, 0)
            
            !chrTipoSocio = IIf(vlblnDependiente, "D", "T")
            !intCuota = cboCuotaSocio.ListIndex
            !intFormaPago = cboFormaPago.ListIndex
            !intDerechos = cboDerechos.ListIndex

'            If !chrHispanidad = IIf(optHispanidad(0), "ES", IIf(optHispanidad(1), "HE", IIf(optHispanidad(2), "NE", IIf(optHispanidad(3), "BE", IIf(optHispanidad(4), "CE", "VE"))))) Or IsNull(!chrHispanidad) Then
'                !chrHispanidad = IIf(optHispanidad(0), "ES", IIf(optHispanidad(1), "HE", IIf(optHispanidad(2), "NE", IIf(optHispanidad(3), "BE", IIf(optHispanidad(4), "CE", "VE")))))
'            Else
             If Not vlblnNuevoSocio Then
                pCambioHispanidad !intcvesocio, vlstrHispanidadAntes, IIf(optHispanidad(0), "ES", IIf(optHispanidad(1), "HE", IIf(optHispanidad(2), "NE", IIf(optHispanidad(3), "BE", IIf(optHispanidad(4), "CE", "VE")))))
             End If
                !chrHispanidad = IIf(optHispanidad(0), "ES", IIf(optHispanidad(1), "HE", IIf(optHispanidad(2), "NE", IIf(optHispanidad(3), "BE", IIf(optHispanidad(4), "CE", "VE")))))
           


            !bitAutorizadoPagoCaja = IIf(chkAutorizadoPagoCaja, 1, 0)
            
            If mskFechaIngreso = "  /  /    " Then
                !dtmfechaingreso = Null
            Else
                !dtmfechaingreso = CDate(mskFechaIngreso)
            End If
            
            If mskFechaBaja = "  /  /    " Then
                !dtmfechaBaja = Null
            Else
                !dtmfechaBaja = CDate(mskFechaBaja)
            End If
            
            If mskFechaUltimoPago = "  /  /    " Then
                !dtmfechaUltimoPago = Null
            Else
                !dtmfechaUltimoPago = CDate(mskFechaUltimoPago)
            End If
            
            If mskFechaEmisionCredencial = "  /  /    " Then
                !dtmFechaCredencial = Null
            Else
                !dtmFechaCredencial = CDate(mskFechaEmisionCredencial)
            End If
           
            '!numSaldoActual = txtSaldoActual.Text
            !vchObservaciones = Trim(txtObservaciones.Text)
            !BITSOLICITUDInscripcion = IIf(chkSolicitudInscripcion, 1, 0)
            !BITACTANACIMIENTOTITULAR = IIf(chkActaNac, 1, 0)
            !BITACTAMATRIMONIO = IIf(chkActaMatrimonio, 1, 0)
            !BITACTANACIMIENTODEPENDIENTE = IIf(chkActaNacDep, 1, 0)
            !BITSOLCITUDCAMBIOREGISTRO = IIf(chkSolicitudCambio, 1, 0)
            
            If Trim(txtRutaImagen.Text) <> "" Then
                stmImagen.Type = adTypeBinary
                '|  Se carga el archivo en el objeto stream, para agregarlo al recordset.
                stmImagen.Open
                stmImagen.LoadFromFile txtRutaImagen.Text
                '|  se llena el campo del recordset con el objeto stream.
                !blbFoto = stmImagen.Read
            Else
                !blbFoto = Null
                vgblnFotoExistente = False
            End If
            
            If !vchseriecredencial = Trim(txtSerie.Text) Or IsNull(!vchseriecredencial) Then
                vlstrSerieAnterior = IIf(IsNull(!vchseriecredencial), txtSerie.Text, !vchseriecredencial)
                !vchseriecredencial = Trim(txtSerie.Text)
                vlblnCambioCredencial = False
            Else
                vlstrSerieAnterior = !vchseriecredencial
                !vchseriecredencial = Trim(txtSerie.Text)
                vlblnCambioCredencial = True
            End If
            
            If txtCredencial = "" Then
                !intnumeroCredencial = Null
            Else
                If (!intnumeroCredencial = CInt(txtCredencial) And txtSerie.Text = vlstrSerieAnterior) Or IsNull(!intnumeroCredencial) Then
                    !intnumeroCredencial = CInt(txtCredencial)
                    vlblnCambioCredencial = False
                Else
                    vlintCredencialAnterior = !intnumeroCredencial
                    !intnumeroCredencial = CInt(txtCredencial)
                    vlblnCambioCredencial = True
                End If
            End If
            
            !vchAcreditacionHispanidad = Trim(txtAcreditacionHispanidad.Text)
            !intnumerocuentacontable = Trim(vlintCuentaContable)
            
            If vlblnCambioCredencial = True Then
             vlstrSentenciaSQL = "insert into SOHISTORICOCREDENCIAL values(null," & _
                                 !intcvesocio & ",'" & vlstrSerieAnterior & "'," & vlintCredencialAnterior & ",'" & !vchseriecredencial & "'," & _
                                 !intnumeroCredencial & "," & fstrFechaSQL(CStr(Date)) & ")"
              
              pEjecutaSentencia vlstrSentenciaSQL
            
'                    'rsHistoricoCredencial.Open
'                    rsHistoricoCredencial.AddNew
'                    rsHistoricoCredencial!intcvesocio = !intcvesocio
'                    rsHistoricoCredencial!vchseriecredencialanterior = vlstrSerieAnterior
'                    rsHistoricoCredencial!intnumerocredencialanterior = vlintCredencialAnterior
'                    rsHistoricoCredencial!vchSeriecredencialactual = !vchserieCredencial
'                    rsHistoricoCredencial!intnumerocredencialactual = !intnumeroCredencial
'                    rsHistoricoCredencial!dtmFechacambiocredencial = Date
'                    rsHistoricoCredencial.Update
'                    rsHistoricoCredencial.Requery
'                    'vlblnCambioCredencial = False
            End If
            
            .Update
            .Requery
            If vlblnNuevoSocio Then
'            .MoveLast
             vllngCveSocio = frsRegresaRs("SELECT max(intcvesocio) from sosocio", adLockOptimistic).Fields(0)
            '= fSigConsecutivo("intcvesocio", "sosocio")
'            .MoveFirst
            End If
        End If
    End With
    
    pInsertaDomicilios vllngCveSocio
    pInsertaTelefonos vllngCveSocio
    pGrabaDictamen vllngCveSocio, mskFechaDictamen.Text, txtObservacionesDictamen.Text
    pGrabaNumerarios vllngCveSocio
    pGrabaDependiente vllngCveSocio
    
    
    vlchrSerie = txtSerie.Text
   
    If vlblnSerieNueva Then
       vlstrSentenciaSQL = "insert into SoFolioCredencial values('" & txtSerie.Text & "'," & txtCredencial.Text & ")"
    Else
       If CInt(txtCredencial.Text) > vlintCredencialActual And vlintCredencialActual <> 0 Then
       vlstrSentenciaSQL = "update SoFolioCredencial set vchseriecredencial = '" & txtSerie.Text & "', intnumerocredencialactual = " & txtCredencial.Text & _
                           " where vchseriecredencial = '" & txtSerie.Text & "'"
        End If
    End If
    pEjecutaSentencia vlstrSentenciaSQL
    
'    With rsCredencial
'        If vlblnSerieNueva Then
'            .AddNew
'            !vchseriecredencial = txtSerie.Text
'            !intnumerocredencialactual = txtCredencial.Text
'            .Update
'            .Requery
'        Else
'            If CInt(txtCredencial.Text) > vlintCredencialActual And vlintCredencialActual <> 0 Then
'                For i = 1 To .RecordCount
'                    If txtSerie.Text = !vchseriecredencial Then
'                        !vchseriecredencial = txtSerie.Text
'                        !intnumerocredencialactual = txtCredencial.Text
'                        .Update
'                        .Requery
'                        Exit For
'                    End If
'                    .MoveNext
'                Next
'
'            End If
'        End If
'    End With
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
    '---Limpia variables y campos manualmente despues consulta al socio que se guardo
    
    vlblnSerieNueva = False
    txtNombreSocio1.Text = ""
    txtCveSocioNum1.Text = ""
    txtNombreSocio2.Text = ""
    txtCveSocioNum2.Text = ""
    pBuscaSocio vllngCveSocio
    pPonEstado stConsulta
    If cmdImprimir.Enabled Then
        cmdImprimir.SetFocus
    End If

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, Me.Name & ":cmdGrabarRegistro_Click")
End Sub

Private Sub cmdImprimir_Click()
     pImprime "P"
  
End Sub
Private Sub cmdQuitarSocio1_Click()
    txtNombreSocio1.Text = ""
    txtCveSocioNum1.Text = ""
End Sub
Private Sub cmdQuitarSocio2_Click()
    txtNombreSocio2.Text = ""
    txtCveSocioNum2.Text = ""
End Sub
Private Sub Form_Activate()
    If stEstado = stEspera Then
        mskFechaIngreso.Mask = ""
        mskFechaIngreso.Text = fdtmServerFecha
        mskFechaIngreso.Mask = "##/##/####"
        
        mskFechaIngresoD.Mask = ""
        mskFechaIngresoD.Text = fdtmServerFecha
        mskFechaIngresoD.Mask = "##/##/####"
    End If
    If vlblnCargaGrupoSanguineoUnaVEZ = False Then
       vlblnCargaGrupoSanguineoUnaVEZ = True
       pCargaGrupoSanguineo
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    Select Case KeyCode
        Case vbKeyReturn
             SendKeys vbTab
            'KeyCode = 0
             DoEvents
         Case vbKeyEscape
            
            If stEstado = stEspera Then
                KeyCode = 0
                Unload Me
            Else
                 '¿Desea abandonar la operación?
                    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        pInicializaForma True
                        pPonEstado stEspera
                        KeyCode = 0
                    Else
                        Exit Sub
                    End If
            End If
    End Select

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyDown"))
End Sub
Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    
    vlstrsql = "SELECT * FROM SOSOCIO where INTCVESOCIO = -1 ORDER BY INTCVESOCIO"
    Set rsSocio = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
'    vlstrsql = "SELECT * FROM SOFOLIOCREDENCIAL ORDER BY VCHSERIECREDENCIAL"
'    Set rsCredencial = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
'    vlstrsql = "SELECT * FROM SOHISTORICOCREDENCIAL ORDER BY INTCVESOCIO"
'    Set rsHistoricoCredencial = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
'    vlstrsql = "SELECT * FROM SOHISTORICOHISPANIDAD ORDER BY INTCVESOCIO"
'    Set rsHispanidad = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
'    vlstrsql = "SELECT * FROM SODICTAMENES ORDER BY INTCVESOCIO"
'    Set rsDictamenes = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
'    vlstrsql = "SELECT * FROM SOSOCIONUMERARIO ORDER BY INTCVESOCIO"
'    Set rsNumerarios = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
'    vlstrsql = "SELECT * FROM SOSOCIODEPENDIENTE ORDER BY INTCVESOCIO"
'    Set rsDependientes = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
'
'    vlstrsql = "SELECT * FROM SOSOCIO ORDER BY INTCVESOCIO"
'    Set rsSocioTitular = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)

    
   'pInstanciaReporte vgrptReporte, "rptImpresionSocio.rpt" 'el reporte se instanacia en la funcion pimprime
    
    vlblnCargaGrupoSanguineoUnaVEZ = False
    
    pLlenaEstadoCivil
    cboDerechos.Clear
    cboDerechos.AddItem "BAJA"
    cboDerechos.AddItem "ACTIVO"
    cboDerechos.AddItem "SUSPENDIDO"
    
    cboDerechosD.Clear
    cboDerechosD.AddItem "BAJA"
    cboDerechosD.AddItem "ACTIVO"
    cboDerechosD.AddItem "SUSPENDIDO"
    
    cboCuotaSocio.Clear
    cboCuotaSocio.AddItem "INDIVIDUAL"
    cboCuotaSocio.AddItem "FAMILIAR"
    
    cboFormaPago.Clear
    cboFormaPago.AddItem "SEMESTRAL"
    cboFormaPago.AddItem "ANUAL"
    cboFormaPago.AddItem "MENSUAL"
    
    
    fblnLlenaCiudadesCbo cboCiudad
    fblnLlenaCiudadesCbo cboCiudadD
        
    cboEstadoCivil.ListIndex = 0
    vlStrEstructuraRFC = mskRFC.Mask
    
    pConfiguraGridCredenciales
    pConfiguraGridHispanidad
    pConfiguraGridDependientes

    vlblnAsignaTitular = False
    
    pConfiguraDependiente
    pPonEstado stEspera
    
    
    
    
    
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub txtComentarios_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub txtComentarios_GotFocus()
pSelTextBox txtComentarios
End Sub
Private Sub txtComentarios_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
    
End Sub
Private Sub txtFactorRH_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub txtFactorRH_GotFocus()
pSelTextBox txtFactorRH
End Sub
Private Sub txtFactorRH_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 43 And Not KeyAscii = 45 And Not KeyAscii = vbKeyReturn And Not KeyAscii = vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub txtLugarNacD_Change()
   If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtLugarNacD_GotFocus()
  pSelTextBox txtLugarNacD
End Sub
Private Sub txtLugarNacD_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TxtNombreEmergencia_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub TxtNombreEmergencia_GotFocus()
pSelTextBox TxtNombreEmergencia
End Sub
Private Sub TxtNombreEmergencia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNumeroExterior_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
    
    If Trim(txtNumeroExterior.Text) = "" Then
        txtNumeroInterior.Text = ""
        txtNumeroInterior.Enabled = False
    Else
        txtNumeroInterior.Enabled = True
    End If
 End Sub

Private Sub txtNumeroExterior_GotFocus()
    pSelTextBox txtNumeroExterior
End Sub

Private Sub txtNumeroExterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumeroExteriorD_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
    
    If Trim(txtNumeroExteriorD.Text) = "" Then
        txtNumeroInteriorD.Text = ""
        txtNumeroInteriorD.Enabled = False
    Else
        txtNumeroInteriorD.Enabled = True
    End If
End Sub

Private Sub txtNumeroExteriorD_GotFocus()
pSelTextBox txtNumeroExteriorD
End Sub

Private Sub txtNumeroExteriorD_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumeroInterior_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub mskFechaBaja_GotFocus()
    pSelMkTexto mskFechaBaja
End Sub

Private Sub mskFechaBaja_LostFocus()
    If Not IsDate(mskFechaBaja.Text) And Not mskFechaBaja.Text = "  /  /    " Then
'        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
'        pEnfocaMkTexto mskFechaBaja
'        Exit Sub
    Else
        If Not mskFechaBaja.Text = "  /  /    " Then
            If (CDate(mskFechaBaja.Text) > fdtmServerFecha) Or (CDate(mskFechaBaja.Text) < CDate("01/01/1900")) Then
                MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
                mskFechaBaja.SetFocus
            Else
                If CDate(mskFechaBaja.Text) < CDate(mskFechaIngreso.Text) Then
                    MsgBox SIHOMsg(379), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaBaja.SetFocus
                End If
            End If
        End If
    End If
        
    
End Sub

Private Sub mskFechaBajaD_Change()
If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub mskFechaBajaD_GotFocus()
    pSelMkTexto mskFechaBajaD
End Sub

Private Sub mskFechaBajaD_LostFocus()
    If Not IsDate(mskFechaBajaD.Text) And Not mskFechaBajaD.Text = "  /  /    " Then
'        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
'        pEnfocaMkTexto mskFechaBajaD
'        Exit Sub
    Else
        If Not mskFechaBajaD.Text = "  /  /    " Then
            If (mskFechaBajaD.Text > fdtmServerFecha) Or (mskFechaBajaD.Text < CDate("01/01/1900")) Then
                MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
                mskFechaBajaD.SetFocus
            Else
                If CDate(mskFechaBajaD.Text) < CDate(mskFechaIngresoD.Text) Then
                    MsgBox SIHOMsg(379), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaBajaD.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub mskFechaDictamen_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub mskFechaDictamen_GotFocus()
    pSelMkTexto mskFechaDictamen
End Sub

Private Sub mskFechaDictamen_LostFocus()
    If Not IsDate(mskFechaDictamen.Text) And Not mskFechaBaja.Text = "  /  /    " Then
        'MsgBox SIHOMsg(29), vbCritical, "Mensaje"
       ' pEnfocaMkTexto mskFechaDictamen
        'Exit Sub
    Else
        If Not mskFechaDictamen.Text = "  /  /    " Then
            If IsDate(mskFechaDictamen.Text) Then
                If (CDate(mskFechaDictamen.Text) > fdtmServerFecha) Or (CDate(mskFechaDictamen.Text) < CDate("01/01/1900")) Then
                    MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
                    pEnfocaMkTexto mskFechaDictamen
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub mskFechaEmisionCredencial_GotFocus()
    pSelMkTexto mskFechaEmisionCredencial
End Sub

Private Sub mskFechaEmisionCredencial_LostFocus()
    If Not IsDate(mskFechaEmisionCredencial.Text) And Not mskFechaEmisionCredencial.Text = "  /  /    " Then
'        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
'        pEnfocaMkTexto mskFechaEmisionCredencial
'        Exit Sub
    Else
        If Not mskFechaEmisionCredencial.Text = "  /  /    " Then
            If (CDate(mskFechaEmisionCredencial.Text) > fdtmServerFecha) Or (CDate(mskFechaEmisionCredencial.Text) < CDate("01/01/1900")) Then
                MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
                mskFechaEmisionCredencial.SetFocus
'            Else
'                If CDate(mskFechaEmisionCredencial.Text) < CDate(mskFechaIngreso.Text) Then
'                    MsgBox SIHOMsg(379), vbOKOnly + vbInformation, "Mensaje"
'                    mskFechaEmisionCredencial.SetFocus
'                End If
            End If
        End If
    End If
End Sub

Private Sub mskFechaEmisionCredencialD_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub mskFechaEmisionCredencialD_GotFocus()
    pSelMkTexto mskFechaEmisionCredencialD
End Sub

Private Sub mskFechaEmisionCredencialD_LostFocus()
     If Not IsDate(mskFechaEmisionCredencialD.Text) And Not mskFechaEmisionCredencialD.Text = "  /  /    " Then
'        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
'        pEnfocaMkTexto mskFechaEmisionCredencialD
'        Exit Sub
    Else
        If Not mskFechaEmisionCredencial.Text = "  /  /    " Then
            
            If (CDate(mskFechaEmisionCredencial.Text) > CDate(fdtmServerFecha)) Or (CDate(mskFechaEmisionCredencial.Text) < CDate("01/01/1900")) Then
                MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
                mskFechaEmisionCredencialD.SetFocus
            Else
                If mskFechaEmisionCredencial.Text < mskFechaIngresoD.Text Then
                    MsgBox SIHOMsg(379), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaEmisionCredencialD.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub mskFechaIngreso_GotFocus()
      pSelMkTexto mskFechaIngreso
End Sub

Private Sub mskFechaIngreso_LostFocus()
    If Not IsDate(mskFechaIngreso.Text) Then
'        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
'        pEnfocaMkTexto mskFechaIngreso
'        Exit Sub
    Else
        If (CDate(mskFechaIngreso.Text) > fdtmServerFecha) Or (CDate(mskFechaIngreso.Text) < CDate("01/01/1900")) Then
            MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
            mskFechaIngreso.SetFocus
        End If
    End If
End Sub

Private Sub mskFechaIngresoD_GotFocus()
    pSelMkTexto mskFechaIngresoD
End Sub

Private Sub mskFechaIngresoD_LostFocus()
    If Not IsDate(mskFechaIngresoD.Text) Then
'        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
'        pEnfocaMkTexto mskFechaIngresoD
'        Exit Sub
    Else
        If (CDate(mskFechaIngresoD.Text) > fdtmServerFecha) Or (CDate(mskFechaIngresoD.Text) < CDate("01/01/1900")) Then
            MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
            mskFechaIngresoD.SetFocus
        End If
    End If
End Sub

Private Sub mskFechaNac_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub mskFechaNac_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskFechaNac
    

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaNac_GotFocus"))
End Sub

Private Sub mskFechaNac_LostFocus()
On Error GoTo NotificaError
    If IsDate(mskFechaNac.Text) Then
        If (CDate(mskFechaNac.Text) > fdtmServerFecha) Or (CDate(mskFechaNac.Text) < CDate("01/01/1900")) Then
            MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
            mskFechaNac.SetFocus
            Exit Sub
        End If
        txtEdad.Text = fstrObtieneEdad(CDate(mskFechaNac.Text), fdtmServerFecha)
        If Trim(txtApePaterno.Text) <> "" And Trim(txtApeMaterno.Text) <> "" And Trim(txtNombre.Text) <> "" And IsDate(mskFechaNac.Text) Then
            mskRFC.Mask = ""
            mskRFC.Text = fstrRFC(Trim(txtApePaterno.Text), Trim(txtApeMaterno.Text), Trim(txtNombre.Text), Trim(mskFechaNac.Text))
            mskRFC.SetFocus
            mskRFC.SelStart = Len(mskRFC)
            mskRFC.MaxLength = 13
            
        End If
        pGeneraCLaveUnica
    Else
        'MsgBox SIHOMsg(29), vbCritical, "Mensaje"
        'pEnfocaMkTexto mskFechaNac
        'txtEdad.Text = ""
    End If
       
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaNac_LostFocus"))
    Unload Me
End Sub

Private Sub mskFechaUltimoPago_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub mskFechaUltimoPago_GotFocus()
    pSelMkTexto mskFechaUltimoPago
End Sub

Private Sub mskFechaUltimoPago_LostFocus()
    If Not IsDate(mskFechaUltimoPago.Text) And Not mskFechaUltimoPago.Text = "  /  /    " Then
'        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
'        pEnfocaMkTexto mskFechaUltimoPago
'        Exit Sub
    Else
        If Not mskFechaUltimoPago.Text = "  /  /    " Then
            If (CDate(mskFechaUltimoPago.Text) > fdtmServerFecha) Or (CDate(mskFechaUltimoPago.Text) < CDate("01/01/1900")) Then
                MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
                mskFechaUltimoPago.SetFocus
'            Else
'                If CDate(mskFechaUltimoPago.Text) < CDate(mskFechaIngreso.Text) Then
'                    MsgBox SIHOMsg(379), vbOKOnly + vbInformation, "Mensaje"
'                    mskFechaUltimoPago.SetFocus
'                End If
            End If
        End If
    End If
        
'    If CDate(mskFechaUltimoPago.Text) < CDate(mskFechaIngreso.Text) Then
'        MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
'        mskFechaUltimoPago.SetFocus
'    End If
End Sub

Private Sub mskRFC_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub mskRFC_GotFocus()
    pSelMkTexto mskRFC
End Sub

Private Sub mskRFC_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub optHispanidad_Click(Index As Integer)
    If stEstado = stNuevo Or stEstado = stedicion Then
        pGeneraCLaveUnica
    End If
End Sub

Private Sub optHispanidad_GotFocus(Index As Integer)
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub optSexo_GotFocus(Index As Integer)
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub optTipoSocio_Click(Index As Integer)
    If optTipoSocio(0).Value Then
        optHispanidad(3).Enabled = False
        optHispanidad(4).Enabled = False
        
        optHispanidad(5).Enabled = True
    Else
        optHispanidad(3).Enabled = True
        optHispanidad(4).Enabled = True
        
        optHispanidad(5).Enabled = False
        optHispanidad(11).Enabled = False ' agrego caso 6882
    End If
End Sub

Private Sub pInicializaForma(blnTodos As Boolean)
    Dim vlControl As Control
    Dim vlintContador As Integer

    On Error GoTo NotificaError
    
    
    If vlblnDependiente Then
        sstOpcion.Tab = 4
    Else
        sstOpcion.Tab = 0
    End If
    '| Limpia los controles tipo TextBox y ComboBox de la forma
    For Each vlControl In frmSocios.Controls
        If TypeOf vlControl Is TextBox Then
            vlControl.Text = ""
        ElseIf TypeOf vlControl Is ComboBox Then
            vlControl.ListIndex = -1
        End If
    Next
    
    cboEstadoCivil.ListIndex = 0
    cboDerechos.ListIndex = 1
    cboDerechosD.ListIndex = 1
    cboCuotaSocio.ListIndex = 0
    cboFormaPago.ListIndex = 0
    
    If cboGrupoSanguineo.ListCount > 0 Then
       cboGrupoSanguineo.ListIndex = 0
    End If
    txtNumeroInterior.Enabled = False
       txtNumeroInteriorD.Enabled = False
    
    pMkTextAsignaValor mskFechaNac, ""
    pMkTextAsignaValor mskFechaIngreso, ""
    pMkTextAsignaValor mskFechaBaja, ""
    pMkTextAsignaValor mskFechaUltimoPago, ""
    pMkTextAsignaValor mskFechaEmisionCredencial, ""
    pMkTextAsignaValor mskFechaDictamen, ""
    pMkTextAsignaValor mskFechaIngresoD, ""
    pMkTextAsignaValor mskFechaBajaD, ""
    pMkTextAsignaValor mskFechaEmisionCredencialD, ""
    mskRFC.Mask = ""
    mskRFC.Text = ""
    mskRFC.Mask = vlStrEstructuraRFC
    mskFechaIngreso.Mask = ""
    mskFechaIngreso.Text = fdtmServerFecha
    mskFechaIngreso.Mask = "##/##/####"
    mskFechaIngresoD.Mask = ""
    mskFechaIngresoD.Text = fdtmServerFecha
    mskFechaIngresoD.Mask = "##/##/####"
    chkAutorizadoPagoCaja.Value = 0
    chkSolicitudInscripcion.Value = 0
    chkActaNac.Value = 0
    chkActaMatrimonio.Value = 0
    chkActaNacDep.Value = 0
    chkSolicitudCambio.Value = 0
    optSexo(0).Value = True
    optTipoSocio(0).Value = True
    If vlblnDependiente Then
    optHispanidad(6).Value = True
       phabilitacambioshispanidad 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0 ' se activan los opt de hispanidad segun si es la ventana de dependientes o titulares
    Else
    optHispanidad(0).Value = True
       phabilitacambioshispanidad 1, 1, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0
    End If
    pEnfocaTextBox txtSerie
    
    With grdhCredenciales
        .Clear
        .Cols = 4
        .Rows = 2
    End With
    
    With grdhHispanidad
        .Clear
        .Cols = 4
        .Rows = 2
    End With
    
    With grdhDependientes
        .Clear
        .Cols = 4
        .Rows = 2
    End With
    
    pConfiguraGridCredenciales
    pConfiguraGridHispanidad
    pConfiguraGridDependientes
    
    If vlstrRutaImagen <> "" Then
        Kill vlstrRutaImagen
        vlstrRutaImagen = ""
    End If
    
   
    mskFechaBajaD.Enabled = False
    mskFechaBaja.Enabled = False
    
    vlblnSerieNueva = False
    
        
    If vlblnDependiente Then
     cboCiudadD.ListIndex = flngLocalizaCbo(cboCiudadD, Str(vgintCveCiudadCH))
    Else
     cboCiudad.ListIndex = flngLocalizaCbo(cboCiudad, Str(vgintCveCiudadCH))
    End If
   
   
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pInicializaForma"))
End Sub
'-----------------------------------------------------------------------------------------
'|  Esta subrutina habilita/deshabilita los componentes del mantenimiento para que se
'|  realicen las operaciones de Espera, Edicion, Consulta, Nuevo de la manera correcta
'-----------------------------------------------------------------------------------------
Public Sub pPonEstado(enmStatus As enmStatus)
    Select Case enmStatus
        Case stEspera
            '|  Inicializa los componentes de la forma
            pInicializaForma True
            '|  Habilita botonera según la acción a ejecutar
          
            cmdBuscar.Enabled = True
            cmdGrabarRegistro.Enabled = False
            cmdImprimir.Enabled = False
            cmdCambioSocio.Enabled = True
            frmSociosBusqueda.vlblnCambioTipoSocio = False
            
            Set picImagen.Picture = LoadPicture("")
            stEstado = enmStatus
            vlblnNuevoSocio = False
            
        Case stConsulta
            cmdBuscar.Enabled = True
            cmdGrabarRegistro.Enabled = False
            cmdImprimir.Enabled = True
            cmdCambioSocio.Enabled = False
            pEnfocaTextBox txtSerie
            stEstado = enmStatus
            vlblnNuevoSocio = False
            
        Case stNuevo
            '|  Inicializa los componentes de la forma
            pInicializaForma False
            '|  Habilita botonera segun la acción a ejecutar
            cmdBuscar.Enabled = False
            cmdGrabarRegistro.Enabled = True
            cmdImprimir.Enabled = False
            cmdCambioSocio.Enabled = False
            pEnfocaTextBox txtSerie
            Set picImagen.Picture = LoadPicture("")
            stEstado = enmStatus
            vlblnNuevoSocio = True
            
        Case stedicion
            '|  Habilita botonera segun la acción a ejecutar
            cmdBuscar.Enabled = False
            cmdGrabarRegistro.Enabled = True
            cmdImprimir.Enabled = False
            cmdCambioSocio.Enabled = False
            stEstado = enmStatus
            vlblnNuevoSocio = False
            
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPonEstado"))
End Sub

Private Sub txtAcreditacionHispanidad_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtAcreditacionHispanidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sstOpcion.Tab = 3
        mskFechaDictamen.SetFocus
    End If
End Sub

Private Sub txtAcreditacionHispanidad_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtApeMaterno_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtApeMaterno_GotFocus()
    pSelTextBox txtApeMaterno
End Sub

Private Sub txtApeMaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtApePaterno_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtApePaterno_GotFocus()
    pSelTextBox txtApePaterno
End Sub

Private Sub txtApePaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAportacionesEA_KeyPress(KeyAscii As Integer)
     If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7

End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
' ESTE PROCESO SE QUITO YA QUE SE USABA EN LAS FLECHAS DE BUSQUEDA QUE FUERON DESACTIVADAS----------------------------------------------------
'Private Sub pConsultaSocio()
'    Dim stmImagen As New ADODB.Stream
'    Dim vlstrSentencia As String
'    Dim rsCredito As New ADODB.Recordset
'    Dim vlStrEstadoActual As String
'    Dim rsDomicilio As New ADODB.Recordset
'    Dim rsTelefono As New ADODB.Recordset
'    Dim vlstrTemp As String
'    Dim rsCredencialesTemp As New ADODB.Recordset
'    Dim rsCuentaTemp As New ADODB.Recordset
'    Dim rsHispanidadTemp As New ADODB.Recordset
'
'
'    On Error GoTo NotificaError
'
'    vgblnFotoExistente = False
'
'    If rsSocio.EOF Then
'        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
'        txtSerie.SetFocus
'        Exit Sub
'    End If
'
'    pInicializaForma True
'
'    With rsSocio
'
'        txtSerie.Text = fstrIsNull(!vchseriecredencial)
'        txtCredencial.Text = fintIsNull(!intnumerocredencial)
'        pCargaCuentaContable !intnumerocuentacontable
'        txtRegistroSBE.Text = fstrIsNull(!vchregistrosbe)
'        txtNombre.Text = fstrIsNull(!vchNombre)
'        txtApeMaterno.Text = fstrIsNull(!vchApellidoMaterno)
'        txtApePaterno.Text = fstrIsNull(!vchApellidoPaterno)
'        mskRFC.Mask = ""
'        mskRFC.Text = fstrIsNull(!vchRFC)
'        txtCurp.Text = fstrIsNull(!vchCURP)
'        chkBitExtranjero.Value = fintIsNull(!bitExtranjero)
'
'        If Not IsNull(!dtmFechaNacimiento) Then
'            pMkTextAsignaValor mskFechaNac, !dtmFechaNacimiento
'        Else
'            pMkTextAsignaValor mskFechaNac, ""
'        End If
'
'        If !chrSexo = "M" Then
'            optSexo(0).Value = True
'        Else
'            optSexo(1).Value = True
'        End If
'
'        ' Comprobaciones de Fechas en caso de que tengan valor Null
'        If Not vlblnDependiente Then
'
'            If Not IsNull(!dtmFechaCredencial) Then
'                pMkTextAsignaValor mskFechaEmisionCredencial, !dtmFechaCredencial
'            Else
'                pMkTextAsignaValor mskFechaEmisionCredencial, ""
'            End If
'
'            If Not IsNull(!dtmfechaingreso) Then
'                pMkTextAsignaValor mskFechaIngreso, !dtmfechaingreso
'            Else
'                pMkTextAsignaValor mskFechaIngreso, ""
'            End If
'
'            If Not IsNull(!dtmfechaUltimoPago) Then
'                pMkTextAsignaValor mskFechaUltimoPago, !dtmfechaUltimoPago
'            Else
'                pMkTextAsignaValor mskFechaUltimoPago, ""
'            End If
'
'            If Not IsNull(!dtmfechaBaja) Then
'                pMkTextAsignaValor mskFechaBaja, !dtmfechaBaja
'            Else
'                pMkTextAsignaValor mskFechaBaja, ""
'            End If
'        Else
'            If Not IsNull(!dtmFechaCredencial) Then
'                pMkTextAsignaValor mskFechaEmisionCredencialD, !dtmFechaCredencial
'            Else
'                pMkTextAsignaValor mskFechaEmisionCredencialD, ""
'            End If
'
'            If Not IsNull(!dtmfechaingreso) Then
'                pMkTextAsignaValor mskFechaIngresoD, !dtmfechaingreso
'            Else
'                pMkTextAsignaValor mskFechaIngresoD, ""
'            End If
'
'            If Not IsNull(!dtmfechaBaja) Then
'                pMkTextAsignaValor mskFechaBajaD, !dtmfechaBaja
'            Else
'                pMkTextAsignaValor mskFechaBajaD, ""
'            End If
'        End If
'
'        cboEstadoCivil.ListIndex = fintLocalizaCbo(cboEstadoCivil, fintIsNull(!intCveEstadoCivil, 0))
'        txtCurp.Text = fstrIsNull(!vchCURP)
'        txtCorreoElectronico.Text = fstrIsNull(!VCHCORREOELECTRONICO)
'
'        If vlblnDependiente Then
'            txtPoblacionD.Text = fstrIsNull(!vchPoblacion)
'        End If
'
'        txtLugarNac.Text = fstrIsNull(!vchLuarNacimiento)
'        txtLugarTrabajo.Text = fstrIsNull(!vchLugarTrabajo)
'        txtProfesion.Text = fstrIsNull(!vchProfesion)
'
'        ' Asigna el tipo de socio
'        If !chrTipoSocio = "T" Then
'            optTipoSocio(0) = True
'        Else
'            optTipoSocio(1) = True
'        End If
'
'        ' Asigna Hispanidad
'        If vlblnDependiente = False Then
'            Select Case !chrHispanidad
'                Case "ES"
'                    opthispanidad(0) = True
'                Case "HE"
'                    opthispanidad(1) = True
'                Case "NE"
'                    opthispanidad(2) = True
'                Case "BE"
'                    opthispanidad(3) = True
'                Case "CE"
'                    opthispanidad(4) = True
'               ' Case "VE"
'                    'optHispanidad(5) = True
'            End Select
'        Else
'            Select Case !chrHispanidad
'                Case "ES"
'                    opthispanidad(6) = True
'                Case "HE"
'                    opthispanidad(7) = True
'                Case "NE"
'                    opthispanidad(8) = True
'                'Case "BE"
'                    'optHispanidad(9) = True
'                'Case "CE"
'                  '  optHispanidad(10) = True
'                Case "VE"
'                    opthispanidad(11) = True
'            End Select
'        End If
'
'        cboCuotaSocio.ListIndex = fintIsNull(!intCuota, 0)
'        cboFormaPago.ListIndex = fintIsNull(!intFormaPago, 0)
'
'        If vlblnDependiente = False Then
'            cboDerechos.ListIndex = fintIsNull(!intDerechos, 0)
'        Else
'            cboDerechosD.ListIndex = fintIsNull(!intDerechos, 0)
'        End If
'
'        txtObservaciones = fstrIsNull(!vchobservaciones)
'
'        chkAutorizadoPagoCaja = IIf(IsNull(!bitAutorizadoPagoCaja), 0, !bitAutorizadoPagoCaja)
'        chkSolicitudInscripcion = IIf(IsNull(!BITSOLICITUDInscripcion), 0, !BITSOLICITUDInscripcion)
'        chkActaNac = IIf(IsNull(!BITACTANACIMIENTOTITULAR), 0, !BITACTANACIMIENTOTITULAR)
'        chkActaMatrimonio = IIf(IsNull(!BITACTAMATRIMONIO), 0, !BITACTAMATRIMONIO)
'        chkActaNacDep = IIf(IsNull(!BITACTANACIMIENTODEPENDIENTE), 0, !BITACTANACIMIENTODEPENDIENTE)
'        chkSolicitudCambio = IIf(IsNull(!BITSOLCITUDCAMBIOREGISTRO), 0, !BITSOLCITUDCAMBIOREGISTRO)
'        txtClaveUnica.Text = fstrIsNull(!VCHCLAVESOCIO)
'
'        vlstrHispanidadAntes = Mid(txtClaveUnica.Text, 1, 2) ' esta variable obtiene la
'        'clave de hispanidad de la consulta realizada
'
'        ' Cargamos la Imagen del Socio
'        If Not IsNull(!blbFoto) Then
'            stmImagen.Type = adTypeBinary
'            stmImagen.Open
'            stmImagen.Write !blbFoto
'            stmImagen.SaveToFile App.Path & "\" & txtClaveUnica.Text, adSaveCreateOverWrite
'            ' Retorna la imagen a la función
'            Set picImagen.Picture = LoadPicture(App.Path & "\" & txtClaveUnica.Text, vbLPLarge, vbLPColor)
'            txtRutaImagen.Text = App.Path & "\" & txtClaveUnica.Text
'            vlstrRutaImagen = txtRutaImagen.Text
'            vgblnFotoExistente = True
'            stmImagen.Close
'        Else
'             vgblnFotoExistente = False
'             Set picImagen.Picture = LoadPicture("", vbLPLarge, vbLPColor)
'        End If
'
'        vllngCveSocio = !intcvesocio
'
'        End With
'
'
'
'        pCargaDomicilioTelefono vllngCveSocio
'        pCargaGridCredencial vllngCveSocio
'        pCargaGridHispanidad vllngCveSocio
'        pCargaDictamen vllngCveSocio
'        pCargaNumerarios vllngCveSocio
'        pCargaGridDependientes vllngCveSocio
'        pFormatoFechaGrid
'
'        If vlblnDependiente Then
'            pCargaTitular vllngCveSocio
'        End If
'
'
'Exit Sub
'NotificaError:
'    EntornoSIHO.ConeccionSIHO.RollbackTrans
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConsultaSocio"))
'End Sub--------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub pGeneraCLaveUnica()
    Dim vlstrClaveUnica As String
    Dim vlintPunto As Integer
    Dim i As Integer
    On Error GoTo NotificaError
    If Not vlblnDependiente Then 'optiTpoSocio(0).Value Then
        vlstrClaveUnica = "T"
    Else
        vlstrClaveUnica = "D"
    End If
    
 ' se cambiaron IF then else por elseif caso 6882
    If vlblnDependiente Then
        If optHispanidad(6).Value Then
               vlstrClaveUnica = vlstrClaveUnica & "E-"
        ElseIf optHispanidad(7).Value Then
               vlstrClaveUnica = vlstrClaveUnica & "H-"
        ElseIf optHispanidad(8).Value Then
               vlstrClaveUnica = vlstrClaveUnica & "N-"
        ElseIf optHispanidad(9).Value Then
               vlstrClaveUnica = vlstrClaveUnica & "B-"
        ElseIf optHispanidad(10).Value Then
               vlstrClaveUnica = vlstrClaveUnica & "C-"
       'se comento en caso 6882
       'ElseIf opthispanidad(11).Value Then
              'vlstrClaveUnica = vlstrClaveUnica & "V-"
        End If
    Else
        If optHispanidad(0).Value Then
            vlstrClaveUnica = "E" & vlstrClaveUnica & "-"
        ElseIf optHispanidad(1).Value Then
            vlstrClaveUnica = "H" & vlstrClaveUnica & "-"
        ElseIf optHispanidad(2).Value Then
            vlstrClaveUnica = "N" & vlstrClaveUnica & "-"
       'se comento en caso 6882
       'ElseIf opthispanidad(3).Value Then
           'vlstrClaveUnica = "B" & vlstrClaveUnica & "-"
       'ElseIf opthispanidad(4).Value Then
           'vlstrClaveUnica = "C" & vlstrClaveUnica & "-"
        ElseIf optHispanidad(5).Value Then
            vlstrClaveUnica = "V" & vlstrClaveUnica & "-"
        End If
    End If
    
    For i = 1 To Len(txtClaveContabilidad.Text)
        If vlintPunto = 2 Then
            vlstrClaveUnica = vlstrClaveUnica & Mid(txtClaveContabilidad.Text, (i), 4) & "-00-"
            Exit For
        End If
        
        If Mid(txtClaveContabilidad, i, 1) = "." Then
            vlintPunto = vlintPunto + 1
        End If
    Next
    
    Trim (mskFechaNac)
   'se comento en caso 6882
   ' vlstrClaveUnica = vlstrClaveUnica & Right(mskFechaNac, 2)
   ' vlstrClaveUnica = vlstrClaveUnica & Mid(mskFechaNac, 4, 2)
   ' vlstrClaveUnica = vlstrClaveUnica & Left(mskFechaNac, 2)
   ' txtClaveUnica.Text = vlstrClaveUnica
   'se comento y se cambio por las siguientes lineas caso 6882
   vlstrClaveUnica = vlstrClaveUnica & Right(mskFechaNac.FormattedText, 2)
    vlstrClaveUnica = vlstrClaveUnica & fregresames(Mid(mskFechaNac.FormattedText, 4, 3))
    vlstrClaveUnica = vlstrClaveUnica & Left(mskFechaNac.FormattedText, 2)
    txtClaveUnica.Text = vlstrClaveUnica
    
   
    If Not vlblnDependiente Then
        txtNombreDictamen.Text = Trim(txtApePaterno.Text) & " " & Trim(txtApeMaterno.Text) & " " & Trim(txtNombre.Text)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGeneraCLaveUnica"))
End Sub
Private Function fregresames(mes As String) As String ' regresa el numero del mes que le mandes, esto ayuda a formar la clave unica con la fecha ya formateada
Select Case LCase(mes)
Case "ene"
fregresames = "01"
Case "feb"
fregresames = "02"
Case "mar"
fregresames = "03"
Case "abr"
fregresames = "04"
Case "may"
fregresames = "05"
Case "jun"
fregresames = "06"
Case "jul"
fregresames = "07"
Case "ago"
fregresames = "08"
Case "sep"
fregresames = "09"
Case "oct"
fregresames = "10"
Case "nov"
fregresames = "11"
Case "dic"
fregresames = "12"
Case Else
fregresames = ""
End Select
End Function
Public Sub pEdadRFC()

        If mskRFC.Text = "" Then
            If Trim(txtApePaterno.Text) <> "" And Trim(txtApeMaterno.Text) <> "" And Trim(txtNombre.Text) <> "" And IsDate(mskFechaNac.Text) Then
                mskRFC.Mask = ""
                mskRFC.Text = fstrRFC(Trim(txtApePaterno.Text), Trim(txtApeMaterno.Text), Trim(txtNombre.Text), Trim(mskFechaNac.Text))
            End If
        End If
        
        If IsDate(mskFechaNac.Text) Then
            If (mskFechaNac.Text > fdtmServerFecha) Or (mskFechaNac.Text < CDate("01/01/1900")) Then
                MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
                mskFechaNac.SetFocus
            End If
            txtEdad.Text = fstrObtieneEdad(CDate(mskFechaNac.Text), fdtmServerFecha)
        Else
            txtEdad.Text = ""
        End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEdadRFC"))
End Sub
Public Sub pGrabaSocio()
    'On Error GoTo NotificaError
    On Error GoTo NotificaError
    
    Dim stmImagen As New ADODB.Stream
    'Dim vlintContador As Integer
    'Dim intRow As Integer
    'Dim vgstrParametrosSP As String

    'If Not fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 1347, 2078), "E") Or Not fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 1347, 2078), "C") Then
        '|  ¡El usuario no tiene permiso para grabar datos!
        'MsgBox SIHOMsg(65), vbCritical, "Mensaje"
        'Exit Sub
    'End If
    
    If Not fblnDatosCorrectos() Then Exit Sub
       
    'vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
      
    'If vllngPersonaGraba = 0 Then Exit Sub
   ' pGeneraCLaveUnica
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    With rsSocio
    
        If stEstado = stNuevo Then
            .AddNew
        End If
            !VCHCLAVESOCIO = Trim(txtClaveUnica.Text)
            'registro SBE
            !vchNombre = Trim(txtNombre.Text)
            !vchApellidoPaterno = Trim(txtApePaterno.Text)
            !vchApellidoMaterno = Trim(txtApeMaterno.Text)
            !dtmFechaNacimiento = CDate(mskFechaNac)
            !chrSexo = IIf(optSexo(0), "M", "f")
            !intCveEstadoCivil = cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex)
            !vchCorreoElectronico = Trim(txtCorreoElectronico.Text)
            !vchLuarNacimiento = Trim(txtLugarNac.Text)
            !vchLugarTrabajo = Trim(txtLugarTrabajo.Text)
            !chrTipoSocio = IIf(optTipoSocio(0), "T", "D")
            '!intDerechos = cboDerechos.ItemData(cboDerechos.ListIndex)
            !chrHispanidad = IIf(optHispanidad(0), "ES", IIf(optHispanidad(1), "HE", IIf(optHispanidad(2), "NE", IIf(optHispanidad(3), "BE", IIf(optHispanidad(4), "CE", "VE")))))
            !bitAutorizadoPagoCaja = IIf(chkAutorizadoPagoCaja, 1, 0)
            !dtmfechaingreso = CDate(mskFechaIngreso)
            !dtmfechaBaja = CDate(mskFechaBaja)
            !dtmfechaUltimoPago = CDate(mskFechaUltimoPago)
            '!numSaldoActual = txtSaldoActual.Text
            !vchObservaciones = Trim(txtObservaciones.Text)
            !BITSOLICITUDInscripcion = IIf(chkSolicitudInscripcion, 1, 0)
            !BITACTANACIMIENTOTITULAR = IIf(chkActaNac, 1, 0)
            !BITACTAMATRIMONIO = IIf(chkActaMatrimonio, 1, 0)
            !BITACTANACIMIENTODEPENDIENTE = IIf(chkActaNacDep, 1, 0)
            !BITSOLCITUDCAMBIOREGISTRO = IIf(chkSolicitudCambio, 1, 0)
            
            If Trim(txtRutaImagen.Text) <> "" Then
                stmImagen.Type = adTypeBinary
                '|  Se carga el archivo en el objeto stream, para agregarlo al recordset.
                stmImagen.Open
                stmImagen.LoadFromFile txtRutaImagen.Text
                '|  se llena el campo del recordset con el objeto stream.
                !blbFoto = stmImagen.Read
            Else
                !blbFoto = Null
                vgblnFotoExistente = False
            End If
            
            !vchseriecredencial = Trim(txtSerie.Text)
            !intnumeroCredencial = txtCredencial.Text
            !dtmFechaCredencial = CDate(mskFechaEmisionCredencial)
            !intCuota = cboCuotaSocio.ItemData(cboCuotaSocio.ListIndex)
            !intFormaPago = cboFormaPago.ItemData(cboFormaPago.ListIndex)
            !vchAcreditacionHispanidad = Trim(txtAcreditacionHispanidad.Text)
            !intnumerocuentacontable = Trim(txtClaveContabilidad.Text)
            
            .Update
            MsgBox "Registro Grabado", vbCritical, "Mensaje"
            
    End With
    
    EntornoSIHO.ConeccionSIHO.CommitTrans

    pPonEstado stEspera
    vllngEstatusForma = stEspera
    pEnfocaTextBox txtSerie

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, Me.Name & ":cmdGrabarRegistro_Click")
End Sub

Private Sub txtClaveContabilidad_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
        pGeneraCLaveUnica
    End If
    If stEstado = stNuevo Then
        pGeneraCLaveUnica
    End If
End Sub

Public Sub txtClaveContabilidad_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vlstrTemp As String
Dim rsCuentaTemp As New ADODB.Recordset

    If KeyCode = vbKeyReturn Then
        If frmBusquedaCuentasContables.vllngNumeroCuenta <> 0 Then
            pCargaCuentaContable frmBusquedaCuentasContables.vllngNumeroCuenta
            frmBusquedaCuentasContables.vllngNumeroCuenta = 0

        Else
            frmBusquedaCuentasContables.Show vbModal, Me
            If frmBusquedaCuentasContables.vllngNumeroCuenta <> 0 Then
                pCargaCuentaContable frmBusquedaCuentasContables.vllngNumeroCuenta
                frmBusquedaCuentasContables.vllngNumeroCuenta = 0
            End If
        End If
    End If
End Sub

Private Sub txtClaveUnica_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
    If Trim(txtClaveUnica.Text) <> "" Then
        vlchrClaveUnica = txtClaveUnica.Text
    End If
End Sub

Private Sub txtClaveUnica_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtClaveUnica_Validate(Cancel As Boolean)
    If Trim(txtClaveUnica.Text) = "" Then
        txtClaveUnica.Text = vlchrClaveUnica
    End If
End Sub

Private Sub txtColonia_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtColonia_GotFocus()
    pSelTextBox txtColonia
End Sub

Private Sub txtColonia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtColoniaD_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtColoniaD_GotFocus()
    pSelTextBox txtColoniaD
End Sub

Private Sub txtColoniaD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCorreoElectronico_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtCorreoElectronico_GotFocus()
    pSelTextBox txtCorreoElectronico
End Sub

Private Sub txtCorreoElectronico_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCP_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtCP_GotFocus()
    pSelTextBox txtCP
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtCPD_Change()
        If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtCPD_GotFocus()
    pSelTextBox txtCPD
End Sub

Private Sub txtCPD_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtCPT_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtCPT_GotFocus()
    pSelTextBox txtCPT
End Sub

Private Sub txtCPT_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtCredencial_KeyPress(KeyAscii As Integer)
     If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtCurp_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtCurp_GotFocus()
    pSelTextBox txtCurp
End Sub

Private Sub txtCurp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDomicilio_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtDomicilio_GotFocus()
    pSelTextBox txtDomicilio
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDomicilioD_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtDomicilioD_GotFocus()
    pSelTextBox txtDomicilioD
End Sub

Private Sub txtDomicilioD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDomicilioT_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtDomicilioT_GotFocus()
    pSelTextBox txtDomicilioT
End Sub

Private Sub txtDomicilioT_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFax_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtFax_GotFocus()
    pSelTextBox txtFax
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtFaxD_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtFaxD_GotFocus()
    pSelTextBox txtFaxD
End Sub

Private Sub txtFaxD_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtLugarNac_Change()
If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtLugarNac_GotFocus()
    pSelTextBox txtLugarNac
End Sub

Private Sub txtLugarNac_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLugarTrabajo_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtLugarTrabajo_GotFocus()
    pSelTextBox txtLugarTrabajo
End Sub

Private Sub txtLugarTrabajo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombre_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtNombre_GotFocus()
    pSelTextBox txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombreDictamen_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombreSocio1_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtNombreSocio2_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtNumeroInterior_GotFocus()
    pSelTextBox txtNumeroInterior
End Sub

Private Sub txtNumeroInterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNumeroInteriorD_Change()
   If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub txtNumeroInteriorD_GotFocus()
    pSelTextBox txtNumeroInteriorD
End Sub
Private Sub txtNumeroInteriorD_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtObservaciones_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sstOpcion.Tab = 2
        cmdCambiarSocio1.SetFocus
    End If
End Sub
Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtObservacionesDictamen_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub
Private Sub txtObservacionesDictamen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If stEstado = stedicion Then
                cmdGrabarRegistro.SetFocus
        End If
    End If
End Sub
Private Sub txtObservacionesDictamen_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtPoblacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPoblacionD_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtPoblacionD_GotFocus()
    pSelTextBox txtPoblacionD
End Sub

Private Sub txtPoblacionD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPoblacionT_Change()
If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtPoblacionT_GotFocus()
pSelTextBox txtPoblacionT
End Sub

Private Sub txtPoblacionT_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtProfesion_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtProfesion_GotFocus()
    pSelTextBox txtProfesion
End Sub

Private Sub txtProfesion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sstOpcion.Tab = 1
        '|  Pone el foco en la hispanidad que se encuentre seleccionada
        If optHispanidad(0).Value Then
           optHispanidad(0).SetFocus
        ElseIf optHispanidad(1).Value Then
               optHispanidad(1).SetFocus
        ElseIf optHispanidad(2).Value Then
               optHispanidad(2).SetFocus
        ElseIf optHispanidad(3).Value Then
               optHispanidad(3).SetFocus
        ElseIf optHispanidad(4).Value Then
               optHispanidad(4).SetFocus
        ElseIf optHispanidad(5).Value Then
               optHispanidad(5).SetFocus
        ElseIf optHispanidad(6).Value Then
               optHispanidad(6).SetFocus
        End If
  End If
End Sub

Private Sub txtProfesion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRegistroSBE_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then 'Or stEstado = stedicion Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtRegistroSBE_GotFocus()
    If vlblnDependiente And stEstado = stNuevo Then ' And Trim(txtRegistroSBE.Text) = "" Then
        If Trim(txtRegistroSBE.Text) = "" Then
        txtApePaterno.SetFocus
            txtRegistroSBE_KeyDown 13, 0
        End If
    Else
    pSelTextBox txtRegistroSBE
    End If
End Sub

Private Sub txtRegistroSBE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And vlblnDependiente Then
        vlblnAsignaTitular = True
        
        
        frmSociosBusqueda.vgblnAsignaTitular = vlblnAsignaTitular
        frmSociosBusqueda.Show vbModal, Me
        
        If Not frmSociosBusqueda.vgblnEscape Then
           txtRegistroSBE.Text = frmSociosBusqueda.vgstrNombreSocio '| Nombre del socio
           txtCveSocioNum1.Text = frmSociosBusqueda.vglngClaveSocio '| Clave del socio
           vllngTitular = frmSociosBusqueda.vglngClaveSocio
           frmBusquedaCuentasContables.vllngNumeroCuenta = frmSociosBusqueda.vglngNumeroCuenta
           txtApePaterno.SetFocus
         Else
           'txtRegistroSBE.SetFocus
         End If
            
        Unload frmSociosBusqueda
        vlblnAsignaTitular = False
       
        If frmBusquedaCuentasContables.vllngNumeroCuenta <> 0 And Trim(txtRegistroSBE.Text) <> "" Then
            pCargaCuentaContable frmBusquedaCuentasContables.vllngNumeroCuenta
            frmBusquedaCuentasContables.vllngNumeroCuenta = 0
        Else
            If Not vlblnDependiente Then
                frmBusquedaCuentasContables.Show vbModal, Me
                If frmBusquedaCuentasContables.vllngNumeroCuenta <> 0 Then
                    pCargaCuentaContable frmBusquedaCuentasContables.vllngNumeroCuenta
                    frmBusquedaCuentasContables.vllngNumeroCuenta = 0
                End If
            End If
        End If
        txtApePaterno.SetFocus
    Else
        vlblnAsignaTitular = False
    End If
End Sub

Private Sub txtRegistroSBE_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSaldoActual_KeyPress(KeyAscii As Integer)
 If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim vlintCredencial
    
    If KeyCode = vbKeyReturn Then
            If Trim(txtSerie.Text) = "" Then
                With frmSociosBusqueda
                       .vgblnAsignaTitular = vlblnAsignaTitular
                       .vgblnDependiente = vlblnDependiente
                       .Show vbModal, Me
                    If Not .vgblnEscape Then
                        pLlenaInformacionSocio
                    End If
                End With
                Unload frmSociosBusqueda
                txtSerie.SetFocus
                Exit Sub
'        cmdBuscar.Enabled = True
'        cmdCambioSocio.Enabled = True
'        With frmSociosBusqueda
'                .vgblnAsignaTitular = vlblnAsignaTitular
'                .vgblnDependiente = vlblnDependiente
'                .Show vbModal, Me
'                pLlenaInformacionSocio
'            End With
'            txtSerie.SetFocus
'            Exit Sub
        End If
        vlchrSerie = Trim(txtSerie.Text)
    
        If Not stEstado = stConsulta Then
            If Not stEstado = stedicion Then
                pPonEstado stNuevo
            Else
                pPonEstado stedicion
                vllngEstatusForma = stedicion
                vlblnCambioSerie = True
            End If
        Else
            pPonEstado stedicion
                vllngEstatusForma = stedicion
                vlblnCambioSerie = True
        End If

        txtSerie.Text = vlchrSerie
        txtCredencial.SetFocus
        txtCredencial.Locked = True
            
        
     vlstrSentenciaSQL = "Select * from SoFolioCredencial where vchseriecredencial = '" & vlchrSerie & "'"
     Set rsCredencial = frsRegresaRs(vlstrSentenciaSQL, adLockOptimistic)
     If rsCredencial.RecordCount = 0 Then ' no encontro la credencial
        txtCredencial.Text = 1
        vlblnSerieNueva = True
        vlintCredencialActual = 0
     Else ' si encontro la serie de la credencial
'       If rsSocio.RecordCount > 0 Then ' si hay un registro cargado
'        'comparamos la series
'         If rsSocio!vchseriecredencial <> vlchrSerie Then ' si son diferentes entonces si se debe hacer el cambio de serie
'            txtCredencial.Text = rsCredencial!intnumerocredencialactual
'            vlintCredencialActual = rsCredencial!intnumerocredencialactual
'            txtCredencial.Text = (CInt(txtCredencial.Text) + 1)
'            vlblnSerieNueva = False
''         End If
'       End If
       If stEstado = stNuevo Then  ' solamente si esta en estado de nuevo, pero y si quiere cambiar la serie de la credencial???
          txtCredencial.Text = rsCredencial!intnumerocredencialactual
          vlintCredencialActual = rsCredencial!intnumerocredencialactual
          txtCredencial.Text = (CInt(txtCredencial.Text) + 1)
          vlblnSerieNueva = False
       ElseIf stEstado = stedicion Then
           If rsSocio.RecordCount > 0 Then
              If rsSocio!vchseriecredencial <> vlchrSerie Then
                 txtCredencial.Text = rsCredencial!intnumerocredencialactual
                vlintCredencialActual = rsCredencial!intnumerocredencialactual
                 txtCredencial.Text = (CInt(txtCredencial.Text) + 1)
                vlblnSerieNueva = False
              End If
           End If
       End If
     End If
'    vlstrsql = "SELECT * FROM SOFOLIOCREDENCIAL ORDER BY VCHSERIECREDENCIAL"
'    Set rsCredencial = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
'
'        With rsCredencial
'           If rsCredencial.EOF Then
'                txtCredencial.Text = 1
'                vlblnSerieNueva = True
'            Else
'                .MoveFirst
'                For i = 1 To .RecordCount
'                    If vlchrSerie = !vchseriecredencial Then
'                        txtCredencial.Text = !intnumerocredencialactual
'                        vlintCredencialActual = !intnumerocredencialactual
'                        txtCredencial.Text = (CInt(txtCredencial.Text) + 1)
'                        vlblnSerieNueva = False
'                        Exit For
'                    End If
'                    .MoveNext
'                    txtCredencial.Text = 1
'                    vlblnSerieNueva = True
'                Next
'                .MoveFirst
'            End If
'            End With
    End If
End Sub


Private Sub pLlenaInformacionSocio()
    With frmSociosBusqueda
        '----------------------------------------------------------------------------------
        '------ Este código se realizaba al seleccionar un socio en la forma de búsqueda
        '----------------------------------------------------------------------------------
        '|  No esta vacío el nombre del socio
        If Not .vgstrNombreSocio = "" Then
            If vlblnAsignaTitular Then
                frmBusquedaCuentasContables.vllngNumeroCuenta = .grdhBuscaSocios.TextMatrix(.grdhBuscaSocios.Row, 7) '| Cuenta contable del socio
                vllngTitular = .vglngClaveSocio  '| Clave del socio
                txtRegistroSBE.Text = .vgstrNombreSocio '| Nombre del socio
                vlblnAsignaTitular = False
            Else
                pBuscaSocio .vglngClaveSocio  '| Clave del socio
                pPonEstado stConsulta
            End If
        End If
        Unload frmSociosBusqueda
    End With
    '--------------------------------------------------------------------------------
    vllngEstatusForma = stConsulta

End Sub


Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSerie_Validate(Cancel As Boolean)
    
    If Not stEstado = stConsulta Then
        If txtSerie.Text = "" And stEstado = stedicion Then
            txtSerie.Text = rsSocio!vchseriecredencial
        End If
        If txtSerie.Text = "" And stEstado = stNuevo Then
            txtSerie.Text = vlchrSerie
        End If
        If txtSerie.Text <> "" Then
            If stEstado <> stEspera And stEstado <> stNuevo Then
                If txtSerie = rsSocio!vchseriecredencial Then
                     txtCredencial.Text = rsSocio!intnumeroCredencial
                Else
                    txtSerie_KeyDown 13, 0
                End If
            Else
                If stEstado = stEspera Then
                txtSerie_KeyDown 13, 0
                Else
                    If txtSerie.Text <> vlchrSerie Then
                        txtSerie_KeyDown 13, 0
                    End If
                End If
            End If
        End If
    Else
        If txtSerie.Text = "" And stEstado <> stEspera Then
            txtSerie.Text = rsSocio!vchseriecredencial
            Exit Sub
        End If
        If txtSerie.Text <> rsSocio!vchseriecredencial Then
            txtSerie_KeyDown 13, 0
        Else
            txtCredencial.Text = rsSocio!intnumeroCredencial
        End If
    End If
End Sub

Private Sub txtTelefono_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtTelefono_GotFocus()
    pSelTextBox txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub pValidaEdicion()
    If stEstado = stConsulta Then
    pPonEstado stedicion
    vllngEstatusForma = stedicion
    End If
End Sub
                
Private Sub pInsertaTelefonos(lngCveSocio As Long)
    On Error GoTo NotificaError
'    Dim intContador As Integer 'Contador
    Dim lngCveTelefono As Long 'Consecutivo de los teléfonos
    Dim rsTelefonos As New ADODB.Recordset 'Teléfonos del paciente
    'Tipos de teléfono: 1 = local, 2 = foráneo
    
' Se eliminó el ciclo para ver si se corrije el problema de los duplicados de teléfonos.
'    For intContador = 1 To 2
    
    
        '*********** LOCAL SOCIO ***********
        
        vgstrParametrosSP = Str(lngCveSocio) & "|" & "3"
        
        Set rsTelefonos = frsEjecuta_SP(vgstrParametrosSP, "SP_SOSELTELEFONOSOC")
        If vlblnDependiente Then
            vgstrParametrosSP = "3|" & " " _
            & "|" & Trim(txtTelefonoD.Text)
            
            If rsTelefonos.RecordCount = 0 Then
                'Alta del teléfono:
                vgstrParametrosSP = vgstrParametrosSP & "|0"
            Else
                'Modificación:
                vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsTelefonos!intCveTelefono)
            End If
        Else
            vgstrParametrosSP = "3|" & " " _
            & "|" & Trim(txtTelefono.Text)
            
            If rsTelefonos.RecordCount = 0 Then
                'Alta del teléfono:
                vgstrParametrosSP = vgstrParametrosSP & "|0"
            Else
                'Modificación:
                vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsTelefonos!intCveTelefono)
            End If
        End If
        lngCveTelefono = 1
        frsEjecuta_SP vgstrParametrosSP, "SP_GNINSTELEFONO", True, lngCveTelefono
            
        If rsTelefonos.RecordCount = 0 Then
            'Dar de alta el teléfono relacionándolo con el paciente:
            vgstrParametrosSP = Str(lngCveSocio) & "|" & Str(lngCveTelefono)
            frsEjecuta_SP vgstrParametrosSP, "SP_SOINSSOCIOTELEFONO"
        End If
        
        
        'vgstrParametrosSP = Str(lngCveSocio) & "|" & "3"
        
        'Set rsTelefonos = frsEjecuta_SP(vgstrParametrosSP, "SP_SOSELTELEFONOSOC")
        
        'vgstrParametrosSP = "3|" & " " _
        '& "|" & Trim(txtTelefono.Text)
        
        'If rsTelefonos.RecordCount = 0 Then
            'Alta del teléfono:
            'vgstrParametrosSP = vgstrParametrosSP & "|0"
        'Else
            'Modificación:
            'vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsTelefonos!intCveTelefono)
        'End If
        
        'lngCveTelefono = 1
        'frsEjecuta_SP vgstrParametrosSP, "SP_GNINSTELEFONO", True, lngCveTelefono
            
        'If rsTelefonos.RecordCount = 0 Then
            'Dar de alta el teléfono relacionándolo con el paciente:
            'vgstrParametrosSP = Str(lngCveSocio) & "|" & Str(lngCveTelefono)
            'frsEjecuta_SP vgstrParametrosSP, "SP_SOINSSOCIOTELEFONO"
        'End If
        
        '*********** TRABAJO SOCIO ***********
        vgstrParametrosSP = Str(lngCveSocio) & "|" & "4"
        
        Set rsTelefonos = frsEjecuta_SP(vgstrParametrosSP, "SP_SOSELTELEFONOSOC")
        
        vgstrParametrosSP = "4|" & " " _
        & "|" & Trim(txtTelefonoT.Text)
        
        If rsTelefonos.RecordCount = 0 Then
            'Alta del teléfono:
            vgstrParametrosSP = vgstrParametrosSP & "|0"
        Else
            'Modificación:
            vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsTelefonos!intCveTelefono)
        End If
        
        lngCveTelefono = 1
        frsEjecuta_SP vgstrParametrosSP, "SP_GNINSTELEFONO", True, lngCveTelefono
            
        If rsTelefonos.RecordCount = 0 Then
            'Dar de alta el teléfono relacionándolo con el paciente:
            vgstrParametrosSP = Str(lngCveSocio) & "|" & Str(lngCveTelefono)
            frsEjecuta_SP vgstrParametrosSP, "SP_SOINSSOCIOTELEFONO"
        End If
        
        '*********** FAX SOCIO ***********
'        intContador = 1 'Local
        vgstrParametrosSP = Str(lngCveSocio) & "|" & "5"
        
        Set rsTelefonos = frsEjecuta_SP(vgstrParametrosSP, "SP_SOSELTELEFONOSOC")
        If vlblnDependiente Then
            vgstrParametrosSP = "5|" & " " _
            & "|" & Trim(txtFaxD.Text)
            
            If rsTelefonos.RecordCount = 0 Then
                'Alta del teléfono:
                vgstrParametrosSP = vgstrParametrosSP & "|0"
            Else
                'Modificación:
                vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsTelefonos!intCveTelefono)
            End If
        Else
            vgstrParametrosSP = "5|" & " " _
            & "|" & Trim(txtFax.Text)
            
            If rsTelefonos.RecordCount = 0 Then
                'Alta del teléfono:
                vgstrParametrosSP = vgstrParametrosSP & "|0"
            Else
                'Modificación:
                vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsTelefonos!intCveTelefono)
            End If
        End If
        lngCveTelefono = 1
        frsEjecuta_SP vgstrParametrosSP, "SP_GNINSTELEFONO", True, lngCveTelefono
            
        If rsTelefonos.RecordCount = 0 Then
            'Dar de alta el teléfono relacionándolo con el paciente:
            vgstrParametrosSP = Str(lngCveSocio) & "|" & Str(lngCveTelefono)
            frsEjecuta_SP vgstrParametrosSP, "SP_SOINSSOCIOTELEFONO"
        End If

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pInsertaTelefonos"))
    Unload Me
End Sub
Private Sub pConfiguraGridCredenciales()
    On Error GoTo NotificaError
    
    'Configura el grid de la búsqueda de Socios
    Dim vlintseq As Integer
    
    grdhCredenciales.Enabled = True
    With grdhCredenciales
    
        .FormatString = "NA|De|A|Fecha"
        .ColWidth(0) = 100
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 1300
        .ScrollBars = flexScrollBarBoth
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCredenciales"))
End Sub

Private Sub pInsertaDomicilios(lngCveSocio)
    On Error GoTo NotificaError

    Dim lngCveDomicilio As Long 'Consecutivo en domicilios
    Dim rsDomicilios As New ADODB.Recordset 'Domicilios del socio
    Dim vlstrciudad As String
        
    vgstrParametrosSP = Str(lngCveSocio) & "|3"
            
    Set rsDomicilios = frsEjecuta_SP(vgstrParametrosSP, "SP_SOSELDOMICILIOSOCIO")
            If vlblnDependiente Then
                vgstrParametrosSP = _
                "3" _
                & "|" & cboCiudadD.ItemData(cboCiudadD.ListIndex) _
                & "|" & Trim(txtDomicilioD.Text) _
                & "|" & Trim(txtNumeroExteriorD.Text) _
                & "|" & Trim(txtNumeroInteriorD.Text) _
                & "|" & Trim(txtColoniaD.Text) _
                & "|" & Trim(txtCPD.Text)
                
                If rsDomicilios.RecordCount = 0 Then
                    'Alta del domicilio:
                    vgstrParametrosSP = vgstrParametrosSP & "|0"
                Else
                    'Modificación:
                    vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsDomicilios!intcveDomicilio)
                End If
            Else
        
            
                vgstrParametrosSP = _
                "3" _
                & "|" & cboCiudad.ItemData(cboCiudad.ListIndex) _
                & "|" & Trim(txtDomicilio.Text) _
                & "|" & Trim(txtNumeroExterior.Text) _
                & "|" & Trim(txtNumeroInterior.Text) _
                & "|" & Trim(txtColonia.Text) _
                & "|" & Trim(txtCP.Text)
                
                If rsDomicilios.RecordCount = 0 Then
                    'Alta del domicilio:
                    vgstrParametrosSP = vgstrParametrosSP & "|0"
                Else
                    'Modificación:
                    vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsDomicilios!intcveDomicilio)
                End If
            End If
            
            lngCveDomicilio = 1
            frsEjecuta_SP vgstrParametrosSP, "SP_GNINSDOMICILIO", True, lngCveDomicilio
                
            If rsDomicilios.RecordCount = 0 Then
                'Dar de alta el domicilio relacionándolo con el paciente:
                vgstrParametrosSP = Str(lngCveSocio) & "|" & Str(lngCveDomicilio)
                frsEjecuta_SP vgstrParametrosSP, "SP_SOINSSOCIODOMICILIO"
            End If
            
            ' Guardar domicilio del trabajo
            
            vgstrParametrosSP = Str(lngCveSocio) & "|4" '
            
            Set rsDomicilios = frsEjecuta_SP(vgstrParametrosSP, "SP_SOSELDOMICILIOSOCIO")
            
            vgstrParametrosSP = _
            "4" _
            & "|" _
            & "|" & Trim(txtDomicilioT.Text) _
            & "|" _
            & "|" _
            & "|" _
            & "|" & Trim(txtCPT.Text)
            
            If rsDomicilios.RecordCount = 0 Then
                'Alta del domicilio:
                vgstrParametrosSP = vgstrParametrosSP & "|0"
            Else
                'Modificación:
                vgstrParametrosSP = vgstrParametrosSP & "|" & Str(rsDomicilios!intcveDomicilio)
            End If
            
            lngCveDomicilio = 1
            frsEjecuta_SP vgstrParametrosSP, "SP_GNINSDOMICILIO", True, lngCveDomicilio
                
            If rsDomicilios.RecordCount = 0 Then
                'Dar de alta el domicilio relacionándolo con el paciente:
                vgstrParametrosSP = Str(lngCveSocio) & "|" & Str(lngCveDomicilio)
                frsEjecuta_SP vgstrParametrosSP, "SP_SOINSSOCIODOMICILIO"
            End If

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pInsertaDomicilios"))
    Unload Me
End Sub

Public Sub pBuscaSocio(lngClaveSocio As Long)

    Dim stmImagen As New ADODB.Stream
    Dim rsDomicilio As New ADODB.Recordset
    Dim rsTelefono As New ADODB.Recordset
    Dim vlstrTemp As String
    Dim rsCredencialesTemp As New ADODB.Recordset
    
    On Error GoTo NotificaError
    vlstrSentenciaSQL = "Select * from sosocio where intcvesocio = " & lngClaveSocio
    Set rsSocio = frsRegresaRs(vlstrSentenciaSQL, adLockOptimistic, adOpenDynamic)
    With rsSocio

     If rsSocio.RecordCount > 0 Then
        If vlblnDependiente Then
            txtSerie.Text = fstrIsNull(!vchseriecredencial)
            txtCredencial.Text = fintIsNull(!intnumeroCredencial)
            
            pCargaCuentaContable !intnumerocuentacontable
            txtRegistroSBE.Text = fstrIsNull(!vchregistrosbe)
            txtNombre.Text = fstrIsNull(!vchNombre)
            txtApeMaterno.Text = fstrIsNull(!vchApellidoMaterno)
            txtApePaterno.Text = fstrIsNull(!vchApellidoPaterno)
            mskRFC.Mask = ""
            mskRFC.Text = fstrIsNull(!vchRFC)

            txtCurp.Text = fstrIsNull(!vchCURP)
            chkBitExtranjero.Value = fintIsNull(!bitExtranjero)
            
            TxtNombreEmergencia.Text = fstrIsNull(Trim(!VCHNOMBREEMERGENCIA))
            txtTelefonoEmergencia.Text = fstrIsNull(Trim(!VCHTELEMERGENCIA))
            txtComentarios.Text = fstrIsNull(Trim(!VCHCOMENTARIOS))
            txtFactorRH = fstrIsNull(Trim(!CHRFACTORRH))
            cboGrupoSanguineo.ListIndex = fintLocalizaCritCbo(cboGrupoSanguineo, IIf(IsNull(!chrGruposanguineo), " ", Trim(!chrGruposanguineo)))
            txtLugarNacD.Text = fstrIsNull(Trim(!vchLuarNacimiento))
                     
            
            If Not IsNull(!dtmFechaNacimiento) Then
                pMkTextAsignaValor mskFechaNac, !dtmFechaNacimiento
            Else
                pMkTextAsignaValor mskFechaNac, ""
            End If
            
            If !chrSexo = "M" Then
                optSexo(0).Value = True
            Else
                optSexo(1).Value = True
            End If
            
            ' Comprobaciones de Fechas en caso de que tengan valor Null
            If Not IsNull(!dtmFechaCredencial) Then
                pMkTextAsignaValor mskFechaEmisionCredencialD, !dtmFechaCredencial
            Else
                pMkTextAsignaValor mskFechaEmisionCredencialD, ""
            End If
            
            If Not IsNull(!dtmfechaingreso) Then
                pMkTextAsignaValor mskFechaIngresoD, !dtmfechaingreso
            Else
                pMkTextAsignaValor mskFechaIngresoD, ""
            End If
            
            If Not IsNull(!dtmfechaBaja) Then
                pMkTextAsignaValor mskFechaBajaD, !dtmfechaBaja
            Else
                pMkTextAsignaValor mskFechaBajaD, ""
            End If
            
            cboEstadoCivil.ListIndex = fintLocalizaCbo(cboEstadoCivil, fintIsNull(!intCveEstadoCivil, 0))
            txtCurp.Text = fstrIsNull(!vchCURP)
            txtCorreoElectronico.Text = fstrIsNull(!vchCorreoElectronico)
            txtPoblacionD.Text = fstrIsNull(!vchPoblacion)
            cboDerechosD.ListIndex = fintIsNull(!intDerechos, 0)
            
            ' Asigna Hispanidad
            Select Case !chrHispanidad
                Case "ES"
                    optHispanidad(6) = True
                Case "HE"
                    optHispanidad(7) = True
                Case "NE"
                    optHispanidad(8) = True
                Case "BE"
                    optHispanidad(9) = True
                Case "CE"
                    optHispanidad(10) = True
                Case "VE"
                    optHispanidad(11) = True
            End Select
            txtClaveUnica.Text = fstrIsNull(!VCHCLAVESOCIO)
            
            'se asignan los posibles cambios de hispanidad que se pueden presentar
            'segun la hispanidad del socio
            vlstrHispanidadAntes = Mid(fstrIsNull(!VCHCLAVESOCIO), 1, 2)
            pPosiblesCambiosHispanidad (vlstrHispanidadAntes)
           
            ' Cargamos la Imagen del Socio
            If Not IsNull(!blbFoto) Then
                stmImagen.Type = adTypeBinary
                stmImagen.Open
                stmImagen.Write !blbFoto
                stmImagen.SaveToFile App.Path & "\" & txtClaveUnica.Text, adSaveCreateOverWrite
                ' Retorna la imagen a la función
                Set picImagen.Picture = LoadPicture(App.Path & "\" & txtClaveUnica.Text, vbLPLarge, vbLPColor)
                txtRutaImagen.Text = App.Path & "\" & txtClaveUnica.Text
                vlstrRutaImagen = txtRutaImagen.Text
                vgblnFotoExistente = True
                stmImagen.Close
            Else
                 vgblnFotoExistente = False
                 Set picImagen.Picture = LoadPicture("", vbLPLarge, vbLPColor)
            End If
            
            vllngCveSocio = !intcvesocio
            
                pCargaDomicilioTelefono vllngCveSocio
                pCargaTitular vllngCveSocio
                pEdadRFC
                pCargaSocioTit
                'Exit Do
        Else
            txtSerie.Text = fstrIsNull(!vchseriecredencial)
            txtCredencial.Text = fintIsNull(!intnumeroCredencial)
            
            pCargaCuentaContable !intnumerocuentacontable
            txtRegistroSBE.Text = fstrIsNull(!vchregistrosbe)
            txtNombre.Text = fstrIsNull(!vchNombre)
            txtApeMaterno.Text = fstrIsNull(!vchApellidoMaterno)
            txtApePaterno.Text = fstrIsNull(!vchApellidoPaterno)
            mskRFC.Mask = ""
            'mskRFC.Mask = "????######AAA"
            mskRFC.Text = fstrIsNull(!vchRFC)
            txtCurp.Text = fstrIsNull(!vchCURP)
            chkBitExtranjero.Value = fintIsNull(!bitExtranjero)
            
            TxtNombreEmergencia.Text = fstrIsNull(Trim(!VCHNOMBREEMERGENCIA))
            txtTelefonoEmergencia.Text = fstrIsNull(Trim(!VCHTELEMERGENCIA))
            txtComentarios.Text = fstrIsNull(Trim(!VCHCOMENTARIOS))
            txtFactorRH = fstrIsNull(Trim(!CHRFACTORRH))
            cboGrupoSanguineo.ListIndex = fintLocalizaCritCbo(cboGrupoSanguineo, IIf(IsNull(!chrGruposanguineo), " ", Trim(!chrGruposanguineo)))
            txtPoblacionT.Text = fstrIsNull(!vchPoblacion)
            
            If Not IsNull(!dtmFechaNacimiento) Then
                pMkTextAsignaValor mskFechaNac, !dtmFechaNacimiento
            Else
                pMkTextAsignaValor mskFechaNac, ""
            End If
            
            If !chrSexo = "M" Then
                optSexo(0).Value = True
            Else
                optSexo(1).Value = True
            End If
            
            ' Comprobaciones de Fechas en caso de que tengan valor Null
            If Not IsNull(!dtmFechaCredencial) Then
                pMkTextAsignaValor mskFechaEmisionCredencial, !dtmFechaCredencial
            Else
                pMkTextAsignaValor mskFechaEmisionCredencial, ""
            End If
            
            If Not IsNull(!dtmfechaingreso) Then
                pMkTextAsignaValor mskFechaIngreso, !dtmfechaingreso
            Else
                pMkTextAsignaValor mskFechaIngreso, ""
            End If
            
            If Not IsNull(!dtmfechaUltimoPago) Then
                pMkTextAsignaValor mskFechaUltimoPago, !dtmfechaUltimoPago
            Else
                pMkTextAsignaValor mskFechaUltimoPago, ""
            End If
            
            If Not IsNull(!dtmfechaBaja) Then
                pMkTextAsignaValor mskFechaBaja, !dtmfechaBaja
            Else
                pMkTextAsignaValor mskFechaBaja, ""
            End If
            
            cboEstadoCivil.ListIndex = fintLocalizaCbo(cboEstadoCivil, fintIsNull(!intCveEstadoCivil, 0))
            txtCurp.Text = fstrIsNull(!vchCURP)
            txtCorreoElectronico.Text = fstrIsNull(!vchCorreoElectronico)
            txtLugarNac.Text = fstrIsNull(!vchLuarNacimiento)
            txtLugarTrabajo.Text = fstrIsNull(!vchLugarTrabajo)
            txtProfesion.Text = fstrIsNull(!vchProfesion)
            
            ' Asigna el tipo de socio
            If !chrTipoSocio = "T" Then
                optTipoSocio(0) = True
            Else
                optTipoSocio(1) = True
            End If
            
            ' Asigna Hispanidad
            Select Case !chrHispanidad
                Case "ES"
                    optHispanidad(0) = True
                Case "HE"
                    optHispanidad(1) = True
                Case "NE"
                    optHispanidad(2) = True
                Case "BE"
                    optHispanidad(3) = True
                Case "CE"
                    optHispanidad(4) = True
                Case "VE"
                    optHispanidad(5) = True
                Case Else
                    optHispanidad(0) = True
            End Select
            
         
            txtAcreditacionHispanidad.Text = fstrIsNull(!vchAcreditacionHispanidad)
            cboCuotaSocio.ListIndex = fintIsNull(!intCuota, 0)
            cboFormaPago.ListIndex = fintIsNull(!intFormaPago, 0)
            cboDerechos.ListIndex = fintIsNull(!intDerechos, 0)
            chkAutorizadoPagoCaja = IIf(IsNull(!bitAutorizadoPagoCaja), 0, !bitAutorizadoPagoCaja)
            txtObservaciones = fstrIsNull(!vchObservaciones)
            chkSolicitudInscripcion = IIf(IsNull(!BITSOLICITUDInscripcion), 0, !BITSOLICITUDInscripcion)
            chkActaNac = IIf(IsNull(!BITACTANACIMIENTOTITULAR), 0, !BITACTANACIMIENTOTITULAR)
            chkActaMatrimonio = IIf(IsNull(!BITACTAMATRIMONIO), 0, !BITACTAMATRIMONIO)
            chkActaNacDep = IIf(IsNull(!BITACTANACIMIENTODEPENDIENTE), 0, !BITACTANACIMIENTODEPENDIENTE)
            chkSolicitudCambio = IIf(IsNull(!BITSOLCITUDCAMBIOREGISTRO), 0, !BITSOLCITUDCAMBIOREGISTRO)
            txtClaveUnica.Text = fstrIsNull(!VCHCLAVESOCIO)
            
           'se asignan los posibles cambios de hispanidad que se pueden presentar
            'segun la hispanidad del socio
            vlstrHispanidadAntes = Mid(fstrIsNull(!VCHCLAVESOCIO), 1, 2)
            pPosiblesCambiosHispanidad (vlstrHispanidadAntes)
            
            ' Cargamos la Imagen del Socio
            If Not IsNull(!blbFoto) Then
                stmImagen.Type = adTypeBinary
                stmImagen.Open
                stmImagen.Write !blbFoto
                stmImagen.SaveToFile App.Path & "\" & txtClaveUnica.Text, adSaveCreateOverWrite
                ' Retorna la imagen a la función
                Set picImagen.Picture = LoadPicture(App.Path & "\" & txtClaveUnica.Text, vbLPLarge, vbLPColor)
                txtRutaImagen.Text = App.Path & "\" & txtClaveUnica.Text
                vlstrRutaImagen = txtRutaImagen.Text
                vgblnFotoExistente = True
                stmImagen.Close
            Else
                 vgblnFotoExistente = False
                 Set picImagen.Picture = LoadPicture("", vbLPLarge, vbLPColor)
            End If
            
            vllngCveSocio = !intcvesocio
            
                pCargaDomicilioTelefono vllngCveSocio
                pCargaGridCredencial vllngCveSocio
                pCargaGridHispanidad vllngCveSocio
                pCargaDictamen vllngCveSocio
                pCargaNumerarios vllngCveSocio
                pCargaGridDependientes vllngCveSocio
                pEdadRFC
                pFormatoFechaGrid
                
                'Exit Do
        End If
    End If
  End With
            
    Exit Sub
NotificaError:
    'EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pBuscaSocio"))
End Sub

Private Sub pCargaCuentaContable(vllngNumeroCuenta)
Dim vlstrTemp As String
Dim rsCuentaTemp As New ADODB.Recordset

    On Error GoTo NotificaError
    
    vlstrTemp = "select vchcuentacontable " & _
                    "from cncuenta where intnumerocuenta = " & vllngNumeroCuenta
            
        Set rsCuentaTemp = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
    
        txtClaveContabilidad.Text = rsCuentaTemp!vchCuentaContable
        vlintCuentaContable = vllngNumeroCuenta
        rsCuentaTemp.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraCuentaContable"))
End Sub

Private Sub pCambioHispanidad(vllngClaveSocio, vlHispanidadAnterior, vlHispanidadActual)
    On Error GoTo NotificaError
    If vlblnDependiente Then 'si se trata de un dependiente
       'codigo ya no es necesario hispanidadanterior ya llega bien
       ' Select Case vlHispanidadAnterior
       '     Case "ES"
       '         vlHispanidadAnterior = "DE"
       '     Case "HE"
       '         vlHispanidadAnterior = "DH"
       '     Case "NE"
       '         vlHispanidadAnterior = "DN"
       '     Case "BE"
       '         vlHispanidadAnterior = "DB"
       '     Case "CE"
       '         vlHispanidadAnterior = "DC"
           'Case "VE" en los dependientes no aplica el VIUDO
        'End Select
    
     Select Case vlHispanidadActual
            Case "ES"
                vlHispanidadActual = "DE"
            Case "HE"
                vlHispanidadActual = "DH"
            Case "NE"
                vlHispanidadActual = "DN"
            Case "BE"
                vlHispanidadActual = "DB"
            Case "CE"
                vlHispanidadActual = "DC"
            'Case "VE" en los dependientes no aplica el VIUDO
        End Select
    
    Else ' en el caso de que se trate de un titular
    'codigo ya no es necesario hispanidadanterior ya llega bien
      '    Select Case vlHispanidadAnterior
      '      Case "ES"
      '          vlHispanidadAnterior = "ET"
      '      Case "HE"
      '          vlHispanidadAnterior = "HT"
      '      Case "NE"
      '          vlHispanidadAnterior = "NT"
      '       Case "VE"
      '          vlHispanidadAnterior = "VT"
      '      'Case "BE" 'un bisnieto no puede ser titular
            'Case "CE" 'un conyuge no puede ser titular
       '  End Select
    
     Select Case vlHispanidadActual
               Case "ES"
                vlHispanidadActual = "ET"
            Case "HE"
                vlHispanidadActual = "HT"
            Case "NE"
                vlHispanidadActual = "NT"
             Case "VE"
                vlHispanidadActual = "VT"
            'Case "BE" 'un bisnieto no puede ser titular
            'Case "CE" 'un conyuge no puede ser titular
        End Select
       
    End If
    
    If vlHispanidadActual <> vlHispanidadAnterior Then
                 vlstrSentenciaSQL = "insert into SOHISTORICOHISPANIDAD values(null," & _
                                 vllngClaveSocio & ",'" & vlHispanidadAnterior & "','" & vlHispanidadActual & "'," & fstrFechaSQL(CStr(Date)) & ")"

              pEjecutaSentencia vlstrSentenciaSQL
    End If
'    With rsHispanidad
'        .AddNew
'        !intcvesocio = vllngClaveSocio
'        !chrHispanidadAnterior = vlHispanidadAnterior
'        !chrHispanidadActual = vlHispanidadActual
'        !dtmFechaCambioHispanidad = Date
'        .Update
'        .Requery
'    End With

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCambioHispanidad"))
End Sub

Private Sub pConfiguraGridHispanidad()
    On Error GoTo NotificaError
    
    'Configura el grid de la búsqueda de Socios
    Dim vlintseq As Integer
    
    grdhHispanidad.Enabled = True
    With grdhHispanidad
    
        .FormatString = "NA|De|A|Fecha"
        .ColWidth(0) = 100
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 1300
        .ScrollBars = flexScrollBarBoth
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridHispanidad"))
End Sub

Private Sub pCargaGridHispanidad(vllngCveSocio As Long)
Dim vlstrTemp As String
Dim rsHispanidadTemp As New ADODB.Recordset

    On Error GoTo NotificaError
    
    With grdhHispanidad
            .Clear
            .Cols = 4
            .Rows = 2
        End With
    
    vlstrTemp = "SELECT CHRHISPANIDADANTERIOR, " & _
                    "CHRHISPANIDADACTUAL, DTMFECHACAMBIOHISPANIDAD " & _
                    "FROM SOHISTORICOHISPANIDAD " & _
                    "WHERE INTCVESOCIO = " & vllngCveSocio & " ORDER BY DTMFECHACAMBIOHISPANIDAD"
                 
        Set rsHispanidadTemp = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
        
        If rsHispanidadTemp.RecordCount > 0 Then
            pLlenarMshFGrdRs grdhHispanidad, rsHispanidadTemp
            pConfiguraGridHispanidad
        Else
            pConfiguraGridHispanidad
        End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaGridHispanidad"))
End Sub
Private Sub pCargaGridCredencial(vllngCveSocio As Long)
Dim vlstrTemp As String
Dim rsCredencialesTemp As New ADODB.Recordset

    On Error GoTo NotificaError
     With grdhCredenciales
            .Clear
            .Cols = 4
            .Rows = 2
        End With
        
        vlstrTemp = "SELECT (VCHSERIECREDENCIALANTERIOR||INTNUMEROCREDENCIALANTERIOR) Anterior, " & _
                    "(VCHSERIECREDENCIALACTUAL||INTNUMEROCREDENCIALACTUAL) Actual, DTMFECHACAMBIOCREDENCIAL " & _
                    "FROM SOHISTORICOCREDENCIAL " & _
                    "WHERE INTCVESOCIO = " & vllngCveSocio & " ORDER BY DTMFECHACAMBIOCREDENCIAL"
                 
        Set rsCredencialesTemp = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)

        If rsCredencialesTemp.RecordCount > 0 Then
          pLlenarMshFGrdRs grdhCredenciales, rsCredencialesTemp
          pConfiguraGridCredenciales
        Else
          pConfiguraGridCredenciales
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaGridCrdencial"))
End Sub

Private Sub pGrabaDictamen(vllngCveSocio As Long, vldtmFecha, vlvchObservaciones)
Dim vlintContador As Integer
    
    On Error GoTo NotificaError
    
    If vlblnDependiente Then ' los dependientes no pueden tener dictamentes
       Exit Sub
    Else ' en la pantalla de manejo de socios
       vlstrSentenciaSQL = "Select * from sodictamenes where intcvesocio = " & vllngCveSocio
       Set rsDictamenes = frsRegresaRs(vlstrSentenciaSQL, adLockOptimistic)
       
       If rsDictamenes.RecordCount = 0 Then ' no existe dictamen para este socio, se da de alta uno nuevo
          vlstrSentenciaSQL = "Insert into soDictamenes values (null," & vllngCveSocio & ","
          If mskFechaDictamen = "  /  /    " Then
            vlstrSentenciaSQL = vlstrSentenciaSQL & "null,"
          Else
            vlstrSentenciaSQL = vlstrSentenciaSQL & fstrFechaSQL(mskFechaDictamen.Text) & ","
          End If
            vlstrSentenciaSQL = vlstrSentenciaSQL & IIf(Trim(txtObservacionesDictamen.Text) = "", "null", "'" & txtObservacionesDictamen.Text & "'") & ")"
       Else ' si existe dictamen, se actualiza
            vlstrSentenciaSQL = "UPDATE soDictamenes set dtmFechaDictamen = "
          If mskFechaDictamen = "  /  /    " Then
            vlstrSentenciaSQL = vlstrSentenciaSQL & "null,"
          Else
            vlstrSentenciaSQL = vlstrSentenciaSQL & fstrFechaSQL(mskFechaDictamen.Text) & ","
          End If
            vlstrSentenciaSQL = vlstrSentenciaSQL & " vchobservaciones = " & IIf(Trim(txtObservacionesDictamen.Text) = "", "null", "'" & txtObservacionesDictamen.Text & "'")
            vlstrSentenciaSQL = vlstrSentenciaSQL & " where intcvesocio = " & vllngCveSocio
       End If
       pEjecutaSentencia vlstrSentenciaSQL
      
    End If
    
'    With rsDictamenes
'    If .EOF Then
'        .AddNew
'        !intcvesocio = vllngCveSocio
'        If mskFechaDictamen = "  /  /    " Then
'            !dtmFechaDictamen = Null
'        Else
'            !dtmFechaDictamen = CDate(mskFechaDictamen)
'        End If
'        !vchobservaciones = txtObservacionesDictamen
'        .Update
'        .Requery
'    Else
'    .MoveFirst
'        For vlintContador = 1 To .RecordCount
'            If !intcvesocio = vllngCveSocio Then
'                !intcvesocio = vllngCveSocio
'                If mskFechaDictamen = "  /  /    " Then
'                    !dtmFechaDictamen = Null
'                Else
'                    !dtmFechaDictamen = CDate(mskFechaDictamen)
'                End If
'                !vchobservaciones = txtObservacionesDictamen
'                .Update
'                .Requery
'                Exit Sub
'            End If
'            .MoveNext
'        Next
'        .AddNew
'        !intcvesocio = vllngCveSocio
'        If mskFechaDictamen = "  /  /    " Then
'            !dtmFechaDictamen = Null
'        Else
'            !dtmFechaDictamen = CDate(mskFechaDictamen)
'        End If
'        !vchobservaciones = txtObservacionesDictamen
'        .Update
'        .Requery
'    End If
'    End With

Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGrabaDictamen"))
End Sub
Private Sub pCargaDictamen(vllngCveSocio As Long)
    
 vlstrSentenciaSQL = "Select * from sodictamenes where intcvesocio = " & vllngCveSocio
 Set rsDictamenes = frsRegresaRs(vlstrSentenciaSQL, adLockOptimistic)
    
 If rsDictamenes.RecordCount = 0 Then
     pDictamenFolio
 Else
    With rsDictamenes
         txtFolioDictamen.Text = !intcvedictamen
         If Not IsNull(!dtmFechaDictamen) Then
                 pMkTextAsignaValor mskFechaDictamen, !dtmFechaDictamen
         Else
                 pMkTextAsignaValor mskFechaDictamen, ""
         End If
         txtObservacionesDictamen.Text = IIf(IsNull(!vchObservaciones), "", !vchObservaciones)
         txtNombreDictamen.Text = Trim(txtApePaterno.Text) & " " & Trim(txtApeMaterno.Text) & " " & Trim(txtNombre.Text)
    End With
 End If
    
    
    
'    Dim vlintContador As Integer
'
'    With rsDictamenes
'    .Requery
'    If Not .EOF Then
'    .MoveFirst
'        For vlintContador = 1 To .RecordCount
'            If !intcvesocio = vllngCveSocio Then
'                txtFolioDictamen.Text = !intcvedictamen
'                If Not IsNull(!dtmFechaDictamen) Then
'                    pMkTextAsignaValor mskFechaDictamen, !dtmFechaDictamen
'                Else
'                    pMkTextAsignaValor mskFechaDictamen, ""
'                End If
'                txtObservacionesDictamen.Text = IIf(IsNull(!vchobservaciones), "", !vchobservaciones)
'                txtNombreDictamen.Text = Trim(txtApePaterno.Text) & " " & Trim(txtApeMaterno.Text) & " " & Trim(txtNombre.Text)
'                Exit Sub
'            End If
'            .MoveNext
'        Next
'        pDictamenFolio
'    Else
'        txtFolioDictamen.Text = fSigConsecutivo("intcvedictamen", "sodictamenes")
'        txtNombreDictamen.Text = Trim(txtApePaterno.Text) & " " & Trim(txtApeMaterno.Text) & " " & Trim(txtNombre.Text)
'    End If
'
'
'    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaDictamen"))
End Sub
Private Sub pDictamenFolio()
Dim vlstrTemp As String
Dim rsFolioDictamen As New ADODB.Recordset
    On Error GoTo NotificaError
    vlstrTemp = "select max(intcvedictamen) maximo from sodictamenes"
    Set rsFolioDictamen = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
    
    If Not rsFolioDictamen.EOF Then
        rsFolioDictamen.Requery
        txtFolioDictamen.Text = rsFolioDictamen!Maximo + 1
        txtNombreDictamen.Text = Trim(txtApePaterno.Text) & " " & Trim(txtApeMaterno.Text) & " " & Trim(txtNombre.Text)
        'rsFolioDictamen.Requery
    End If

Exit Sub
NotificaError:
        EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDictamenFolio"))
End Sub
Private Sub pGrabaNumerarios(vllngCveSocio As Long)
'''Dim vlintContador As Integer
'''Dim vlintCuentaSocios As Integer
'''Dim vlblnNumerariosNuevos As Boolean
'''Dim vlstrTemp As String
'''Dim rsSocioNumDel As New ADODB.Recordset

    On Error GoTo NotificaError
        pEjecutaSentencia "Delete From SoSocioNumerario Where INTCVESOCIO = " & vllngCveSocio
        If txtNombreSocio1.Text <> "" Then
            pEjecutaSentencia "Insert Into SoSocioNumerario (intcvesocio, intcvesocionumerario, intorden) Values (" & vllngCveSocio & ", " & txtCveSocioNum1.Text & ", 1)"
        End If
        If txtNombreSocio2.Text <> "" Then
            pEjecutaSentencia "Insert Into SoSocioNumerario (intcvesocio, intcvesocionumerario, intorden) Values (" & vllngCveSocio & ", " & txtCveSocioNum2.Text & ", 2)"
        End If
    
'''    If vlblnDependiente Or (txtNombreSocio1 = "" And txtNombreSocio2 = "") Then
'''        vlstrTemp = "select count(*) from sosocionumerario where sosocionumerario.INTCVESOCIO = " & vllngCveSocio
'''        Set rsSocioNumDel = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
'''        If Not rsSocioNumDel.RecordCount > 0 Then
'''            Exit Sub
'''        End If
'''    End If
'''
'''    With rsNumerarios
'''        If Not .EOF Then
'''            .MoveFirst
'''            For vlintContador = 1 To .RecordCount
'''                If !intcvesocio = vllngCveSocio Then
'''                    !intcvesocio = vllngCveSocio
'''                    If txtCveSocioNum1.Text = "" Then
'''                        .Delete
'''                    Else
'''                        !intcvesocionumerario = txtCveSocioNum1.Text
'''                    End If
'''                    vlblnNumerariosNuevos = False
'''                    Exit For
'''                Else
'''                    vlblnNumerariosNuevos = True
'''                End If
'''                .MoveNext
'''            Next
'''        Else
'''            vlblnNumerariosNuevos = True
'''        End If
'''
'''        If Not txtCveSocioNum1.Text = "" Then
'''            If vlblnNumerariosNuevos Then
'''                .AddNew
'''                !intcvesocio = vllngCveSocio
'''                !intcvesocionumerario = IIf((Trim(txtCveSocioNum1.Text) = ""), vllngCveSocio, CInt(txtCveSocioNum1.Text))
'''                .Update
'''            End If
'''            .Update
'''            .Requery
'''        End If
'''
'''        If Not .EOF Then
'''        .MoveFirst
'''            For vlintContador = 1 To .RecordCount
'''                If !intcvesocio = vllngCveSocio And Not !intcvesocionumerario = txtCveSocioNum1.Text Then
'''                    !intcvesocio = vllngCveSocio
'''                    If txtCveSocioNum2.Text = "" Then
'''                        .Delete
'''                    Else
'''                        !intcvesocionumerario = txtCveSocioNum2.Text
'''                    End If
'''                    vlblnNumerariosNuevos = False
'''                    Exit For
'''                Else
'''                    vlblnNumerariosNuevos = True
'''                End If
'''            .MoveNext
'''            Next
'''        Else
'''            vlblnNumerariosNuevos = True
'''        End If
'''
'''        If Not txtCveSocioNum2.Text = "" Then
'''        If vlblnNumerariosNuevos Then
'''            .AddNew
'''            !intcvesocio = vllngCveSocio
'''            !intcvesocionumerario = IIf((Trim(txtCveSocioNum2.Text) = ""), Null, CInt(txtCveSocioNum2.Text))
'''            .Update
'''        End If
'''        .Update
'''        .Requery
'''        End If
'''        End With
'''    'End If
'''    vlintCuentaSocios = 0
Exit Sub
NotificaError:
        EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGrabaNumerarios"))
End Sub
'Private Sub pCargaDictament(vllngCveSocio As Long)
    
    'Dim vlintContador As Integer
    'Dim vlstrTemp As String
    'Dim rsFolioDictamen As New ADODB.Recordset

    'vlstrTemp = "select rtrim(vchapellidopaterno) || ' ' || rtrim(vchapellidomaterno) || ' ' || rtrim(vchnombre) nombre from sosocio where intcvesocio = " & rs!intccvesocionumerario
    'Set rsFolioDictamen = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
    
    'With rsDictamenes
    '.Requery
    '.MoveFirst
    'If Not .EOF Then
        'For vlintContador = 1 To .RecordCount
            'If !INTCVESOCIO = vllngCveSocio Then
                'If Not !intcvesocionumerario = vllngCveSocio Then
                    'If txtNombreSocio1.Text <> "" Then
                        'vlstrTemp = "select rtrim(vchapellidopaterno) || ' ' || rtrim(vchapellidomaterno) || ' ' || rtrim(vchnombre) nombre from sosocio where intcvesocio = " & !intccvesocionumerario
                        'Set rsFolioDictamen = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
                    'End If
                'End If
            'End If
        '.MoveNext
        'Next
    'End If
    
    
    'End With
'Exit Sub
'NotificaError:
    'Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
'End Sub

Private Sub pCargaNumerarios(vllngCveSocio As Long)
    Dim rsSocioNumerario As New ADODB.Recordset
    Dim strSentencia As String
    Dim intNumero As Integer
    
    strSentencia = " Select intcvesocionumerario, (VCHAPELLIDOPATERNO || ' ' || VCHAPELLIDOMATERNO || ' ' || VCHNOMBRE) Nombre, intorden " & _
                   "   From SOSOCIONUMERARIO " & _
                   "        Inner Join SOSOCIO On (SOSOCIONUMERARIO.intcvesocionumerario = SOSOCIO.intcvesocio) " & _
                   "  Where SOSOCIONUMERARIO.intcvesocio = " & vllngCveSocio
    
    Set rsSocioNumerario = frsRegresaRs(strSentencia)
    intNumero = 1
    With rsSocioNumerario
        While Not .EOF
            If IsNull(!intOrden) Then
                If intNumero = 1 Then
                    txtCveSocioNum1.Text = !intcvesocionumerario
                    txtNombreSocio1.Text = !Nombre
                Else
                    txtCveSocioNum2.Text = !intcvesocionumerario
                    txtNombreSocio2.Text = !Nombre
                End If
            Else
                If !intOrden = 1 Then
                    txtCveSocioNum1.Text = !intcvesocionumerario
                    txtNombreSocio1.Text = !Nombre
                Else
                    txtCveSocioNum2.Text = !intcvesocionumerario
                    txtNombreSocio2.Text = !Nombre
                End If
            End If
            intNumero = intNumero + 1
            .MoveNext
        Wend
    End With
        
'''    Dim vlstrTemp As String
'''    Dim rsSocioNumTemp As New ADODB.Recordset
'''    Dim vlintCuentaSocios As Integer
'''    Dim vlintContador As Integer
'''    Dim socio1 As Integer
'''    Dim socio2 As Integer
'''
'''    vlintCuentaSocios = 0
'''
'''    With rsNumerarios
'''        .Requery
'''        If Not .EOF Then
'''        .MoveFirst
'''            For vlintContador = 1 To .RecordCount
'''                If !INTCVESOCIO = vllngCveSocio Then
'''                    If vlintCuentaSocios = 0 Then
'''                    txtCveSocioNum1.Text = !intcvesocionumerario
'''                    socio1 = !intcvesocionumerario
'''                    Else
'''                    txtCveSocioNum2.Text = !intcvesocionumerario
'''                    socio2 = !intcvesocionumerario
'''                    Exit For
'''                    End If
'''                    vlintCuentaSocios = vlintCuentaSocios + 1
'''                End If
'''            .MoveNext
'''            Next
'''            vlintCuentaSocios = 0
'''
'''            vlstrTemp = "select rtrim(vchapellidopaterno) || ' ' || rtrim(vchapellidomaterno) || ' ' || rtrim(vchnombre) nombre " & _
'''                "from sosocio where intcvesocio = " & socio1
'''            Set rsSocioNumTemp = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
'''
'''            If Not rsSocioNumTemp.EOF Then
'''                txtNombreSocio1.Text = rsSocioNumTemp!Nombre
'''            End If
'''
'''
'''            vlstrTemp = "select rtrim(vchapellidopaterno) || ' ' || rtrim(vchapellidomaterno) || ' ' || rtrim(vchnombre) nombre " & _
'''                        "from sosocio where intcvesocio = " & socio2
'''            Set rsSocioNumTemp = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
'''
'''            If Not rsSocioNumTemp.EOF Then
'''                txtNombreSocio2.Text = rsSocioNumTemp!Nombre
'''            End If
'''
'''        End If
'''        .Requery
'''    End With
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaNumerarios"))
End Sub

Private Sub pConfiguraDependiente()
    
    If vlblnDependiente Then
        sstOpcion.TabsPerRow = 1
        optTipoSocio(1).Value = True
        lblRegistroSBE.Caption = "Nombre del titular"
        cmdCambioSocio.Caption = "Titular a dependiente"
        cmdCambioSocio.ToolTipText = "Elegir y cambiar socio de titular a dependiente"
        frmSocios.Caption = "Manejo de dependientes"
        txtRegistroSBE.ToolTipText = "Nombre del titular asociado"
        txtRegistroSBE.Locked = True
        txtClaveContabilidad.TabStop = False
    Else
        sstOpcion.TabsPerRow = 4
        optTipoSocio(0).Value = True
        lblRegistroSBE.Caption = "Registro SBE"
        cmdCambioSocio.Caption = "Dependiente a titular"
        cmdCambioSocio.ToolTipText = "Elegir y cambiar socio de dependiente a titular"
        txtRegistroSBE.ToolTipText = "Nombre para el registro SBE"
        txtRegistroSBE.Locked = False
        txtClaveContabilidad.TabStop = True
        frmSocios.Caption = "Manejo de socios"
    End If
    
    If vlblnMostrarTabDependientes = True Then
        sstOpcion.TabVisible(4) = True
    Else
        sstOpcion.TabVisible(4) = False
    End If
    
    If vlblnMostrarTabDictamenes = True Then
        sstOpcion.TabVisible(3) = True
    Else
        sstOpcion.TabVisible(3) = False
    End If
    
    If vlblnMostrarTabDocumentacion = True Then
        sstOpcion.TabVisible(2) = True
    Else
        sstOpcion.TabVisible(2) = False
    End If
    
    If vlblnMostrarTabDomicilios = True Then
        sstOpcion.TabVisible(0) = True
    Else
        sstOpcion.TabVisible(0) = False
    End If

    If vlblnMostrarTabEstado = True Then
        sstOpcion.TabVisible(1) = True
    Else
        sstOpcion.TabVisible(1) = False
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraDependiente"))
End Sub

Private Sub pCargaDomicilioTelefono(vllngCveSocio As Long)
Dim rsDomicilio As New ADODB.Recordset
Dim rsTelefono As New ADODB.Recordset

        'Domicilio particular
            If vlblnDependiente = False Then
                '---Titulares---
                Set rsDomicilio = frsEjecuta_SP(vllngCveSocio & "|3", "SP_SOSELDOMICILIOSOCIO")
                If rsDomicilio.RecordCount > 0 Then
                    txtDomicilio.Text = Trim(IIf(IsNull(rsDomicilio!vchCalle), "", rsDomicilio!vchCalle))
                    txtColonia.Text = Trim(IIf(IsNull(rsDomicilio!vchcolonia), "", rsDomicilio!vchcolonia))
                    txtCP.Text = Trim(IIf(IsNull(rsDomicilio!vchCodigoPostal), "", rsDomicilio!vchCodigoPostal))
                    txtNumeroExterior.Text = fstrIsNull(rsDomicilio!VCHNUMEROEXTERIOR)
                    txtNumeroInterior.Text = fstrIsNull(rsDomicilio!VCHNUMEROINTERIOR)
                    cboCiudad.ListIndex = fintLocalizaCbo(cboCiudad, fintIsNull(rsDomicilio!intCveCiudad))
                Else
                    txtDomicilio.Text = ""
                    txtColonia.Text = ""
                    txtCP.Text = ""
                    txtNumeroExterior.Text = ""
                    txtNumeroInterior.Text = ""
                    cboCiudad.ListIndex = -1
                End If
                    If Trim(txtNumeroInteriorD.Text) = "" Then txtNumeroInteriorD.Enabled = False
                
            'Domicilio Trabajo
                Set rsDomicilio = frsEjecuta_SP(vllngCveSocio & "|4", "SP_SOSELDOMICILIOSOCIO")
                If rsDomicilio.RecordCount > 0 Then
                    txtDomicilioT.Text = Trim(IIf(IsNull(rsDomicilio!vchCalle), "", rsDomicilio!vchCalle))
                    txtCPT.Text = Trim(IIf(IsNull(rsDomicilio!vchCodigoPostal), "", rsDomicilio!vchCodigoPostal))
                Else
                    txtDomicilioT.Text = ""
                    txtCPT.Text = ""
                End If
                rsDomicilio.Close
            
            'Telefono local socio
                Set rsTelefono = frsEjecuta_SP(vllngCveSocio & "|3", "SP_SOSELTELEFONOSOC")
                If rsTelefono.RecordCount > 0 Then
                    txtTelefono.Text = Trim(IIf(IsNull(rsTelefono!vchtelefono), "", rsTelefono!vchtelefono))
                Else
                    txtTelefono.Text = ""
                End If
            
            'Telefono trabajo socio
                Set rsTelefono = frsEjecuta_SP(vllngCveSocio & "|4", "SP_SOSELTELEFONOSOC")
                If rsTelefono.RecordCount > 0 Then
                    txtTelefonoT.Text = Trim(IIf(IsNull(rsTelefono!vchtelefono), "", rsTelefono!vchtelefono))
                Else
                    txtTelefono.Text = ""
                End If
    
            'Fax local socio
                Set rsTelefono = frsEjecuta_SP(vllngCveSocio & "|5", "SP_SOSELTELEFONOSOC")
                If rsTelefono.RecordCount > 0 Then
                    txtFax.Text = Trim(IIf(IsNull(rsTelefono!vchtelefono), "", rsTelefono!vchtelefono))
                Else
                    txtFax.Text = ""
                End If
                rsTelefono.Close
                
            Else
            
                '---Dependientes---
                Set rsDomicilio = frsEjecuta_SP(vllngCveSocio & "|3", "SP_SOSELDOMICILIOSOCIO")
                If rsDomicilio.RecordCount > 0 Then
                    txtDomicilioD.Text = Trim(IIf(IsNull(rsDomicilio!vchCalle), "", rsDomicilio!vchCalle))
                    txtColoniaD.Text = Trim(IIf(IsNull(rsDomicilio!vchcolonia), "", rsDomicilio!vchcolonia))
                    txtCPD.Text = Trim(IIf(IsNull(rsDomicilio!vchCodigoPostal), "", rsDomicilio!vchCodigoPostal))
                    txtNumeroExteriorD.Text = fstrIsNull(rsDomicilio!VCHNUMEROEXTERIOR)
                    txtNumeroInteriorD.Text = fstrIsNull(rsDomicilio!VCHNUMEROINTERIOR)
                    cboCiudadD.ListIndex = fintLocalizaCbo(cboCiudadD, fintIsNull(rsDomicilio!intCveCiudad))
                Else
                    txtDomicilioD.Text = ""
                    txtColoniaD.Text = ""
                    txtCPD.Text = ""
                    txtNumeroExteriorD.Text = ""
                    txtNumeroInteriorD.Text = ""
                    cboCiudadD.ListIndex = -1
                End If
                    If Trim(txtNumeroInteriorD.Text) = "" Then txtNumeroInteriorD.Enabled = False
                
            'Telefono local socio
                Set rsTelefono = frsEjecuta_SP(vllngCveSocio & "|3", "SP_SOSELTELEFONOSOC")
                If rsTelefono.RecordCount > 0 Then
                    txtTelefonoD.Text = Trim(IIf(IsNull(rsTelefono!vchtelefono), "", rsTelefono!vchtelefono))
                Else
                    txtTelefonoD.Text = ""
                End If
            
            'Fax local socio
                Set rsTelefono = frsEjecuta_SP(vllngCveSocio & "|5", "SP_SOSELTELEFONOSOC")
                If rsTelefono.RecordCount > 0 Then
                    txtFaxD.Text = Trim(IIf(IsNull(rsTelefono!vchtelefono), "", rsTelefono!vchtelefono))
                Else
                    txtFaxD.Text = ""
                End If
                rsTelefono.Close
            End If
        
Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaDomicilioTelefono"))
End Sub

Private Sub txtTelefonoD_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtTelefonoD_GotFocus()
    pSelTextBox txtTelefonoD
End Sub

Private Sub txtTelefonoD_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub txtTelefonoEmergencia_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtTelefonoEmergencia_GotFocus()
pSelTextBox txtTelefonoEmergencia
End Sub
Private Sub txtTelefonoEmergencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And sstOpcion.Tab = 4 Then
txtDomicilioD.SetFocus
End If

End Sub

Private Sub txtTelefonoEmergencia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    
    
End Sub

Private Sub txtTelefonoT_Change()
    If Not stEstado = stNuevo And Not stEstado = stEspera Then
        pPonEstado stedicion
    End If
End Sub

Private Sub txtTelefonoT_GotFocus()
    pSelTextBox txtTelefonoT
End Sub

Private Sub txtTelefonoT_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
End Sub

Private Sub pConfiguraGridDependientes()
    On Error GoTo NotificaError
    
    'Configura el grid de la búsqueda de Socios
    Dim vlintseq As Integer
    
    grdhDependientes.Enabled = True
    With grdhDependientes
    
        .FormatString = "NA|Nombre|Clave única|Fecha alta"
        .ColWidth(0) = 100
        .ColWidth(1) = 5500
        .ColWidth(2) = 2500
        .ColWidth(3) = 2000
        .ScrollBars = flexScrollBarBoth
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridDependientes"))
End Sub

Private Sub pCargaGridDependientes(vllngCveSocio As Long)
Dim vlstrTemp As String
Dim rsDependientesTemp As New ADODB.Recordset

    On Error GoTo NotificaError
     With grdhDependientes
            .Clear
            .Cols = 4
            .Rows = 2
        End With
        
        vlstrTemp = "SELECT (VCHAPELLIDOPATERNO||' '||VCHAPELLIDOMATERNO||' '||VCHNOMBRE) Nombre," & _
        " VCHCLAVESOCIO Clave, " & _
        "dtmfechaingreso Alta " & _
        "From SOSOCIO " & _
        "INNER JOIN SOSOCIODEPENDIENTE ON SOSOCIODEPENDIENTE.INTCVEDEPENDIENTE = SOSOCIO.INTCVESOCIO " & _
        "Where SOSOCIODEPENDIENTE.INTCVESOCIO = " & vllngCveSocio
                 
        Set rsDependientesTemp = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
        txtDependientes.Text = rsDependientesTemp.RecordCount
        If rsDependientesTemp.RecordCount > 0 Then
          pLlenarMshFGrdRs grdhDependientes, rsDependientesTemp
          pConfiguraGridDependientes
          txtDependientes.Text = rsDependientesTemp.RecordCount
        Else
          pConfiguraGridDependientes
          txtDependientes.Text = 0
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaGridDependientes"))
End Sub

Private Sub pGrabaDependiente(vllngCveSocio As Long)
Dim vlintContador As Integer
Dim vlintContador2 As Integer
Dim vlblnDependienteReg As Boolean
Dim vlstrNombreTit As String

    On Error GoTo NotificaError

    If Not vlblnDependiente Then
       If optTipoSocio(1) Then
          vlstrSentenciaSQL = "Delete from SOSOCIODEPENDIENTE where intcvedependiente = " & CStr(vllngCveSocio)
          pEjecutaSentencia vlstrSentenciaSQL
       End If
'        With rsDependientes
'            If Not .EOF Then
'                If optTipoSocio(1) Then
'                    .MoveFirst
'                    For vlintContador = 1 To .RecordCount
'                        If !intcvedependiente = vllngCveSocio Then 'And vllngTitular = !intcvesocio Then
'                            .Delete
'                            .Update
'                        End If
'                        .MoveNext
'                    Next
'
'                    .Requery
'                End If
'            End If
'        End With
'        Exit Sub
    Else
        vlstrSentenciaSQL = "SELECT * from SoSocioDependiente where intcvedependiente = " & CStr(vllngCveSocio)
        Set rsDependientes = frsRegresaRs(vlstrSentenciaSQL, adLockOptimistic)
        If rsDependientes.RecordCount = 0 Then ' si no encuentra nada de nada
           vlstrSentenciaSQL = "insert into SosocioDependiente values (" & vllngTitular & "," & vllngCveSocio & ")"
        Else ' si encontro registrado el dependiente
            vlstrSentenciaSQL = "UPDATE SosocioDependiente set intcvesocio = " & vllngTitular & ", intcvedependiente = " & vllngCveSocio & _
                                " where intcvedependiente = " & vllngCveSocio
    
        End If
        pEjecutaSentencia vlstrSentenciaSQL
        
        ''ahora algo que no entiendo muy bien que digamos pero esta en el codigo original asi que no le vallamos a quitar funcionalidad a esto
        'algo asi como que si el socio paso a ser dependiente, el titular de este socio a hora sera el titular de los dependientes que tenia el socio
        'cuando era titular
        vlstrSentenciaSQL = "UPDATE SosocioDependiente set intcvesocio = " & vllngTitular & " where intcvesocio = " & vllngCveSocio
        
        pEjecutaSentencia vlstrSentenciaSQL
        
        
    
    End If
    



'    With rsDependientes
'        If .EOF Then
'                .AddNew
'                !intcvesocio = vllngTitular
'                !intcvedependiente = vllngCveSocio
'                .Update
'                .Requery
'        Else
'            .MoveFirst
'            For vlintContador = 1 To .RecordCount
'                If vllngCveSocio = !intcvedependiente Then 'And vllngTitular = !intcvesocio Then
'                    vlblnDependienteReg = False
'                    Exit For
'                End If
'                .MoveNext
'            Next
'
'            If vlblnDependienteReg Then
'                .AddNew
'                !intcvesocio = vllngTitular
'                !intcvedependiente = vllngCveSocio
'                .Update
'                .Requery
'            Else
'                !intcvesocio = vllngTitular
'                !intcvedependiente = vllngCveSocio
'                .Update
'                .Requery
'            End If
'
'            If vllngTitular <> 0 Then
'                .MoveFirst
'                For vlintContador = 1 To .RecordCount
'                    If !intcvesocio = vllngCveSocio Then 'And vllngTitular = !intcvesocio Then
'                        !intcvesocio = vllngTitular
'                        .Update
'                    End If
'                    .MoveNext
'                Next
'
'                .Requery
'            End If
'        End If
'
'
'    End With
    
Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGrabaDependiente"))
End Sub
Private Sub pCargaTitular(vllngCveSocio As Long)
    On Error GoTo NotificaError
    
    vlstrSentenciaSQL = "Select * from SoSocioDependiente where intcvedependiente = " & CStr(vllngCveSocio)
    Set rsDependientes = frsRegresaRs(vlstrSentenciaSQL, adLockOptimistic)
    
    If rsDependientes.RecordCount = 0 Then
       vllngTitular = 0
    Else
       vllngTitular = rsDependientes!intcvesocio
    End If
    
'    With rsDependientes
'    If .RecordCount > 0 Then
'        .MoveFirst
'    End If
'        For vlintContador = 1 To .RecordCount
'            If !intcvedependiente = vllngCveSocio Then
'                vllngTitular = !intcvesocio
'                Exit For
'            End If
'            .MoveNext
'        Next
'    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTitular"))
End Sub
Private Sub pImprime(vlstrDestino As String)
    On Error GoTo NotificaError
    Dim rptReporte As CRAXDRT.Report
    Dim rsReporte As ADODB.Recordset
    'Dim vlstrTipoDescuento As String
    'Dim vldblClave As Double
    'Dim vlStrTipoPaciente As String
    'Dim vlstrTipoCargo As String
    'Dim vldblVigencia As Double
    'Dim vlstrIniVigencia As String
    'Dim vlstrFinVigencia As String
    Dim alstrParametros(0) As String
    Dim vlstrTituloReporte As String
    
        
    Dim vlblnContinuar As Boolean
    
    Set rsReporte = frsEjecuta_SP(Trim(vllngCveSocio) & "|*", "SP_SORPTSELSOCIO")
    If rsReporte.RecordCount = 0 Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        pInstanciaReporte rptReporte, "rptImpresionSocio.rpt"
        rptReporte.DiscardSavedData
        'vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        pCargaParameterFields alstrParametros, rptReporte
        pImprimeReporte rptReporte, rsReporte, vlstrDestino, "Hoja frontal del socio"
    End If
    rsReporte.Close
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
    Unload Me
End Sub

Private Sub pFormatoFechaGrid()
    Dim lRow As Long
    '-- Formatear las fechas 'dd/mmm/yyyy'
    
    With grdhHispanidad
        .Redraw = False
        For lRow = .FixedRows To .Rows - 1
            .TextMatrix(lRow, 3) = Format(.TextMatrix(lRow, 3), "dd/mmm/yyyy")
        Next lRow
        .Redraw = True
    End With
    
    With grdhCredenciales
        .Redraw = False
        For lRow = .FixedRows To .Rows - 1
            .TextMatrix(lRow, 3) = Format(.TextMatrix(lRow, 3), "dd/mmm/yyyy")
        Next lRow
        .Redraw = True
    End With
    
    With grdhDependientes
        .Redraw = False
        For lRow = .FixedRows To .Rows - 1
            .TextMatrix(lRow, 3) = Format(.TextMatrix(lRow, 3), "dd/mmm/yyyy")
        Next lRow
        .Redraw = True
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pFormatoFechaGrid"))
End Sub
Private Sub pCargaSocioTit()
    Dim vlintContador As Integer
    Dim vlintContador2 As Integer
    Dim vlstrNombreTit As String
    
    vlstrSentenciaSQL = "select * from Sosocio where intcvesocio = " & CStr(vllngTitular)
    Set rsSocioTitular = frsRegresaRs(vlstrSentenciaSQL, adLockOptimistic)
    If rsSocioTitular.RecordCount = 0 Then
      vlstrNombreTit = ""
      vllngTitular = 0
    Else
      vlstrNombreTit = rsSocioTitular!vchApellidoPaterno & " " & rsSocioTitular!vchApellidoMaterno & " " & rsSocioTitular!vchNombre
      vllngTitular = rsSocioTitular!intcvesocio
      txtRegistroSBE.Text = vlstrNombreTit
      pPonEstado stConsulta
    End If
    
    
    
    

'    With rsDependientes
'            'If stEstado = stConsulta Or stEstado = stespera Then
'            If .RecordCount > 0 Then
'            .MoveFirst
'            End If
'            For vlintContador = 1 To .RecordCount
'                If !intcvesocio = vllngTitular Then
'                        rsSocioTitular.MoveFirst
'                        For vlintContador2 = 1 To rsSocioTitular.RecordCount
'                            If Not rsSocioTitular.EOF And Not .EOF Then
'                            If rsSocioTitular!intcvesocio = !intcvesocio Then
'                                'ya jala solo que hay que ponerle el concat de los nombres mejor jeje
'                                vlstrNombreTit = rsSocioTitular!vchApellidoPaterno & " " & rsSocioTitular!vchApellidoMaterno & " " & rsSocioTitular!vchNombre
'                                vllngTitular = !intcvesocio
'                                Exit For
'                            End If
'                            rsSocioTitular.MoveNext
'                            End If
'                        Next
'                        rsSocioTitular.MoveFirst
'                        For vlintContador2 = 1 To rsSocioTitular.RecordCount
'                            If Not rsSocioTitular.EOF And Not .EOF Then
'                            If rsSocioTitular!intcvesocio = !intcvedependiente Then
'                                txtRegistroSBE.Text = vlstrNombreTit
'                                pPonEstado stConsulta
'                                '.Update
'                                Exit For
'                            End If
'                            rsSocioTitular.MoveNext
'                            '.MoveNext
'                            End If
'                        Next
'                    'Exit For
'                End If
'                If Not .EOF Then
'                    .MoveNext
'                End If
'            Next
'            If .RecordCount > 0 Then
'                .MoveFirst
'            End If
'        'End If
'    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaSocioTit"))
End Sub
