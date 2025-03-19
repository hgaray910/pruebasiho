VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMtoConceptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de facturación"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   675
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   Begin TabDlg.SSTab SSTObj 
      Height          =   10620
      Left            =   -120
      TabIndex        =   20
      Top             =   -825
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   18733
      _Version        =   393216
      Tabs            =   4
      TabHeight       =   661
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Mantenimiento"
      TabPicture(0)   =   "frmMtoConcepto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBotonera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraConcepto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Departamentos que lo utilizan"
      TabPicture(1)   =   "frmMtoConcepto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fradepas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmMtoConcepto.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdExepcionesContables"
      Tab(2).Control(1)=   "grdConceptoEmpresas"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(3)=   "chkSaldarCuentas"
      Tab(2).Control(4)=   "cboEmpresaContable"
      Tab(2).Control(5)=   "Frame3"
      Tab(2).Control(6)=   "txtConceptosDe"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtEstructuraem"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "grdExepcionesContables2"
      Tab(2).Control(9)=   "Frame4"
      Tab(2).Control(10)=   "Label2"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmMtoConcepto.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblFilas"
      Tab(3).Control(1)=   "grdPrecios"
      Tab(3).Control(2)=   "lstPesos"
      Tab(3).Control(3)=   "UpDown1"
      Tab(3).ControlCount=   4
      Begin VB.ListBox UpDown1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         ItemData        =   "frmMtoConcepto.frx":0070
         Left            =   -74400
         List            =   "frmMtoConcepto.frx":007D
         TabIndex        =   91
         Top             =   2400
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ListBox lstPesos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         ItemData        =   "frmMtoConcepto.frx":00BB
         Left            =   -69000
         List            =   "frmMtoConcepto.frx":00C5
         TabIndex        =   90
         Top             =   2400
         Visible         =   0   'False
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExepcionesContables 
         Height          =   855
         Left            =   -74760
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   10590
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1508
         _Version        =   393216
         Rows            =   0
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConceptoEmpresas 
         Height          =   855
         Left            =   -74760
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   10470
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1508
         _Version        =   393216
         Rows            =   0
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cuentas contables"
         Height          =   4020
         Left            =   -74865
         TabIndex        =   74
         Top             =   1680
         Width           =   8520
         Begin VB.TextBox txtCuentaDEficienciaPaquetes 
            Height          =   315
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   3525
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaEficienciaPaquetes 
            Height          =   315
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   3110
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaIngresosSocialCU 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1870
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaDescuentosPendientes 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   2695
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaIngresosPendientes 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   2300
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaDescNota 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1460
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaIngreso 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   620
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaDescuento 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1040
            Width           =   4230
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   200
            Width           =   6400
         End
         Begin MSMask.MaskEdBox mskCuentaIngreso 
            Height          =   315
            Left            =   1980
            TabIndex        =   39
            ToolTipText     =   "Cuenta de ingresos"
            Top             =   620
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescuento 
            Height          =   315
            Left            =   1980
            TabIndex        =   41
            ToolTipText     =   "Cuenta de descuento"
            Top             =   1040
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescNota 
            Height          =   315
            Left            =   1980
            TabIndex        =   43
            ToolTipText     =   "Cuenta de descuento por nota de crédito"
            Top             =   1460
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaIngresosPendientes 
            Height          =   315
            Left            =   1980
            TabIndex        =   47
            ToolTipText     =   "Cuenta puente para ingresos pendientes de facturar"
            Top             =   2300
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescuentosPendientes 
            Height          =   315
            Left            =   1980
            TabIndex        =   49
            ToolTipText     =   "Cuenta puente para descuentos pendientes de facturar"
            Top             =   2695
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentadescuentoCUSocial 
            Height          =   315
            Left            =   1980
            TabIndex        =   45
            ToolTipText     =   "Cuenta de descuento por asistencia social"
            Top             =   1870
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaEficienciaPaquetes 
            Height          =   315
            Left            =   1980
            TabIndex        =   51
            ToolTipText     =   "Cuenta para ingresos por eficiencia en paquetes"
            Top             =   3110
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDeficienciaPaquetes 
            Height          =   315
            Left            =   1980
            TabIndex        =   53
            ToolTipText     =   "Cuenta para descuentos por deficiencia en paquetes"
            Top             =   3525
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
            Caption         =   "Cuenta de descuentos deficiencia en paquetes"
            Height          =   510
            Left            =   120
            TabIndex        =   89
            Top             =   3480
            Width           =   2025
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label21 
            Caption         =   "Cuenta para ingresos por eficiencia en paquetes"
            Height          =   390
            Left            =   120
            TabIndex        =   88
            Top             =   3070
            Width           =   1785
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label19 
            Caption         =   "Cuenta de descuento por asistencia social"
            Height          =   390
            Left            =   120
            TabIndex        =   86
            Top             =   1820
            Width           =   1785
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label16 
            Caption         =   "Cuenta para descuentos pendientes de facturar"
            Height          =   390
            Left            =   120
            TabIndex        =   83
            Top             =   2655
            Width           =   1785
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label15 
            Caption         =   "Cuenta para ingresos pendientes de facturar"
            Height          =   390
            Left            =   135
            TabIndex        =   82
            Top             =   2240
            Width           =   1785
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label13 
            Caption         =   "Cuenta de descuento por nota de crédito"
            Height          =   395
            Left            =   135
            TabIndex        =   80
            Top             =   1400
            Width           =   1790
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Departamento"
            Height          =   255
            Left            =   135
            TabIndex        =   35
            Top             =   230
            Width           =   1790
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta de ingreso"
            Height          =   255
            Left            =   135
            TabIndex        =   34
            Top             =   650
            Width           =   1790
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de descuento"
            Height          =   255
            Left            =   135
            TabIndex        =   33
            Top             =   1070
            Width           =   1790
         End
      End
      Begin VB.CheckBox chkSaldarCuentas 
         Caption         =   "Saldar cuentas de ingresos y descuentos"
         BeginProperty DataFormat 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   7
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         ToolTipText     =   "Saldar cuentas de ingresos y descuentos si se tiene la misma cuenta configurada"
         Top             =   1395
         Width           =   3315
      End
      Begin VB.ComboBox cboEmpresaContable 
         Height          =   315
         Left            =   -73440
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1020
         Width           =   7095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Excepciones contables"
         Height          =   3105
         Left            =   -74865
         TabIndex        =   75
         Top             =   5850
         Width           =   8520
         Begin VB.TextBox txtCuentaDescSocialEX1 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   1854
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaDescuentosPendientes1 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   2665
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaIngresosPendientes1 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   2265
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaDescNota1 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   1460
            Width           =   4230
         End
         Begin VB.ComboBox cboDepartamento1 
            Height          =   315
            Left            =   1980
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   200
            Width           =   6400
         End
         Begin VB.TextBox txtCuentaDescuento1 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   1040
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaIngreso1 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   620
            Width           =   4230
         End
         Begin MSMask.MaskEdBox mskCuentaIngreso1 
            Height          =   315
            Left            =   1980
            TabIndex        =   56
            ToolTipText     =   "Cuenta de ingresos"
            Top             =   620
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescuento1 
            Height          =   315
            Left            =   1980
            TabIndex        =   58
            ToolTipText     =   "Cuenta de descuentos"
            Top             =   1040
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescNota1 
            Height          =   315
            Left            =   1980
            TabIndex        =   60
            ToolTipText     =   "Cuenta de descuento por nota de crédito"
            Top             =   1440
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaIngresosPendientes1 
            Height          =   315
            Left            =   1980
            TabIndex        =   64
            ToolTipText     =   "Cuenta puente para ingresos pendientes de facturar"
            Top             =   2265
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescuentosPendientes1 
            Height          =   315
            Left            =   1980
            TabIndex        =   66
            ToolTipText     =   "Cuenta puente para descuentos pendientes de facturar"
            Top             =   2665
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescExSocial1 
            Height          =   315
            Left            =   1980
            TabIndex        =   62
            ToolTipText     =   "Cuenta de descuento por asistencia social"
            Top             =   1854
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label20 
            Caption         =   "Cuenta de descuento por asistencia social"
            Height          =   390
            Left            =   120
            TabIndex        =   87
            Top             =   1810
            Width           =   1785
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label18 
            Caption         =   "Cuenta para descuentos pendientes de facturar"
            Height          =   390
            Left            =   135
            TabIndex        =   85
            Top             =   2625
            Width           =   1785
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label17 
            Caption         =   "Cuenta para ingresos pendientes de facturar"
            Height          =   390
            Left            =   135
            TabIndex        =   84
            Top             =   2205
            Width           =   1785
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label14 
            Caption         =   "Cuenta de descuento por nota de crédito"
            Height          =   395
            Left            =   135
            TabIndex        =   81
            Top             =   1400
            Width           =   1790
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label6 
            Caption         =   "Cuenta de descuento"
            Height          =   255
            Left            =   135
            TabIndex        =   32
            Top             =   1070
            Width           =   1790
         End
         Begin VB.Label Label7 
            Caption         =   "Cuenta de ingreso"
            Height          =   255
            Left            =   135
            TabIndex        =   31
            Top             =   650
            Width           =   1790
         End
         Begin VB.Label Label8 
            Caption         =   "Departamento"
            Height          =   255
            Left            =   135
            TabIndex        =   30
            Top             =   230
            Width           =   1790
         End
      End
      Begin VB.TextBox txtConceptosDe 
         Height          =   285
         Left            =   -69480
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   10830
         Width           =   2775
      End
      Begin VB.TextBox txtEstructuraem 
         Height          =   285
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   10470
         Width           =   2775
      End
      Begin VB.Frame fraConcepto 
         Height          =   2130
         Left            =   155
         TabIndex        =   0
         Top             =   870
         Width           =   8520
         Begin VB.CheckBox chkPredeterminadoPaquetes 
            Caption         =   "Mostrar como predeterminado en presupuestos"
            Height          =   195
            Left            =   4320
            TabIndex        =   6
            ToolTipText     =   "Para uso de registro de paquete en base a un presupuesto"
            Top             =   1365
            Width           =   3930
         End
         Begin VB.CheckBox chkExentoIVA 
            Caption         =   "Exento de IVA"
            Height          =   195
            Left            =   4320
            TabIndex        =   4
            ToolTipText     =   "Exento de IVA"
            Top             =   1050
            Width           =   1530
         End
         Begin VB.ComboBox cboIvas 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Impuesto "
            Top             =   1365
            Width           =   2145
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Activo"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   7
            ToolTipText     =   "Estado"
            Top             =   1760
            Width           =   915
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1680
            MaxLength       =   3900
            TabIndex        =   2
            ToolTipText     =   "Descripción del concepto de facturación"
            Top             =   610
            Width           =   6585
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Tipo de concepto de factura"
            Top             =   990
            Width           =   2145
         End
         Begin MSMask.MaskEdBox txtCveConcepto 
            Height          =   315
            Left            =   1680
            TabIndex        =   1
            ToolTipText     =   "Clave del concepto de facturación"
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IVA"
            Height          =   195
            Left            =   135
            TabIndex        =   24
            Top             =   1395
            Width           =   255
         End
         Begin VB.Label lblClave 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   135
            TabIndex        =   23
            Top             =   300
            Width           =   405
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   135
            TabIndex        =   22
            Top             =   670
            Width           =   840
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de concepto"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   1050
            Width           =   1260
         End
      End
      Begin VB.Frame fraBotonera 
         Height          =   660
         Left            =   1653
         TabIndex        =   8
         Top             =   3150
         Width           =   5515
         Begin VB.CommandButton cmdListasPrecio 
            Caption         =   "Listas de precios"
            Height          =   480
            Left            =   4505
            TabIndex        =   17
            ToolTipText     =   "Configuración del tipo de incremento en las listas de precios del concepto"
            Top             =   135
            Width           =   975
         End
         Begin VB.CommandButton cmdCuentas 
            Caption         =   "Cuentas contables"
            Height          =   480
            Left            =   3520
            TabIndex        =   16
            Top             =   135
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   480
            Left            =   3015
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoConcepto.frx":00D9
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Borrar el registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Height          =   480
            Left            =   2520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoConcepto.frx":027B
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Guardar el registro"
            Top             =   135
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2025
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoConcepto.frx":05BD
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Ultimo registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1530
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoConcepto.frx":072F
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Siguiente registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1035
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoConcepto.frx":08A1
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Búsqueda"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   540
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoConcepto.frx":0A13
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Anterior registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   45
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoConcepto.frx":0B85
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Primer registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame fradepas 
         Height          =   3200
         Left            =   -74880
         TabIndex        =   18
         Top             =   930
         Width           =   8610
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
            CausesValidation=   0   'False
            DragIcon        =   "frmMtoConcepto.frx":0CF7
            Height          =   2960
            Left            =   45
            TabIndex        =   19
            ToolTipText     =   "Doble click para seleccionar un concepto"
            Top             =   150
            Width           =   8490
            _ExtentX        =   14975
            _ExtentY        =   5212
            _Version        =   393216
            ForeColor       =   0
            Rows            =   16
            Cols            =   5
            BackColorBkg    =   -2147483643
            GridColor       =   12632256
            GridColorFixed  =   -2147483632
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
            HighLight       =   0
            MergeCells      =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
            _Band(0).GridLineWidthBand=   1
            _Band(0).TextStyleBand=   0
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExepcionesContables2 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   10470
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.Frame Frame4 
         Caption         =   "Excepciones contables para socios"
         Height          =   1305
         Left            =   -74865
         TabIndex        =   76
         Top             =   9030
         Width           =   8520
         Begin VB.TextBox txtCuentaIngreso2 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   545
            Width           =   4230
         End
         Begin VB.TextBox txtCuentaDescuento2 
            Height          =   315
            Left            =   4150
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   870
            Width           =   4230
         End
         Begin VB.ComboBox cboPaciente 
            Height          =   315
            Left            =   1980
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   220
            Width           =   6400
         End
         Begin MSMask.MaskEdBox mskCuentaIngreso2 
            Height          =   315
            Left            =   1980
            TabIndex        =   69
            ToolTipText     =   "Cuenta de ingresos"
            Top             =   545
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaDescuento2 
            Height          =   315
            Left            =   1980
            TabIndex        =   71
            ToolTipText     =   "Cuenta de descuentos"
            Top             =   870
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de paciente"
            Height          =   255
            Left            =   135
            TabIndex        =   79
            Top             =   270
            Width           =   1790
         End
         Begin VB.Label Label10 
            Caption         =   "Cuenta de ingreso"
            Height          =   255
            Left            =   135
            TabIndex        =   78
            Top             =   575
            Width           =   1790
         End
         Begin VB.Label Label11 
            Caption         =   "Cuenta de descuento"
            Height          =   255
            Left            =   135
            TabIndex        =   77
            Top             =   900
            Width           =   1790
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPrecios 
         Height          =   3045
         Left            =   -74640
         TabIndex        =   92
         ToolTipText     =   "Captura de los precios"
         Top             =   1320
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   5371
         _Version        =   393216
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   -2147483638
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblFilas 
         Caption         =   $"frmMtoConcepto.frx":1001
         Height          =   1095
         Left            =   -74880
         TabIndex        =   93
         Top             =   4680
         Width           =   8415
      End
      Begin VB.Label Label2 
         Caption         =   "Empresa contable"
         Height          =   255
         Left            =   -74865
         TabIndex        =   73
         Top             =   1050
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMtoConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmMtoConceptos
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento del catálogo de Conceptos de Facturación
'|           PvConceptoFacturacion
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 31/Octubre/2000
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit 'Permite forzar la declaración de las variables

Const cintTipoHospital = 0 'Tipo de concepto para cargos
Const cintTipoSeguros = 1 'Tipo de concepto para facturas de seguros
Const cintTipoAdmitivo = 2 'Tipo de concepto para facturas directas

Const vllngSizeNormal = 3705
Const vllngSizeGrande = 9900
Const vllngSizeMediana = 8495



Public rsConceptos As New ADODB.Recordset
Public llngNumOpcion As Long 'Opcion para guardar datos

Dim vgblnNuevoRegistro As Boolean
Dim vglngDesktop  As Long
Dim vlstrx As String
Dim vlstrsql As String
Dim blnClaveManualCatalogo As Boolean
Dim blnEnfocando As Boolean
Dim vlUsaSocios As Boolean 'Indica si se utilizan socios en el sistema

Private Type arrListaPrecio
    vllngClaveLista As Long
    vlstrListaPrecios As String
    vlblnIncrementoAutomatico As Boolean
    vlstrTipoIncremento As String
    vldblmargenutilidad As Double
    vlblnUsaTabulador As Boolean
    vldblPrecio As Double
    vlblnPredeterminada As Boolean
    vlblnNuevoEnLista As Boolean
End Type
Dim aListaPrecio() As arrListaPrecio
Public rsArticuloPrecio As New ADODB.Recordset

'Columnas del grid de precios:
Const cintColClave = 1
Const cintColDescripcion = 2
Const cintColTipo = 3 '10
Const cintColTipoDes = 4 '11
Const cintColCveFact = 5 '14
Const cintColFacturacion = 6 '15
Const cintColPrecio = 7 '8
Const cintColCosto = 8 '7
Const cintColIncremetoAutomatico = 9 '3
Const cintColTipoIncremento = 10 '4
Const cintColUtilidad = 11 '5
Const cintColTabulador = 12 '6
Const cintColCostoUltimaEntrada = 13 '16
Const cintColCostoMasAlto = 14 '17
Const cintColPrecioMaximopublico = 15 '18
Const cintColPrecioNuevo = 16
Const cintColMoneda = 17


Const cstrUltimaCompra = "ÚLTIMA COMPRA"
Const cstrCompraMasAlta = "COMPRA MÁS ALTA"
Const cstrPrecioMaximoPublico = "PRECIO MÁXIMO AL PÚBLICO"



Private Function fblnDatosValidos() As Boolean
    Dim rsConcepto As New ADODB.Recordset
    Dim rsCargos As New ADODB.Recordset
    Dim rsPaquetes As New ADODB.Recordset
    Dim rsConceptoFacturacion As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    Dim vlintCboIVA As Integer

    fblnDatosValidos = True
    
    If Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2) + Chr(13) + txtDescripcion.ToolTipText, vbExclamation, "Mensaje"
        pEnfocaTextBox txtDescripcion
    End If
    
    If fblnDatosValidos And (chkExentoIVA.Value = 0 And cboIvas.ListIndex = -1) Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2) + Chr(13) + cboIvas.ToolTipText, vbExclamation, "Mensaje"
        cboIvas.SetFocus
    End If
    
    If fblnDatosValidos And Not vgblnNuevoRegistro Then
        Set rsConcepto = frsRegresaRs("SELECT SMYIVA FROM PVCONCEPTOFACTURACION WHERE SMICVECONCEPTO = " & Trim(txtCveConcepto.Text))
        If rsConcepto.RecordCount <> 0 Then
            If cboIvas.ListIndex = -1 Then
                vlintCboIVA = cboIvas.ListIndex
            Else
                vlintCboIVA = cboIvas.ItemData(cboIvas.ListIndex)
            End If
            If rsConcepto!smyIVA <> vlintCboIVA Then
                Set rsCargos = frsRegresaRs("SELECT COUNT(*) CANTIDADCARGOS " & _
                                            "From PVCARGO " & _
                                                "INNER JOIN EXPACIENTEINGRESO ON EXPACIENTEINGRESO.INTNUMCUENTA = PVCARGO.INTMOVPACIENTE " & _
                                                    "AND EXPACIENTEINGRESO.INTCVETIPOINGRESO IN (SELECT INTCVETIPOINGRESO " & _
                                                                                                "From SITIPOINGRESO " & _
                                                                                                "WHERE CHRTIPOINGRESO = PVCARGO.CHRTIPOPACIENTE) " & _
                                                    "AND EXPACIENTEINGRESO.INTCUENTAFACTURADA = 0 " & _
                                            "Where PVCARGO.SMICVECONCEPTO = " & Trim(txtCveConcepto.Text) & " " & _
                                                "AND (PVCARGO.CHRFOLIOFACTURA IS NULL OR TRIM(PVCARGO.CHRFOLIOFACTURA) = '')")
                If rsCargos!CANTIDADCARGOS <> 0 Then
                    fblnDatosValidos = False
                End If
                rsCargos.Close
                
                Set rsPaquetes = frsRegresaRs("SELECT COUNT(*) PENDIENTES FROM ( " & _
                                                "SELECT PVPAQUETEPACIENTE.* " & _
                                                    ", NVL((SELECT SUM(INTCANTIDADFACTURADA) FACTURADO " & _
                                                       "From PVPAQUETEPACIENTEFACTURADO " & _
                                                       "Where PVPAQUETEPACIENTEFACTURADO.INTMOVPACIENTE = PVPAQUETEPACIENTE.INTMOVPACIENTE " & _
                                                            "AND TRIM(PVPAQUETEPACIENTEFACTURADO.CHRTIPOPACIENTE) = TRIM(PVPAQUETEPACIENTE.CHRTIPOPACIENTE) " & _
                                                            "AND PVPAQUETEPACIENTEFACTURADO.INTNUMPAQUETE = PVPAQUETEPACIENTE.INTNUMPAQUETE " & _
                                                            "AND TRIM(PVPAQUETEPACIENTEFACTURADO.CHRESTATUS) = 'F'),0) FACTURADO " & _
                                                "From PVPAQUETEPACIENTE " & _
                                                    "INNER JOIN PVPAQUETE on PVPAQUETE.INTNUMPAQUETE = PVPAQUETEPACIENTE.INTNUMPAQUETE " & _
                                                "WHERE PVPAQUETE.SMICONCEPTOFACTURA = " & Trim(txtCveConcepto.Text) & ") PAQUETESDEFACTURAR " & _
                                                    "INNER JOIN EXPACIENTEINGRESO ON EXPACIENTEINGRESO.INTNUMCUENTA = PAQUETESDEFACTURAR.INTMOVPACIENTE " & _
                                                        "AND EXPACIENTEINGRESO.INTCVETIPOINGRESO IN (SELECT INTCVETIPOINGRESO " & _
                                                                                                    "From SITIPOINGRESO " & _
                                                                                                    "WHERE CHRTIPOINGRESO = PAQUETESDEFACTURAR.CHRTIPOPACIENTE) " & _
                                                        "AND EXPACIENTEINGRESO.INTCUENTAFACTURADA = 0 " & _
                                                "WHERE INTCANTIDAD > FACTURADO")
                If rsPaquetes!PENDIENTES <> 0 Then
                    fblnDatosValidos = False
                End If
                rsPaquetes.Close
                
                If Not fblnDatosValidos Then
                    ' No es posible cambiar el IVA del concepto de facturación, debido a que existen cargos o paquetes pendientes de facturarse relacionados a dicho concepto, favor de verificar.
                    MsgBox SIHOMsg(1402), vbOKOnly + vbInformation, "Mensaje"
                    cboIvas.Enabled = True
                    cboIvas.SetFocus
                    If chkExentoIVA.Value = 1 Then
                        chkExentoIVA.Value = 0
                    End If
                End If
            End If
            If fblnDatosValidos Then
                If rsConceptos!bitactivo <> chkActivo.Value Then
                    Set rsConceptoFacturacion = frsRegresaRs("SELECT  (SELECT COUNT(*) CONCEPTOS FROM IVARTICULO Where SMICVECONCEPTFACT = " & Trim(txtCveConcepto.Text) & " OR SMICVECONCEPTFACT2 = " & Trim(txtCveConcepto.Text) & ") +" & _
                             "(SELECT COUNT(*) CONCEPTOS FROM LaExamen Where SMICONFACT = " & Trim(txtCveConcepto.Text) & ") +" & _
                             "(SELECT COUNT(*) CONCEPTOS FROM LaGrupoExamen Where SMICONFACT = " & Trim(txtCveConcepto.Text) & ") +" & _
                             "(SELECT COUNT(*) CONCEPTOS FROM ImEstudio Where SMICONFACT = " & Trim(txtCveConcepto.Text) & ") +" & _
                             "(SELECT COUNT(*) CONCEPTOS FROM PVOTROCONCEPTO Where SMICONCEPTOFACT = " & Trim(txtCveConcepto.Text) & ")+" & _
                             "(SELECT COUNT(*) CONCEPTOS FROM PvPaquete Where SMICONCEPTOFACTURA = " & Trim(txtCveConcepto.Text) & ") as conceptos FROM dual")
                    If rsConceptoFacturacion!CONCEPTOS > 0 Then
                        fblnDatosValidos = False
                        chkActivo.Value = 1
                        chkActivo.SetFocus
                    End If
                    rsConceptoFacturacion.Close
                    If Not fblnDatosValidos Then
                        ' No es posible inactivar el concepto de facturación, debido a que existen cargos relacionados con dicho concepto, favor de verificar.
                        MsgBox SIHOMsg(20005), vbOKOnly + vbInformation, "Mensaje"
                    End If
                End If
            End If
        End If
        rsConcepto.Close
    End If
    
    If fblnDatosValidos And cboTipo.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbExclamation, "Mensaje"
        cboTipo.SetFocus
    End If
    
    If fblnDatosValidos Then
        fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "C", True)
        If Not fblnDatosValidos Then
            '¡El usuario no tiene permiso para grabar datos!
            MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    
    'Verifica que no exista otro concepto marcado como predeterminado
    If fblnDatosValidos Then
        If chkPredeterminadoPaquetes.Value = 1 Then
            vlstrSentencia = "SELECT * From pvconceptofacturacion " & _
                             "INNER JOIN pvconceptofacturacionempresa ON pvconceptofacturacion.smicveconcepto = pvconceptofacturacionempresa.intcveconceptofactura " & _
                             "WHERE pvconceptofacturacionempresa.intCveDepartamento = " & vgintNumeroDepartamento & _
                             " AND pvconceptofacturacion.bitpaquetepresupuesto = 1"
            Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rs.RecordCount <> 0 Then
                If rs!smicveconcepto <> Val(txtCveConcepto.Text) Then
                    'Ya se encuentra el concepto  como predeterminado para el departamento del concepto
                    MsgBox "Ya se encuentra el concepto " & rs!chrdescripcion & " como predeterminado para este departamento.", vbExclamation + vbOKOnly, "Mensaje"
                    chkPredeterminadoPaquetes.Value = 0
                    fblnDatosValidos = False
                End If
            End If
        End If
    End If


End Function

Private Sub pCargaIvas()
On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "SELECT relPorcentaje, vchDescripcion AS Descripcion FROM CnImpuesto WHERE bitActivo = 1"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboIvas, rs, 0, 1
        cboIvas.ListIndex = 0
    End If
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaIvas"))
    Unload Me
End Sub

Private Sub cboDepartamento_Click()
    Dim intRow As Integer
    Dim rsListas As ADODB.Recordset
    
    intRow = fintLocalizaRow()
    If intRow > -1 Then
        If cboDepartamento.ListIndex > -1 Then
            vgstrParametrosSP = cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & txtCveConcepto.Text & "|" & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
            Set rsListas = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelListaDeptos")
            If rsListas.RecordCount > 0 Then
                ' Si el concepto existía en una lista predeterminada , pero no existe una lista predeterminada para el nuevo departamento
                If rsListas!listaactual > 0 And rsListas!listanueva = 0 Then
                    If MsgBox(SIHOMsg(914) & " " & cboDepartamento.Text & " ¿Desea continuar?", vbYesNo + vbCritical, "Mensaje") = vbYes Then
                        grdConceptoEmpresas.TextMatrix(intRow, 2) = cboDepartamento.ItemData(cboDepartamento.ListIndex)
                    Else
                        cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, grdConceptoEmpresas.TextMatrix(intRow, 2))
                    End If
                Else
                    grdConceptoEmpresas.TextMatrix(intRow, 2) = cboDepartamento.ItemData(cboDepartamento.ListIndex)
                End If
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 2) = cboDepartamento.ItemData(cboDepartamento.ListIndex)
            End If
        Else
            grdConceptoEmpresas.TextMatrix(intRow, 2) = ""
        End If
    End If
End Sub

Private Function fintLocalizaRow()
    Dim intIndex As Integer
    
    If cboEmpresaContable.ListIndex > -1 Then
        For intIndex = 0 To grdConceptoEmpresas.Rows - 1
            If CLng(grdConceptoEmpresas.TextMatrix(intIndex, 0)) = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) Then
                fintLocalizaRow = intIndex
                Exit Function
            End If
        Next
        grdConceptoEmpresas.AddItem ""
        grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 0) = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
        grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 1) = txtCveConcepto.Text
        fintLocalizaRow = grdConceptoEmpresas.Rows - 1
        Exit Function
    Else
        fintLocalizaRow = -1
    End If
    
    fintLocalizaRow = -1
End Function

Private Function fintLocalizaRow1()
    Dim intIndex As Integer
        
    If cboDepartamento1.ListIndex > -1 Then
        For intIndex = 0 To grdExepcionesContables.Rows - 1
            If CLng(grdExepcionesContables.TextMatrix(intIndex, 1)) = cboDepartamento1.ItemData(cboDepartamento1.ListIndex) Then
                fintLocalizaRow1 = intIndex
                Exit Function
            End If
        Next
        grdExepcionesContables.AddItem ""
        grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 0) = txtCveConcepto.Text
        fintLocalizaRow1 = grdExepcionesContables.Rows - 1
        Exit Function
    Else
        fintLocalizaRow1 = -1
    End If
    
    fintLocalizaRow1 = -1
End Function

Private Function fintLocalizaRow2()
    Dim intIndex As Integer
    
    If cboPaciente.ListIndex > -1 Then
        For intIndex = 0 To grdExepcionesContables2.Rows - 1
            If CLng(grdExepcionesContables2.TextMatrix(intIndex, 1)) = cboPaciente.ItemData(cboPaciente.ListIndex) Then
                fintLocalizaRow2 = intIndex
                Exit Function
            End If
        Next
        grdExepcionesContables2.AddItem ""
        grdExepcionesContables2.TextMatrix(grdExepcionesContables2.Rows - 1, 0) = txtCveConcepto.Text
        fintLocalizaRow2 = grdExepcionesContables2.Rows - 1
        Exit Function
    Else
        fintLocalizaRow2 = -1
    End If
    fintLocalizaRow2 = -1
End Function

Private Sub cboDepartamento_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboDepartamento1_Click()
    Dim intRow  As Integer
    
    mskCuentaIngreso1.Mask = ""
    mskCuentaDescuento1.Mask = ""
    mskCuentaDescNota1.Mask = ""
    mskCuentaDescExSocial1.Mask = ""
    mskCuentaIngresosPendientes1.Mask = ""
    mskCuentaDescuentosPendientes1.Mask = ""
    
    mskCuentaIngreso1.Text = ""
    mskCuentaDescuento1.Text = ""
    mskCuentaDescNota1.Text = ""
    mskCuentaDescExSocial1.Text = ""
    mskCuentaIngresosPendientes1.Text = ""
    mskCuentaDescuentosPendientes1.Text = ""
    
    intRow = fintLocalizaRow1
    If intRow > -1 Then
        If cboDepartamento1.ListIndex > -1 Then
            grdExepcionesContables.TextMatrix(intRow, 1) = cboDepartamento1.ItemData(cboDepartamento1.ListIndex)
        Else
            grdExepcionesContables.TextMatrix(intRow, 2) = ""
        End If
        
        If cboEmpresaContable.ListIndex > -1 Then
            If IsNumeric(grdExepcionesContables.TextMatrix(intRow, 2)) Then
                mskCuentaIngreso1.Text = fstrCuentaContable(CLng(grdExepcionesContables.TextMatrix(intRow, 2)))
                txtCuentaIngreso1.Text = fstrDescripcionCuenta(mskCuentaIngreso1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
            
            If IsNumeric(grdExepcionesContables.TextMatrix(intRow, 3)) Then
                mskCuentaDescuento1.Text = fstrCuentaContable(CLng(grdExepcionesContables.TextMatrix(intRow, 3)))
                txtCuentaDescuento1.Text = fstrDescripcionCuenta(mskCuentaDescuento1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
            
            '- (CR) Agregado para caso 6808 -'
            If IsNumeric(grdExepcionesContables.TextMatrix(intRow, 4)) Then
                mskCuentaDescNota1.Text = fstrCuentaContable(CLng(grdExepcionesContables.TextMatrix(intRow, 4)))
                txtCuentaDescNota1.Text = fstrDescripcionCuenta(mskCuentaDescNota1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
            
            If IsNumeric(grdExepcionesContables.TextMatrix(intRow, 5)) Then
                mskCuentaIngresosPendientes1.Text = fstrCuentaContable(CLng(grdExepcionesContables.TextMatrix(intRow, 5)))
                txtCuentaIngresosPendientes1.Text = fstrDescripcionCuenta(mskCuentaIngresosPendientes1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
            
            If IsNumeric(grdExepcionesContables.TextMatrix(intRow, 6)) Then
                mskCuentaDescuentosPendientes1.Text = fstrCuentaContable(CLng(grdExepcionesContables.TextMatrix(intRow, 6)))
                txtCuentaDescuentosPendientes1.Text = fstrDescripcionCuenta(mskCuentaDescuentosPendientes1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
            
           
            If IsNumeric(grdExepcionesContables.TextMatrix(intRow, 7)) Then
                mskCuentaDescExSocial1.Text = fstrCuentaContable(CLng(grdExepcionesContables.TextMatrix(intRow, 7)))
                txtCuentaDescSocialEX1.Text = fstrDescripcionCuenta(mskCuentaDescExSocial1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
            
        End If
    End If
    
    mskCuentaIngreso1.Mask = txtEstructuraem.Text
    mskCuentaDescuento1.Mask = txtEstructuraem.Text
    mskCuentaDescNota1.Mask = txtEstructuraem.Text
    mskCuentaDescExSocial1.Mask = txtEstructuraem.Text
    mskCuentaIngresosPendientes1.Mask = txtEstructuraem.Text
    mskCuentaDescuentosPendientes1.Mask = txtEstructuraem.Text
End Sub

Private Sub cboDepartamento1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskCuentaIngreso1
    End If
End Sub

Private Sub cboDepartamento1_LostFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboEmpresaContable_Click()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim intIndex As Integer
    
    cboDepartamento.Clear
    cboDepartamento1.Clear
    cboPaciente.Clear
    
    '- Cuentas contables -'
    mskCuentaIngreso.Mask = ""
    mskCuentaDescuento.Mask = ""
    mskCuentaDescNota.Mask = ""
    mskCuentadescuentoCUSocial.Mask = ""
    mskCuentaIngresosPendientes.Mask = ""
    mskCuentaDescuentosPendientes.Mask = ""
    mskCuentaEficienciaPaquetes.Mask = ""
    mskCuentaDeficienciaPaquetes.Mask = ""
    
    mskCuentaIngreso.Text = ""
    mskCuentaDescuento.Text = ""
    mskCuentaDescNota.Text = ""
    mskCuentadescuentoCUSocial.Text = ""
    mskCuentaEficienciaPaquetes.Text = ""
    mskCuentaDeficienciaPaquetes.Text = ""
   
    
    mskCuentaIngresosPendientes.Text = ""
    mskCuentaDescuentosPendientes.Text = ""
    
    txtCuentaIngreso.Text = ""
    txtCuentaDescuento.Text = ""
    txtCuentaDescNota.Text = ""
    txtCuentaIngresosSocialCU = ""
    txtCuentaEficienciaPaquetes = ""
    txtCuentaDEficienciaPaquetes = ""
    
    '- Excepciones contables -'
    mskCuentaIngreso1.Mask = ""
    mskCuentaDescuento1.Mask = ""
    mskCuentaDescNota1.Mask = ""
    mskCuentaDescExSocial1.Mask = ""
    mskCuentaIngresosPendientes1.Mask = ""
    mskCuentaDescuentosPendientes1.Mask = ""
   
    
    mskCuentaIngreso1.Text = ""
    mskCuentaDescuento1.Text = ""
    mskCuentaDescNota1.Text = ""
    mskCuentaDescExSocial1.Text = ""
   
    mskCuentaIngresosPendientes1.Text = ""
    mskCuentaDescuentosPendientes1.Text = ""
   
    
    txtCuentaIngreso1.Text = ""
    txtCuentaDescuento1.Text = ""
    txtCuentaDescNota1.Text = ""
    txtCuentaDescSocialEX1 = ""
    
    
    '- Excepciones contables para socios -'
    mskCuentaIngreso2.Mask = ""
    mskCuentaDescuento2.Mask = ""
    mskCuentaIngreso2.Text = ""
    mskCuentaDescuento2.Text = ""
    txtCuentaIngreso2.Text = ""
    txtCuentaDescuento2.Text = ""
    
    txtEstructuraem.Text = ""
    
    If cboEmpresaContable.ListIndex > -1 Then
        strSQL = "SELECT smiCveDepartamento, vchDescripcion FROM NODepartamento WHERE bitEstatus <> 0 AND tnyClaveEmpresa = " & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) & " ORDER BY vchDescripcion"
        Set rs = frsRegresaRs(strSQL)
        pLlenarCboRs cboDepartamento, rs, 0, 1
        pLlenarCboRs cboDepartamento1, rs, 0, 1
        rs.Close
        
        'Carga el tipo de paciente configurado como socio
        strSQL = "SELECT SI.VCHVALOR, AD.VCHDESCRIPCION FROM SIPARAMETRO SI INNER JOIN ADTIPOPACIENTE AD ON AD.TNYCVETIPOPACIENTE = SI.VCHVALOR WHERE TRIM(VCHNOMBRE) = 'INTCVETIPOPACIENTESOCIO' AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
        Set rs = frsRegresaRs(strSQL)
        If rs.RecordCount > 0 Then
            pLlenarCboRs cboPaciente, rs, 0, 1
            cboPaciente.ListIndex = 0
        End If
        
        'strSQL = "select * from CnParametro where tnyClaveEmpresa=" & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
        'Set rs = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)
        Set rs = frsSelParametros("CN", cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex), "VCHESTRUCTURACUENTACONTABLE")
        If Not rs.EOF Then
            txtEstructuraem.Text = rs!valor
        End If
        rs.Close
        
        If grdConceptoEmpresas.Rows > 0 Then
        For intIndex = 0 To grdConceptoEmpresas.Rows - 1
            If CLng(grdConceptoEmpresas.TextMatrix(intIndex, 0)) = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) Then
                cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, grdConceptoEmpresas.TextMatrix(intIndex, 2))
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 3)) Then
                    mskCuentaIngreso.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 3)))
                    txtCuentaIngreso.Text = fstrDescripcionCuenta(mskCuentaIngreso.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
                
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 4)) Then
                    mskCuentaDescuento.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 4)))
                    txtCuentaDescuento.Text = fstrDescripcionCuenta(mskCuentaDescuento.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
                
                '- (CR) Agregado para caso 6808 -'
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 5)) Then
                    mskCuentaDescNota.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 5)))
                    txtCuentaDescNota.Text = fstrDescripcionCuenta(mskCuentaDescNota.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
                
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 6)) Then
                    mskCuentaIngresosPendientes.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 6)))
                    txtCuentaIngresosPendientes.Text = fstrDescripcionCuenta(mskCuentaIngresosPendientes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
                
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 7)) Then
                    mskCuentaDescuentosPendientes.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 7)))
                    txtCuentaDescuentosPendientes.Text = fstrDescripcionCuenta(mskCuentaDescuentosPendientes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
               
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 8)) Then
                    mskCuentadescuentoCUSocial.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 8)))
                    txtCuentaIngresosSocialCU.Text = fstrDescripcionCuenta(mskCuentadescuentoCUSocial.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
              
                
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 9)) Then
                    mskCuentaEficienciaPaquetes.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 9)))
                    txtCuentaEficienciaPaquetes.Text = fstrDescripcionCuenta(mskCuentaEficienciaPaquetes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
               
               
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 10)) Then
                    mskCuentaDeficienciaPaquetes.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 10)))
                    txtCuentaDEficienciaPaquetes.Text = fstrDescripcionCuenta(mskCuentaDeficienciaPaquetes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
               
            End If
        Next
    Else
        cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
    End If
    End If
    
    mskCuentaIngreso.Mask = txtEstructuraem.Text
    mskCuentaDescuento.Mask = txtEstructuraem.Text
    mskCuentaDescNota.Mask = txtEstructuraem.Text
    mskCuentaIngresosPendientes.Mask = txtEstructuraem.Text
    mskCuentaDescuentosPendientes.Mask = txtEstructuraem.Text
    mskCuentadescuentoCUSocial.Mask = txtEstructuraem.Text
    mskCuentaEficienciaPaquetes.Mask = txtEstructuraem.Text
    mskCuentaDeficienciaPaquetes.Mask = txtEstructuraem.Text
    
    mskCuentaIngreso1.Mask = txtEstructuraem.Text
    mskCuentaDescuento1.Mask = txtEstructuraem.Text
    mskCuentaDescNota1.Mask = txtEstructuraem.Text
    mskCuentaDescExSocial1.Mask = txtEstructuraem.Text
    mskCuentaIngresosPendientes1.Mask = txtEstructuraem.Text
    mskCuentaDescuentosPendientes1.Mask = txtEstructuraem.Text
    
    
            
    mskCuentaIngreso2.Mask = txtEstructuraem.Text
    mskCuentaDescuento2.Mask = txtEstructuraem.Text
End Sub

Private Sub cboEmpresaContable_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboIvas_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboIVAS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If fblnCanFocus(chkPredeterminadoPaquetes) Then
            chkPredeterminadoPaquetes.SetFocus
        Else
            If fblnCanFocus(chkActivo) Then
                chkActivo.SetFocus
            Else
                cmdGrabarRegistro.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboIVAS_KeyDown"))
    Unload Me
End Sub

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskCuentaIngreso
    End If
End Sub

Private Sub cboPaciente_Click()
    Dim intRow  As Integer
    
    mskCuentaIngreso2.Mask = ""
    mskCuentaDescuento2.Mask = ""
    mskCuentaIngreso2.Text = ""
    mskCuentaDescuento2.Text = ""
    
    intRow = fintLocalizaRow2
    If intRow > -1 Then
        If cboPaciente.ListIndex > -1 Then
            grdExepcionesContables2.TextMatrix(intRow, 1) = cboPaciente.ItemData(cboPaciente.ListIndex)
        Else
            grdExepcionesContables2.TextMatrix(intRow, 2) = ""
        End If
        
        If cboEmpresaContable.ListIndex > -1 Then
            If IsNumeric(grdExepcionesContables2.TextMatrix(intRow, 2)) Then
                mskCuentaIngreso2.Text = fstrCuentaContable(CLng(grdExepcionesContables2.TextMatrix(intRow, 2)))
                txtCuentaIngreso2.Text = fstrDescripcionCuenta(mskCuentaIngreso2.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
            
            If IsNumeric(grdExepcionesContables2.TextMatrix(intRow, 3)) Then
                mskCuentaDescuento2.Text = fstrCuentaContable(CLng(grdExepcionesContables2.TextMatrix(intRow, 3)))
                txtCuentaDescuento2.Text = fstrDescripcionCuenta(mskCuentaDescuento2.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            End If
        End If
    End If
    
    mskCuentaIngreso2.Mask = txtEstructuraem.Text
    mskCuentaDescuento2.Mask = txtEstructuraem.Text
End Sub

Private Sub cboPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskCuentaIngreso2
    End If
End Sub

Private Sub cboPaciente_LostFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboTipo_Click()
    If cboTipo.List(cboTipo.ListIndex) = "HOSPITALARIO" And Not vgblnNuevoRegistro Then
        cmdListasPrecio.Enabled = True
    Else
        cmdListasPrecio.Enabled = False
    End If
    If cboTipo.List(cboTipo.ListIndex) = "HOSPITALARIO" Then
        chkPredeterminadoPaquetes.Enabled = True
    Else
        chkPredeterminadoPaquetes.Value = 0
        chkPredeterminadoPaquetes.Enabled = False
    End If
    
    If cboTipo.List(cboTipo.ListIndex) = "ADMINISTRATIVO" Then
        chkExentoIVA.Enabled = True
    Else
        chkExentoIVA.Value = 0
        chkExentoIVA.Enabled = False
    End If
    
End Sub

Private Sub cboTipo_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkActivo_Click()
On Error GoTo NotificaError
    
    If vgblnNuevoRegistro Then
        chkActivo.Value = 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_Click"))
    Unload Me
End Sub

Private Sub chkActivo_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkActivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkExentoIVA_Click()
    If chkExentoIVA.Value = 1 Then
        cboIvas.ListIndex = -1
        cboIvas.Enabled = False
    Else
        If cboIvas.ListCount > 0 And vgblnNuevoRegistro Then
            cboIvas.ListIndex = 0
        End If
        cboIvas.Enabled = True
    End If
End Sub

Private Sub chkExentoIVA_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkExentoIVA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If chkExentoIVA.Value = 1 Then
            If chkActivo.Enabled = True Then
                chkActivo.SetFocus
            Else
                cmdGrabarRegistro.SetFocus
            End If
        Else
            cboIvas.SetFocus
        End If
    End If
End Sub


Private Sub chkPredeterminadoPaquetes_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkPredeterminadoPaquetes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If chkActivo.Enabled = True Then
            chkActivo.SetFocus
        Else
            cmdGrabarRegistro.SetFocus
        End If
    End If
End Sub


Private Sub chkSaldarCuentas_Click()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkSaldarCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo NotificaError
    
    sstObj.Tab = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
    Unload Me
End Sub

Private Sub cmdCuentas_Click()
    pCargaConceptos
    
    cboEmpresaContable.ListIndex = -1
    cboEmpresaContable.ListIndex = flngLocalizaCbo(cboEmpresaContable, CStr(vgintClaveEmpresaContable))
    
    sstObj.Tab = 2
    cboEmpresaContable.SetFocus
End Sub

Private Sub pCargaConceptos()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    If txtConceptosDe.Text <> txtCveConcepto.Text Then
        txtConceptosDe.Text = txtCveConcepto.Text
        grdConceptoEmpresas.Rows = 0
        grdExepcionesContables.Rows = 0
        grdExepcionesContables2.Rows = 0
        
        strSQL = "SELECT * FROM PVConceptoFacturacionEmpresa WHERE intCveConceptoFactura = " & txtCveConcepto.Text
        Set rs = frsRegresaRs(strSQL)
        Do Until rs.EOF
            grdConceptoEmpresas.AddItem ""
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 0) = rs!intCveEmpresaContable
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 1) = rs!intCveConceptoFactura
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 2) = rs!intCveDepartamento
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 3) = IIf(IsNull(rs!intNumCtaIngreso), "", rs!intNumCtaIngreso)
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 4) = IIf(IsNull(rs!intNumCtaDescuento), "", rs!intNumCtaDescuento)
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 5) = IIf(IsNull(rs!intNumCtaDescNota), "", rs!intNumCtaDescNota) '(CR) Agregado para caso 6808
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 6) = IIf(IsNull(rs!INTNUMCTAINGRESOPENDIENTE), "", rs!INTNUMCTAINGRESOPENDIENTE)
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 7) = IIf(IsNull(rs!INTNUMCTADESCUENTOPENDIENTE), "", rs!INTNUMCTADESCUENTOPENDIENTE)
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 8) = IIf(IsNull(rs!INTNUMCTADESCSOCIAL), "", rs!INTNUMCTADESCSOCIAL)
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 9) = IIf(IsNull(rs!INTNUMCTAEFICIENCIAPAQUETE), "", rs!INTNUMCTAEFICIENCIAPAQUETE)
            grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 10) = IIf(IsNull(rs!INTNUMCTADEFICIENCIAPAQUETE), "", rs!INTNUMCTADEFICIENCIAPAQUETE)
            rs.MoveNext
        Loop
        rs.Close
        
        strSQL = "SELECT * FROM PVConceptoFacturacionDepartame WHERE smiCveConcepto = " & txtCveConcepto.Text
        Set rs = frsRegresaRs(strSQL)
        Do Until rs.EOF
            grdExepcionesContables.AddItem ""
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 0) = rs!smicveconcepto
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 1) = rs!smicvedepartamento
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 2) = IIf(IsNull(rs!intNumCuentaIngreso), "", rs!intNumCuentaIngreso)
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 3) = IIf(IsNull(rs!INTNUMCUENTADESCUENTO), "", rs!INTNUMCUENTADESCUENTO)
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 4) = IIf(IsNull(rs!intNumCuentaDescNota), "", rs!intNumCuentaDescNota) '(CR) Agregado para caso 6808
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 5) = IIf(IsNull(rs!INTNUMCTAINGRESOPENDIENTE), "", rs!INTNUMCTAINGRESOPENDIENTE)
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 6) = IIf(IsNull(rs!INTNUMCTADESCUENTOPENDIENTE), "", rs!INTNUMCTADESCUENTOPENDIENTE)
            grdExepcionesContables.TextMatrix(grdExepcionesContables.Rows - 1, 7) = IIf(IsNull(rs!INTNUMCTADESCSOCIAL), "", rs!INTNUMCTADESCSOCIAL)
            
            rs.MoveNext
        Loop
        rs.Close
        
        strSQL = "SELECT * FROM PVConceptoFactPaciente WHERE smiCveConcepto = " & txtCveConcepto.Text
        Set rs = frsRegresaRs(strSQL)
        Do Until rs.EOF
            grdExepcionesContables2.AddItem ""
            grdExepcionesContables2.TextMatrix(grdExepcionesContables2.Rows - 1, 0) = rs!smicveconcepto
            grdExepcionesContables2.TextMatrix(grdExepcionesContables2.Rows - 1, 1) = rs!SMICVETIPOPACIENTE
            grdExepcionesContables2.TextMatrix(grdExepcionesContables2.Rows - 1, 2) = IIf(IsNull(rs!intNumCuentaIngreso), "", rs!intNumCuentaIngreso)
            grdExepcionesContables2.TextMatrix(grdExepcionesContables2.Rows - 1, 3) = IIf(IsNull(rs!INTNUMCUENTADESCUENTO), "", rs!INTNUMCUENTADESCUENTO)
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ValidaIntegridad
    
    Dim vllngPersonaGraba As Long
    Dim vlstrSentecia As String
    Dim rsIntegridad As ADODB.Recordset 'Permitirá verificar si se puede borrar o no este concepto de facturación


    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        '----------------------------------------------------------------------'
        ' Verificación de que el concepto no esté en el catálogo de descuentos '
        '----------------------------------------------------------------------'
        vlstrSentecia = "SELECT COUNT(smiCveConcepto) NumRegistros " & _
                        " FROM PvDescuento " & _
                        " WHERE chrTipoCargo = 'CF' " & _
                        " AND smiCveConcepto = " & txtCveConcepto.Text
        Set rsIntegridad = frsRegresaRs(vlstrSentecia, adLockReadOnly, adOpenForwardOnly)
        If rsIntegridad!numRegistros > 0 Then
            rsIntegridad.Close
            'No se puede borrar. Tiene otros datos relacionados
            MsgBox SIHOMsg(672) & ":" & Chr(13) & "Descuentos.", vbCritical, "Mensaje"
            Exit Sub
        Else
            rsIntegridad.Close
        End If
        
        '---------------------------------------------------------------------'
        ' Verificación de que el concepto no esté en el catálogo de Artículos '
        '---------------------------------------------------------------------'
        vlstrSentecia = "SELECT COUNT(smiCveConceptFact) NumRegistros " & _
                        " FROM IvArticulo " & _
                        " WHERE smiCveConceptFact = " & txtCveConcepto.Text
        Set rsIntegridad = frsRegresaRs(vlstrSentecia, adLockReadOnly, adOpenForwardOnly)
        If rsIntegridad!numRegistros > 0 Then
            rsIntegridad.Close
            'No se puede borrar. Tiene otros datos relacionados
            MsgBox SIHOMsg(672) & ":" & Chr(13) & "Artículos.", vbCritical, "Mensaje"
            Exit Sub
        Else
            rsIntegridad.Close
        End If
        
        '-----------------------------------------------------------------------'
        ' Verificación de que el concepto no esté en LAS NOTAS DE CARGO-CREDITO '
        '-----------------------------------------------------------------------'
        vlstrSentecia = "SELECT COUNT(intConcepto) NumRegistros " & _
                        " FROM CcNotaDetalle " & _
                        " WHERE chrTipoCargo = 'CF' AND intConcepto = " & txtCveConcepto.Text
        Set rsIntegridad = frsRegresaRs(vlstrSentecia, adLockReadOnly, adOpenForwardOnly)
        If rsIntegridad!numRegistros > 0 Then
            rsIntegridad.Close
            'No se puede borrar. Tiene otros datos relacionados
            MsgBox SIHOMsg(672) & ":" & Chr(13) & "Notas de crédito o cargo.", vbCritical, "Mensaje"
            Exit Sub
        Else
            rsIntegridad.Close
        End If
        
        '-----------------------------------------------------------------------'
        ' Verificación de que el concepto no esté en OTROS CONCEPTOS '
        '-----------------------------------------------------------------------'
        vlstrSentecia = "SELECT COUNT(smiConceptoFact) NumRegistros " & _
                        " FROM PvOtroConcepto " & _
                        " WHERE smiConceptoFact = " & txtCveConcepto.Text
        Set rsIntegridad = frsRegresaRs(vlstrSentecia, adLockReadOnly, adOpenForwardOnly)
        If rsIntegridad!numRegistros > 0 Then
            rsIntegridad.Close
            'No se puede borrar. Tiene otros datos relacionados
            MsgBox SIHOMsg(672) & ":" & Chr(13) & "Otros conceptos de cargo.", vbCritical, "Mensaje"
            Exit Sub
        Else
            rsIntegridad.Close
        End If
        
        '-----------------------------------------------------------------------'
        ' Verificación de que el concepto no esté en ESTUDIOS '
        '-----------------------------------------------------------------------'
        vlstrSentecia = "SELECT COUNT(smiConFact) NumRegistros " & _
                        " FROM ImEstudio " & _
                        " WHERE smiConFact = " & txtCveConcepto.Text
        Set rsIntegridad = frsRegresaRs(vlstrSentecia, adLockReadOnly, adOpenForwardOnly)
        If rsIntegridad!numRegistros > 0 Then
            rsIntegridad.Close
            'No se puede borrar. Tiene otros datos relacionados
            MsgBox SIHOMsg(672) & ":" & Chr(13) & "Estudios.", vbCritical, "Mensaje"
            Exit Sub
        Else
            rsIntegridad.Close
        End If
        
        '-----------------------------------------------------------------------'
        ' Verificación de que el concepto no esté en EXAMENES '
        '-----------------------------------------------------------------------'
        vlstrSentecia = "SELECT COUNT(smiConFact) NumRegistros " & _
                        " FROM LaExamen " & _
                        " WHERE smiConFact = " & txtCveConcepto.Text
        Set rsIntegridad = frsRegresaRs(vlstrSentecia, adLockReadOnly, adOpenForwardOnly)
        If rsIntegridad!numRegistros > 0 Then
            rsIntegridad.Close
            'No se puede borrar. Tiene otros datos relacionados
            MsgBox SIHOMsg(672) & ":" & Chr(13) & "Exámenes.", vbCritical, "Mensaje"
            Exit Sub
        Else
            rsIntegridad.Close
        End If
        
        '-----------------------------------------------------------------------'
        ' Verificación de que el concepto no esté en GRUPOS EXAMENES '
        '-----------------------------------------------------------------------'
        vlstrSentecia = "SELECT COUNT(smiConFact) NumRegistros " & _
                        " FROM LaGrupoExamen " & _
                        " WHERE smiConFact = " & txtCveConcepto.Text
        Set rsIntegridad = frsRegresaRs(vlstrSentecia, adLockReadOnly, adOpenForwardOnly)
        If rsIntegridad!numRegistros > 0 Then
            rsIntegridad.Close
            'No se puede borrar. Tiene otros datos relacionados
            MsgBox SIHOMsg(672) & ":" & Chr(13) & "Grupos de exámenes.", vbCritical, "Mensaje"
            Exit Sub
        Else
            rsIntegridad.Close
        End If
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        frsEjecuta_SP txtCveConcepto.Text, "Sp_PVDelCuentasConcFact"
        rsConceptos.Delete
        rsConceptos.Requery
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "CONCEPTO DE FACTURACION", txtCveConcepto.Text)
        pNuevoRegistro
    End If
    
    Exit Sub
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdGrabarRegistro_Click()
On Error GoTo NotificaError
'***************************************************'
' Procedimiento para grabar una alta o modificacion '
'***************************************************'

    Dim vlintContador As Integer
    Dim vllngPersonaGraba As Long, vllngSecuencia As Long
    Dim vllngNumeroCuenta As Long
    Dim vlintErrorCuenta As Integer
    Dim vlintPorcentajeIVA As Integer
    Dim vlblnSeCambioLista As Boolean
    Dim rsChecarListaModificada As New ADODB.Recordset
    Dim vlstrQuery As String
    
    If Not fblnDatosValidos Then Exit Sub
   
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        With rsConceptos
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            '-----------------------------------'
            ' Grabar el Concepto de Facturación '
            '-----------------------------------'
            If vgblnNuevoRegistro Then
                .AddNew
            End If
            If rsConceptos.Fields("smiCveConcepto").Attributes <> 16 And rsConceptos.Fields("smiCveConcepto").Attributes <> 32784 Then
                !smicveconcepto = Me.txtCveConcepto.Text
            End If
            If chkExentoIVA.Value = 1 Then
                vlintPorcentajeIVA = 0
            Else
                vlintPorcentajeIVA = cboIvas.ItemData(cboIvas.ListIndex)
            End If
            !chrdescripcion = Trim(txtDescripcion.Text)
            !smyIVA = vlintPorcentajeIVA
            !intTipo = cboTipo.ListIndex
            !bitactivo = chkActivo.Value
            !bitsaldarcuentas = chkSaldarCuentas.Value
            !bitExentoIva = chkExentoIVA.Value
            !bitPaquetePresupuesto = chkPredeterminadoPaquetes.Value
            On Error GoTo UpdateErr
            .Update
            
            pGuardaDetalle
            
            If cboTipo.List(cboTipo.ListIndex) = "HOSPITALARIO" And Not vgblnNuevoRegistro Then
                'Se modifico para que solo aparezca el mensaje si se modificaron las listas
                vlblnSeCambioLista = False
                If fCantidadElementos > 0 Then
                    For vlintContador = 1 To UBound(aListaPrecio) - 1
                        Set rsChecarListaModificada = frsRegresaRs("SELECT COUNT(*) TOTAL FROM PVPOLITICALISTAPRECIOCONCEPTO WHERE PVPOLITICALISTAPRECIOCONCEPTO.SMICVECONCEPTO = " & Me.txtCveConcepto.Text & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTCVELISTA = " & aListaPrecio(vlintContador).vllngClaveLista)
                        If rsChecarListaModificada.RecordCount > 0 Then
                            'Verifica que la lista existe, sino la da de alta
                            If rsChecarListaModificada!Total > 0 Then
                                rsChecarListaModificada.Close
                        
                                vlstrQuery = "SELECT COUNT(*) TOTAL FROM PVPOLITICALISTAPRECIOCONCEPTO "
                                vlstrQuery = vlstrQuery & "WHERE PVPOLITICALISTAPRECIOCONCEPTO.SMICVECONCEPTO = " & Me.txtCveConcepto.Text
                                vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTCVELISTA = " & aListaPrecio(vlintContador).vllngClaveLista
                                vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.CHRTIPOINCREMENTO = '" & CStr(aListaPrecio(vlintContador).vlstrTipoIncremento) & "' "
                                vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.NUMMARGENUTILIDAD = " & aListaPrecio(vlintContador).vldblmargenutilidad
                                vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTINCREMENTO = " & IIf(aListaPrecio(vlintContador).vlblnIncrementoAutomatico = True, 1, 0)
                                vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTTABULADOR = " & IIf(aListaPrecio(vlintContador).vlblnUsaTabulador = True, 1, 0)
                                vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.MNYPRECIO = " & aListaPrecio(vlintContador).vldblPrecio
                                Set rsChecarListaModificada = frsRegresaRs(vlstrQuery)
                                'Verifica que los valores del arreglo sean iguales a los que estan en la tabla, si son diferentes modifica la lista y los artículos
                                If rsChecarListaModificada.RecordCount > 0 Then
                                    If rsChecarListaModificada!Total = 0 Then
                                        vlblnSeCambioLista = True
                                        Exit For
                                    End If
                                End If
                            Else
                                vlblnSeCambioLista = True
                                Exit For
                            End If
                        End If
                    Next vlintContador
                End If
                '-------------------------
                If vlblnSeCambioLista = True Then
                    'Se actualizarán los datos en las listas de precios modificadas. ¿Desea continuar?
                    If MsgBox("Se actualizarán los datos en las listas de precios modificadas. ¿Desea continuar?", vbYesNo + vbCritical, "Mensaje") = vbYes Then
                        pGuardaListaPrecios vllngPersonaGraba, True
                    Else
                        pGuardaListaPrecios vllngPersonaGraba, False
                    End If
                End If
            End If
            
            On Error GoTo NotificaError
            If vgblnNuevoRegistro Then
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "CONCEPTO DE FACTURACION", txtCveConcepto.Text)
            Else
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "CONCEPTO DE FACTURACION", txtCveConcepto.Text)
            End If
            
            EntornoSIHO.ConeccionSIHO.CommitTrans
        End With
        rsConceptos.Requery
        pNuevoRegistro
  End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Unload Me
    Exit Sub
UpdateErr:
    MsgBox SIHOMsg(649), , "Mensaje"
    If rsConceptos.State = 1 Then
        If Not (rsConceptos.BOF Or rsConceptos.EOF) Then
            rsConceptos.CancelUpdate
        End If
    End If
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    blnEnfocando = True
    pEnfocaMkTexto txtCveConcepto
End Sub

Private Sub pGuardaDetalle()
    Dim intIndex As Integer
    Dim strParametros As String
    
    For intIndex = 0 To grdConceptoEmpresas.Rows - 1
        'If grdConceptoEmpresas.TextMatrix(intIndex, 2) <> "" And grdConceptoEmpresas.TextMatrix(intIndex, 3) <> "" And grdConceptoEmpresas.TextMatrix(intIndex, 4) <> "" And grdConceptoEmpresas.TextMatrix(intIndex, 5) <> "" Then
        '    strParametros = grdConceptoEmpresas.TextMatrix(intIndex, 0) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 2) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 1) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 3) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 4) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 5)
        'Else
        '    strParametros = grdConceptoEmpresas.TextMatrix(intIndex, 0) & "|-1|" & grdConceptoEmpresas.TextMatrix(intIndex, 1) & "|0|0|0"
        'End If
        
'        If grdConceptoEmpresas.TextMatrix(intIndex, 2) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 3) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 4) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 5) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 6) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 7) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 8) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 9) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 10) = "" Then
'            grdConceptoEmpresas.TextMatrix(intIndex, 2) = "-1"
'        End If
        If grdConceptoEmpresas.TextMatrix(intIndex, 2) = "" Then
'            grdConceptoEmpresas.TextMatrix(intIndex, 2) = "-1"
            MsgBox SIHOMsg(242), vbExclamation, "Mensaje"
        Else
'            If grdConceptoEmpresas.TextMatrix(intIndex, 3) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 4) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 5) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 6) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 7) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 8) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 9) = "" And grdConceptoEmpresas.TextMatrix(intIndex, 10) = "" Then
'                grdConceptoEmpresas.TextMatrix(intIndex, 2) = "-1"
'                MsgBox SIHOMsg(211), vbExclamation, "Mensaje"
'            End If
        strParametros = grdConceptoEmpresas.TextMatrix(intIndex, 0) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 2) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 1) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 3) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 4) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 5) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 6) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 7) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 8) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 9) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 10)
        frsEjecuta_SP strParametros, "Sp_PvUpdConcFactDetEmpresa"
        End If
        
'        strParametros = grdConceptoEmpresas.TextMatrix(intIndex, 0) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 2) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 1) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 3) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 4) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 5) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 6) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 7) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 8) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 9) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 10)
'        frsEjecuta_SP strParametros, "Sp_PvUpdConcFactDetEmpresa"
    Next
    
    For intIndex = 0 To grdExepcionesContables.Rows - 1
        'If grdExepcionesContables.TextMatrix(intIndex, 1) <> "" And grdExepcionesContables.TextMatrix(intIndex, 2) <> "" And grdExepcionesContables.TextMatrix(intIndex, 3) <> "" Then
        '    strParametros = grdExepcionesContables.TextMatrix(intIndex, 1) & "|" & grdExepcionesContables.TextMatrix(intIndex, 0) & "|" & grdExepcionesContables.TextMatrix(intIndex, 2) & "|" & grdExepcionesContables.TextMatrix(intIndex, 3)
        'Else
        '    strParametros = grdExepcionesContables.TextMatrix(intIndex, 1) & "|" & grdExepcionesContables.TextMatrix(intIndex, 0) & "|-1|0"
        'End If
        
 
        If grdExepcionesContables.TextMatrix(intIndex, 2) = "" And grdExepcionesContables.TextMatrix(intIndex, 3) = "" And grdExepcionesContables.TextMatrix(intIndex, 4) = "" And grdExepcionesContables.TextMatrix(intIndex, 5) = "" And grdExepcionesContables.TextMatrix(intIndex, 6) = "" And grdExepcionesContables.TextMatrix(intIndex, 7) = "" Then
            grdExepcionesContables.TextMatrix(intIndex, 2) = "-1"
        End If
        
        strParametros = grdExepcionesContables.TextMatrix(intIndex, 1) & "|" & grdExepcionesContables.TextMatrix(intIndex, 0) & "|" & grdExepcionesContables.TextMatrix(intIndex, 2) & "|" & grdExepcionesContables.TextMatrix(intIndex, 3) & "|" & grdExepcionesContables.TextMatrix(intIndex, 4) & "|" & grdExepcionesContables.TextMatrix(intIndex, 5) & "|" & grdExepcionesContables.TextMatrix(intIndex, 6) & "|" & grdExepcionesContables.TextMatrix(intIndex, 7)
        frsEjecuta_SP strParametros, "Sp_PvUpdConcFactDetExcepcio"
    Next
    
    If vlUsaSocios = True Then
        For intIndex = 0 To grdExepcionesContables2.Rows - 1
            'If grdExepcionesContables2.TextMatrix(intIndex, 1) <> "" And grdExepcionesContables2.TextMatrix(intIndex, 2) <> "" And grdExepcionesContables2.TextMatrix(intIndex, 3) <> "" Then
            '    strParametros = grdExepcionesContables2.TextMatrix(intIndex, 1) & "|" & grdExepcionesContables2.TextMatrix(intIndex, 0) & "|" & grdExepcionesContables2.TextMatrix(intIndex, 2) & "|" & grdExepcionesContables2.TextMatrix(intIndex, 3)
            'Else
            '    strParametros = grdExepcionesContables2.TextMatrix(intIndex, 1) & "|" & grdExepcionesContables2.TextMatrix(intIndex, 0) & "|-1|0"
            'End If
            
            If grdExepcionesContables2.TextMatrix(intIndex, 2) = "" Then grdExepcionesContables2.TextMatrix(intIndex, 2) = "-1"
            If grdExepcionesContables2.TextMatrix(intIndex, 3) = "" Then grdExepcionesContables2.TextMatrix(intIndex, 3) = "0"
            
            strParametros = grdExepcionesContables2.TextMatrix(intIndex, 1) & "|" & grdExepcionesContables2.TextMatrix(intIndex, 0) & "|" & grdExepcionesContables2.TextMatrix(intIndex, 2) & "|" & grdExepcionesContables2.TextMatrix(intIndex, 3)
            frsEjecuta_SP strParametros, "Sp_PVUpdConcFactDetExcepcioPac"
        Next
    End If
End Sub

Private Sub cmdListasPrecio_Click()
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim rsListasPrecios As ADODB.Recordset
    Dim vlIntContArray As Integer   'contador arreglo
    Dim vlIntContGrid As Integer    'contador grid
    Dim vlIntColGrid As Integer      'columna
    Dim rsConceptoEmpresa As ADODB.Recordset
    
    vlstrSentencia = "SELECT COUNT(*) CONTADOR FROM PVCONCEPTOFACTURACIONEMPRESA WHERE INTCVECONCEPTOFACTURA = " & txtCveConcepto.Text & " AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
    Set rsConceptoEmpresa = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsConceptoEmpresa.RecordCount <> 0 Then
        If rsConceptoEmpresa!contador = 0 Then
            'No se encontraron listas de precios activas del departamento del concepto.
            MsgBox SIHOMsg(1614), vbExclamation + vbOKOnly, "Mensaje"
        Exit Sub
        End If
    End If
    
    vlstrSentencia = txtCveConcepto.Text & "|" & vgintClaveEmpresaContable
    Set rsListasPrecios = frsEjecuta_SP(vlstrSentencia, "SP_PVSELPOLITICALSTPRECIO")
    
    If rsListasPrecios.RecordCount = 0 Then ' no hay listas de precios dadas de alta o activas
        '¡No se encontraron listas de precios configuradas o activas!
        MsgBox SIHOMsg(1218), vbExclamation + vbOKOnly, "Mensaje"
        Exit Sub ' para afuera
    Else ' si hay listas dadas de alta, entonces se deben agregar al datagrid
        Load frmListasPreciosConcepto
        pLlenarMshFGrdRs frmListasPreciosConcepto.grdPrecios, rsListasPrecios
    End If
    
    With frmListasPreciosConcepto
         If fCantidadElementos > 0 Then 'se debe de colocar la informacion del arreglo en el data grid
            For vlIntContArray = 0 To UBound(aListaPrecio())
                For vlIntContGrid = 1 To .grdPrecios.Rows - 1
                    If .grdPrecios.TextMatrix(vlIntContGrid, 1) = aListaPrecio(vlIntContArray).vllngClaveLista Then
                       .grdPrecios.TextMatrix(vlIntContGrid, 3) = IIf(aListaPrecio(vlIntContArray).vlblnIncrementoAutomatico = True, 1, 0)
                       .grdPrecios.TextMatrix(vlIntContGrid, 4) = IIf(aListaPrecio(vlIntContArray).vlstrTipoIncremento = "M", "PRECIO", (IIf(aListaPrecio(vlIntContArray).vlstrTipoIncremento = "C", "ÚLTIMA", "COMPRA")))
                       .grdPrecios.TextMatrix(vlIntContGrid, 5) = aListaPrecio(vlIntContArray).vldblmargenutilidad
                       .grdPrecios.TextMatrix(vlIntContGrid, 6) = IIf(aListaPrecio(vlIntContArray).vlblnUsaTabulador = True, 1, 0)
                       .grdPrecios.TextMatrix(vlIntContGrid, 7) = FormatCurrency(aListaPrecio(vlIntContArray).vldblPrecio, 4)
                       .grdPrecios.TextMatrix(vlIntContGrid, 8) = IIf(aListaPrecio(vlIntContArray).vlblnPredeterminada = True, 1, 0)
                       .grdPrecios.TextMatrix(vlIntContGrid, 9) = IIf(aListaPrecio(vlIntContArray).vlblnNuevoEnLista = True, 1, 0)
                       Exit For
                    End If
                Next vlIntContGrid
            Next vlIntContArray
         End If
 
    For vlIntContGrid = 1 To .grdPrecios.Rows - 1
        If Val(.grdPrecios.TextMatrix(vlIntContGrid, 7)) = 0 Then
        ' a pintar de rojo la información de la lista
        For vlIntColGrid = 3 To .grdPrecios.Cols - 1
            .grdPrecios.Col = vlIntColGrid
            .grdPrecios.Row = vlIntContGrid
            .grdPrecios.CellForeColor = &HC0&
            .grdPrecios.CellFontBold = True
         Next vlIntColGrid
        End If
        
        If .grdPrecios.TextMatrix(vlIntContGrid, 8) = 1 Then
            ' a pintar de azul el nombre de la lista predeterminada
            .grdPrecios.Col = 1
            .grdPrecios.Row = vlIntContGrid
            .grdPrecios.CellForeColor = &HC00000
            .grdPrecios.CellFontBold = True
            .grdPrecios.Col = 2
            .grdPrecios.Row = vlIntContGrid
            .grdPrecios.CellForeColor = &HC00000
            .grdPrecios.CellFontBold = True
        End If
        
            'formato a la información
            .grdPrecios.TextMatrix(vlIntContGrid, 3) = IIf(.grdPrecios.TextMatrix(vlIntContGrid, 3) = 1, "*", "")
            .grdPrecios.TextMatrix(vlIntContGrid, 5) = Format(.grdPrecios.TextMatrix(vlIntContGrid, 5), "0.0000") & "%"
            .grdPrecios.TextMatrix(vlIntContGrid, 6) = IIf(.grdPrecios.TextMatrix(vlIntContGrid, 6) = 1, "*", "")
            .grdPrecios.TextMatrix(vlIntContGrid, 7) = FormatCurrency(.grdPrecios.TextMatrix(vlIntContGrid, 7), 2)
        Next vlIntContGrid
    End With
    
    pInicializaFormaListasPrecios
    pConfGridListasPrecios
    frmListasPreciosConcepto.Show vbModal
    
    If frmListasPreciosConcepto.vgblnCancel Then ' si se cancelo la accion no hay cambios en el arreglo y se cierra el formulario
        frmListasPreciosConcepto.vgblnCancel = False
    Else ' aqui si se tienen que 'TODOS' guardar los datos del grid en el arreglo, empezamos inicializando el arreglo
            
        Erase aListaPrecio
        ReDim aListaPrecio(0)
        ReDim aListaPrecio(frmListasPreciosConcepto.grdPrecios.Rows)
        For vlIntContGrid = 1 To frmListasPreciosConcepto.grdPrecios.Rows - 1
            aListaPrecio(vlIntContGrid).vllngClaveLista = frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 1)
            aListaPrecio(vlIntContGrid).vlblnIncrementoAutomatico = IIf(frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 3) = "*", 1, 0)
            aListaPrecio(vlIntContGrid).vlstrTipoIncremento = IIf(frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 4) = "PRECIO", "M", IIf(frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 4) = "ÚLTIMA", "C", "A"))
            aListaPrecio(vlIntContGrid).vldblmargenutilidad = Replace(frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 5), "%", "")
            aListaPrecio(vlIntContGrid).vlblnUsaTabulador = IIf(frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 6) = "*", 1, 0)
            aListaPrecio(vlIntContGrid).vldblPrecio = IIf(Val(Replace(frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 7), "$", "")) = 0, 0, Format(Replace(frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 7), "$", ""), "#############.####"))
            aListaPrecio(vlIntContGrid).vlblnPredeterminada = frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 8)
            aListaPrecio(vlIntContGrid).vlblnNuevoEnLista = frmListasPreciosConcepto.grdPrecios.TextMatrix(vlIntContGrid, 9)
        Next vlIntContGrid
            pHabilita 0, 0, 0, 0, 0, 1, 0
    End If
    Unload frmListasPreciosConcepto
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdListasPrecio_Click"))
End Sub

Private Sub Form_Activate()
    vgstrNombreForm = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyEscape Then
        KeyCode = 7
        Unload Me
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
        
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
        
    Me.Icon = frmMenuPrincipal.Icon
    Me.sstObj.Tab = 0
    
    Me.Height = vllngSizeNormal
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    'Se inicializa la variable de uso de socios
    vlUsaSocios = False
    
    'Se verifica si se utilizan Socios en la empresa
    vlstrsql = "SELECT vchValor FROM SiParametro WHERE TRIM(vchNombre) = 'BITUTILIZASOCIOS'"
    Set rs = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount > 0 Then
        If rs!vchvalor = "1" Then
            vlUsaSocios = True
            Frame4.Enabled = True
        Else
            vlUsaSocios = False
            Frame4.Enabled = False
        End If
        rs.Close
    End If
    Frame4.Visible = vlUsaSocios
    
    pCargaCombos
    blnClaveManualCatalogo = fblnClaveManualCatalogo("CONCEPTOS DE FACTURACION")

    vgstrNombreForm = Me.Name
    
    vlstrsql = "SELECT * FROM PvConceptoFacturacion"
    Set rsConceptos = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    pCargaIvas
    
    cboTipo.AddItem "HOSPITALARIO", cintTipoHospital
    cboTipo.AddItem "ASEGURADORA", cintTipoSeguros
    cboTipo.AddItem "ADMINISTRATIVO", cintTipoAdmitivo
    pHabilita 1, 1, 1, 1, 1, 0, 0
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub pCargaCombos()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT * FROM CNEmpresaContable WHERE bitActiva <> 0 ORDER BY vchNombre"
    Set rs = frsRegresaRs(strSQL)
    If Not rs.EOF Then
        pLlenarCboRs cboEmpresaContable, rs, 0, 1
    End If
    rs.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError
    
    If sstObj.Tab <> 0 Then
        sstObj.Tab = 0
        If txtDescripcion.Enabled Then
            txtDescripcion.SetFocus
        Else
            txtCveConcepto.SetFocus
        End If
        Cancel = True
    Else
        If Me.Height = vllngSizeGrande Then
            Cancel = True
            
            Me.Height = vllngSizeNormal
            Me.Top = Int((vglngDesktop - Me.Height) / 2)
            
            fraBotonera.Enabled = True
            fraConcepto.Enabled = True
        Else
            If cmdGrabarRegistro.Enabled = True Then
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    txtDescripcion.SetFocus
                    txtCveConcepto.SetFocus
                End If
                Cancel = True
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub pNuevoRegistro()
On Error GoTo NotificaError

    txtCveConcepto.Text = fintSigNumRs(rsConceptos, 0)
    txtDescripcion.Text = ""
    txtDescripcion.Enabled = False
    cboIvas.Enabled = False
    
    cboTipo.ListIndex = -1
    cboTipo.Enabled = False
    
    chkActivo.Enabled = False
    cmdCuentas.Enabled = False
    cmdListasPrecio.Enabled = False
    chkSaldarCuentas.Value = 0
    chkSaldarCuentas.Enabled = False
    chkExentoIVA.Value = 0
    chkExentoIVA.Enabled = True
    chkPredeterminadoPaquetes.Value = 0
    chkPredeterminadoPaquetes.Enabled = False
    
    If cboIvas.ListCount > 0 Then
        cboIvas.ListIndex = 0
    Else
        MsgBox SIHOMsg(13) + Chr(13) + cboIvas.ToolTipText, vbExclamation, "Mensaje"
        Unload Me
        Exit Sub
    End If
    
    chkActivo.Value = 1
    vgblnNuevoRegistro = True
    
    pHabilita 1, 1, 1, 1, 1, 0, 0
    txtConceptosDe.Text = ""
    
    Call pEnfocaMkTexto(txtCveConcepto)
    pConfiguraGrid
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pNuevoRegistro"))
    Unload Me
End Sub

Private Sub pLlenaGrid()
On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsConceptos As New ADODB.Recordset
    Dim vlintContador As Integer
    
    GrdHBusqueda.Clear
    GrdHBusqueda.Rows = 2
    GrdHBusqueda.Cols = 3
    
    vgstrParametrosSP = "0|-1|-1"
    Set rsConceptos = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelConceptoFactura")
    If rsConceptos.RecordCount > 0 Then
        vlintContador = 1
        Do While Not rsConceptos.EOF
            GrdHBusqueda.TextMatrix(vlintContador, 1) = rsConceptos!smicveconcepto
            GrdHBusqueda.TextMatrix(vlintContador, 2) = rsConceptos!chrdescripcion
            vlintContador = vlintContador + 1
            GrdHBusqueda.Rows = GrdHBusqueda.Rows + 1
            rsConceptos.MoveNext
        Loop
        GrdHBusqueda.Rows = GrdHBusqueda.Rows - 1
    Else
        sstObj.Tab = 0
        cmdBuscar.SetFocus
    End If
    rsConceptos.Close
    
    pConfiguraGrid

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenaGrid"))
    Unload Me
End Sub

Private Sub pConfiguraGrid()
On Error GoTo NotificaError
    
    With GrdHBusqueda
        .FormatString = "|Clave|Descripción"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 1000 'Clave
        .ColWidth(2) = 5500 'Descripción
        
        .ScrollBars = flexScrollBarVertical
    End With
    With grdConceptoEmpresas
     .FormatString = "|||||||||||"
     .ColWidth(0) = 1000  '
     .ColWidth(1) = 1000  '
     .ColWidth(2) = 1000  '
     .ColWidth(3) = 1000  '
     .ColWidth(4) = 1000  '
     .ColWidth(5) = 1000  '
     .ColWidth(6) = 1000  '
     .ColWidth(7) = 1000  '
     .ColWidth(8) = 1000  '
     .ColWidth(9) = 1000  '
     .ColWidth(10) = 1000  '
    End With
    
    With grdExepcionesContables
     .FormatString = "|||||||"
     .ColWidth(0) = 1000  '
     .ColWidth(1) = 1000  '
     .ColWidth(2) = 1000  '
     .ColWidth(3) = 1000  '
     .ColWidth(4) = 1000  '
     .ColWidth(5) = 1000  '
     .ColWidth(6) = 1000  '
     .ColWidth(7) = 1000  '
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
    Unload Me
End Sub

Private Sub grdHBusqueda_DblClick()
On Error GoTo NotificaError
    
    If fintLocalizaPkRs(rsConceptos, 0, GrdHBusqueda.TextMatrix(GrdHBusqueda.Row, 1)) > 0 Then
        pModificaRegistro
        pHabilita 1, 1, 1, 1, 1, 0, 1
        sstObj.Tab = 0
        cmdBuscar.SetFocus
    Else
        Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
        Call pEnfocaMkTexto(txtCveConcepto)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_DblClick"))
    Unload Me
End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdHBusqueda_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":GrdHBusqueda_KeyDown"))
    Unload Me
End Sub
Private Sub mskCuentaDescExSocial1_Change()
    pAsignaCuentaImproved mskCuentaDescExSocial1, txtCuentaDescSocialEX1
End Sub

Private Sub mskCuentaDescExSocial1_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescExSocial1
End Sub

Private Sub mskCuentaDescExSocial1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescExSocial1, txtCuentaDescSocialEX1
        If mskCuentaDescExSocial1.Text <> "" Then
            mskCuentaIngresosPendientes1.SetFocus
        End If
    End If
End Sub

Private Sub mskCuentaDescExSocial1_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow1
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDescExSocial1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables.TextMatrix(intRow, 7) = lngCuenta
            Else
                grdExepcionesContables.TextMatrix(intRow, 7) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescNota_Change()
    pAsignaCuentaImproved mskCuentaDescNota, txtCuentaDescNota
End Sub

Private Sub mskCuentadescuentoCUSocial_Change()
    pAsignaCuentaImproved mskCuentadescuentoCUSocial, txtCuentaIngresosSocialCU
End Sub

Private Sub mskCuentadescuentoCUSocial_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentadescuentoCUSocial

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentadescuentoCUSocial_GotFocus"))
    Unload Me
End Sub

Private Sub mskCuentadescuentoCUSocial_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentadescuentoCUSocial, txtCuentaIngresosSocialCU
        If txtCuentaIngresosSocialCU.Text <> "" Then
            mskCuentaIngresosPendientes.SetFocus
            'mskCuentadescuentoCUSocial.SetFocus
            
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentadescuentoCUSocial_KeyPress"))
    Unload Me
End Sub

Private Sub mskCuentadescuentoCUSocial_LostFocus()
     Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentadescuentoCUSocial.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 8) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 8) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaEficienciaPaquetes_Change()
     pAsignaCuentaImproved mskCuentaEficienciaPaquetes, txtCuentaEficienciaPaquetes
End Sub

Private Sub mskCuentaEficienciaPaquetes_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaEficienciaPaquetes

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaEficienciaPaquetes_GotFocus"))
    Unload Me

End Sub

Private Sub mskCuentaEficienciaPaquetes_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaEficienciaPaquetes, txtCuentaEficienciaPaquetes
        If txtCuentaEficienciaPaquetes.Text <> "" Then
            mskCuentaDeficienciaPaquetes.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaEficienciaPaquetes_KeyPress"))
    Unload Me
End Sub

Private Sub mskCuentaEficienciaPaquetes_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaEficienciaPaquetes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 9) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 9) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDeficienciaPaquetes_Change()
     pAsignaCuentaImproved mskCuentaDeficienciaPaquetes, txtCuentaDEficienciaPaquetes
End Sub

Private Sub mskCuentaDeficienciaPaquetes_GotFocus()
 pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDeficienciaPaquetes
End Sub

Private Sub mskCuentaDeficienciaPaquetes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDeficienciaPaquetes, txtCuentaDEficienciaPaquetes
        If txtCuentaDEficienciaPaquetes.Text <> "" Then
             cboDepartamento1.SetFocus
        End If
    End If
End Sub

Private Sub mskCuentaDeficienciaPaquetes_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDeficienciaPaquetes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 10) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 10) = ""
            End If
        End If
    End If

End Sub

Private Sub mskCuentaIngresosPendientes_Change()
    pAsignaCuentaImproved mskCuentaIngresosPendientes, txtCuentaIngresosPendientes
End Sub

Private Sub mskCuentaDescuentosPendientes_Change()
    pAsignaCuentaImproved mskCuentaDescuentosPendientes, txtCuentaDescuentosPendientes
End Sub

Private Sub mskCuentaDescNota_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescNota

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaDescNota_GotFocus"))
    Unload Me
End Sub

Private Sub mskCuentaIngresosPendientes_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaIngresosPendientes

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaIngresosPendientes_GotFocus"))
    Unload Me
End Sub

Private Sub mskCuentaDescuentosPendientes_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescuentosPendientes

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaDescuentosPendientes_GotFocus"))
    Unload Me
End Sub

Private Sub mskCuentaDescNota_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescNota, txtCuentaDescNota
        If txtCuentaDescNota.Text <> "" Then
            mskCuentadescuentoCUSocial.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaDescNota_KeyPress"))
    Unload Me
End Sub

Private Sub mskCuentaIngresosPendientes_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaIngresosPendientes, txtCuentaIngresosPendientes
        If txtCuentaIngresosPendientes.Text <> "" Then
            mskCuentaDescuentosPendientes.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaIngresosPendientes_KeyPress"))
    Unload Me
End Sub

Private Sub mskCuentaDescuentosPendientes_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescuentosPendientes, txtCuentaDescuentosPendientes
        If txtCuentaDescuentosPendientes.Text <> "" Then
            mskCuentaEficienciaPaquetes.SetFocus
            'cboDepartamento1.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaDescuentosPendientes_KeyPress"))
    Unload Me
End Sub

Private Sub mskCuentaDescNota_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDescNota.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 5) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 5) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaIngresosPendientes_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaIngresosPendientes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 6) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 6) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescuentosPendientes_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDescuentosPendientes.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 7) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 7) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescNota1_Change()
    pAsignaCuentaImproved mskCuentaDescNota1, txtCuentaDescNota1
End Sub

Private Sub mskCuentaIngresosPendientes1_Change()
    pAsignaCuentaImproved mskCuentaIngresosPendientes1, txtCuentaIngresosPendientes1
End Sub

Private Sub mskCuentaDescuentosPendientes1_Change()
    pAsignaCuentaImproved mskCuentaDescuentosPendientes1, txtCuentaDescuentosPendientes1
End Sub

Private Sub mskCuentaDescNota1_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescNota1
End Sub

Private Sub mskCuentaIngresosPendientes1_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaIngresosPendientes1
End Sub

Private Sub mskCuentaDescuentosPendientes1_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescuentosPendientes1
End Sub

Private Sub mskCuentaDescNota1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescNota1, txtCuentaDescNota1
        If mskCuentaDescNota1.Text <> "" Then
            mskCuentaDescExSocial1.SetFocus
        End If
    End If
End Sub

Private Sub mskCuentaIngresosPendientes1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaIngresosPendientes1, txtCuentaIngresosPendientes1
        If mskCuentaIngresosPendientes1.Text <> "" Then
            mskCuentaDescuentosPendientes1.SetFocus
        End If
    End If
End Sub

Private Sub mskCuentaDescuentosPendientes1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescuentosPendientes1, txtCuentaDescuentosPendientes1
        
        If vlUsaSocios = True Then
            If txtCuentaDescuentosPendientes1.Text <> "" Then
                cboPaciente.SetFocus
            End If
        Else
            If txtCuentaDescuentosPendientes1.Text <> "" Then
                cboEmpresaContable.SetFocus
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescNota1_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow1
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDescNota1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables.TextMatrix(intRow, 4) = lngCuenta
            Else
                grdExepcionesContables.TextMatrix(intRow, 4) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaIngresosPendientes1_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow1
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaIngresosPendientes1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables.TextMatrix(intRow, 5) = lngCuenta
            Else
                grdExepcionesContables.TextMatrix(intRow, 5) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescuentosPendientes1_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow1
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDescuentosPendientes1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables.TextMatrix(intRow, 6) = lngCuenta
            Else
                grdExepcionesContables.TextMatrix(intRow, 6) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescuento_Change()
    pAsignaCuentaImproved mskCuentaDescuento, txtCuentaDescuento
End Sub

Private Sub mskCuentaDescuento_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaDescuento_GotFocus"))
    Unload Me
End Sub

Private Sub mskCuentaDescuento_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescuento, txtCuentaDescuento
        If txtCuentaDescuento.Text <> "" Then
            'cboDepartamento1.SetFocus
            mskCuentaDescNota.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaDescuento_KeyPress"))
    Unload Me
End Sub

Private Sub pAsignaCuentaImproved(mskCuenta As MaskEdBox, txtCuenta As TextBox)
    Dim rs As New ADODB.Recordset
  
    txtCuenta.Text = ""
    If cboEmpresaContable.ListIndex > -1 Then
        Set rs = frsRegresaRs("SELECT vchCuentaContable, vchDescripcionCuenta, intNumeroCuenta FROM CnCuenta WHERE BitEstatusMovimientos = 1 AND vchCuentaContable = '" & mskCuenta.Text & "' AND tnyClaveEmpresa = " & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) & " ORDER BY vchCuentaContable")
        If (rs.State <> adStateClosed) Then
            If rs.RecordCount > 0 Then
                txtCuenta.Text = rs!vchDescripcionCuenta
            End If
        rs.Close
        End If
    End If
End Sub

Private Sub mskCuentaDescuento_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
        lngCuenta = flngNumeroCuenta(mskCuentaDescuento.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 4) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 4) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescuento1_Change()
    pAsignaCuentaImproved mskCuentaDescuento1, txtCuentaDescuento1
End Sub

Private Sub mskCuentaDescuento1_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescuento1
End Sub

Private Sub mskCuentaDescuento1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescuento1, txtCuentaDescuento1
        
        'If vlUsaSocios = True Then
            If txtCuentaDescuento1.Text <> "" Then
                'cboPaciente.SetFocus
                mskCuentaDescNota1.SetFocus
            End If
        'Else
        '    If txtCuentaDescuento1.Text <> "" Then
        '        cboEmpresaContable.SetFocus
        '    End If
        'End If
    End If
End Sub

Private Sub mskCuentaDescuento1_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow1
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDescuento1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables.TextMatrix(intRow, 3) = lngCuenta
            Else
                grdExepcionesContables.TextMatrix(intRow, 3) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaDescuento2_Change()
    pAsignaCuentaImproved mskCuentaDescuento2, txtCuentaDescuento2
End Sub

Private Sub mskCuentaDescuento2_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaDescuento2
End Sub

Private Sub mskCuentaDescuento2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaDescuento2, txtCuentaDescuento2
        If txtCuentaDescuento2.Text <> "" Then
            cboEmpresaContable.SetFocus
        End If
    End If
End Sub

Private Sub mskCuentaDescuento2_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow2
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaDescuento2.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables2.TextMatrix(intRow, 3) = lngCuenta
            Else
                grdExepcionesContables2.TextMatrix(intRow, 3) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaIngreso_Change()
    pAsignaCuentaImproved mskCuentaIngreso, txtCuentaIngreso
End Sub

Private Sub mskCuentaIngreso_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaIngreso

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaIngreso_GotFocus"))
    Unload Me
End Sub

Private Sub mskCuentaIngreso_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaIngreso, txtCuentaIngreso
        If txtCuentaIngreso.Text <> "" Then
            mskCuentaDescuento.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaIngreso_KeyPress"))
End Sub

Private Sub pAsignaCuenta(mskObject As MaskEdBox, txtObject As TextBox)
On Error GoTo NotificaError
    Dim vllngNumeroCuenta As Long
    Dim vlstrCuentaCompleta As String
    
    If cboEmpresaContable.ListIndex = -1 Then
        Exit Sub
    End If

    If Trim(mskObject.ClipText) = "" Then
        vllngNumeroCuenta = flngBusquedaCuentasContables(False, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
        If vllngNumeroCuenta <> 0 Then
            mskObject.Text = fstrCuentaContable(vllngNumeroCuenta)
        Else
            Exit Sub
        End If
    End If
    
    vlstrCuentaCompleta = fstrCuentaCompleta(mskObject.Text)
    
    mskObject.Mask = ""
    mskObject.Text = vlstrCuentaCompleta
    mskObject.Mask = vgstrEstructuraCuentaContable
    
    vllngNumeroCuenta = flngNumeroCuenta(mskObject.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
    If vllngNumeroCuenta <> 0 Then
        txtObject.Text = fstrDescripcionCuenta(mskObject.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            'Que las cuentas contables acepten movimientos:
            If fblnMovimientos(vllngNumeroCuenta) = False Then
               'La cuenta seleccionada no acepta movimientos.
               MsgBox SIHOMsg(375), vbExclamation, "Mensaje"
                mskObject.Mask = ""
                mskObject.Text = ""
                mskObject.Mask = vgstrEstructuraCuentaContable
                txtObject.Text = ""
               Exit Sub
            End If
    Else
        mskObject.Mask = ""
        mskObject.Text = ""
        mskObject.Mask = vgstrEstructuraCuentaContable
        txtObject.Text = ""
        MsgBox SIHOMsg(222), vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsignaCuenta"))
    Unload Me
End Sub

Private Sub mskCuentaIngreso_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
   
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaIngreso.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 3) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 3) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaIngreso1_Change()
    pAsignaCuentaImproved mskCuentaIngreso1, txtCuentaIngreso1
End Sub

Private Sub mskCuentaIngreso1_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaIngreso1
End Sub

Private Sub mskCuentaIngreso1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaIngreso1, txtCuentaIngreso1
        If txtCuentaIngreso1.Text <> "" Then
            mskCuentaDescuento1.SetFocus
        End If
    End If
End Sub

Private Sub mskCuentaIngreso1_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow1
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaIngreso1.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables.TextMatrix(intRow, 2) = lngCuenta
            Else
                grdExepcionesContables.TextMatrix(intRow, 2) = ""
            End If
        End If
    End If
End Sub

Private Sub mskCuentaIngreso2_Change()
    pAsignaCuentaImproved mskCuentaIngreso2, txtCuentaIngreso2
End Sub

Private Sub mskCuentaIngreso2_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaIngreso2
End Sub

Private Sub mskCuentaIngreso2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaIngreso2, txtCuentaIngreso2
        If txtCuentaIngreso2.Text <> "" Then
            mskCuentaDescuento2.SetFocus
        End If
    End If
End Sub

Private Sub mskCuentaIngreso2_LostFocus()
    Dim intRow As Integer
    Dim lngCuenta As Long
    
    intRow = fintLocalizaRow2
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
            lngCuenta = flngNumeroCuenta(mskCuentaIngreso2.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdExepcionesContables2.TextMatrix(intRow, 2) = lngCuenta
            Else
                grdExepcionesContables2.TextMatrix(intRow, 2) = ""
            End If
        End If
    End If
End Sub

Private Sub SSTObj_Click(PreviousTab As Integer)
On Error GoTo NotificaError

    Dim vllngDesktop As Long
    
    If sstObj.Tab = 1 Then
        pLlenaGrid
        GrdHBusqueda.Enabled = True
        GrdHBusqueda.SetFocus
        Me.Height = vllngSizeNormal
        
    ElseIf sstObj.Tab = 2 Then
        If vlUsaSocios = True Then
            Me.Height = vllngSizeGrande
        Else
            Me.Height = vllngSizeMediana
        End If
        
    Else
        Me.Height = vllngSizeNormal
    End If
    
    If sstObj.Tab = 3 Then
       Me.Height = vllngSizeGrande
    End If
    
'    vllngDesktop = (Me.Top * 2) + Me.Height
'    Me.Top = Int((vllngDesktop - Me.Height) / 2)

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":SSTObj_Click"))
    Unload Me
End Sub

Private Sub txtCveConcepto_GotFocus()
On Error GoTo NotificaError

    If Not blnEnfocando Then
        pNuevoRegistro
    End If
    blnEnfocando = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveConcepto_GotFocus"))
    Unload Me
End Sub

Private Sub txtCveConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
'--------------------------------------------------------------------------------------------'
' Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o '
' modificar uno que ya existe                                                                '
'--------------------------------------------------------------------------------------------'

    Dim vlintNumero As Integer
    Dim rsBusca As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            'Buscar criterio
            If (Len(txtCveConcepto.Text) <= 0) Then
                txtCveConcepto.Text = "0"
            End If
            ReDim aListaPrecio(0)
            txtDescripcion.Enabled = True
            cboIvas.Enabled = True
            cboTipo.Enabled = True
            chkActivo.Enabled = True
            cmdCuentas.Enabled = True
            chkSaldarCuentas.Enabled = True
            If fintSigNumRs(rsConceptos, 0) = CLng(txtCveConcepto.Text) Then
                pHabilita 0, 0, 0, 0, 0, 1, 0
                txtDescripcion.SetFocus
                chkActivo.Value = 1
                chkActivo.Enabled = False
                chkSaldarCuentas.Value = 0
                Call pEnfocaTextBox(txtDescripcion)
            Else
                If fintLocalizaPkRs(rsConceptos, 0, txtCveConcepto.Text) > 0 Then
                    pHabilita 0, 0, 0, 0, 0, 1, 1
                    pModificaRegistro
                    Call pEnfocaTextBox(txtDescripcion)
                    chkActivo.Enabled = True
                Else
                    If blnClaveManualCatalogo Then
                       Set rsBusca = frsRegresaRs("SELECT * FROM PvConceptoFacturacion WHERE smiCveConcepto = " & Me.txtCveConcepto.Text, adLockReadOnly, adOpenForwardOnly)
                        If rsBusca.EOF Then
                            pHabilita 0, 0, 0, 0, 0, 1, 1
                            txtDescripcion.SetFocus
                            chkActivo.Value = 1
                            chkActivo.Enabled = False
                            Call pEnfocaTextBox(txtDescripcion)
                        Else
                            Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
                            Call pEnfocaMkTexto(txtCveConcepto)
                            txtCveConcepto_GotFocus
                        End If
                        rsBusca.Close
                    Else
                        Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
                        Call pEnfocaMkTexto(txtCveConcepto)
                        txtCveConcepto_GotFocus
                    End If
                End If
            End If
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveConcepto_KeyDown"))
    Unload Me
End Sub

Private Sub pModificaRegistro()
On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim vlintContador As Integer
    
    '-------------------------------------------------------------------'
    ' Permite realizar la modificación de la descripción de un registro '
    '-------------------------------------------------------------------'
    vgblnNuevoRegistro = False
    
    txtDescripcion.Enabled = True
    If cboIvas.ListCount > 0 And chkExentoIVA = 0 Then
        cboIvas.Enabled = True
    Else
        cboIvas.Enabled = False
    End If
        
    ReDim aListaPrecio(0)
    cboTipo.Enabled = True
    chkActivo.Enabled = True
    cmdCuentas.Enabled = True
    chkSaldarCuentas.Enabled = True
    
    With rsConceptos
        '------------------------------------'
        ' Carga los conceptos de facturación '
        '------------------------------------'
        txtCveConcepto.Text = !smicveconcepto
        txtDescripcion.Text = Trim(!chrdescripcion)
        
        cboIvas.ListIndex = IIf(!bitExentoIva = 1, -1, fintLocalizaCbo(cboIvas, !smyIVA))
        cboTipo.ListIndex = IIf(IsNull(!intTipo), 0, !intTipo)
        chkActivo.Value = IIf(!bitactivo Or !bitactivo = 1, 1, 0)
        chkSaldarCuentas.Value = IIf(!bitsaldarcuentas Or !bitsaldarcuentas = 1, 1, 0)
        chkExentoIVA.Value = IIf(!bitExentoIva = 1, 1, 0)
        chkPredeterminadoPaquetes.Value = IIf(!bitPaquetePresupuesto = 1, 1, 0)
        
        sstObj.TabEnabled(1) = True
    End With
    If cboTipo.List(cboTipo.ListIndex) = "HOSPITALARIO" Then cmdListasPrecio.Enabled = True
     
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pModificaRegistro"))
    Unload Me
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    Select Case KeyCode
        Case vbKeyReturn
            Call pEnfocaCbo(cboTipo)
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Txtdescripcion_KeyDown"))
    Unload Me
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))
    Unload Me
End Sub

Private Sub cmdAnteriorRegistro_Click()
On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsConceptos, "A")
    pModificaRegistro
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnteriorRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdPrimerRegistro_Click()
On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsConceptos, "I")
    pModificaRegistro
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimerRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdSiguienteRegistro_Click()
On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsConceptos, "S")
    pModificaRegistro
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguienteRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdUltimoRegistro_Click()
On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsConceptos, "U")
    pModificaRegistro
    pHabilita 1, 1, 1, 1, 1, 0, 1
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimoRegistro_Click"))
    Unload Me
End Sub

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intDelete As Integer)
    cmdPrimerRegistro.Enabled = intTop = 1
    cmdAnteriorRegistro.Enabled = intBack = 1
    cmdBuscar.Enabled = intlocate = 1
    cmdSiguienteRegistro.Enabled = intNext = 1
    cmdUltimoRegistro.Enabled = intEnd = 1
    cmdGrabarRegistro.Enabled = intSave = 1
    cmdDelete.Enabled = intDelete = 1
End Sub

Private Function fintValidaCuenta(vlngNumero As Long) As Integer
'=============================================================================='
' Función para validar la cuenta antes de incluirla en el detalle de la póliza '
' Regresa también en la variable <vlintOrden> si es una cuenta de orden o no   '
'=============================================================================='
On Error GoTo NotificaError
    
    Dim rsCuenta As New ADODB.Recordset
    Dim vlstrSentencia As String
   
    ' Valores de regreso (Errores):
    ' 1 = Que la cuenta no acepte movimientos
    ' 2 = Que la fecha de la cuenta sea mayor a la fecha de la póliza
    ' 0 = No hay error
    
    fintValidaCuenta = 0
    
    vlstrSentencia = "SELECT * FROM CnCuenta WHERE intNumeroCuenta = " & vlngNumero
    Set rsCuenta = frsRegresaRs(vlstrSentencia)
    If rsCuenta.RecordCount <> 0 Then
        If rsCuenta!Bitestatusmovimientos = 0 Then
            fintValidaCuenta = 1
        End If
    End If

Exit Function
NotificaError:
   Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintValidaCuenta"))
End Function

Private Function fblnMovimientos(llngNumero As Long) As Boolean
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsTotalMovimientos As New ADODB.Recordset
    
    fblnMovimientos = False
    
    vlstrSentencia = "SELECT * FROM CNCUENTA WHERE BITESTATUSMOVIMIENTOS = 1 and INTNUMEROCUENTA = " & llngNumero
    
    Set rsTotalMovimientos = frsRegresaRs(vlstrSentencia)
    If rsTotalMovimientos.RecordCount <> 0 Then
        fblnMovimientos = True
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnMovimientos"))
End Function

Private Sub pConfGridListasPrecios()

     With frmListasPreciosConcepto
         .grdPrecios.FixedCols = 1
         .grdPrecios.FixedRows = 1
                              '    00|   01|              02|                   03|             04|             05|            06|        07|    08|09|10|11|12|13|
         .grdPrecios.FormatString = "|Clave|Lista de precios|Incremento automático|Tipo incremento|Margen utilidad|Usar tabulador|Precio"
         .grdPrecios.RowHeight(0) = 450
         .grdPrecios.ColWidth(0) = 150  'Fix                    ' seleccion 00
         .grdPrecios.ColWidth(1) = 0                          ' clave de la lista de precios 01
         .grdPrecios.ColWidth(2) = 4000                         ' lista de precios 02
         .grdPrecios.ColWidth(3) = 900                          ' incremento automatico 03
         .grdPrecios.ColWidth(4) = 1000                         ' tipo incremento 04
         .grdPrecios.ColWidth(5) = 1050                          ' margen de utilidad 05
         .grdPrecios.ColWidth(6) = 800                          ' usar tabulador 06
         .grdPrecios.ColWidth(7) = 1550                         ' precio 08
         .grdPrecios.ColWidth(8) = 0                           ' lista predeterminada 12
         .grdPrecios.ColWidth(9) = 0                           ' nuevo en la lista de precios 13
                
         .grdPrecios.ColAlignmentFixed(1) = flexAlignLeftCenter
         .grdPrecios.ColAlignmentFixed(2) = flexAlignCenterCenter
         .grdPrecios.ColAlignmentFixed(3) = flexAlignCenterCenter
         .grdPrecios.ColAlignmentFixed(4) = flexAlignCenterCenter
         .grdPrecios.ColAlignmentFixed(5) = flexAlignCenterCenter
         .grdPrecios.ColAlignmentFixed(6) = flexAlignCenterCenter
         .grdPrecios.ColAlignmentFixed(7) = flexAlignCenterCenter
        
         .grdPrecios.ColAlignment(0) = flexAlignCenterCenter
         .grdPrecios.ColAlignment(2) = flexAlignLeftCenter
         .grdPrecios.ColAlignment(3) = flexAlignCenterCenter
         .grdPrecios.ColAlignment(6) = flexAlignCenterCenter
         .grdPrecios.ColAlignment(7) = flexAlignRightCenter
         
     End With
End Sub

Private Sub pInicializaFormaListasPrecios()
     With frmListasPreciosConcepto
         'cargamos los labels de clave y nombre del articulo
         .lblClaveConcepto.Caption = txtCveConcepto.Text
         .lblNombreConcepto = txtDescripcion.Text
           
         'llenamos el combobox
         .cboTipoIncremento.Clear
         .cboTipoIncremento.AddItem "ÚLTIMA COMPRA", 0
         .cboTipoIncremento.AddItem "COMPRA MÁS ALTA", 1
         .cboTipoIncremento.AddItem "PRECIO MÁXIMO AL PÚBLICO", 2
         .cboTipoIncremento.ListIndex = 0
         
         'limpiamos los campos del frame modificar listas
         '.txtCostoBase.Text = "$0.0000"
         .txtMargenUtilidad.Text = "0.0000%"
         .txtPrecio.Text = "$0.00"
         .chkIncrementoAutomatico.Value = 0
         .chkUsarTabulador.Value = 0
         .cmdAplicar.Enabled = False
     End With
End Sub

Private Function fCantidadElementos() As Integer
On Error GoTo ArregloSinElementos

fCantidadElementos = UBound(aListaPrecio())

Exit Function
ArregloSinElementos:
     If Err = 9 Then
        fCantidadElementos = 0
     Else
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & " :fcantidadElementos"))
     End If
End Function

Private Sub pCalcularPrecio(lngRow As Long, lngLista As Long)
    On Error GoTo NotificaError
    
    Dim dblPrecio As Double
    Dim dblAumentoTabulador As Double
    Dim rs As ADODB.Recordset
    Dim strParametros As String
    Dim dblCosto As Double
    Dim dblUtilidad As Double
    Dim strSQL As String
    If grdPrecios.TextMatrix(lngRow, cintColIncremetoAutomatico) = "*" Then
        dblCosto = CDbl(grdPrecios.TextMatrix(lngRow, cintColCosto))
        dblUtilidad = CDbl(Replace(grdPrecios.TextMatrix(lngRow, cintColUtilidad), "%", ""))
        dblAumentoTabulador = 0
        If grdPrecios.TextMatrix(lngRow, cintColTabulador) = "*" And grdPrecios.TextMatrix(lngRow, cintColTipo) = "AR" Then
            strSQL = "select sp_IVSelTabulador(" & dblCosto & ", '" & Trim(grdPrecios.TextMatrix(lngRow, cintColClave)) & "', " & fintTabuladorListaPrecio(lngLista) & ") aumento from dual"
            Set rs = frsRegresaRs(strSQL)
            If Not rs.EOF Then
                dblAumentoTabulador = rs!aumento
            End If
            rs.Close
        End If
        dblCosto = dblCosto * (1 + (dblUtilidad / 100))
        dblCosto = dblCosto * (1 + (dblAumentoTabulador / 100))
        'grdPrecios.TextMatrix(lngRow, cintColPrecioNuevo) = Format(dblCosto, "$###,###,###,##0.00####")
        grdPrecios.TextMatrix(lngRow, cintColPrecioNuevo) = dblCosto
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalcularPrecio"))
End Sub

Private Sub pGuardaArticuloPrecio(lngLista As Long, lngConcepto As Long, StrIAutomatico As String, StrTIncremento As String, lngMUtilidad As Double, StrUTabulador As String)
On Error GoTo NotificaError
    Dim vlstrParametros As String

    grdPrecios.Rows = 1
    grdPrecios.Cols = 20
    'grdPrecios.FormatString = "|Clave|Descripción cargo|Tipo|Tipo descripción|Clave concepto|Concepto facturación|Precio|Costo|Incremento automático|Tipo incremento|Margen utilidad|Usar tabulador|Costo última entrada|Costo más alto|Precio máximo al público|Moneda|Precio nuevo|"

    'Obtiene los artículos de la lista
    Set rsArticuloPrecio = frsRegresaRs("SELECT SMIDEPARTAMENTO FROM PVLISTAPRECIO WHERE INTCVELISTA =  " & lngLista)
    If rsArticuloPrecio.RecordCount = 0 Then
        Exit Sub
    End If
    vgstrParametrosSP = rsArticuloPrecio!SMIDEPARTAMENTO & "|" & lngLista & "|*|" & lngConcepto & "|3|0|0|0||0|1"
    rsArticuloPrecio.Close
    Set rsArticuloPrecio = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELELEMENTOSLISTASPRECIOS")

    Do While Not rsArticuloPrecio.EOF
            ' Actualización de la barra de estado
            'If pgbBarra.Value + dblAvance < 100 Then
            '    pgbBarra.Value = pgbBarra.Value + dblAvance
            'Else
            '    pgbBarra.Value = 100
            'End If
            'pgbBarra.Refresh
            
            'Artículo
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColClave) = rsArticuloPrecio!clave
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColDescripcion) = rsArticuloPrecio!Descripcion
            'Tipo
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipo) = rsArticuloPrecio!tipo
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoDes) = rsArticuloPrecio!TipoDescripcion
            'Concepto de facturación
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCveFact) = rsArticuloPrecio!CveConceptoFacturacion
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColFacturacion) = rsArticuloPrecio!ConceptoFacturacion
            'Precio
            'grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecio) = Format(rsArticuloPrecio!precio, "$###,###,###,##0.00####")
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecio) = rsArticuloPrecio!precio
            'Costo
            'grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCosto) = Format(IIf(StrTIncremento = "C", rsArticuloPrecio!costo, IIf(StrTIncremento = "M", rsArticuloPrecio!PrecioMaximoPublico, rsArticuloPrecio!CostoMasAlto)), "$###,###,###,##0.0000##")
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCosto) = Format(IIf(rsArticuloPrecio!tipo = "AR", IIf(StrTIncremento = "C", rsArticuloPrecio!costo, IIf(StrTIncremento = "M", rsArticuloPrecio!PrecioMaximoPublico, rsArticuloPrecio!CostoMasAlto)), rsArticuloPrecio!costo), "$###,###,###,##0.0000##")
            'Incremento automático
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColIncremetoAutomatico) = StrIAutomatico
            'Tipo incremento
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoIncremento) = IIf(rsArticuloPrecio!tipo = "AR", IIf(StrTIncremento = "C", cstrUltimaCompra, IIf(StrTIncremento = "M", cstrPrecioMaximoPublico, cstrCompraMasAlta)), "NA")
            'Margen de utilidad
            'grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColUtilidad) = Format(lngMUtilidad, "0.0000") & "%"
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColUtilidad) = lngMUtilidad
            'Usar tabulador
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTabulador) = StrUTabulador
            'Costo ultima entrada
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCostoUltimaEntrada) = rsArticuloPrecio!costo
            'Costo más alto
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCostoMasAlto) = rsArticuloPrecio!CostoMasAlto
            'Precio máximo al público
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioMaximopublico) = rsArticuloPrecio!PrecioMaximoPublico
            'Moneda
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColMoneda) = rsArticuloPrecio!TipoMoneda
           
            If CDbl(grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecio)) > 0 Then
                If grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColIncremetoAutomatico) = "*" Then
                    'Obtiene precio nuevo
                    pCalcularPrecio grdPrecios.Rows - 1, lngLista
                Else
                    grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioNuevo) = grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecio)
                End If
                
                'Modifica valores de la lista
                vlstrParametros = lngLista & "|" & _
                    Trim(grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColClave)) & "|" & _
                    grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipo) & "|" & _
                    grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioNuevo) & "|" & _
                    IIf(rsArticuloPrecio!tipo = "AR", IIf(grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoIncremento) = cstrUltimaCompra, "C", IIf(grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoIncremento) = cstrPrecioMaximoPublico, "M", "A")), "C") & "|" & _
                    grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColUtilidad) & "|" & _
                    IIf(rsArticuloPrecio!tipo = "AR", IIf(grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTabulador) = "*", 1, 0), 0) & "|" & _
                    IIf(grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColIncremetoAutomatico) = "*", 1, 0) & "|" & _
                    IIf(grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColMoneda) = "PESOS", 1, 0)
                          
                'GUARDA LOS CAMBIOS EN LA LISTA CON EL NUEVO PRECIO
                frsEjecuta_SP vlstrParametros, "SP_PVUPDDETALLELISTAPRECIO"
            End If
            
            grdPrecios.Rows = grdPrecios.Rows + 1
            rsArticuloPrecio.MoveNext
        Loop
        
        grdPrecios.Clear
        'sstObj.Tab = 3

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGuardaArticuloPrecio"))
End Sub

Private Sub pGuardaListaPrecios(lngGraba As Long, blnCambiaPrecio As Boolean)
On Error GoTo NotificaError
    Dim vlstrParametros As String
    Dim vlintContador As Integer
    Dim rsListaModificada As New ADODB.Recordset
    Dim vlstrQuery As String

   ' If lblnPermisoListasPrecios Then ' se tiene el permiso
        If fCantidadElementos > 0 Then
            For vlintContador = 1 To UBound(aListaPrecio) - 1
                
                Set rsListaModificada = frsRegresaRs("SELECT COUNT(*) TOTAL FROM PVPOLITICALISTAPRECIOCONCEPTO WHERE PVPOLITICALISTAPRECIOCONCEPTO.SMICVECONCEPTO = " & Me.txtCveConcepto.Text & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTCVELISTA = " & aListaPrecio(vlintContador).vllngClaveLista)
                If rsListaModificada.RecordCount > 0 Then
                    'Verifica que la lista existe, sino la da de alta
                    If rsListaModificada!Total > 0 Then
                        rsListaModificada.Close
                        
                        vlstrQuery = "SELECT COUNT(*) TOTAL FROM PVPOLITICALISTAPRECIOCONCEPTO "
                        vlstrQuery = vlstrQuery & "WHERE PVPOLITICALISTAPRECIOCONCEPTO.SMICVECONCEPTO = " & Me.txtCveConcepto.Text
                        vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTCVELISTA = " & aListaPrecio(vlintContador).vllngClaveLista
                        vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.CHRTIPOINCREMENTO = '" & CStr(aListaPrecio(vlintContador).vlstrTipoIncremento) & "' "
                        vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.NUMMARGENUTILIDAD = " & aListaPrecio(vlintContador).vldblmargenutilidad
                        vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTINCREMENTO = " & IIf(aListaPrecio(vlintContador).vlblnIncrementoAutomatico = True, 1, 0)
                        vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.INTTABULADOR = " & IIf(aListaPrecio(vlintContador).vlblnUsaTabulador = True, 1, 0)
                        vlstrQuery = vlstrQuery & " AND PVPOLITICALISTAPRECIOCONCEPTO.MNYPRECIO = " & aListaPrecio(vlintContador).vldblPrecio
                        Set rsListaModificada = frsRegresaRs(vlstrQuery)
                        'Verifica que los valores del arreglo sean iguales a los que estan en la tabla, si son diferentes modifica la lista y los artículos
                        If rsListaModificada.RecordCount > 0 Then
                            If rsListaModificada!Total = 0 Then
                                'Modifica valores de la lista
                                vlstrParametros = aListaPrecio(vlintContador).vllngClaveLista & "|" & _
                                Me.txtCveConcepto.Text & "|" & _
                                aListaPrecio(vlintContador).vldblPrecio & "|" & _
                                CStr(aListaPrecio(vlintContador).vlstrTipoIncremento) & "|" & _
                                aListaPrecio(vlintContador).vldblmargenutilidad & "|" & _
                                IIf(aListaPrecio(vlintContador).vlblnUsaTabulador = True, 1, 0) & "|" & _
                                IIf(aListaPrecio(vlintContador).vlblnIncrementoAutomatico = True, 1, 0)
                                'NOTA PARA HACER CAMBIO CUANDO SE RETOME
                                'QUE SOLO SE GUARDEN EN EL ARREGLO LOS QUE SE MODIFICARON PARA QUE SOLO ESTEN ESAS LISTAS
                                 frsEjecuta_SP vlstrParametros, " SP_PVINSUPDPOLITPRECIOCONCEP"
                                
                                If blnCambiaPrecio = True Then
                                    'Modifica valores de los artículos
                                    pGuardaArticuloPrecio aListaPrecio(vlintContador).vllngClaveLista, Me.txtCveConcepto.Text, IIf(aListaPrecio(vlintContador).vlblnIncrementoAutomatico = True, "*", ""), CStr(aListaPrecio(vlintContador).vlstrTipoIncremento), aListaPrecio(vlintContador).vldblmargenutilidad, IIf(aListaPrecio(vlintContador).vlblnUsaTabulador = True, "*", "")
                                End If
                            End If
                        End If
                    Else
                        vlstrParametros = aListaPrecio(vlintContador).vllngClaveLista & "|" & _
                        Me.txtCveConcepto.Text & "|" & _
                        aListaPrecio(vlintContador).vldblPrecio & "|" & _
                        CStr(aListaPrecio(vlintContador).vlstrTipoIncremento) & "|" & _
                        aListaPrecio(vlintContador).vldblmargenutilidad & "|" & _
                        IIf(aListaPrecio(vlintContador).vlblnUsaTabulador = True, 1, 0) & "|" & _
                        IIf(aListaPrecio(vlintContador).vlblnIncrementoAutomatico = True, 1, 0)
                        'NOTA PARA HACER CAMBIO CUANDO SE RETOME
                        'QUE SOLO SE GUARDEN EN EL ARREGLO LOS QUE SE MODIFICARON PARA QUE SOLO ESTEN ESAS LISTAS
                         frsEjecuta_SP vlstrParametros, " SP_PVINSUPDPOLITPRECIOCONCEP"
                    End If
                End If
            Next vlintContador
            
            pGuardarLogTransaccion Me.Name, EnmGrabar, lngGraba, "POLITICA DE LISTA DE PRECIOS DESDE PV", txtCveConcepto.Text
        End If
    'End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGuardaListaPrecios"))
End Sub
