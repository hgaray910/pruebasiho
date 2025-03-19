VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmSolicitudServicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisición de servicios internos"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabPrincipal 
      Height          =   8670
      Left            =   -10
      TabIndex        =   28
      Top             =   -10
      Width           =   11500
      _ExtentX        =   20294
      _ExtentY        =   15293
      _Version        =   393216
      TabOrientation  =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmSolicitudServicio.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSolicito"
      Tab(0).Control(1)=   "tmrCarga"
      Tab(0).Control(2)=   "Frame42"
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(4)=   "cmdCargarInformacion"
      Tab(0).Control(5)=   "cmdNuevo"
      Tab(0).Control(6)=   "grdSolicitud"
      Tab(0).Control(7)=   "pgbTermometro"
      Tab(0).Control(8)=   "Frame7"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmSolicitudServicio.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSTabServicio"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MyTabHeader1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmSolicitudServicio.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2230
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   11100
         Begin VB.TextBox txtEstadoActual 
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
            Height          =   375
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Estado actual de la solicitud"
            Top             =   710
            Width           =   3285
         End
         Begin VB.TextBox txtPersonaSolicita 
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
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Persona que solicita"
            Top             =   1110
            Width           =   3195
         End
         Begin VB.TextBox txtDepartamentoSolicita 
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
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Departamento que solicita"
            Top             =   700
            Width           =   3195
         End
         Begin VB.TextBox txtNumeroSolicitud 
            Alignment       =   1  'Right Justify
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
            Height          =   375
            Left            =   2400
            TabIndex        =   0
            ToolTipText     =   "Número de solicitud"
            Top             =   300
            Width           =   1170
         End
         Begin VB.CheckBox chkUrgente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Urgente"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   9840
            TabIndex        =   48
            Top             =   360
            Width           =   1155
         End
         Begin HSFlatControls.MyCombo cboTipoServicio 
            Height          =   375
            Left            =   2400
            TabIndex        =   10
            ToolTipText     =   "Selección del tipo de servicio"
            Top             =   1710
            Width           =   8570
            _ExtentX        =   15108
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   ""
            ItemData        =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin HSFlatControls.MyCombo cboDepartamentoProporciona 
            Height          =   375
            Left            =   7680
            TabIndex        =   11
            ToolTipText     =   "Selección del departamento que proporciona el servicio"
            Top             =   1110
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   ""
            ItemData        =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox mskFechaSolicitud 
            Height          =   375
            Left            =   7680
            TabIndex        =   6
            ToolTipText     =   "Fecha de la solicitud"
            Top             =   300
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mmm/yyyy hh:mm"
            Mask            =   "##/##/#### ##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Estado actual"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   30
            Top             =   765
            Width           =   1350
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Departamento proporciona"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   31
            Top             =   1170
            Width           =   1815
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo de servicio"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   135
            TabIndex        =   32
            Top             =   1750
            Width           =   1485
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Persona que solicita"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   135
            TabIndex        =   33
            Top             =   1170
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Departamento solicita"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   135
            TabIndex        =   34
            Top             =   760
            Width           =   2190
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   35
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Número de requisición"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   135
            TabIndex        =   36
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame fraSolicito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   730
         Left            =   -67080
         TabIndex        =   40
         Top             =   30
         Visible         =   0   'False
         Width           =   3255
         Begin VB.OptionButton optSolicito 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Solicito"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Filtrar las solicitudes que realizó el deparamento"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optSolicito 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Me solicitan"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   4
            ToolTipText     =   "Filtrar las solicitudes que son para el deparamento"
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.Timer tmrCarga 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -67200
         Top             =   840
      End
      Begin VB.Frame Frame42 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rango de fechas"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   -74850
         TabIndex        =   41
         Top             =   60
         Width           =   3945
         Begin MSMask.MaskEdBox mskFechaInicialConsulta 
            Height          =   375
            Left            =   600
            TabIndex        =   1
            ToolTipText     =   "Fecha inicial de la consulta"
            Top             =   250
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaFinalConsulta 
            Height          =   375
            Left            =   2400
            TabIndex        =   2
            ToolTipText     =   "Fecha final de la consulta"
            Top             =   250
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000005&
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   51
            Top             =   315
            Width           =   375
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000005&
            Caption         =   "al"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2160
            TabIndex        =   52
            Top             =   310
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   700
         Left            =   -70760
         TabIndex        =   53
         Top             =   60
         Width           =   3525
         Begin HSFlatControls.MyCombo cboEstadoBusqueda 
            Height          =   375
            Left            =   120
            TabIndex        =   54
            ToolTipText     =   "Estado actual de la requisición"
            Top             =   240
            Width           =   3280
            _ExtentX        =   5794
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   ""
            ItemData        =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MyCommandButton.MyButton cmdCargarInformacion 
         Height          =   375
         Left            =   -66600
         TabIndex        =   57
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         BackColorDown   =   -2147483643
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   15790320
         Caption         =   "Cargar información"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdNuevo 
         Height          =   600
         Left            =   -69650
         TabIndex        =   47
         Top             =   7350
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   16777215
         Picture         =   "frmSolicitudServicio.frx":0054
         BackColorDown   =   -2147483643
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   15790320
         Caption         =   ""
         CaptionPosition =   4
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmSolicitudServicio.frx":3310
         PictureAlignment=   5
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin HSFlatControls.MyTabHeader MyTabHeader1 
         Height          =   420
         Left            =   120
         TabIndex        =   58
         Top             =   2315
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   741
         Tabs            =   3
         TabCurrent      =   0
         TabWidth        =   3715
         Caption         =   $"frmSolicitudServicio.frx":3C94
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSolicitud 
         Height          =   5535
         Left            =   -74850
         TabIndex        =   5
         ToolTipText     =   "Lista de solicitudes"
         Top             =   1680
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   9763
         _Version        =   393216
         ForeColor       =   0
         Cols            =   9
         ForeColorFixed  =   0
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   -2147483638
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         Appearance      =   0
         FormatString    =   "|Fecha|Número|Departamento solicita|Persona solicita|Departamento proporciona|Estado actual|Urgente|Personal asignado"
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
         _Band(0).Cols   =   9
      End
      Begin MSComctlLib.ProgressBar pgbTermometro 
         Height          =   150
         Left            =   -74850
         TabIndex        =   42
         Top             =   1560
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin TabDlg.SSTab SSTabServicio 
         Height          =   5100
         Left            =   120
         TabIndex        =   37
         Top             =   2330
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   8996
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabMaxWidth     =   5
         BackColor       =   0
         TabCaption(0)   =   "Detalle del servicio solicitado"
         TabPicture(0)   =   "frmSolicitudServicio.frx":3CD4
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Line1"
         Tab(0).Control(1)=   "Frame3"
         Tab(0).Control(2)=   "Frame2(0)"
         Tab(0).Control(3)=   "Frame9"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Estado del servicio"
         TabPicture(1)   =   "frmSolicitudServicio.frx":3CF0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label9"
         Tab(1).Control(1)=   "Label11"
         Tab(1).Control(2)=   "Label10"
         Tab(1).Control(3)=   "mskHora"
         Tab(1).Control(4)=   "mskFechaEstado"
         Tab(1).Control(5)=   "grdEstadoSolicitud"
         Tab(1).Control(6)=   "cboEstado"
         Tab(1).Control(7)=   "cmdCambiarEstado"
         Tab(1).Control(8)=   "Frame8"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Personal asignado"
         TabPicture(2)   =   "frmSolicitudServicio.frx":3D0C
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label12"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "grdEmpleadoAsignado"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "cboEmpleado"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "cmdAsignarPersonal"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Frame4"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).ControlCount=   5
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   -64080
            TabIndex        =   61
            Top             =   360
            Width           =   255
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   10
            TabIndex        =   60
            Top             =   4950
            Width           =   5380
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Servicio solicitado"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2310
            Index           =   0
            Left            =   -75000
            TabIndex        =   39
            Top             =   490
            Width           =   11100
            Begin VB.TextBox txtServicioSolicitado 
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
               Height          =   1860
               Left            =   120
               MaxLength       =   1000
               MultiLine       =   -1  'True
               TabIndex        =   12
               ToolTipText     =   "Descripción del servicio solicitado"
               Top             =   285
               Width           =   10870
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   -74990
            TabIndex        =   59
            Top             =   4950
            Width           =   5380
         End
         Begin MyCommandButton.MyButton cmdAsignarPersonal 
            Height          =   375
            Left            =   2400
            TabIndex        =   26
            ToolTipText     =   "Asignar personal"
            Top             =   960
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   "Asignar el personal >>"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdCambiarEstado 
            Height          =   375
            Left            =   -73080
            TabIndex        =   23
            ToolTipText     =   "Cambiar estado agregar"
            Top             =   1360
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   661
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   "Asignar el estado actual >>"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin HSFlatControls.MyCombo cboEstado 
            Height          =   375
            Left            =   -73800
            TabIndex        =   21
            ToolTipText     =   "Selección del estado actual"
            Top             =   555
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   ""
            ItemData        =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin HSFlatControls.MyCombo cboEmpleado 
            Height          =   375
            Left            =   1440
            TabIndex        =   25
            ToolTipText     =   "Selección del empleado"
            Top             =   555
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   ""
            ItemData        =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdEstadoSolicitud 
            Height          =   4710
            Left            =   -69600
            TabIndex        =   24
            ToolTipText     =   "Estados de la solicitud"
            Top             =   390
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   8308
            _Version        =   393216
            ForeColor       =   0
            Cols            =   5
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            Appearance      =   0
            FormatString    =   "|Fecha|Estado"
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
            _Band(0).Cols   =   5
         End
         Begin MSMask.MaskEdBox mskFechaEstado 
            Height          =   375
            Left            =   -73800
            TabIndex        =   22
            ToolTipText     =   "Fecha del estado"
            Top             =   960
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdEmpleadoAsignado 
            Height          =   4710
            Left            =   5400
            TabIndex        =   27
            Top             =   390
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   8308
            _Version        =   393216
            ForeColor       =   0
            Cols            =   4
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            Appearance      =   0
            FormatString    =   "|Empleado"
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
            _Band(0).Cols   =   4
         End
         Begin MSMask.MaskEdBox mskHora 
            Height          =   375
            Left            =   -70750
            TabIndex        =   49
            ToolTipText     =   "Hora del estado"
            Top             =   960
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   " "
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Comentario del departamento que proporciona el servicio"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2290
            Left            =   -75000
            TabIndex        =   38
            Top             =   2810
            Width           =   11100
            Begin VB.TextBox txtComentario 
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
               Height          =   1890
               Left            =   120
               MaxLength       =   1000
               MultiLine       =   -1  'True
               TabIndex        =   13
               ToolTipText     =   "Comentario del departamento"
               Top             =   285
               Width           =   10870
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   -75000
            X2              =   -75000
            Y1              =   480
            Y2              =   360
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74775
            TabIndex        =   44
            Top             =   610
            Width           =   660
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74775
            TabIndex        =   45
            Top             =   1020
            Width           =   585
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Empleado"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   225
            TabIndex        =   46
            Top             =   610
            Width           =   1005
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Hora"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -71520
            TabIndex        =   50
            Top             =   1020
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   3510
         TabIndex        =   43
         Top             =   7380
         Width           =   4320
         Begin MyCommandButton.MyButton cmdPrimero 
            Height          =   600
            Left            =   60
            TabIndex        =   14
            ToolTipText     =   "Primer registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmSolicitudServicio.frx":3D28
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmSolicitudServicio.frx":46AA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdAnterior 
            Height          =   600
            Left            =   660
            TabIndex        =   15
            ToolTipText     =   "Anterior registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmSolicitudServicio.frx":502C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmSolicitudServicio.frx":59AE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBuscar 
            Height          =   600
            Left            =   1260
            TabIndex        =   16
            ToolTipText     =   "Búsqueda"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmSolicitudServicio.frx":6330
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmSolicitudServicio.frx":6CB4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSiguiente 
            Height          =   600
            Left            =   1860
            TabIndex        =   17
            ToolTipText     =   "Siguiente registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmSolicitudServicio.frx":7638
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmSolicitudServicio.frx":7FBA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdUltimo 
            Height          =   600
            Left            =   2460
            TabIndex        =   18
            ToolTipText     =   "Último registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmSolicitudServicio.frx":893C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmSolicitudServicio.frx":92BE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdGrabar 
            Height          =   600
            Left            =   3060
            TabIndex        =   19
            ToolTipText     =   "Guardar el registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmSolicitudServicio.frx":9C40
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmSolicitudServicio.frx":A5C4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdImprimir 
            Height          =   600
            Left            =   3660
            TabIndex        =   20
            ToolTipText     =   "Imprimir requisición"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmSolicitudServicio.frx":AF48
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmSolicitudServicio.frx":B8CC
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Empleado"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   -74880
         TabIndex        =   55
         Top             =   750
         Width           =   5295
         Begin HSFlatControls.MyCombo cboEmpleados 
            Height          =   375
            Left            =   150
            TabIndex        =   56
            Top             =   250
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   ""
            ItemData        =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmSolicitudServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : prjInventario
'| Nombre del Formulario    : frmSolicitudServicio
'-------------------------------------------------------------------------------------
'| Objetivo: Registro, consulta y seguimiento de servicios entre departamentos del
'|           del hospital.
'-------------------------------------------------------------------------------------
'| Fecha de Creación        : 09/Octubre/2003
'| Análisis y diseño        : Juan Rodolfo Ramos García - Rosenda Hernández Anaya
'| Desarrollo               : Rosenda Hernández Anaya
'-------------------------------------------------------------------------------------

Option Explicit

Public vllngNumeroOpcion As Long                        'Número de opción en el módulo que está corriendo
Dim vlstrSentencia As String                            'Sentencias SQL
Dim rsDatos As New ADODB.Recordset                      'Recordset general
Dim vlblnDepartamentoProporcionaServicios As Boolean    'Indica si el departamento logueado presta servicios
Dim vlblnConsultandoDelGrid As Boolean                  'Indica si se está consultando del grid de la consulta de solicitudes
Dim vlblnConsulta As Boolean                            'Indica si se está consultando una solicitud
Dim rsGnSolicitudServicio As New ADODB.Recordset        'Recordset dinámico de solicitudes
Dim vlvarColorNueva As Variant                          'Color en que aparecerán las nuevas solicitudes
Dim vlvarColorOtras As Variant                          'Color en que aparecerán las otras solicitudes
Private vgrptReporte As CRAXDRT.Report
Dim vlstrSQL As String
Dim rsDepartamento As ADODB.Recordset
Dim vlintNumDepartamento As Integer
Dim strSQL As String
Dim rstipodepapel As ADODB.Recordset


Private Sub cboDepartamentoProporciona_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        SSTabServicio.Tab = 0
        txtServicioSolicitado.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamentoProporciona_KeyDown"))
    Unload Me
End Sub

Private Sub cboEmpleado_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_GotFocus"))
    Unload Me
End Sub

Private Sub cboEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cmdAsignarPersonal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_KeyDown"))
    Unload Me
End Sub

Private Sub cboEmpleados_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cmdCargarInformacion.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleados_KeyDown"))
    Unload Me
End Sub

Private Sub cboEstado_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False
    mskFechaEstado.Mask = ""
    mskFechaEstado.Text = ""
    mskFechaEstado.Mask = "##/##/####"
    
    mskHora.Mask = ""
    mskHora.Text = ""
    mskHora.Mask = "##:##"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstado_GotFocus"))
    Unload Me
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        mskFechaEstado.Mask = ""
        mskFechaEstado.Text = fdtmServerFecha
        mskFechaEstado.Mask = "##/##/####"
        
        mskHora.Mask = ""
        mskHora.Text = ""
        mskHora.Mask = "##:##"
        
        mskFechaEstado.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstado_KeyDown"))
    Unload Me
End Sub

Private Sub cboEstadoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        
        If vlblnDepartamentoProporcionaServicios Then
            If optSolicito(0).Value Then
                optSolicito(0).SetFocus
            Else
                optSolicito(1).SetFocus
            End If
        Else
            cboEmpleados.SetFocus
        End If
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstadoBusqueda_KeyDown"))
    Unload Me
End Sub

Private Sub cboTipoServicio_Click()
    On Error GoTo NotificaError
    
    If cboTipoServicio.ListIndex <> -1 Then
        pCargaDepartamentoProporciona
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoServicio_Click"))
    Unload Me
End Sub

Private Sub pCargaDepartamentoProporciona()
    On Error GoTo NotificaError
    
    cboDepartamentoProporciona.Clear

    vlstrSentencia = "" & _
    "select " & _
        "GnTipoServicioDepartamento.intCveDepartamentoProporciona," & _
        "NoDepartamento.vchDescripcion " & _
    "From " & _
        "GnTipoServicioDepartamento " & _
        "inner join NoDepartamento on " & _
        "GnTipoServicioDepartamento.intCveDepartamentoProporciona = NoDepartamento.smiCveDepartamento " & _
    "Where " & _
        "GnTipoServicioDepartamento.intCveTipoServicio = " & cboTipoServicio.ItemData(cboTipoServicio.ListIndex)
    Set rsDatos = frsRegresaRs(vlstrSentencia)
    If rsDatos.RecordCount <> 0 Then
        pLlenarCboRs_new cboDepartamentoProporciona, rsDatos, 0, 1
        cboDepartamentoProporciona.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaDepartamentoProporciona"))
    Unload Me
End Sub

Private Sub cboTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cboDepartamentoProporciona.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoServicio_KeyDown"))
    Unload Me
End Sub

Private Sub chkUrgente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cboTipoServicio.SetFocus
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkUrgente_KeyDown"))
    Unload Me
End Sub

Private Sub cmdAnterior_Click()
    On Error GoTo NotificaError
    
    If vlblnConsultandoDelGrid Then
        If grdSolicitud.Row <> 1 Then
            grdSolicitud.Row = grdSolicitud.Row - 1
        End If
        If fintLocalizaPkRs(rsGnSolicitudServicio, 0, grdSolicitud.TextMatrix(grdSolicitud.Row, 2)) <> 0 Then
            pMuestra
        End If
    Else
        If rsGnSolicitudServicio.RecordCount > 0 Then
          rsGnSolicitudServicio.MovePrevious
          If rsGnSolicitudServicio.BOF Then
              rsGnSolicitudServicio.MoveNext
          End If
          pMuestra
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnterior_Click"))
    Unload Me
End Sub

Private Sub cmdAsignarPersonal_Click()
    On Error GoTo NotificaError
    Dim vllngRenglon  As Long
    
    If cboEmpleado.ListCount = 0 Then
        cboEmpleado.SetFocus
    Else
        If Trim(grdEmpleadoAsignado.TextMatrix(1, 1)) = "" Then
            vllngRenglon = 1
        Else
            grdEmpleadoAsignado.Rows = grdEmpleadoAsignado.Rows + 1
            vllngRenglon = grdEmpleadoAsignado.Rows - 1
        End If
        
        With grdEmpleadoAsignado
            .TextMatrix(vllngRenglon, 1) = cboEmpleado.List(cboEmpleado.ListIndex)
            .TextMatrix(vllngRenglon, 2) = cboEmpleado.ItemData(cboEmpleado.ListIndex)
        End With
        
        cboEmpleado.RemoveItem cboEmpleado.ListIndex
        
        If cboEmpleado.ListCount <> 0 Then
            cboEmpleado.ListIndex = 0
        End If
        
        cboEmpleado.SetFocus
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAsignarPersonal_Click"))
    Unload Me
End Sub

Private Sub cmdAsignarPersonal_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAsignarPersonal_GotFocus"))
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError
    
    SSTabPrincipal.Tab = 0
    If mskFechaInicialConsulta.Enabled And mskFechaInicialConsulta.Visible Then
      mskFechaInicialConsulta.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
    Unload Me
End Sub

Private Sub cmdCambiarEstado_Click()
    On Error GoTo NotificaError
    Dim vllngRenglon As Long

    If fblnCambiarEstado Then
                
        If Trim(grdEstadoSolicitud.TextMatrix(1, 1)) = "" Then
            vllngRenglon = 1
        Else
            grdEstadoSolicitud.Rows = grdEstadoSolicitud.Rows + 1
            vllngRenglon = grdEstadoSolicitud.Rows - 1
        End If
        
        With grdEstadoSolicitud
            .TextMatrix(vllngRenglon, 1) = mskFechaEstado.Text
            .TextMatrix(vllngRenglon, 2) = Format(mskHora, "hh:mm")
            .TextMatrix(vllngRenglon, 3) = cboEstado.List(cboEstado.ListIndex)
            .TextMatrix(vllngRenglon, 4) = cboEstado.ItemData(cboEstado.ListIndex)
        End With
        
        pCargaEstadosSiguientes cboEstado.ItemData(cboEstado.ListIndex)
        
        If cboEstado.ListCount <> 0 Then
            cboEstado.ListIndex = 0
        End If
        
        cboEstado.SetFocus

    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCambiarEstado_Click"))
    Unload Me
End Sub

Private Function fblnFechaCorrecta(vlstrFecha As String) As Boolean
    On Error GoTo NotificaError
    Dim vlstrUltimaFecha As String
    
    fblnFechaCorrecta = True
    
    If Trim(grdEstadoSolicitud.TextMatrix(1, 1)) <> "" Then
        
        vlstrUltimaFecha = Trim(grdEstadoSolicitud.TextMatrix(grdEstadoSolicitud.Rows - 1, 1))
        
        If CDate(vlstrUltimaFecha) > CDate(vlstrFecha) Then
            fblnFechaCorrecta = False
        End If
                
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnFechaCorrecta"))
    Unload Me
End Function

Private Sub cmdCambiarEstado_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCambiarEstado_GotFocus"))
    Unload Me
End Sub

Private Sub cmdCargarInformacion_Click()
On Error GoTo NotificaError

    pCargaSolicitud

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargarInformacion_Click"))
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo NotificaError
    Dim rsGnSolicitudServicioEstado As New ADODB.Recordset
    Dim rsGnSolicitudServicioPersona As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    Dim vllngContador As Long
    Dim vllngSecuencia As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then
        If fblnDatosValidos() Then
            
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            
            If vllngPersonaGraba <> 0 Then
            
                EntornoSIHO.ConeccionSIHO.BeginTrans
                
                '---------------------------------------------------------------------
                'Solicitud del servicio
                '---------------------------------------------------------------------
                With rsGnSolicitudServicio
                    If Not vlblnConsulta Then
                        .AddNew
                        !dtmFechaHoraSolicitud = fdtmServerFechaHora
                        !intCveEstadoActual = 0
                    End If
                    If Not vlblnConsulta Then
                        !intCveDepartamentoSolicita = vgintNumeroDepartamento
                        !intCveEmpleadoSolicita = vllngPersonaGraba
                    End If
                    !intCveTipoServicio = cboTipoServicio.ItemData(cboTipoServicio.ListIndex)
                    !chrDescripcionServicioSolicita = Trim(txtServicioSolicitado.Text)
                    !intCveDepartamentoProporciona = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex)
                    !chrComentarioDepartamentoPropo = IIf(Trim(txtComentario.Text) = "", " ", Trim(txtComentario.Text))
                    !bitUrgente = chkUrgente.Value
                    !intCveEstadoActual = 0 'Este dato se actualiza mas adelante
                    .Update
                    vllngSecuencia = 0
                    If Not vlblnConsulta Then
                        vllngSecuencia = flngObtieneIdentity("SEC_GnSolicitudServicio", IIf(IsNull(rsGnSolicitudServicio!intCveSolicitud), 0, rsGnSolicitudServicio!intCveSolicitud))
                    Else
                        vllngSecuencia = rsGnSolicitudServicio!intCveSolicitud
                    End If
                End With
                
                '---------------------------------------------------------------------
                'Estados de la solicitud del servicio
                '---------------------------------------------------------------------
                vlstrSentencia = "Delete From GnSolicitudServicioEstado where intCveSolicitud = " & Str(vllngSecuencia)
                pEjecutaSentencia vlstrSentencia
                
                If Trim(grdEstadoSolicitud.TextMatrix(1, 1)) <> "" Then
                
                    vlstrSentencia = "Select * From GnSolicitudServicioEstado where intCveSolicitud = -1 "
                    Set rsGnSolicitudServicioEstado = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                    
                    With rsGnSolicitudServicioEstado
                        
                        vllngContador = 1
                        Do While vllngContador <= grdEstadoSolicitud.Rows - 1
                            .AddNew
                            !intCveSolicitud = vllngSecuencia
                            !intCveEstadoServicio = grdEstadoSolicitud.TextMatrix(vllngContador, 4)
                            !dtmFechaHoraEstado = FormatDateTime(grdEstadoSolicitud.TextMatrix(vllngContador, 1), vbShortDate) & " " & FormatDateTime(grdEstadoSolicitud.TextMatrix(vllngContador, 2), vbLongTime)
                            .Update
                            vllngContador = vllngContador + 1
                            
                            If vllngContador > grdEstadoSolicitud.Rows - 1 Then
                                rsGnSolicitudServicio!intCveEstadoActual = !intCveEstadoServicio
                                rsGnSolicitudServicio.Update
                                txtEstadoActual.Text = Trim(grdEstadoSolicitud.TextMatrix(vllngContador - 1, 3))
                            End If
                        Loop
                    End With
                End If
                
                '---------------------------------------------------------------------
                'Empleados asignados
                '---------------------------------------------------------------------
                vlstrSentencia = "delete FROM GnSolicitudServicioPersona where intCveSolicitud=" & Str(vllngSecuencia)
                pEjecutaSentencia vlstrSentencia
                If Trim(grdEmpleadoAsignado.TextMatrix(1, 1)) <> "" Then
                    vlstrSentencia = "select * from GnSolicitudServicioPersona where intCveSolicitud = -1 "
                    Set rsGnSolicitudServicioPersona = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                    With rsGnSolicitudServicioPersona
                        vllngContador = 1
                        Do While vllngContador <= grdEmpleadoAsignado.Rows - 1
                            .AddNew
                            !intCveSolicitud = vllngSecuencia
                            !intcveempleado = grdEmpleadoAsignado.TextMatrix(vllngContador, 2)
                            .Update
                            vllngContador = vllngContador + 1
                        Loop
                    End With
                End If
                
                Call pGuardarLogTransaccion(Me.Name, IIf(vlblnConsulta, EnmCambiar, EnmGrabar), vllngPersonaGraba, "SOLICITUD DE SERVICIO", CStr(vllngSecuencia))
                
                EntornoSIHO.ConeccionSIHO.CommitTrans
                
                If Not vlblnConsulta Then
                  pImpresionRemota "SS", vllngSecuencia, cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex)
                  rsGnSolicitudServicio.Close
                  vlstrSentencia = "select * from GnSolicitudServicio where intCveDepartamentoSolicita = " & Str(vgintNumeroDepartamento) & " or intCveDepartamentoProporciona = " & Str(vgintNumeroDepartamento) & " order by intCveSolicitud"
                  Set rsGnSolicitudServicio = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                  rsGnSolicitudServicio.MoveFirst
                  rsGnSolicitudServicio.Find ("intCveSolicitud=" & vllngSecuencia)
                End If
                
                'La operación se realizó satisfactoriamente.
                MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                If vlblnDepartamentoProporcionaServicios Then
                    SSTabPrincipal.Tab = 0
                    If mskFechaInicialConsulta.Enabled And mskFechaInicialConsulta.Visible Then
                      mskFechaInicialConsulta.SetFocus
                    End If
                Else
                    txtNumeroSolicitud.SetFocus
                End If
        
            End If
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabar_Click"))
    Unload Me
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnDatosValidos = True
    
    If Trim(txtNumeroSolicitud.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtNumeroSolicitud.SetFocus
    End If
    If fblnDatosValidos And cboTipoServicio.ListCount = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboTipoServicio.SetFocus
    End If
    If fblnDatosValidos And cboDepartamentoProporciona.ListCount = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboDepartamentoProporciona.SetFocus
    End If
    If fblnDatosValidos And Trim(Replace(Replace(txtServicioSolicitado.Text, Chr(10), ""), Chr(13), "")) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtServicioSolicitado.SetFocus
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
    Unload Me
End Function

Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    
    vlstrx = Trim(vgstrNombreHospitalCH)                    'Nombre del hospital
    vlstrx = vlstrx & "|" & Val(txtNumeroSolicitud.Text)    'Número de solicitud
    vlstrx = vlstrx & "|" & 0                               'CveDepartamento solicita
    vlstrx = vlstrx & "|" & 0                               'CveDepartamento proporciona
    vlstrx = vlstrx & "|" & 0                               'Urgente
    vlstrx = vlstrx & "|" & 0                               'Estado actual
    vlstrx = vlstrx & "|" & "NADA"                          'Fecha inicial
    vlstrx = vlstrx & "|" & "NADA"                          'Fecha final
    Set rsReporte = frsEjecuta_SP(vlstrx, "sp_GnRptSolicitudServicio")
    
    If rsReporte.RecordCount > 0 Then
      vgrptReporte.DiscardSavedData
        'Para la impresion a media carta,carta etc
        If Not rstipodepapel.EOF Then
            vgrptReporte.SetUserPaperSize rstipodepapel!intHeight, rstipodepapel!intWidth
            vgrptReporte.PaperSize = crPaperUser
        End If
    
      pImprimeReporte vgrptReporte, rsReporte, "P", "Solicitud de servicio"
    Else
      'No existe información con esos parámetros.
      MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    If rsReporte.State <> adStateClosed Then rsReporte.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprimir_Click"))
    Unload Me
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo NotificaError
    
    SSTabPrincipal.Tab = 1
    If txtNumeroSolicitud.Enabled And txtNumeroSolicitud.Visible Then
      txtNumeroSolicitud.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNuevo_Click"))
    Unload Me
End Sub

Private Sub cmdPrimero_Click()
    On Error GoTo NotificaError
    
    If vlblnConsultandoDelGrid Then
        If fintLocalizaPkRs(rsGnSolicitudServicio, 0, grdSolicitud.TextMatrix(1, 2)) <> 0 Then
            pMuestra
        End If
    Else
        If rsGnSolicitudServicio.RecordCount > 0 Then
          rsGnSolicitudServicio.MoveFirst
          pMuestra
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimero_Click"))
    Unload Me
End Sub

Private Sub cmdSiguiente_Click()
    On Error GoTo NotificaError
    
    If vlblnConsultandoDelGrid Then
        
        If grdSolicitud.Row <> grdSolicitud.Rows - 1 Then
            grdSolicitud.Row = grdSolicitud.Row + 1
        End If
        If fintLocalizaPkRs(rsGnSolicitudServicio, 0, grdSolicitud.TextMatrix(grdSolicitud.Row, 2)) <> 0 Then
            pMuestra
        End If
        
    Else
    
        If rsGnSolicitudServicio.RecordCount > 0 Then
          rsGnSolicitudServicio.MoveNext
          If rsGnSolicitudServicio.EOF Then
              rsGnSolicitudServicio.MovePrevious
          End If
          pMuestra
        End If
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguiente_Click"))
    Unload Me
End Sub

Private Sub cmdUltimo_Click()
    On Error GoTo NotificaError

    If vlblnConsultandoDelGrid Then
        If fintLocalizaPkRs(rsGnSolicitudServicio, 0, grdSolicitud.TextMatrix(grdSolicitud.Rows - 1, 2)) <> 0 Then
            pMuestra
        End If
    Else
        If rsGnSolicitudServicio.RecordCount > 0 Then
          rsGnSolicitudServicio.MoveLast
          pMuestra
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimo_Click"))
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    vgstrNombreForm = Me.Name
    
    If vlblnDepartamentoProporcionaServicios Then
        SSTabPrincipal.Tab = 0
    Else
        SSTabPrincipal.Tab = 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim rsTemp As ADODB.Recordset
    Dim intcontador As Integer
    
    
    'Color de Tab
    SetStyle SSTabServicio.hwnd, 0
    SetSolidColor SSTabServicio.hwnd, 16777215
    SSTabSubclass SSTabServicio.hwnd
    
    'Color de Tab
    SetStyle SSTabPrincipal.hwnd, 0
    SetSolidColor SSTabPrincipal.hwnd, 16777215
    SSTabSubclass SSTabPrincipal.hwnd
    
    
    Me.Icon = frmMenuPrincipal.Icon
    
    'Esto es para la impresion a media carta
    vlstrSQL = "SELECT SMICVEDEPARTAMENTO FROM LOGIN WHERE INTNUMEROLOGIN=" & vglngNumeroLogin
    Set rsDepartamento = frsRegresaRs(vlstrSQL, adLockOptimistic, adOpenDynamic)

    If Not rsDepartamento.EOF Then
        vlintNumDepartamento = rsDepartamento!smicvedepartamento
    End If
    
    strSQL = "SELECT INTWIDTH,INTHEIGHT FROM SIPAPELDOCDEPTOMODULO INNER JOIN SITIPODEPAPEL ON SIPAPELDOCDEPTOMODULO.INTIDPAPEL=SITIPODEPAPEL.INTCVETIPODEPAPEL   WHERE VCHMODULO= '" & Trim(cgstrModulo) & "' AND VCHTIPODOCUMENTO='SS' AND INTDEPTO=" & vlintNumDepartamento & ""
    Set rstipodepapel = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)
    
    pInstanciaReporte vgrptReporte, "rptSolicitudServicio.rpt"
    
    'Recordset dinámico GnSolicitudServicio
    vlstrSentencia = "select * from GnSolicitudServicio where intCveDepartamentoSolicita = " & Str(vgintNumeroDepartamento) & " or intCveDepartamentoProporciona = " & Str(vgintNumeroDepartamento) & " order by intCveSolicitud"
    Set rsGnSolicitudServicio = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    vlvarColorNueva = &HFF0000
    vlvarColorOtras = &H80000008

    vlblnDepartamentoProporcionaServicios = False
    vlstrSentencia = "select count(*) from GnTipoServicioDepartamento where intCveDepartamentoProporciona = " & Str(vgintNumeroDepartamento)
    If frsRegresaRs(vlstrSentencia).Fields(0) <> 0 Then
        vlblnDepartamentoProporcionaServicios = True
    End If
    
    
    cboEstadoBusqueda.AddItem "<TODOS>", 0
    cboEstadoBusqueda.AddItem "NUEVA", 1
    cboEstadoBusqueda.ItemData(1) = -1
    Set rsTemp = frsEjecuta_SP("-1", "Sp_GnSelEstadoServicio")
    If rsTemp.RecordCount > 0 Then
        With cboEstadoBusqueda
            For intcontador = 1 To rsTemp.RecordCount
                .AddItem Trim(rsTemp!chrDescripcion), intcontador
                .ItemData(intcontador) = rsTemp!intCveEstadoServicio
                rsTemp.MoveNext
            Next intcontador
        End With
    End If
    
    'cboEstadoBusqueda.AddItem "NUEVA", cboEstadoBusqueda.ListCount - 1
    'cboEstadoBusqueda.ItemData(cboEstadoBusqueda.ListCount - 1) = -1
    
    Set rsTemp = frsEjecuta_SP(CStr(vgintNumeroDepartamento), "Sp_GnSelEmpleadoDepartamento")
    pLlenarCboRs_new cboEmpleados, rsTemp, 0, 1, 3
    
    pLimpiaTab0
    pLimpiaTab1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub pCargaTiposServicio()
    On Error GoTo NotificaError
    
    vlstrSentencia = "select intCveTipoServicio, chrDescripcion from GnTipoServicio where bitActivo = 1"
    Set rsDatos = frsRegresaRs(vlstrSentencia)
    If rsDatos.RecordCount <> 0 Then
        pLlenarCboRs_new cboTipoServicio, rsDatos, 0, 1
        cboTipoServicio.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTiposServicio"))
    Unload Me
End Sub

Private Sub pLimpiaTab0()
    On Error GoTo NotificaError
    
    mskFechaInicialConsulta.Mask = ""
    mskFechaInicialConsulta.Text = fdtmServerFecha
    mskFechaInicialConsulta.Mask = "##/##/####"

    mskFechaFinalConsulta.Mask = ""
    mskFechaFinalConsulta.Text = fdtmServerFecha
    mskFechaFinalConsulta.Mask = "##/##/####"
    
    cboEstadoBusqueda.ListIndex = 0
    cboEmpleados.ListIndex = 0
    
    pgbTermometro.Visible = False
    
    cmdNuevo.Visible = vlblnDepartamentoProporcionaServicios
    optSolicito(0).Value = Not vlblnDepartamentoProporcionaServicios
    optSolicito(1).Value = vlblnDepartamentoProporcionaServicios
    fraSolicito.Visible = vlblnDepartamentoProporcionaServicios
    
    If vlblnDepartamentoProporcionaServicios Then
        grdSolicitud.Height = 5100
    Else
        grdSolicitud.Height = 5800
    End If
    
    cmdCargarInformacion_Click
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaTab0"))
    Unload Me
End Sub

Private Sub pLimpiaTab1()
    On Error GoTo NotificaError
    
    'Limpiar variables
    vlblnConsultandoDelGrid = False
    vlblnConsulta = False

    'Habilitar / deshabilitar controles
    cboTipoServicio.Enabled = True
    cboDepartamentoProporciona.Enabled = True
    
    txtServicioSolicitado.Enabled = True
    txtComentario.Enabled = False
    SSTabServicio.TabEnabled(1) = False
    MyTabHeader1.TabEnabled(1) = False
    SSTabServicio.TabEnabled(2) = False
    MyTabHeader1.TabEnabled(2) = False
    
    chkUrgente.Enabled = True

    'Limpiar controles
    txtNumeroSolicitud.Text = frsRegresaRs("select isnull(max(intCveSolicitud),0)+1 from GnSolicitudServicio").Fields(0)
    
    chkUrgente.Value = 0
    
    mskFechaSolicitud.Mask = ""
    mskFechaSolicitud.Text = Format(fdtmServerFechaHora, "dd/mm/yyyy hh:mm")
    mskFechaSolicitud.Mask = "##/##/#### ##:##"
    
    txtDepartamentoSolicita.Text = vgstrNombreDepartamento
    txtPersonaSolicita.Text = ""
    
    txtEstadoActual.Text = "NUEVA"
    
    pCargaTiposServicio
    If cboTipoServicio.ListCount <> 0 Then
        cboTipoServicio.ListIndex = 0
    End If
    
    txtServicioSolicitado.Text = ""
    txtComentario.Text = ""
    
    mskFechaEstado.Mask = ""
    mskFechaEstado.Text = ""
    mskFechaEstado.Mask = "##/##/####"
    
    mskHora.Mask = ""
    mskHora.Text = ""
    mskHora.Mask = "##:##"
    
    'Estados del servicio
    pCargaEstadosSiguientes 0
    pLimpiaGridEstado
    pConfiguraGridEstado
    
    'Empleados asignados
    cboEmpleado.Clear
    pLimpiaGridEmpleado
    pConfiguraGridEmpleado
    
    SSTabServicio.Tab = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaTab1"))
    Unload Me
End Sub

Private Sub pCargaEmpleado()
    On Error GoTo NotificaError
    Dim vllngDepartamento As Long
    Dim rsTemp As ADODB.Recordset

    cboEmpleado.Clear

    vllngDepartamento = rsGnSolicitudServicio!intCveDepartamentoProporciona
        
    Set rsTemp = frsEjecuta_SP(CStr(vllngDepartamento), "Sp_GnSelEmpleadoDepartamento")
    If rsTemp.RecordCount <> 0 Then
    
        Do While Not rsTemp.EOF
            If Not fblnEstaEmpleado(rsTemp!intcveempleado) Then
                cboEmpleado.AddItem rsTemp!Nombre
                cboEmpleado.ItemData(cboEmpleado.NewIndex) = rsTemp!intcveempleado
            End If
        
            rsTemp.MoveNext
        Loop
        
        If cboEmpleado.ListCount <> 0 Then
            cboEmpleado.ListIndex = 0
        End If
    End If
    rsTemp.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaEmpleado"))
    Unload Me
End Sub

Private Function fblnEstaEmpleado(vllngCveEmpleado As Long) As Boolean
    On Error GoTo NotificaError
    Dim vllngContador As Long

    fblnEstaEmpleado = False
    
    If Trim(grdEmpleadoAsignado.TextMatrix(1, 1)) <> "" Then
        
        For vllngContador = 1 To grdEmpleadoAsignado.Rows - 1
            If Val(grdEmpleadoAsignado.TextMatrix(vllngContador, 2)) = vllngCveEmpleado Then
                fblnEstaEmpleado = True
            End If
        
        Next vllngContador
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnEstaEmpleado"))
    Unload Me
End Function

Private Sub pConfiguraGridEmpleado()
    On Error GoTo NotificaError
    
    With grdEmpleadoAsignado
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Personal asignado"
        
        .ColWidth(0) = 100
        .ColWidth(1) = 5500 'Personal asignado
        .ColWidth(2) = 0    'Cve empleado
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridEmpleado"))
    Unload Me
End Sub

Private Sub pLimpiaGridEmpleado()
    On Error GoTo NotificaError
    Dim vllngContador As Long
    
    With grdEmpleadoAsignado
        .Rows = 2
        .Cols = 3
        
        For vllngContador = 1 To .Cols - 1
            .TextMatrix(1, vllngContador) = ""
        Next vllngContador
    
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridEmpleado"))
    Unload Me
End Sub

Private Sub pConfiguraGridEstado()
    On Error GoTo NotificaError
    
    With grdEstadoSolicitud
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Fecha|Hora|Estado"
        
        .ColWidth(0) = 100
        .ColWidth(1) = 1350 'Fecha
        .ColWidth(2) = 800  'Hora
        .ColWidth(3) = 3500 'Estado
        .ColWidth(4) = 0    'CveEstado
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
    
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
    
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridEstado"))
    Unload Me
End Sub

Private Sub pLimpiaGridEstado()
    On Error GoTo NotificaError
    Dim vllngContador As Long
    
    With grdEstadoSolicitud
        .Cols = 5
        .Rows = 2
        
        For vllngContador = 1 To .Cols - 1
            .TextMatrix(1, vllngContador) = ""
        Next vllngContador
        
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridEstado"))
    Unload Me
End Sub

Private Sub pCargaSolicitud()
    On Error GoTo NotificaError
    
    Dim vllngRenglon As Long
    Dim vlblnAvanceTermometro
    Dim vlvarColor As Variant
    Dim rsGnSelSolicitudServicioDepartamento As New ADODB.Recordset
    
    grdSolicitud.Visible = False
    pLimpiaGridSolicitud
    
    vgstrParametrosSP = fstrFechaSQL(mskFechaInicialConsulta.Text, "00:00:00", True) & _
                    "|" & fstrFechaSQL(mskFechaFinalConsulta.Text, "23:59:59", True) & _
                    "|" & CStr(vgintNumeroDepartamento) & _
                    "|" & CStr(IIf(optSolicito(0).Value, 1, 0)) & _
                    "|" & cboEstadoBusqueda.ItemData(cboEstadoBusqueda.ListIndex) & _
                    "|" & cboEmpleados.ItemData(cboEmpleados.ListIndex)
    Set rsGnSelSolicitudServicioDepartamento = frsEjecuta_SP(vgstrParametrosSP, "sp_GnSelSolicitudServicioDepar")
    
    If rsGnSelSolicitudServicioDepartamento.RecordCount <> 0 Then
      
      vlblnAvanceTermometro = (rsGnSelSolicitudServicioDepartamento.RecordCount / 100) * 100
      pgbTermometro.Value = 0
      pgbTermometro.Visible = rsGnSelSolicitudServicioDepartamento.RecordCount > 50
      
      Do While Not rsGnSelSolicitudServicioDepartamento.EOF
        
        If Trim(grdSolicitud.TextMatrix(1, 1)) = "" Then
          vllngRenglon = 1
        Else
          vllngRenglon = grdSolicitud.Rows
          grdSolicitud.Rows = grdSolicitud.Rows + 1
        End If
        
        With grdSolicitud
            If Trim(rsGnSelSolicitudServicioDepartamento!EstadoActual) = "NUEVA" And optSolicito(1).Value Then
              vlvarColor = vlvarColorNueva
            Else
              vlvarColor = vlvarColorOtras
            End If
            .Row = vllngRenglon
            .Col = 1
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 1) = Format(rsGnSelSolicitudServicioDepartamento!FechaSolicitud, "dd/mmm/yyyy hh:mm")
            .Col = 2
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 2) = rsGnSelSolicitudServicioDepartamento!NumeroSolicitud
            .Col = 3
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 3) = rsGnSelSolicitudServicioDepartamento!DepartamentoSolicita
            .Col = 4
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 4) = rsGnSelSolicitudServicioDepartamento!EmpleadoSolicita
            .Col = 5
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 5) = rsGnSelSolicitudServicioDepartamento!DepartamentoProporciona
            .Col = 6
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 6) = rsGnSelSolicitudServicioDepartamento!EstadoActual
            .Col = 7
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 7) = IIf(IsNull(rsGnSelSolicitudServicioDepartamento!Urgente), "", rsGnSelSolicitudServicioDepartamento!Urgente)
            .Col = 8
            .CellForeColor = vlvarColor
            .TextMatrix(vllngRenglon, 8) = IIf(IsNull(rsGnSelSolicitudServicioDepartamento!PersonalAsignado), "", rsGnSelSolicitudServicioDepartamento!PersonalAsignado)
        End With
        
        vlblnAvanceTermometro = (pgbTermometro.Value / rsGnSelSolicitudServicioDepartamento.RecordCount) * 100
        pgbTermometro.Value = CInt(vlblnAvanceTermometro)
        rsGnSelSolicitudServicioDepartamento.MoveNext
      Loop
      
      pgbTermometro.Visible = False
    End If
    
    rsGnSelSolicitudServicioDepartamento.Close
    pConfiguraGridSolicitud
    grdSolicitud.Row = 1
    grdSolicitud.Col = 1
    grdSolicitud.Visible = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaSolicitud"))
    Unload Me
End Sub

Private Sub pConfiguraGridSolicitud()
    On Error GoTo NotificaError
    
    With grdSolicitud
        .FixedCols = 1
        .FixedRows = 1
        
        .FormatString = "|Fecha|Número|Departamento solicita|Persona solicita|Departamento proporciona|Estado|Urgente"
        
        .ColWidth(0) = 100
        .ColWidth(1) = 2000     'Fecha solicitud
        .ColWidth(2) = 850      'Número de solicitud
        .ColWidth(3) = 2700     'Departamento que solicita
        .ColWidth(4) = 3000     'Empleado que solicita
        .ColWidth(5) = 3200     'Departamento que proporciona
        .ColWidth(6) = 1500     'Estado del servicio
        .ColWidth(7) = 1000      'Urgente
        .ColWidth(8) = 3000     'Personal asignado
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignLeftCenter
        
        .ColAlignmentFixed = flexAlignCenterCenter
        
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridSolicitud"))
    Unload Me
End Sub

Private Sub pLimpiaGridSolicitud()
    On Error GoTo NotificaError
    Dim vllngContador As Long
    
    grdSolicitud.Rows = 2
    grdSolicitud.Cols = 9

    For vllngContador = 1 To grdSolicitud.Cols - 1
        grdSolicitud.TextMatrix(1, vllngContador) = ""
    Next vllngContador
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridSolicitud"))
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    If vlblnDepartamentoProporcionaServicios Then
            
        If SSTabPrincipal.Tab <> 0 Then
            Cancel = True
            
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                SSTabPrincipal.Tab = 0
                If mskFechaInicialConsulta.Enabled And mskFechaInicialConsulta.Visible Then
                  mskFechaInicialConsulta.SetFocus
                End If
            End If
        End If
    
    Else
    
        If SSTabPrincipal.Tab = 0 Then
            SSTabPrincipal.Tab = 1
            If txtNumeroSolicitud.Enabled And txtNumeroSolicitud.Visible Then
              txtNumeroSolicitud.SetFocus
            End If
            Cancel = True
        Else
            If vlblnConsulta Or cmdGrabar.Enabled Then
                Cancel = True
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    txtNumeroSolicitud.SetFocus
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub grdEmpleadoAsignado_DblClick()
    On Error GoTo NotificaError
    
    If Trim(grdEmpleadoAsignado.TextMatrix(1, 1)) <> "" Then
        If grdEmpleadoAsignado.Rows = 2 Then
            pLimpiaGridEmpleado
        Else
            grdEmpleadoAsignado.RemoveItem grdEmpleadoAsignado.Row
        End If
        
        pCargaEmpleado
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEmpleadoAsignado_DblClick"))
    Unload Me
End Sub

Private Sub grdEmpleadoAsignado_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEmpleadoAsignado_GotFocus"))
    Unload Me
End Sub

Private Sub grdEmpleadoAsignado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdEmpleadoAsignado_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEmpleadoAsignado_KeyDown"))
    Unload Me
End Sub

Private Sub grdEstadoSolicitud_DblClick()
    On Error GoTo NotificaError
    Dim vllngEstadoActual As Long

    If Trim(grdEstadoSolicitud.TextMatrix(1, 1)) <> "" Then
        If grdEstadoSolicitud.Rows = 2 Then
            pLimpiaGridEstado
            vllngEstadoActual = 0
        Else
            grdEstadoSolicitud.RemoveItem grdEstadoSolicitud.Row
            vllngEstadoActual = Val(grdEstadoSolicitud.TextMatrix(grdEstadoSolicitud.Rows - 1, 4))
        End If
        
        pCargaEstadosSiguientes vllngEstadoActual
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEstadoSolicitud_DblClick"))
    Unload Me
End Sub

Private Sub grdEstadoSolicitud_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEstadoSolicitud_GotFocus"))
    Unload Me
End Sub

Private Sub grdEstadoSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdEstadoSolicitud_DblClick
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEstadoSolicitud_KeyDown"))
    Unload Me
End Sub

Private Sub grdSolicitud_DblClick()
    On Error GoTo NotificaError
    
    If Trim(grdSolicitud.TextMatrix(1, 1)) <> "" Then
        If fintLocalizaPkRs(rsGnSolicitudServicio, 0, grdSolicitud.TextMatrix(grdSolicitud.Row, 2)) <> 0 Then
            vlblnConsultandoDelGrid = True
            pMuestra
            pHabilitaBotonera True, True, True, True, True, False, True
            SSTabPrincipal.Tab = 1
            If cmdPrimero.Enabled And cmdPrimero.Visible Then
              cmdPrimero.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdSolicitud_DblClick"))
    Unload Me
End Sub

Private Sub grdSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdSolicitud_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdSolicitud_KeyDown"))
    Unload Me
End Sub

Private Sub mskFechaEstado_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False
    pSelMkTexto mskFechaEstado
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaEstado_GotFocus"))
    Unload Me
End Sub

Private Sub mskFechaEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        mskHora.Mask = ""
        mskHora.Text = Format(fdtmServerHora, "hh:mm")
        mskHora.Mask = "##:##"
        
        mskHora.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaEstado_KeyDown"))
    Unload Me
End Sub

Private Sub mskFechaFinalConsulta_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFechaFinalConsulta

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinalConsulta_GotFocus"))
    Unload Me
End Sub

Private Sub mskFechaFinalConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cboEstadoBusqueda.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinalConsulta_KeyDown"))
    Unload Me
End Sub

Private Sub mskFechaFinalConsulta_LostFocus()
    On Error GoTo NotificaError
    
    If Not IsDate(mskFechaFinalConsulta.Text) Then
        mskFechaFinalConsulta.Mask = ""
        mskFechaFinalConsulta.Text = fdtmServerFecha
        mskFechaFinalConsulta.Mask = "##/##/####"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinalConsulta_LostFocus"))
    Unload Me
End Sub

Private Sub mskFechaInicialConsulta_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFechaInicialConsulta
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicialConsulta_GotFocus"))
    Unload Me
End Sub

Private Sub mskFechaInicialConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        mskFechaFinalConsulta.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicialConsulta_KeyDown"))
    Unload Me
End Sub

Private Sub mskFechaInicialConsulta_LostFocus()
    On Error GoTo NotificaError
    
    If Not IsDate(mskFechaInicialConsulta.Text) Then
        mskFechaInicialConsulta.Mask = ""
        mskFechaInicialConsulta.Text = fdtmServerFecha
        mskFechaInicialConsulta.Mask = "##/##/####"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicialConsulta_LostFocus"))
    Unload Me
End Sub

Private Sub mskHora_GotFocus()
On Error GoTo NotificaError

    pHabilitaBotonera False, False, False, False, False, True, False
    pSelMkTexto mskHora
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHora_GotFocus"))
    Unload Me
End Sub

Private Sub mskHora_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cmdCambiarEstado.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHora_KeyDown"))
    Unload Me
End Sub

Private Sub MyTabHeader1_Click(Index As Integer)
    MyTabHeader1.TabCurrent = Index
    SSTabServicio.Tab = Index
End Sub

Private Sub optSolicito_Click(Index As Integer)
    
    If Index = 0 Then
        cboEmpleados.ToolTipText = "Empleado que solicita"
    Else
        cboEmpleados.ToolTipText = "Empleado asignado"
    End If
    
End Sub

Private Sub optSolicito_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        cboEmpleados.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optSolicito_KeyDown"))
    Unload Me
End Sub

Private Sub SSTabPrincipal_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If SSTabPrincipal.Tab = 0 Then
        pLimpiaTab0
        If mskFechaInicialConsulta.Enabled And mskFechaInicialConsulta.Visible Then
          mskFechaInicialConsulta.SetFocus
        End If
    End If
    
    If SSTabPrincipal.Tab = 1 Then
        If Not vlblnConsulta Then
            If txtNumeroSolicitud.Enabled And txtNumeroSolicitud.Visible Then
              txtNumeroSolicitud.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":SSTabPrincipal_Click"))
    Unload Me
End Sub

Private Sub SSTabServicio_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If SSTabServicio.Tab = 1 And cboEstado.Enabled Then
        cboEstado.SetFocus
    End If
    If SSTabServicio.Tab = 2 And cboEmpleado.Enabled Then
        cboEmpleado.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":SSTabServicio_Click"))
    Unload Me
End Sub

Private Sub txtComentario_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtComentario_GotFocus"))
    Unload Me
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtComentario_KeyPress"))
    Unload Me
End Sub

Private Sub txtNumeroSolicitud_GotFocus()
    On Error GoTo NotificaError
    
    pLimpiaTab1
    pHabilitaBotonera True, True, True, True, True, False, False
    
    pSelTextBox txtNumeroSolicitud

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumeroSolicitud_GotFocus"))
    Unload Me
End Sub

Private Sub pHabilitaBotonera( _
vlblnTop As Boolean, _
vlblnBack As Boolean, _
vlblnLook As Boolean, _
vlblnnext As Boolean, _
vlblnlast As Boolean, _
vlblnsave As Boolean, _
vlblnPrint As Boolean)
    On Error GoTo NotificaError
    
    cmdPrimero.Enabled = IIf(rsGnSolicitudServicio.RecordCount = 0, False, vlblnTop)
    cmdAnterior.Enabled = IIf(rsGnSolicitudServicio.RecordCount = 0, False, vlblnBack)
    cmdBuscar.Enabled = IIf(rsGnSolicitudServicio.RecordCount = 0, False, vlblnLook)
    cmdSiguiente.Enabled = IIf(rsGnSolicitudServicio.RecordCount = 0, False, vlblnnext)
    cmdUltimo.Enabled = IIf(rsGnSolicitudServicio.RecordCount = 0, False, vlblnlast)
    cmdGrabar.Enabled = vlblnsave
    cmdImprimir.Enabled = IIf(rsGnSolicitudServicio.RecordCount = 0, False, vlblnPrint)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaBotonera"))
    Unload Me
End Sub

Private Sub txtNumeroSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        
        If fintLocalizaPkRs(rsGnSolicitudServicio, 0, Str(Val(txtNumeroSolicitud.Text))) = 0 Then
            pHabilitaBotonera False, False, False, False, False, True, False
            
            txtNumeroSolicitud.Text = frsRegresaRs("select isnull(max(intCveSolicitud),0)+1 from GnSolicitudServicio").Fields(0)
                      
            chkUrgente.SetFocus
            
        Else
            pMuestra
            pHabilitaBotonera True, True, True, True, True, False, True
            cmdPrimero.SetFocus
        End If
    
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumeroSolicitud_KeyDown"))
    Unload Me
End Sub

Private Sub pMuestra()
    On Error GoTo NotificaError
    
    vlblnConsulta = True
    
    With rsGnSolicitudServicio
        txtNumeroSolicitud.Text = !intCveSolicitud
        txtDepartamentoSolicita.Text = frsRegresaRs("select vchDescripcion from NoDepartamento where smiCveDepartamento = " & Str(!intCveDepartamentoSolicita)).Fields(0)
        txtPersonaSolicita.Text = frsRegresaRs("select rtrim(vchApellidoPaterno)||' '||rtrim(vchApellidoMaterno)||' '||rtrim(vchNombre) from NoEmpleado where intCveEmpleado = " & Str(!intCveEmpleadoSolicita)).Fields(0)
        
        mskFechaSolicitud.Mask = ""
        mskFechaSolicitud.Text = Format(!dtmFechaHoraSolicitud, "dd/mm/yyyy hh:mm")
        mskFechaSolicitud.Mask = "##/##/#### ##:##"
        
        If ((!bitUrgente = True) Or (!bitUrgente = 1)) Then
          chkUrgente.Value = 1
        Else
          chkUrgente.Value = 0
        End If
        If !intCveEstadoActual = 0 Then
            txtEstadoActual.Text = "NUEVA"
        Else
            txtEstadoActual.Text = frsRegresaRs("select rtrim(chrDescripcion) from GnEstadoServicio where intCveEstadoServicio=" & Str(!intCveEstadoActual)).Fields(0)
        End If
        
        cboTipoServicio.Clear
        cboTipoServicio.AddItem frsRegresaRs("select rtrim(chrDescripcion) from GnTipoServicio where intCveTipoServicio = " & Str(!intCveTipoServicio)).Fields(0)
        cboTipoServicio.ItemData(cboTipoServicio.NewIndex) = !intCveTipoServicio
        cboTipoServicio.ListIndex = 0
        
        cboDepartamentoProporciona.Clear
        cboDepartamentoProporciona.AddItem frsRegresaRs("select Trim(vchDescripcion) from NoDepartamento where smiCveDepartamento = " & Str(!intCveDepartamentoProporciona)).Fields(0)
        cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.NewIndex) = !intCveDepartamentoProporciona
        cboDepartamentoProporciona.ListIndex = 0
        
        txtServicioSolicitado.Text = Trim(!chrDescripcionServicioSolicita)
        txtComentario.Text = Trim(!chrComentarioDepartamentoPropo)
        
        pCargaEstadoSolicitud
        
        pCargaPersonalAsignado
        
        SSTabServicio.TabEnabled(1) = True
        MyTabHeader1.TabEnabled(1) = True
        
        chkUrgente.Enabled = Not cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento And !intCveEstadoActual = 0
        
        txtServicioSolicitado.Enabled = Not cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento And !intCveEstadoActual = 0
        txtComentario.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        cboEstado.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        mskFechaEstado.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        mskHora.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        cmdCambiarEstado.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        grdEstadoSolicitud.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        
        SSTabServicio.TabEnabled(2) = True
        MyTabHeader1.TabEnabled(2) = True
        
        cboEmpleado.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        cmdAsignarPersonal.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        grdEmpleadoAsignado.Enabled = cboDepartamentoProporciona.ItemData(cboDepartamentoProporciona.ListIndex) = vgintNumeroDepartamento
        
        cboTipoServicio.Enabled = False
        cboDepartamentoProporciona.Enabled = False
        
        pCargaEstadosSiguientes !intCveEstadoActual
        pCargaEmpleado

    End With
    
    SSTabServicio.Tab = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra"))
    Unload Me
End Sub

Private Sub pCargaEstadosSiguientes(vllngEstadoActual As Long)
    On Error GoTo NotificaError
    Dim rsDatosEstadoActual As New ADODB.Recordset
    Dim vllngOrdenEstadoActual As Long
    
    cboEstado.Clear
    
    vllngOrdenEstadoActual = 0
    
    vlstrSentencia = "select * from GnEstadoServicio where intCveEstadoServicio =" & Str(vllngEstadoActual)
    Set rsDatosEstadoActual = frsRegresaRs(vlstrSentencia)
    If rsDatosEstadoActual.RecordCount <> 0 Then
        vllngOrdenEstadoActual = rsDatosEstadoActual!intOrdenEstado
    End If

    vlstrSentencia = "select * from GnEstadoServicio where bitActivo = 1 order by intOrdenEstado "
    Set rsDatos = frsRegresaRs(vlstrSentencia)
    
    If rsDatos.RecordCount <> 0 Then
    
        Do While Not rsDatos.EOF
        
            'If rsDatos!intOrdenEstado > vllngOrdenEstadoActual Then
            
                If Not fblnEstaAsignado(rsDatos!intCveEstadoServicio) Then
                    cboEstado.AddItem Trim(rsDatos!chrDescripcion)
                    cboEstado.ItemData(cboEstado.NewIndex) = rsDatos!intCveEstadoServicio
                End If
            
            'End If
        
            rsDatos.MoveNext
            
        Loop
    
    End If

    If cboEstado.ListCount <> 0 Then
        cboEstado.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaEstadosSiguientes"))
    Unload Me
End Sub

Private Function fblnEstaAsignado(vllngCveEstado As Long) As Boolean
    On Error GoTo NotificaError
    Dim vllngContador As Long
    
    fblnEstaAsignado = False
    
    If Trim(grdEstadoSolicitud.TextMatrix(1, 1)) <> "" Then
        
        For vllngContador = 1 To grdEstadoSolicitud.Rows - 1
            If Val(grdEstadoSolicitud.TextMatrix(vllngContador, 4)) = vllngCveEstado Then
                fblnEstaAsignado = True
            End If
        Next vllngContador
        
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnEstaAsignado"))
    Unload Me
End Function

Private Sub pCargaPersonalAsignado()
    On Error GoTo NotificaError
    
    pLimpiaGridEmpleado
    
    vlstrSentencia = "" & _
    "select " & _
        "rtrim(NoEmpleado.vchApellidoPaterno)||' '||rtrim(NoEmpleado.vchApellidoMaterno)||' '||rtrim(NoEmpleado.vchNombre) NombreEmpleado," & _
        "GnSolicitudServicioPersona.intCveEmpleado " & _
    "From " & _
        "GnSolicitudServicioPersona " & _
        "inner join NoEmpleado on " & _
        "GnSolicitudServicioPersona.intCveEmpleado = NoEmpleado.intCveEmpleado " & _
    "Where " & _
        "GnSolicitudServicioPersona.intCveSolicitud = " & Str(rsGnSolicitudServicio!intCveSolicitud) & " " & _
    "Order By " & _
        "NombreEmpleado"
        
    Set rsDatos = frsRegresaRs(vlstrSentencia)
    If rsDatos.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdEmpleadoAsignado, rsDatos
    End If
    
    pConfiguraGridEmpleado

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaPersonalAsignado"))
    Unload Me
End Sub

Private Sub pCargaEstadoSolicitud()
    On Error GoTo NotificaError
    Dim intcontador As Integer
    
    pLimpiaGridEstado
    
    vlstrSentencia = "" & _
    "select " & _
        "GnSolicitudServicioEstado.dtmFechaHoraEstado FechaEstado," & _
        "GnEstadoServicio.chrDescripcion Estado," & _
        "GnSolicitudServicioEstado.intCveEstadoServicio ClaveEstado " & _
    "From " & _
        "GnSolicitudServicioEstado " & _
        "inner join GnEstadoServicio on " & _
        "GnSolicitudServicioEstado.intCveEstadoServicio = GnEstadoServicio.intCveEstadoServicio " & _
    "Where " & _
        "GnSolicitudServicioEstado.intCveSolicitud = " & Str(rsGnSolicitudServicio!intCveSolicitud) & " " & _
    "Order By " & _
        "GnSolicitudServicioEstado.dtmFechaHoraEstado "
    Set rsDatos = frsRegresaRs(vlstrSentencia)
    If rsDatos.RecordCount <> 0 Then
        
        grdEstadoSolicitud.Rows = rsDatos.RecordCount + 1
        For intcontador = 1 To rsDatos.RecordCount
        
            grdEstadoSolicitud.TextMatrix(intcontador, 1) = Format(rsDatos!FechaEstado, "dd/mmm/yyyy")
            grdEstadoSolicitud.TextMatrix(intcontador, 2) = Format(rsDatos!FechaEstado, "hh:mm")
            grdEstadoSolicitud.TextMatrix(intcontador, 3) = rsDatos!Estado
            grdEstadoSolicitud.TextMatrix(intcontador, 4) = rsDatos!ClaveEstado
            rsDatos.MoveNext

        Next intcontador
    
    End If
    pConfiguraGridEstado

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaEstadoSolicitud"))
    Unload Me
End Sub

Private Sub txtNumeroSolicitud_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumeroSolicitud_KeyPress"))
    Unload Me
End Sub

Private Sub txtServicioSolicitado_GotFocus()
    On Error GoTo NotificaError
    
    pHabilitaBotonera False, False, False, False, False, True, False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtServicioSolicitado_GotFocus"))
    Unload Me
End Sub

Private Sub txtServicioSolicitado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
   ' If KeyCode = vbKeyReturn Then
    '    cmdGrabar.SetFocus
   ' End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtServicioSolicitado_KeyDown"))
    Unload Me
End Sub

Private Sub txtServicioSolicitado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtServicioSolicitado_KeyPress"))
    Unload Me
End Sub

Private Function fblnCambiarEstado() As Boolean
On Error GoTo NotificaError

    fblnCambiarEstado = True

    If cboEstado.ListCount = 0 Then
        fblnCambiarEstado = False
        cboEstado.SetFocus
    End If
    
    If Not IsDate(mskFechaEstado.Text) And fblnCambiarEstado Then
        fblnCambiarEstado = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        mskFechaEstado.SetFocus
    End If
            
    If Not fblnFechaCorrecta(mskFechaEstado.Text) And fblnCambiarEstado Then
        fblnCambiarEstado = False
        '¡Fecha no válida!
        MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
        mskFechaEstado.SetFocus
    End If
    
    If mskHora <> "  :  " And fblnCambiarEstado Then
        If Val(Mid(mskHora, 1, 2)) < 24 And Val(Mid(mskHora, 4, 2)) < 60 And Len(Trim(mskHora)) = 5 Then
            If Not fblnValidaHora(mskHora) Then
                fblnCambiarEstado = False
                '¡Hora no válida!, formato de hora hh:mm
                MsgBox SIHOMsg(41), vbOKOnly + vbInformation, "Mensaje"
                mskHora.SetFocus
            End If
        Else
            fblnCambiarEstado = False
            '¡Hora no válida!, formato de hora hh:mm
            MsgBox SIHOMsg(41), vbOKOnly + vbInformation, "Mensaje"
            mskHora.SetFocus
        End If
    ElseIf mskHora = "  :  " And fblnCambiarEstado Then
        fblnCambiarEstado = False
        '¡Hora no válida!, formato de hora hh:mm
        MsgBox SIHOMsg(41), vbOKOnly + vbInformation, "Mensaje"
        mskHora.SetFocus
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCambiarEstado"))
    Unload Me
End Function




