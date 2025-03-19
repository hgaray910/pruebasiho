VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmRequisCargoDirecto 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisición de activo fijo"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   8835
      Left            =   -10
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   -10
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   15584
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabHeight       =   529
      BackColor       =   12632256
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmRequisCargoDirecto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmBotonera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmDetalle"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCabecera"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmRequisCargoDirecto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEstatusGrid"
      Tab(1).Control(1)=   "grdHBusqueda"
      Tab(1).Control(2)=   "cboEstatusGrid"
      Tab(1).ControlCount=   3
      Begin VB.Frame fraCabecera 
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
         Height          =   1635
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   10485
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
            Left            =   5440
            TabIndex        =   1
            ToolTipText     =   "Urgente"
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox txtEstatusMaestro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   2
            ToolTipText     =   "Estado actual de la requisición"
            Top             =   300
            Width           =   1740
         End
         Begin VB.TextBox txtEmpleadoSolicito 
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
            Left            =   2480
            Locked          =   -1  'True
            TabIndex        =   33
            ToolTipText     =   "Empleado que solicitó"
            Top             =   1110
            Width           =   3975
         End
         Begin HSFlatControls.MyCombo cboDepartamento 
            Height          =   375
            Left            =   2475
            TabIndex        =   4
            ToolTipText     =   "Departamento que realiza la requisición"
            Top             =   705
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
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   375
            Left            =   4080
            TabIndex        =   3
            ToolTipText     =   "Fecha de la requisición"
            Top             =   300
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            _Version        =   393216
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtNumero 
            Height          =   375
            Left            =   2475
            TabIndex        =   0
            ToolTipText     =   "Número de la requisición"
            Top             =   300
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label lblNumRequisCD 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Requisición"
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
            Left            =   195
            TabIndex        =   20
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblDepartamentoCD 
            BackColor       =   &H80000005&
            Caption         =   "Departamento solicitó"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   195
            TabIndex        =   21
            Top             =   760
            Width           =   2220
         End
         Begin VB.Label lblEstatusCD 
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
            Left            =   7880
            TabIndex        =   22
            Top             =   360
            Width           =   825
         End
         Begin VB.Label lblFechaCD 
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
            Left            =   3420
            TabIndex        =   23
            Top             =   360
            Width           =   585
         End
         Begin VB.Label lblEmpleadoCD 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Empleado solicitó"
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
            Left            =   195
            TabIndex        =   24
            Top             =   1170
            Width           =   1740
         End
      End
      Begin VB.Frame frmDetalle 
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
         Height          =   5625
         Left            =   120
         TabIndex        =   26
         Top             =   1620
         Width           =   10485
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHRequisicionCD 
            Height          =   3075
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Requisición de artículos de cargos directos"
            Top             =   1290
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   5424
            _Version        =   393216
            ForeColor       =   0
            Rows            =   0
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            AllowBigSelection=   0   'False
            HighLight       =   0
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            MergeCells      =   1
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
            _Band(0).Cols   =   6
         End
         Begin VB.Frame fraObservaciones 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Observaciones"
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
            Height          =   1135
            Left            =   120
            TabIndex        =   34
            Top             =   4350
            Width           =   9165
            Begin VB.TextBox txtObservaciones 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   780
               Left            =   10
               MaxLength       =   3949
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               Top             =   250
               Width           =   9045
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H80000015&
               Height          =   810
               Left            =   0
               Top             =   240
               Width           =   9070
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones"
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
               Height          =   255
               Left            =   60
               TabIndex        =   38
               Top             =   0
               Width           =   1440
            End
         End
         Begin VB.TextBox txtDescripcionLarga 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   2490
            MaxLength       =   3949
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            ToolTipText     =   "Especificación detallada"
            Top             =   670
            Width           =   7840
         End
         Begin VB.TextBox txtEstatusDetalle 
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
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Estado del artículo"
            Top             =   250
            Width           =   1140
         End
         Begin MyCommandButton.MyButton cmdBorragrid 
            Height          =   495
            Left            =   9885
            TabIndex        =   37
            ToolTipText     =   "Eliminar artículo"
            Top             =   4590
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
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
            Picture         =   "frmRequisCargoDirecto.frx":0038
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            CaptionPosition =   4
            DepthEvent      =   1
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   13003064
            ForeColorFocus  =   13003064
            ForeColorDown   =   13003064
            PictureDisabled =   "frmRequisCargoDirecto.frx":074A
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdAgregarGrid 
            Height          =   495
            Left            =   9405
            TabIndex        =   10
            ToolTipText     =   "Agregar artículo"
            Top             =   4590
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
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
            Picture         =   "frmRequisCargoDirecto.frx":0E5E
            AppearanceThemes=   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            CaptionPosition =   4
            DepthEvent      =   1
            DropDownPicture =   "frmRequisCargoDirecto.frx":1572
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   13003064
            ForeColorFocus  =   13003064
            ForeColorDown   =   13003064
            PictureDisabled =   "frmRequisCargoDirecto.frx":158E
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin HSFlatControls.MyCombo cboArticulo 
            Height          =   375
            Left            =   2475
            TabIndex        =   5
            ToolTipText     =   "Artículo a solicitar"
            Top             =   255
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
         Begin MSMask.MaskEdBox txtCantidad 
            Height          =   375
            Left            =   7620
            TabIndex        =   6
            ToolTipText     =   "Cantidad"
            Top             =   255
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000015&
            Height          =   600
            Left            =   2470
            Top             =   660
            Width           =   7900
         End
         Begin VB.Label lblArticuloCD 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Artículo"
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
            Left            =   195
            TabIndex        =   27
            Top             =   315
            Width           =   735
         End
         Begin VB.Label lblEstatusArt 
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
            Left            =   8520
            TabIndex        =   28
            Top             =   315
            Width           =   660
         End
         Begin VB.Label lblCantidadArt 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Cantidad"
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
            Left            =   6600
            TabIndex        =   29
            Top             =   315
            Width           =   945
         End
         Begin VB.Label lblDescripcionLgaArt 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción larga"
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
            Left            =   195
            TabIndex        =   30
            Top             =   720
            Width           =   1695
         End
      End
      Begin HSFlatControls.MyCombo cboEstatusGrid 
         Height          =   375
         Left            =   -72480
         TabIndex        =   31
         ToolTipText     =   "Estatus de la requisición"
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
         DragIcon        =   "frmRequisCargoDirecto.frx":1CA2
         Height          =   7420
         Left            =   -74880
         TabIndex        =   18
         ToolTipText     =   "Búsqueda de requisiciones de cargo directo"
         Top             =   570
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   13097
         _Version        =   393216
         ForeColor       =   0
         Rows            =   28
         Cols            =   10
         ForeColorFixed  =   0
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   -2147483638
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         MergeCells      =   1
         Appearance      =   0
         FormatString    =   "|intNumRequisCarDir||Departamento||Empleado|Estatus|Urgente|Fecha|Hora"
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
         _Band(0).Cols   =   10
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
      End
      Begin VB.Frame frmBotonera 
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
         Left            =   2880
         TabIndex        =   25
         Top             =   7190
         Width           =   4920
         Begin MyCommandButton.MyButton cmdImprimir 
            Height          =   600
            Left            =   4260
            TabIndex        =   16
            ToolTipText     =   "Imprimir"
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
            Picture         =   "frmRequisCargoDirecto.frx":1FAC
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":2930
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdCancelarRequisicion 
            Height          =   600
            Left            =   3665
            TabIndex        =   15
            ToolTipText     =   "Cancelar la requisición"
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
            Picture         =   "frmRequisCargoDirecto.frx":32B2
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":3C36
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdPrimerRegistro 
            Height          =   600
            Left            =   60
            TabIndex        =   11
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
            Picture         =   "frmRequisCargoDirecto.frx":45BA
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":4F3C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdAnteriorRegistro 
            Height          =   600
            Left            =   660
            TabIndex        =   12
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
            Picture         =   "frmRequisCargoDirecto.frx":58BE
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":6240
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBuscar 
            Height          =   600
            Left            =   1260
            TabIndex        =   35
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
            Picture         =   "frmRequisCargoDirecto.frx":6BC2
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":7546
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSiguienteRegistro 
            Height          =   600
            Left            =   1860
            TabIndex        =   13
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
            Picture         =   "frmRequisCargoDirecto.frx":7ECA
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":884C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdUltimoRegistro 
            Height          =   600
            Left            =   2460
            TabIndex        =   14
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
            Picture         =   "frmRequisCargoDirecto.frx":91CE
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":9B50
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdGrabarRegistro 
            Height          =   600
            Left            =   3060
            TabIndex        =   36
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
            Picture         =   "frmRequisCargoDirecto.frx":A4D2
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisCargoDirecto.frx":AE56
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.Label lblEstatusGrid 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Estado de la requisición"
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
         Left            =   -74880
         TabIndex        =   32
         Top             =   180
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmRequisCargoDirecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rosenda

'-------------------------------------------------------------------------------------
'Requisiciones de Cargos Directos para el Módulo de Inventarios'
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : CargoDirecto
'| Nombre del Formulario    : frmRequisCargoDirecto
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza la Requisición de Artículos de Cargo Directo
'|
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Ursula Orrantia - Inés Saláis
'| Autor                    : Ursula Orrantia - Inés Saláis
'| Fecha de Creación        : 24/Enero/2000
'-------------------------------------------------------------------------------------
'| Modificó                 : Oneida Almodóvar
'| Fecha última modificación: 09/Junio/2011
'-------------------------------------------------------------------------------------

Option Explicit

Public vllngNumeroOpcion As Long 'Número de opción en el módulo en que está corriendo

Dim vlblnModificaRegistro As Boolean 'Bandera que detecta si se esta modificando un registro
Dim vlblnmodificaarticulo As Boolean 'Bandera que detecta si se esta modificando un artículo en el grid
Dim vlblnNuevoRegistro As Boolean 'Bandera que detecta si se trata de un registro nuevo
Dim vlintPosRegGrid As Integer 'es el renglón donde se dió dbclick o enter en el grid
Dim vlblnCmdFiltrado As Boolean 'Detecta si se está haciendo una consulta de la tabla o del select

Dim lintSoloMaestro As Integer
Dim rsIvRequisCarDirMaestro As New ADODB.Recordset
Dim rsIvRequisCarDirDetalle As New ADODB.Recordset
Dim rsDatos As New ADODB.Recordset
Dim rsSelConsRCDMaestro As New ADODB.Recordset

Dim vlstrSentencia As String 'Sentencias SQL
Dim vgrptReporte As CRAXDRT.Report

Dim vllngNumeroOpcionCancelar As Long  'Número de opción para cancelar requisiones

Private Function fEstatusDetalle(Fila As Integer) As String
'------------------------------------------------------------------------------------------
' Determina si hay artículos pendientes o cotizados en el detalle de la requisición, para su
' posible modificación
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    fEstatusDetalle = ""
    If grdHRequisicionCD.TextMatrix(Fila, 3) = "PENDIENTE" Then
        fEstatusDetalle = "PENDIENTE"
    End If
    If grdHRequisicionCD.TextMatrix(Fila, 3) = "COTIZADA" Then
        fEstatusDetalle = "COTIZADA"
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fEstatusDetalle"))
    Unload Me
End Function

Private Sub pAgregarArtGrid()
'------------------------------------------------------------------------------------------
' Si desea agregar el artículo al grid se llama este procedimiento
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    txtDescripcionLarga = fstrFormatTxt(txtDescripcionLarga, "*", ">", 3949, True)
    Call pAgregaRegMshFGrid(grdHRequisicionCD, 6, cboArticulo.Text, txtCantidad.Text, txtEstatusDetalle.Text, txtDescripcionLarga.Text, CStr(cboArticulo.ItemData(cboArticulo.ListIndex)))
    Call pConfFGrid(grdHRequisicionCD, "|Artículo|Cantidad|Estado|Descripción larga|")
    grdHRequisicionCD.Col = 2
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregarArtGrid"))
    Unload Me
End Sub

Private Sub pAgregarArticulo()
'------------------------------------------------------------------------------------------
' Prepara la pantalla para poder seleccionar un nuevo artículo
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    txtCantidad.Text = ""
    txtCantidad.Enabled = False
    txtEstatusDetalle.Text = ""
    txtDescripcionLarga.Text = ""
    txtDescripcionLarga.Enabled = False
    cmdAgregarGrid.Enabled = False
    
    pLlenarcboArticulo True
    If cboArticulo.ListCount <> 0 Then
        cboArticulo.ListIndex = 0
    End If
    cboArticulo.Enabled = True
    cboArticulo.SetFocus
    If grdHRequisicionCD.Row > 0 Then
        cmdGrabarRegistro.Enabled = True
        cmdImprimir.Enabled = False
        grdHRequisicionCD.Enabled = True
        cmdBorragrid.Enabled = True
    Else
        cmdGrabarRegistro.Enabled = False
        cmdImprimir.Enabled = True
        grdHRequisicionCD.Enabled = False
        cmdBorragrid.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregarArticulo"))
    Unload Me
End Sub

Private Sub pAgregarRegistro(vlstrSelTxt As String)
'-------------------------------------------------------------------------------------------
' Prepara el estado de un alta de registro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    'Lo siguiente es para substituir el requery que se le quitó
    txtFecha.Text = Format((CDate(fdtmServerFechaHora)), "dd/mmm/yyyy") 'Muestra la fecha del sistema con formato
    txtNumero.Text = CStr(fintSigNumRs(rsIvRequisCarDirMaestro, 0))  'Muestra el siguiente consecutivo del campo Clave
    txtNumero.Enabled = True 'Habilita el ingreso de una clave para su búsqueda
    txtEstatusMaestro.Enabled = False
    txtEstatusMaestro.Text = "PENDIENTE"
    chkUrgente.Enabled = False
    chkUrgente.Value = 0
    txtFecha.Enabled = False
    txtEmpleadoSolicito.Text = ""
    txtEmpleadoSolicito.Enabled = False
    cboArticulo.Enabled = False
    
    pLlenarcboArticulo True
    
    If cboArticulo.ListCount <> 0 Then
        cboArticulo.ListIndex = 0
    End If
    txtCantidad.Text = ""
    txtCantidad.Enabled = False
    txtEstatusDetalle.Text = ""
    txtEstatusDetalle.Enabled = False
    txtDescripcionLarga.Text = ""
    txtDescripcionLarga.Enabled = False
    txtObservaciones.Text = ""
    txtObservaciones.Enabled = True
    
    grdHRequisicionCD.Clear
    Call pIniciaMshFGrid(grdHRequisicionCD)
    Call pLimpiaMshFGrid(grdHRequisicionCD)
    grdHRequisicionCD.Enabled = False
    cmdAgregarGrid.Enabled = False
    cmdBorragrid.Enabled = False
    grdhBusqueda.Enabled = False
    cmdCancelarRequisicion.Enabled = False
    cmdImprimir.Enabled = False
    
    pHabilitaBotonBuscar True
    If vlstrSelTxt <> "I" Then
        Call pEnfocaMkTexto(txtNumero)
    Else
        Call pSelMkTexto(txtNumero)
    End If
    vlblnNuevoRegistro = True
    vlblnModificaRegistro = False
    vlblnmodificaarticulo = False
    vlblnCmdFiltrado = False
    
    cboDepartamento.ListIndex = flngLocalizaCbo_new(cboDepartamento, CStr(vgintNumeroDepartamento)) 'se posiciona en el depto con el que se dio login
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregarRegistro"))
    Unload Me
End Sub

Private Sub pConfFGrid(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
'-------------------------------------------------------------------------------------------
' Configuraciones del grdHRequisicionCD
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintseq, vlintLargo As Integer
    
    If ObjGrid.Rows > 0 Then
        ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas
        
        ' Configura el ancho de las columnas del grdHRequisicionCD
        With ObjGrid
            .ColWidth(0) = 300 'cabecera
            .ColWidth(1) = 3900 'articulo de cargo directo
            .ColWidth(2) = 1100 'cantidad
            .ColWidth(3) = 1600 'estatus
            .ColWidth(4) = 9700 'descripcion larga
            .ColWidth(5) = 0 'clave del articulo
            
            .ScrollBars = flexScrollBarBoth
        End With
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfFGrid"))
    Unload Me
End Sub

Private Sub pConfFGridBus(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
'-------------------------------------------------------------------------------------------
' Configuraciones del grdHBusqueda
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintseq, vlintLargo As Integer
    
    If ObjGrid.Rows > 0 Then
        ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas
        
        ' Configura el ancho de las columnas del grdHBusqueda
        With ObjGrid
            .ColWidth(0) = 300 'cabecera de fila
            .ColWidth(1) = 1250 'articulo de cargo
            .ColWidth(2) = 0 'clave del departamento
            .ColWidth(3) = 2500 'departamento
            .ColWidth(4) = 0 'clave del empleado
            .ColWidth(5) = 3500 'empleado
            .ColWidth(6) = 2000 'estatus
            .ColWidth(7) = 850 'urgente
            .ColWidth(8) = 1350 'fecha
            .ColWidth(9) = 0 'hora
           
            For vlintseq = 1 To ObjGrid.Rows - 1
                If .TextMatrix(vlintseq, 7) <> "" Then
                  If .TextMatrix(vlintseq, 7) = True Then
                    .TextMatrix(vlintseq, 7) = "SI"
                  Else
                    .TextMatrix(vlintseq, 7) = "NO"
                  End If
                End If
            Next vlintseq
            
            For vlintseq = 1 To ObjGrid.Rows - 1
                .TextMatrix(vlintseq, 8) = UCase(Format(.TextMatrix(vlintseq, 8), "DD/MMM/YYYY"))
            Next vlintseq
            
            .ScrollBars = flexScrollBarBoth
        End With
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfFGridBus"))
    Unload Me
End Sub

Private Sub pConsultarGrid()
    On Error GoTo NotificaError
    
    sstObj.Tab = 1 'Se localiza en el segundo tabulador para la consulta
    pSelEstatusCD
    pFiltraGrid
    pDeshabilitaStabuno
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConsultarGrid"))
    Unload Me
End Sub

Private Sub pDeshabilitaStabuno()
'------------------------------------------------------------------------------------------
'Deshabilita los controles del primer stab
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    txtNumero.Enabled = False
    chkUrgente.Enabled = False
    cboArticulo.Enabled = False
    txtCantidad.Enabled = False
    txtDescripcionLarga.Enabled = False
    grdHRequisicionCD.Enabled = False
    cmdAgregarGrid.Enabled = False
    cmdBorragrid.Enabled = False
    cmdPrimerRegistro.Enabled = False
    cmdAnteriorRegistro.Enabled = False
    cmdBuscar.Enabled = False
    cmdSiguienteRegistro.Enabled = False
    cmdUltimoRegistro.Enabled = False
    cmdGrabarRegistro.Enabled = False
    cmdCancelarRequisicion.Enabled = False
    cmdImprimir.Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDeshabilitaStabuno"))
    Unload Me
End Sub

Private Sub pFiltraGrid()
'------------------------------------------------------------------------------------------
' Selecciona un estatus para filtrar la consulta del grid de búsqueda
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
          
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCancelar, "C") Then
        vgstrParametrosSP = cboEstatusGrid.Text & "|" & CStr(0)
    Else
        vgstrParametrosSP = cboEstatusGrid.Text & "|" & CStr(vgintNumeroDepartamento)
    End If
    
    Set rsSelConsRCDMaestro = frsEjecuta_SP(vgstrParametrosSP, "sp_IvSelConsRCDMaestro")
    Call pIniciaMshFGrid(grdhBusqueda)
    If rsSelConsRCDMaestro.RecordCount = 0 Then
        Call pLimpiaMshFGrid(grdhBusqueda)
    Else
        Call pLlenarMshFGrdRs(grdhBusqueda, rsSelConsRCDMaestro)
        Call pConfFGridBus(grdhBusqueda, "|Requisición||Departamento||Empleado|Estado|Urgente|Fecha|Hora")
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFiltraGrid"))
    Unload Me
End Sub

Private Sub pHabilitaBotonBuscar(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilita el botón de Buscar y deshabilita los demás botones
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    cmdPrimerRegistro.Enabled = Not vlblnHabilita
    cmdAnteriorRegistro.Enabled = Not vlblnHabilita
    cmdBuscar.Enabled = vlblnHabilita
    cmdSiguienteRegistro.Enabled = Not vlblnHabilita
    cmdUltimoRegistro.Enabled = Not vlblnHabilita
    cmdGrabarRegistro.Enabled = Not vlblnHabilita
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaBotonBuscar"))
    Unload Me
End Sub

Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    cmdPrimerRegistro.Enabled = vlblnHabilita
    cmdAnteriorRegistro.Enabled = vlblnHabilita
    cmdBuscar.Enabled = vlblnHabilita
    cmdSiguienteRegistro.Enabled = vlblnHabilita
    cmdUltimoRegistro.Enabled = vlblnHabilita
 
    If txtEstatusMaestro.Text = "PENDIENTE" Then
       cmdGrabarRegistro.Enabled = vlblnHabilita
        cmdCancelarRequisicion.Enabled = True
    Else
        cmdCancelarRequisicion.Enabled = False
        cmdGrabarRegistro.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaBotonModifica"))
    Unload Me
End Sub

Private Sub pLlenarcboArticulo(vlSoloActivos As Boolean)
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP(IIf(vlSoloActivos, 1, -1), "SP_IVSELARTICULOCARGODIRECTO")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs_new cboArticulo, rs, 0, 1, -1
        cboArticulo.ListIndex = 0
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboArticulo"))
    Unload Me
End Sub

Private Sub pLlenarCboDepartamento()
    On Error GoTo NotificaError
    Dim rsDepartamento As ADODB.Recordset
    
    vlstrSentencia = "SELECT * FROM NoDepartamento"
    Set rsDepartamento = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic) 'CmdTable Abre la conexión con la tabla de Departamentos utilizando un RS
    Call pLlenarCboRs_new(cboDepartamento, rsDepartamento, 0, 1, -1)
    
    rsDepartamento.Close
    
    cboDepartamento.ListIndex = flngLocalizaCbo_new(cboDepartamento, CStr(vgintNumeroDepartamento)) 'se posiciona en el depto con el que se dio login
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboDepartamento"))
    Unload Me
End Sub

Private Sub pLlenarCboEstGrid()
'----------------------------------------------------------------------------------------
' Llena el combo del estatus en el grid de búsqueda para poder filtrar las requisiciones
'----------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    cboEstatusGrid.AddItem "PENDIENTE", 0
    cboEstatusGrid.AddItem "CANCELADA", 1
    cboEstatusGrid.AddItem "COTIZADA", 2
    cboEstatusGrid.AddItem "COTIZADA PARCIAL", 3
    cboEstatusGrid.AddItem "AUTORIZADA", 4
    cboEstatusGrid.AddItem "AUTORIZADA PARCIAL", 5
    cboEstatusGrid.AddItem "NO AUTORIZADA", 6
    cboEstatusGrid.AddItem "ORDENADA", 7
    cboEstatusGrid.AddItem "ORDENADA PARCIAL", 8
    cboEstatusGrid.AddItem "RECIBIDA", 9
    cboEstatusGrid.AddItem "RECIBIDA PARCIAL", 10
    cboEstatusGrid.AddItem "SURTIDA", 11
    cboEstatusGrid.AddItem "SURTIDA PARCIAL", 12

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboEstGrid"))
    Unload Me
End Sub

Private Sub pMuestraRegistro()
'-------------------------------------------------------------------------------------------
' Permite realizar la consulta de la descripción de un registro al teclear el número de
' requisición en el txtNumero
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintCveDepto As Integer
    Dim vlintCveEmp As Integer
    Dim vlintnumrequisicion As Long
    
    
    pLlenarcboArticulo True
    cboArticulo.Enabled = False
    txtCantidad.Text = ""
    txtCantidad.Enabled = False
    txtEstatusDetalle.Text = ""
    txtDescripcionLarga.Text = ""
    txtDescripcionLarga.Enabled = False
    grdHRequisicionCD.Enabled = True
    cmdAgregarGrid.Enabled = False
    cmdBorragrid.Enabled = False
    cmdImprimir.Enabled = True
    
    
    vlintnumrequisicion = rsIvRequisCarDirMaestro!INTNUMREQUISCARDIR
    
    'actualización del registro maestro
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCancelar, "C") Then
       vlstrSentencia = "select INTNUMREQUISCARDIR, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, DTMFECHAREQUIS, DTMHORAREQUIS, BITURGENTEREQUIS, VCHESTATUSREQUIS, DTMFECHAAUTORIZA, VCHOBSERVACIONES from IvRequisCarDirMaestro order by intNumRequisCarDir"
    Else
       vlstrSentencia = "select INTNUMREQUISCARDIR, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, DTMFECHAREQUIS, DTMHORAREQUIS, BITURGENTEREQUIS, VCHESTATUSREQUIS, DTMFECHAAUTORIZA, VCHOBSERVACIONES from IvRequisCarDirMaestro where smiCveDeptoRequis = " & Str(vgintNumeroDepartamento) & " order by intNumRequisCarDir"
    End If
    
    Set rsIvRequisCarDirMaestro = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    rsIvRequisCarDirMaestro.MoveFirst
    rsIvRequisCarDirMaestro.Find ("intNumRequisCarDir=" & vlintnumrequisicion)
    '------
    
    'Datos del Maestro
    txtEstatusMaestro.Text = rsIvRequisCarDirMaestro!vchEstatusRequis
    If ((rsIvRequisCarDirMaestro!bitUrgenteRequis) Or (rsIvRequisCarDirMaestro!bitUrgenteRequis = 1)) Then
        chkUrgente.Value = 1
    Else
        chkUrgente.Value = 0
    End If
    txtNumero.Text = rsIvRequisCarDirMaestro!INTNUMREQUISCARDIR
    txtFecha.Text = Format(rsIvRequisCarDirMaestro!dtmFechaRequis, "dd/mmm/yyyy")
    cboDepartamento.ListIndex = flngLocalizaCbo_new(cboDepartamento, CStr(rsIvRequisCarDirMaestro!smiCveDeptoRequis))
    txtEmpleadoSolicito.Text = frsRegresaRs("select rtrim(NoEmpleado.vchApellidoPaterno)||' '||rtrim(NoEmpleado.vchApellidoMaterno)||' '||rtrim(NoEmpleado.vchNombre) from NoEmpleado where intCveEmpleado = " & Str(rsIvRequisCarDirMaestro!intCveEmpleaRequis)).Fields(0)
    
    txtObservaciones.Text = IIf(IsNull(Trim(rsIvRequisCarDirMaestro!vchObservaciones)), "", Trim(rsIvRequisCarDirMaestro!vchObservaciones))
        
    If rsIvRequisCarDirMaestro!vchEstatusRequis = "CANCELADA" Or rsIvRequisCarDirMaestro!vchEstatusRequis = "NO AUTORIZADA" Or rsIvRequisCarDirMaestro!vchEstatusRequis = "RECIBIDA" Then
        txtObservaciones.Enabled = False
    Else
        txtObservaciones.Enabled = True
        If rsIvRequisCarDirMaestro!vchEstatusRequis = "PENDIENTE" Then
            lintSoloMaestro = 0
        Else
            lintSoloMaestro = 1
        End If
    End If
    
    'Datos del detalle
    vlstrSentencia = "" & _
    "select " & _
        "IvArticuloCargoDirecto.vchDescripcion," & _
        "IvRequisCarDirDetalle.smiCantidad," & _
        "IvRequisCarDirDetalle.vchEstatusRequis," & _
        "IvRequisCarDirDetalle.vchDescLarga," & _
        "IvArticuloCargoDirecto.intCveArticuloCarDir " & _
    "from " & _
        "IvRequisCarDirDetalle " & _
        "inner join IvArticuloCargoDirecto on " & _
        "IvRequisCarDirDetalle.intCveArticuloCarDir = IvArticuloCargoDirecto.intCveArticuloCarDir " & _
    "where " & _
        "IvRequisCarDirDetalle.intNumRequisCarDir = " & txtNumero.Text
    Set rsDatos = frsRegresaRs(vlstrSentencia)
    If rsDatos.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdHRequisicionCD, rsDatos
    End If
    pConfFGrid grdHRequisicionCD, "|Artículo|Cantidad|Estado|Descripción larga|"
        
    grdhBusqueda.Enabled = False
    Call pHabilitaBotonModifica(True)
    cmdGrabarRegistro.Enabled = False
    vlblnNuevoRegistro = False
    
    If txtEstatusMaestro = "PENDIENTE" Then
        pModificaRegistro
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraRegistro"))
    Unload Me
End Sub

Private Sub pModificaArticulo()
'------------------------------------------------------------------------------------------
' Permite editar un artículo contenido en el grid para su modificación
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    
    vlblnmodificaarticulo = True
    
    pLlenarcboArticulo False
    cboArticulo.Text = grdHRequisicionCD.TextMatrix(grdHRequisicionCD.Row, 1)
    txtCantidad.Text = grdHRequisicionCD.TextMatrix(grdHRequisicionCD.Row, 2)
    txtEstatusDetalle.Text = grdHRequisicionCD.TextMatrix(grdHRequisicionCD.Row, 3)
    txtDescripcionLarga.Text = grdHRequisicionCD.TextMatrix(grdHRequisicionCD.Row, 4)
    
    cboArticulo.Enabled = False
    txtCantidad.Enabled = False
    cmdBorragrid.Enabled = False
    cmdAgregarGrid.Enabled = False
    
    If fEstatusDetalle(grdHRequisicionCD.Row) = "PENDIENTE" Then
        txtCantidad.Enabled = True
        cmdBorragrid.Enabled = True
        cmdAgregarGrid.Enabled = True
        txtDescripcionLarga.Enabled = True
        If cboDepartamento.ItemData(cboDepartamento.ListIndex) = vgintNumeroDepartamento Then
            txtDescripcionLarga.Locked = False
        Else
            txtDescripcionLarga.Locked = True
            txtCantidad.Enabled = False
            cmdAgregarGrid.Enabled = False
        End If
        Call pEnfocaMkTexto(txtCantidad)
    Else
        txtDescripcionLarga.Enabled = True
        txtDescripcionLarga.Locked = True
    End If
    
    If fEstatusDetalle(grdHRequisicionCD.Row) = "COTIZADA" Then
        txtCantidad.Enabled = True
        Call pEnfocaMkTexto(txtCantidad)
    End If
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pModificaArticulo"))
    Unload Me
End Sub

Private Sub pModificaRegistro()
'------------------------------------------------------------------------------------------
' Permite modificar una requisición, siempre que esté pendiente y la modifique la misma
' persona que la generó
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    cboArticulo.Enabled = True
    cboArticulo.SetFocus
    txtCantidad.Enabled = False
    txtDescripcionLarga.Enabled = False
    grdHRequisicionCD.Enabled = True
    cmdAgregarGrid.Enabled = True
    cmdBorragrid.Enabled = True
    vlblnModificaRegistro = True
    cmdGrabarRegistro.Enabled = False
    cmdAgregarGrid.Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pModificaRegistro"))
    Unload Me
End Sub

Private Sub pModificaRequis()
'------------------------------------------------------------------------------------------
' Prepara la requisición cuando hay pendientes o cotizadas en el detalle
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    grdHRequisicionCD.SetFocus
    cboArticulo.ListIndex = 0
    txtCantidad.Enabled = False
    txtCantidad.Text = ""
    txtEstatusDetalle.Text = ""
    txtDescripcionLarga.Enabled = False
    txtDescripcionLarga.Text = ""
    cmdGrabarRegistro.Enabled = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pModificaRequis"))
    Unload Me
End Sub

Private Sub pSelEstatusCD()
'------------------------------------------------------------------------------------------
' Se posiciona en la primera posición del combo para filtrar la consulta del grid de
' búsqueda al momento de cargar la forma
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    cboEstatusGrid.Enabled = True
    cboEstatusGrid.SetFocus
    cboEstatusGrid.ListIndex = 0
    grdhBusqueda.Enabled = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSelEstatusCD"))
    Unload Me
End Sub

Private Sub pValidaciones()
    On Error GoTo NotificaError

    Call pValidaMkText(txtCantidad, "N", ">", 6, True) 'que la cantidad no esté vacío
    If vgblnErrorIngreso Then
        Call pEnfocaMkTexto(txtCantidad)
    Else
        Call pValidaTextBox(txtDescripcionLarga, "*", ">", 3949, True) 'que la descripción no esté vacío
        If vgblnErrorIngreso Then
            Call pEnfocaTextBox(txtDescripcionLarga)
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pValidaciones"))
    Unload Me
End Sub

Private Sub pValidaEstatus()
'---------------------------------------------------------------------------------------------
' Valida que los estatus tecleados en el detalle correspondan al estatus del maestro
'---------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintseq As Integer
    Dim vlContCot As Integer 'Cuenta los artículos cotizados en el grid
    Dim vlContAut As Integer 'Cuenta los artículos autorizados en el grid
    Dim vlContOrd As Integer 'Cuenta los artículos ordenados en el grid
    Dim vlContRec As Integer 'Cuenta los artículos recibidos en el grid
    Dim vlContSur As Integer 'Cuenta los artículos surtidos en el grid
    
    vlContCot = 0
    vlContAut = 0
    vlContOrd = 0
    vlContRec = 0
    vlContSur = 0
    
    With grdHRequisicionCD
        For vlintseq = 1 To grdHRequisicionCD.Rows - 1
            If .TextMatrix(vlintseq, 3) = "COTIZADA" Then
                vlContCot = vlContCot + 1
            End If
    
            If .TextMatrix(vlintseq, 3) = "AUTORIZADA" Then
                vlContAut = vlContAut + 1
            End If
        
            If .TextMatrix(vlintseq, 3) = "ORDENADA" Then
                vlContOrd = vlContOrd + 1
            End If
            
            If .TextMatrix(vlintseq, 3) = "RECIBIDA" Then
                vlContRec = vlContRec + 1
            End If
            
            If .TextMatrix(vlintseq, 3) = "SURTIDA" Then
                vlContSur = vlContSur + 1
            End If
        Next vlintseq
    End With
    
    If vlContCot = grdHRequisicionCD.Rows - 1 Then
        txtEstatusMaestro.Text = "COTIZADA"
    End If
    
    If vlContAut = grdHRequisicionCD.Rows - 1 Then
        txtEstatusMaestro.Text = "AUTORIZADA"
    End If

    If vlContOrd = grdHRequisicionCD.Rows - 1 Then
        txtEstatusMaestro.Text = "ORDENADA"
    End If

    If vlContRec = grdHRequisicionCD.Rows - 1 Then
        txtEstatusMaestro.Text = "RECIBIDA"
    End If

    If vlContSur = grdHRequisicionCD.Rows - 1 Then
        txtEstatusMaestro.Text = "SURTIDA"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pValidaEstatus"))
    Unload Me
End Sub

Private Sub cboArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
    Case vbKeyReturn
       txtCantidad.Enabled = True
       txtDescripcionLarga.Enabled = True
       If cboDepartamento.ItemData(cboDepartamento.ListIndex) = vgintNumeroDepartamento Then
           txtDescripcionLarga.Locked = False
           cmdAgregarGrid.Enabled = True
       Else
           txtDescripcionLarga.Locked = True
           cmdAgregarGrid.Enabled = False
           MsgBox SIHOMsg(1261), vbOKOnly + vbExclamation, "Mensaje"
       End If

       cmdGrabarRegistro.Enabled = False
       Call pEnfocaMkTexto(txtCantidad)
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboArticulo_KeyDown"))
    Unload Me
End Sub

Private Sub cboEstatusGrid_Click()
    On Error GoTo NotificaError
    
    If cboEstatusGrid.ListIndex <> -1 Then
        pFiltraGrid
        If grdhBusqueda.Rows = 0 Then 'No existen requisiciones con ese estatus
            grdhBusqueda.Enabled = False
            MsgBox SIHOMsg("5"), vbExclamation, "Mensaje"
            cboEstatusGrid.SetFocus
        Else
            grdhBusqueda.Enabled = True
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstatusGrid_Click"))
    Unload Me
End Sub

Private Sub cboEstatusGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            pFiltraGrid
            If grdhBusqueda.Rows = 0 Then 'No existen requisiciones con ese estatus
                grdhBusqueda.Enabled = False
                MsgBox SIHOMsg("5"), vbExclamation, "Mensaje"
                cboEstatusGrid.SetFocus
            Else
                grdhBusqueda.Enabled = True
            End If
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstatusGrid_KeyDown"))
    Unload Me
End Sub

Private Sub chkUrgente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Select Case KeyCode
    Case vbKeyReturn
        pAgregarArticulo
        cboArticulo.SetFocus
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkUrgente_KeyDown"))
    Unload Me
End Sub

Private Sub cmdAgregarGrid_Click()
    On Error GoTo NotificaError

    Dim vlintNumReg As Integer
    Dim vlintPosReg As Integer
    Dim vlintPosRegDesc As Integer
    Dim vstrReg As String
    Dim vlintResul As Integer
    Dim vlstrMensaje As String
    
    If cboDepartamento.ItemData(cboDepartamento.ListIndex) <> vgintNumeroDepartamento Then
        MsgBox SIHOMsg(1261), vbOKOnly + vbExclamation, "Mensaje"
        Exit Sub
    End If
    
    If cboArticulo.ListCount <> 0 Then
    
        vlintPosReg = fintLocRegMshFGrd(grdHRequisicionCD, txtDescripcionLarga.Text, 4) 'busca la descripción tecleada
        If vlintPosReg > 0 Or vlblnmodificaarticulo Then 'si lo encontró en el grid o se dio dbclick en el grid
            'vuelve a buscar la descripción en el grid
            vlintPosRegDesc = fintLocRegMshFGrd(grdHRequisicionCD, txtDescripcionLarga.Text, 4)
            If vlintPosRegDesc > 0 And vlblnmodificaarticulo And vlintPosRegDesc <> vlintPosRegGrid Then 'advertir que existe un registro igual
            'si lo encuentra en el grid y hubo un doble clik en el grid y se corrige con una descripcion repetida en el grid
                vlstrMensaje = SIHOMsg("19") & Chr(13) & "¿Desea sustituir?" 'Existe información con el mismo contenido
                vlintResul = MsgBox(vlstrMensaje, (vbYesNo + vbQuestion), "Mensaje")
                If vlintResul = vbYes Then 'Borrar el registro editado y actualizar el otro
                    grdHRequisicionCD.Row = vlintPosRegGrid 'es el renglón donde se dió dbclick o enter
                    Call pActualizaRegMshFGrid(grdHRequisicionCD, vlintPosRegDesc, cboArticulo.Text, txtCantidad.Text, txtEstatusDetalle.Text, fstrFormatTxt(txtDescripcionLarga, "*", ">", 3949, True), CStr(cboArticulo.ItemData(cboArticulo.ListIndex)))
                    Call pBorrarRegMshFGrd(grdHRequisicionCD, grdHRequisicionCD.Row)
                End If
                vlblnmodificaarticulo = False
                
                If txtEstatusMaestro = "PENDIENTE" Then 'ahorita
                    pAgregarArticulo
                Else
                    pModificaRequis
                End If
                
            Else
                pValidaciones
                If vgblnErrorIngreso = False Then 'Desea actualizar los datos?
                    vlintResul = MsgBox(SIHOMsg("7"), (vbYesNo + vbExclamation), "Mensaje")
                    If vlintResul = vbYes Then
                        txtDescripcionLarga = fstrFormatTxt(txtDescripcionLarga, "*", ">", 3949, True)
                        If vlblnmodificaarticulo = False Then
                            Call pActualizaRegMshFGrid(grdHRequisicionCD, vlintPosReg, cboArticulo.Text, txtCantidad.Text, txtEstatusDetalle.Text, txtDescripcionLarga.Text, CStr(cboArticulo.ItemData(cboArticulo.ListIndex)))
                        Else
                            Call pActualizaRegMshFGrid(grdHRequisicionCD, vlintPosRegGrid, cboArticulo.Text, txtCantidad.Text, txtEstatusDetalle.Text, txtDescripcionLarga.Text, CStr(cboArticulo.ItemData(cboArticulo.ListIndex)))
                        End If
                        grdHRequisicionCD.Col = 1
                    End If
                    vlblnmodificaarticulo = False

                    If txtEstatusMaestro = "PENDIENTE" Then
                        pAgregarArticulo
                    Else
                        pModificaRequis
                    End If
                End If
            End If
        Else
            pValidaciones
            If vgblnErrorIngreso = False Then 'Desea agregar el nuevo artículo?
                vlintResul = MsgBox(SIHOMsg("8"), (vbYesNo + vbQuestion), "Mensaje")
                If vlintResul = vbYes Then
                    pAgregarArtGrid
                    Call pAgregarArticulo
                End If
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregarGrid_Click"))
    Unload Me
End Sub

Private Sub cmdAnteriorRegistro_Click()
'-------------------------------------------------------------------------------------------
' Manda llamar los procedimientos pPosicionaRegRs y ModificaRegistro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If vlblnCmdFiltrado Then
        If grdhBusqueda.Row <> 1 Then
            grdhBusqueda.Row = grdhBusqueda.Row - 1
        End If
        If fintLocalizaPkRs(rsIvRequisCarDirMaestro, 0, grdhBusqueda.TextMatrix(grdhBusqueda.Row, 1)) <> 0 Then
            pMuestraRegistro
        End If
    Else
        rsIvRequisCarDirMaestro.MovePrevious
        If rsIvRequisCarDirMaestro.BOF Then
            rsIvRequisCarDirMaestro.MoveNext
        End If
        pMuestraRegistro
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnteriorRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdBorraGrid_Click()
'------------------------------------------------------------------------------------------
' Borra del grid el articulo que esté seleccionado
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlstrMensaje As String
    Dim vlintResultado As Integer
    Dim vlintNumReg As Integer

    'If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCancelar, "C") Then
    If cboDepartamento.ItemData(cboDepartamento.ListIndex) <> vgintNumeroDepartamento Then
        MsgBox SIHOMsg(1261), vbOKOnly + vbExclamation, "Mensaje"
    Else
        If (grdHRequisicionCD.Rows - 1) > 0 Then
            vlstrMensaje = SIHOMsg("6") & Chr(13)
            vlintResultado = MsgBox(vlstrMensaje, (vbYesNo + vbQuestion), "Mensaje") '"¿Está seguro de eliminar los datos?"
            If vlintResultado = vbYes Then
                Call pBorrarRegMshFGrd(grdHRequisicionCD, grdHRequisicionCD.Row)
                grdHRequisicionCD.Refresh
            End If
            
            If txtEstatusMaestro = "PENDIENTE" Then
                pAgregarArticulo
            Else
                pModificaRequis
            End If
        End If
        If vlblnmodificaarticulo = True Then
            vlblnmodificaarticulo = False
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBorraGrid_Click"))
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
'-------------------------------------------------------------------------------------------
' Manda el enfoque al Tab 1 del sstObj para visualizar la búsqueda y actualizar el Grid
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If rsIvRequisCarDirMaestro.RecordCount <> 0 Then
        pConsultarGrid
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
    Unload Me
End Sub

Private Sub cmdCancelarRequisicion_Click()
'-------------------------------------------------------------------------------------------
'  Permite cancelar una requisición, siempre que sea por la misma persona que la generó y
'  su estatus sea de PENDIENTE
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vllngPersonaGraba As Long
       
    'If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C") Then
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            With rsIvRequisCarDirMaestro
                !vchEstatusRequis = "CANCELADA"
                .Update
            End With
            
            vlstrSentencia = "update IvRequisCarDirDetalle set vchEstatusRequis='CANCELADA' where intNumRequisCarDir = " & txtNumero.Text
            pEjecutaSentencia vlstrSentencia
            
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            pAgregarRegistro ("")
        End If
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCancelarRequisicion_Click"))
    Unload Me
End Sub

Private Sub cmdGrabarRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite crear un nuevo registro o actualizar la información de un registro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    Dim vlintNumRequisCD As Double
    Dim vlintSeqFil As Integer
    Dim vlstrMensaje As String
    Dim rs As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    Dim vlmsgActivosInactivos As String

    If cboDepartamento.ItemData(cboDepartamento.ListIndex) <> vgintNumeroDepartamento Then
        MsgBox SIHOMsg(1261), vbOKOnly + vbExclamation, "Mensaje"
        Exit Sub
    End If
    
    If vlblnNuevoRegistro Then
        vlmsgActivosInactivos = ""
        For vlintSeqFil = 1 To grdHRequisicionCD.Rows - 1
            Set rs = frsRegresaRs("SELECT vchdescripcion, bitactivo FROM IVARTICULOCARGODIRECTO WHERE intcvearticulocardir = " & grdHRequisicionCD.TextMatrix(vlintSeqFil, 5))
            If rs.RecordCount <> 0 Then
                If rs!bitactivo = 0 Then
                    vlmsgActivosInactivos = IIf(Trim(vlmsgActivosInactivos) = "", Trim(rs!vchDescripcion), vlmsgActivosInactivos & Chr(13) & Trim(rs!vchDescripcion))
                End If
            End If
        Next vlintSeqFil
        
        If Trim(vlmsgActivosInactivos) <> "" Then
            MsgBox SIHOMsg(1256) & Chr(13) & vlmsgActivosInactivos, vbInformation, "Mensaje"
            
            If grdHRequisicionCD.Enabled Then
                grdHRequisicionCD.SetFocus
            End If
            Exit Sub
        End If
    End If

    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then

        If Not vlblnNuevoRegistro Then
            pValidaEstatus
        End If
            
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
                With rsIvRequisCarDirMaestro
                    If vlblnNuevoRegistro Then
                        .AddNew
                        !smiCveDeptoRequis = cboDepartamento.ItemData(cboDepartamento.ListIndex)
                        !intCveEmpleaRequis = vllngPersonaGraba
                        !dtmFechaRequis = fdtmServerFecha
                        !dtmHoraRequis = fdtmServerHora
                        !bitUrgenteRequis = IIf(chkUrgente.Value = 1, 1, 0)
                        !vchEstatusRequis = Trim(txtEstatusMaestro.Text)
                        !dtmFechaAutoriza = Null
                        !vchObservaciones = Trim(txtObservaciones.Text)
                        .Update
                        txtNumero.Text = CStr(flngObtieneIdentity("SEC_IvRequisCarDirMaestro", !INTNUMREQUISCARDIR))
                        lintSoloMaestro = 0
                    Else
                        !vchEstatusRequis = Trim(txtEstatusMaestro.Text)
                        If txtObservaciones.Enabled Then !vchObservaciones = Trim(txtObservaciones.Text)
                        .Update
                        txtNumero.Text = CStr(!INTNUMREQUISCARDIR)
                    End If
                End With
                
                rsIvRequisCarDirMaestro.Close
                
                If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCancelar, "C") Then
                    vlstrSentencia = "select INTNUMREQUISCARDIR, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, DTMFECHAREQUIS, DTMHORAREQUIS, BITURGENTEREQUIS, VCHESTATUSREQUIS, DTMFECHAAUTORIZA, VCHOBSERVACIONES from IvRequisCarDirMaestro order by intNumRequisCarDir"
                Else
                    vlstrSentencia = "select INTNUMREQUISCARDIR, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, DTMFECHAREQUIS, DTMHORAREQUIS, BITURGENTEREQUIS, VCHESTATUSREQUIS, DTMFECHAAUTORIZA, VCHOBSERVACIONES from IvRequisCarDirMaestro where smiCveDeptoRequis = " & Str(vgintNumeroDepartamento) & " order by intNumRequisCarDir"
                End If
                Set rsIvRequisCarDirMaestro = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                rsIvRequisCarDirMaestro.MoveFirst
                rsIvRequisCarDirMaestro.Find ("intNumRequisCarDir=" & txtNumero.Text)
                
                If lintSoloMaestro = 0 Then
                
                    vlstrSentencia = "delete FROM IvRequisCarDirDetalle where intNumRequisCarDir=" & txtNumero.Text
                    pEjecutaSentencia vlstrSentencia
                
                    With rsIvRequisCarDirDetalle
                        For vlintSeqFil = 1 To grdHRequisicionCD.Rows - 1
                            .AddNew
                            !INTNUMREQUISCARDIR = CLng(txtNumero.Text)
                            !INTCVEARTICULOCARDIR = grdHRequisicionCD.TextMatrix(vlintSeqFil, 5)
                            !vchDescLarga = grdHRequisicionCD.TextMatrix(vlintSeqFil, 4)
                            !SMICANTIDAD = CLng(grdHRequisicionCD.TextMatrix(vlintSeqFil, 2))
                            !vchEstatusRequis = grdHRequisicionCD.TextMatrix(vlintSeqFil, 3)
                            .Update
                        Next vlintSeqFil
                    End With
                    
                    rsIvRequisCarDirDetalle.Close
                    vlstrSentencia = "select INTNUMREQUISCARDIR, INTCVEARTICULOCARDIR, VCHDESCLARGA, SMICANTIDAD, VCHESTATUSREQUIS from IvRequisCarDirDetalle where intNumRequisCarDir = -1"
                    Set rsIvRequisCarDirDetalle = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                    rsIvRequisCarDirDetalle.Find ("intNumRequisCarDir=" & txtNumero.Text)
                End If
                
                pImpresionRemota "RD", CLng(txtNumero.Text), 0
                Call pGuardarLogTransaccion(Me.Name, IIf(vlblnNuevoRegistro, EnmGrabar, EnmCambiar), vllngPersonaGraba, "REQUISICION DE CARGOS DIRECTOS", txtNumero.Text)
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            pMuestraRegistro
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    Dim alstrParametros(0) As String
    Dim vlstrsql As String
    vlstrx = CStr(Val(txtNumero.Text))
    Dim rsDepartamento As ADODB.Recordset
    Dim vlintnumdepartamento As Integer
    Dim strSql As String
    Dim rstipodepapel As ADODB.Recordset
    Set rsReporte = frsEjecuta_SP(vlstrx, "sp_CORptRequisicioCargoDir")
    
    
    vlstrsql = "SELECT SMICVEDEPARTAMENTO FROM LOGIN WHERE INTNUMEROLOGIN=" & vglngNumeroLogin
                    Set rsDepartamento = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
                
                    If Not rsDepartamento.EOF Then
                        vlintnumdepartamento = rsDepartamento!smiCveDepartamento
                    End If
                    
                    strSql = "SELECT INTWIDTH,INTHEIGHT FROM SIPAPELDOCDEPTOMODULO INNER JOIN SITIPODEPAPEL ON SIPAPELDOCDEPTOMODULO.INTIDPAPEL=SITIPODEPAPEL.INTCVETIPODEPAPEL   WHERE VCHMODULO= '" & Trim(cgstrModulo) & "' AND VCHTIPODOCUMENTO='RD' AND INTDEPTO=" & vlintnumdepartamento & ""
                    Set rstipodepapel = frsRegresaRs(strSql, adLockOptimistic, adOpenDynamic)
    
    If rsReporte.RecordCount > 0 Then
      vgrptReporte.DiscardSavedData
      
      If Not rstipodepapel.EOF Then
                    vgrptReporte.SetUserPaperSize rstipodepapel!INTHEIGHT, rstipodepapel!INTWIDTH
                    vgrptReporte.PaperSize = crPaperUser
      End If
      
      alstrParametros(0) = "Empresa;" & Trim(vgstrNombreHospitalCH)
      pCargaParameterFields alstrParametros, vgrptReporte
      pImprimeReporte vgrptReporte, rsReporte, "P", "Requisición de artículos"
    Else
      'No existe información con esos parámetros.
      MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    If rsReporte.State <> adStateClosed Then rsReporte.Close
    'txtNumero.SetFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprimir_Click"))
    Unload Me
End Sub

Private Sub cmdPrimerRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite localizarse en el primer registro del RS
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If vlblnCmdFiltrado Then
        If fintLocalizaPkRs(rsIvRequisCarDirMaestro, 0, grdhBusqueda.TextMatrix(1, 1)) <> 0 Then
            pMuestraRegistro
        End If
    Else
        rsIvRequisCarDirMaestro.MoveFirst
        pMuestraRegistro
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimerRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdSiguienteRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite localizarse en el siguiente registro del RS
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If vlblnCmdFiltrado Then
        If grdhBusqueda.Row <> grdhBusqueda.Rows - 1 Then
            grdhBusqueda.Row = grdhBusqueda.Row + 1
        End If
        If fintLocalizaPkRs(rsIvRequisCarDirMaestro, 0, grdhBusqueda.TextMatrix(grdhBusqueda.Row, 1)) <> 0 Then
            pMuestraRegistro
        End If
    Else
        rsIvRequisCarDirMaestro.MoveNext
        If rsIvRequisCarDirMaestro.EOF Then
            rsIvRequisCarDirMaestro.MovePrevious
        End If
        pMuestraRegistro
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguienteRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdUltimoRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite localizarse en el último registro del RS
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If vlblnCmdFiltrado Then
        If fintLocalizaPkRs(rsIvRequisCarDirMaestro, 0, grdhBusqueda.TextMatrix(grdhBusqueda.Rows - 1, 1)) <> 0 Then
            pMuestraRegistro
        End If
    Else
        rsIvRequisCarDirMaestro.MoveLast
        pMuestraRegistro
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimoRegistro_Click"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError 'Manejo del error

    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError 'Manejo del error
    
    'Color de tabs
    SetStyle sstObj.hwnd, 0
    SetSolidColor sstObj.hwnd, 16777215
    SSTabSubclass sstObj.hwnd
    
    Me.Icon = frmMenuPrincipal.Icon
    pInstanciaReporte vgrptReporte, "rptReqCarDir.rpt"
    vgstrNombreForm = Me.Name
    
    vllngNumeroOpcionCancelar = flngObtenOpcion(cmdCancelarRequisicion.Name)
    
    '----------------------------------
    ' IvRequisCarDirMaestro
    '----------------------------------
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCancelar, "C") Then
        vlstrSentencia = "select INTNUMREQUISCARDIR, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, DTMFECHAREQUIS, DTMHORAREQUIS, BITURGENTEREQUIS, VCHESTATUSREQUIS, DTMFECHAAUTORIZA, VCHOBSERVACIONES from IvRequisCarDirMaestro order by intNumRequisCarDir"
    Else
        vlstrSentencia = "select INTNUMREQUISCARDIR, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, DTMFECHAREQUIS, DTMHORAREQUIS, BITURGENTEREQUIS, VCHESTATUSREQUIS, DTMFECHAAUTORIZA, VCHOBSERVACIONES from IvRequisCarDirMaestro where smiCveDeptoRequis = " & Str(vgintNumeroDepartamento) & " order by intNumRequisCarDir"
    End If
    Set rsIvRequisCarDirMaestro = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    '----------------------------------
    ' IvRequisCarDirDetalle
    '----------------------------------
    vlstrSentencia = "select INTNUMREQUISCARDIR, INTCVEARTICULOCARDIR, VCHDESCLARGA, SMICANTIDAD, VCHESTATUSREQUIS from IvRequisCarDirDetalle where intNumRequisCarDir = -1"
    Set rsIvRequisCarDirDetalle = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    vgstrNombreForm = Me.Name 'Nombre del formulario que se utiliza actualmente
    vgblnExistioError = False 'Inicia la bandera sin errores
    vgblnErrorIngreso = False
    vlblnmodificaarticulo = False
    vlblnModificaRegistro = False
    vlblnCmdFiltrado = False
    sstObj.Tab = 0 'Se localiza en el primer tabulador para la alta
    vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
    vgintTipoOrd = 1 'Que tipo de ordenamiento realizará de inicio en el grdHBusqueda
    vgintColLoc = 1 'Localiza la búsqueda de registros para la primera columna del grdHBusqueda
    
    pLlenarCboDepartamento
    pLlenarCboEstGrid
    
    pAgregarRegistro ("I") 'Permite agregar un registro nuevo
    
    cboDepartamento.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    If sstObj.Tab <> 0 Then
        Cancel = True
        sstObj.Tab = 0
        pAgregarRegistro ("")
    Else
        If cmdGrabarRegistro.Enabled Or Not vlblnNuevoRegistro Then
            Cancel = True
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                pAgregarRegistro ("")
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub



Private Sub grdHBusqueda_Click()
'-------------------------------------------------------------------------------------------
' Refresca el GrdHBusqueda y asigna bajo que columna se va a hacer la búsqueda
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If grdhBusqueda.Rows > 0 Then
        grdhBusqueda.Refresh
        vgintColLoc = grdhBusqueda.Col
        vgstrAcumTextoBusqueda = ""
        grdhBusqueda.Col = vgintColLoc
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_Click"))
    Unload Me
End Sub

Private Sub grdhBusqueda_DblClick()
'-------------------------------------------------------------------------------------------
' Muestra la información del registro encontrado y habilita su posible modificación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Integer
    
    vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
    
    If grdhBusqueda.Rows > 0 Then
        If grdhBusqueda.MouseRow >= grdhBusqueda.FixedRows Then
            
            If fintLocalizaPkRs(rsIvRequisCarDirMaestro, 0, grdhBusqueda.TextMatrix(grdhBusqueda.Row, 1)) <> 0 Then
                vlblnCmdFiltrado = True
                pMuestraRegistro
                sstObj.Tab = 0
            End If
        End If
        vgintColOrdAnt = vgintColOrd 'Guarda la columna de ordenación anterior
        vgintColOrd = grdhBusqueda.Col  'Configura la columna a ordenar
        
        'Escoge el Tipo de Ordenamiento
        If vgintTipoOrd = 1 Then
            vgintTipoOrd = 2
        Else
            vgintTipoOrd = 1
        End If
        Call pOrdColMshFGrid(grdhBusqueda, vgintTipoOrd)
        Call pDesSelMshFGrid(grdhBusqueda)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_DblClick"))
    Unload Me
End Sub

Private Sub grdHRequisicionCD_DblClick()
'-------------------------------------------------------------------------------------------
' Muestra la información del registro encontrado y habilita su posible modificación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Integer
       
    vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
    ' Ordena solamente cuando un encabezado de columna es seleccionado con un click
    If grdHRequisicionCD.Rows > 0 Then
        If (grdHRequisicionCD.Row <= grdHRequisicionCD.Rows - 1) Then
            If grdHRequisicionCD.MouseRow >= grdHRequisicionCD.FixedRows Then
                vlintPosRegGrid = grdHRequisicionCD.MouseRow
                pModificaArticulo
                'cmdAgregarGrid.Enabled = True
                Exit Sub
            End If
            
            vgintColOrdAnt = vgintColOrd 'Guarda la columna de ordenación anterior
            vgintColOrd = grdHRequisicionCD.Col  'Configura la columna a ordenar
            
            'Escoge el Tipo de Ordenamiento
            If vgintTipoOrd = 1 Then
                 vgintTipoOrd = 2
                Else
                    vgintTipoOrd = 1
                End If
            Call pOrdColMshFGrid(grdHRequisicionCD, vgintTipoOrd)
            Call pDesSelMshFGrid(grdHRequisicionCD)
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHRequisicionCD_DblClick"))
    Unload Me
End Sub

Private Sub grdHRequisicionCD_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
' Muestra la información del registro encontrado y habilita su posible modificación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Integer
    Dim vlintResul As Integer
    Dim vlstrMensaje As String
    
    Select Case KeyCode
        Case vbKeyReturn
            grdHRequisicionCD_DblClick
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHRequisicionCD_KeyDown"))
    Unload Me
End Sub

Private Sub MyButton1_Click()

End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
    Case vbKeyReturn
        If txtDescripcionLarga.Enabled Then txtDescripcionLarga.SetFocus
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidad_KeyDown"))
    Unload Me
End Sub

Private Sub txtCantidad_LostFocus()
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Call pValidaMkText(txtCantidad, "N", ">", 6, True)
    If vgblnErrorIngreso = False Then
       If vlblnNuevoRegistro Then
         txtEstatusDetalle.Text = "PENDIENTE"
         txtDescripcionLarga.SelStart = Len(txtDescripcionLarga)
         txtDescripcionLarga.SetFocus
       Else
         If txtEstatusMaestro.Text = "PENDIENTE" Then
             txtEstatusDetalle.Text = "PENDIENTE"
             txtDescripcionLarga.SelStart = Len(txtDescripcionLarga)
             txtDescripcionLarga.SetFocus
         End If
         
         If grdHRequisicionCD.Rows - 1 > 0 Then
             If fEstatusDetalle(grdHRequisicionCD.Row) = "COTIZADA" Then
                If vlblnmodificaarticulo = False Then
                    cmdAgregarGrid_Click
                End If
             Else
               txtDescripcionLarga.SelStart = Len(txtDescripcionLarga)
               txtDescripcionLarga.SetFocus
             End If
         End If
       End If
    Else
       Call pEnfocaMkTexto(txtCantidad)
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidad_LostFocus"))
    Unload Me
End Sub

Private Sub txtDescripcionLarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 124 Then
        KeyAscii = 0
        Beep
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtNumero_GotFocus()
pSelMkTexto txtNumero

End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
'modificar uno que ya existe
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    Dim vlstrMensaje As String
    
    Select Case KeyCode
        Case vbKeyReturn
             
            'Buscar criterio
            If (Len(txtNumero.Text) <= 0) Then
                txtNumero.Text = "0"
            End If
            If fintSigNumRs(rsIvRequisCarDirMaestro, 0) = CDbl(txtNumero.Text) Then
                txtNumero.Enabled = False
                chkUrgente.Enabled = True
                chkUrgente.SetFocus
            Else
                If fintLocalizaPkRs(rsIvRequisCarDirMaestro, 0, txtNumero.Text) <> 0 Then
                    pMuestraRegistro
                    txtNumero.Enabled = False
                    Me.grdHRequisicionCD.SetFocus
                Else 'La información no existe
                    Call MsgBox(SIHOMsg("12"), vbExclamation, "Mensaje")
                    pAgregarRegistro ("")
                    Call pEnfocaMkTexto(txtNumero)
                End If
            End If
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumero_KeyDown"))
    Unload Me
End Sub

Private Sub txtObservaciones_GotFocus()
    If Not vlblnNuevoRegistro Then cmdGrabarRegistro.Enabled = True
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fblnCanFocus(cmdGrabarRegistro) Then cmdGrabarRegistro.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


