VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmPermisos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de permisos al módulo"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      Left            =   4170
      TabIndex        =   33
      Top             =   9600
      Width           =   4320
      Begin MyCommandButton.MyButton cmdTop 
         Height          =   600
         Left            =   60
         TabIndex        =   15
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
         Picture         =   "frmPermisos.frx":0000
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmPermisos.frx":0982
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdBack 
         Height          =   600
         Left            =   660
         TabIndex        =   16
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
         Picture         =   "frmPermisos.frx":1304
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         DropDownPicture =   "frmPermisos.frx":1C86
         PictureDisabled =   "frmPermisos.frx":1CA2
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdLocate 
         Height          =   600
         Left            =   1260
         TabIndex        =   17
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
         Picture         =   "frmPermisos.frx":2624
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmPermisos.frx":2FA8
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdNext 
         Height          =   600
         Left            =   1860
         TabIndex        =   18
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
         Picture         =   "frmPermisos.frx":392C
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmPermisos.frx":42AE
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdEnd 
         Height          =   600
         Left            =   2460
         TabIndex        =   19
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
         Picture         =   "frmPermisos.frx":4C30
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmPermisos.frx":55B2
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdSave 
         Height          =   600
         Left            =   3060
         TabIndex        =   20
         ToolTipText     =   "Grabar información"
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
         Picture         =   "frmPermisos.frx":5F34
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmPermisos.frx":68B8
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdDelete 
         Height          =   600
         Left            =   3660
         TabIndex        =   21
         ToolTipText     =   "Borrar información"
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
         Picture         =   "frmPermisos.frx":723C
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmPermisos.frx":7BBE
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   4680
         Y1              =   0
         Y2              =   0
      End
   End
   Begin HSFlatControls.MyTabHeader MyTabHeader1 
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   741
      Tabs            =   2
      TabCurrent      =   0
      TabWidth        =   6340
      Caption         =   $"frmPermisos.frx":8542
   End
   Begin TabDlg.SSTab sstObj 
      Height          =   12720
      Left            =   -15
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   360
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   22437
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Usuarios"
      TabPicture(0)   =   "frmPermisos.frx":8558
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Perfiles"
      TabPicture(1)   =   "frmPermisos.frx":8574
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
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
         Height          =   9250
         Left            =   -74880
         TabIndex        =   34
         Top             =   0
         Width           =   12460
         Begin VB.TextBox txtDescripcionPerfil 
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
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   37
            ToolTipText     =   "Descripción del Perfil "
            Top             =   700
            Width           =   9820
         End
         Begin VB.TextBox txtClavePerfil 
            Alignment       =   1  'Right Justify
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
            Left            =   2520
            TabIndex        =   35
            ToolTipText     =   "Clave del Perfil"
            Top             =   300
            Width           =   1215
         End
         Begin HSFlatControls.MyCombo cboModuloPerfiles 
            Height          =   375
            Left            =   3765
            TabIndex        =   36
            Top             =   300
            Width           =   8580
            _ExtentX        =   15134
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPermisos1 
            Height          =   7660
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Opciones del módulo"
            Top             =   1455
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   13520
            _Version        =   393216
            ForeColor       =   0
            Rows            =   0
            FixedRows       =   0
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            ScrollTrack     =   -1  'True
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
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción del Perfil"
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
            Left            =   120
            TabIndex        =   38
            Top             =   760
            Width           =   1995
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Permisos de acceso al sistema"
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
            Height          =   360
            Left            =   120
            TabIndex        =   40
            Top             =   1110
            Width           =   12220
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Clave"
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
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   585
         End
      End
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
         Height          =   9260
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   12460
         Begin VB.TextBox txtPersonaAutoriza 
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   12
            ToolTipText     =   "17"
            Top             =   3540
            Width           =   9340
         End
         Begin VB.TextBox txtClaveLogin 
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
            Left            =   3000
            TabIndex        =   1
            ToolTipText     =   "Clave del login"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtNombreDepartamento 
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
            Left            =   3000
            TabIndex        =   8
            ToolTipText     =   "Nombre del departamento"
            Top             =   2320
            Width           =   9340
         End
         Begin VB.TextBox txtContraseña 
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
            IMEMode         =   3  'DISABLE
            Left            =   3000
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   4
            ToolTipText     =   "Contraseña para acceso al módulo"
            Top             =   1110
            Width           =   2220
         End
         Begin VB.TextBox txtUsuario 
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
            Left            =   3000
            MaxLength       =   20
            TabIndex        =   3
            ToolTipText     =   "Nombre de usuario para acceso al módulo"
            Top             =   700
            Width           =   9340
         End
         Begin VB.TextBox txtVerifContraseña 
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
            IMEMode         =   3  'DISABLE
            Left            =   10130
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   5
            ToolTipText     =   "Verificar contraseña para acceso al módulo"
            Top             =   1110
            Width           =   2220
         End
         Begin HSFlatControls.MyCombo cboPerfiles 
            Height          =   375
            Left            =   4600
            TabIndex        =   11
            ToolTipText     =   "Perfiles"
            Top             =   3135
            Width           =   7740
            _ExtentX        =   13653
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
         Begin HSFlatControls.MyCombo cboDepartamento 
            Height          =   375
            Left            =   3000
            TabIndex        =   7
            ToolTipText     =   "Selección del departamento con que entrará el usuario"
            Top             =   1920
            Width           =   9340
            _ExtentX        =   16484
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
            Left            =   3000
            TabIndex        =   6
            ToolTipText     =   "Selección del empleado al que pertenece el login de usuario"
            Top             =   1510
            Width           =   9340
            _ExtentX        =   16484
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
         Begin HSFlatControls.MyCombo cboModulo 
            Height          =   375
            Left            =   8130
            TabIndex        =   2
            Top             =   300
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPermisos 
            Height          =   4840
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Opciones del módulo"
            Top             =   4290
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   8546
            _Version        =   393216
            ForeColor       =   0
            Rows            =   0
            FixedRows       =   0
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            ScrollTrack     =   -1  'True
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            AllowUserResizing=   1
            Appearance      =   0
            RowSizingMode   =   1
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
         End
         Begin MSMask.MaskEdBox mskFechaFinal 
            Height          =   375
            Left            =   3000
            TabIndex        =   10
            ToolTipText     =   "Fecha en que expira el login"
            Top             =   3135
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox mskFechaInicial 
            Height          =   375
            Left            =   3000
            TabIndex        =   9
            ToolTipText     =   "Fecha de inicio del login"
            Top             =   2730
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Clave"
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
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   585
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   180
            X2              =   180
            Y1              =   4620
            Y2              =   4950
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Persona que autoriza"
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
            Left            =   120
            TabIndex        =   23
            Top             =   3600
            Width           =   2055
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Permisos de acceso al sistema"
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
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Top             =   3945
            Width           =   12220
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fecha final"
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
            Left            =   120
            TabIndex        =   25
            Top             =   3200
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fecha inicial"
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
            Left            =   120
            TabIndex        =   26
            Top             =   2790
            Width           =   1200
         End
         Begin VB.Label Label5 
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
            Left            =   120
            TabIndex        =   27
            Top             =   1570
            Width           =   1005
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "Nombre del departamento"
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
            Left            =   120
            TabIndex        =   28
            Top             =   2380
            Width           =   3030
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Departamento"
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
            Left            =   120
            TabIndex        =   29
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Contraseña"
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
            Left            =   120
            TabIndex        =   30
            Top             =   1170
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Usuario"
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
            Left            =   120
            TabIndex        =   31
            Top             =   760
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Perfil"
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
            Left            =   4600
            TabIndex        =   42
            Top             =   2790
            Width           =   450
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000005&
            Caption         =   "Módulo"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7320
            TabIndex        =   43
            Top             =   360
            Width           =   735
         End
         Begin VB.Label ConfirmarContraseña 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Confirmar contraseña"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7920
            TabIndex        =   44
            Top             =   1200
            Width           =   2130
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   8280
         TabIndex        =   0
         Top             =   7200
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
' Programa para asignación de permisos a las opciones de los módulos
' Fecha de programación: 30 de Noviembre del 2000
'-----------------------------------------------------------------------------------
' Ultimas modificaciones:
' Fecha:
' Descripción del cambio:
'-----------------------------------------------------------------------------------
Option Explicit
Dim rsEmpleados As New ADODB.Recordset
Dim rsDepartamentos As New ADODB.Recordset
Dim rsOpciones As New ADODB.Recordset
Dim rsLogin As New ADODB.Recordset
Dim rsPermiso As New ADODB.Recordset
Dim rsPerfiles As New ADODB.Recordset
Dim rsPerfilesDetalle As New ADODB.Recordset

Dim vlstrsql As String
Dim vlintPerfilUsuario As Integer
Dim vlblnActivaTabPerfiles As Boolean
Dim vlblnConsulta As Boolean
Dim vlblnNoCombo As Boolean
Dim vlblnHaCambiado As Boolean

Public vllngNumeroOpcion As Long
Dim vlblnNoEsPrimerEntrada As Boolean

Private Sub grdPermisos_DblClick()

Dim vlintRenglon As Integer

    With grdPermisos
        If .MouseRow = 0 Then
            If .Col > 2 Then
                If .Col = 3 Then
                    For vlintRenglon = 1 To grdPermisos.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = "x"
                        .TextMatrix(vlintRenglon, 4) = ""
                        .TextMatrix(vlintRenglon, 5) = ""
                        .TextMatrix(vlintRenglon, 6) = ""
                    Next vlintRenglon
                    vlblnHaCambiado = True
                End If
                If .Col = 4 Then
                    For vlintRenglon = 1 To grdPermisos.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = ""
                        .TextMatrix(vlintRenglon, 4) = "x"
                        .TextMatrix(vlintRenglon, 5) = ""
                        .TextMatrix(vlintRenglon, 6) = ""
                    Next vlintRenglon
                    vlblnHaCambiado = True
                End If
                If .Col = 5 Then
                    For vlintRenglon = 1 To grdPermisos.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = ""
                        .TextMatrix(vlintRenglon, 4) = ""
                        .TextMatrix(vlintRenglon, 5) = "x"
                        .TextMatrix(vlintRenglon, 6) = ""
                    Next vlintRenglon
                    vlblnHaCambiado = True
                End If
                If .Col = 6 Then
                    For vlintRenglon = 1 To grdPermisos.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = ""
                        .TextMatrix(vlintRenglon, 4) = ""
                        .TextMatrix(vlintRenglon, 5) = ""
                        .TextMatrix(vlintRenglon, 6) = "x"
                    Next vlintRenglon
                    vlblnHaCambiado = True
                End If
            End If
        End If
    End With

End Sub

Private Sub grdPermisos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = 13 Then
        With grdPermisos
            If .Row > 0 Then
                If .Col > 2 Then
                    .TextMatrix(.Row, 3) = ""
                    .TextMatrix(.Row, 4) = ""
                    .TextMatrix(.Row, 5) = ""
                    .TextMatrix(.Row, 6) = ""
                    If .Col = 3 Then
                        .TextMatrix(.Row, 3) = "x"
                        vlblnHaCambiado = True
                    End If
                    If .Col = 4 Then
                        .TextMatrix(.Row, 4) = "x"
                        vlblnHaCambiado = True
                    End If
                    If .Col = 5 Then
                        .TextMatrix(.Row, 5) = "x"
                        vlblnHaCambiado = True
                    End If
                    If .Col = 6 Then
                        .TextMatrix(.Row, 6) = "x"
                        vlblnHaCambiado = True
                    End If
                End If
            End If
        End With
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPermisos_Click"))


End Sub

Private Sub Label14_Click()

End Sub

Private Sub grdPermisos1_DblClick()

Dim vlintRenglon As Integer

    With grdPermisos1
        If .MouseRow = 0 Then
            If .Col > 2 Then
                If .Col = 3 Then
                    For vlintRenglon = 1 To grdPermisos1.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = "x"
                        .TextMatrix(vlintRenglon, 4) = ""
                        .TextMatrix(vlintRenglon, 5) = ""
                        .TextMatrix(vlintRenglon, 6) = ""
                    Next vlintRenglon
                End If
                If .Col = 4 Then
                    For vlintRenglon = 1 To grdPermisos1.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = ""
                        .TextMatrix(vlintRenglon, 4) = "x"
                        .TextMatrix(vlintRenglon, 5) = ""
                        .TextMatrix(vlintRenglon, 6) = ""
                    Next vlintRenglon
                End If
                If .Col = 5 Then
                    For vlintRenglon = 1 To grdPermisos1.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = ""
                        .TextMatrix(vlintRenglon, 4) = ""
                        .TextMatrix(vlintRenglon, 5) = "x"
                        .TextMatrix(vlintRenglon, 6) = ""
                    Next vlintRenglon
                End If
                If .Col = 6 Then
                    For vlintRenglon = 1 To grdPermisos1.Rows - 1
                        .TextMatrix(vlintRenglon, 3) = ""
                        .TextMatrix(vlintRenglon, 4) = ""
                        .TextMatrix(vlintRenglon, 5) = ""
                        .TextMatrix(vlintRenglon, 6) = "x"
                    Next vlintRenglon
                End If
            End If
        End If
    End With
    
End Sub

Private Sub MyTabHeader1_Click(Index As Integer)
        sstObj.Tab = Index
End Sub

Private Sub MyTabHeader1_GotFocus()
On Error GoTo NotificaError
    
    If sstObj.Tab = 1 Then
       txtClavePerfil.SetFocus
    Else
       vlblnNoCombo = False
       txtClaveLogin.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstObj_GotFocus"))

End Sub

Private Sub MyTabHeader1_LostFocus()
On Error GoTo NotificaError
    
    If sstObj.Tab = 1 Then
       txtClavePerfil.SetFocus
    Else
       vlblnNoCombo = False
       txtClaveLogin.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstObj_GotFocus"))

End Sub

Private Sub sstObj_GotFocus()
    On Error GoTo NotificaError
    
    If sstObj.Tab = 1 Then
       txtClavePerfil.SetFocus
    Else
       vlblnNoCombo = False
       txtClaveLogin.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstObj_GotFocus"))

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtVerifContraseña

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtVerifContraseña_GotFocus"))

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cboEmpleado.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtVerifContraseña_KeyPress"))

End Sub

Private Sub txtDescripcionPerfil_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        grdPermisos1.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcionPerfil_KeyPress"))

End Sub

Private Sub grdPermisos_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPermisos_GotFocus"))

End Sub

Private Sub grdPermisos1_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPermisos1_GotFocus"))

End Sub

Private Sub mskFechaFinal_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskFechaFinal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinal_GotFocus"))

End Sub

Private Sub mskFechaFinal_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        grdPermisos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinal_KeyPress"))

End Sub

Private Sub mskFechaFinal_LostFocus()
    On Error GoTo NotificaError
    
    If mskFechaFinal.ClipText = "" Then
        mskFechaFinal.Text = fdtmServerFecha
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinal_LostFocus"))

End Sub

Private Sub mskFechaInicial_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskFechaInicial

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicial_GotFocus"))

End Sub

Private Sub mskFechaInicial_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFechaFinal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicial_KeyPress"))

End Sub

Private Sub mskFechaInicial_LostFocus()
    On Error GoTo NotificaError
    
    If mskFechaInicial.ClipText = "" Then
        mskFechaInicial.Text = fdtmServerFecha
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicial_LostFocus"))

End Sub

Private Sub txtClaveLogin_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 1, 1, 1, 1, 1, 0, 0
    pLimpia
    pSelTextBox txtClaveLogin

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClaveLogin_GotFocus"))

End Sub

Private Sub txtClavePerfil_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 1, 1, 1, 1, 1, 0, 0
    pLimpia
    pSelTextBox txtClavePerfil

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClavePerfil_GotFocus"))

End Sub

Private Sub txtClaveLogin_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        vlblnNoCombo = True
        pBusca txtClaveLogin.Text
        pHabilita 1, 1, 1, 1, 1, 0, 1
        txtUsuario.SetFocus
        vlblnNoCombo = False
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClaveLogin_KeyPress"))

End Sub

Private Sub txtClavePerfil_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        pBusca txtClavePerfil.Text
        txtDescripcionPerfil.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClavePerfil_KeyPress"))

End Sub

Private Sub txtContraseña_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtContraseña

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtContraseña_GotFocus"))

End Sub

Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
            txtVerifContraseña.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtContraseña_KeyPress"))

End Sub

Private Sub cboPerfiles_KeyDown(KeyAscii As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        grdPermisos.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboPerfiles_KeyPress"))

End Sub

Private Sub txtNombreDepartamento_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtNombreDepartamento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNombreDepartamento_GotFocus"))

End Sub


Private Sub txtNombreDepartamento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFechaInicial.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNombreDepartamento_KeyPress"))

End Sub

Private Sub txtPersonaAutoriza_GotFocus()
    grdPermisos.SetFocus
End Sub

Private Sub txtUsuario_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtUsuario

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtUsuario_GotFocus"))

End Sub

Private Sub txtDescripcionPerfil_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescripcionPerfil

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtUsuario_GotFocus"))

End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtContraseña.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtUsuario_KeyPress"))

End Sub

Private Sub cboDepartamento_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_GotFocus"))

End Sub

Private Sub cboDepartamento_KeyDown(KeyAscii As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If cboDepartamento.ListIndex >= 0 Then
            txtNombreDepartamento.Text = cboDepartamento.List(cboDepartamento.ListIndex)
        End If
        txtNombreDepartamento.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_KeyPress"))

End Sub

Private Sub cboEmpleado_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_GotFocus"))

End Sub

Private Sub cboEmpleado_KeyDown(KeyAscii As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cboDepartamento.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_KeyPress"))

End Sub

Private Sub cboPerfiles_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboPerfiles_GotFocus"))

End Sub

Private Sub cboModuloPerfiles_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboModuloPerfiles_GotFocus"))

End Sub

Private Sub cboModulo_Click()
    On Error GoTo NotificaError
    
    Dim vlstrsql As String
    Dim rsHayPerfilGuardado As New ADODB.Recordset

   If vlblnNoCombo Then
        Exit Sub
   End If

   vlstrsql = "select vchDescripcion,intNumeroOpcion from Opcion where smiNumeroModulo=" + STR(cboModulo.ItemData(cboModulo.ListIndex)) + " order by intOrdenOpcion, vchDescripcion "
   Set rsOpciones = frsRegresaRs(vlstrsql)
   
   If rsOpciones.RecordCount <> 0 Then
      pLlenarMshFGrdRs grdPermisos, rsOpciones, 1
       pConfigura
      pCargarYPosicionarPeril "M", cboModulo.ItemData(cboModulo.ListIndex)
      If rsLogin.State <> 0 Then
         If vlblnConsulta Then
            pHabilita 0, 0, 0, 0, 0, 1, 0
         Else
            pHabilita 1, 1, 1, 1, 1, 0, 1
         End If
      End If
      
        vlstrsql = "select * from SIPERFILESGUARDADOS where INTCVEMODULO = " + CStr(cboModulo.ItemData(cboModulo.ListIndex)) + " AND INTCVEUSUARIO = " + txtClaveLogin.Text
        Set rsHayPerfilGuardado = frsRegresaRs(vlstrsql)
        If rsHayPerfilGuardado.RecordCount > 0 Then
            If rsHayPerfilGuardado!INTCVEPERFIL <> 0 Then
                pCargarYPosicionarPeril "P", rsHayPerfilGuardado!INTCVEPERFIL
            Else
                pPosiciona cboPerfiles, 0
                pCargarPermisosSinPerfil
            End If
        End If
   End If
   
   
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboModulo_Click"))
End Sub

Private Sub cboModuloPerfiles_Click()
   
   vlblnActivaTabPerfiles = True
   
   vlstrsql = "select vchDescripcion,intNumeroOpcion from Opcion where smiNumeroModulo=" + STR(cboModuloPerfiles.ItemData(cboModuloPerfiles.ListIndex)) + " order by intOrdenOpcion, vchDescripcion"
   Set rsOpciones = frsRegresaRs(vlstrsql)
   If rsOpciones.RecordCount <> 0 Then
      pLlenarMshFGrdRs grdPermisos1, rsOpciones, 1
      pConfiguraModuloPerfiles
      If rsPerfiles.State <> 0 Then
         pLimpiaGridPerfiles
      End If
   End If
   
   vlblnActivaTabPerfiles = False
   
End Sub

Private Sub cboPerfiles_Click()
   Dim X As Integer
   Dim vlstrParametro As String
   
   vlstrParametro = STR(cboPerfiles.ItemData(cboPerfiles.ListIndex))
   If Trim(vlstrParametro) <> "0" Then
     vlstrsql = "select vchDescripcion,Opcion.intNumeroOpcion from SiPerfilesDetalle JOIN Opcion ON Opcion.intNumeroOpcion=SiPerfilesDetalle.intNumeroOpcion where intPerfil=" + STR(cboPerfiles.ItemData(cboPerfiles.ListIndex)) + " and OPCION.SMINUMEROMODULO = " + STR(cboModulo.ItemData(cboModulo.ListIndex)) + " order by intOrdenOpcion "
   Else
      vlstrsql = "select vchDescripcion,intNumeroOpcion from Opcion where smiNumeroModulo=" + STR(cboModulo.ItemData(cboModulo.ListIndex)) + " order by intOrdenOpcion"
   End If
   
   vlblnHaCambiado = False
   Set rsOpciones = frsRegresaRs(vlstrsql)
   If rsOpciones.RecordCount <> 0 Then
      pLlenarMshFGrdRs grdPermisos, rsOpciones, 1
      pConfigura
      If rsLogin.State <> 0 Then
         pBuscaPerfil vlstrParametro
      End If
   Else
      pLimpiaGrid
   End If
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    Dim rsPerfil As ADODB.Recordset
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then
        If sstObj.Tab = 0 Then
        ' ¿Desea eliminar el usuario?
            If MsgBox(SIHOMsg(243), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            
                vlstrsql = "Delete from  Permiso where intNumeroLogin=" + txtClaveLogin.Text
                pEjecutaSentencia vlstrsql
                
                vlstrsql = "select * from SILOGINPERFIL where SILOGINPERFIL.INTNUMEROLOGIN = " + txtClaveLogin.Text
                Set rsPerfil = frsRegresaRs(vlstrsql)
                    
                    If rsPerfil.RecordCount > 0 Then
                         vlstrsql = "Delete from SILOGINPERFIL where SILOGINPERFIL.INTNUMEROLOGIN = " + txtClaveLogin.Text
                         pEjecutaSentencia vlstrsql
                    End If
                
        
                vlstrsql = "Delete from  Login where intNumeroLogin=" + txtClaveLogin.Text
                pEjecutaSentencia vlstrsql
        
                rsLogin.Requery
                rsPermiso.Requery
                Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "USUARIO", txtClaveLogin.Text)
        
                txtClaveLogin.SetFocus
            End If
        Else
        ' ¿Desea eliminar el Perfil?
            If MsgBox(SIHOMsg(584), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                vlstrsql = "Delete from  SiPerfilesDetalle where intPerfil=" + txtClavePerfil.Text
                pEjecutaSentencia vlstrsql
           
                vlstrsql = "Delete from  SiPerfiles where intPerfil=" + txtClavePerfil.Text
                pEjecutaSentencia vlstrsql
        
           
                rsPerfiles.Requery
                rsPerfilesDetalle.Requery
                Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "PERFIL", txtClavePerfil.Text)
        
                txtClavePerfil.SetFocus
            End If
        End If
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    Dim vllngNumeroLogin As Long
    Dim vlintPerfilFinal As Integer
    Dim vlintModulo As Integer
    Dim vlintUsuario As Integer
    Dim X As Integer
    Dim rsUsuario As New ADODB.Recordset
    Dim strSentencia As String
    Dim lintBan As Integer
    Dim lstrPass As String
    Dim vlstrsql As String
    Dim rsHayPerfilGuardado As New ADODB.Recordset
    
    If sstObj.Tab = 1 Then
       cmdSavePerfil_Click
       pLimpia
       pHabilita 1, 1, 1, 1, 1, 0, 0
       Exit Sub
    End If
    
    If cboPerfiles.ListIndex <> -1 Then
        vlintPerfilFinal = cboPerfiles.ItemData(cboPerfiles.ListIndex)
    Else
        vlintPerfilFinal = 0
    End If
    
    If vlblnHaCambiado = True Then
        vlintPerfilFinal = 0
    End If
    
    vlintModulo = cboModulo.ItemData(cboModulo.ListIndex)
    vlintUsuario = txtClaveLogin.Text
    
    'Vamos a guardar el historico de los perfiles, basados en el usuario, modulo y perfil
    vlstrsql = "select * from SIPERFILESGUARDADOS where INTCVEMODULO = " + CStr(vlintModulo) + " AND INTCVEUSUARIO = " + CStr(vlintUsuario)
    Set rsHayPerfilGuardado = frsRegresaRs(vlstrsql)
    If rsHayPerfilGuardado.RecordCount > 0 Then
        vlstrsql = "UPDATE SIPERFILESGUARDADOS SET INTCVEPERFIL = " + CStr(vlintPerfilFinal) + " WHERE INTCVEMODULO = " + CStr(vlintModulo) + " AND INTCVEUSUARIO = " + CStr(vlintUsuario)
        pEjecutaSentencia vlstrsql
    Else
        vlstrsql = "INSERT INTO SIPERFILESGUARDADOS (INTCVEMODULO, INTCVEUSUARIO, INTCVEPERFIL) VALUES (" + CStr(vlintModulo) + ", " + CStr(vlintUsuario) + ", " + CStr(vlintPerfilFinal) + ")"
        pEjecutaSentencia vlstrsql
    End If
    '----------------------
    
    lintBan = 0
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then
        If DatosValidos() Then
            lintBan = 1
            'Nueva forma de guardar los datos
            With rsLogin
              If Not vlblnConsulta Then
                .AddNew
                !vchUsuario = Trim(txtUsuario.Text)
                If txtContraseña.Text <> "" Then
                  !vchPassword = fstrEncrypt(txtContraseña.Text, txtUsuario.Text)
                End If
                !intCveEmpleado = cboEmpleado.ItemData(cboEmpleado.ListIndex)
                !smicvedepartamento = cboDepartamento.ItemData(cboDepartamento.ListIndex)
                !intCveEmpAutoriza = vglngNumeroEmpleado
                !dtmFechaInicial = CDate(mskFechaInicial)
                !dtmFechaFinal = CDate(mskFechaFinal)
                !vchNombreDepartamento = txtNombreDepartamento
                !intPerfil = STR(vlintPerfilFinal)
                .Update
              Else
                !vchUsuario = Trim(txtUsuario.Text)
                If txtContraseña.Text <> "" Then
                  !vchPassword = fstrEncrypt(txtContraseña.Text, txtUsuario.Text)
                End If
                !intCveEmpleado = cboEmpleado.ItemData(cboEmpleado.ListIndex)
                !smicvedepartamento = cboDepartamento.ItemData(cboDepartamento.ListIndex)
                !intCveEmpAutoriza = vglngNumeroEmpleado
                !dtmFechaInicial = CDate(mskFechaInicial)
                !dtmFechaFinal = CDate(mskFechaFinal)
                !vchNombreDepartamento = txtNombreDepartamento
                !intPerfil = STR(vlintPerfilFinal)
                rsLogin.Update
              End If
            End With
            
            If Not vlblnConsulta Then
              vllngNumeroLogin = flngObtieneIdentity("SEC_LOGIN", rsLogin!intNumeroLogin)
            Else
              vllngNumeroLogin = rsLogin!intNumeroLogin
            End If
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "USUARIO", STR(vllngNumeroLogin))
            
            'Grabar nuevos permisos
            For X = 1 To grdPermisos.Rows - 1
                With rsPermiso
                    If fblnExistePermiso(CStr(vllngNumeroLogin), grdPermisos.RowData(X)) Then
                        strSentencia = _
                        "update Permiso set " & _
                            "chrPermiso = '" & fstrPermiso(X) & "' " & _
                        "Where " & _
                            "intNumeroLogin = " & STR(vllngNumeroLogin) & _
                            " and intNumeroOpcion = " & STR(grdPermisos.RowData(X))
                        pEjecutaSentencia strSentencia
                            
                    Else
                        .AddNew
                        !intNumeroLogin = vllngNumeroLogin
                        !intNumeroOpcion = grdPermisos.RowData(X)
                        !chrpermiso = fstrPermiso(X)
                        .Update
                    End If
                End With
            Next X
            
            rsLogin.Requery
            rsPermiso.Requery
            
            txtClaveLogin.SetFocus
            pLimpia
            pHabilita 1, 1, 1, 1, 1, 0, 0
            pCargarYPosicionarPeril "M", cboModulo.ItemData(cboModulo.ListIndex)
        End If
    Else
        pLimpia
        pHabilita 1, 1, 1, 1, 1, 0, 0
    End If
    
    pCargarYPosicionarPeril "M", cboModulo.ItemData(cboModulo.ListIndex)
    
    If lintBan = 1 Then
        rsPermiso.Requery
        rsLogin.Requery
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))

End Sub

Private Sub cmdSavePerfil_Click()
    On Error GoTo NotificaError
    
    Dim vllngNumeroLogin As Long
    Dim vlstrPerfil As String
    Dim X As Integer

    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Then
        If DatosValidos() Then
            If Not vlblnConsulta Then
                With rsPerfiles
                    .AddNew
                    !vchDescripcionPerfil = txtDescripcionPerfil.Text
                    !intNumeroModulo = cboModuloPerfiles.ItemData(cboModuloPerfiles.ListIndex)
                    !vchNombreModulo = cboModuloPerfiles.Text
                    .Update
                End With
                vlstrPerfil = flngObtieneIdentity("SEC_SIPERFILES", rsPerfiles!intPerfil)
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "PERFIL", vlstrPerfil)
            Else
                vlstrPerfil = txtClavePerfil.Text
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "PERFIL", vlstrPerfil)
                vlstrsql = "update SiPerfiles set vchDescripcionPerfil='" + txtDescripcionPerfil.Text + "',"
                vlstrsql = vlstrsql + "intNumeroModulo=" + STR(cboModuloPerfiles.ItemData(cboModuloPerfiles.ListIndex)) + ","
                vlstrsql = vlstrsql + "vchNombreModulo='" + cboModuloPerfiles.Text + "' "
                vlstrsql = vlstrsql + "where intPerfil=" + txtClavePerfil.Text
                pEjecutaSentencia vlstrsql
            End If

            'Borrar los permisos anteriores
            For X = 1 To grdPermisos1.Rows - 1
                vlstrsql = "Delete from  SiPerfilesDetalle where intPerfil=" + vlstrPerfil + " and intNumeroOpcion=" + STR(grdPermisos1.RowData(X))
                pEjecutaSentencia vlstrsql
            Next X
            
            '| nuevos permisos
            For X = 1 To grdPermisos1.Rows - 1
                With rsPerfilesDetalle
                    .AddNew
                    !intPerfil = vlstrPerfil
                    !intNumeroOpcion = grdPermisos1.RowData(X)
                    !chrpermiso = fstrPermiso(X)
                    .Update
                End With
            Next X
            
            rsPerfiles.Requery
            rsPerfilesDetalle.Requery
            
            txtDescripcionPerfil.SetFocus
            pCargarYPosicionarPeril "M", cboModulo.ItemData(cboModulo.ListIndex)

        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSavePerfil_Click"))

End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    If sstObj.Tab = 0 Then
       If rsLogin.RecordCount <> 0 Then
           frmPermisosBusqueda.Show vbModal
           vlblnNoCombo = True
           If frmPermisosBusqueda.vllngLoginSeleccionado <> 0 Then
               pBusca STR(frmPermisosBusqueda.vllngLoginSeleccionado)
               pHabilita 1, 1, 1, 1, 1, 0, 1
           Else
               txtClaveLogin.SetFocus
           End If
       Else
           '¡No existe información!
           MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    Else
       If rsPerfiles.RecordCount <> 0 Then
           frmPermisosBusquedaPerfil.Show vbModal
           If frmPermisosBusquedaPerfil.vllngPerfilSeleccionado <> 0 Then
               pBusca STR(frmPermisosBusquedaPerfil.vllngPerfilSeleccionado)
               pHabilita 1, 1, 1, 1, 1, 0, 1
           Else
               txtClavePerfil.SetFocus
           End If
       Else
           '¡No existe información!
           MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
    
    If sstObj.Tab = 0 Then
       If rsLogin.RecordCount <> 0 Then
           rsLogin.MoveLast
           pBusca STR(rsLogin!intNumeroLogin)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
          MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    Else
       If rsPerfiles.RecordCount <> 0 Then
           rsPerfiles.MoveLast
           pBusca STR(rsPerfiles!intPerfil)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
          MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError
    
    If sstObj.Tab = 0 Then
       If rsLogin.RecordCount <> 0 Then
           rsLogin.MoveNext
           If rsLogin.EOF Then
               rsLogin.MovePrevious
           End If
           pBusca STR(rsLogin!intNumeroLogin)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
           MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    Else
       If rsPerfiles.RecordCount <> 0 Then
           rsPerfiles.MoveNext
           If rsPerfiles.EOF Then
               rsPerfiles.MovePrevious
           End If
           pBusca STR(rsPerfiles!intPerfil)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
           MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdBack_Click()
    On Error GoTo NotificaError
    
    If sstObj.Tab = 0 Then
       If rsLogin.RecordCount <> 0 Then
           rsLogin.MovePrevious
           If rsLogin.BOF Then
               rsLogin.MoveNext
           End If
           pBusca STR(rsLogin!intNumeroLogin)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
           MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    Else
       If rsPerfiles.RecordCount <> 0 Then
           rsPerfiles.MovePrevious
           If rsPerfiles.BOF Then
               rsPerfiles.MoveFirst
           End If
           pBusca STR(rsPerfiles!intPerfil)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
           MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
    
    If sstObj.Tab = 0 Then
       If rsLogin.RecordCount <> 0 Then
           rsLogin.MoveFirst
           pBusca STR(rsLogin!intNumeroLogin)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
          MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    Else
       If rsPerfiles.RecordCount <> 0 Then
           rsPerfiles.MoveFirst
           pPosicionaRegRs rsPerfiles, "I"
           pBusca STR(rsPerfiles!intPerfil)
           pHabilita 1, 1, 1, 1, 1, 0, 1
       Else
           '¡No existe información!
          MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
       End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Function fstrPermiso(vlintRenglon As Integer) As String
    On Error GoTo NotificaError
    
    If sstObj.Tab = 0 Then
       fstrPermiso = IIf(grdPermisos.TextMatrix(vlintRenglon, 3) = "x", "L", IIf(grdPermisos.TextMatrix(vlintRenglon, 4) = "x", "E", IIf(grdPermisos.TextMatrix(vlintRenglon, 5) = "x", "C", "S")))
    Else
       fstrPermiso = IIf(grdPermisos1.TextMatrix(vlintRenglon, 3) = "x", "L", IIf(grdPermisos1.TextMatrix(vlintRenglon, 4) = "x", "E", IIf(grdPermisos1.TextMatrix(vlintRenglon, 5) = "x", "C", "S")))
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fstrPermiso"))

End Function

Private Function DatosValidos() As Boolean
    On Error GoTo NotificaError
    
    Dim rsUsuariosRepetidos As New ADODB.Recordset
    
    DatosValidos = True
    
    If Trim(txtUsuario.Text) = "" And sstObj.Tab = 0 Then
        DatosValidos = False
        txtUsuario.SetFocus
    End If

    If Not vlblnConsulta And txtContraseña.Text = "" And sstObj.Tab = 0 Then
            DatosValidos = False
            txtContraseña.SetFocus
    End If
    If DatosValidos And Trim(txtNombreDepartamento.Text) = "" And sstObj.Tab = 0 Then
        DatosValidos = False
        txtNombreDepartamento.SetFocus
    End If
    If DatosValidos And Trim(txtDescripcionPerfil.Text) = "" And sstObj.Tab = 1 Then
        DatosValidos = False
        txtDescripcionPerfil.SetFocus
    End If
    If Not DatosValidos Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
    Else
        If cboEmpleado.ListIndex < 0 And sstObj.Tab = 0 Then
            DatosValidos = False
            cboEmpleado.SetFocus
            ' Seleccione el empleado
            MsgBox SIHOMsg(241), vbOKOnly + vbInformation, "Mensaje"
        End If
        If DatosValidos And cboDepartamento.ListIndex < 0 And sstObj.Tab = 0 Then
            DatosValidos = False
            cboDepartamento.SetFocus
            ' Seleccione el departamento
            MsgBox SIHOMsg(242), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        If DatosValidos And sstObj.Tab = 0 And Len(txtContraseña.Text) > 0 Then
            If Mid(txtContraseña.Text, 1, 1) = Mid(txtUsuario.Text, 1, 1) Then
                DatosValidos = False
                txtContraseña.SetFocus
                ' ¡La contraseña no puede iniciar con ese caracter!
                MsgBox SIHOMsg(774) & Chr(13), vbExclamation, "Mensaje"
            End If
        End If
        
        If DatosValidos And txtContraseña.Text <> txtVerifContraseña.Text And sstObj.Tab = 0 Then
            DatosValidos = False
            txtContraseña.SetFocus
            MsgBox SIHOMsg(763) & Chr(13), vbExclamation, "Mensaje"
        End If
        
        If DatosValidos And Not fblnValidaFecha(mskFechaInicial) And sstObj.Tab = 0 Then
            DatosValidos = False
            mskFechaInicial.SetFocus
            ' ¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        End If
        If DatosValidos And Not fblnValidaFecha(mskFechaFinal) And sstObj.Tab = 0 Then
            DatosValidos = False
            mskFechaFinal.SetFocus
            ' ¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        End If
        If DatosValidos And (CDate(mskFechaInicial.Text) > CDate(mskFechaFinal.Text)) And sstObj.Tab = 0 Then
            DatosValidos = False
            mskFechaInicial.SetFocus
            ' ¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
        End If
        If DatosValidos And sstObj.Tab = 0 Then
            If vlblnConsulta Then
                vlstrsql = "select count(*) from Login where intNumeroLogin<>" + txtClaveLogin.Text + " and vchUsuario=" + "'" + txtUsuario.Text + "'"
            Else
                vlstrsql = "select count(*) from Login where vchUsuario=" + "'" + txtUsuario.Text + "'"
            End If
            Set rsUsuariosRepetidos = frsRegresaRs(vlstrsql)
            If rsUsuariosRepetidos.Fields(0) <> 0 Then
                DatosValidos = False
                txtUsuario.SetFocus
                ' ¡Existe otro usuario con el mismo nombre!
                MsgBox SIHOMsg(245), vbOKOnly + vbInformation, "Mensaje"
            End If
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":DatosValidos"))

End Function

Private Sub Form_Activate()
    On Error GoTo NotificaError
    

    
    Dim X As Integer
   
    If rsEmpleados.RecordCount = 0 Then
        ' No existen empleados registrados.
        MsgBox SIHOMsg(238), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If
    If rsDepartamentos.RecordCount = 0 Then
        ' No existen departamentos registrados.
        MsgBox SIHOMsg(239), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If
    If cboModulo.ListCount = 0 Then
        ' No existen opciones registradas para este módulo.
        MsgBox SIHOMsg(240), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    Else
        'Posicionar el cbo de módulos
         For X = 0 To cboModulo.ListCount - 1
             If cboModulo.ItemData(X) = vgintNumeroModulo And Not vlblnNoCombo Then
                 cboModulo.ListIndex = X
                 cboModuloPerfiles.ListIndex = X
                 pLimpiaGrid
                 Exit For
             End If
         Next X
    End If

    vlblnNoCombo = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If cmdSave.Enabled Then
            ' ¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
               If sstObj.Tab = 1 Then
                  txtClavePerfil.SetFocus
               Else
                  txtClaveLogin.SetFocus
               End If
            End If
        Else
            Unload Me
        End If
    Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))

End Sub

Private Sub pHabilita(a As Integer, B As Integer, c As Integer, d As Integer, e As Integer, f As Integer, G As Integer)
    On Error GoTo NotificaError
    
    If a = 1 Then
        cmdTop.Enabled = True
    Else
        cmdTop.Enabled = False
    End If
    If B = 1 Then
        cmdBack.Enabled = True
    Else
        cmdBack.Enabled = False
    End If
    If c = 1 Then
        cmdLocate.Enabled = True
    Else
        cmdLocate.Enabled = False
    End If
    If d = 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
    If e = 1 Then
        cmdEnd.Enabled = True
    Else
        cmdEnd.Enabled = False
    End If
    If f = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    If G = 1 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If

    If f = 1 And cmdSave.Enabled = True Then
       If sstObj.Tab = 1 Then
          sstObj.TabEnabled(0) = False
       Else
          sstObj.TabEnabled(1) = False
       End If
    Else
       If sstObj.Tab = 1 Or vlblnActivaTabPerfiles Then
          sstObj.TabEnabled(0) = True
       Else
          sstObj.TabEnabled(1) = True
       End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))

End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    vlblnConsulta = False
    
    If sstObj.Tab = 0 And Not vlblnActivaTabPerfiles Then
       txtUsuario.Text = ""
       txtContraseña.Text = ""
       txtVerifContraseña.Text = ""
       txtNombreDepartamento.Text = ""
       mskFechaInicial.Text = fdtmServerFecha
       mskFechaFinal.Text = fdtmServerFecha
       cboDepartamento.ListIndex = -1
       cboEmpleado.ListIndex = -1
       If vlblnNoEsPrimerEntrada = False Then
            vlblnNoEsPrimerEntrada = True
       Else
            pCargarYPosicionarPeril "M", cboModulo.ItemData(cboModulo.ListIndex)
       End If
       pPersonaAutoriza vglngNumeroEmpleado
       pUltimaClaveLogin
       pLimpiaGrid
    Else
       txtDescripcionPerfil.Text = ""
       pUltimaClavePerfil
       pLimpiaGridPerfiles
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))

End Sub
Private Sub pPersonaAutoriza(vllngxNumeroEmpleado As Long)
    On Error GoTo NotificaError
    
    Dim rsPersonaAutoriza As New ADODB.Recordset
    
    txtPersonaAutoriza.Text = ""
    vlstrsql = "select vchApellidoPaterno||' '||vchApellidoMaterno||' '||vchNombre as Nombre from Noempleado where intCveEmpleado=" + STR(vllngxNumeroEmpleado)
    Set rsPersonaAutoriza = frsRegresaRs(vlstrsql)
    If rsPersonaAutoriza.RecordCount <> 0 Then
        txtPersonaAutoriza.Text = rsPersonaAutoriza!Nombre
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPersonaAutoriza"))

End Sub

Private Sub pLimpiaGrid()
    On Error GoTo NotificaError
    
    Dim X As Integer
    
    With grdPermisos
       For X = 1 To .Rows - 1
            .TextMatrix(X, 3) = ""
            .TextMatrix(X, 4) = ""
            .TextMatrix(X, 5) = ""
            .TextMatrix(X, 6) = "x"
       Next X
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGrid"))

End Sub

Private Sub pLimpiaGridPerfiles()
    On Error GoTo NotificaError
    
    Dim X As Integer
    
    With grdPermisos1
        For X = 1 To .Rows - 1
            .TextMatrix(X, 3) = ""
            .TextMatrix(X, 4) = ""
            .TextMatrix(X, 5) = ""
            .TextMatrix(X, 6) = "x"
        Next X
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridPerfiles"))

End Sub

Private Sub pUltimaClaveLogin()
    On Error GoTo NotificaError
    
    Dim rsUltimaClave As New ADODB.Recordset
    
    If rsLogin.RecordCount = 0 Then
        txtClaveLogin.Text = 1
    Else
        vlstrsql = "select max(intNumeroLogin)+1 as Ultimo from Login"
        Set rsUltimaClave = frsRegresaRs(vlstrsql)
        txtClaveLogin.Text = rsUltimaClave!Ultimo
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pUltimaClaveLogin"))

End Sub

Private Sub pUltimaClavePerfil()
    On Error GoTo NotificaError
    
    Dim rsUltimaClave As New ADODB.Recordset
    
    If rsPerfiles.RecordCount = 0 Then
        txtClavePerfil.Text = 1
    Else
        vlstrsql = "select max(intPerfil)+1 as Ultimo from SiPerfiles"
        Set rsUltimaClave = frsRegresaRs(vlstrsql)
        txtClavePerfil.Text = rsUltimaClave!Ultimo
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pUltimaClavePerfil"))

End Sub

Private Sub pConfigura()
    On Error GoTo NotificaError
    
    With grdPermisos
        .Cols = 6
        .FormatString = "|Opción||L|E|C|S"
        .ColWidth(0) = 0
        .ColWidth(1) = 11020 'Opción
        .ColWidth(2) = 0    'Numero opcion
        .ColWidth(3) = 220  'Lectura
        .ColWidth(4) = 220  'Escritura
        .ColWidth(5) = 220  'Control total
        .ColWidth(6) = 220  'Sin acceso
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfigura"))

End Sub

Private Sub pConfiguraModuloPerfiles()
    On Error GoTo NotificaError
    
    With grdPermisos1
        .Cols = 6
        .FormatString = "|Opción||L|E|C|S"
        .ColWidth(0) = 0
        .ColWidth(1) = 11020 'Opción
        .ColWidth(2) = 0    'Numero opcion
        .ColWidth(3) = 220  'Lectura
        .ColWidth(4) = 220  'Escritura
        .ColWidth(5) = 220  'Control total
        .ColWidth(6) = 220  'Sin acceso
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraModuloPerfiles"))

End Sub
Private Sub grdPermisos_Click()
    On Error GoTo NotificaError
    
    With grdPermisos
        If .Row > 0 Then
            If .Col > 2 Then
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 4) = ""
                .TextMatrix(.Row, 5) = ""
                .TextMatrix(.Row, 6) = ""
                If .Col = 3 Then
                    .TextMatrix(.Row, 3) = "x"
                    vlblnHaCambiado = True
                End If
                If .Col = 4 Then
                    .TextMatrix(.Row, 4) = "x"
                    vlblnHaCambiado = True
                End If
                If .Col = 5 Then
                    .TextMatrix(.Row, 5) = "x"
                    vlblnHaCambiado = True
                End If
                If .Col = 6 Then
                    .TextMatrix(.Row, 6) = "x"
                    vlblnHaCambiado = True
                End If
            End If
        End If
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPermisos_Click"))
End Sub

Private Sub grdPermisos1_Click()
    On Error GoTo NotificaError
    
    With grdPermisos1
        If .Row > 0 Then
            If .Col > 2 Then
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 4) = ""
                .TextMatrix(.Row, 5) = ""
                .TextMatrix(.Row, 6) = ""
                If .Col = 3 Then
                    .TextMatrix(.Row, 3) = "x"
                End If
                If .Col = 4 Then
                    .TextMatrix(.Row, 4) = "x"
                End If
                If .Col = 5 Then
                    .TextMatrix(.Row, 5) = "x"
                End If
                If .Col = 6 Then
                    .TextMatrix(.Row, 6) = "x"
                End If
            End If
        End If
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPermisos1_Click"))
End Sub

Private Sub pCargarYPosicionarPeril(vlstrTipo As String, vlintPerfil As Integer)
    On Error GoTo NotificaError

    Dim rsTmpPerfiles As New ADODB.Recordset

    vlstrsql = "select * from SiPerfiles where intNumeroModulo=" + STR(cboModulo.ItemData(cboModulo.ListIndex))
    
    Set rsTmpPerfiles = frsRegresaRs(vlstrsql)

    If rsTmpPerfiles.RecordCount <> 0 Then
       pLlenarCboRs_new cboPerfiles, rsTmpPerfiles, 0, 1
       cboPerfiles.AddItem " ", 0
       cboPerfiles.ItemData(0) = 0
       If vlstrTipo = "M" Then
          pPosiciona cboPerfiles, 0
          pCargarPermisosSinPerfil
       Else
          pPosiciona cboPerfiles, STR(vlintPerfil)
       End If
    Else
       pLlenarCboRs_new cboPerfiles, rsTmpPerfiles, 0, 1
       cboPerfiles.AddItem " ", 0
       cboPerfiles.ItemData(0) = 0
       pPosiciona cboPerfiles, 0
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargarYPosicionarPeril"))
End Sub

Private Sub pBuscaryPosicionaModulo()
    On Error GoTo NotificaError
    Dim rsModulo As New ADODB.Recordset

    vlblnNoCombo = True
       
    vlstrsql = "SELECT ISNULL(MAX(smiNumeroModulo),0) AS Modulo FROM permiso " & _
               "JOIN Opcion ON Opcion.intNumeroOpcion = Permiso.intNumeroOpcion " & _
               "WHERE intNumeroLogin=" & txtClaveLogin.Text
    Set rsModulo = frsRegresaRs(vlstrsql)

    pPosiciona cboModulo, rsModulo!modulo

    pCargarYPosicionarPeril "M", cboModulo.ItemData(cboModulo.ListIndex)

    rsModulo.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBuscaryPosicionaModulo"))
End Sub

Private Sub pBusca(vlstrNumeroLogin As String)
    On Error GoTo NotificaError

    Dim X As Integer
    Dim rsxLogin As New ADODB.Recordset
    Dim rsxPerfil As New ADODB.Recordset
    Dim rsHayPerfilGuardado As New ADODB.Recordset
    Dim vlstrsql As String
    
    If sstObj.Tab = 0 And Not vlblnActivaTabPerfiles Then
        X = STR(fintLocalizaPkRs(rsLogin, 0, vlstrNumeroLogin))
       vlstrsql = "select * from Login where intNumeroLogin=" + vlstrNumeroLogin
       Set rsxLogin = frsRegresaRs(vlstrsql)
       If rsxLogin.RecordCount <> 0 Then
          vlblnConsulta = True
          vlblnHaCambiado = False
          txtClaveLogin.Text = rsxLogin!intNumeroLogin
          txtUsuario.Text = rsxLogin!vchUsuario
          txtContraseña.Text = ""
          txtNombreDepartamento.Text = rsxLogin!vchNombreDepartamento
          vlintPerfilUsuario = rsxLogin!intPerfil
          mskFechaInicial.Text = rsxLogin!dtmFechaInicial
          mskFechaFinal.Text = rsxLogin!dtmFechaFinal
          
           If IIf(IsNull(rsxLogin!intCveEmpleado), 0, rsxLogin!intCveEmpleado) <> 0 Then
             pPosiciona cboEmpleado, rsxLogin!intCveEmpleado ' esta estaba en lugar del IF
           Else
              cboEmpleado.ListIndex = -1
           End If
        

          pPosiciona cboDepartamento, rsxLogin!smicvedepartamento
          pPersonaAutoriza rsxLogin!intCveEmpAutoriza
'         pBuscaryPosicionaModulo
          pLimpiaGrid
        'Vamos a revisar si hay historico del perfil o no
        vlstrsql = "select * from SIPERFILESGUARDADOS where INTCVEMODULO = " + CStr(cboModulo.ItemData(cboModulo.ListIndex)) + " AND INTCVEUSUARIO = " + txtClaveLogin.Text
        Set rsHayPerfilGuardado = frsRegresaRs(vlstrsql)
        If rsHayPerfilGuardado.RecordCount > 0 Then
          If rsHayPerfilGuardado!INTCVEPERFIL <> 0 Then
            pCargarYPosicionarPeril "P", rsHayPerfilGuardado!INTCVEPERFIL
          Else
            pPosiciona cboPerfiles, 0
            pCargarPermisosSinPerfil
          End If
        Else
          If rsxLogin!intPerfil <> 0 Then
            pCargarYPosicionarPeril "P", rsxLogin!intPerfil
          Else
            pPosiciona cboPerfiles, 0
            pCargarPermisosSinPerfil
          End If
         End If
       Else
           pLimpia
       End If
    Else
       X = STR(fintLocalizaPkRs(rsPerfiles, 0, vlstrNumeroLogin))
       vlstrsql = "select * from SiPerfiles where intPerfil=" + Trim(vlstrNumeroLogin)
       Set rsxPerfil = frsRegresaRs(vlstrsql)
       If rsxPerfil.RecordCount <> 0 Then
          vlblnConsulta = True
          vlblnNoCombo = True
          txtClavePerfil.Text = rsxPerfil!intPerfil
          txtDescripcionPerfil.Text = rsxPerfil!vchDescripcionPerfil
          pPosiciona cboModuloPerfiles, rsxPerfil!intNumeroModulo
          pLimpiaGridPerfiles
          pPonPermiso1 vlstrNumeroLogin
       Else
          pLimpia
       End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBusca"))
End Sub

Private Sub pBuscaPerfil(vlstrPerfil As String)
    On Error GoTo NotificaError

    Dim X As Integer
    Dim rsxPerfilesDetalle As New ADODB.Recordset
    
    vlstrsql = "select * from SiPerfilesDetalle where intPerfil=" + vlstrPerfil
    Set rsxPerfilesDetalle = frsRegresaRs(vlstrsql)
    If rsxPerfilesDetalle.RecordCount <> 0 Then
       pLimpiaGrid
       pPonPermisoPerfil rsxPerfilesDetalle!intPerfil
    Else
       pLimpiaGrid
       pCargarPermisosSinPerfil
       pPonPermiso txtClaveLogin.Text
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBuscaPerfil"))
End Sub

Private Sub pCargarPermisosSinPerfil()
    On Error GoTo NotificaError
    '--Verificacion de socios para mostrar los permisos de socios.
    Dim rs As ADODB.Recordset
    Dim vlintsocios As Integer
    Set rs = frsRegresaRs("SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITUTILIZASOCIOS'")
    If Not rs.EOF Then
        vlintsocios = rs!vchvalor
    Else
        vlintsocios = 0
    End If
    rs.Close
    If Not vlintsocios = 1 Then
        vlstrsql = "SELECT vchDescripcion,intNumeroOpcion FROM Opcion WHERE smiNumeroModulo=" + STR(cboModulo.ItemData(cboModulo.ListIndex)) + " AND INTORDENOPCION NOT LIKE '02.06%'ORDER BY intOrdenOpcion, vchDescripcion"
    Else
        vlstrsql = "SELECT vchDescripcion,intNumeroOpcion FROM Opcion WHERE smiNumeroModulo=" + STR(cboModulo.ItemData(cboModulo.ListIndex)) + " ORDER BY intOrdenOpcion, vchDescripcion"
    End If
    
   'vlstrSQL = "SELECT vchDescripcion,intNumeroOpcion FROM Opcion WHERE smiNumeroModulo=" + Str(cboModulo.ItemData(cboModulo.ListIndex)) + " ORDER BY intOrdenOpcion, vchDescripcion"
   Set rsOpciones = frsRegresaRs(vlstrsql)
   If rsOpciones.RecordCount <> 0 Then
      pLlenarMshFGrdRs grdPermisos, rsOpciones, 1
      pConfigura
      pPonPermiso txtClaveLogin.Text
    Else
      pLimpiaGrid
    End If
      
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargarPermisosSinPerfil"))
End Sub

Private Sub pPonPermiso(vlstrLogin As String)
    On Error GoTo NotificaError
    
    Dim rsPermiso As New ADODB.Recordset
    Dim vlintRenglon As Integer
    
    vlstrsql = "select chrPermiso, Permiso.intNumeroOpcion from Permiso JOIN Opcion ON Opcion.intNumeroOpcion = Permiso.intNumeroOpcion where intNumeroLogin=" + vlstrLogin + " and Opcion.smiNumeroModulo=" & STR(cboModulo.ItemData(cboModulo.ListIndex)) & " order by intOrdenOpcion"
    Set rsPermiso = frsRegresaRs(vlstrsql)
    vlintRenglon = 1

    If rsPermiso.EOF Then
        pLimpiaGrid
        Exit Sub
    End If

    Do While Not rsPermiso.EOF
        For vlintRenglon = 1 To grdPermisos.Rows - 1
          If rsPermiso!intNumeroOpcion = grdPermisos.TextMatrix(vlintRenglon, 2) Then
    
            grdPermisos.TextMatrix(vlintRenglon, 3) = ""
            grdPermisos.TextMatrix(vlintRenglon, 4) = ""
            grdPermisos.TextMatrix(vlintRenglon, 5) = ""
            grdPermisos.TextMatrix(vlintRenglon, 6) = ""
            If rsPermiso!chrpermiso = "L" Then
                grdPermisos.TextMatrix(vlintRenglon, 3) = "x"
            End If
            If rsPermiso!chrpermiso = "E" Then
                grdPermisos.TextMatrix(vlintRenglon, 4) = "x"
            End If
            If rsPermiso!chrpermiso = "C" Then
                grdPermisos.TextMatrix(vlintRenglon, 5) = "x"
            End If
            If rsPermiso!chrpermiso = "S" Then
                grdPermisos.TextMatrix(vlintRenglon, 6) = "x"
            End If
          End If
        Next vlintRenglon
        rsPermiso.MoveNext
    Loop
     
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPonPermiso"))
End Sub

Private Sub pPonPermisoPerfil(vlstrPerfil As String)
    On Error GoTo NotificaError
    
    Dim rsPermiso As New ADODB.Recordset
    Dim vlintRenglon As Integer
    
    vlstrsql = "select chrPermiso, SiPerfilesDetalle.intNumeroOpcion from SiPerfilesDetalle JOIN Opcion ON opcion.intNumeroOpcion = siPerfilesDetalle.intNumeroOpcion where intPerfil=" + vlstrPerfil + " order by intOrdenOpcion"
    Set rsPermiso = frsRegresaRs(vlstrsql)
    vlintRenglon = 1
    
    Do While Not rsPermiso.EOF
        For vlintRenglon = 1 To grdPermisos.Rows - 1
          If rsPermiso!intNumeroOpcion = grdPermisos.TextMatrix(vlintRenglon, 2) Then
            grdPermisos.TextMatrix(vlintRenglon, 3) = ""
            grdPermisos.TextMatrix(vlintRenglon, 4) = ""
            grdPermisos.TextMatrix(vlintRenglon, 5) = ""
            grdPermisos.TextMatrix(vlintRenglon, 6) = ""
            If rsPermiso!chrpermiso = "L" Then
                grdPermisos.TextMatrix(vlintRenglon, 3) = "x"
            End If
            If rsPermiso!chrpermiso = "E" Then
                grdPermisos.TextMatrix(vlintRenglon, 4) = "x"
            End If
            If rsPermiso!chrpermiso = "C" Then
                grdPermisos.TextMatrix(vlintRenglon, 5) = "x"
            End If
            If rsPermiso!chrpermiso = "S" Then
                grdPermisos.TextMatrix(vlintRenglon, 6) = "x"
            End If
          End If
        Next vlintRenglon
        rsPermiso.MoveNext
    Loop
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPonPermisoPerfil"))
End Sub

Private Sub pPonPermiso1(vlstrPerfil As String)
    On Error GoTo NotificaError
    
    Dim rsPerfil As New ADODB.Recordset
    Dim vlintRenglon As Integer
    
    'vlstrSQL = "select chrPermiso, SiPerfilesDetalle.intNumeroOpcion from SiPerfilesDetalle JOIN Opcion ON opcion.intNumeroOpcion = siPerfilesDetalle.intNumeroOpcion where intPerfil=" + vlstrPerfil + " order by intOrdenOpcion"
    'Set rsPerfil = frsRegresaRs(vlstrSQL)
    '---------------------------------------------------
    vgstrParametrosSP = vlstrPerfil
    Set rsPerfil = frsEjecuta_SP(vgstrParametrosSP, "SP_SiSelPerfilesDetalle")
    '---------------------------------------------------
    
    vlintRenglon = 1
    
    Do While Not rsPerfil.EOF
      For vlintRenglon = 1 To grdPermisos1.Rows - 1
        If rsPerfil!intNumeroOpcion = grdPermisos1.TextMatrix(vlintRenglon, 2) Then
          grdPermisos1.TextMatrix(vlintRenglon, 3) = ""
          grdPermisos1.TextMatrix(vlintRenglon, 4) = ""
          grdPermisos1.TextMatrix(vlintRenglon, 5) = ""
          grdPermisos1.TextMatrix(vlintRenglon, 6) = ""
          If rsPerfil!chrpermiso = "L" Then
              grdPermisos1.TextMatrix(vlintRenglon, 3) = "x"
          End If
          If rsPerfil!chrpermiso = "E" Then
              grdPermisos1.TextMatrix(vlintRenglon, 4) = "x"
          End If
          If rsPerfil!chrpermiso = "C" Then
              grdPermisos1.TextMatrix(vlintRenglon, 5) = "x"
          End If
          If rsPerfil!chrpermiso = "S" Then
              grdPermisos1.TextMatrix(vlintRenglon, 6) = "x"
          End If
        End If
      Next vlintRenglon
      rsPerfil.MoveNext
    Loop
     
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPonPermiso1"))

End Sub

Private Sub pPosiciona(cboNombre As MyCombo, vlintNumero As Long)
    On Error GoTo NotificaError
    
    Dim X As Integer
    
    For X = 0 To cboNombre.ListCount - 1
        If cboNombre.ItemData(X) = vlintNumero Then
            cboNombre.ListIndex = X
            Exit For
        End If
    Next X

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPosiciona"))

End Sub

Private Sub Form_Load()
   On Error GoTo NotificaError
   Dim rs As New ADODB.Recordset

    vlblnNoEsPrimerEntrada = False

    'Color del SSTab
    SetStyle sstObj.hwnd, 0
    SetSolidColor sstObj.hwnd, 16777215
    SSTabSubclass sstObj.hwnd
   
   
    Me.Icon = frmMenuPrincipal.Icon
   
    fblnHabilitaObjetos frmPermisos
    sstObj.Tab = 0
    vlblnNoCombo = False
   
    ' Empleados
    vlstrsql = "select vchApellidoPaterno||' '||vchApellidoMaterno||' '||vchNombre as Nombre,intCveEmpleado from NoEmpleado where bitActivo=1"
    Set rsEmpleados = frsRegresaRs(vlstrsql)
    If rsEmpleados.RecordCount <> 0 Then
        pLlenarCboRs_new cboEmpleado, rsEmpleados, 1, 0
    End If
   
    ' Departamentos
    vlstrsql = "select vchDescripcion,smiCveDepartamento from NoDepartamento WHERE BITESTATUS = 1 and tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rsDepartamentos = frsRegresaRs(vlstrsql)
    If rsDepartamentos.RecordCount <> 0 Then
        pLlenarCboRs_new cboDepartamento, rsDepartamentos, 1, 0
    End If
   
    ' Persona autoriza
    pPersonaAutoriza vglngNumeroEmpleado
   
    '------------------------------
    ' Tablas
    '------------------------------
    vlstrsql = "SELECT * FROM SiPerfiles "
    Set rsPerfiles = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    vlstrsql = "select * from Login Order by intNumerologin"
    
    Set rsLogin = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    vlstrsql = "select * from Permiso"
    Set rsPermiso = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    vlstrsql = "select * from SiPerfilesDetalle Order by intPerfil, intNumeroOpcion"
    Set rsPerfilesDetalle = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
   
    ' Modulos
    vlstrsql = "select * from modulo"
    Set rs = frsRegresaRs(vlstrsql)
   
    pLlenarCboRs_new cboModulo, rs, 0, 1
    If cgstrModulo <> "SI" Then cboModulo.Enabled = False
   
    pLlenarCboRs_new cboModuloPerfiles, rs, 0, 1
    rs.Close
   
    pLimpia
    pHabilita 1, 1, 1, 1, 1, 0, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Function fblnExistePermiso(vlstrLogin As String, vlstrCvePermiso As String) As Boolean

Dim vlrsPermiso As New ADODB.Recordset

    fblnExistePermiso = False
    Set vlrsPermiso = frsRegresaRs("SELECT intNumeroOpcion " & _
                                    "FROM Permiso " & _
                                    "WHERE intNumeroLogin = " & vlstrLogin & " AND " & _
                                    "      intNumeroOpcion = " & vlstrCvePermiso, adLockOptimistic, adOpenDynamic)
    If vlrsPermiso.RecordCount Then fblnExistePermiso = True
End Function
Private Sub txtVerifContraseña_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cboEmpleado.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtVerifContraseña_KeyPress"))

End Sub


