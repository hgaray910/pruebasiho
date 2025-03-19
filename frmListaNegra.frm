VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmListaNegra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de deudores incobrables"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAutoriza 
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
      Height          =   1090
      Left            =   7785
      TabIndex        =   40
      Top             =   8480
      Visible         =   0   'False
      Width           =   3060
      Begin MyCommandButton.MyButton cmdNoAutoriza 
         Height          =   375
         Left            =   90
         TabIndex        =   44
         Top             =   650
         Width           =   1455
         _ExtentX        =   2566
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
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   15790320
         Caption         =   "Cancelar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin VB.TextBox txtPassword 
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
         Left            =   1530
         PasswordChar    =   "*"
         TabIndex        =   42
         Top             =   240
         Width           =   1450
      End
      Begin MyCommandButton.MyButton cmdSiAutoriza 
         Height          =   375
         Left            =   1530
         TabIndex        =   43
         Top             =   650
         Width           =   1455
         _ExtentX        =   2566
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
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   15790320
         Caption         =   "Aceptar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   3060
         X2              =   3060
         Y1              =   0
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   3120
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Left            =   90
         TabIndex        =   45
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   615
      Left            =   255
      TabIndex        =   26
      Top             =   10440
      Width           =   1455
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Siguiente"
         Default         =   -1  'True
         Height          =   255
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10935
      Left            =   -270
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   -120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   19288
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmListaNegra.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInst"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraControl"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraInfo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmListaNegra.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraGrid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmListaNegra.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraGrid 
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
         Height          =   8600
         Left            =   -74640
         TabIndex        =   31
         Top             =   120
         Width           =   10755
         Begin VSFlex7LCtl.VSFlexGrid grdBusqueda 
            Height          =   8205
            Left            =   120
            TabIndex        =   32
            Top             =   270
            Width           =   10515
            _cx             =   18547
            _cy             =   14473
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   5
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   8600
         Left            =   -74640
         TabIndex        =   35
         Top             =   120
         Width           =   10755
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
            Height          =   3470
            Left            =   120
            TabIndex        =   49
            Top             =   5010
            Width           =   10515
            Begin VB.TextBox txtObservacionesFamiliar 
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
               Height          =   1935
               Left            =   2280
               Locked          =   -1  'True
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   66
               Top             =   1400
               Width           =   8175
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Paciente"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   52
               Top             =   300
               Width           =   840
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Fecha de registro"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   53
               Top             =   650
               Width           =   1695
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Empleado registró"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   54
               Top             =   1000
               Width           =   1785
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Observaciones"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   55
               Top             =   1350
               Width           =   1455
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Expediente"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   6480
               TabIndex        =   56
               Top             =   300
               Width           =   1080
            End
            Begin VB.Label lblNombrePac 
               BackColor       =   &H80000005&
               Caption         =   "nombrePac"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   2280
               TabIndex        =   57
               Top             =   300
               Width           =   3615
            End
            Begin VB.Label lblExpediente 
               BackColor       =   &H80000005&
               Caption         =   "expediente"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   7680
               TabIndex        =   58
               Top             =   300
               Width           =   1575
            End
            Begin VB.Label lblFechaRegistro 
               BackColor       =   &H80000005&
               Caption         =   "fecha"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   2280
               TabIndex        =   59
               Top             =   650
               Width           =   3375
            End
            Begin VB.Label lblEmpleadoRegistro 
               BackColor       =   &H80000005&
               Caption         =   "empleado"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   2280
               TabIndex        =   60
               Top             =   1000
               Width           =   6375
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid grdPersonas 
            Height          =   2820
            Left            =   120
            TabIndex        =   41
            Top             =   1920
            Width           =   10515
            _cx             =   18547
            _cy             =   4974
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Datos del paciente en la lista de deudores incobrables"
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
            TabIndex        =   48
            Top             =   4790
            Width           =   5295
         End
         Begin VB.Label Label10 
            BackColor       =   &H80000005&
            Caption         =   "Se encontró la siguiente información en la lista de deudores incobrables"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   10515
         End
         Begin VB.Label lblMensajeCoinciden 
            BackColor       =   &H80000005&
            Caption         =   $"frmListaNegra.frx":0054
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   10515
         End
      End
      Begin VB.Frame fraInfo 
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
         Height          =   8600
         Left            =   360
         TabIndex        =   0
         Top             =   120
         Width           =   10755
         Begin VB.Frame fraBusqueda 
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
            Height          =   2010
            Left            =   3550
            TabIndex        =   62
            Top             =   150
            Width           =   7080
            Begin VB.TextBox txtIniciales 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   64
               ToolTipText     =   "Iniciales del nombre del cliente"
               Top             =   220
               Width           =   5685
            End
            Begin VB.ListBox lstBusqueda 
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
               Height          =   1305
               Left            =   90
               TabIndex        =   63
               Top             =   630
               Width           =   6885
            End
            Begin MyCommandButton.MyButton cmdCargar 
               Height          =   375
               Left            =   5810
               TabIndex        =   65
               ToolTipText     =   "Cargar la lista"
               Top             =   225
               Width           =   1170
               _ExtentX        =   2064
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
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   15790320
               Caption         =   "Cargar"
               DepthEvent      =   1
               ShowFocus       =   -1  'True
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000015&
               X1              =   0
               X2              =   7080
               Y1              =   0
               Y2              =   0
            End
         End
         Begin VB.OptionButton optTipoCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Médico"
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
            Index           =   2
            Left            =   2040
            TabIndex        =   4
            ToolTipText     =   "Tipo de cliente"
            Top             =   1240
            Width           =   1380
         End
         Begin VB.OptionButton optTipoCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Paciente"
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
            Index           =   0
            Left            =   2040
            TabIndex        =   2
            ToolTipText     =   "Tipo de cliente"
            Top             =   700
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton optTipoCliente 
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
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   3
            ToolTipText     =   "Tipo de cliente"
            Top             =   970
            Width           =   1440
         End
         Begin VB.TextBox txtNumeroCliente 
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
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   1
            ToolTipText     =   "Número de cliente"
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optTipoCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Empresa"
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
            Index           =   3
            Left            =   2040
            TabIndex        =   5
            ToolTipText     =   "Tipo de cliente"
            Top             =   1510
            Width           =   1290
         End
         Begin VB.CheckBox chkActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Activo"
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
            Left            =   2040
            TabIndex        =   17
            ToolTipText     =   "Cliente activo"
            Top             =   8240
            Width           =   1170
         End
         Begin VB.TextBox txtObservaciones 
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
            Height          =   825
            Left            =   2040
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   7360
            Width           =   8595
         End
         Begin VB.TextBox txtNombreCliente 
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   2600
            Width           =   8595
         End
         Begin VB.TextBox txtFecha 
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   6560
            Width           =   1815
         End
         Begin VB.TextBox txtReferencia 
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2190
            Width           =   1455
         End
         Begin VB.TextBox txtEmpleado 
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   6960
            Width           =   8595
         End
         Begin MyCommandButton.MyButton cmdBuscar 
            Height          =   375
            Left            =   2040
            TabIndex        =   6
            Top             =   1790
            Width           =   1455
            _ExtentX        =   2566
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
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   "Buscar"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin VB.Frame fraFamiliares 
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
            Height          =   3670
            Left            =   2040
            TabIndex        =   37
            Top             =   2850
            Width           =   8595
            Begin MyCommandButton.MyButton cmdAsignar 
               Height          =   375
               Left            =   7170
               TabIndex        =   10
               ToolTipText     =   "Asignar como parte de la lista de deudores incobrables"
               Top             =   1590
               Width           =   1335
               _ExtentX        =   2355
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
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   15790320
               Caption         =   "Asignar"
               DepthEvent      =   1
               ShowFocus       =   -1  'True
            End
            Begin VSFlex7LCtl.VSFlexGrid grdRegistrados 
               Height          =   1095
               Left            =   90
               TabIndex        =   9
               ToolTipText     =   "Personas registradas"
               Top             =   465
               Width           =   8415
               _cx             =   14843
               _cy             =   1931
               _ConvInfo       =   1
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483638
               GridColorFixed  =   -2147483638
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   1
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin VSFlex7LCtl.VSFlexGrid grdAsignados 
               Height          =   1095
               Left            =   90
               TabIndex        =   11
               ToolTipText     =   "Personas asignadas"
               Top             =   1995
               Width           =   8415
               _cx             =   14843
               _cy             =   1931
               _ConvInfo       =   1
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483638
               GridColorFixed  =   -2147483638
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   1
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   0
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin MyCommandButton.MyButton cmdBorrar 
               Height          =   495
               Left            =   7530
               TabIndex        =   12
               ToolTipText     =   "Agregar otra persona a la lista de deudores incobrables"
               Top             =   3120
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
               Picture         =   "frmListaNegra.frx":0130
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   15790320
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmListaNegra.frx":0844
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdAgregar 
               Height          =   495
               Left            =   8010
               TabIndex        =   13
               ToolTipText     =   "Agregar una persona a la lista de deudores incobrables"
               Top             =   3120
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
               Picture         =   "frmListaNegra.frx":0F58
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   15790320
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmListaNegra.frx":166C
               PictureAlignment=   5
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Registrados"
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
               Left            =   90
               TabIndex        =   38
               Top             =   200
               Width           =   1110
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Asignados"
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
               Left            =   90
               TabIndex        =   39
               Top             =   1730
               Width           =   990
            End
         End
         Begin VB.Label lblMensaje 
            BackColor       =   &H80000005&
            Caption         =   "Se encontró la siguiente información en la lista de deudores incobrables"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   5400
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Número"
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
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo de cliente"
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
            Top             =   690
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nombre"
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
            TabIndex        =   24
            Top             =   2660
            Width           =   795
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
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
            Height          =   250
            Left            =   120
            TabIndex        =   25
            Top             =   7420
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000005&
            Caption         =   "Fecha de ingreso"
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
            Left            =   120
            TabIndex        =   29
            Top             =   6620
            Width           =   2175
         End
         Begin VB.Label lblRef 
            BackColor       =   &H80000005&
            Caption         =   "Referencia"
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
            TabIndex        =   30
            Top             =   2240
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000005&
            Caption         =   "Empleado que registra"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   33
            Top             =   6885
            Width           =   1815
         End
         Begin VB.Label lblFamiliares 
            BackColor       =   &H80000005&
            Caption         =   "Familiares, persona para casos de emergencia y responsable de la cuenta"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Left            =   120
            TabIndex        =   36
            Top             =   3060
            Width           =   1740
         End
      End
      Begin VB.Frame fraControl 
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
         Left            =   4770
         TabIndex        =   18
         Top             =   8640
         Width           =   1920
         Begin MyCommandButton.MyButton cmdLocate 
            Height          =   600
            Left            =   60
            TabIndex        =   19
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
            Picture         =   "frmListaNegra.frx":1D80
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmListaNegra.frx":2704
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSave 
            Height          =   600
            Left            =   660
            TabIndex        =   20
            ToolTipText     =   "Grabar"
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
            Picture         =   "frmListaNegra.frx":3088
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmListaNegra.frx":3A0C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   1260
            TabIndex        =   21
            ToolTipText     =   "Borrar"
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
            Picture         =   "frmListaNegra.frx":4390
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmListaNegra.frx":4D12
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.Label lblInst 
         BackColor       =   &H80000005&
         Caption         =   "Teclee la contraseña y presione continuar para seguir con la operación"
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
         Left            =   480
         TabIndex        =   51
         Top             =   8760
         Visible         =   0   'False
         Width           =   7425
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000005&
         Caption         =   "Teclee la contraseña y presione continuar para seguir con la operación"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -74520
         TabIndex        =   50
         Top             =   8760
         Width           =   7425
      End
   End
End
Attribute VB_Name = "frmListaNegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim vlblnConsulta As Boolean
Dim lblnContinuar As Boolean
Dim lstrPassword As String
Public vllngNumeroOpcion As Long
Dim lstrParentescos As String
Public llngCveFamiliarListaNegra As Long
Public lblnCoincidePersonListaNegra As Boolean
Dim llngParentescoPadre As Long
Dim llngParentescoMadre As Long
Dim llngParentescoConyuge As Long

Private Sub chkActivo_GotFocus()
    On Error GoTo NotificaError

    If vlblnConsulta Then pHabilita 0, 1, 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkActivo_GotFocus"))
End Sub

Private Sub cmdAgregar_Click()
On Error GoTo NotificaError
    grdAsignados.AddItem ""
    grdAsignados.Col = 1
    grdAsignados.Row = grdAsignados.Rows - 1
    cmdBorrar.Enabled = True
    pHabilita 0, 1, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAgregar_Click"))
End Sub

Private Sub cmdAsignar_Click()
On Error GoTo NotificaError
    Dim lintAux As Integer
    
    With grdRegistrados
        For lintAux = 1 To .Rows - 1
            If .Cell(flexcpChecked, lintAux, 1) = flexChecked Then
                If fblnFamiliarNoAsignado(Val(.TextMatrix(lintAux, 5))) Then
                    grdAsignados.AddItem ""
                    grdAsignados.TextMatrix(grdAsignados.Rows - 1, 1) = IIf(IsNull(.TextMatrix(lintAux, 2)), "", .TextMatrix(lintAux, 2))
                    grdAsignados.TextMatrix(grdAsignados.Rows - 1, 2) = IIf(IsNull(.TextMatrix(lintAux, 6)), "", .TextMatrix(lintAux, 6))
                    grdAsignados.TextMatrix(grdAsignados.Rows - 1, 3) = IIf(IsNull(.TextMatrix(lintAux, 7)), "", .TextMatrix(lintAux, 7))
                    grdAsignados.TextMatrix(grdAsignados.Rows - 1, 4) = IIf(IsNull(.TextMatrix(lintAux, 8)), "", .TextMatrix(lintAux, 8))
                    grdAsignados.TextMatrix(grdAsignados.Rows - 1, 5) = IIf(IsNull(.TextMatrix(lintAux, 4)), "", .TextMatrix(lintAux, 4))
                    grdAsignados.TextMatrix(grdAsignados.Rows - 1, 6) = IIf(IsNull(.TextMatrix(lintAux, 5)), "", .TextMatrix(lintAux, 5))
                End If
                .Cell(flexcpChecked, lintAux, 1) = flexUnchecked
            End If
        Next lintAux
    End With
    cmdBorrar.Enabled = grdAsignados.Rows > 1
    pHabilita 0, 1, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAsignar_Click"))
End Sub
Private Function fblnFamiliarNoAsignado(lcveParentesco As Long) As Boolean
    Dim intAux As Integer
    fblnFamiliarNoAsignado = True
    If lcveParentesco <> 0 Then
         For intAux = 1 To grdAsignados.Rows - 1
            If Val(grdAsignados.TextMatrix(intAux, 6)) = llngParentescoPadre And lcveParentesco = llngParentescoPadre Then
                MsgBox "Ya se ha asignado el padre del paciente.", vbOKOnly + vbInformation, "Mensaje"
                fblnFamiliarNoAsignado = False
                Exit For
            ElseIf Val(grdAsignados.TextMatrix(intAux, 6)) = llngParentescoMadre And lcveParentesco = llngParentescoMadre Then
                MsgBox "Ya se ha asignado la madre del paciente.", vbOKOnly + vbInformation, "Mensaje"
                fblnFamiliarNoAsignado = False
                Exit For
            ElseIf Val(grdAsignados.TextMatrix(intAux, 6)) = llngParentescoConyuge And lcveParentesco = llngParentescoConyuge Then
                MsgBox "Ya se ha asignado el cónyuge del paciente.", vbOKOnly + vbInformation, "Mensaje"
                fblnFamiliarNoAsignado = False
                Exit For
            End If
        Next intAux
    End If
End Function
Private Sub cmdBorrar_Click()
    If grdAsignados.Row > 0 Then
        grdAsignados.RemoveItem grdAsignados.Row
        cmdBorrar.Enabled = grdAsignados.Rows > 1
        pHabilita 0, 1, 0
    End If
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError

    fraBusqueda.Visible = True
    txtIniciales.Text = ""
    lstBusqueda.Clear
    txtIniciales.SetFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBuscar_Click"))
End Sub

Private Sub cmdCargar_Click()
    On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim intCampoItemData As Integer
    
    lstBusqueda.Clear
    lstBusqueda.Visible = False
    
    If txtIniciales.Text <> "" Then
        If optTipoCliente(0).Value Then
            'Pacientes
            Set rs = frsEjecuta_SP(Trim(txtIniciales.Text), "SP_GNSELNOMBREPACIENTE")
            intCampoItemData = 1
        Else
            intCampoItemData = 1
            If optTipoCliente(1).Value Then
                'Empleados activos
                Set rs = frsEjecuta_SP(Trim(txtIniciales.Text) & "|" & vgintClaveEmpresaContable, "SP_CCSelNombreEmpleados")
            Else
                If optTipoCliente(2).Value Then
                    'Médicos activos
                    Set rs = frsEjecuta_SP(Trim(txtIniciales.Text), "SP_CCSelNombreMedicos")
                Else
                    'Empresas activas
                    Set rs = frsEjecuta_SP(Trim(txtIniciales.Text), "SP_CCSelNombreEmpresas")
                End If
            End If
        End If
        If rs.State <> adStateClosed Then
            If rs.RecordCount <> 0 Then
                pLlenarListRs lstBusqueda, rs, intCampoItemData, 0
                lstBusqueda.ListIndex = 0
                lstBusqueda.SetFocus
            Else
                'No existe información con esos parámetros.
                MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
            End If
            rs.Close
        End If
        lstBusqueda.Visible = True
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCargar_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ValidaIntegridad
    
    Dim strSentencia As String
    Dim lngPersonaGraba As Long
    
    lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If lngPersonaGraba = 0 Then Exit Sub
    EntornoSIHO.ConeccionSIHO.BeginTrans
    strSentencia = "delete from CCListaNegraFamiliares where intCveListaNegra = " & txtNumeroCliente.Text
    pEjecutaSentencia strSentencia
    strSentencia = "delete from CCListaNegra where intCveListaNegra = " & txtNumeroCliente.Text
    pEjecutaSentencia strSentencia
    pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "LISTA DE DEUDORES INCOBRABLES", txtNumeroCliente.Text
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
    txtNumeroCliente.SetFocus
    
    Exit Sub
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError

    SSTab1.Tab = 1
    grdBusqueda.SetFocus
    pCargaDatosBusqueda
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdModPass_Click()
    On Error GoTo NotificaError

    SSTab1.Tab = 2
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdModPass_Click"))
End Sub

Private Sub cmdNoAutoriza_Click()
    On Error GoTo NotificaError

    lblnContinuar = False
    Me.Hide
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdNoAutoriza_Click"))
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo NotificaError

    If Not fraAutoriza.Visible Then
        Unload Me
    Else
        cmdNoAutoriza_Click
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSalir_Click"))
End Sub

Private Sub cmdSave_Click()
On Error GoTo NotificaError
    
    Dim rs As ADODB.Recordset
    Dim rsFam As ADODB.Recordset
    Dim strSentencia As String
    Dim lngPersonaGraba As Long
    Dim vllngNumeroCuenta As Long
    Dim vlintErrorCuenta As Integer
    Dim intOperacion As TipoOperacion
    Dim strTipo As String
    Dim lintAux As Integer
    
    'Checar el pemiso que le mandan
    If fblnRevisaPermiso(vglngNumeroLogin, 2267, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2267, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2370, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2370, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2053, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2053, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2054, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2054, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2379, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2379, "C", True) Then
        If fblnDatosValidos() Then
            lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If lngPersonaGraba = 0 Then Exit Sub
            
            If optTipoCliente(0).Value Then
                strTipo = "PA"  'Paciente
            ElseIf optTipoCliente(1).Value Then
                strTipo = "EM"  'Empleado
            ElseIf optTipoCliente(2).Value Then
                strTipo = "ME"  'Médico
            ElseIf optTipoCliente(3).Value Then
                strTipo = "CO"  'Empresa
            End If
            
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            strSentencia = "select * from CCListaNegra where intCveListaNegra = " & IIf(vlblnConsulta, txtNumeroCliente.Text, "-1")
            Set rs = frsRegresaRs(strSentencia, adLockOptimistic, adOpenStatic)
            With rs
                intOperacion = TipoOperacion.EnmCambiar
                If Not vlblnConsulta Then
                  .AddNew
                  !INTNUMREFERENCIA = txtReferencia.Text
                  !chrTipoCliente = strTipo
                  !dtmFechaRegistro = fdtmServerFechaHora
                  !intCveEmpleadoRegistra = lngPersonaGraba
                  intOperacion = TipoOperacion.EnmGrabar
                End If
                !VCHCOMENTARIOS = txtObservaciones.Text
                !bitactivo = IIf(chkActivo.Value = vbChecked, 1, 0)
                .Update
                If Not vlblnConsulta Then
                  txtNumeroCliente.Text = flngObtieneIdentity("sec_CCListaNegra", !intCveListaNegra)
                End If
            End With
            rs.Close
            pGuardarLogTransaccion Me.Name, intOperacion, lngPersonaGraba, "LISTA DE DEUDORES INCOBRABLES", txtNumeroCliente.Text
            If strTipo = "PA" Then
                If vlblnConsulta Then
                    strSentencia = "delete from CCListaNegraFamiliares where intCveListaNegra = " & txtNumeroCliente.Text
                    pEjecutaSentencia strSentencia
                End If
                If grdAsignados.Rows > 1 Then
                    strSentencia = "select * from CCListaNegraFamiliares where intCveListaNegra = -1"
                    Set rsFam = frsRegresaRs(strSentencia, adLockOptimistic, adOpenStatic)
                    With rsFam
                        For lintAux = 1 To grdAsignados.Rows - 1
                            .AddNew
                            !intCveListaNegra = Val(txtNumeroCliente.Text)
                            !vchApellidoPaterno = grdAsignados.TextMatrix(lintAux, 2)
                            !vchApellidoMaterno = grdAsignados.TextMatrix(lintAux, 3)
                            !vchNombre = grdAsignados.TextMatrix(lintAux, 4)
                            !dtmFechaNacimiento = CDate(grdAsignados.TextMatrix(lintAux, 5))
                            !intCveParentesco = Val(grdAsignados.TextMatrix(lintAux, 6))
                            .Update
                        Next lintAux
                    End With
                    rsFam.Close
                End If
            End If
            
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            txtNumeroCliente.SetFocus
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub
Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError
    Dim lintAux As Integer
    fblnDatosValidos = True
    ' si es paciente, se validan los familiares asignados
    If optTipoCliente(0).Value = True Then
        With grdAsignados
            If grdAsignados.Rows > 1 Then
                For lintAux = 1 To grdAsignados.Rows - 1
                    'parentesco
                    If Val(grdAsignados.TextMatrix(lintAux, 6)) = 0 Then
                        fblnDatosValidos = False
                        .Col = 1
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Parentesco", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    'paterno
                    ElseIf IsNull(grdAsignados.TextMatrix(lintAux, 2)) Then
                        fblnDatosValidos = False
                        .Col = 2
                        .Row = lintAux
                        .SetFocus
                        '¡No ha ingresado datos!
                        MsgBox SIHOMsg(2) & vbCrLf & "Apellido paterno", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    ElseIf grdAsignados.TextMatrix(lintAux, 2) = "" Then
                        fblnDatosValidos = False
                        .Col = 2
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Apellido paterno", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    'materno
                    ElseIf IsNull(grdAsignados.TextMatrix(lintAux, 3)) Then
                        fblnDatosValidos = False
                        .Col = 3
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Apellido materno", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    ElseIf grdAsignados.TextMatrix(lintAux, 3) = "" Then
                        fblnDatosValidos = False
                        .Col = 3
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Apellido materno", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    'nombre
                    ElseIf IsNull(grdAsignados.TextMatrix(lintAux, 4)) Then
                        fblnDatosValidos = False
                        .Col = 4
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Nombre", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    ElseIf grdAsignados.TextMatrix(lintAux, 4) = "" Then
                        fblnDatosValidos = False
                        .Col = 4
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Nombre", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    'fecha nacimiento
                    ElseIf IsNull(grdAsignados.TextMatrix(lintAux, 5)) Then
                        fblnDatosValidos = False
                        .Col = 5
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Fecha nacimiento", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    ElseIf grdAsignados.TextMatrix(lintAux, 5) = "  /  /    " Or grdAsignados.TextMatrix(lintAux, 5) = "" Then
                        fblnDatosValidos = False
                        .Col = 5
                        .Row = lintAux
                        .SetFocus
                        MsgBox SIHOMsg(2) & vbCrLf & "Fecha nacimiento", vbInformation + vbOKOnly, "Mensaje"
                        Exit For
                    End If
                Next lintAux
            End If
        End With
    End If
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function
Private Sub cmdSiAutoriza_Click()
    On Error GoTo NotificaError

    Dim lstrContrasenaDada As String
    Dim lintx As Integer
    Dim lblnPasswordOK As Integer
    Dim lintAux As Integer
    
    lstrContrasenaDada = fstrEncrypt(UCase(txtPassword.Text), "LISTANEGRA")
    If gbitNewEncrypt = 1 Then
            lblnPasswordOK = IIf(lstrContrasenaDada = lstrPassword, 1, 0)
    Else
        If Len(lstrContrasenaDada) = Len(lstrPassword) Then
            lblnPasswordOK = 1
            For lintx = 1 To Len(lstrContrasenaDada)
                If Asc(Mid(lstrContrasenaDada, lintx, 1)) <> Asc(Mid(lstrPassword, lintx, 1)) Then lblnPasswordOK = 0
            Next lintx
        End If
    End If
    If lblnPasswordOK = 1 Then
        lblnContinuar = True
        If SSTab1.Tab = 2 Then
            'Si es un familiar o aval en lista negra, se identificá si se seleccionó alguna persona del grid
            For lintAux = 1 To grdPersonas.Rows - 1
                If grdPersonas.Cell(flexcpChecked, lintAux, 1) = flexChecked Then
                    llngCveFamiliarListaNegra = Val(grdPersonas.TextMatrix(lintAux, 12))
                    Exit For
                End If
            Next lintAux
        End If
        Me.Hide
    Else
        'La contraseña no coincide, verificar nuevamente
        MsgBox SIHOMsg(763), vbExclamation, "Mensaje"
        pEnfocaTextBox txtPassword
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSiAutoriza_Click"))
End Sub

Private Sub cmdSiguiente_Click()
    On Error GoTo NotificaError

    If Me.ActiveControl.Name = "txtNumeroCliente" Then
        If IsNumeric(txtNumeroCliente.Text) Then
            pMuestraCliente CLng(txtNumeroCliente.Text)
        End If
    ElseIf Me.ActiveControl.Name = "lstBusqueda" Then
        lstBusqueda_DblClick
    ElseIf Me.ActiveControl.Name = "grdBusqueda" Then
        grdBusqueda_DblClick
    Else
        SendKeys vbTab
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSiguiente_Click"))
End Sub

Private Sub Form_Activate()
    If SSTab1.Tab = 0 And fraAutoriza.Visible Then
        If fblnCanFocus(txtPassword) Then txtPassword.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        grdPersonas.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    'Color del Tab
    SetStyle SSTab1.hwnd, 0
    SetSolidColor SSTab1.hwnd, 16777215
    SSTabSubclass SSTab1.hwnd
    ''''''''''''''''''''''''
    
    Me.Icon = frmMenuPrincipal.Icon
    SSTab1.Tab = 0
    pLimpia
    pConfiguraGrid
    pConfiguraGridFamiliares
    fraBusqueda.BorderStyle = 0
    lstrParentescos = fstrParentesco
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub
Private Function fstrParentesco() As String
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim lstrSentencia As String
    lstrSentencia = "select * from SiParentesco order by vchdescripcion"
    Set rs = frsRegresaRs(lstrSentencia)
    fstrParentesco = ""
    Do While Not rs.EOF
        If rs!bitPadre = 1 Then
            llngParentescoPadre = rs!intCveParentesco
        ElseIf rs!bitMadre = 1 Then
            llngParentescoMadre = rs!intCveParentesco
        ElseIf rs!bitConyuge = 1 Then
            llngParentescoConyuge = rs!intCveParentesco
        End If
        
        fstrParentesco = fstrParentesco & "|#" & Trim(Str(rs!intCveParentesco)) & ";" & Trim(rs!VCHDESCRIPCION)
        rs.MoveNext
    Loop
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrParentesco"))
End Function
Private Sub pConfiguraGridFamiliares()
On Error GoTo NotificaError
    With grdRegistrados
        .Clear
        .Cols = 9
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "||Parentesco|Nombre|Fecha nacimiento||||"
        .ColWidth(0) = 100
        .ColWidth(1) = 300  'check seleccionar
        .ColWidth(2) = 1150 'Parentesco
        .ColWidth(3) = 4000 'Nombre completo
        .ColWidth(4) = 1900 'Fecha de nacimiento
        .ColWidth(5) = 0    'cve parentesco
        .ColWidth(6) = 0    'paterno
        .ColWidth(7) = 0    'materno
        .ColWidth(8) = 0    'nombre
        .ColDataType(1) = flexDTBoolean
        .ColAlignment(1) = flexAlignCenterCenter
    End With
    grdRegistrados.Editable = flexEDKbdMouse
            
    With grdAsignados
        .Clear
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Parentesco|Apellido paterno|Apellido materno|Nombre|Fecha nacimiento|"
        .ColWidth(0) = 100
        .ColWidth(1) = 1210    'Prentesco
        .ColWidth(2) = 1750     'Paterno
        .ColWidth(3) = 1830     'Materno
        .ColWidth(4) = 1520     'Nombre
        .ColWidth(5) = 1900     'Fecha nacimiento
        .ColWidth(6) = 0        'cve parentesco
        .ColFormat(5) = "dd/mm/yyyy"
        .ColEditMask(5) = "##/##/####"
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGridFamiliares"))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    If Me.Visible Then
        If UnloadMode = vbFormControlMenu And fraAutoriza.Visible Then
            Cancel = True
        Else
            If SSTab1.Tab <> 0 Then
                SSTab1.Tab = 0
                txtNumeroCliente.SetFocus
                Cancel = True
            Else
                If fraBusqueda.Visible Then
                    cmdBuscar.SetFocus
                    Cancel = True
                    fraBusqueda.Visible = False
                Else
                    If cmdSave.Enabled Or vlblnConsulta Then
                        '¿Desea abandonar la operación?
                        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                            txtNumeroCliente.SetFocus
                        End If
                        Cancel = True
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer)
    On Error GoTo NotificaError
    
    cmdLocate.Enabled = vlb1 = 1
    cmdSave.Enabled = vlb2 = 1
    cmdDelete.Enabled = vlb3 = 1
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

Private Sub grdAsignados_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'Parentesco
    If Col = 1 Then
        If grdAsignados.ComboIndex <> -1 Then
            grdAsignados.TextMatrix(Row, 6) = grdAsignados.ComboData(grdAsignados.ComboIndex)
        Else
            grdAsignados.TextMatrix(Row, 6) = ""
        End If
    End If
End Sub

Private Sub grdAsignados_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Then
        grdAsignados.ComboList = lstrParentescos
    Else
        grdAsignados.ComboList = ""
    End If
End Sub

Private Sub grdAsignados_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If (Col <> 5) Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    pHabilita 0, 1, 0
End Sub

Private Sub grdBusqueda_DblClick()
    On Error GoTo NotificaError

    If grdBusqueda.Row > 0 Then
        SSTab1.Tab = 0
        pMuestraCliente CLng(grdBusqueda.TextMatrix(grdBusqueda.Row, 1))
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdBusqueda_DblClick"))
End Sub

Private Sub grdAsignados_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim intAux As Integer
    Dim lcveParentesco As Long
    
    If Col = 5 Then
        If IsDate(grdAsignados.EditText) Then
            If CDate(grdAsignados.EditText) > CDate(fdtmServerFecha) Or CDate(grdAsignados.EditText) < CDate("01/01/1900") Then
                Cancel = True
            Else
                grdAsignados.EditMaxLength = 10
            End If
        Else
            Cancel = True
        End If
    ElseIf Col = 2 Or Col = 3 Or Col = 4 Then
        grdAsignados.EditMaxLength = 100
    ElseIf Col = 1 Then
        If grdAsignados.ComboIndex <> -1 Then
            lcveParentesco = grdAsignados.ComboData(grdAsignados.ComboIndex)
            For intAux = 1 To grdAsignados.Rows - 1
            If intAux <> Row And Val(grdAsignados.TextMatrix(intAux, 6)) = llngParentescoPadre And lcveParentesco = llngParentescoPadre Then
                MsgBox "Ya se ha asignado el padre del paciente.", vbOKOnly + vbInformation, "Mensaje"
                grdAsignados.ComboIndex = 0
                Cancel = True
                Exit For
            ElseIf intAux <> Row And Val(grdAsignados.TextMatrix(intAux, 6)) = llngParentescoMadre And lcveParentesco = llngParentescoMadre Then
                MsgBox "Ya se ha asignado la madre del paciente.", vbOKOnly + vbInformation, "Mensaje"
                grdAsignados.ComboIndex = 0
                Cancel = True
                Exit For
            ElseIf intAux <> Row And Val(grdAsignados.TextMatrix(intAux, 6)) = llngParentescoConyuge And lcveParentesco = llngParentescoConyuge Then
                MsgBox "Ya se ha asignado el cónyuge del paciente.", vbOKOnly + vbInformation, "Mensaje"
                grdAsignados.ComboIndex = 0
                Cancel = True
                Exit For
            End If
            Next intAux
        End If
    End If
    
End Sub

Private Sub grdPersonas_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lintAux As Integer
    If Col = 1 Then
        If grdPersonas.Cell(flexcpChecked, Row, 1) = flexChecked Then
            ' se des seleccionan las demas personas
            For lintAux = 1 To grdPersonas.Rows - 1
                If lintAux <> Row And grdPersonas.Cell(flexcpChecked, lintAux, 1) = flexChecked Then
                    grdPersonas.Cell(flexcpChecked, lintAux, 1) = flexUnchecked
                End If
            Next lintAux
        End If
    End If
End Sub

Private Sub grdPersonas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub

Private Sub grdPersonas_Click()
    If grdPersonas.Row > 0 Then
        lblNombrePac.Caption = grdPersonas.TextMatrix(grdPersonas.Row, 7) 'paciente
        lblExpediente.Caption = grdPersonas.TextMatrix(grdPersonas.Row, 8) 'expediente
        lblFechaRegistro.Caption = grdPersonas.TextMatrix(grdPersonas.Row, 9) 'fechaRegistro
        lblEmpleadoRegistro.Caption = grdPersonas.TextMatrix(grdPersonas.Row, 10) 'empleado
        txtObservacionesFamiliar.Text = grdPersonas.TextMatrix(grdPersonas.Row, 11) 'observaciones
    End If
End Sub

Private Sub grdRegistrados_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    End If
End Sub

Private Sub lstBusqueda_DblClick()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim strTipo As String

    If lstBusqueda.ListIndex >= 0 Then
        txtNombreCliente.Text = lstBusqueda.List(lstBusqueda.ListIndex)
        txtReferencia.Text = lstBusqueda.ItemData(lstBusqueda.ListIndex)
        fraBusqueda.Visible = False
        
        'Verifica si ya está como cliente
        If optTipoCliente(0).Value Then
                strTipo = "PA"  'Paciente
            ElseIf optTipoCliente(1).Value Then
                strTipo = "EM"  'Empleado
            ElseIf optTipoCliente(2).Value Then
                strTipo = "ME"  'Médico
            ElseIf optTipoCliente(3).Value Then
                strTipo = "CO"  'Empresa
            End If
        Set rs = frsRegresaRs("select * from CCListaNegra where intNumReferencia = " & Trim(txtReferencia.Text) & " and chrTipoCliente = '" & strTipo & "'")
        If rs.RecordCount <> 0 Then
            txtNumeroCliente.Text = rs!intCveListaNegra
            pMuestraCliente CLng(txtNumeroCliente.Text)
            pHabilita 0, 1, 0
            If fblnCanFocus(txtObservaciones) Then
                txtObservaciones.SetFocus
                pSelTextBox txtObservaciones
            End If
        Else
            grdRegistrados.Rows = 1
            grdAsignados.Rows = 1
            If optTipoCliente(0).Value = True Then
                pConsultaFamiliares
            Else
                txtObservaciones.SetFocus
            End If
            pHabilita 0, 1, 0
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstBusqueda_DblClick"))
End Sub
Private Sub pConsultaFamiliares()
On Error GoTo NotificaError
Dim rs As New ADODB.Recordset
    Set rs = frsEjecuta_SP(txtReferencia.Text, "SP_CCSELFAMILIARESAVALES")
    With grdRegistrados
        .Rows = 1
        Do Until rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 2) = IIf(IsNull(rs!parentesco), "", rs!parentesco)
            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(rs!nombreCompleto), "", rs!nombreCompleto)
            .TextMatrix(.Rows - 1, 4) = IIf(IsNull(rs!FechaNacimiento), "", Format(rs!FechaNacimiento, "dd/MMM/yyyy"))
            .TextMatrix(.Rows - 1, 5) = IIf(IsNull(rs!cveParentesco), "", rs!cveParentesco)
            .TextMatrix(.Rows - 1, 6) = IIf(IsNull(rs!Paterno), "", rs!Paterno)
            .TextMatrix(.Rows - 1, 7) = IIf(IsNull(rs!Materno), "", rs!Materno)
            .TextMatrix(.Rows - 1, 8) = IIf(IsNull(rs!Nombre), "", rs!Nombre)
            rs.MoveNext
        Loop
        If rs.RecordCount <> 0 Then
            grdRegistrados.Col = 1
            grdRegistrados.Row = 1
            If vlblnConsulta = False Then grdRegistrados.SetFocus
            cmdAsignar.Enabled = True
        Else
            cmdAsignar.Enabled = False
            If vlblnConsulta = False Then txtObservaciones.SetFocus
        End If
        cmdAgregar.Enabled = True
        cmdBorrar.Enabled = grdAsignados.Rows > 1
    End With
    rs.Close
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConsultaFamiliares"))
End Sub

Private Sub optTipoCliente_Click(Index As Integer)
    On Error GoTo NotificaError

    fraBusqueda.Visible = False
    
    Select Case Index
        Case 0
            lblRef.Caption = "Expediente"
            pHabilitaFamiliares True
        Case 1
            lblRef.Caption = "Clave empleado"
            pHabilitaFamiliares False
        Case 2
            lblRef.Caption = "Clave médico"
            pHabilitaFamiliares False
        Case 3
            lblRef.Caption = "Clave empresa"
            pHabilitaFamiliares False
    End Select

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoCliente_Click"))
End Sub
Private Sub pHabilitaFamiliares(blnHabilitar As Boolean)
On Error GoTo NotificaError
    lblFamiliares.Enabled = blnHabilitar
    fraFamiliares.Enabled = blnHabilitar
    cmdAsignar.Enabled = False
    cmdAgregar.Enabled = False
    cmdBorrar.Enabled = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaFamiliares"))
End Sub

Private Sub txtIniciales_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cmdCargar.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIniciales_KeyDown"))
End Sub

Private Sub txtIniciales_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIniciales_KeyPress"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    Dim rs As ADODB.Recordset
    
    vlblnConsulta = False
    cmdBuscar.Enabled = True
    Set rs = frsRegresaRs("select isnull(max(intCveListaNegra), 0) + 1 from CCListaNegra")
    If rs.RecordCount <> 0 Then
        txtNumeroCliente.Text = rs.Fields(0)
    Else
        txtNumeroCliente.Text = "1"
    End If
    optTipoCliente(0).Value = True
    optTipoCliente(1).Value = False
    optTipoCliente(2).Value = False
    optTipoCliente(3).Value = False
    optTipoCliente_Click 0
    txtIniciales.Text = ""
    txtReferencia.Text = ""
    txtEmpleado.Text = ""
    txtObservaciones.Text = ""
    txtFecha.Text = ""
    lstBusqueda.Clear
    fraBusqueda.Visible = False
    txtNombreCliente.Text = ""
    chkActivo.Value = 1
    
    grdRegistrados.Rows = 1
    grdAsignados.Rows = 1
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Private Sub txtNumeroCliente_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 1, 0, 0
    pLimpia
    pSelTextBox txtNumeroCliente
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCliente_GotFocus"))
End Sub

Private Sub pValidateNumeric(ByRef KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 8 Then Exit Sub
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pValidateNumeric"))
End Sub

Private Sub txtNumeroCliente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    pValidateNumeric KeyAscii
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCliente_KeyPress"))
End Sub

Private Sub pMuestraCliente(vllngxNumero As Long)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim rsFam As New ADODB.Recordset
    Dim intIndex As Integer
    
    vgstrParametrosSP = Str(vllngxNumero)
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_CCSelListaNegra")
    If rs.RecordCount <> 0 Then
        vlblnConsulta = True
        cmdBuscar.Enabled = False
        txtNumeroCliente = rs!intCveListaNegra
        txtReferencia.Text = rs!INTNUMREFERENCIA
        optTipoCliente(0).Value = rs!chrTipoCliente = "PA"
        optTipoCliente(1).Value = rs!chrTipoCliente = "EM"
        optTipoCliente(2).Value = rs!chrTipoCliente = "ME"
        optTipoCliente(3).Value = rs!chrTipoCliente = "CO"
        For intIndex = 0 To 3
            If optTipoCliente(intIndex).Value Then
                optTipoCliente_Click intIndex
            End If
        Next intIndex
        txtNombreCliente.Text = rs!CLIENTE
        txtEmpleado.Text = rs!EmpleadoRegistra
        txtObservaciones.Text = IIf(IsNull(rs!VCHCOMENTARIOS), " ", rs!VCHCOMENTARIOS)
        txtFecha.Text = Format(rs!dtmFechaRegistro, "dd/MM/yyyy HH:mm")
        chkActivo.Value = rs!bitactivo
        grdRegistrados.Rows = 1
        grdAsignados.Rows = 1
        If rs!chrTipoCliente = "PA" Then
            vgstrParametrosSP = Str(vllngxNumero)
            Set rsFam = frsEjecuta_SP(vgstrParametrosSP, "SP_CCSELLISTANEGRAFAMILIARES")
            Do Until rsFam.EOF
                grdAsignados.AddItem ""
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 1) = rsFam!parentesco
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 2) = rsFam!vchApellidoPaterno
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 3) = rsFam!vchApellidoMaterno
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 4) = rsFam!vchNombre
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 5) = Format(rsFam!dtmFechaNacimiento, "dd/MMM/yyyy")
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 6) = rsFam!intCveParentesco
                rsFam.MoveNext
            Loop
            pConsultaFamiliares
        End If
        pHabilita 1, 0, 1
        cmdLocate.SetFocus
    Else
        optTipoCliente(0).SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraCliente"))
End Sub

Private Sub txtObservaciones_GotFocus()
On Error GoTo NotificaError
    If vlblnConsulta Then pHabilita 0, 1, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtObservaciones_GotFocus"))
End Sub

Private Sub pCargaDatosBusqueda()
    On Error GoTo NotificaError
    
    Dim rs As ADODB.Recordset
    Dim strTipoCliente As String
    grdBusqueda.Rows = 1
    Set rs = frsEjecuta_SP("-1", "sp_CCSelListaNegra")
    Do Until rs.EOF
        Select Case rs!chrTipoCliente
            Case "PA"
                strTipoCliente = "PACIENTE"
            Case "CO"
                strTipoCliente = "EMPRESA"
            Case "EM"
                strTipoCliente = "EMPLEADO"
            Case "ME"
                strTipoCliente = "MEDICO"
        End Select
        grdBusqueda.AddItem ""
        grdBusqueda.TextMatrix(grdBusqueda.Rows - 1, 1) = rs!intCveListaNegra
        grdBusqueda.TextMatrix(grdBusqueda.Rows - 1, 2) = rs!CLIENTE
        grdBusqueda.TextMatrix(grdBusqueda.Rows - 1, 3) = rs!INTNUMREFERENCIA
        grdBusqueda.TextMatrix(grdBusqueda.Rows - 1, 4) = strTipoCliente
        grdBusqueda.TextMatrix(grdBusqueda.Rows - 1, 5) = Format(rs!dtmFechaRegistro, "dd/MM/yyyy")
        grdBusqueda.TextMatrix(grdBusqueda.Rows - 1, 6) = rs!EmpleadoRegistra
        grdBusqueda.TextMatrix(grdBusqueda.Rows - 1, 7) = IIf(rs!bitactivo = 0, "NO", "SI")
        rs.MoveNext
    Loop
    rs.Close
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDatosBusqueda"))
End Sub

Private Sub pConfiguraGrid()
    On Error GoTo NotificaError

    grdBusqueda.Cols = 8
    grdBusqueda.Rows = 1
    grdBusqueda.TextMatrix(0, 2) = "Nombre cliente"
    grdBusqueda.TextMatrix(0, 3) = "Referencia"
    grdBusqueda.TextMatrix(0, 4) = "Tipo cliente"
    grdBusqueda.TextMatrix(0, 5) = "Fecha registro"
    grdBusqueda.TextMatrix(0, 6) = "Empleado que registra"
    grdBusqueda.TextMatrix(0, 7) = "Activo"
    grdBusqueda.ColWidth(0) = 100
    grdBusqueda.ColWidth(1) = 1000
    grdBusqueda.ColWidth(2) = 3200
    grdBusqueda.ColWidth(3) = 1000
    grdBusqueda.ColWidth(4) = 1800
    grdBusqueda.ColWidth(5) = 1100
    grdBusqueda.ColWidth(6) = 3200
    grdBusqueda.ColWidth(7) = 600
    grdBusqueda.ColHidden(1) = True
    grdBusqueda.ColAlignment(5) = flexAlignCenterCenter
    grdBusqueda.ColAlignment(7) = flexAlignCenterCenter
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrid"))
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtObservaciones_KeyPress"))
End Sub

Public Function fblnAutorizacion(lngCveLista As Long)
    On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim strPass As String
    If fblnMuestraDatos(lngCveLista) Then

        Set rs = frsSelParametros("AD", -1, "VCHPASSWORDLISTANEGRA")
        
        If Not rs.EOF Then
            lstrPassword = IIf(IsNull(rs!valor), "", rs!valor)
        Else
            lstrPassword = ""
        End If
        rs.Close
        
        'fraInfo.Enabled = False
        txtNumeroCliente.Enabled = False
        optTipoCliente(0).Enabled = False
        optTipoCliente(1).Enabled = False
        optTipoCliente(2).Enabled = False
        optTipoCliente(3).Enabled = False
        cmdBuscar.Enabled = False
        txtReferencia.Enabled = False
        txtNombreCliente.Enabled = False
        grdRegistrados.Enabled = False
        cmdAsignar.Enabled = False
        grdAsignados.Editable = False
        cmdBorrar.Enabled = False
        cmdAgregar.Enabled = False
        txtFecha.Enabled = False
        txtEmpleado.Enabled = False
        'txtObservaciones.Enabled = False
        txtObservaciones.Locked = True
        chkActivo.Enabled = False
        fraGrid.Enabled = False
        fraControl.Visible = False
        cmdBuscar.Visible = False
        lblMensaje.Visible = True
        lblInst.Visible = True
        fraAutoriza.Visible = True
        Me.Show vbModal
    Else
        lblnContinuar = True
    End If
    fblnAutorizacion = lblnContinuar
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnAutorizacion"))
End Function
Public Function fblnAutorizacionFamiliar(lstrPaterno As String, lstrMaterno As String, nombreCompleto As String, fechaNac As String)
On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim strPass As String
    Dim rsPwd As ADODB.Recordset
    Dim lstrFechaNacimiento As String
    
    Set rs = frsEjecuta_SP(lstrPaterno & "|" & lstrMaterno, "SP_ADSELFAMILIARESLISTANEGRA")
    llngCveFamiliarListaNegra = -1
    lblnCoincidePersonListaNegra = False
    If rs.RecordCount <> 0 Then
        lblnCoincidePersonListaNegra = True
        'contraseña
        Set rsPwd = frsSelParametros("AD", -1, "VCHPASSWORDLISTANEGRA")
        If Not rsPwd.EOF Then
            lstrPassword = IIf(IsNull(rsPwd!valor), "", rsPwd!valor)
        Else
            lstrPassword = ""
        End If
        rsPwd.Close
        lstrFechaNacimiento = ""
        If fechaNac <> "" Then
            If IsDate(CDate(fechaNac)) Then
                lstrFechaNacimiento = " con fecha de nacimiento " & Format(fechaNac, "dd/MMM/yyyy")
            End If
        End If
        lblMensajeCoinciden.Caption = "El paciente " & nombreCompleto & lstrFechaNacimiento & _
                                        " coincide con las siguientes personas que están registradas en la lista de deudores incobrables, las cuales fueron relacionadas en la atención de un paciente en la lista de deudores incobrables."
        'Configuracion datos
        pConfiguraGridPersonas
        With grdPersonas
        .Rows = 1
            Do Until rs.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, 2) = IIf(IsNull(rs!vchApellidoPaterno), "", rs!vchApellidoPaterno)
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(rs!vchApellidoMaterno), "", rs!vchApellidoMaterno)
                .TextMatrix(.Rows - 1, 4) = IIf(IsNull(rs!vchNombre), "", rs!vchNombre)
                .TextMatrix(.Rows - 1, 5) = IIf(IsNull(rs!dtmFechaNacimiento), "", Format(rs!dtmFechaNacimiento, "dd/MMM/yyyy"))
                .TextMatrix(.Rows - 1, 6) = IIf(IsNull(rs!parentesco), "", rs!parentesco)
                .TextMatrix(.Rows - 1, 7) = IIf(IsNull(rs!Paciente), "", rs!Paciente)
                .TextMatrix(.Rows - 1, 8) = IIf(IsNull(rs!expediente), "", rs!expediente)
                .TextMatrix(.Rows - 1, 9) = IIf(IsNull(rs!FECHAREGISTRO), "", Format(rs!FECHAREGISTRO, "dd/MMM/yyyy HH:mm"))
                .TextMatrix(.Rows - 1, 10) = IIf(IsNull(rs!Empleado), "", rs!Empleado)
                .TextMatrix(.Rows - 1, 11) = IIf(IsNull(rs!Observaciones), "", rs!Observaciones)
                .TextMatrix(.Rows - 1, 12) = rs!intConsecutivo
                rs.MoveNext
            Loop
            .Col = 1
            .Row = 1
        End With
        
        rs.MoveFirst
        lblNombrePac.Caption = IIf(IsNull(rs!Paciente), "", rs!Paciente)
        lblFechaRegistro.Caption = IIf(IsNull(rs!FECHAREGISTRO), "", Format(rs!FECHAREGISTRO, "dd/MMM/yyyy HH:mm"))
        lblEmpleadoRegistro.Caption = IIf(IsNull(rs!Empleado), "", rs!Empleado)
        txtObservacionesFamiliar.Text = IIf(IsNull(rs!Observaciones), "", rs!Observaciones)
        lblExpediente.Caption = IIf(IsNull(rs!expediente), "", rs!expediente)
        
        lblMensaje.Visible = True
        lblInst.Visible = True
        fraAutoriza.Visible = True
        
        SSTab1.Tab = 2
        Me.Show vbModal
    Else
        lblnContinuar = True
    End If
    fblnAutorizacionFamiliar = lblnContinuar
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnAutorizacionFamiliar"))
End Function
Private Sub pConfiguraGridPersonas()
On Error GoTo NotificaError
    With grdPersonas
        .Clear
        .Cols = 13
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "||Apellido paterno|Apellido materno|Nombre|Fecha nacimiento|Parentesco|||||"
        .ColWidth(0) = 100
        .ColWidth(1) = 300  'check seleccionar
        .ColWidth(2) = 2000 'apellido paterno
        .ColWidth(3) = 2000 'apellido materno
        .ColWidth(4) = 2000 'Nombre completo
        .ColWidth(5) = 1500 'Fecha de nacimiento
        .ColWidth(6) = 1100 'Parentesco
        .ColWidth(7) = 0    'Paciente
        .ColWidth(8) = 0    'expediente
        .ColWidth(9) = 0    'fecha registro
        .ColWidth(10) = 0   'empleado
        .ColWidth(11) = 0   'Observaciones
        .ColWidth(12) = 0   'cve familiar en lista negra
        .ColDataType(1) = flexDTBoolean
        .ColAlignment(1) = flexAlignCenterCenter
        .Editable = flexEDKbdMouse
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGridPersonas"))
End Sub
Private Function fblnMuestraDatos(vllngxNumero As Long) As Boolean
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim rsFam As New ADODB.Recordset
    Dim intIndex As Integer
    
    vgstrParametrosSP = Str(vllngxNumero)
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_CCSelListaNegra")
    If Not rs.EOF Then
        txtNumeroCliente = rs!intCveListaNegra
        txtReferencia.Text = rs!INTNUMREFERENCIA
        optTipoCliente(0).Value = rs!chrTipoCliente = "PA"
        optTipoCliente(1).Value = rs!chrTipoCliente = "EM"
        optTipoCliente(2).Value = rs!chrTipoCliente = "ME"
        optTipoCliente(3).Value = rs!chrTipoCliente = "CO"
        For intIndex = 0 To 3
            If optTipoCliente(intIndex).Value Then
                optTipoCliente_Click intIndex
            End If
        Next
        txtNombreCliente.Text = rs!CLIENTE
        txtEmpleado.Text = rs!EmpleadoRegistra
        txtObservaciones.Text = IIf(IsNull(rs!VCHCOMENTARIOS), "", rs!VCHCOMENTARIOS)
        txtFecha.Text = Format(rs!dtmFechaRegistro, "dd/MM/yyyy HH:mm")
        chkActivo.Value = rs!bitactivo
        grdAsignados.Rows = 1
        grdRegistrados.Rows = 1
        If rs!chrTipoCliente = "PA" Then
            vgstrParametrosSP = Str(vllngxNumero)
            Set rsFam = frsEjecuta_SP(vgstrParametrosSP, "SP_CCSELLISTANEGRAFAMILIARES")
            Do Until rsFam.EOF
                grdAsignados.AddItem ""
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 1) = rsFam!parentesco
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 2) = rsFam!vchApellidoPaterno
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 3) = rsFam!vchApellidoMaterno
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 4) = rsFam!vchNombre
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 5) = Format(rsFam!dtmFechaNacimiento, "dd/MMM/yyyy")
                grdAsignados.TextMatrix(grdAsignados.Rows - 1, 6) = rsFam!intCveParentesco
                rsFam.MoveNext
            Loop
        End If
        fblnMuestraDatos = True
    Else
       fblnMuestraDatos = False
    End If
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnMuestraDatos"))
End Function

Public Sub pConfiguraConsultaInicial()

    Label7.Visible = False
    txtPassword.Visible = False
    cmdSiAutoriza.Visible = False
    cmdNoAutoriza.Caption = "Aceptar"
    
    fraAutoriza.Top = 8480
    fraAutoriza.Left = 9230
    fraAutoriza.Height = 700
    fraAutoriza.Width = 1620
    
    cmdNoAutoriza.Top = 240
    cmdNoAutoriza.Left = 90
    
    lblInst.Left = -7000
    
    frmListaNegra.Height = 9645
    
End Sub
