VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmParametrizacionTicket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formato de ticket"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPresentacionPreliminar 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   11710
   End
   Begin VB.PictureBox pbxCveTicket 
      Height          =   855
      Left            =   4080
      ScaleHeight     =   795
      ScaleWidth      =   3555
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   3375
         Begin VB.TextBox txtCveTicket 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2160
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Introduzca clave del ticket"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   300
            Width           =   2295
         End
      End
   End
   Begin VB.ListBox lstCandidatos 
      Height          =   645
      ItemData        =   "frmParametrizacionTicket.frx":0000
      Left            =   7800
      List            =   "frmParametrizacionTicket.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame fraBotonera 
      Height          =   660
      Left            =   3893
      TabIndex        =   28
      Top             =   7440
      Width           =   4095
      Begin VB.CheckBox chkPresentacionPreliminar 
         Height          =   480
         Left            =   3510
         Picture         =   "frmParametrizacionTicket.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   135
         Width           =   495
      End
      Begin VB.CommandButton cmdPrimer 
         Height          =   480
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrizacionTicket.frx":27A6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Primer registro"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   480
         Left            =   540
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrizacionTicket.frx":2918
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Anterior registro"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   480
         Left            =   1035
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrizacionTicket.frx":2A8A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Búsqueda"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdSiguiente 
         Height          =   480
         Left            =   1530
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrizacionTicket.frx":2BFC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Siguiente registro"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdUltimo 
         Height          =   480
         Left            =   2025
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrizacionTicket.frx":2D6E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ultimo registro"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdGrabar 
         Height          =   480
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrizacionTicket.frx":2EE0
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Guardar el registro"
         Top             =   135
         Width           =   495
      End
      Begin VB.CommandButton cmdEliminar 
         Height          =   480
         Left            =   3015
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrizacionTicket.frx":3222
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Borrar el registro"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Datos generales del ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7455
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2300
         TabIndex        =   33
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Frame Frame1 
         Height          =   23
         Left            =   0
         TabIndex        =   32
         Top             =   1920
         Width           =   7455
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   5930
         TabIndex        =   27
         Top             =   1135
         Width           =   23
      End
      Begin VB.Frame Frame2 
         Height          =   23
         Left            =   0
         TabIndex        =   26
         Top             =   1080
         Width           =   7455
      End
      Begin VB.CommandButton cmdGeneraGrids 
         Caption         =   "&Generar"
         Height          =   355
         Left            =   6120
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtRowsPie 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Text            =   "1"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtRowsCuerpo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Text            =   "1"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtRowsEncabezado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Text            =   "1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtClave 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   6015
      End
      Begin VB.TextBox txtCols 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Text            =   "1"
         Top             =   1200
         Width           =   375
      End
      Begin VSFlex7LCtl.VSFlexGrid grdEncabezado 
         Height          =   1425
         Left            =   120
         TabIndex        =   7
         Top             =   2260
         Width           =   7215
         _cx             =   12726
         _cy             =   2514
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         BackColorSel    =   -2147483645
         ForeColorSel    =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   14737632
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   250
         ColWidthMax     =   250
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdCuerpo 
         Height          =   1425
         Left            =   120
         TabIndex        =   8
         Top             =   3985
         Width           =   7215
         _cx             =   12726
         _cy             =   2514
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         BackColorSel    =   -2147483645
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   14737632
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   250
         ColWidthMax     =   250
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPie 
         Height          =   1425
         Left            =   120
         TabIndex        =   9
         Top             =   5680
         Width           =   7215
         _cx             =   12726
         _cy             =   2514
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         BackColorSel    =   -2147483645
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   14737632
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   250
         ColWidthMax     =   250
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   2300
         TabIndex        =   34
         Top             =   1550
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   5540
         TabIndex        =   35
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   285
         Left            =   5540
         TabIndex        =   36
         Top             =   1560
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   5460
         Width           =   7215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cuerpo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   3740
         Width           =   7215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Encabezado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2020
         Width           =   7215
      End
      Begin VB.Label Label5 
         Caption         =   "Renglones de pie de ticket"
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Renglones del cuerpo"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Renglones del encabezado"
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   1245
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Clave"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Columnas del ticket"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1245
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos insertables"
      Height          =   7215
      Left            =   7680
      TabIndex        =   18
      Top             =   120
      Width           =   4160
      Begin VSFlex7LCtl.VSFlexGrid grdDatos 
         Height          =   6735
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3915
         _cx             =   6906
         _cy             =   11880
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         BackColorSel    =   -2147483645
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   14737632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmParametrizacionTicket.frx":33C4
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdControl 
            Height          =   2175
            Left            =   0
            TabIndex        =   37
            Top             =   4560
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   0
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   7
         End
      End
   End
End
Attribute VB_Name = "frmParametrizacionTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const cgColorDatoInsertable = &HC0FFFF '[  Amarillo tenue  ]
Const cgValorBlanco = "×" '[  Valor que se insertará en la tabla para simbolizar un espacio (" ")  ]
Private WithEvents vgrsFormatoTicket As ADODB.Recordset '[  RecordSet general para el mantenimiento  ]
Attribute vgrsFormatoTicket.VB_VarHelpID = -1
Private Enum Status '[  Conjunto de estados posibles en el mantenimiento  ]
    stNuevo = 0
    stedicion = 1
    stConsulta = 2
    stPresentacionPreliminar = 3
    stEspera = 4
End Enum
Private vgstsStatus As Status '[  Indica el estado en el que está el mantenimiento  ]
Dim vlblnLicenciaIEPS As Boolean
Private Sub chkPredeterminado_Click()
    pChange
End Sub
Private Sub chkPresentacionPreliminar_Click()
    If chkPresentacionPreliminar.Value Then
        pbxCveTicket.Visible = True
        txtCveTicket.Text = ""
        pHabilitaBotones stPresentacionPreliminar
        txtClave.Enabled = False
        txtCveTicket.SetFocus
    Else
        pbxCveTicket.Visible = False
        txtPresentacionPreliminar.Visible = False
        txtPresentacionPreliminar.Text = ""
        txtClave.Enabled = True
        pHabilitaBotones stConsulta
    End If
End Sub

Private Sub cmdAnterior_Click()
    
    vgrsFormatoTicket.MovePrevious
    If vgrsFormatoTicket.BOF Then
        vgrsFormatoTicket.MoveNext
    End If
    txtClave.Text = vgrsFormatoTicket!intCveFormatoTicket
    txtClave_KeyDown vbKeyReturn, 0
    pHabilita 1, 1, 1, 1, 1
        
End Sub


Private Sub cmdBuscar_Click()
    If fblnExistenFormatos Then
        frmBusquedaFormatoTicket.Show vbModal
        If frmBusquedaFormatoTicket.vgblnBuscar Then
            txtClave.Text = frmBusquedaFormatoTicket.grdBusqueda.TextMatrix(frmBusquedaFormatoTicket.grdBusqueda.Row, 0)
            Unload frmBusquedaFormatoTicket
            txtClave_KeyDown vbKeyReturn, 0
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim vlstrSentencia As String
    Dim vlrsFormato As New ADODB.Recordset
    
    On Error GoTo NotificaError
    '[  ¿Está seguro de eliminar los datos?  ]
    If MsgBox(SIHOMsg(6), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
        Set vlrsFormato = frsRegresaRs("Select Count(*) As Co From pvTicketDepartamento Where intCveFormatoTicket = " & vgrsFormatoTicket!intCveFormatoTicket, adLockOptimistic, adOpenDynamic)
        If vlrsFormato!CO > 0 Then
            '[  !No se pueden borrar los datos!  ]
            MsgBox SIHOMsg(257), vbCritical, "Mensaje"
        Else
            EntornoSIHO.ConeccionSIHO.BeginTrans
            pEjecutaSentencia "Delete From pvTicketDepartamento Where intCveFormatoTicket = " & vgrsFormatoTicket!intCveFormatoTicket & " And smiDepartamento = " & CStr(vgintNumeroDepartamento)
            pEjecutaSentencia "Delete From pvDetalleFormatoTicket Where intCveFormatoTicket = " & vgrsFormatoTicket!intCveFormatoTicket
            pEjecutaSentencia "Delete From pvFormatoTicket Where intCveFormatoTicket = " & vgrsFormatoTicket!intCveFormatoTicket
            EntornoSIHO.ConeccionSIHO.CommitTrans
            vgrsFormatoTicket.Requery
            If vgrsFormatoTicket.RecordCount > 0 Then
                txtClave.Text = vgrsFormatoTicket!intCveFormatoTicket
                txtClave_KeyDown vbKeyReturn, 0
            Else
                txtClave.SetFocus
            End If
        End If
    End If
    Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEliminar_Click"))
End Sub

Private Sub cmdGeneraGrids_Click()
    txtCols.Text = Val(txtCols.Text)
    txtRowsEncabezado.Text = Val(txtRowsEncabezado.Text)
    txtRowsCuerpo.Text = Val(txtRowsCuerpo.Text)
    txtRowsPie.Text = Val(txtRowsPie.Text)
    If CInt(txtRowsCuerpo.Text) > 0 Then
        If fblnValidaColumnas Then
            pChange
            '[  Establece el número de columnas y renglones especificado en los siguientes grids  ]
            pGeneraGrid grdEncabezado, CInt(txtRowsEncabezado.Text)
            pGeneraGrid grdCuerpo, CInt(txtRowsCuerpo.Text)
            pGeneraGrid grdPie, CInt(txtRowsPie.Text)
        End If
    Else
        '[  No se puede realizar la operación con cantidad cero o menor que cero  ]
        MsgBox SIHOMsg(651), vbCritical, "Mensaje"
        txtRowsCuerpo.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim vlrsFormato As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlintCol As Integer
    Dim vlintRow As Integer
    Dim vlintGrids As Integer
    Dim vlgrdGrid As VSFlexGrid
    
    On Error GoTo NotificaError

    If pDatosValidos Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
        '-------------------------------------------------------------------
        '««  Inserta o actualiza el formato del ticket (pvFormatoTicket)  »»
        '-------------------------------------------------------------------
        vlstrSentencia = "Select * " & _
                         "From PVFORMATOTICKET " & _
                         "Where INTCVEFORMATOTICKET = " & IIf(vgstsStatus = stNuevo, "-1", txtClave.Text)
        Set vlrsFormato = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        If vgstsStatus = stNuevo Then vlrsFormato.AddNew
        vlrsFormato!vchDescripcion = txtDescripcion.Text
        vlrsFormato!intColumnas = grdEncabezado.Cols
        vlrsFormato!intRenglonesEncabezado = grdEncabezado.Rows
        vlrsFormato!intRenglonesCuerpo = grdCuerpo.Rows
        vlrsFormato!intRenglonesPie = grdPie.Rows
        vlrsFormato.Update
        If vgstsStatus = stNuevo Then txtClave.Text = flngObtieneIdentity("SEC_PVFORMATOTICKET", vlrsFormato!intCveFormatoTicket)
        '-------------------------------------------------------------------------------------
        '««  Inserta o actualiza el detalle del formato del ticket (pvDetalleFormatoTicket) »»
        '-------------------------------------------------------------------------------------
        pEjecutaSentencia "Delete from PvDetalleFormatoTicket Where intCveFormatoTicket = " & txtClave.Text
        For vlintGrids = 0 To 2
            '[  Selecciona el grid que se evaluará  ]
            Select Case vlintGrids
                    Case 0
                        Set vlgrdGrid = grdEncabezado
                    Case 1
                        Set vlgrdGrid = grdCuerpo
                    Case 2
                        Set vlgrdGrid = grdPie
            End Select
            '-----------------------------------------------------
            '««  Recorre el grid para grabar los valores fijos  »»
            '-----------------------------------------------------
            For vlintRow = 0 To vlgrdGrid.Rows - 1
                vlgrdGrid.Row = vlintRow
                For vlintCol = 0 To vlgrdGrid.Cols - 1
                    vlgrdGrid.Col = vlintCol
                    '[  Si no es un dato insertable  ]
                    If vlgrdGrid.CellBackColor <> cgColorDatoInsertable Then
                            pEjecutaSentencia "Insert into PvDetalleFormatoTicket ( " & _
                                              "intCveFormatoTicket, " & _
                                              "vchTipoValor, " & _
                                              "intRenglon, " & _
                                              "intColumna, " & _
                                              "intLonguitud, " & _
                                              "vchValor, " & _
                                              "vchSeccion) Values (" & _
                                              txtClave.Text & ", " & _
                                              "'F', " & _
                                              vlintRow & ", " & _
                                              vlintCol & ", " & _
                                              "1, " & _
                                              "'" & IIf(Trim(vlgrdGrid.TextMatrix(vlintRow, vlintCol)) = "", cgValorBlanco, vlgrdGrid.TextMatrix(vlintRow, vlintCol)) & "', " & _
                                              IIf(vlintGrids = 0, "'E'", IIf(vlintGrids = 1, "'C'", "'P'")) & ")"
                    End If
                Next
            Next
            '--------------------------------------------------------------------
            '««  Recorre el grid de control para grabar los valores variables  »»
            '--------------------------------------------------------------------
            For vlintRow = 0 To grdControl.Rows - 1
                If (vlintGrids = 0 And grdControl.TextMatrix(vlintRow, 5) = "E") Or _
                   (vlintGrids = 1 And grdControl.TextMatrix(vlintRow, 5) = "C") Or _
                   (vlintGrids = 2 And grdControl.TextMatrix(vlintRow, 5) = "P") Then
                    grdDatos.Row = fintRenglonDeClave(grdControl.TextMatrix(vlintRow, 4))
                    pEjecutaSentencia "Insert into PvDetalleFormatoTicket ( " & _
                                      "intCveFormatoTicket, " & _
                                      "vchTipoValor, " & _
                                      "intRenglon, " & _
                                      "intColumna, " & _
                                      "intLonguitud, " & _
                                      "vchValor, " & _
                                      "vchSeccion, " & _
                                      "intTipoDato, " & _
                                      "intLongMax) Values (" & _
                                      txtClave.Text & ", " & _
                                      "'V', " & _
                                      grdControl.TextMatrix(vlintRow, 1) & ", " & _
                                      grdControl.TextMatrix(vlintRow, 2) & ", " & _
                                      grdControl.TextMatrix(vlintRow, 3) & ", '" & _
                                      grdControl.TextMatrix(vlintRow, 0) & "', '" & _
                                      grdControl.TextMatrix(vlintRow, 5) & "', " & _
                                      grdDatos.TextMatrix(grdDatos.Row, 3) & ", " & _
                                      grdControl.TextMatrix(vlintRow, 6) & ")"
                End If
            Next
        Next
        EntornoSIHO.ConeccionSIHO.CommitTrans
        vgrsFormatoTicket.Requery
        vgrsFormatoTicket.Find "intCveFormatoTicket = " & txtClave.Text
        txtClave_KeyDown vbKeyReturn, 0
    End If
    Exit Sub

NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabar_Click"))
End Sub

Private Sub cmdOcultarVP_Click()
End Sub

Private Sub cmdPrimer_Click()
    vgrsFormatoTicket.MoveFirst
    txtClave.Text = vgrsFormatoTicket!intCveFormatoTicket
    txtClave_KeyDown vbKeyReturn, 0
    pHabilita 1, 1, 1, 1, 1

End Sub

Private Sub cmdSiguiente_Click()
    If Not vgrsFormatoTicket.EOF Then
        vgrsFormatoTicket.MoveNext
    End If
    If vgrsFormatoTicket.EOF Then
        vgrsFormatoTicket.MovePrevious
    End If
    txtClave.Text = vgrsFormatoTicket!intCveFormatoTicket
    txtClave_KeyDown vbKeyReturn, 0
    pHabilita 1, 1, 1, 1, 1
End Sub

Private Sub cmdUltimo_Click()
    vgrsFormatoTicket.MoveLast
    txtClave.Text = vgrsFormatoTicket!intCveFormatoTicket
    txtClave_KeyDown vbKeyReturn, 0
    pHabilita 1, 1, 1, 1, 1

End Sub

Private Sub Command1_Click()
        
End Sub

Private Sub cmdVistaPreliminar_Click()
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If chkPresentacionPreliminar.Value = 0 Then
                Unload Me
                KeyCode = 0
            Else
                pbxCveTicket.Visible = False
                chkPresentacionPreliminar.Value = 0
            End If
        Case vbKeyReturn
            If chkPresentacionPreliminar.Value = 0 And Me.ActiveControl.Name <> "grdDatos" Then SendKeys vbTab
    End Select
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmMenuPrincipal.Icon
    vlblnLicenciaIEPS = fblLicenciaIEPS
        
    Set vgrsFormatoTicket = frsRegresaRs("Select intCveFormatoTicket, vchDescripcion From PvFormatoTicket Order by intCveFormatoTicket", adLockOptimistic, adOpenDynamic)
    '[  Establece los valores que se podrán seleccionar en la columna "Tipo de dato"  ]
    grdDatos.ColComboList(3) = "#0;Cadena|#1;Fecha|#2;Moneda|#3;Número|#4;Hora"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If vgstsStatus = stPresentacionPreliminar Then
        Cancel = True
        chkPresentacionPreliminar.Value = 0
    Else
        If vgstsStatus <> stConsulta And vgstsStatus <> stEspera Then
            Cancel = True
            '[  ¿Desea abandonar la operación?  ]
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                txtClave.SetFocus
            End If
        End If
    End If
End Sub

Private Sub grdCuerpo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grdCuerpo.CellBackColor = cgColorDatoInsertable Then
        Cancel = True
        pPonLongMax grdCuerpo
        grdDatos.Row = fintRenglonDeClave(grdCuerpo.TextMatrix(grdCuerpo.Row, grdCuerpo.Col))
    Else
        grdCuerpo.EditMaxLength = 1
    End If
End Sub

Private Sub grdCuerpo_CellChanged(ByVal Row As Long, ByVal Col As Long)
    pChange
End Sub

Private Sub grdCuerpo_KeyDown(KeyCode As Integer, Shift As Integer)
    pKeyDown grdCuerpo, KeyCode
End Sub

Private Sub grdCuerpo_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If grdCuerpo.CellBackColor = cgColorDatoInsertable Then KeyCode = 0
End Sub

Private Sub grdCuerpo_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode <> vbKeyEscape Then
    pKeyUpEdit grdCuerpo
End If
End Sub

Private Sub grdCuerpo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pMouseUp grdCuerpo, Button
End Sub


Private Sub grdDatos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    pChange
End Sub

Private Sub grdDatos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grdDatos.Col = 1 Then
        Cancel = True
    End If
End Sub

Private Sub grdDatos_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If grdDatos.Col < grdDatos.Cols - 1 Then
            grdDatos.Col = grdDatos.Col + 1
        Else
            If grdDatos.Row < grdDatos.Rows - 1 Then
                grdDatos.Col = 2
                grdDatos.Row = grdDatos.Row + 1
            End If
        End If
    End If
End Sub

Private Sub grdEncabezado_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grdEncabezado.CellBackColor = cgColorDatoInsertable Then
        grdDatos.Row = fintRenglonDeClave(grdEncabezado.TextMatrix(grdEncabezado.Row, grdEncabezado.Col))
        pPonLongMax grdEncabezado
        Cancel = True
    Else
        grdEncabezado.EditMaxLength = 1
    End If
End Sub

Private Sub grdEncabezado_CellChanged(ByVal Row As Long, ByVal Col As Long)
    pChange
End Sub

Private Sub grdEncabezado_KeyDown(KeyCode As Integer, Shift As Integer)
    pKeyDown grdEncabezado, KeyCode
End Sub

Private Sub grdEncabezado_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If grdEncabezado.CellBackColor = cgColorDatoInsertable Then KeyCode = 0
End Sub

Private Sub grdEncabezado_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    pKeyUpEdit grdEncabezado
End Sub


Private Sub grdEncabezado_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pMouseUp grdEncabezado, Button
End Sub

Private Sub pGeneraGrid(pgrdGrid As VSFlexGrid, pstrRenglones As String)
    Dim vlintCont As Integer
    
    If fblnSePuedeGenerar(pgrdGrid, pstrRenglones) Then
        With pgrdGrid
            .Cols = txtCols.Text
            .Rows = CInt(pstrRenglones)
            For vlintCont = 0 To .Cols - 1
                .ColWidth(vlintCont) = 250
            Next
        End With
    End If
End Sub




Private Sub grdPie_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grdPie.CellBackColor = cgColorDatoInsertable Then
        Cancel = True
        pPonLongMax grdPie
        grdDatos.Row = fintRenglonDeClave(grdPie.TextMatrix(grdPie.Row, grdPie.Col))
    Else
        grdPie.EditMaxLength = 1
    End If
End Sub

Sub pMouseUp(pgrdGrid As VSFlexGrid, Button)
    
    Dim vlintCont As Integer
    Dim vlblnBand As Boolean
    
    '[  Se dió click en el botón secundario del mouse y el mouse está posicionado en alguna cuadrícula del grid  ]
    If Button = 2 And pgrdGrid.MouseRow <> -1 And pgrdGrid.MouseCol <> -1 Then
        With pgrdGrid
            .SetFocus
            .Row = .MouseRow
            .Col = .MouseCol
            '-----------------------------------------------------------------------
            '««  Validación de si cabe el dato que se va a insertar en el grid y  »»
            '««  si no sobreescribe nada                                          »»
            '-----------------------------------------------------------------------
            vlblnBand = True
            If .Col + CInt(grdDatos.TextMatrix(grdDatos.Row, 2)) <= .Cols Then
                For vlintCont = .Col To .Col + CInt(grdDatos.TextMatrix(grdDatos.Row, 2)) - 1
                    If .TextMatrix(.Row, vlintCont) <> "" Then
                        vlblnBand = False
                        Exit For
                    End If
                Next
            Else
                vlblnBand = False
            End If
            '[  Si cabe el dato que se va a insertar en el grid y si no sobreescribe nada  ]
            If vlblnBand Then
                '---------------------------------------------------------------------------------------
                '««  Pone el grid de datos color amarillo para indicar los datos que se están usando  »»
                '---------------------------------------------------------------------------------------
                For vlintCont = 1 To grdDatos.Cols - 1
                    grdDatos.Col = vlintCont
                    grdDatos.CellBackColor = cgColorDatoInsertable
                Next
                
                '-----------------------------------------
                '««  Almacena datos en grid de control  »»
                '-----------------------------------------
                With grdControl
                    .Rows = grdControl.Rows + 1
                    '[  Nombre del dato  ]
                    .TextMatrix(.Rows - 1, 0) = grdDatos.TextMatrix(grdDatos.Row, 1)
                    '[  Renglón  ]
                    .TextMatrix(.Rows - 1, 1) = pgrdGrid.Row
                    '[  Columna  ]
                    .TextMatrix(.Rows - 1, 2) = pgrdGrid.Col
                    '[  Longuitud ]
                    .TextMatrix(.Rows - 1, 3) = grdDatos.TextMatrix(grdDatos.Row, 2)
                    '[  Número del dato  ]
                    .TextMatrix(.Rows - 1, 4) = grdDatos.TextMatrix(grdDatos.Row, 0)
                    '[  Grid donde se insertó  ]
                    .TextMatrix(.Rows - 1, 5) = IIf(pgrdGrid.Name = "grdEncabezado", "E", IIf(pgrdGrid.Name = "grdCuerpo", "C", "P"))
                    '[  Especifica si el campo puede crecer  ]
                    .TextMatrix(.Rows - 1, 6) = grdDatos.TextMatrix(grdDatos.Row, 4)
                End With
                '-------------------------------------------------------------
                '««  Muestra el dato insertable en el grid correspondiente  »»
                '-------------------------------------------------------------
                For vlintCont = .Col To .Col + CInt(grdDatos.TextMatrix(grdDatos.Row, 2)) - 1
                    .Col = vlintCont
                    .CellBackColor = cgColorDatoInsertable
                    .CellFontBold = True
                    If grdDatos.TextMatrix(grdDatos.Row, 4) <> 0 Then
                        .CellFontUnderline = True
                    Else
                        .CellFontUnderline = False
                    End If
                    .TextMatrix(.Row, vlintCont) = grdDatos.TextMatrix(grdDatos.Row, 0)
                Next
            Else
                '[  No es posible insertar el dato porque excede el límite establecido de columnas o porque sobreescribiría información.  ]
                MsgBox SIHOMsg(663), vbCritical, "Mensaje"
            End If
        End With
    End If
End Sub

Private Sub grdPie_CellChanged(ByVal Row As Long, ByVal Col As Long)
    pChange
End Sub

Private Sub grdPie_KeyDown(KeyCode As Integer, Shift As Integer)
    pKeyDown grdPie, KeyCode
End Sub

Private Sub grdPie_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If grdPie.CellBackColor = cgColorDatoInsertable Then KeyCode = 0
End Sub

Private Sub grdPie_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    pKeyUpEdit grdPie
End Sub

Private Sub grdPie_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pMouseUp grdPie, Button
End Sub

Sub pKeyUpEdit(pgrdGrid As VSFlexGrid)
    With pgrdGrid
        .CellBackColor = vbWindowBackground
        .CellFontBold = False
        .CellFontUnderline = False

        If Len(.EditText) >= 1 Then
            If .Col < (.Cols - 1) Then
                .Col = (.Col + 1)
                .SetFocus
            Else
                If .Row < (.Rows - 1) Then
                    .Col = 0
                    .Row = .Row + 1
                    .SetFocus
                Else
                    Me.SetFocus
                    pgrdGrid.SetFocus
                End If
            End If
        End If
    End With
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtClave_GotFocus()
    pLimpiaForma
    txtClave.Text = fintSigNumRs(vgrsFormatoTicket, 0)
    pHabilitaBotones stEspera
    '[  Valida si no existe ningun registro para deshabilitar todo  ]
    If vgrsFormatoTicket.RecordCount = 0 Then
        pHabilita 0, 0, 0, 0, 0
    End If
End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
    Dim vlrsFormato As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        '-------------------------------------------------------------
        '««  Busca si existe el formato según la clave introducida  »»
        '-------------------------------------------------------------
        vlstrSentencia = "Select * From pvFormatoTicket Where intCveFormatoTicket = " & txtClave.Text
        Set vlrsFormato = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        If vlrsFormato.RecordCount > 0 Then
            pLimpiaForma
            txtClave.Text = vlrsFormato!intCveFormatoTicket
            txtDescripcion.Text = vlrsFormato!vchDescripcion
            txtCols.Text = vlrsFormato!intColumnas
            txtRowsEncabezado.Text = vlrsFormato!intRenglonesEncabezado
            txtRowsCuerpo.Text = vlrsFormato!intRenglonesCuerpo
            txtRowsPie.Text = vlrsFormato!intRenglonesPie
            cmdGeneraGrids_Click
            pCargaDetalle txtClave.Text
            pHabilitaBotones stConsulta
        Else
            pNuevoRegistro
        End If
    End If
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtCols_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCveTicket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If RTrim(txtCveTicket.Text) <> "" Then
            pbxCveTicket.Visible = False
            pHabilitaBotones stPresentacionPreliminar
            pImprimeTicket txtCveTicket.Text
        Else
            MsgBox "Introduzca una clave de ticket.", vbCritical, "Mensaje"
            txtCveTicket.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then
        chkPresentacionPreliminar.Value = 0
        chkPresentacionPreliminar_Click
    End If
End Sub

Private Sub txtCveTicket_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtCveTicket_LostFocus()
    If txtCveTicket.Visible Then txtCveTicket.SetFocus
End Sub

Private Sub txtDescripcion_Change()
    pChange
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRowsCuerpo_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRowsEncabezado_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRowsPie_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
    
Private Function fblnSePuedeGenerar(pgrdGrid As VSFlexGrid, pstrRenglones As String) As Boolean
    Dim vlintCol As Integer
    Dim vlintRow As Integer
    Dim vlintGrids As Integer
    Dim vlgrdGrid As VSFlexGrid
    Dim vlintNumRows As Integer
    
    fblnSePuedeGenerar = False
    '[  Valida que los parámetros sean mayores a cero  ]
    If CInt(pstrRenglones) >= 0 Then
       '[  Si se estrablece un mayor número de renglones  ]
       If CInt(pstrRenglones) > CInt(pgrdGrid.Rows) Then
            fblnSePuedeGenerar = True
       Else
            For vlintRow = CInt(pstrRenglones) To pgrdGrid.Rows - 1
                For vlintCol = 0 To pgrdGrid.Cols - 1
                    If pgrdGrid.TextMatrix(vlintRow, vlintCol) <> "" Then
                        '[  No se puede establecer un número menor de renglones porque ocasionaría una pérdida de información  ]
                        MsgBox SIHOMsg(664), vbCritical, "Mensaje"
                        pgrdGrid.SetFocus
                        Exit Function
                    End If
                Next
            Next
            fblnSePuedeGenerar = True
       End If
    Else
        '[  No se puede realizar la operación con cantidad cero o menor que cero  ]
        MsgBox SIHOMsg(651), vbCritical, "Mensaje"
        If pgrdGrid.Name = "grdEncabezado" Then
            txtRowsEncabezado.SetFocus
        Else
            If pgrdGrid.Name = "grdCuerpo" Then
                txtRowsCuerpo.SetFocus
            Else
                If pgrdGrid.Name = "grdPie" Then
                    txtRowsPie.SetFocus
                End If
            End If
        End If
    End If
End Function

Private Function fblnValidaColumnas() As Boolean
    Dim vlintCol As Integer
    Dim vlintRow As Integer
    Dim vlintGrids As Integer
    Dim vlgrdGrid As VSFlexGrid
    
    fblnValidaColumnas = False
    '[  Valida que el número de columnas sea mayor a cero  ]
    If CInt(txtCols.Text) > 0 Then
       '[  Si se establece un mayor número de columnas  ]
       If CInt(txtCols.Text) > CInt(grdEncabezado.Cols) Then
            fblnValidaColumnas = True
       Else
            '--------------------------------------------------------------
            '««  Valida que no se tenga información que se pueda perder  »»
            '--------------------------------------------------------------
            For vlintGrids = 0 To 2
                '[  Selecciona el grid que se evaluará  ]
                Select Case vlintGrids
                        Case 0
                            Set vlgrdGrid = grdEncabezado
                        Case 1
                            Set vlgrdGrid = grdCuerpo
                        Case 2
                            Set vlgrdGrid = grdPie
                End Select
                '-----------------------------------------------
                '««  Recorre el grid en busca de información  »»
                '-----------------------------------------------
                For vlintRow = 0 To vlgrdGrid.Rows - 1
                    For vlintCol = CInt(txtCols.Text) To vlgrdGrid.Cols - 1
                        If vlgrdGrid.TextMatrix(vlintRow, vlintCol) <> "" Then
                            '[  No se puede establecer un número menor de columnas porque ocasionaría una pérdida de información  ]
                            MsgBox SIHOMsg(665), vbCritical, "Mensaje"
                            vlgrdGrid.SetFocus
                            Exit Function
                        End If
                    Next
                Next
            Next
            fblnValidaColumnas = True
       End If
    Else
        '[  No se puede realizar la operación con cantidad cero o menor que cero  ]
        MsgBox SIHOMsg(651), vbCritical, "Mensaje"
        txtCols.SetFocus
    End If
End Function

Private Sub UpDown1_UpClick()
    txtCols.Text = CStr(CInt(txtCols.Text) + 1)
End Sub

Private Sub UpDown1_DownClick()
    If CInt(txtCols.Text) - 1 >= 0 Then txtCols.Text = CStr(CInt(txtCols.Text) - 1)
End Sub

Private Sub UpDown2_UpClick()
    txtRowsCuerpo.Text = CStr(CInt(txtRowsCuerpo.Text) + 1)
End Sub

Private Sub UpDown2_DownClick()
    If CInt(txtRowsCuerpo.Text) - 1 >= 0 Then txtRowsCuerpo.Text = CStr(CInt(txtRowsCuerpo.Text) - 1)
End Sub

Private Sub UpDown3_UpClick()
    txtRowsEncabezado.Text = CStr(CInt(txtRowsEncabezado.Text) + 1)
End Sub

Private Sub UpDown3_DownClick()
    If CInt(txtRowsEncabezado.Text) - 1 >= 0 Then txtRowsEncabezado.Text = CStr(CInt(txtRowsEncabezado.Text) - 1)
End Sub

Private Sub UpDown4_UpClick()
    txtRowsPie.Text = CStr(CInt(txtRowsPie.Text) + 1)
End Sub

Private Sub UpDown4_DownClick()
    If CInt(txtRowsPie.Text) - 1 >= 0 Then txtRowsPie.Text = CStr(CInt(txtRowsPie.Text) - 1)
End Sub

Private Sub pKeyDown(pgrdGrid As VSFlexGrid, pintKeyCode As Integer)
    Dim vlintCol As Integer
    Dim vlintCont As Integer
    Dim vlintPosOriginal As Integer
    Dim vlblnExiste As Boolean
    
    '[  Si pulsó la tecla "Delete"  ]
    If pintKeyCode = 46 Then
        '[  Si no es parte de un dato insertable  ]
        If pgrdGrid.CellBackColor <> cgColorDatoInsertable Then
            pgrdGrid.TextMatrix(pgrdGrid.Row, pgrdGrid.Col) = ""
        Else
            '[  ¿Está seguro de eliminar los datos?  ]
            If MsgBox(SIHOMsg(6), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                vlintPosOriginal = pgrdGrid.Col
                '---------------------------------------------------------------
                '««  Localiza renglón que será eliminado del grid de control  »»
                '---------------------------------------------------------------
                lstCandidatos.Clear
                For vlintCont = 0 To grdControl.Rows - 1
                    '[  Si es el mismo dato, el mismo renglón y queda dentro del rango establecido  ]
                    If grdControl.TextMatrix(vlintCont, 4) = pgrdGrid.TextMatrix(pgrdGrid.Row, pgrdGrid.Col) And _
                       grdControl.TextMatrix(vlintCont, 1) = pgrdGrid.Row And _
                       (grdControl.TextMatrix(vlintCont, 2) <= pgrdGrid.Col And _
                        pgrdGrid.Col < CStr(CInt(grdControl.TextMatrix(vlintCont, 2)) + CInt(grdControl.TextMatrix(vlintCont, 3)))) Then
                        '|  Dado que ésta condición puede ser cumplida por varios registros,       |
                        '|  guardo las diferencias de la posición actual menos el inicio del       |
                        '|  dato en el registro, para posteriormente utilizar la menor diferencia  |
                        lstCandidatos.AddItem pgrdGrid.Col - CInt(grdControl.TextMatrix(vlintCont, 2))
                        lstCandidatos.ItemData(lstCandidatos.newIndex) = vlintCont
                    End If
                Next
                '------------------------------------------------
                '««  Elimina el dato del grid correspondiente  »»
                '------------------------------------------------
                vlintCont = lstCandidatos.ItemData(0)
                pgrdGrid.Col = CInt(grdControl.TextMatrix(vlintCont, 2))
                For vlintCol = CInt(grdControl.TextMatrix(vlintCont, 2)) To (CInt(grdControl.TextMatrix(vlintCont, 2)) + CInt(grdControl.TextMatrix(vlintCont, 3))) - 1
                    pgrdGrid.CellBackColor = vbWindowBackground
                    pgrdGrid.CellFontBold = False
                    pgrdGrid.CellFontUnderline = False
                    pgrdGrid.TextMatrix(pgrdGrid.Row, pgrdGrid.Col) = ""
                    If pgrdGrid.Col < pgrdGrid.Cols - 1 Then pgrdGrid.Col = pgrdGrid.Col + 1
                Next
                '-----------------------------------------------
                '««  Elimina el registro del grid de control  »»
                '-----------------------------------------------
                If grdControl.Rows > 1 Then
                    grdControl.RemoveItem (vlintCont)
                Else
                    grdControl.Rows = 0
                End If
                '-------------------------------------------------------------------------------------
                '««  Si ya no hay mas datos insertados, quita el color del fondo del grid de datos  »»
                '-------------------------------------------------------------------------------------
                vlblnExiste = False
                For vlintCont = 0 To grdControl.Rows - 1
                    If grdControl.TextMatrix(vlintCont, 0) = grdDatos.TextMatrix(grdDatos.Row, 1) Then vlblnExiste = True
                Next
                If Not vlblnExiste Then
                    For vlintCont = 1 To grdDatos.Cols - 1
                        grdDatos.Col = vlintCont
                        grdDatos.CellBackColor = vbWindowBackground
                    Next
                End If
                                
                pgrdGrid.Col = vlintPosOriginal
            End If
        End If
    End If
End Sub

Private Sub pNuevoRegistro()
    Dim vlrsDatos As New ADODB.Recordset
    Dim vlintRows As Integer
    
    pLimpiaForma
    txtClave.Text = fintSigNumRs(vgrsFormatoTicket, 0)
    pHabilitaBotones stNuevo
    
End Sub

Private Sub pHabilitaBotones(pstStatus As Status)
    Select Case pstStatus
        Case 0, 1 '[  Nuevo, Edición  ]
            cmdPrimer.Enabled = False
            cmdAnterior.Enabled = False
            cmdBuscar.Enabled = False
            cmdSiguiente.Enabled = False
            cmdUltimo.Enabled = False
            cmdGrabar.Enabled = True
            cmdEliminar.Enabled = False
            chkPresentacionPreliminar.Enabled = False
        Case 2 '[  Consulta  ]
            If Not cmdPrimer.Enabled Then cmdPrimer.Enabled = True
            If Not cmdAnterior.Enabled Then cmdAnterior.Enabled = True
            If Not cmdSiguiente.Enabled Then cmdSiguiente.Enabled = True
            If Not cmdUltimo.Enabled Then cmdUltimo.Enabled = True
            cmdBuscar.Enabled = True
            cmdGrabar.Enabled = False
            cmdEliminar.Enabled = True
            chkPresentacionPreliminar.Enabled = True
        Case 3 '[  Presentación preliminar  ]
            cmdPrimer.Enabled = False
            cmdAnterior.Enabled = False
            cmdBuscar.Enabled = False
            cmdSiguiente.Enabled = False
            cmdUltimo.Enabled = False
            cmdGrabar.Enabled = False
            cmdEliminar.Enabled = False
            chkPresentacionPreliminar.Enabled = True
        Case 4 '[  Espera  ]
            If Not cmdPrimer.Enabled Then cmdPrimer.Enabled = True
            If Not cmdAnterior.Enabled Then cmdAnterior.Enabled = True
            If Not cmdSiguiente.Enabled Then cmdSiguiente.Enabled = True
            If Not cmdUltimo.Enabled Then cmdUltimo.Enabled = True
            cmdBuscar.Enabled = True
            cmdGrabar.Enabled = False
            cmdEliminar.Enabled = False
            chkPresentacionPreliminar.Enabled = False
    End Select
    vgstsStatus = pstStatus
End Sub

Private Function pDatosValidos() As Boolean
    Dim vlstrDatosFaltantes As String
    
    vlstrDatosFaltantes = ""
    pDatosValidos = True
    If txtDescripcion.Text = "" Then
        vlstrDatosFaltantes = "     - Descripción" & vbCrLf
        txtDescripcion.SetFocus
    End If
    If vlstrDatosFaltantes <> "" Then
        MsgBox "Datos faltantes: " & vbCrLf & vbCrLf & vlstrDatosFaltantes, vbCritical, "Mensaje"
        pDatosValidos = False
    End If
End Function

Private Sub pCargaDetalle(pstrClave As String)
    Dim vlstrSentencia As String
    Dim vlrsDetalleFormato As New ADODB.Recordset
    Dim vlintGrids As Integer
    Dim vlgrdGrid As VSFlexGrid
    Dim vlintRow As Integer
    Dim vlintCont As Integer
    
    '----------------------------------------------
    '««  Llena los grids con los datos grabados  »»
    '----------------------------------------------
    For vlintGrids = 0 To 2
        '[  Selecciona el grid que se llenará  ]
        Select Case vlintGrids
                Case 0
                    Set vlgrdGrid = grdEncabezado
                Case 1
                    Set vlgrdGrid = grdCuerpo
                Case 2
                    Set vlgrdGrid = grdPie
        End Select
        '-------------------------------------------------------
        '««  Llena el grid con los valores fijos y variables  »»
        '-------------------------------------------------------
        vlstrSentencia = "Select * " & _
                         "From pvDetalleFormatoTicket " & _
                         "Where intCveFormatoTicket = " & pstrClave & " And " & _
                         "      vchSeccion = '" & IIf(vlintGrids = 0, "E", IIf(vlintGrids = 1, "C", "P")) & "' " & _
                         " Order by vchSeccion, intRenglon, intColumna"
                         '"      vchTipoValor = 'F' " &
        Set vlrsDetalleFormato = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        Do While Not vlrsDetalleFormato.EOF
            If vlrsDetalleFormato!vchTipoValor = "F" Then '[  Valor fijo  ]
                If vlrsDetalleFormato!vchvalor <> cgValorBlanco Then vlgrdGrid.TextMatrix(vlrsDetalleFormato!intRenglon, vlrsDetalleFormato!intColumna) = vlrsDetalleFormato!vchvalor
            Else '[  Valor variable  ]
                '-----------------------------------------
                '««  Almacena datos en grid de control  »»
                '-----------------------------------------
                With grdControl
                    .Rows = grdControl.Rows + 1
                    '[  Nombre del dato  ]
                    .TextMatrix(.Rows - 1, 0) = vlrsDetalleFormato!vchvalor
                    '[  Renglón  ]
                    .TextMatrix(.Rows - 1, 1) = vlrsDetalleFormato!intRenglon
                    '[  Columna  ]
                    .TextMatrix(.Rows - 1, 2) = vlrsDetalleFormato!intColumna
                    '[  Longuitud ]
                    .TextMatrix(.Rows - 1, 3) = vlrsDetalleFormato!intLonguitud
                    '[  Número del dato  ]
                    .TextMatrix(.Rows - 1, 4) = fstrNumeroDato(vlrsDetalleFormato!vchvalor, vlintRow)
                    '[  Grid donde se insertó  ]
                    .TextMatrix(.Rows - 1, 5) = IIf(vlgrdGrid.Name = "grdEncabezado", "E", IIf(vlgrdGrid.Name = "grdCuerpo", "C", "P"))
                    '[  Longuitud máxima  de caracteres permitida en el campo insertable ]
                    .TextMatrix(.Rows - 1, 6) = IIf(IsNull(vlrsDetalleFormato!intLongMax), 0, vlrsDetalleFormato!intLongMax)
                End With
                grdDatos.TextMatrix(vlintRow, 3) = IIf(IsNull(vlrsDetalleFormato!intTipoDato), 0, vlrsDetalleFormato!intTipoDato)
                grdDatos.TextMatrix(vlintRow, 2) = vlrsDetalleFormato!intLonguitud
                grdDatos.TextMatrix(vlintRow, 4) = vlrsDetalleFormato!intLongMax
                grdDatos.Row = vlintRow
                For vlintCont = 1 To grdDatos.Cols - 1
                    grdDatos.Col = vlintCont
                    grdDatos.CellBackColor = cgColorDatoInsertable
                Next
                
                '-------------------------------------------------------------
                '««  Muestra el dato insertable en el grid correspondiente  »»
                '-------------------------------------------------------------
                vlgrdGrid.Row = vlrsDetalleFormato!intRenglon
                For vlintCont = vlrsDetalleFormato!intColumna To (vlrsDetalleFormato!intColumna + vlrsDetalleFormato!intLonguitud) - 1
                    vlgrdGrid.Col = vlintCont
                    vlgrdGrid.CellBackColor = cgColorDatoInsertable
                    vlgrdGrid.CellFontBold = True
                    If vlrsDetalleFormato!intLongMax > 0 Then
                        vlgrdGrid.CellFontUnderline = True
                    Else
                        vlgrdGrid.CellFontUnderline = False
                    End If
                    vlgrdGrid.TextMatrix(vlrsDetalleFormato!intRenglon, vlintCont) = grdDatos.TextMatrix(vlintRow, 0)
                Next
                                
            End If
            vlrsDetalleFormato.MoveNext
        Loop
    Next
    pHabilitaBotones stConsulta
    pEnfocaTextBox txtDescripcion
End Sub

Private Sub pLimpiaForma()
    Dim vlrsDatos As New ADODB.Recordset
    Dim vlintRows As Integer
    Dim vlintCont As Integer
    Dim vlintCols As Integer
    Dim vlblnDesgloseIEPS As Integer
    '-----------------------------------------------------------
    '««  Limpia toda la forma y reestablece valores iniciales »»
    '-----------------------------------------------------------
    txtClave.Text = ""
    txtDescripcion.Text = ""
    txtCols.Text = "1"
    txtRowsEncabezado.Text = "1"
    txtRowsCuerpo.Text = "1"
    txtRowsPie.Text = "1"
    grdEncabezado.Clear
    grdCuerpo.Clear
    grdPie.Clear
    grdControl.Clear
    grdControl.Rows = 0
    lstCandidatos.Clear
    grdEncabezado.Rows = 0
    grdCuerpo.Rows = 0
    grdPie.Rows = 0
    For vlintRows = 1 To grdDatos.Rows - 1
        For vlintCols = 1 To grdDatos.Cols - 1
            grdDatos.Col = vlintCols
            grdDatos.Row = vlintRows
            grdDatos.CellBackColor = vbWindowBackground
        Next
    Next
    '-----------------------------------------------------------
    '««  Carga los datos que se pueden incluir en el formato  »»
    '-----------------------------------------------------------
    vlblnDesgloseIEPS = fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0)
    Set vlrsDatos = frsEjecuta_SP("-1|" & vlblnDesgloseIEPS, "Sp_Pvselticket")
    grdDatos.Rows = vlrsDatos.Fields.Count + IIf(vlblnLicenciaIEPS And vlblnDesgloseIEPS = 1, 1, -1)
    vlintCont = 1
    For vlintRows = 1 To vlrsDatos.Fields.Count
     If (vlrsDatos.Fields(vlintRows - 1).Name = "IEPS" And vlblnLicenciaIEPS And vlblnDesgloseIEPS = 1) Or (vlrsDatos.Fields(vlintRows - 1).Name = "ImporteGravadoIEPS" And vlblnLicenciaIEPS And vlblnDesgloseIEPS = 1) Or _
        (vlrsDatos.Fields(vlintRows - 1).Name <> "IEPS" And vlrsDatos.Fields(vlintRows - 1).Name <> "ImporteGravadoIEPS") Then
         grdDatos.TextMatrix(vlintCont, 0) = vlintRows '[  Id  ]
         grdDatos.TextMatrix(vlintCont, 1) = vlrsDatos.Fields(vlintRows - 1).Name '[  Dato  ]
         grdDatos.TextMatrix(vlintCont, 2) = "1" '[  Long  ]
         Select Case vlrsDatos.Fields.Item(vlintRows - 1).Type '[  Tipo de dato  ]
            Case adChar, adBSTR, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar '[  Cadena  ]
                grdDatos.TextMatrix(vlintCont, 3) = "0"
            Case adDate, adDBDate, adDBTimeStamp '[  Fecha  ]
                grdDatos.TextMatrix(vlintCont, 3) = "1"
            Case adCurrency, adDecimal, adDouble, adSingle, adVarNumeric, adNumeric  '[  Moneda  ]
                grdDatos.TextMatrix(vlintCont, 3) = "2"
            Case adBigInt, adInteger, adTinyInt, adSmallInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt '[  Número  ]
                grdDatos.TextMatrix(vlintCont, 3) = "3"
            Case adDBTime '[  Hora  ]
                grdDatos.TextMatrix(vlintCont, 3) = "3"
         End Select
         grdDatos.TextMatrix(vlintCont, 4) = "0" '[  L. Max.  ]
         vlintCont = vlintCont + 1
     End If
    Next
    grdDatos.Row = 1
    vlrsDatos.Close
    
End Sub

'|  Función que sirve para saber cual es el número que le corresponde  |
'|  a un nombre de dato y su index en el grid de datos                 |
Private Function fstrNumeroDato(pstrNombreDato As String, pintIndice As Integer) As String
    Dim vlintCont As Integer
    
    fstrNumeroDato = ""
    For vlintCont = 0 To grdDatos.Rows - 1
        If grdDatos.TextMatrix(vlintCont, 1) = pstrNombreDato Then
            fstrNumeroDato = grdDatos.TextMatrix(vlintCont, 0)
            pintIndice = vlintCont
            Exit For
        End If
    Next

End Function

Private Sub pChange()
    If vgstsStatus = stConsulta Then pHabilitaBotones stedicion
End Sub

Private Function fintRenglonDeClave(pstrClave As String) As Integer
    Dim vlintCont As Integer
    If pstrClave = "" Then
        pstrClave = 1
    End If
    fintRenglonDeClave = 0
    For vlintCont = 1 To grdDatos.Rows
        If grdDatos.TextMatrix(vlintCont, 0) = pstrClave Then
            fintRenglonDeClave = vlintCont
            Exit For
        End If
    Next
End Function

Private Sub pPonLongMax(pgrdGrid As VSFlexGrid)
    Dim vlintCont As Integer
    
    '-------------------------------------------
    '««  Localiza renglón en grid de control  »»
    '-------------------------------------------
    lstCandidatos.Clear
    For vlintCont = 0 To grdControl.Rows - 1
        '[  Si es el mismo dato, el mismo renglón y queda dentro del rango establecido  ]
        If grdControl.TextMatrix(vlintCont, 4) = pgrdGrid.TextMatrix(pgrdGrid.Row, pgrdGrid.Col) And _
           grdControl.TextMatrix(vlintCont, 1) = pgrdGrid.Row And _
           (grdControl.TextMatrix(vlintCont, 2) <= pgrdGrid.Col And _
            pgrdGrid.Col < CStr(CInt(grdControl.TextMatrix(vlintCont, 2)) + CInt(grdControl.TextMatrix(vlintCont, 3)))) Then
            '|  Dado que ésta condición puede ser cumplida por varios registros,       |
            '|  guardo las diferencias de la posición actual menos el inicio del       |
            '|  dato en el registro, para posteriormente utilizar la menor diferencia  |
            lstCandidatos.AddItem pgrdGrid.Col - CInt(grdControl.TextMatrix(vlintCont, 2))
            lstCandidatos.ItemData(lstCandidatos.newIndex) = vlintCont
        End If
    Next
    '---------------------------------------------------------------
    '««  Pone la longuitud máxima que se estableció para el dato  »»
    '---------------------------------------------------------------
    vlintCont = lstCandidatos.ItemData(0)
    grdDatos.TextMatrix(grdDatos.Row, 4) = grdControl.TextMatrix(vlintCont, 6)
End Sub

Private Function fintObtieneFormato() As Integer
    Dim vlstrSentencia As String
    Dim vlrsFormato As New ADODB.Recordset

    fintObtieneFormato = -1
    '---------------------------------------
    '««  Busca un formato predeterminado  »»
    '---------------------------------------
    vlstrSentencia = "select intCveFormatoTicket from pvTicketDepartamento where smiDepartamento = " & CStr(vgintNumeroDepartamento)
    Set vlrsFormato = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If vlrsFormato.RecordCount > 0 Then
        fintObtieneFormato = vlrsFormato!intCveFormatoTicket
    End If
End Function

Private Sub pImprimeTicket(pstrCveTicket As String)
    Dim vlintCont As Integer       '[  Contador general  ]
    Dim vlintSecciones As Integer  '[  Contador de secciones  ]
    Dim vlintCveFormato As Integer '[  Clave del formato  ]
    Dim vlintRenglon As Integer    '[  Sirve para identificar el cambio de renglón  ]
    Dim vlintPosicion As Integer   '[  Posición del campo(field) en el RecordSet  ]
    Dim vlstrLinea As String       '[  Cadena que se va a imprimir  ]
    Dim vlstrSeccion As String     '[  Indica la sección, E = Encabezado, C = Cuerpo, P = Pie]
    Dim vlstrSentencia As String
    Dim vlrsValoresTicket As New ADODB.Recordset
    Dim vlrsFormatoTicket As New ADODB.Recordset
    
'    '[  Si existe un formato  ]
        Set vlrsValoresTicket = frsEjecuta_SP(pstrCveTicket & "|" & fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), "Sp_Pvselticket")
        '[  Si existe un ticket registrado con esa clave  ]
        If vlrsValoresTicket.RecordCount > 0 Then
            txtPresentacionPreliminar.Visible = True
            '[  Es un ciclo de tres vueltas porque son tres secciones (Encabezado, Cuerpo y Pie)  ]
            For vlintSecciones = 0 To 2
                Select Case vlintSecciones
                    Case 0 '[  Encabezado  ]
                        vlstrSeccion = "E"
                    Case 1 '[  Cuerpo  ]
                        vlstrSeccion = "C"
                    Case 2 '[  Pie  ]
                        vlstrSeccion = "P"
                End Select
                '[  Obtiene los detalles del formato del ticket  ]
                vlstrSentencia = "                  Select * "
                vlstrSentencia = vlstrSentencia & " From pvFormatoTicket "
                vlstrSentencia = vlstrSentencia & "      Inner Join pvDetalleFormatoTicket On (pvFormatoTicket.intCveFormatoTicket = pvDetalleFormatoTicket.intCveFormatoTicket)"
                vlstrSentencia = vlstrSentencia & " Where pvFormatoTicket.intCveFormatoTicket = " & txtClave.Text & " And"
                vlstrSentencia = vlstrSentencia & "       pvDetalleFormatoTicket.vchSeccion = '" & vlstrSeccion & "'"
                vlstrSentencia = vlstrSentencia & " Order by intRenglon, intColumna"
                Set vlrsFormatoTicket = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                With vlrsFormatoTicket
                    '[  Dado que solo el cuerpo provocará un ciclo, forza a que las demás secciones solo den una vuelta  ]
                    If vlstrSeccion <> "C" Then vlrsValoresTicket.MoveLast
                    Do While Not vlrsValoresTicket.EOF
                        vlintRenglon = 0
                        '------------------------------------------------------
                        '««  Recorre el formato del ticket según su sección  »»
                        '------------------------------------------------------
                        Do While Not .EOF
                            vlintRenglon = vlrsFormatoTicket!intRenglon
                            vlstrLinea = ""
                            Do While vlintRenglon = vlrsFormatoTicket!intRenglon
                                '[  Valores fijos  ]
                                If vlrsFormatoTicket!vchTipoValor = "F" Then
                                    vlstrLinea = vlstrLinea & IIf(vlrsFormatoTicket!vchvalor = "×", " ", vlrsFormatoTicket!vchvalor)
                                Else '[  Campos insertables  ]
                                    '------------------------------------------------------
                                    '««  Localiza la posición del campo en el RecordSet  »»
                                    '------------------------------------------------------
                                    For vlintCont = 0 To vlrsValoresTicket.Fields.Count - 1
                                        If UCase(vlrsValoresTicket.Fields(vlintCont).Name) = UCase(vlrsFormatoTicket!vchvalor) Then
                                            vlintPosicion = vlintCont
                                            Exit For
                                        End If
                                    Next
                                    Select Case vlrsFormatoTicket!intTipoDato
                                        Case 0 '[  Cadena  ]
                                            If vlrsFormatoTicket!intLongMax > 0 Then
                                                vlstrLinea = vlstrLinea & RTrim(Mid(vlrsValoresTicket.Fields(vlintPosicion), 1, vlrsFormatoTicket!intLongMax))
                                            Else
                                                vlstrLinea = vlstrLinea & Mid(vlrsValoresTicket.Fields(vlintPosicion), 1, vlrsFormatoTicket!intLonguitud) & Space(vlrsFormatoTicket!intLonguitud - Len(Mid(IIf(IsNull(vlrsValoresTicket.Fields(vlintPosicion)), "", vlrsValoresTicket.Fields(vlintPosicion)), 1, vlrsFormatoTicket!intLonguitud)))
                                            End If
                                        Case 1 '[  Fecha   ]
                                            vlstrLinea = vlstrLinea & RTrim(Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "dd/mmm/yyyy"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud)))
                                        Case 2 '[  Moneda  ]
                                            vlstrLinea = vlstrLinea & Space((Len(vlstrLinea) + vlrsFormatoTicket!intLonguitud - Len(Format(vlrsValoresTicket.Fields(vlintPosicion), "###,###,##0.00")) - Len(vlstrLinea))) & Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "###,###,##0.00"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud))
                                        Case 3 '[  Número  ]
                                            vlstrLinea = vlstrLinea & Space((Len(vlstrLinea) + vlrsFormatoTicket!intLonguitud - Len(Format(vlrsValoresTicket.Fields(vlintPosicion), "###,###,###")) - Len(vlstrLinea))) & Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "###,###,###"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud))
                                        Case 4 '[   Hora   ]
                                            vlstrLinea = vlstrLinea & RTrim(Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "hh:mm"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud)))
                                    End Select
                                    
                                End If
                                .MoveNext
                                If .EOF Then Exit Do
                            Loop
                            txtPresentacionPreliminar.Text = txtPresentacionPreliminar.Text & vlstrLinea & vbCrLf
                        Loop
                        vlrsValoresTicket.MoveNext
                        If vlstrSeccion = "C" Then .MoveFirst
                    Loop
                    vlrsValoresTicket.MoveFirst
                End With
                txtPresentacionPreliminar.Visible = True
            Next
        Else
            '[  ¡No existe información!  ]
            MsgBox SIHOMsg(13), vbCritical, "Mensaje"
            chkPresentacionPreliminar.Value = 0
        End If

End Sub


Private Function fblnExistenFormatos() As Boolean
    Dim vlstrSentencia As String
    Dim vlrsFormatos As New ADODB.Recordset

    fblnExistenFormatos = True
    vlstrSentencia = "Select pvFormatoTicket.INTCVEFORMATOTICKET, pvFormatoTicket.VCHDESCRIPCION From pvFormatoTicket"
    Set vlrsFormatos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If vlrsFormatos.RecordCount = 0 Then
        '[  ¡No existe información!  ]
        MsgBox SIHOMsg(13), vbCritical, "Mensaje"
        fblnExistenFormatos = False
    End If
End Function

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer, vlb4 As Integer, vlb5 As Integer)
    On Error GoTo NotificaError
    
    cmdPrimer.Enabled = vlb1 = 1
    cmdAnterior.Enabled = vlb2 = 1
    cmdBuscar.Enabled = vlb3 = 1
    cmdSiguiente.Enabled = vlb4 = 1
    cmdUltimo.Enabled = vlb5 = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
End Sub

