VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmManejoMedicamentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manejo de medicamentos"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTobj 
      Height          =   4095
      Left            =   -10
      TabIndex        =   15
      Top             =   -10
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   7223
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmManejoMedicamentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "comColorLetra"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "optOrden(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmManejoMedicamentos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfManejos"
      Tab(1).ControlCount=   1
      Begin VB.OptionButton optOrden 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Clave"
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
         Height          =   390
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   -300
         Width           =   1080
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
         Height          =   2490
         Left            =   120
         TabIndex        =   18
         Top             =   40
         Width           =   6780
         Begin VB.TextBox txtColorLetra 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   13
            ToolTipText     =   "Color de letra del manejo"
            Top             =   1110
            Width           =   750
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
            Height          =   250
            Left            =   1800
            TabIndex        =   6
            ToolTipText     =   "Manejo activo"
            Top             =   2100
            Value           =   1  'Checked
            Width           =   2910
         End
         Begin VB.TextBox txtDescripcion 
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   "Descripción del manejo"
            Top             =   710
            Width           =   4800
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   0
            ToolTipText     =   "Clave del manejo"
            Top             =   300
            Width           =   750
         End
         Begin VB.CheckBox chkControlado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Controlado"
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
            Height          =   270
            Left            =   5150
            TabIndex        =   4
            ToolTipText     =   "Manejo cotrolado"
            Top             =   1170
            Width           =   1470
         End
         Begin VB.CheckBox chkVerificacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Requiere doble verificación al administrar el medicamento"
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
            Height          =   495
            Left            =   1800
            TabIndex        =   5
            ToolTipText     =   "Requiere doble verificación de contraseña en la administración del medicamento"
            Top             =   1560
            Width           =   4575
         End
         Begin HSFlatControls.MyCombo cboSimbolo 
            Height          =   375
            Left            =   4200
            TabIndex        =   3
            ToolTipText     =   "Símbolo del manejo"
            Top             =   1110
            Width           =   855
            _ExtentX        =   1508
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
         Begin MyCommandButton.MyButton cmdColorLetra 
            Height          =   390
            Left            =   2580
            TabIndex        =   2
            ToolTipText     =   "Seleccionar color de letra del manejo"
            Top             =   1110
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   688
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
            Picture         =   "frmManejoMedicamentos.frx":0038
            BackColorDown   =   -2147483633
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483642
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":0569
            PictureAlignment=   4
            ShowFocus       =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Color de letra"
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
            TabIndex        =   19
            Top             =   1170
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción"
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
            TabIndex        =   20
            Top             =   770
            Width           =   1125
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   195
            TabIndex        =   21
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000005&
            Caption         =   "Símbolo"
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
            Left            =   3240
            TabIndex        =   22
            Top             =   1170
            Width           =   960
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid vsfManejos 
         Height          =   3200
         Left            =   -74910
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   120
         Width           =   6820
         _cx             =   12030
         _cy             =   5644
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
         TreeColor       =   -2147483643
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmManejoMedicamentos.frx":0C7B
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
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
      Begin MSComDlg.CommonDialog comColorLetra 
         Left            =   6240
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame4 
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
         Left            =   1360
         TabIndex        =   17
         Top             =   2480
         Width           =   4320
         Begin MyCommandButton.MyButton cmdPrimerRegistro 
            Height          =   600
            Left            =   60
            TabIndex        =   9
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
            Picture         =   "frmManejoMedicamentos.frx":0CFB
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":167D
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBusqueda 
            Height          =   600
            Left            =   1260
            TabIndex        =   8
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
            Picture         =   "frmManejoMedicamentos.frx":1FFF
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":2983
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSiguienteRegistro 
            Height          =   600
            Left            =   1860
            TabIndex        =   11
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
            Picture         =   "frmManejoMedicamentos.frx":3307
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":3C89
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdUltimoRegistro 
            Height          =   600
            Left            =   2460
            TabIndex        =   12
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
            Picture         =   "frmManejoMedicamentos.frx":460B
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":4F8D
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdGrabar 
            Height          =   600
            Left            =   3060
            TabIndex        =   7
            ToolTipText     =   "Guardar el manejo de medicamento"
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
            Picture         =   "frmManejoMedicamentos.frx":590F
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":6293
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBorrar 
            Height          =   600
            Left            =   3660
            TabIndex        =   14
            ToolTipText     =   "Eliminar el manejo de medicamento"
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
            Picture         =   "frmManejoMedicamentos.frx":6C17
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":7599
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdAnteriorRegistro 
            Height          =   600
            Left            =   660
            TabIndex        =   10
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
            Picture         =   "frmManejoMedicamentos.frx":7F1D
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmManejoMedicamentos.frx":889F
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmManejoMedicamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnSoloBusqueda As Boolean

Dim rsManejos As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset

Dim lngPersonaGraba As Long

Dim blnConsulta As Boolean

Private Sub cboSimbolo_Change()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboSimbolo_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboSimbolo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkActivo_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkControlado_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkControlado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkVerificacion_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkVerificacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdAnteriorRegistro_Click()

    If rsManejos.RecordCount > 0 Then
    
        If Not rsManejos.BOF Then rsManejos.MovePrevious
        If rsManejos.BOF Then rsManejos.MoveNext
        
        pMuestra rsManejos!intCveManejo
        
    End If
    
End Sub

Private Sub cmdBorrar_Click()
On Error GoTo NotificaError
Dim lngManejo As Long
Dim rs As New ADODB.Recordset

    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "IV", 2438, 2439), "E") Then
        
        Set rs = frsRegresaRs("Select * From IvArticuloManejo Where intCveManejo = " & CStr(rsManejos!intCveManejo))
        If rs.EOF Then
        
            lngManejo = rsManejos!intCveManejo
        
            lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento, "E")
            If lngPersonaGraba = 0 Then Exit Sub
            
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            pEjecutaSentencia ("Delete From IvManejoMedicamento Where intCveManejo = " & Trim(txtNumero.Text))
            
            pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "MANEJO DE MEDICAMENTOS", CStr(lngManejo)
        
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            rsManejos.Requery
            pLimpia
            txtNumero.SetFocus
        
        Else
        
            'No se puede eliminar la información, ya ha sido utilizada.
             MsgBox SIHOMsg(771), vbExclamation, "Mensaje"
             
        End If
        rs.Close
    
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBorrar_Click"))
End Sub

Private Sub cmdBusqueda_Click()
    sstObj.Tab = 1
    pBusqueda
End Sub

Private Sub cmdColorLetra_Click()

    With comColorLetra
        
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .Color = txtColorLetra.BackColor
        .ShowColor
        
        If .Color = 8388608 Then
            'Color repetido o restringido su uso, por favor seleccione otro color
            MsgBox SIHOMsg(1071), vbExclamation, "Mensaje"
        Else
            txtColorLetra.BackColor = .Color
            SendKeys vbTab
        End If
        
    End With

End Sub

Private Sub cmdColorLetra_GotFocus()
    If blnConsulta Then pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cmdColorLetra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo NotificaError
Dim strColor As String
Dim lngCveManejo As Long
    
    If Trim(txtDescripcion) = "" Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation, "Mensaje"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    
    Set rsTemp = frsRegresaRs("Select Count(*) Manejo From IvManejoMedicamento Where vchSimbolo = '" & cboSimbolo.List(cboSimbolo.ListIndex) & "' And vchColor = '" & txtColorLetra.BackColor & "'" & IIf(blnConsulta, " And intCveManejo <> " & Trim(CStr(txtNumero.Text)), ""))
    If rsTemp!MANEJO > 0 Then
        'Color repetido o restringido su uso, por favor seleccione otro color
        MsgBox SIHOMsg(1071), vbExclamation, "Mensaje"
        Exit Sub
    End If
    rsTemp.Close
    
    strColor = txtColorLetra.BackColor
    
    lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento, "E")
    If lngPersonaGraba = 0 Then Exit Sub

    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    Set rsTemp = frsRegresaRs("Select * From IvManejoMedicamento Where intCveManejo = " & Trim(txtNumero.Text), adLockOptimistic, adOpenDynamic)
    
    With rsTemp
        If rsTemp.RecordCount = 0 Then .AddNew
        !vchDescripcion = Trim(txtDescripcion.Text)
        !vchColor = strColor
        !vchSimbolo = cboSimbolo.List(cboSimbolo.ListIndex)
        !bitControlado = chkControlado.Value
        !bitactivo = chkActivo.Value
        !BitRequiereAutorizacion = chkVerificacion.Value
        .Update
    End With
    
    If Not blnConsulta Then
      lngCveManejo = flngObtieneIdentity("SEC_IVMANEJOMEDICAMENTO", rsTemp!intCveManejo)
    Else
      lngCveManejo = rsTemp!intCveManejo
    End If
    rsTemp.Close
    
    pGuardarLogTransaccion Me.Name, IIf(Not blnConsulta, EnmGrabar, EnmCambiar), lngPersonaGraba, "MANEJO DE MEDICAMENTOS", Trim(txtNumero.Text)
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
    rsManejos.Requery
    pMuestra lngCveManejo
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabar_Click"))
End Sub

Private Sub cmdPrimerRegistro_Click()
    
    If rsManejos.RecordCount > 0 Then
        rsManejos.MoveFirst
        pMuestra rsManejos!intCveManejo
    End If
    
End Sub

Private Sub cmdSiguienteRegistro_Click()

    If rsManejos.RecordCount > 0 Then
    
        If Not rsManejos.EOF Then rsManejos.MoveNext
        If rsManejos.EOF Then rsManejos.MovePrevious

        pMuestra rsManejos!intCveManejo

    End If
    
End Sub

Private Sub cmdUltimoRegistro_Click()

    If rsManejos.RecordCount > 0 Then
        rsManejos.MoveLast
        pMuestra rsManejos!intCveManejo
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then
    
        If blnSoloBusqueda Then
            Unload Me
        Else
            If blnConsulta Or sstObj.Tab = 1 Then
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pLimpia
                    KeyAscii = 0
                    txtNumero.SetFocus
                End If
            Else
                Unload Me
            End If
        End If
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError

    'Color de Tab
    SetStyle sstObj.hwnd, 0
    SetSolidColor sstObj.hwnd, 16777215
    SSTabSubclass sstObj.hwnd


    vgstrNombreForm = Me.Name
    Me.Icon = frmMenuPrincipal.Icon
    
    blnConsulta = False
    
    Set rsManejos = frsEjecuta_SP("-1|" & IIf(blnSoloBusqueda, "1", "-1") & "|-1", "Sp_IvSelManejos")
    
    If blnSoloBusqueda Then
    
        If Not rsManejos.EOF Then
            sstObj.Tab = 1
            pBusqueda
        Else
            '¡No existen manejos de medicamentos!
            MsgBox SIHOMsg(1107), vbOKOnly + vbExclamation, "Mensaje"
            Unload Me
        End If
        
    ElseIf Not blnSoloBusqueda Then
        sstObj.Tab = 0
        
        cboSimbolo.Clear
        cboSimbolo.AddItem "l"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 1
        cboSimbolo.AddItem "¡"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 2
        cboSimbolo.AddItem "u"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 3
        cboSimbolo.AddItem "p"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 4
        cboSimbolo.AddItem "x"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 5
        cboSimbolo.AddItem "y"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 6
        cboSimbolo.AddItem "v"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 7
        cboSimbolo.AddItem "T"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 8
        cboSimbolo.AddItem "z"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 9
        cboSimbolo.AddItem "i"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 10
        cboSimbolo.AddItem "°"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 11
        cboSimbolo.AddItem "2"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 12
        cboSimbolo.AddItem "5"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 13
        cboSimbolo.AddItem "6"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 14
        cboSimbolo.AddItem "¿"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 15
        cboSimbolo.AddItem "?"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 16
        cboSimbolo.AddItem "S"
        cboSimbolo.ItemData(cboSimbolo.NewIndex) = 17
        
        pLimpia
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If sstObj.Tab = 1 And Not blnSoloBusqueda Then
        Cancel = True
        sstObj.Tab = 0
    Else
        rsManejos.Close
    End If
        
End Sub

Private Sub txtDescripcion_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
End Sub

Private Sub txtNumero_GotFocus()
    pLimpia
    pSelTextBox txtNumero
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = vbKeyReturn Then
        
        If Trim(txtNumero.Text) = "" Then txtNumero.Text = flngSiguiente("intCveManejo", "IvManejoMedicamento")
        
        If fintLocalizaPkRs(rsManejos, 0, txtNumero.Text) = 0 Then
            txtNumero.Text = flngSiguiente("intCveManejo", "IvManejoMedicamento")
        Else
            pMuestra CLng(txtNumero.Text)
        End If

        SendKeys vbTab
        
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumero_KeyPress"))
End Sub

Private Sub pBusqueda()
On Error GoTo NotificaError
Dim intcontador As Integer

    vsfManejos.Rows = 1
    
    If rsManejos.RecordCount > 0 Then
        
        rsManejos.MoveFirst
        With vsfManejos
        
            For intcontador = 1 To rsManejos.RecordCount
                .Rows = .Rows + 1
                .Col = 1
                .Row = intcontador
                .CellFontName = "Century Gothic"
                .TextMatrix(intcontador, 1) = rsManejos!intCveManejo
                .Col = 2
                .Row = intcontador
                .CellFontName = "Wingdings"
                .CellFontSize = 12
                .CellForeColor = CLng(rsManejos!vchColor)
                .TextMatrix(intcontador, 2) = rsManejos!vchSimbolo
                .Col = 3
                .Row = intcontador
                .FontName = "Century Gothic"
                .TextMatrix(intcontador, 3) = rsManejos!vchDescripcion
                rsManejos.MoveNext
            Next intcontador
            rsManejos.MoveFirst
        
        End With
    
    Else
        '¡No existen manejos de medicamentos!
        MsgBox SIHOMsg(1107), vbOKOnly + vbExclamation, "Mensaje"
        If blnSoloBusqueda Then
            Unload Me
        Else
            sstObj.Tab = 0
            Exit Sub
        End If
    End If
    
    If fblnCanFocus(vsfManejos) Then vsfManejos.SetFocus
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBusqueda"))
End Sub

Private Sub pHabilita(intPrimero As Integer, intAnterior As Integer, intBusqueda As Integer, intSiguiente As Integer, intUltimo As Integer, intGrabar As Integer, intBorrar As Integer)

    cmdPrimerRegistro.Enabled = intPrimero = 1
    cmdAnteriorRegistro.Enabled = intAnterior = 1
    cmdBusqueda.Enabled = intBusqueda = 1
    cmdSiguienteRegistro.Enabled = intSiguiente = 1
    cmdUltimoRegistro.Enabled = intUltimo = 1
    cmdGrabar.Enabled = intGrabar = 1
    cmdBorrar.Enabled = intBorrar = 1

End Sub

Private Sub pLimpia()
On Error GoTo NotificaError
    
    sstObj.Tab = IIf(blnSoloBusqueda, 1, 0)
    blnConsulta = False
    txtNumero.Text = flngSiguiente("intCveManejo", "IvManejoMedicamento")
    txtDescripcion.Text = ""
    txtColorLetra.BackColor = 986895
    If Not blnSoloBusqueda Then cboSimbolo.ListIndex = 1
    chkControlado.Value = 0
    chkVerificacion.Value = 0
    Set rsTemp = frsRegresaRs("Select Count(*) Controlado From IvManejoMedicamento Where bitControlado = 1 And bitActivo = 1")
    If rsTemp!Controlado > 0 Then
        chkControlado.Enabled = False
    Else
        chkControlado.Enabled = True
    End If
    rsTemp.Close
    chkActivo.Value = 1
    
    pHabilita 1, 1, 1, 1, 1, 0, 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Sub pMuestra(lngClaveManejo As Long)
On Error GoTo NotificaError

    sstObj.Tab = 0
    blnConsulta = True
    
    rsManejos.Find "intCveManejo = '" & CStr(lngClaveManejo) & "'"
    txtNumero.Text = rsManejos!intCveManejo
    txtDescripcion.Text = rsManejos!vchDescripcion
    txtColorLetra.BackColor = rsManejos!vchColor
    cboSimbolo.ListIndex = fintLocalizaCritCbo_new(cboSimbolo, CStr(rsManejos!vchSimbolo))
    chkControlado.Value = rsManejos!bitControlado
    chkVerificacion.Value = rsManejos!BitRequiereAutorizacion
    If rsManejos!bitControlado = 0 Then
        Set rsTemp = frsRegresaRs("Select Count(*) Controlado From IvManejoMedicamento Where bitControlado = 1 And bitActivo = 1")
        If rsTemp!Controlado > 0 Then
            chkControlado.Enabled = False
        Else
            chkControlado.Enabled = True
        End If
        rsTemp.Close
    Else
        chkControlado.Enabled = True
    End If
    chkActivo.Value = rsManejos!bitactivo
    
    pHabilita 1, 1, 1, 1, 1, 0, 1
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra"))
End Sub

Private Sub vsfManejos_DblClick()

    If Not blnSoloBusqueda Then
        pMuestra CLng(vsfManejos.TextMatrix(vsfManejos.Row, 1))
    End If
    
End Sub

