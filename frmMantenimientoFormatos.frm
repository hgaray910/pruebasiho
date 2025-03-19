VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantenimientoFormatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formatos de impresión"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabFormato 
      Height          =   5790
      Left            =   -45
      TabIndex        =   24
      Top             =   -360
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   10213
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantenimientoFormatos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdContenido"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCoordenada"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantenimientoFormatos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   2535
         Left            =   155
         TabIndex        =   36
         Top             =   1920
         Visible         =   0   'False
         Width           =   9460
         Begin VB.TextBox txtNombreACrystal5 
            Height          =   315
            Left            =   4680
            TabIndex        =   15
            ToolTipText     =   "Nombre que tendrá el archivo de Crystal Reports"
            Top             =   2040
            Width           =   3735
         End
         Begin VB.TextBox txtNombreACrystal4 
            Height          =   315
            Left            =   4680
            TabIndex        =   39
            ToolTipText     =   "Nombre que tendrá el archivo de Crystal Reports"
            Top             =   950
            Width           =   3735
         End
         Begin VB.TextBox txtNombreACrystal3 
            Height          =   315
            Left            =   4680
            MaxLength       =   100
            TabIndex        =   14
            ToolTipText     =   "Nombre que tendrá el archivo de Crystal Reports"
            Top             =   1440
            Width           =   3735
         End
         Begin VB.TextBox txtNombreACrystal2 
            Height          =   315
            Left            =   4680
            MaxLength       =   100
            TabIndex        =   12
            ToolTipText     =   "Nombre que tendrá el archivo de Crystal Reports"
            Top             =   960
            Width           =   3735
         End
         Begin VB.TextBox txtNombreACrystal1 
            Height          =   315
            Left            =   4680
            MaxLength       =   100
            TabIndex        =   10
            ToolTipText     =   "Nombre que tendrá el archivo de Crystal Reports"
            Top             =   480
            Width           =   3735
         End
         Begin VB.OptionButton optDesglozadoCargo 
            Caption         =   "Desglosado por cargo"
            Height          =   255
            Left            =   1200
            TabIndex        =   13
            ToolTipText     =   "La factura se imprimirá sin realizar ningún tipo de agrupación, es decir mostrará todos los cargos que tenga la cuenta"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.OptionButton optAcargo 
            Caption         =   "Agrupado por cargo"
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            ToolTipText     =   "La factura se imprimirá agrupando los cargos por clave, tipo, precio y descuento"
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optAconceptoFacturacion 
            Caption         =   "Agrupado por concepto de facturación"
            Height          =   255
            Left            =   1200
            TabIndex        =   9
            ToolTipText     =   "La factura se imprimirá agrupando los cargos por concepto de facturación"
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label7 
            Caption         =   "Formato único CFDI 3.3"
            Height          =   255
            Left            =   1470
            TabIndex        =   40
            Top             =   2070
            Width           =   3015
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            X1              =   120
            X2              =   9240
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label lblNombreACrystal2 
            Caption         =   "Nombre del archivo Crystal Reports"
            Height          =   255
            Left            =   1440
            TabIndex        =   38
            Top             =   980
            Width           =   2775
         End
         Begin VB.Label lblNombreACrystal 
            Caption         =   "Nombre del archivo Crystal Reports"
            Height          =   255
            Left            =   4680
            TabIndex        =   37
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1515
         Left            =   150
         TabIndex        =   29
         Top             =   360
         Width           =   9450
         Begin VB.CheckBox chkComprobanteFiscal 
            Caption         =   "Comprobante fiscal digital"
            Height          =   375
            Left            =   2280
            TabIndex        =   1
            ToolTipText     =   "Indíca si es un formato físico o digital"
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkAgrupadoCargos 
            Caption         =   "Agrupar por cargo"
            Height          =   255
            Left            =   5540
            TabIndex        =   6
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txtRenglones 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8670
            TabIndex        =   7
            Top             =   990
            Width           =   585
         End
         Begin VB.ComboBox cboTamano 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Seleccione el tamaño de letra para la impresión del formato"
            Top             =   990
            Width           =   705
         End
         Begin VB.TextBox txtTotalLineas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2520
            MaxLength       =   9
            TabIndex        =   4
            ToolTipText     =   "Total de lineas del formato"
            Top             =   990
            Width           =   705
         End
         Begin VB.TextBox txtTotalCols 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            MaxLength       =   9
            TabIndex        =   3
            ToolTipText     =   "Total de Columnas del formato"
            Top             =   990
            Width           =   705
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   2
            ToolTipText     =   "Descripción del formato"
            Top             =   645
            Width           =   8055
         End
         Begin VB.TextBox txtClave 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            MaxLength       =   9
            TabIndex        =   0
            ToolTipText     =   "Clave del formato"
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tamaño de letra"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   195
            Left            =   3480
            TabIndex        =   35
            Top             =   1050
            Width           =   1155
         End
         Begin VB.Label lblRenglones 
            AutoSize        =   -1  'True
            Caption         =   "Renglones detalle"
            Height          =   195
            Left            =   7320
            TabIndex        =   34
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Largo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   195
            Left            =   2040
            TabIndex        =   33
            Top             =   1050
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ancho"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   1050
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   705
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5340
         Left            =   -74825
         TabIndex        =   27
         Top             =   360
         Width           =   9405
         Begin VSFlex7LCtl.VSFlexGrid vsfFormato 
            Height          =   5090
            Left            =   75
            TabIndex        =   28
            ToolTipText     =   "Formatos configurados"
            Top             =   180
            Width           =   9240
            _cx             =   16298
            _cy             =   8978
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12632256
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   12
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
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
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   1
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
         Height          =   735
         Left            =   3135
         TabIndex        =   25
         Top             =   4800
         Width           =   3645
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   90
            Picture         =   "frmMantenimientoFormatos.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Primer formato"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   585
            Picture         =   "frmMantenimientoFormatos.frx":015A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Anterior formato"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1080
            Picture         =   "frmMantenimientoFormatos.frx":02CC
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Búsqueda de formatos"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1575
            Picture         =   "frmMantenimientoFormatos.frx":043E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Siguiente formato"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2085
            Picture         =   "frmMantenimientoFormatos.frx":05B0
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Ultimo formato"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3075
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantenimientoFormatos.frx":0722
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Borrar formato"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2580
            Picture         =   "frmMantenimientoFormatos.frx":08C4
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Grabar formato"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.TextBox txtCoordenada 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   255
         MaxLength       =   6
         TabIndex        =   23
         Text            =   "0"
         ToolTipText     =   "Coordenada"
         Top             =   4695
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdContenido 
         Height          =   2450
         Left            =   170
         TabIndex        =   8
         ToolTipText     =   "Contenido del formato"
         Top             =   1920
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   4313
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         GridColor       =   12632256
         HighLight       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         Caption         =   "(*) Si no se captura el ancho, se imprimirá el dato completo"
         Height          =   240
         Left            =   135
         TabIndex        =   26
         Top             =   4425
         Width           =   4230
      End
   End
End
Attribute VB_Name = "frmMantenimientoFormatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Programa para dar mantenimiento a formatos de impresión.
' Ejemplo cheques y facturas.
' Fecha de desarrollo: Lunes 8 de Enero del 2001
'-------------------------------------------------------------------------------
' Ultimas modificaciones:
' Fecha:
' Descripción del cambio:
'-------------------------------------------------------------------------------
Dim rsContenidoFormato As New ADODB.Recordset
Dim rsFormato As New ADODB.Recordset                    'Es usado para los comprobantes fiscales NO DIGITALES
Dim rsFormatoDetalle As New ADODB.Recordset
Dim rsFormatoDigital As New ADODB.Recordset             'Es utilizado para los comprobantes fiscales DIGITALES
Dim vlstrsql As String
Dim vlblnConsulta As Boolean
Dim vllngNumeroRegistros As Long
Dim llngPersonaGraba As Long
Dim blnCargando As Boolean
Public vllngNumeroOpcionModulo As Long

Private Sub pHabilita(a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer, g As Integer)
    On Error GoTo NotificaError
    
    cmdTop.Enabled = a = 1
    cmdBack.Enabled = b = 1
    cmdLocate.Enabled = c = 1
    cmdNext.Enabled = d = 1
    cmdEnd.Enabled = e = 1
    cmdSave.Enabled = f = 1
    cmdDelete.Enabled = g = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))

End Sub

Private Sub cboTamano_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTamano_GotFocus"))
End Sub

Private Sub cboTamano_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If fblnCanFocus(chkAgrupadoCargos) Then
            chkAgrupadoCargos.SetFocus
        Else
            grdContenido.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTamano_KeyDown"))
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkAgrupadoCargos_Click()
    If chkAgrupadoCargos.Value = vbChecked Then
        txtRenglones.Enabled = True
    Else
        txtRenglones.Text = ""
        txtRenglones.Enabled = False
    End If
    If Not blnCargando Then
        pMuestraDetalle
    End If
End Sub

Private Sub chkAgrupadoCargos_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkAgrupadoCargos_GotFocus"))

End Sub

Private Sub chkAgrupadoCargos_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If fblnCanFocus(txtRenglones) Then
            txtRenglones.SetFocus
        Else
            grdContenido.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkAgrupadoCargos_KeyDown"))
End Sub

Private Sub chkComprobanteFiscal_Click()
pLimpia
pHabilita 1, 1, 1, 1, 1, 0, 0
If chkComprobanteFiscal.Value = vbChecked Then
    grdContenido.Visible = False
    txtTotalCols.Visible = False
    txtTotalLineas.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    cboTamano.Visible = False
'AGREGADO
If vglngNumeroTipoFormato = 2 Then
    chkAgrupadoCargos.Visible = False
    lblRenglones.Visible = False
    txtRenglones.Visible = False
    lblNombreACrystal.Visible = True
End If
If vglngNumeroTipoFormato = 9 Then
    lblNombreACrystal2.Visible = True
    txtNombreACrystal4.Visible = True
End If
'FIN AGREGADO

'TEST PARA EL FORMATO DE NOTAS
If vglngNumeroTipoFormato = 8 Then
    lblNombreACrystal2.Visible = True
    txtNombreACrystal4.Visible = True
End If
'END TEST

    Frame4.Visible = True
    optAconceptoFacturacion.Value = True
Else
    grdContenido.Visible = True
    txtTotalCols.Visible = True
    txtTotalLineas.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    cboTamano.Visible = True
'AGREGADO
If vglngNumeroTipoFormato = 2 Then
    chkAgrupadoCargos.Visible = True
    lblRenglones.Visible = True
    txtRenglones.Visible = True
    lblNombreACrystal.Visible = True
End If
If vglngNumeroTipoFormato = 9 Then
    lblNombreACrystal2.Visible = True
    txtNombreACrystal4.Visible = True
End If
'FIN AGREGADO

'TEST PARA EL FORMATO DE NOTAS
If vglngNumeroTipoFormato = 8 Then
    lblNombreACrystal2.Visible = False
    txtNombreACrystal4.Visible = False
End If
'END TEST

    Frame4.Visible = False
    optAconceptoFacturacion.Value = False
    txtNombreACrystal1.Text = ""
    txtNombreACrystal2.Text = ""
    txtNombreACrystal3.Text = ""
    txtNombreACrystal4.Text = ""
End If

End Sub

Private Sub chkComprobanteFiscal_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
            txtDescripcion.SetFocus
    End If
End Sub

Private Sub cmdBack_Click()
    On Error GoTo NotificaError
    Dim vlblnTermina As Boolean

'Si no es un comprobante fiscal digital; hacer...
If chkComprobanteFiscal.Value = vbUnchecked Then
    If vllngNumeroRegistros <> 0 Then
        If Not rsFormato.BOF Then
            rsFormato.MovePrevious
        End If
        vlblnTermina = False
        Do While Not rsFormato.BOF And Not vlblnTermina
            If rsFormato!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                rsFormato.MovePrevious
            Else
                vlblnTermina = True
            End If
        Loop
        If rsFormato.BOF Then
            rsFormato.MoveNext
            Do While Not rsFormato.EOF And Not vlblnTermina
                If rsFormato!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                    rsFormato.MoveNext
                Else
                    vlblnTermina = True
                End If
            Loop
        End If
        If Not rsFormato.BOF Then
            pMuestra
            pHabilita 1, 1, 1, 1, 1, 0, 1
        End If
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If
'Si no; hacer...
Else
    If vllngNumeroRegistros <> 0 Then
        If Not rsFormatoDigital.BOF Then
            rsFormatoDigital.MovePrevious
        End If
        vlblnTermina = False
        Do While Not rsFormatoDigital.BOF And Not vlblnTermina
            If rsFormatoDigital!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                rsFormatoDigital.MovePrevious
            Else
                vlblnTermina = True
            End If
        Loop
        If rsFormatoDigital.BOF Then
            rsFormatoDigital.MoveNext
            Do While Not rsFormatoDigital.EOF And Not vlblnTermina
                If rsFormatoDigital!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                    rsFormatoDigital.MoveNext
                Else
                    vlblnTermina = True
                End If
            Loop
        End If
        If Not rsFormatoDigital.BOF Then
            pMuestra
            pHabilita 1, 1, 1, 1, 1, 0, 1
        End If
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If

End If
'Fin de agregado


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionModulo, "C") Then
      '--------------------------------------------------------
      ' Persona que graba
      '--------------------------------------------------------
      vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
      If vllngPersonaGraba <> 0 Then
        If fblnIntegridadValida() Then
            If MsgBox(SIHOMsg(6), vbYesNo + vbExclamation, "Mensaje") = vbYes Then
                Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "FORMATO DE IMPRESION", txtClave.Text)
                EntornoSIHO.ConeccionSIHO.BeginTrans
                        pEjecutaSentencia ("Delete from  FormatoDetalle where intNumeroFormato=" + txtClave.Text)
                            rsFormatoDetalle.Requery
            'Agregado
            'Si es comprobante fiscal digital hacer...
                    If chkComprobanteFiscal.Value = vbChecked Then
                        rsFormatoDigital.Delete
                        rsFormatoDigital.Update
            'Si no; hacer...
                    Else
                        rsFormato.Delete
                        rsFormato.Update
                    End If
            'Fin agregado
                    Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "FORMATO DE IMPRESION", txtClave.Text)
                EntornoSIHO.ConeccionSIHO.CommitTrans
                pNumeroRegistros
                txtClave.SetFocus
            End If
        Else
            MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        End If
      End If
    End If
rsFormato.Requery
rsFormatoDigital.Requery
'AGREGADO: Desactivar el chkComprobante fiscal después de borrar o cancelar una transacción
'Se optó mejor por la opcion de dar el SetFocus al txtClave
txtClave.SetFocus '<-- chkComprobanteFiscal.Value = vbUnchecked
'FIN AGREGADO
pLimpia
Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Sub cmdDelete_Click"))
End Sub

Private Function fblnIntegridadValida() As Boolean
    On Error GoTo NotificaError
    
    Dim rsEncontrados As New ADODB.Recordset
    
    fblnIntegridadValida = True
    vlstrsql = "select count(*) from CpBanco where intNumeroFormato=" + txtClave.Text
    Set rsEncontrados = frsRegresaRs(vlstrsql)
    If rsEncontrados.Fields(0) <> 0 Then
        fblnIntegridadValida = False
    Else
        vlstrsql = "select count(*) from PvDocumentoDepartamento where intNumFormato=" + txtClave.Text
        Set rsEncontrados = frsRegresaRs(vlstrsql)
        If rsEncontrados.Fields(0) <> 0 Then
            fblnIntegridadValida = False
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnIntegridadValida"))
End Function

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
'Agregado
'Si NO es comprobante fiscal digital hacer..
If chkComprobanteFiscal.Value = vbUnchecked Then
    If vllngNumeroRegistros <> 0 Then
        rsFormato.MoveLast
        Do While rsFormato!intNumeroTipoFormato <> vglngNumeroTipoFormato
            rsFormato.MovePrevious
        Loop
        pMuestra
        pHabilita 1, 1, 1, 1, 1, 0, 1
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If
'Si no...
Else
    If vllngNumeroRegistros <> 0 Then
        rsFormatoDigital.MoveLast
        Do While rsFormatoDigital!intNumeroTipoFormato <> vglngNumeroTipoFormato
            rsFormatoDigital.MovePrevious
        Loop
        pMuestra
        pHabilita 1, 1, 1, 1, 1, 0, 1
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If
End If
'Fin de agregado

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    If vllngNumeroRegistros <> 0 Then
        'frmMantenimientoFormatosConsulta.Show vbModal
        SSTabFormato.Tab = 1
        pCargaFormatos
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub pCargaFormatos()
    Dim rs As New ADODB.Recordset
    
    With vsfFormato
    
        .Clear
        .Rows = 2
        .Cols = 3
        .FormatString = "|Número|Descripción|Tipo"
        
        Set rs = frsEjecuta_SP(CStr(vglngNumeroTipoFormato), "SP_GNSELFORMATOMAESTRO")
        
        Do While Not rs.EOF
            .TextMatrix(.Rows - 1, 1) = rs!intNumeroFormato
            .TextMatrix(.Rows - 1, 2) = rs!VCHDESCRIPCION
            '.TextMatrix(.Rows - 1, 3) = rs!bitComprobantefiscaldigital
            .TextMatrix(.Rows - 1, 3) = IIf(rs!bitComprobantefiscaldigital = 1, "Digital", "Físico")
            .Rows = .Rows + 1
            rs.MoveNext
        Loop
        .Rows = .Rows - 1
        
        .ColWidth(1) = 1000
        .ColWidth(2) = 6000
        .ColWidth(3) = 1800
        
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
    End With

End Sub


Private Sub cmdNext_Click()
    On Error GoTo NotificaError
    
Dim vlblnTermina As Boolean
'Agregado
'Si no es un comprobante fiscal digital hacer...
If chkComprobanteFiscal.Value = vbUnchecked Then
    If vllngNumeroRegistros <> 0 Then
        If Not rsFormato.EOF Then
            rsFormato.MoveNext
        End If
        vlblnTermina = False
        Do While Not rsFormato.EOF And Not vlblnTermina
            If rsFormato!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                rsFormato.MoveNext
            Else
                vlblnTermina = True
            End If
        Loop
        If rsFormato.EOF Then
            rsFormato.MovePrevious
            Do While Not rsFormato.BOF And Not vlblnTermina
                If rsFormato!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                    rsFormato.MovePrevious
                Else
                    vlblnTermina = True
                End If
            Loop
        End If
        If Not rsFormato.EOF Then
            pMuestra
            pHabilita 1, 1, 1, 1, 1, 0, 1
        End If
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If
' Si no; hacer...
Else
    If vllngNumeroRegistros <> 0 Then
        If Not rsFormatoDigital.EOF Then
            rsFormatoDigital.MoveNext
        End If
        vlblnTermina = False
        Do While Not rsFormatoDigital.EOF And Not vlblnTermina
            If rsFormatoDigital!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                rsFormatoDigital.MoveNext
            Else
                vlblnTermina = True
            End If
        Loop
        If rsFormatoDigital.EOF Then
            rsFormatoDigital.MovePrevious
            Do While Not rsFormatoDigital.BOF And Not vlblnTermina
                If rsFormatoDigital!intNumeroTipoFormato <> vglngNumeroTipoFormato Then
                    rsFormatoDigital.MovePrevious
                Else
                    vlblnTermina = True
                End If
            Loop
        End If
        If Not rsFormatoDigital.EOF Then
            pMuestra
            pHabilita 1, 1, 1, 1, 1, 0, 1
        End If
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If
End If
'Fin de agregado


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub
Private Sub cmdSave_Click()

On Error GoTo NotificaError
Dim X As Integer
Dim vllngPersonaGraba As Long
    
         '--------------------------------------------------------------------
        '|                 ******** FORMATO FISICO ********                   |'
         '--------------------------------------------------------------------
         
If chkComprobanteFiscal.Value = vbUnchecked Then
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionModulo, "E", True) Then
        If fblnDatosValidos() Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            ' Si es modificación hacer... ---------------------------------
            If vlblnConsulta Then
                rsFormato!VCHDESCRIPCION = Trim(txtDescripcion.Text)
                rsFormato!intAncho = Int(Val(txtTotalCols.Text))
                rsFormato!intLargo = Int(Val(txtTotalLineas.Text))
                rsFormato!INTLETRA = Val(cboTamano.List(cboTamano.ListIndex))
                If vglngNumeroTipoFormato = 2 Then
                    rsFormato!intRenglonesDetalle = IIf(chkAgrupadoCargos.Value = vbChecked, Int(Val(txtRenglones.Text)), Null)
                    rsFormato!bitAgruparCargos = IIf(chkAgrupadoCargos.Value = vbChecked, 1, 0)
                End If
                rsFormato.Update
                rsFormato.Requery
                vlstrsql = "Delete from  FormatoDetalle where intNumeroFormato=" + txtClave.Text
                pEjecutaSentencia vlstrsql
            Else
            'Si es registro nuevo hacer... ---------------------------------
                With rsFormato
                    .AddNew
                    !VCHDESCRIPCION = Trim(txtDescripcion.Text)
                    !intNumeroTipoFormato = vglngNumeroTipoFormato
                    !intAncho = Int(Val(txtTotalCols.Text))
                    !intLargo = Int(Val(txtTotalLineas.Text))
                    !INTLETRA = Val(cboTamano.List(cboTamano.ListIndex))
                    If vglngNumeroTipoFormato = 2 Then
                        !intRenglonesDetalle = IIf(chkAgrupadoCargos.Value = vbChecked, Int(Val(txtRenglones.Text)), Null)
                        !bitAgruparCargos = IIf(chkAgrupadoCargos.Value = vbChecked, 1, 0)
                    End If
                End With
                rsFormato.Update
                rsFormato.Requery
                txtClave.Text = flngObtieneIdentity("SEC_FORMATO", rsFormato!intNumeroFormato)
            End If
            
            'Agregar a la tabla FormatoDetalle ---------------------------------
            With rsFormatoDetalle
                For X = 1 To grdContenido.Rows - 1
                    .AddNew
                    !intNumeroFormato = txtClave.Text
                    !intNumeroContenidoFormato = grdContenido.RowData(X)
                    !intCoordenadaX = IIf(Trim(grdContenido.TextMatrix(X, 2)) = "", 0, grdContenido.TextMatrix(X, 2))
                    !intCoordenadaY = IIf(Trim(grdContenido.TextMatrix(X, 3)) = "", 0, grdContenido.TextMatrix(X, 3))
                    !intAnchoContenido = IIf(Trim(grdContenido.TextMatrix(X, 4)) = "", 0, grdContenido.TextMatrix(X, 4))
                    .Update
                Next X
            End With
            Call pGuardarLogTransaccion(Me.Name, IIf(vlblnConsulta, EnmGrabar, EnmCambiar), vllngPersonaGraba, "FORMATO DE IMPRESION", txtClave.Text)
            EntornoSIHO.ConeccionSIHO.CommitTrans
            pNumeroRegistros
            txtClave.SetFocus
            pLimpia
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
Else
         '--------------------------------------------------------------------
        '|               ******** FORMATO DIGITAL ********                    |'
         '--------------------------------------------------------------------
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionModulo, "E", True) Then
        If fblnDatosValidos() Then
            'Si es modificación hacer... ---------------------------------
            If vglngNumeroTipoFormato = 2 Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
                If vlblnConsulta Then
                    rsFormatoDigital!VCHDESCRIPCION = Trim(txtDescripcion.Text)
                    If optAconceptoFacturacion.Value = True Then
                        rsFormatoDigital!inttipoagrupadigital = 1
                    End If
                    If optAcargo.Value = True Then
                        rsFormatoDigital!inttipoagrupadigital = 2
                    End If
                    If optDesglozadoCargo.Value = True Then
                        rsFormatoDigital!inttipoagrupadigital = 3
                    End If
                    rsFormatoDigital!vchDescripcionAgrupa1 = Trim(txtNombreACrystal1.Text)
                    rsFormatoDigital!vchDescripcionAgrupa2 = Trim(txtNombreACrystal2.Text)
                    rsFormatoDigital!vchDescripcionAgrupa3 = Trim(txtNombreACrystal3.Text)
                    rsFormatoDigital!vchDescripcionAgrupa5 = Trim(txtNombreACrystal5.Text)
'                    rsFormato.Update
'                    rsFormato.Requery
                    rsFormatoDigital.Update
                    rsFormatoDigital.Requery
                Else
                    'Si es registro nuevo; hacer... ---------------------------------
                    With rsFormatoDigital
                        .AddNew
                        !VCHDESCRIPCION = Trim(txtDescripcion.Text)
                        !intNumeroTipoFormato = vglngNumeroTipoFormato
                        !bitComprobantefiscaldigital = IIf(chkComprobanteFiscal.Value = vbChecked, 1, 0)
                        If optAconceptoFacturacion.Value = True Then
                            !inttipoagrupadigital = 1
                        End If
                        If optAcargo.Value = True Then
                            !inttipoagrupadigital = 2
                        End If
                        If optDesglozadoCargo.Value = True Then
                            !inttipoagrupadigital = 3
                        End If
                        !vchDescripcionAgrupa1 = Trim(txtNombreACrystal1.Text)
                        !vchDescripcionAgrupa2 = Trim(txtNombreACrystal2.Text)
                        !vchDescripcionAgrupa3 = Trim(txtNombreACrystal3.Text)
                        !vchDescripcionAgrupa5 = Trim(txtNombreACrystal5.Text)
                    End With
'                    rsFormato.Update
'                    rsFormato.Requery
                    rsFormatoDigital.Update
                    rsFormatoDigital.Requery
                    txtClave.Text = flngObtieneIdentity("SEC_FORMATO", rsFormatoDigital!intNumeroFormato)
                End If
                Call pGuardarLogTransaccion(Me.Name, IIf(vlblnConsulta, EnmGrabar, EnmCambiar), vllngPersonaGraba, "FORMATO DE IMPRESION", txtClave.Text)
                EntornoSIHO.ConeccionSIHO.CommitTrans
                pNumeroRegistros
                txtClave.SetFocus
            End If
            
'****************************** Diferencia entre tipo de formato (Factura o Nota)********************************

            'Si es modificación hacer... ---------------------------------
            If vglngNumeroTipoFormato = 8 Or vglngNumeroTipoFormato = 9 Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
                If vlblnConsulta Then
                    rsFormatoDigital!VCHDESCRIPCION = Trim(txtDescripcion.Text)
                    rsFormatoDigital!vchDescripcionAgrupa4 = Trim(txtNombreACrystal4.Text)
                    rsFormatoDigital!vchDescripcionAgrupa5 = Trim(txtNombreACrystal5.Text)
'                    rsFormato.Update
'                    rsFormato.Requery
                    rsFormatoDigital.Update
                    rsFormatoDigital.Requery
                Else
                    'Si es registro nuevo; hacer... ---------------------------------
                    With rsFormatoDigital
                        .AddNew
                        !VCHDESCRIPCION = Trim(txtDescripcion.Text)
                        !intNumeroTipoFormato = vglngNumeroTipoFormato
                        !bitComprobantefiscaldigital = IIf(chkComprobanteFiscal.Value = vbChecked, 1, 0)
                        !vchDescripcionAgrupa4 = Trim(txtNombreACrystal4.Text)
                        !vchDescripcionAgrupa5 = Trim(txtNombreACrystal5.Text)
                    End With
'                    rsFormato.Update
'                    rsFormato.Requery
                    rsFormatoDigital.Update
                    rsFormatoDigital.Requery
                    txtClave.Text = flngObtieneIdentity("SEC_FORMATO", rsFormatoDigital!intNumeroFormato)
                End If
                Call pGuardarLogTransaccion(Me.Name, IIf(vlblnConsulta, EnmGrabar, EnmCambiar), vllngPersonaGraba, "FORMATO DE IMPRESION", txtClave.Text)
                EntornoSIHO.ConeccionSIHO.CommitTrans
                pNumeroRegistros
                txtClave.SetFocus
            End If
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub
Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    Dim X As Integer
    
    fblnDatosValidos = True
    
    If Trim(txtClave.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtClave.SetFocus
    End If
    If fblnDatosValidos And Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    If fblnDatosValidos And cboTamano.ListIndex = -1 And chkComprobanteFiscal.Value = vbUnchecked Then
        fblnDatosValidos = False
        'Seleccione el dato.
        MsgBox SIHOMsg(431), vbExclamation + vbOKOnly, "Mensaje"
        cboTamano.SetFocus
    End If
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbUnchecked Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        fblnDatosValidos = llngPersonaGraba <> 0
    End If
    
'******************Agregado validacion de datos sobre el comprobante fiscal digital
If vglngNumeroTipoFormato = 2 Then
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked And optAconceptoFacturacion.Value = True And Trim(txtNombreACrystal1.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtNombreACrystal1.SetFocus
    End If
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked And optAcargo.Value = True And Trim(txtNombreACrystal2.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtNombreACrystal2.SetFocus
    End If
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked And optDesglozadoCargo.Value = True And Trim(txtNombreACrystal3.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtNombreACrystal3.SetFocus
    End If
        If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        fblnDatosValidos = llngPersonaGraba <> 0
    End If
End If

'AGREGADO PARA FORMATOS DE NOTAS DIGITALES
If vglngNumeroTipoFormato = 8 Or vglngNumeroTipoFormato = 9 Then
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked And Trim(txtNombreACrystal4.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtNombreACrystal4.SetFocus
    End If
        If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        fblnDatosValidos = llngPersonaGraba <> 0
    End If
End If
'FIN AGREGADO PARA FORMATOS DE NOTAS DIGITALES

'Fin agregado
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))

End Function

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
'Agregado
'Si no es un comprobante fiscal digital; hacer...
If chkComprobanteFiscal.Value = vbUnchecked Then
    If vllngNumeroRegistros <> 0 Then
        rsFormato.MoveFirst
        Do While rsFormato!intNumeroTipoFormato <> vglngNumeroTipoFormato
            rsFormato.MoveNext
        Loop
        pMuestra
        pHabilita 1, 1, 1, 1, 1, 0, 1
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
    End If
'Si no; hacer
Else
    If vllngNumeroRegistros <> 0 Then
        rsFormatoDigital.MoveFirst
        Do While rsFormatoDigital!intNumeroTipoFormato <> vglngNumeroTipoFormato
            rsFormatoDigital.MoveNext
        Loop
        pMuestra
        pHabilita 1, 1, 1, 1, 1, 0, 1
    Else
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
    End If
End If
'Fin de agregado

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    If rsContenidoFormato.RecordCount = 0 Then
        'No existe contenido para el tipo de formato.
        MsgBox SIHOMsg(271), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))

End Sub

Private Sub pCargaTamanos()
    On Error GoTo NotificaError

    cboTamano.AddItem "8", 0
    cboTamano.ItemData(cboTamano.newIndex) = 8
    cboTamano.AddItem "9", 1
    cboTamano.ItemData(cboTamano.newIndex) = 9
    cboTamano.AddItem "10", 2
    cboTamano.ItemData(cboTamano.newIndex) = 10
    cboTamano.AddItem "11", 3
    cboTamano.ItemData(cboTamano.newIndex) = 11
    cboTamano.AddItem "12", 4
    cboTamano.ItemData(cboTamano.newIndex) = 12
    cboTamano.AddItem "14", 5
    cboTamano.ItemData(cboTamano.newIndex) = 14
    cboTamano.AddItem "16", 6
    cboTamano.ItemData(cboTamano.newIndex) = 16
    cboTamano.AddItem "18", 7
    cboTamano.ItemData(cboTamano.newIndex) = 18
    cboTamano.AddItem "20", 8
    cboTamano.ItemData(cboTamano.newIndex) = 20
    cboTamano.AddItem "22", 9
    cboTamano.ItemData(cboTamano.newIndex) = 22
    cboTamano.AddItem "24", 10
    cboTamano.ItemData(cboTamano.newIndex) = 24
    cboTamano.AddItem "26", 11
    cboTamano.ItemData(cboTamano.newIndex) = 26
    cboTamano.AddItem "28", 12
    cboTamano.ItemData(cboTamano.newIndex) = 28
    cboTamano.AddItem "36", 13
    cboTamano.ItemData(cboTamano.newIndex) = 36
    cboTamano.AddItem "48", 14
    cboTamano.ItemData(cboTamano.newIndex) = 48
    cboTamano.AddItem "72", 15
    cboTamano.ItemData(cboTamano.newIndex) = 72
    cboTamano.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTamanos"))
End Sub

Private Sub pCargaContenido()
    
    
    grdContenido.Rows = 1
    vlstrsql = "select vchDescripcion, 0 as x, 0 as y, intAnchoContenidodefault, intNumeroContenidoFormato from ContenidoFormato where intNumeroTipoFormato = " & vglngNumeroTipoFormato & " and bitAgruparCargos = " & IIf(chkAgrupadoCargos.Value = vbChecked, "1", "0") & " order by vchDescripcion"
       
    Set rsContenidoFormato = frsRegresaRs(vlstrsql)
    If rsContenidoFormato.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdContenido, rsContenidoFormato, 4
        pConfigura
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Dim rsTipoFormato As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    blnCargando = False
    'Tamaño de fuente:
    pCargaTamanos
    
    pCargaContenido

    vlstrsql = "select vchDescripcion from TipoFormato where intNumeroTipoFormato=" + Str(vglngNumeroTipoFormato)
    Set rsTipoFormato = frsRegresaRs(vlstrsql)
    If rsTipoFormato.RecordCount <> 0 Then
        'Me.Caption = "Mantenimiento de formatos de " & LCase(rsTipoFormato!vchDescripcion)
        Me.Caption = "Formatos de " & LCase(rsTipoFormato!VCHDESCRIPCION)
    End If
    
'AGREGADO PARA MOSTRAR LAS OPCIONES DE COMPROBANTE FISCAL DIGITAL SEGÚN EL TIPO DE DOCUMENTO
If vglngNumeroTipoFormato = 2 Or vglngNumeroTipoFormato = 9 Or vglngNumeroTipoFormato = 8 Then
    chkComprobanteFiscal.Enabled = True
Else
    chkComprobanteFiscal.Enabled = False
End If
'FIN AGREGADO

    
    lblRenglones.Visible = vglngNumeroTipoFormato = 2
    txtRenglones.Visible = vglngNumeroTipoFormato = 2
    chkAgrupadoCargos.Visible = vglngNumeroTipoFormato = 2
    'lblNombreACrystal.Visible = vglngNumeroTipoFormato = 2

'AGREGADO
If vglngNumeroTipoFormato = 2 Then
    lblNombreACrystal.Visible = True
    optAconceptoFacturacion.Visible = True
    optAcargo.Visible = True
    optDesglozadoCargo.Visible = True
    txtNombreACrystal1.Visible = True
    txtNombreACrystal2.Visible = True
    txtNombreACrystal3.Visible = True
    lblNombreACrystal2.Visible = False
    txtNombreACrystal4.Visible = False
ElseIf vglngNumeroTipoFormato = 8 Or vglngNumeroTipoFormato = 9 Then
    'TEST PARA EL FORMATO DE NOTAS
    lblNombreACrystal2.Visible = True
    txtNombreACrystal4.Visible = True
    lblNombreACrystal.Visible = False
    optAconceptoFacturacion.Visible = False
    optAcargo.Visible = False
    optDesglozadoCargo.Visible = False
    txtNombreACrystal1.Visible = False
    txtNombreACrystal2.Visible = False
    txtNombreACrystal3.Visible = False
End If
    'END TEST
'FIN AGREGADO
    pHabilita 1, 1, 1, 1, 1, 0, 0
    pNumeroRegistros
    
    '-----------------------------
    ' Tablas
    '-----------------------------
'Se agregó en el query del vlstrsql "where bitComprobanteFiscalDigital = 0"
    vlstrsql = "select * from Formato where bitComprobanteFiscalDigital = 0"
    Set rsFormato = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    vlstrsql = "select * from FormatoDetalle"
    Set rsFormatoDetalle = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
'Agregado para el rsFormatoDigital
    vlstrsql = "select * from Formato where bitComprobanteFiscalDigital = 1"
    Set rsFormatoDigital = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    pUltimaClave
'Fin de agregado
    
    SSTabFormato.Tab = 0
    chkAgrupadoCargos_Click
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pNumeroRegistros()
    On Error GoTo NotificaError
    
    Dim rsNumeroRegistros As New ADODB.Recordset
    
    vlstrsql = "select count(*) from Formato where intNumeroTipoFormato=" + Str(vglngNumeroTipoFormato)
    Set rsNumeroRegistros = frsRegresaRs(vlstrsql)
    vllngNumeroRegistros = rsNumeroRegistros.Fields(0)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pNumeroRegistros"))
End Sub

Private Sub pConfigura()
    On Error GoTo NotificaError
    
    With grdContenido
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Descripción|Columna|Renglón|Ancho (*)"
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignmentFixed(1) = flexAlignLeftCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColWidth(0) = 100
        .ColWidth(1) = 6000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 0
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfigura"))

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If SSTabFormato.Tab = 0 Then
    If cmdSave.Enabled = True Then
        Cancel = True
        ' ¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            txtCoordenada.Visible = False
            txtCoordenada.Text = ""
            txtNombreACrystal1.Text = ""
            txtNombreACrystal2.Text = ""
            txtNombreACrystal3.Text = ""
            txtNombreACrystal4.Text = ""
            chkComprobanteFiscal.Value = vbUnchecked
            optAconceptoFacturacion.Value = True
            txtClave.SetFocus
        Else
            Cancel = True
            txtDescripcion.SetFocus
            Exit Sub
        End If
    Else
        Unload Me
    End If
Else
    If SSTabFormato.Tab = 1 Then
        SSTabFormato.Tab = 0
        txtClave.SetFocus
    End If
    Cancel = True
End If

End Sub

Private Sub grdContenido_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdContenido_GotFocus"))

End Sub

Private Sub grdContenido_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii <> 27 Then
    If grdContenido.Col = 2 Or grdContenido.Col = 3 Or grdContenido.Col = 4 Then
        txtCoordenada.Move grdContenido.Left + grdContenido.CellLeft, grdContenido.Top + grdContenido.CellTop, grdContenido.CellWidth - 8, grdContenido.CellHeight - 8
        If IsNumeric(Chr(KeyAscii)) Then
            txtCoordenada.Text = Chr(KeyAscii)
        Else
            txtCoordenada.Text = grdContenido.TextMatrix(grdContenido.Row, grdContenido.Col)
        End If
        txtCoordenada.Visible = True
        txtCoordenada.SelStart = Len(txtCoordenada.Text)
        txtCoordenada.SetFocus
    End If
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdContenido_KeyPress"))

End Sub

Private Sub optAcargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtNombreACrystal1.SetFocus
    End If
End Sub

Private Sub optAconceptoFacturacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtNombreACrystal1.SetFocus
    End If
End Sub

Private Sub optDesglozadoCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtNombreACrystal1.SetFocus
    End If
End Sub

Private Sub txtclave_GotFocus()
    On Error GoTo NotificaError
    
'Agregado (si esta activa la casilla de comprobante digital, desactivarla)
If chkComprobanteFiscal.Value = vbChecked Then
chkComprobanteFiscal.Value = vbUnchecked
End If
'Fin agregado

pLimpia

    pHabilita 1, 1, 1, 1, 1, 0, 0
    pSelTextBox txtClave
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_GotFocus"))

End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    Dim X As Integer
    
    vlblnConsulta = False
    
    pUltimaClave
    txtDescripcion.Text = ""
    txtTotalCols = ""
    txtTotalLineas = ""
    txtRenglones = ""
'Agregado por la opcion de comprobante fiscal digital
    grdContenido.Visible = True
    txtTotalCols.Visible = True
    txtTotalLineas.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    cboTamano.Visible = True
'AGREGADO
If vglngNumeroTipoFormato = 2 Then
    lblNombreACrystal.Visible = True
    chkAgrupadoCargos.Visible = True
    lblRenglones.Visible = True
    txtRenglones.Visible = True
End If
If vglngNumeroTipoFormato = 9 Then
    lblNombreACrystal2.Visible = True
    txtNombreACrystal4.Visible = True
End If
'FIN AGREGADO

'TEST PARA EL FORMATO DE NOTAS
If vglngNumeroTipoFormato = 8 Then
    lblNombreACrystal2.Visible = True
    txtNombreACrystal4.Visible = True
End If
'END TEST

    Frame4.Visible = False
    optAconceptoFacturacion.Value = False
    txtNombreACrystal1.Text = ""
    txtNombreACrystal2.Text = ""
    txtNombreACrystal3.Text = ""
    txtNombreACrystal4.Text = ""
' Adicional
    chkAgrupadoCargos.Value = vbUnchecked
' Fin de agregado
    
    For X = 1 To grdContenido.Rows - 1
        grdContenido.TextMatrix(X, 2) = ""
        grdContenido.TextMatrix(X, 3) = ""
        grdContenido.TextMatrix(X, 4) = ""
    Next X
    grdContenido.Row = 1
    grdContenido.Col = 2
    
    cboTamano.ListIndex = -1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))

End Sub

Private Sub pUltimaClave()
    On Error GoTo NotificaError
    
    Dim rsUltimaClave As New ADODB.Recordset
    
    vlstrsql = "select max(intNumeroFormato) as Ultimo from Formato"
    Set rsUltimaClave = frsRegresaRs(vlstrsql)
    If IsNull(rsUltimaClave!Ultimo) Then
        txtClave.Text = "1"
    Else
        txtClave.Text = Str(rsUltimaClave!Ultimo + 1)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pUltimaClave"))

End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    'Agregado
    If KeyAscii = 13 And txtClave.Text = "" Then
        pUltimaClave
    End If
    'Fin de agregado
    
    If KeyAscii = 13 Then
        pBusca txtClave.Text
        If vlblnConsulta Then
            pHabilita 0, 0, 0, 0, 0, 1, 1
            cmdSave.SetFocus
        Else
            pHabilita 0, 0, 0, 0, 0, 1, 0
            'AGREGADO EN CASO DE QUE EL DOCUMENTO NO TENGA LAS OPCIONES DIGITALES DISPONIBLES
            If chkComprobanteFiscal.Enabled = False Then
                txtDescripcion.SetFocus
            Else
                chkComprobanteFiscal.SetFocus
            End If
            'FIN AGREGADO
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_KeyPress"))

End Sub

Private Sub pBusca(vlstrxClaveFormato As String)
    On Error GoTo NotificaError
    
    Dim vllngNumeroFormato As Long

    'Si es digital hacer...
    vllngNumeroFormato = fintLocalizaPkRs(rsFormatoDigital, 0, txtClave.Text)
    If vllngNumeroFormato = 0 Then
    'Si no es digital hacer...
        vllngNumeroFormato = fintLocalizaPkRs(rsFormato, 0, txtClave.Text)
        If vllngNumeroFormato = 0 Then
            pLimpia
        Else
            If rsFormato!intNumeroTipoFormato = vglngNumeroTipoFormato Then
                pMuestra
            Else
                pLimpia
            End If
        End If
    Else
        If vllngNumeroFormato = 0 Then
            pLimpia
        Else
            If rsFormatoDigital!intNumeroTipoFormato = vglngNumeroTipoFormato Then
                chkComprobanteFiscal.Value = vbChecked
                pMuestra
            Else
                pLimpia
            End If
        End If

    End If
    
                '    'ORIGINAL
                '        vllngNumeroFormato = fintLocalizaPkRs(rsFormato, 0, txtClave.Text)
                '    If vllngNumeroFormato = 0 Then
                '        pLimpia
                '    Else
                '        If rsFormato!intNumeroTipoFormato = vglngNumeroTipoFormato Then
                '            pMuestra
                '        Else
                '            pLimpia
                '        End If
                '    End If
                '    'FIN ORIGINAL

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBusca"))

End Sub
Private Sub pBuscaMuestra(vlstrxClaveFormato As String)
On Error GoTo NotificaError
Dim vllngNumeroFormato2 As Long
'Para reiniciar la busqueda y no se quede buscando solo los NO DIGITALES
chkComprobanteFiscal.Value = vbUnchecked
    'Si es digital hacer...
    vllngNumeroFormato2 = fintLocalizaPkRs(rsFormatoDigital, 0, Val(vlstrxClaveFormato))
    If vllngNumeroFormato2 = 0 Then
    'Si no es digital hacer...
        vllngNumeroFormato2 = fintLocalizaPkRs(rsFormato, 0, Val(vlstrxClaveFormato))
        If vllngNumeroFormato2 = 0 Then
            pLimpia
        Else
            If rsFormato!intNumeroTipoFormato = vglngNumeroTipoFormato Then '<-- 2 Then <-- vglngNumeroTipoFormato2 Then
                pMuestra
            Else
                pLimpia
            End If
        End If
    Else
        If vllngNumeroFormato2 = 0 Then
            pLimpia
        Else
            If rsFormatoDigital!intNumeroTipoFormato = vglngNumeroTipoFormato Then '<-- 2 Then <-- vglngNumeroTipoFormato2 Then= 2 Then 'vglngNumeroTipoFormato2 Then
                chkComprobanteFiscal.Value = vbChecked
                pMuestra
            Else
                pLimpia
            End If
        End If
    End If
Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBuscaMuestra"))
End Sub

Private Sub pMuestra()
    On Error GoTo NotificaError
    
    Dim rsDetalledelFormato As New ADODB.Recordset
    Dim X As Integer
    
    blnCargando = True
    vlblnConsulta = True
'Agregado
'Si es un comprobante fiscal NO digital hacer...
If chkComprobanteFiscal.Value = vbUnchecked Then
    txtClave.Text = rsFormato!intNumeroFormato
    txtDescripcion.Text = rsFormato!VCHDESCRIPCION
    txtTotalCols.Text = IIf(IsNull(rsFormato!intAncho), 0, rsFormato!intAncho)
    txtTotalLineas.Text = IIf(IsNull(rsFormato!intLargo), 0, rsFormato!intLargo)
    txtRenglones.Text = IIf(IsNull(rsFormato!intRenglonesDetalle), "", rsFormato!intRenglonesDetalle)
    cboTamano.ListIndex = flngLocalizaCbo(cboTamano, Str(rsFormato!INTLETRA))
    chkAgrupadoCargos = IIf(IsNull(rsFormato!bitAgruparCargos), vbUnchecked, IIf(rsFormato!bitAgruparCargos = 1, vbChecked, vbUnchecked))
    pMuestraDetalle
    blnCargando = False
'Si no (SI es un comprobante fiscal digital) hacer...
Else
    txtClave.Text = rsFormatoDigital!intNumeroFormato
    txtDescripcion.Text = rsFormatoDigital!VCHDESCRIPCION
    
    If rsFormatoDigital!inttipoagrupadigital = 1 Then
        optAconceptoFacturacion.Value = True
    End If
    If rsFormatoDigital!inttipoagrupadigital = 2 Then
        optAcargo.Value = True
    End If
    If rsFormatoDigital!inttipoagrupadigital = 3 Then
        optDesglozadoCargo.Value = True
    End If
    
    txtNombreACrystal1.Text = IIf(IsNull(rsFormatoDigital!vchDescripcionAgrupa1), "", rsFormatoDigital!vchDescripcionAgrupa1)
    txtNombreACrystal2.Text = IIf(IsNull(rsFormatoDigital!vchDescripcionAgrupa2), "", rsFormatoDigital!vchDescripcionAgrupa2)
    txtNombreACrystal3.Text = IIf(IsNull(rsFormatoDigital!vchDescripcionAgrupa3), "", rsFormatoDigital!vchDescripcionAgrupa3)
    txtNombreACrystal4.Text = IIf(IsNull(rsFormatoDigital!vchDescripcionAgrupa4), "", rsFormatoDigital!vchDescripcionAgrupa4)
    txtNombreACrystal5.Text = IIf(IsNull(rsFormatoDigital!vchDescripcionAgrupa5), "", rsFormatoDigital!vchDescripcionAgrupa5)
    'pMuestraDetalle
    blnCargando = False
End If
'Fin agregado

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra"))
End Sub

Private Sub pMuestraDetalle()
    On Error GoTo NotificaError
    
    Dim rsDetalledelFormato As New ADODB.Recordset
    Dim X As Integer
    
    pCargaContenido
    
    If txtClave.Text <> "" Then
        vgstrParametrosSP = txtClave.Text
        
        Set rsDetalledelFormato = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormatoDetalle")
        
        If rsDetalledelFormato.RecordCount <> 0 Then
            rsDetalledelFormato.MoveFirst
            Do While Not rsDetalledelFormato.EOF
                For X = 1 To grdContenido.Rows - 1
                    If grdContenido.RowData(X) = rsDetalledelFormato!numero Then
                        grdContenido.TextMatrix(X, 2) = IIf(IsNull(rsDetalledelFormato!intCoordenadaX), "", rsDetalledelFormato!intCoordenadaX)
                        grdContenido.TextMatrix(X, 3) = IIf(IsNull(rsDetalledelFormato!intCoordenadaY), "", rsDetalledelFormato!intCoordenadaY)
                        grdContenido.TextMatrix(X, 4) = IIf(IsNull(rsDetalledelFormato!intAnchoContenido), "", rsDetalledelFormato!intAnchoContenido)
                    End If
                Next X
                rsDetalledelFormato.MoveNext
            Loop
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraDetalle"))

End Sub

Private Sub txtCoordenada_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    Else
        If KeyAscii = 13 Then
            grdContenido.TextMatrix(grdContenido.Row, grdContenido.Col) = txtCoordenada.Text
            
            If grdContenido.Col <> 3 And grdContenido.Col <> 2 Then
                If grdContenido.Row + 1 < grdContenido.Rows Then
                    grdContenido.Col = 2
                    grdContenido.Row = grdContenido.Row + 1
                    grdContenido.SetFocus
                Else
                    cmdSave.SetFocus
                End If
            Else
                If grdContenido.Col = 2 Then
                    grdContenido.Col = 3
                    grdContenido.SetFocus
                Else
                    grdContenido.Col = 4
                    grdContenido.SetFocus
                End If
            End If
                        
            txtCoordenada.Visible = False
            
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCoordenada_KeyPress"))
End Sub

Private Sub txtCoordenada_LostFocus()
    On Error GoTo NotificaError

    txtCoordenada.Visible = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCoordenada_LostFocus"))
End Sub

Private Sub txtDescripcion_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescripcion

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_GotFocus"))

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
'Si no esta checada la opcion de comprobante fiscal hacer...
    If KeyAscii = 13 And chkComprobanteFiscal.Value = vbUnchecked Then
        txtTotalCols.SetFocus
    End If
'Si no...
    If KeyAscii = 13 And chkComprobanteFiscal.Value = vbChecked Then
        If vglngNumeroTipoFormato = 2 Then
            optAconceptoFacturacion.SetFocus
        ElseIf vglngNumeroTipoFormato = 8 Or vglngNumeroTipoFormato = 9 Then
            txtNombreACrystal4.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))

End Sub
Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError
    
    Dim vlintTexto As Integer
    
    ' Posiciona el texto
    
    With txtEdit
       .Text = grdContenido.TextMatrix(grdContenido.Row, grdContenido.Col) + Chr(KeyAscii) 'Inicialización del Textbox
        Select Case KeyAscii
            Case 0 To 32
                'Edita el texto de la celda en la que está posicionado
                    .SelStart = 0
                    .SelLength = 1000
            Case 8, 46, 48 To 57
                ' Reemplaza el texto actual solo si se teclean números
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
        End Select
    End With
    
    ' Muestra el textbox en el lugar indicado
    With grid
        If .CellWidth < 0 Then Exit Sub
            txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
    vgstrEstadoManto = vgstrEstadoManto & "E"
    txtEdit.Visible = True
    txtEdit.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEditarColumna"))

End Sub

Private Sub txtNombreACrystal1_GotFocus()
pSelTextBox txtNombreACrystal1
pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub txtNombreACrystal1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            txtNombreACrystal2.SetFocus
    End If
End Sub

Private Sub txtNombreACrystal2_GotFocus()
pSelTextBox txtNombreACrystal2
pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub txtNombreACrystal2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            txtNombreACrystal3.SetFocus
    End If
End Sub

Private Sub txtNombreACrystal3_GotFocus()
pSelTextBox txtNombreACrystal3
pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub txtNombreACrystal3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            txtNombreACrystal5.SetFocus
    End If
End Sub

Private Sub txtNombreACrystal4_GotFocus()
pSelTextBox txtNombreACrystal4
pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub txtNombreACrystal4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            txtNombreACrystal5.SetFocus
    End If
End Sub

Private Sub txtNombreACrystal5_GotFocus()
pSelTextBox txtNombreACrystal5
pHabilita 0, 0, 0, 0, 0, 1, 0

End Sub

Private Sub txtNombreACrystal5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            cmdSave.SetFocus
    End If
End Sub

Private Sub txtRenglones_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtRenglones
End Sub

Private Sub txtRenglones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdContenido.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub txtTotalCols_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtTotalCols
End Sub

Private Sub txtTotalCols_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
               
    If KeyAscii = 13 Then
        txtTotalLineas.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtTotalCols_KeyPress"))
End Sub

Private Sub txtTotalLineas_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtTotalLineas
End Sub

Private Sub txtTotalLineas_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
               
    If KeyAscii = 13 Then
        cboTamano.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtTotalLineas_KeyPress"))

End Sub

Private Sub vsfFormato_DblClick()
'    Dim vllngNumeroFormato As Long
    Call pBuscaMuestra(Val(vsfFormato.TextMatrix(vsfFormato.Row, 1)))
    SSTabFormato.Tab = 0
    pHabilita 1, 1, 1, 1, 1, 0, 1
    
                ''ORIGINAL
                ''Agregado
                ''Si NO es comprobante fiscal digital, hacer...
                'If chkComprobanteFiscal.Value = vbUnchecked Then
                '    If Val(vsfFormato.TextMatrix(vsfFormato.Row, 1)) <> 0 Then
                '        vllngNumeroFormato = fintLocalizaPkRs(rsFormato, 0, Str(Val(vsfFormato.TextMatrix(vsfFormato.Row, 1))))
                '        SSTabFormato.Tab = 0
                '        If vllngNumeroFormato = 0 Then
                '            txtClave.SetFocus
                '        Else
                '            pMuestra
                '            pHabilita 1, 1, 1, 1, 1, 0, 1
                '        End If
                '    End If
                '
                ''Si no; hacer...
                'Else
                '    If Val(vsfFormato.TextMatrix(vsfFormato.Row, 1)) <> 0 Then
                '        vllngNumeroFormato = fintLocalizaPkRs(rsFormatoDigital, 0, Str(Val(vsfFormato.TextMatrix(vsfFormato.Row, 1))))
                '        SSTabFormato.Tab = 0
                '        If vllngNumeroFormato = 0 Then
                '            txtClave.SetFocus
                '        Else
                '            pMuestra
                '            pHabilita 1, 1, 1, 1, 1, 0, 1
                '        End If
                '    End If
                'End If
                ''Fin de agregado
                ''FIN ORIGINAL
End Sub

Private Sub vsfFormato_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        vsfFormato_DblClick
    End If

End Sub
