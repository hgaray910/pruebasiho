VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmListasPreciosPemex 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de precios Pemex"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8055
   ScaleMode       =   0  'User
   ScaleWidth      =   14130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   120
      TabIndex        =   17
      Top             =   8520
      Visible         =   0   'False
      Width           =   13890
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   13800
         _ExtentX        =   24342
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblTextoBarra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cargando datos, por favor espere..."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   11250
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   670
      Left            =   6600
      TabIndex        =   16
      Top             =   7270
      Width           =   1125
      Begin VB.CommandButton cmdCancelaOrd 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00DCDCDC&
         Picture         =   "frmListasPreciosPemex.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Eliminar registro"
         Top             =   140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   495
         Left            =   52
         Picture         =   "frmListasPreciosPemex.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Guardar el Registro"
         Top             =   140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripción completa del cargo"
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
      Height          =   795
      Left            =   120
      TabIndex        =   14
      Top             =   6450
      Width           =   13890
      Begin VB.Label lblNombreCompleto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   390
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Descripción completa del cargo"
         Top             =   270
         Width           =   13665
      End
   End
   Begin VB.Frame cmdImprimir2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   13890
      Begin VB.Timer tmrDespliega 
         Interval        =   1000
         Left            =   6240
         Top             =   1500
      End
      Begin VB.TextBox txtcargo 
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
         Left            =   2190
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Nombre del cargo a buscar"
         Top             =   240
         Width           =   7545
      End
      Begin MyCommandButton.MyButton cmdExportar 
         Height          =   375
         Left            =   11640
         TabIndex        =   9
         ToolTipText     =   "Permite exportar la información filtrada en formato Excel"
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
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
         TransparentColor=   16777215
         Caption         =   "Exportar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdImportar 
         Height          =   375
         Left            =   12689
         TabIndex        =   6
         ToolTipText     =   "Permite importar información de precios desde un archivo Excel"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         TransparentColor=   16777215
         Caption         =   "Importar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9480
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPrecios 
         Height          =   4485
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Lista de precios de PEMEX"
         Top             =   840
         Width           =   13665
         _ExtentX        =   24104
         _ExtentY        =   7911
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
      Begin MyCommandButton.MyButton cmdBuscar 
         Height          =   375
         Left            =   9990
         TabIndex        =   4
         ToolTipText     =   "Permite buscar cargo en la lista de precios"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         TransparentColor=   16777215
         Caption         =   "Buscar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAgregar 
         Height          =   375
         Left            =   11320
         TabIndex        =   5
         ToolTipText     =   "Permite agregar nuevo cargo a la lista de precios"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         TransparentColor=   16777215
         Caption         =   "Agregar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblDescripcionCargo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Descripción cargo"
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
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   2250
      End
   End
   Begin VB.Frame freDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   770
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   13890
      Begin VB.TextBox Text1 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Nombre del departamento"
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.PictureBox SysInfo1 
      Height          =   480
      Left            =   2040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   10200
      Width           =   1200
   End
End
Attribute VB_Name = "frmListasPreciosPemex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmListasPreciosPemex
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento de listas de precios unicamente de pemex
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        :
'| Autor                    :
'| Fecha de Creación        :
'| Modificó                 : Nombre(s)
'| Fecha terminación        :
'| Fecha última modificación:
'-------------------------------------------------------------------------------------
Option Explicit


'Columnas del grid de precios:
Const cintColIDConsecutivo = 1
Const cintColCodigoPCE = 2
Const cintColDescripcion = 3
Const cintColPrecioLayout = 4
Const cintColCantidad = 5
Const cintColPrecioUnidosis = 6
Const cintColumnas = 7

Const clngRojo = &HC0&
Const clngAzul = &HC00000



Const clngLargoRenglon = 240

Public WithEvents grid As MSHFlexGrid
Attribute grid.VB_VarHelpID = -1
Public WithEvents txtEdit As TextBox
Attribute txtEdit.VB_VarHelpID = -1
Public vpblnEsSistemas As Boolean

Private vgrptReporte As CRAXDRT.Report
Dim vgblnNoEditar As Boolean

Dim vgblnEditaPrecio As Boolean 'Para saber si se esta editando una cantidad
Dim vgblnEditaUtilidad As Boolean 'Para saber si se esta editando el margen de utilidad
Dim vgstrEstadoManto As String 'Estatus para saber donde ando en la pantalla
Dim vglngDesktop As Long     'Para saber el tamaño del desktop
Dim vlstrSentencia As String
Dim vlblnCancelCaptura As Boolean  ' Para no permitir capturar si no hay elementos en la lista
Dim vlblnRecalcularImportar As Boolean  ' Para saber si se esta importando

Dim lblnConsulta As Boolean
Dim llngMarcados As Long
Dim blnactivainicializa As Integer
Dim vlstrCveCargo As String

Dim rsListasDepto As New ADODB.Recordset 'Listas de precios del departamento

'-----------------------------------
' vgstrEstadoManto puede tener los siguientes valores
' ""  Nuevo registro, default
' "A" Una Alta de un Elemento
' "B" Que esta en la pantalla de búsqueda
' "M" Una modificacion o Consulta
' "ME" Esta editando un precio en una consulta
' "AE" Esta editando un precio en una alta
'-----------------------------------
Const cgintFactorMovVentana = 150
Const cgIntAltoVentanaMax = 10410
Const cgIntAltoVentanaMin = 3240

Public blnCatalogo As Boolean
Public StrCveArticulo As String
Dim lblnPermisoCosto As Boolean 'tiene permiso para modificar el costo








Private Sub cmdAgregar_Click()

     


    frmCargoPrecioPemex.vlstrChrCveArticulo = ""
                frmCargoPrecioPemex.txtDescripcionLgaArt = ""
                frmCargoPrecioPemex.txtCodigo = ""
                frmCargoPrecioPemex.txtPrecio = ""
                'CStr(Val(Format(Trim(grdPrecios.TextMatrix(i, 4)), "##########0.00####")))
                frmCargoPrecioPemex.txtCantidad = ""
                frmCargoPrecioPemex.txtpreciounidosis = ""

    Load frmCargoPrecioPemex
    frmCargoPrecioPemex.Show vbModal, Me

    pLimpiaSeleccionGrid
        cmdGrabarRegistro.Enabled = False
        cmdCancelaOrd.Enabled = False
        txtcargo.Text = ""
        lblNombreCompleto = ""
        pConfiguraGrid
        txtcargo.SetFocus


End Sub

Private Sub cmdBuscar_Click()
Dim vlstrSentencia As String
    Dim rsConceptoFacturacion As New ADODB.Recordset
    Dim rsbusqueda As New ADODB.Recordset
    Dim vlintRenglon As Integer
    Dim totalecontrado As Integer
    lblNombreCompleto.Caption = ""
     pLimpiaSeleccionGrid
                cmdGrabarRegistro.Enabled = False
                cmdCancelaOrd.Enabled = False
    vlstrSentencia = "select INTIDCONSECUTIVO, CHRCODIGOPCE,VCHDESCRIPCION,  MNYPRECIOLAYOUT,INTCANTIDAD, MNYPRECIOUNIDOSIS from PVLISTAPEMEXPRECIOS where VCHDESCRIPCION like '%" & Trim(txtcargo.Text) & "%' "
    
    Set rsbusqueda = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    totalecontrado = rsbusqueda.RecordCount
    If rsbusqueda.RecordCount > 0 Then
    
        With rsbusqueda
                .MoveFirst
    
    
                vlintRenglon = 1
                 If totalecontrado = 1 Then
                    lblNombreCompleto.Caption = !VCHDESCRIPCION
                 End If
    
                Do While Not .EOF
                        grdPrecios.TextMatrix(vlintRenglon, cintColIDConsecutivo) = !INTIDCONSECUTIVO
                        grdPrecios.TextMatrix(vlintRenglon, cintColCodigoPCE) = !CHRCODIGOPCE
                        grdPrecios.TextMatrix(vlintRenglon, cintColDescripcion) = !VCHDESCRIPCION
                        grdPrecios.TextMatrix(vlintRenglon, cintColPrecioLayout) = Format(Trim(!MNYPRECIOLAYOUT), "$###,###,###,##0.00####")
                        grdPrecios.TextMatrix(vlintRenglon, cintColCantidad) = !intCantidad
                        grdPrecios.TextMatrix(vlintRenglon, cintColPrecioUnidosis) = Format(Trim(!MNYPRECIOUNIDOSIS), "$###,###,###,##0.00")
                                           
                                                            
                       grdPrecios.Rows = grdPrecios.Rows + 1
    
                    vlintRenglon = vlintRenglon + 1
                    .MoveNext
                Loop
                    grdPrecios.Rows = grdPrecios.Rows - 1
            End With
    Else
    
     '|  No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        pLimpiaSeleccionGrid
        cmdGrabarRegistro.Enabled = False
        cmdCancelaOrd.Enabled = False
        txtcargo.Text = ""
        lblNombreCompleto.Caption = ""
        pConfiguraGrid
        txtcargo.SetFocus
    End If
    
End Sub


Private Sub CargarDatosGrid()
Dim vlstrSentencia As String
    Dim rsConceptoFacturacion As New ADODB.Recordset
    Dim rsbusqueda As New ADODB.Recordset
    Dim vlintRenglon As Integer
    
    vlstrSentencia = "select INTIDCONSECUTIVO, CHRCODIGOPCE,VCHDESCRIPCION,  MNYPRECIOLAYOUT,INTCANTIDAD, MNYPRECIOUNIDOSIS from PVLISTAPEMEXPRECIOS"
    
    Set rsbusqueda = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsbusqueda.RecordCount > 0 Then
    
    With rsbusqueda
            .MoveFirst


            vlintRenglon = 1

            Do While Not .EOF
                    grdPrecios.TextMatrix(vlintRenglon, cintColIDConsecutivo) = !INTIDCONSECUTIVO
                    grdPrecios.TextMatrix(vlintRenglon, cintColCodigoPCE) = !CHRCODIGOPCE
                    grdPrecios.TextMatrix(vlintRenglon, cintColDescripcion) = !VCHDESCRIPCION
                    grdPrecios.TextMatrix(vlintRenglon, cintColPrecioLayout) = Format(Trim(!MNYPRECIOLAYOUT), "$###,###,###,##0.00####")
                    grdPrecios.TextMatrix(vlintRenglon, cintColCantidad) = !intCantidad
                    grdPrecios.TextMatrix(vlintRenglon, cintColPrecioUnidosis) = Format(Trim(!MNYPRECIOUNIDOSIS), "$###,###,###,##0.00")
                                       
                   grdPrecios.Rows = grdPrecios.Rows + 1

                vlintRenglon = vlintRenglon + 1
                .MoveNext
            Loop
                grdPrecios.Rows = grdPrecios.Rows - 1
        End With
    
    End If
    
End Sub

Private Sub cmdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        grdPrecios.SetFocus
    End If
End Sub


Private Sub cmdCancelaOrd_Click()

    Dim i, lngPersonaGraba As Long
    Dim BlnTransaccion As Boolean
    BlnTransaccion = False
                     
    lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If lngPersonaGraba <> 0 Then
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
         For i = 1 To grdPrecios.Rows - 1
            If grdPrecios.TextMatrix(i, 0) = "*" Then
            

            'elimina registro de la tabla
             frsEjecuta_SP Trim(grdPrecios.TextMatrix(i, 1)), "SP_PVDELLISTAPRECIOPEMEX", True
            Exit For
           End If
        Next
                 
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, lngPersonaGraba, "LISTA DE PRECIOS PEMEX", CStr(vgintNumeroDepartamento))
        EntornoSIHO.ConeccionSIHO.CommitTrans
                   
         '|La operación se realizó satisfactoriamente.
        MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
          
        pLimpiaSeleccionGrid
        cmdGrabarRegistro.Enabled = False
        cmdCancelaOrd.Enabled = False
        txtcargo.Text = ""
        lblNombreCompleto.Caption = ""
        pConfiguraGrid
        txtcargo.SetFocus
     End If
End Sub

Private Sub cmdExportar_Click()
On Error GoTo NotificaError
'    Dim rsAux As New ADODB.Recordset
    Dim o_Excel As Object
'    Dim o_ExcelAbrir As Object
    Dim o_Libro As Object
    Dim o_Sheet As Object
    Dim intRow As Long
    Dim intCol As Integer
    Dim dblAvance As Double
    Dim intRowExcel As Long
        
    If grdPrecios.Rows > 1 And grdPrecios.TextMatrix(1, 1) <> "" Then
        Set o_Excel = CreateObject("Excel.Application")
        Set o_Libro = o_Excel.Workbooks.Add
        Set o_Sheet = o_Libro.Worksheets(1)
        
        If Not IsObject(o_Excel) Then
            MsgBox "Necesitas Microsoft Excel para utilizar esta funcionalidad", _
               vbExclamation, "Mensaje"
            Exit Sub
        End If
        
        'Columnas titulos
        o_Excel.Cells(1, 1).Value = "CodigoPCE"
        o_Excel.Cells(1, 2).Value = "Descripción"
        o_Excel.Cells(1, 3).Value = "Precio"
        o_Excel.Cells(1, 4).Value = "Cantidad"
        o_Excel.Cells(1, 5).Value = "Precio unidosis"

        
        o_Sheet.range("A1:E1").HorizontalAlignment = -4108
        o_Sheet.range("A1:E1").VerticalAlignment = -4108
        o_Sheet.range("A1:E1").WrapText = True
        o_Sheet.range("A2").Select
        o_Excel.ActiveWindow.FreezePanes = True
        o_Sheet.range("A1:E1").Interior.ColorIndex = 15 '15 48
        
        'o_Sheet.Range(o_Excel.Cells(grdPrecios.Rows + 7, 13), o_Excel.Cells(grdPrecios.Rows + 7, 23)).Interior.ColorIndex = 15
        o_Sheet.range("A:A").Columnwidth = 12
        o_Sheet.range("B:B").Columnwidth = 50
        o_Sheet.range("C:C").Columnwidth = 12
        o_Sheet.range("D:D").Columnwidth = 15
        o_Sheet.range("E:E").Columnwidth = 12

        
        'o_Sheet.Range(o_Excel.Cells(1, 1), o_Excel.Cells(grdPrecios.Rows, 15)).Borders(4).LineStyle = 1
        
        'info del rs
        o_Sheet.range("A:E").Font.Size = 9
        o_Sheet.range("A:E").Font.Name = "Times New Roman" '
        o_Sheet.range("A:E").Font.Bold = False
        
        'o_Sheet.Range("A:A").NumberFormat = "0000000000"
        'titulos
        o_Sheet.range("A1:X1").Font.Bold = True
        o_Sheet.range(o_Excel.Cells(grdPrecios.Rows + 7, 1), o_Excel.Cells(grdPrecios.Rows + 7, 23)).Font.Bold = True
        'o_Sheet.Range(o_Excel.Cells(2, 1), o_Excel.Cells(5, 1)).Font.Bold = True
        'centrado, auto ajustar texto, alinear medio
        o_Sheet.range("C:C").NumberFormat = "$ ###,###,###,##0.00"
       o_Sheet.range("E:E").NumberFormat = "$ ###,###,###,##0.00"
        
        
        
        dblAvance = 100 / grdPrecios.Rows
        '------------------------
        ' Configuración de la Barra de estado
        '------------------------
        lblTextoBarra.Caption = "Exportando información, por favor espere..."
        freBarra.Visible = True
        freBarra.Top = 720
        pgbBarra.Value = 0
        freBarra.Refresh
        pgbBarra.Value = 0
        
        intRowExcel = 2
        'Recorre el grid y llena el Excel
        For intRow = 2 To grdPrecios.Rows '- 1

            ' Actualización de la barra de estado
            If pgbBarra.Value + dblAvance < 100 Then
                pgbBarra.Value = pgbBarra.Value + dblAvance
            Else
                pgbBarra.Value = 100
            End If
            pgbBarra.Refresh

            If grdPrecios.RowHeight(intRow - 1) > 0 Then
                With grdPrecios
                    
                    Dim strTipo As String
'                    strTipo = .TextMatrix(intRow - 1, 10)
'
'                    If strTipo = "AR" Then
'                        o_Sheet.Cells(intRowExcel, 1).NumberFormat = "0000000000"
'                    Else
'                        o_Sheet.Cells(intRowExcel, 1).NumberFormat = "0"
'                    End If
                    
                    o_Sheet.Cells(intRowExcel, 1).Value = .TextMatrix(intRow - 1, 1) & " "
                    o_Sheet.Cells(intRowExcel, 2).Value = .TextMatrix(intRow - 1, 2) & " "
                    '
                    o_Sheet.Cells(intRowExcel, 3).Value = .TextMatrix(intRow - 1, 3) & ""
                    o_Sheet.Cells(intRowExcel, 4).Value = .TextMatrix(intRow - 1, 4) & " "
                    o_Sheet.Cells(intRowExcel, 5).Value = .TextMatrix(intRow - 1, 5) & " "

                End With
                intRowExcel = intRowExcel + 1
            End If
        Next
        
        'La información ha sido exportada exitosamente
        'MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
        freBarra.Visible = False
        o_Excel.Visible = True
        
        Set o_Excel = Nothing
        cmdImportar.SetFocus
        
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
        
Exit Sub
NotificaError:
    ' -- Cierra la hoja y la aplicación Excel
    freBarra.Visible = False
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Sheet Is Nothing Then Set o_Sheet = Nothing
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdExportar_Click"))
End Sub


Private Sub cmdGrabarRegistro_Click()

'    Dim lngContador As Integer
'
'   Dim vllngPersonaGraba As Long
'
'
'
'            vglngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
'            If vglngPersonaGraba <> 0 Then
'                EntornoSIHO.ConeccionSIHO.BeginTrans
'                    vlstrSentencia = "DELETE FROM PVLISTAPEMEXPRECIOS "
'                    pEjecutaSentencia vlstrSentencia
'
'                    For lngContador = 1 To grdPrecios.Rows - 1
'                        With grdPrecios
'
'                            If Trim(.TextMatrix(lngContador, 1)) <> "" Then
'
'                                 vgstrParametrosSP = Trim(grdPrecios.TextMatrix(lngContador, cintColCodigoPCE)) & "|" & Trim(grdPrecios.TextMatrix(lngContador, cintColDescripcion)) & "|" & CStr(Val(Format(grdPrecios.TextMatrix(lngContador, cintColPrecioLayout), "##########0.00####"))) & "|" & grdPrecios.TextMatrix(lngContador, cintColCantidad) & "|" & CStr(Val(Format(grdPrecios.TextMatrix(lngContador, cintColPrecioUnidosis), "##########0.00####")))
'
''                                 vlstrSentencia = "INSERT INTO PVLISTAPEMEXPRECIOS (CHRCODIGOPCE,VCHDESCRIPCION,  MNYPRECIOLAYOUT,INTCANTIDAD,                                    MNYPRECIOUNIDOSIS) VALUES (" & _
''                                                 Trim(grdPrecios.TextMatrix(lngContador, cintColCodigoPCE)) & ",'" & Trim(grdPrecios.TextMatrix(lngContador, cintColDescripcion)) & "'," & CStr(Val(Format(grdPrecios.TextMatrix(lngContador, cintColPrecioLayout), "##########0.00####"))) & "," & grdPrecios.TextMatrix(lngContador, cintColCantidad) & "," & _
''                                                 CStr(Val(Format(grdPrecios.TextMatrix(lngContador, cintColPrecioUnidosis), "##########0.00####"))) & ")"
''                                pEjecutaSentencia vlstrSentencia
'
'                                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSLISTAPRECIOPEMEX"
'
'
'                            End If
'
'                        End With
'                    Next lngContador
'
'                      Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngPersonaGraba, "LISTA DE PRECIOS PEMEX", CStr(vgintNumeroDepartamento))
'                    EntornoSIHO.ConeccionSIHO.CommitTrans
'
'                     '|La operación se realizó satisfactoriamente.
'                    MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
'
'
'            End If



    Dim i As Integer

    For i = 1 To grdPrecios.Rows - 1
            If grdPrecios.TextMatrix(i, 0) = "*" Then
                frmCargoPrecioPemex.vlstrChrCveArticulo = Trim(grdPrecios.TextMatrix(i, 1))
                frmCargoPrecioPemex.vlstrArtDescripcion = Trim(grdPrecios.TextMatrix(i, 3))
                frmCargoPrecioPemex.vlstrCodigo = Trim(grdPrecios.TextMatrix(i, 2))
                frmCargoPrecioPemex.vlintPrecio = FormatNumber(Trim(grdPrecios.TextMatrix(i, 4)), 2)
                'CStr(Val(Format(Trim(grdPrecios.TextMatrix(i, 4)), "##########0.00####")))
                frmCargoPrecioPemex.vlintCantidad = Val((grdPrecios.TextMatrix(i, 5)))
                frmCargoPrecioPemex.vlintPreciounidosis = FormatNumber(Trim(grdPrecios.TextMatrix(i, 6)), 2)
                'CStr(Val(Format(Trim(grdPrecios.TextMatrix(i, 6)), "##########0.00####")))

                Load frmCargoPrecioPemex

                frmCargoPrecioPemex.Show vbModal, Me


                'End If
'                txtcargo.Text = ""
'                txtcargo.SetFocus
'                pLimpiaSeleccionGrid
'                cmdGrabarRegistro.Enabled = False
'                cmdCancelaOrd.Enabled = False
                pLimpiaSeleccionGrid
                  cmdGrabarRegistro.Enabled = False
                  cmdCancelaOrd.Enabled = False
                  txtcargo.Text = ""
                  lblNombreCompleto = ""
                  pConfiguraGrid
                  txtcargo.SetFocus
                Exit For

            End If
        Next


End Sub

Private Sub cmdImportar_Click()
On Error GoTo NotificaError
    Dim objXLApp As Object
    Dim txtRuta As String
    Dim vlstrValidador As String
    Dim intLoopCounter As Integer
    Dim intGridCounter As Integer
    Dim dblAvance As Double
    Dim vlblnModificaVal As Boolean
    Dim vlstrTipoCargo As String
    Dim vlRespuesta As Integer
    Dim intValorCero As Integer
    Dim vlstrSentencia As String
    Dim rngData As Object
    Dim colIndex As Integer
    Dim contieneCaracterEspecialOLetra As Boolean

'    vlstrSentencia = "DELETE FROM PVLISTAPEMEXPRECIOS"
'    pEjecutaSentencia vlstrSentencia
'    pLimpiaSeleccionGrid
'    cmdGrabarRegistro.Enabled = False
'    cmdCancelaOrd.Enabled = False
    
   
    Set objXLApp = CreateObject("Excel.Application")

    CommonDialog1.DialogTitle = "Abrir archivo"
    CommonDialog1.Filter = "Documentos excel|*.xls;*.xlsx;"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowOpen
    If Err Then
        'Si se cancela el cuadro de diálogo
        Exit Sub
    End If
    
     pLimpiaSeleccionGrid
    cmdGrabarRegistro.Enabled = False
    cmdCancelaOrd.Enabled = False
    txtcargo.Text = ""
    lblNombreCompleto = ""
    pConfiguraGrid
    txtcargo.SetFocus
    
   
    
    txtRuta = CommonDialog1.FileName
    vlstrValidador = ""
    With objXLApp
        .Workbooks.Open txtRuta
        .Workbooks(1).Worksheets(1).Select
        
       
        
        dblAvance = 50 / CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
        '------------------------
        ' Configuración de la Barra de estado
        '------------------------
        lblTextoBarra.Caption = "Importando información, por favor espere..."
        freBarra.Visible = True
        freBarra.Top = 890
        pgbBarra.Value = 0
        freBarra.Refresh
        pgbBarra.Value = 0
    
     
    
    
        '---VALIDACIONES DE EXCEL ANTES DE REALIZAR LA IMPORTACIÓN---
        '--- VALIDACIÓN PARA SABER SI NO HAY DATOS EN EL ARCHIVO ---
        If CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) > 1 Then
            '---VALIDACIÓN PARA SABER SI TIENEN LA MISMA CANTIDAD DE FILAS----
            'If grdPrecios.Rows - 1 <> (CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 1) Then
            If grdPrecios.Rows - 1 < 1 And (CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 1) < 1 Then
                freBarra.Visible = False
                pgbBarra.Value = 100
                MsgBox "La cantidad de cargos que se intentan importar son diferentes a los cargos que ya se tienen en la cuadrícula", vbOKOnly + vbInformation, "Mensaje"
                
                .Workbooks(1).Close False
                .Quit
                objXLApp = Nothing
                Exit Sub
            End If
                        
            '---VALIDACIÓN PARA SABER SI TIENE FILAS REPETIDAS O QUE NO ESTEN EN LA CUADRICULA---
            For intLoopCounter = 2 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 2
                vlblnModificaVal = False

                
                '--- SI ENCUENTRA LA FILA EN LA CUADRÍCULA Y NO ESTA REPETIDA, SIGUE CON LAS DEMÁS VALIDACIONES ---
                '--- CLAVE ---
                If Trim(.range("A" & intLoopCounter)) = "" Then 'Clave
                    MsgBox "Renglón " & intLoopCounter & ". Clave de cargo no encontrada. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    'objXLApp = Nothing
                    Set objXLApp = Nothing
                    Exit Sub
                End If
                '--- DESCRIPCIÓN CARGO ---
                If Trim(.range("B" & intLoopCounter)) = "" Then 'Descripción
                    MsgBox "Renglón " & intLoopCounter & ". Descripción de cargo no encontrada. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    Set objXLApp = Nothing
                    Exit Sub
                End If
              
                '--- PRECIO ---
                If Not (IsNumeric(Trim(.range("C" & intLoopCounter)))) Then 'Precio
                   MsgBox "Renglón " & intLoopCounter & ". Formato de la columna Precio es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                   pLimpiaSeleccionGrid
                   freBarra.Visible = False
                   pgbBarra.Value = 100
                   .Workbooks(1).Close False
                   .Quit
                   Set objXLApp = Nothing
                   Exit Sub
                Else
                   ' --- PRECIO IGUAL A 0 ---
                   If CDbl(Trim(.range("C" & intLoopCounter))) = 0 Then
                        intValorCero = intValorCero + 1
                   Else
                        ' --- PRECIO MENOR A 0 ---
                        If CDbl(Trim(.range("C" & intLoopCounter))) < 0 Then
                            MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Precio es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                            pLimpiaSeleccionGrid
                            freBarra.Visible = False
                            pgbBarra.Value = 100
                            .Workbooks(1).Close False
                            .Quit
                            Set objXLApp = Nothing
                            Exit Sub
                        End If
                   End If
                End If
                
                '--- CANTIDAD ---
                If Not (IsNumeric(Trim(.range("D" & intLoopCounter)))) Then 'CANTIDAD
                   MsgBox "Renglón " & intLoopCounter - 1 & ". Columna identificador artículo: Dato incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                   .Workbooks(1).Close False
                   .Quit
                   Set objXLApp = Nothing
                   Exit Sub
                Else
                   ' --- CANTIDAD IGUAL A 0 ---
                   If CDbl(Trim(.range("D" & intLoopCounter))) = 0 Then
                        intValorCero = intValorCero + 1
                   Else
                        ' --- CANTIDAD MENOR A 0 ---
                        If CDbl(Trim(.range("D" & intLoopCounter))) < 0 Then
                            MsgBox "Renglón " & intLoopCounter - 1 & ". Columna cantidad: Dato Incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                            .Workbooks(1).Close False
                            .Quit
                            Set objXLApp = Nothing
                            Exit Sub
                        End If
                   End If
                End If
                '--- PRECIO UNIDOSIS ---
                If Not (IsNumeric(Trim(.range("E" & intLoopCounter)))) Then 'Precio unidosis
                   MsgBox "Renglón " & intLoopCounter & ". Formato de la columna Precio es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                   pLimpiaSeleccionGrid
                   freBarra.Visible = False
                   pgbBarra.Value = 100
                   .Workbooks(1).Close False
                   .Quit
                   Set objXLApp = Nothing
                   Exit Sub
                Else
                   ' --- PRECIO UNIDOSIS IGUAL A 0 ---
                   If CDbl(Trim(.range("E" & intLoopCounter))) = 0 Then
                        intValorCero = intValorCero + 1
                   Else
                        ' --- PRECIO UNIDOSISR A 0 ---
                        If CDbl(Trim(.range("E" & intLoopCounter))) < 0 Then
                            MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Precio unidosis es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                            pLimpiaSeleccionGrid
                            freBarra.Visible = False
                            pgbBarra.Value = 100
                            .Workbooks(1).Close False
                            .Quit
                            Set objXLApp = Nothing
                            Exit Sub
                        End If
                   End If
                End If
                
                ' Actualización de la barra de estado
                If pgbBarra.Value + dblAvance < 50 Then
                    pgbBarra.Value = pgbBarra.Value + dblAvance
                Else
                    pgbBarra.Value = 50
                End If
            Next intLoopCounter
            
            ' --- Existen filas con valor de precio cero. ¿Desea continuar con la importación? ---
            If intValorCero > 0 Then
                vlRespuesta = MsgBox("Existen filas con valor de precio cero. ¿Desea continuar con la importación? ", vbYesNo + vbQuestion, "Mensaje")
      
                If vlRespuesta = 7 Then
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    Set objXLApp = Nothing
                    Exit Sub
                End If
            End If
            
            Dim Cadena As String
            Dim i As Integer
            Dim specialChars As String
            Dim Maximo18Caracteres As Boolean
            specialChars = "!@#$%^&*()_+[]\{}|;':"",./<>?"
            
            contieneCaracterEspecialOLetra = False
            Maximo18Caracteres = False
            For intLoopCounter = 2 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 2
                Cadena = Trim(.range("A" & intLoopCounter))
                For i = 1 To Len(specialChars)
                    If InStr(Cadena, Mid(specialChars, i, 1)) > 0 Then
                        contieneCaracterEspecialOLetra = True
                        Exit For
                    End If
                Next i
                    
                If Cadena Like "*[A-Za-z]*" Then
                    contieneCaracterEspecialOLetra = True
                End If
                    
                If contieneCaracterEspecialOLetra Then
                    Exit For
                End If
          
                If Len(Cadena) > 18 Then
                    Maximo18Caracteres = True
                    Exit For
                End If
            
            Next intLoopCounter
            
            If contieneCaracterEspecialOLetra Or Maximo18Caracteres Then
               
                    MsgBox "Renglón " & intLoopCounter & ": No se puede realizar la importación. Se encontraron inconsistencias, verifique que el dato no sea mayor a 18 dígitos permitidos y no contenga caracteres especiales. ", vbOKOnly + vbInformation, "Mensaje"

               
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    
                    Set objXLApp = Nothing
                    Exit Sub
            End If
            
            '---TERMINAN VALIDACIONES, INICIA IMPORTACIÓN DE DATOS ---
            'se eliminan los registros existentes
            vlstrSentencia = "DELETE FROM PVLISTAPEMEXPRECIOS"
            pEjecutaSentencia vlstrSentencia
    
            For intLoopCounter = 2 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 2


                            'vgstrParametrosSP = Trim(.Range("A" & intLoopCounter)) & "|" & Trim(.Range("B" & intLoopCounter)) & "|" & Trim(.Range("C" & intLoopCounter).Text) & "|" & Trim(.Range("D" & intLoopCounter)) & "|" & Trim(.Range("E" & intLoopCounter).Text)
                             vgstrParametrosSP = Trim(.range("A" & intLoopCounter)) & "|" & Trim(.range("B" & intLoopCounter)) & "|" & Format(Replace(Trim(.range("C" & intLoopCounter).Text), ",", ""), "0.00") & "|" & Trim(.range("D" & intLoopCounter)) & "|" & Format(Replace(Trim(.range("E" & intLoopCounter).Text), ",", ""), "0.00")
                                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSLISTAPRECIOPEMEX"
                                                                                                                                'Replace(Text1, "_", "1")
                                
'                            grdPrecios.Rows = grdPrecios.Rows + 1
'                                grdPrecios.TextMatrix(intLoopCounter - 1, 1) = Trim(.Range("A" & intLoopCounter)) 'clave
'                                grdPrecios.TextMatrix(intLoopCounter - 1, 2) = Trim(.Range("B" & intLoopCounter)) 'Descripción
'                           grdPrecios.TextMatrix(intLoopCounter - 1, 3) = Format(Trim(.Range("C" & intLoopCounter).Text), "$###,###,###,##0.000##") 'precio
'                            grdPrecios.TextMatrix(intLoopCounter - 1, 4) = Trim(.Range("D" & intLoopCounter)) 'cantidad
'                            grdPrecios.TextMatrix(intLoopCounter - 1, 5) = Format(Trim(.Range("E" & intLoopCounter).Text), "$###,###,###,##0.00####") 'Precio unidosis
'

                                      
                ' Actualización de la barra de estado
                If pgbBarra.Value + dblAvance < 100 Then
                    pgbBarra.Value = pgbBarra.Value + dblAvance
                Else
                    pgbBarra.Value = 100
                End If
            Next intLoopCounter
                
            If vlblnRecalcularImportar Then
               ' Call cmdRec_Click
            End If
            
             pgbBarra.Value = 90
            'Realiza la carga de los datos importados al grid
            CargarDatosGrid
            
            freBarra.Visible = False
            pgbBarra.Value = 100
            
            MsgBox "Se realizó la importación exitosa de los datos en la cuadrícula.", vbOKOnly + vbInformation, "Mensaje"

            .Workbooks(1).Close False
            .Quit
        Else
            MsgBox "No existen datos para importar", vbOKOnly + vbInformation, "Mensaje"
            freBarra.Visible = False
            pgbBarra.Value = 100
            .Workbooks(1).Close False
            .Quit
        End If
        
    End With
      
    Set objXLApp = Nothing
    'cmdFiltrar.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImportar_Click"))
    Unload Me
    
End Sub

Private Function contieneCaracterEspecialOLetra(ByVal str As String) As Boolean
    Dim i As Integer
    Dim specialChars As String
    specialChars = "!@#$%^&*()_+[]\{}|;':"",./<>?"
    
    ' Verificar caracteres especiales
    For i = 1 To Len(specialChars)
        If InStr(str, Mid(specialChars, i, 1)) > 0 Then
            contieneCaracterEspecialOLetra = True
            Exit Function
        End If
    Next i
    
    ' Verificar letras
    If str Like "*[A-Za-z]*" Then
        contieneCaracterEspecialOLetra = True
        Exit Function
    End If
    
    contieneCaracterEspecialOLetra = False
End Function

Private Sub Form_Activate()
'Text1.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rsConceptoFacturacion As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    cmdImportar.Enabled = True
       
    Set rs1 = frsEjecuta_SP(str(vgintNumeroDepartamento), "SP_GNSELDEPARTAMENTOS")
    If rs1.RecordCount > 0 Then
        Text1.Text = rs1!VCHDESCRIPCION
    Else
        Text1.Text = ""
    End If
    rs1.Close
    
    pLimpiaSeleccionGrid
    cmdGrabarRegistro.Enabled = False
    cmdCancelaOrd.Enabled = False
        
    pConfiguraGrid
 
    
   
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub






Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = vbKeyEscape And Me.ActiveControl.Name <> "UpDown1" And ActiveControl.Name <> "lstPrecios" Then
        KeyAscii = 0
        Unload Me
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub pConfiguraGrid()
    On Error GoTo NotificaError
    
    
    'cmdExportar.Enabled = True
    'Se agrega para el botón importar
    cmdImportar.Enabled = True
    With grdPrecios
        .Clear
        .Cols = cintColumnas
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Consecutivo|Clave|Descripción|Precio|Cantidad|Precio Unidosis" '"|Clave|Descripción|Precio|Cantidad|Precio Unidosis||||"
        .RowHeight(0) = 580
        .RowHeight(1) = 280
        .ColWidth(0) = 200  'Fix

        .ColWidth(cintColIDConsecutivo) = 0
        .ColWidth(cintColCodigoPCE) = 2000
        .ColWidth(cintColDescripcion) = 5300
        .ColWidth(cintColPrecioLayout) = 1950
        .ColWidth(cintColCantidad) = 1950
        .ColWidth(cintColPrecioUnidosis) = 1950

       ' ColAlignmentFixed (cintColIDConsecutivo)
        .ColAlignmentFixed(cintColCodigoPCE) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDescripcion) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColPrecioLayout) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCantidad) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColPrecioUnidosis) = flexAlignCenterCenter
   
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(cintColCodigoPCE) = flexAlignLeftCenter
        .ColAlignment(cintColDescripcion) = flexAlignLeftCenter
        .ColAlignment(cintColPrecioLayout) = flexAlignRightCenter
        .ColAlignment(cintColCantidad) = flexAlignRightCenter
        .ColAlignment(cintColPrecioUnidosis) = flexAlignRightCenter

    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrid"))
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

'    Select Case vgstrEstadoManto
'        Case "A", "M"
'            Cancel = 1
'            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
'                pEnfocaTextBox txtClave
'            End If
'        Case "AE", "ME"
'            Cancel = 1
'            cmdCapturaPrecios_Click
'        Case "AEE", "MEE"
'            Cancel = 1
'            vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 2)
'        Case "B"
'            Cancel = 1
'            sstListas.Tab = 0
'            pEnfocaTextBox txtClave
'    End Select
    
'    If blnCatalogo = True Then
'        blnCatalogo = False
'        pNuevoRegistro
'        frmMantoListasPrecios.Height = cgIntAltoVentanaMin
'        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
'    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub



Private Sub grdPrecios_Click()
pMuestraNomComercialCompleto
End Sub
Sub pMuestraNomComercialCompleto()
    On Error GoTo NotificaError

    lblNombreCompleto.Caption = grdPrecios.TextMatrix(grdPrecios.Row, cintColDescripcion)
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraNomComercialCompleto"))
End Sub

Private Sub grdPrecios_DblClick()



On Error GoTo NotificaError
Dim existesleccionados As Boolean
Dim bndPermiteSeleccion As Boolean
Dim i As Integer


'***************unicaseleccion*****************
bndPermiteSeleccion = False


    With grdPrecios
        If .TextMatrix(1, 1) <> "" Then
         For i = 1 To grdPrecios.Rows - 1
            If grdPrecios.TextMatrix(i, 0) = "*" Then
                bndPermiteSeleccion = True
                Exit For
            Else
                bndPermiteSeleccion = False

            End If
          Next
            
        End If
    End With

'*****************************

    With grdPrecios
        If .TextMatrix(1, 1) <> "" Then
            If .TextMatrix(.Row, 0) = "*" Then
                .TextMatrix(.Row, 0) = ""
                .Col = 0
                .CellFontBold = True
    
            ElseIf .TextMatrix(.Row, 0) = "" And Trim(.TextMatrix(.Row, 1)) <> "" And bndPermiteSeleccion = False Then
                 .TextMatrix(.Row, 0) = "*"
                .Col = 0
                .CellFontBold = True
                
            End If
        End If
    End With
    
   If bndPermiteSeleccion = False Then
        cmdGrabarRegistro.Enabled = True
        cmdCancelaOrd.Enabled = True
   Else
        cmdGrabarRegistro.Enabled = False
        cmdCancelaOrd.Enabled = False

   End If
   

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPrecios_DblClick"))
End Sub

Private Sub grdPrecios_LostFocus()
    On Error GoTo NotificaError

    vgstrAcumTextoBusqueda = ""
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_LostFocus"))
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtcargo.SetFocus
    End If
End Sub

Private Sub tmrDespliega_Timer()

   ' cmdFiltrar_Click
  '  tmrDespliega.Enabled = False
    
End Sub




'Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intDelete As Integer, intPrint As Integer)
'    On Error GoTo NotificaError
'
'    cmdPrimerRegistro.Enabled = intTop = 1
'    cmdAnteriorRegistro.Enabled = intBack = 1
'    cmdBuscar.Enabled = intlocate = 1
'    cmdSiguienteRegistro.Enabled = intNext = 1
'    cmdUltimoRegistro.Enabled = intEnd = 1
'    cmdGrabarRegistro.Enabled = intSave = 1
'    cmdDelete.Enabled = intDelete = 1
'    cmdImprimir.Enabled = intPrint = 1
'    chkImprimeCargosPrecioCero.Enabled = intPrint = 1
'
'    Exit Sub
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
Private Sub pLimpiaSeleccionGrid()
On Error GoTo NotificaError

    Dim lngContador As Long
     grdPrecios.Rows = 2
    For lngContador = 1 To grdPrecios.Rows - 1
    grdPrecios.TextMatrix(1, lngContador - 1) = ""
        grdPrecios.TextMatrix(lngContador, 0) = ""
        
    Next lngContador

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & "pLimpiaSeleccionGrid"))
End Sub

'Private Sub pLimpiaGrid()
'On Error GoTo NotificaError
'    Dim intCols As Integer
'
'    grdArticulos.Rows = 2
'
'    For intCols = 1 To grdArticulos.Cols - 1
'        grdArticulos.TextMatrix(1, intCols - 1) = ""
'         grdArticulos.TextMatrix(1, intCols) = ""
'    Next
'    pFormatoGrid
'    intMaxManejos = 0
'
'
'Exit Sub
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGrid"))
'    Unload Me
'End Sub



Private Sub txtcargo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)

  KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
