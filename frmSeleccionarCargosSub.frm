VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSeleccionarCargosSub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar cargos"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtidProveedor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   13200
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   -480
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmSeleccionarCargosSub.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProveedorCargSub"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdCargosProveedores"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtProveedorSubrCarg"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmSeleccionarCargosSub.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "grdFacturaMultiemp"
      Tab(1).Control(3)=   "grdDetalleMultiemp"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(5)=   "txtProveedorMultiemp"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtProveedorMultiemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73920
         TabIndex        =   12
         Top             =   585
         Width           =   5535
      End
      Begin VB.Frame Frame2 
         Height          =   655
         Left            =   -70793
         TabIndex        =   8
         Top             =   7560
         Width           =   4590
         Begin VB.CommandButton cmdTerminarSelecMulti 
            Caption         =   "Terminar selección"
            Height          =   480
            Left            =   2280
            TabIndex        =   10
            Top             =   120
            Width           =   2235
         End
         Begin VB.CommandButton cmdInvertirSelcMulti 
            Caption         =   "Invertir selección"
            Height          =   480
            Left            =   40
            TabIndex        =   9
            Top             =   120
            Width           =   2235
         End
      End
      Begin VB.TextBox txtProveedorSubrCarg 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   585
         Width           =   5535
      End
      Begin VB.Frame Frame1 
         Height          =   655
         Left            =   4207
         TabIndex        =   2
         Top             =   7560
         Width           =   4590
         Begin VB.CommandButton cmdInvertirSeleccCargSub 
            Caption         =   "Invertir selección"
            Enabled         =   0   'False
            Height          =   480
            Left            =   40
            TabIndex        =   6
            Top             =   120
            Width           =   2235
         End
         Begin VB.CommandButton cmdTerminarSeleccCarg 
            Caption         =   "Terminar selección"
            Enabled         =   0   'False
            Height          =   480
            Left            =   2280
            TabIndex        =   5
            Top             =   120
            Width           =   2235
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCargosProveedores 
         Height          =   6495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Cargos de servicios subrogados"
         Top             =   960
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   11456
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDetalleMultiemp 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   11
         ToolTipText     =   "Cargos de servicios subrogados"
         Top             =   4590
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   5106
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFacturaMultiemp 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   13
         ToolTipText     =   "Facturas de servicios subrogados"
         Top             =   960
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   5530
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   -74850
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Detalle de la factura"
         Height          =   255
         Left            =   -74880
         TabIndex        =   7
         Top             =   4215
         Width           =   1695
      End
      Begin VB.Label lblProveedorCargSub 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSeleccionarCargosSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim vgstrParametrosSP As String
Public vlintPestañaInicial As Integer
Dim vlincveEmpresaCredito As Integer
Dim vlintcontador As Integer

Private Sub cmdInvertirSelcMulti_Click()
    With grdFacturaMultiemp
        For vlintcontador = 1 To .Rows - 1
            .Row = vlintcontador
            .Col = 0
            .Text = IIf(.Text = "*", "", "*")
            .CellFontBold = True
        Next
    End With
End Sub

Private Sub cmdInvertirSeleccCargSub_Click()
    With grdCargosProveedores
        For vlintcontador = 1 To .Rows - 1
            .Row = vlintcontador
            .Col = 0
            .Text = IIf(.Text = "*", "", "*")
            .CellFontBold = True
        Next
    End With
End Sub

Private Sub cmdTerminarSeleccCarg_Click()
' vlblnError = False
'    pValidaSeleccion
'    If Not vlblnError Then
        Me.Hide
'    End If
End Sub

Private Sub cmdTerminarSelecMulti_Click()
    PagregarProveedoresCargos
    Me.Hide
End Sub

Private Sub PagregarProveedoresCargos()
On Error GoTo NotificaError
Dim rs As ADODB.Recordset
Dim vlintCont As Integer
    vgstrParametrosSP = ""
    With grdCargosProveedores
        
        For vlintCont = 1 To grdFacturaMultiemp.Rows - 1
            If grdFacturaMultiemp.TextMatrix(vlintCont, 0) = "*" Then
                If grdFacturaMultiemp.RowData(vlintCont) <> -1 Then
                    vgstrParametrosSP = grdFacturaMultiemp.TextMatrix(vlintCont, 2)
                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_cpselcargmultiempsub")
                    If rs.RecordCount > 0 Then
                        Do While Not rs.EOF
                                If grdCargosProveedores.RowData(1) <> -1 Then
                                    .Rows = .Rows + 1
                                    .Row = .Rows - 1
                                End If
                                .RowData(.Row) = rs!numcargo
                                .TextMatrix(.Row, 0) = "*"
                                .TextMatrix(.Row, 2) = rs!Nombre 'Nombre del Paciente
                                .TextMatrix(.Row, 3) = Format(rs!dtmFechahora, "dd/mmm/yyyy") 'Descripción del servicio
                                .TextMatrix(.Row, 4) = Format(rs!mnyPrecio, "$ ###,###,###,###,###.00") 'Nombre de la empresa
                                .TextMatrix(.Row, 5) = rs!Descripcion 'Tipo de convenio
                                .TextMatrix(.Row, 6) = Format(rs!MNYCantidad, "$ ###,###,###,###,###.00") 'Importe de la cantidad fija o del porcentaje
                                .TextMatrix(.Row, 7) = Format(rs!MNYIVA, "$ ###,###,###,###,###.00")
                                .TextMatrix(.Row, 8) = rs!intNumeroCuenta
                                .TextMatrix(.Row, 9) = 0
                                .TextMatrix(.Row, 10) = rs!IVA
                                rs.MoveNext
            '                    cmdInvertirSeleccCargSub.Enabled = True
            '                    cmdTerminarSeleccCarg.Enabled = True
                        Loop
                    'rs.Close
                    End If
                End If
            End If
        Next
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":PagregarProveedoresCargos"))
    Unload Me
End Sub

Private Sub Form_Activate()
Dim rs As ADODB.Recordset
On Error GoTo NotificaError
    If Me.SSTab1.Tab = 0 Then
        pLlenargridCargosProveedores
    Else
        Set rs = frsRegresaRs("select intidempresacliente from siempresacliente sim where sim.intcveproveedor = " & txtidProveedor.Text & " and sim.TNYIDEMPRESA = " & vgintClaveEmpresaContable)
        If rs.RecordCount <> 0 Then vlincveEmpresaCredito = rs!intidempresacliente
        rs.Close
        pLlenargridMultiemp
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
Dim rs As ADODB.Recordset

    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    If vlintPestañaInicial = 0 Then
        SSTab1.Tab = 0
        pConfiguraGridCargosProveedores
    Else
        SSTab1.Tab = 1
        pConfiguragrdDetalleMultiemp
        pConfiguragrdFacturaMultiemp
        pConfiguraGridCargosProveedores
        'vlincveEmpresaCredito = "select intidempresacliente from siempresacliente sim where sim.intcveproveedor = " & txtidProveedor.Text & " and sim.TNYIDEMPRESA = " & vgintClaveEmpresaContable
    End If
    'pLlenargridCargosProveedores
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub
Private Sub pLlenargridMultiemp()
On Error GoTo NotificaError
Dim rs As ADODB.Recordset
    
    With grdFacturaMultiemp
        If .RowData(1) = -1 Then
            'aqui tienes id de proveedor que esta relacionado a una empresa en sistemas cambiar por el id de la empresa
            vgstrParametrosSP = vlincveEmpresaCredito & "|" & -1
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_cpselfacMultiempsub")
            If rs.RecordCount > 0 Then
                Do While Not rs.EOF
                    If grdFacturaMultiemp.RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    .RowData(.Row) = rs!intConsecutivo
                    .TextMatrix(.Row, 1) = Format(rs!fecha, "dd/mmm/yyyy")  'Nombre del Paciente
                    .TextMatrix(.Row, 2) = rs!chrfoliocredito 'Descripción del servicio
                    .TextMatrix(.Row, 3) = Format(rs!mnytotalfactura, "$ ###,###,###,###,###.00") 'Nombre de la empresa
                    .TextMatrix(.Row, 4) = Format(rs!MNYDESCUENTO, "$ ###,###,###,###,###.00") 'Tipo de convenio
                    .TextMatrix(.Row, 5) = Format(rs!mnytotalfactura - rs!smyiva, "$ ###,###,###,###,###.00")
                    .TextMatrix(.Row, 6) = Format(rs!smyiva, "$ ###,###,###,###,###.00")
                    .TextMatrix(.Row, 7) = Format(rs!mnytotalfactura, "$ ###,###,###,###,###.00")
'                    .TextMatrix(.Row, 8) = rs!inttipoacuerdo
'                    .TextMatrix(.Row, 9) = rs!IVA
                    rs.MoveNext
                    cmdInvertirSelcMulti.Enabled = True
                    cmdTerminarSelecMulti.Enabled = True
                Loop
            .Row = 1
            grdFacturaMultiemp_Click
            rs.Close
            End If
        End If
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenargridMultiemp"))
    Unload Me
End Sub


Private Sub pConfiguragrdFacturaMultiemp()
On Error GoTo NotificaError

    With grdFacturaMultiemp
        .Cols = 8
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Fecha|Folio|Importe|Descuentos|Subtotal|IVA|Total"
        .ColWidth(0) = 200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1800
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
        .ColWidth(5) = 1800
        .ColWidth(6) = 1800
        .ColWidth(7) = 1800
'        .ColWidth(8) = 3500
'        .ColWidth(9) = 3500
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        grdFacturaMultiemp.RowData(1) = -1
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguragrdFacturaMultiemp"))
    Unload Me
End Sub


Private Sub pConfiguragrdDetalleMultiemp()
On Error GoTo NotificaError

    With grdDetalleMultiemp
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Cuenta|Paciente|Fecha del cargo|Cantidad|Descripción|Cantidad a pagar"
        .ColWidth(0) = 200
        .ColWidth(1) = 1600
        .ColWidth(2) = 3500
        .ColWidth(3) = 1300
        .ColWidth(4) = 1600
        .ColWidth(5) = 3500
        .ColWidth(6) = 1600
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        grdDetalleMultiemp.RowData(1) = -1
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguragrdDetalleMultiemp"))
    Unload Me
End Sub


Private Sub pLlenargridCargosProveedores()
On Error GoTo NotificaError
Dim rs As ADODB.Recordset
    
    With grdCargosProveedores
        If .RowData(1) = -1 Then
'            vgstrParametrosSP = txtidProveedor.Text & "|" & -1 'vgintClaveEmpresaContable
            vgstrParametrosSP = txtidProveedor.Text & "|" & vgintClaveEmpresaContable
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CPSELCARGOSSERVICIOSUB")
            'aqui por la empresa no esta mostrando los cargos revisar .. cambiar por -1 para poder ver cargos lalal pero que pasa si es multiemp ? hmm
            If rs.RecordCount > 0 Then
                Do While Not rs.EOF
                    If grdCargosProveedores.RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    .RowData(.Row) = rs!IntNumCargo
                    .TextMatrix(.Row, 1) = rs!cuentapac 'Nombre del Paciente
                    .TextMatrix(.Row, 2) = rs!Nombre 'Nombre del Paciente
                    .TextMatrix(.Row, 3) = Format(rs!fecha, "dd/mmm/yyyy") 'Descripción del servicio
                    .TextMatrix(.Row, 4) = Format(rs!cantidadcargo, "$ ###,###,###,###,###.00") 'Nombre de la empresa
                    .TextMatrix(.Row, 5) = rs!Descripcion 'Tipo de convenio
                    .TextMatrix(.Row, 6) = IIf(rs!inttipoacuerdo = 1, rs!MNYCantidad & "%", Format(rs!MNYCantidad, "$ ###,###,###,###,###.00")) 'Importe de la cantidad fija o del porcentaje
                    .TextMatrix(.Row, 7) = Format(rs!MNYIVA, "$ ###,###,###,###,###.00")
                    .TextMatrix(.Row, 8) = rs!intNumeroCuenta
                    .TextMatrix(.Row, 9) = rs!inttipoacuerdo
                    .TextMatrix(.Row, 10) = rs!IVA
                    .TextMatrix(.Row, 11) = rs!conceptofac
                    rs.MoveNext
                    cmdInvertirSeleccCargSub.Enabled = True
                    cmdTerminarSeleccCarg.Enabled = True
                Loop
            rs.Close
            End If
        End If
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenargridCargosProveedores"))
    Unload Me
End Sub

Private Sub pConfiguraGridCargosProveedores()
On Error GoTo NotificaError

    With grdCargosProveedores
        .Cols = 12
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Cuenta|Paciente|Fecha del cargo|Cantidad|Descripción|Cantidad a pagar"
        .ColWidth(0) = 200
        .ColWidth(1) = 800
        .ColWidth(2) = 4000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1600
        .ColWidth(5) = 3500
        .ColWidth(6) = 1600
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        grdCargosProveedores.RowData(1) = -1
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridProveedores"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub grdFacturaMultiemp_Click()
    pLimpiargridcargos
    pConfiguragrdDetalleMultiemp
    pCargarCargosMulti
End Sub

Private Sub pLimpiargridcargos()
On Error GoTo NotificaError
Dim intContador As Integer

    With grdDetalleMultiemp
    .Clear
    .RowData(1) = -1
    .Rows = 2
    .Row = 1
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiargridcargos"))
    Unload Me
End Sub

Private Sub pCargarCargosMulti()
On Error GoTo NotificaError
Dim rs As ADODB.Recordset
    
    With grdDetalleMultiemp
        If grdFacturaMultiemp.RowData(1) <> -1 Then
            vgstrParametrosSP = grdFacturaMultiemp.TextMatrix(grdFacturaMultiemp.Row, 2)
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CPSELCARGMULTIEMPSUB")
            If rs.RecordCount > 0 Then
                Do While Not rs.EOF
                    If grdDetalleMultiemp.RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    .RowData(.Row) = rs!numcargo
                    .TextMatrix(.Row, 1) = rs!intNumCuenta 'Nombre del Paciente
                    .TextMatrix(.Row, 2) = rs!Nombre 'Descripción del servicio
                    .TextMatrix(.Row, 3) = Format(rs!dtmFechahora, "dd/mmm/yyyy") 'Nombre de la empresa
                    .TextMatrix(.Row, 4) = Format(rs!mnyPrecio, "$ ###,###,###,###,###.00") 'Tipo de convenio
                    .TextMatrix(.Row, 5) = rs!Descripcion
                    .TextMatrix(.Row, 6) = Format(rs!Cantidad, "$ ###,###,###,###,###.00")
'                    .TextMatrix(.Row, 7) = rs!mnytotalfactura
'                    .TextMatrix(.Row, 8) = rs!inttipoacuerdo
'                    .TextMatrix(.Row, 9) = rs!IVA
                    rs.MoveNext
'                    cmdInvertirSelcMulti.Enabled = True
'                    cmdTerminarSelecMulti.Enabled = True
                Loop
            rs.Close
            End If
        End If
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargarCargosMulti"))
    Unload Me
End Sub

Private Sub grdCargosProveedores_DblClick()
On Error GoTo NotificaError
    With grdCargosProveedores
        If .TextMatrix(.Row, 0) = "*" Then
            .TextMatrix(.Row, 0) = ""
            .Col = 0
            .CellFontBold = True
        Else
            .TextMatrix(.Row, 0) = "*"
            .Col = 0
            .CellFontBold = True
        End If
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargosProveedores_DblClick"))
End Sub

Private Sub grdFacturaMultiemp_DblClick()
   With grdFacturaMultiemp
        If .TextMatrix(.Row, 0) = "*" Then
            .TextMatrix(.Row, 0) = ""
            .Col = 0
            .CellFontBold = True
        Else
            .TextMatrix(.Row, 0) = "*"
            .Col = 0
            .CellFontBold = True
        End If
    End With
End Sub
