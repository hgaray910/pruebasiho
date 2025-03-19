VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantoFormasPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de pago"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabFormas 
      Height          =   7860
      Left            =   -45
      TabIndex        =   30
      Top             =   -480
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   13864
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoFormasPago.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoFormasPago.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmMantoFormasPago.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   6195
         Left            =   -74880
         TabIndex        =   46
         Top             =   1320
         Width           =   8230
         Begin VB.ComboBox cboTipoCargoBancario 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   45
            ToolTipText     =   "Seleccione el tipo de cargo bancario"
            Top             =   250
            Width           =   5130
         End
         Begin VB.CommandButton cmdIncluirComision 
            Height          =   495
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoFormasPago.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Incluir la comisión capturada para la forma de pago"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.ComboBox cboComisionBancaria 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   47
            ToolTipText     =   "Seleccione la comisión bancaria"
            Top             =   705
            Width           =   5130
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdComisiones 
            Height          =   4695
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Comisiones bancarias seleccionadas para la forma de pago"
            Top             =   1320
            Width           =   7965
            _ExtentX        =   14049
            _ExtentY        =   8281
            _Version        =   393216
            Cols            =   6
            GridColor       =   -2147483633
            ScrollBars      =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de cargo bancario"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   285
            Width           =   1650
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Comisión bancaria"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   735
            Width           =   1290
         End
      End
      Begin VB.Frame Frame4 
         Height          =   720
         Left            =   -74880
         TabIndex        =   44
         Top             =   600
         Width           =   8230
         Begin VB.TextBox txtFormaPago 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   53
            ToolTipText     =   "Descripción de la forma de pago"
            Top             =   240
            Width           =   5130
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Forma de pago"
            Height          =   195
            Left            =   140
            TabIndex        =   52
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         Height          =   7035
         Left            =   -74880
         TabIndex        =   37
         Top             =   480
         Width           =   8230
         Begin VB.ComboBox cboDeptoBusqueda 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Seleccione el departamento"
            Top             =   195
            Width           =   5130
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFormas 
            Height          =   6300
            Left            =   120
            TabIndex        =   29
            Top             =   555
            Width           =   7965
            _ExtentX        =   14049
            _ExtentY        =   11113
            _Version        =   393216
            GridColor       =   12632256
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   225
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   1732
         TabIndex        =   36
         Top             =   6880
         Width           =   4920
         Begin VB.CommandButton cmdComisiones 
            Caption         =   "Comisiones bancarias"
            Height          =   495
            Left            =   3600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoFormasPago.frx":0546
            TabIndex        =   27
            ToolTipText     =   "Comisiones bancarias"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3090
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoFormasPago.frx":06E8
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Borrar"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2595
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoFormasPago.frx":088A
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Grabar"
            Top             =   165
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2100
            Picture         =   "frmMantoFormasPago.frx":0BCC
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Ultimo registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1590
            Picture         =   "frmMantoFormasPago.frx":10BE
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Siguiente registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1080
            Picture         =   "frmMantoFormasPago.frx":1230
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Búsqueda"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   585
            Picture         =   "frmMantoFormasPago.frx":13A2
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Anterior registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   90
            Picture         =   "frmMantoFormasPago.frx":1514
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Primer registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6270
         Left            =   150
         TabIndex        =   31
         Top             =   495
         Width           =   8175
         Begin VB.CheckBox chkPinPad 
            Caption         =   "Habilitar interfaz con pinpad"
            Height          =   255
            Left            =   1650
            TabIndex        =   12
            ToolTipText     =   "Habilitar conexion con pinpad"
            Top             =   3000
            Width           =   2415
         End
         Begin VB.Frame fraPinPad 
            Height          =   1455
            Left            =   1650
            TabIndex        =   13
            Top             =   3360
            Width           =   6255
            Begin VB.ComboBox cboTerminal 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   14
               ToolTipText     =   "Terminal asociada a la forma de pago"
               Top             =   360
               Width           =   4695
            End
            Begin VB.TextBox txtImpVoucher 
               Height          =   315
               Left            =   1440
               MaxLength       =   255
               TabIndex        =   15
               ToolTipText     =   "Impresora para los comprobantes de los pagos con tarjeta"
               Top             =   840
               Width           =   4695
            End
            Begin VB.Label Label12 
               Caption         =   "Impresora del comprobante"
               Enabled         =   0   'False
               Height          =   435
               Left            =   120
               TabIndex        =   58
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label11 
               Caption         =   "Terminal"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   390
               Width           =   1335
            End
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Departamento"
            Top             =   5400
            Width           =   6300
         End
         Begin VB.ComboBox cboMetodosSATCFDi 
            Height          =   315
            Left            =   3850
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Selección del método de pago del SAT para CFDi"
            Top             =   1920
            Width           =   4100
         End
         Begin VB.ComboBox cboMetodosSAT 
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "frmMantoFormasPago.frx":1916
            Left            =   3850
            List            =   "frmMantoFormasPago.frx":1918
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Selección del método de pago del SAT para contabilidad electrónica"
            Top             =   1280
            Width           =   4100
         End
         Begin VB.Frame fraTipo 
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   1650
            TabIndex        =   40
            Top             =   945
            Width           =   2100
            Begin VB.OptionButton optTipo 
               Caption         =   "Cheque"
               Height          =   270
               Index           =   4
               Left            =   0
               TabIndex        =   6
               Top             =   1020
               Width           =   2340
            End
            Begin VB.OptionButton optTipo 
               Caption         =   "Transferencia bancaria"
               Height          =   270
               Index           =   3
               Left            =   0
               TabIndex        =   5
               Top             =   795
               Width           =   2340
            End
            Begin VB.OptionButton optTipo 
               Caption         =   "Tarjeta"
               Height          =   270
               Index           =   2
               Left            =   0
               TabIndex        =   4
               Top             =   555
               Width           =   1380
            End
            Begin VB.OptionButton optTipo 
               Caption         =   "Crédito"
               Height          =   270
               Index           =   1
               Left            =   0
               TabIndex        =   3
               Top             =   330
               Width           =   1380
            End
            Begin VB.OptionButton optTipo 
               Caption         =   "Efectivo"
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   2
               Top             =   90
               Width           =   1380
            End
         End
         Begin VB.CheckBox chkActiva 
            Caption         =   "Activa"
            Height          =   315
            Left            =   1650
            TabIndex        =   19
            ToolTipText     =   "Estado"
            Top             =   5760
            Width           =   900
         End
         Begin VB.CheckBox chkFolioReferencia 
            Caption         =   "Folio de referencia"
            Height          =   200
            Left            =   1650
            TabIndex        =   9
            ToolTipText     =   "Forma de pago que hace referencia a un folio"
            Top             =   2295
            Width           =   2280
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Dólares"
            Height          =   200
            Index           =   1
            Left            =   2475
            TabIndex        =   11
            ToolTipText     =   "Moneda en dólares"
            Top             =   2580
            Width           =   840
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Pesos"
            Height          =   200
            Index           =   0
            Left            =   1650
            TabIndex        =   10
            ToolTipText     =   "Moneda en pesos"
            Top             =   2580
            Width           =   840
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1650
            MaxLength       =   100
            TabIndex        =   1
            ToolTipText     =   "Descripción de la forma de pago"
            Top             =   630
            Width           =   6300
         End
         Begin VB.TextBox txtClave 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1650
            MaxLength       =   5
            TabIndex        =   0
            ToolTipText     =   "Clave de la forma de pago"
            Top             =   285
            Width           =   1005
         End
         Begin MSMask.MaskEdBox mskCuenta 
            Height          =   315
            Left            =   1650
            TabIndex        =   16
            ToolTipText     =   "Cuenta contable"
            Top             =   5040
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            Caption         =   "Terminal"
            Height          =   255
            Left            =   225
            TabIndex        =   56
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   225
            TabIndex        =   55
            Top             =   5460
            Width           =   1005
         End
         Begin VB.Label lblMetodoPagoSAT 
            AutoSize        =   -1  'True
            Caption         =   "Método de pago del SAT para contabilidad electrónica"
            Height          =   195
            Left            =   3855
            TabIndex        =   54
            Top             =   1035
            Width           =   3870
         End
         Begin VB.Label lblCuentaBanco 
            Caption         =   "Lo registrado con esta forma de pago afectará el libro de bancos al cerrar el corte."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   43
            Top             =   5760
            Visible         =   0   'False
            Width           =   4995
         End
         Begin VB.Label lblMetodoPagoCFDi 
            AutoSize        =   -1  'True
            Caption         =   "Método de pago del SAT para CFDi"
            Height          =   195
            Left            =   3855
            TabIndex        =   42
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   225
            TabIndex        =   41
            Top             =   1050
            Width           =   315
         End
         Begin VB.Label lblDescripcionCuenta 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3480
            TabIndex        =   17
            Top             =   5040
            Width           =   4455
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   225
            TabIndex        =   38
            Top             =   5820
            Width           =   495
         End
         Begin VB.Label lblMoneda 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   225
            TabIndex        =   35
            Top             =   2550
            Width           =   585
         End
         Begin VB.Label lblCuentaContable 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta contable"
            Height          =   195
            Left            =   225
            TabIndex        =   34
            Top             =   5100
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   690
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   345
            Width           =   405
         End
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Terminal"
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmMantoFormasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
' Programa para dar mantenimiento a las formas de pago (PvFormaPago)
' Fecha de programación: Martes 27 de Febrero de 2001
'-----------------------------------------------------------------------------------
Option Explicit
Const cintColIdForma = 1
Const cintColDescripcion = 2
Const cIntColEstado = 3
Const cintColDepartamento = 4
Const cintCols = 5
Const cstrColumnas = "|Clave|Descripción|Estado|Departamento"

Dim rsPvFormaPago As New ADODB.Recordset
Dim rsDepartamento As New ADODB.Recordset

Dim vlstrx As String
Dim llngNumCuenta As Long   'Id. de la cuenta seleccionada

Dim vlblnConsulta As Boolean
Dim lblnChange As Boolean   'Para ejecutar o no lo que está en el evento mskCuenta_Change
Dim lblnCarga As Boolean    'Para ejecutar o no la cargada de las formas de pago
Dim lstrTipo As String      'Tipo de forma de pago (E = efectivo, C = crédito, T = tarjeta, B = transferencia bancaria)
Dim lintMoneda As Integer   'Moneda del banco si se asoció la forma de pago a una cuenta contable de un banco
Private Sub pCargaTerminales()
    Dim rsTerm As ADODB.Recordset
    Set rsTerm = frsRegresaRs("select * from PVTerminal order by vchNombre", adLockReadOnly, adOpenForwardOnly)
    Do While Not rsTerm.EOF
        cboTerminal.AddItem rsTerm!vchNombre
        cboTerminal.ItemData(cboTerminal.newIndex) = rsTerm!intCveTerminal
        rsTerm.MoveNext
    Loop
    rsTerm.Close
End Sub


Private Sub pConfiguraGrid()
    On Error GoTo NotificaError
    
    With grdComisiones
        .Cols = 9
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "||||Tipo de cargo bancario|Comisión bancaria|Porcentaje|IVA|Predeterminado"
        .ColWidth(0) = 100 'Fix
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 2500 'Tipo de cargo
        .ColWidth(5) = 1800 'Comisión
        .ColWidth(6) = 1000  'Porcentaje
        .ColWidth(7) = 1000  'IVA
        .ColWidth(8) = 1200 'Predeterminado
        .ColAlignment(4) = flexAlignLeftBottom
        .ColAlignment(5) = flexAlignLeftBottom
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterTop
        .ColAlignmentFixed(4) = flexAlignLeftBottom
        .ColAlignmentFixed(5) = flexAlignLeftBottom
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(1, 4) = ""
        .RowData(1) = -1
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
    Unload Me
End Sub

Private Sub cboAgrupadoras_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub cboAgrupadoras_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cboDepartamento_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub cboDeptoBusqueda_Click()
    If cboDeptoBusqueda.ListIndex <> -1 Then
        If lblnCarga Then
            pCarga
        End If
    End If
End Sub

Private Sub pCargaMetodosSAT()
    On Error GoTo NotificaError

    pLlenarCboSentencia cboMetodosSAT, "SELECT INTIDREGISTRO, VCHDESCRIPCION FROM PVFORMAPAGOSAT ORDER BY VCHDESCRIPCION", 1, 0
    cboMetodosSAT.AddItem "<NINGUNA>", 0
    cboMetodosSAT.ItemData(cboMetodosSAT.newIndex) = 0
    
    cboMetodosSAT.ListIndex = 0
    
    pLlenarCboSentencia cboMetodosSATCFDi, "SELECT INTIDREGISTRO, VCHDESCRIPCION FROM PVMETODOPAGOSATCFDI ORDER BY VCHDESCRIPCION", 1, 0
    cboMetodosSATCFDi.AddItem "<NINGUNA>", 0
    cboMetodosSATCFDi.ItemData(cboMetodosSATCFDi.newIndex) = 0
    
    cboMetodosSATCFDi.ListIndex = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaMetodosSAT"))
End Sub

Private Sub pCarga()
    Dim rs As New ADODB.Recordset
    
    vgstrParametrosSP = "-1|-1|-1|" & CStr(cboDeptoBusqueda.ItemData(cboDeptoBusqueda.ListIndex)) & "|-1|*"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormaPago")
    
    With grdFormas
        .Clear
        .Cols = cintCols
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        
        .FormatString = cstrColumnas
        .ColWidth(cintColIdForma) = 1000
        .ColWidth(cintColDescripcion) = IIf(cboDeptoBusqueda.ItemData(cboDeptoBusqueda.ListIndex) = -1, 3000, 4500)
        .ColWidth(cIntColEstado) = 1000
        .ColWidth(cintColDepartamento) = IIf(cboDeptoBusqueda.ItemData(cboDeptoBusqueda.ListIndex) = -1, 2500, 0)
        
        .ColAlignment(cintColIdForma) = flexAlignRightCenter
        .ColAlignment(cintColDescripcion) = flexAlignLeftCenter
        .ColAlignment(cIntColEstado) = flexAlignLeftCenter
        .ColAlignment(cintColDepartamento) = flexAlignLeftCenter
        
        .ColAlignmentFixed(cintColIdForma) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDescripcion) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColEstado) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDepartamento) = flexAlignCenterCenter
        
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                .TextMatrix(.Rows - 1, cintColIdForma) = rs!intFormaPago
                .TextMatrix(.Rows - 1, cintColDescripcion) = Trim(rs!chrDescripcion)
                .TextMatrix(.Rows - 1, cIntColEstado) = IIf(rs!bitestatusactivo = 1, "ACTIVA", "INACTIVA")
                .TextMatrix(.Rows - 1, cintColDepartamento) = rs!NombreDepto
                rs.MoveNext
                .Rows = .Rows + 1
            Loop
            .Rows = .Rows - 1
        End If
    End With
End Sub


Private Sub pLlenaGrid()
    Dim vlstrSentencia As String
    Dim rsformapagotipocargocomision As New ADODB.Recordset
    Dim vlintTipoCargoSel As Integer
    Dim vlintcontador As Integer
    
    If cboTipoCargoBancario.ListIndex = -1 Then
        vlintTipoCargoSel = -1
    Else
        If cboTipoCargoBancario.ItemData(cboTipoCargoBancario.ListIndex) = 0 Then
            vlintTipoCargoSel = -1
        Else
            Exit Sub
        End If
    End If
    
    pLimpiaGrid
    pConfiguraGrid
    vlstrSentencia = "select pvformapagotipocargocomision.intformapago," & _
                            " pvformapagotipocargocomision.intcvetipocargo," & _
                            " pvformapagotipocargocomision.smicvecomision," & _
                            " pvformapagotipocargocomision.bitpredeterminado," & _
                            " pvtipocargobancario.chrdescripcion descTipoCargoBancario," & _
                            " pvComisionBancaria.chrdescripcion descComisionBancaria," & _
                            " pvComisionBancaria.mnycomision," & _
                            " PvComisionBancaria.smyiva" & _
                       " from pvformapagotipocargocomision" & _
                            ", pvtipocargobancario" & _
                            ", pvComisionBancaria" & _
                      " where intFormaPago = " & txtClave.Text & _
                        " and (pvformapagotipocargocomision.intcvetipocargo = " & vlintTipoCargoSel & " or " & vlintTipoCargoSel & " = -1)" & _
                        " and pvformapagotipocargocomision.intcvetipocargo = pvTipoCargoBancario.intcvetipocargo" & _
                        " and pvformapagotipocargocomision.smicvecomision = pvComisionBancaria.smicvecomision" & _
                        " and pvtipocargobancario.bitactivo = 1" & _
                        " and pvComisionBancaria.bitactivo = 1"
                        
    Set rsformapagotipocargocomision = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    If rsformapagotipocargocomision.RecordCount > 0 Then
        With grdComisiones
            vlintcontador = 1
            Do While Not rsformapagotipocargocomision.EOF
                .TextMatrix(vlintcontador, 1) = rsformapagotipocargocomision!intFormaPago
                .TextMatrix(vlintcontador, 2) = rsformapagotipocargocomision!intcvetipocargo
                .TextMatrix(vlintcontador, 3) = rsformapagotipocargocomision!smiCveComision
                .TextMatrix(vlintcontador, 4) = rsformapagotipocargocomision!descTipoCargoBancario
                .TextMatrix(vlintcontador, 5) = rsformapagotipocargocomision!descComisionBancaria
                .TextMatrix(vlintcontador, 6) = FormatPercent(rsformapagotipocargocomision!mnycomision / 100, 2)
                .TextMatrix(vlintcontador, 7) = FormatPercent(rsformapagotipocargocomision!smyIVA / 100, 2)
                .TextMatrix(vlintcontador, 8) = IIf(rsformapagotipocargocomision!bitpredeterminado = 1, "*", "")
                .Row = vlintcontador
                .Col = 8
                .CellFontBold = True
                .CellFontSize = 10
        
                rsformapagotipocargocomision.MoveNext
                vlintcontador = vlintcontador + 1
                .Rows = .Rows + 1
            Loop
            .Rows = .Rows - 1
        End With
    End If
    rsformapagotipocargocomision.Close
End Sub

Private Sub cboMetodosSAT_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub cboMetodosSATCFDi_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub cboTipoCargoBancario_Click()
    cboComisionBancaria.ListIndex = -1
    pLlenaGrid
End Sub

Private Sub chkActiva_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub chkFolioReferencia_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub chkPinPad_Click()
    
    txtImpVoucher.Text = ""

    cboTerminal.ListIndex = -1
    fraPinPad.Enabled = (chkPinPad.Value = vbChecked)
  If chkPinPad.Value = 1 Then
 
  End If
    txtImpVoucher.Enabled = (chkPinPad.Value = vbChecked)
   Label11.Enabled = (chkPinPad.Value = vbChecked)
   Label12.Enabled = (chkPinPad.Value = vbChecked)
    cboTerminal.Enabled = (chkPinPad.Value = vbChecked)
    
End Sub

Private Sub chkPinPad_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub cmdBack_Click()
On Error GoTo NotificaError
    
    If grdFormas.Row > 1 Then
        grdFormas.Row = grdFormas.Row - 1
    End If
    pMuestraForma grdFormas.TextMatrix(grdFormas.Row, cintColIdForma)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdComisiones_Click()
    
    If grdComisiones.Rows <= 2 And grdComisiones.TextMatrix(1, 4) = "" Then cboTipoCargoBancario.Text = "<TODOS>"
    
    txtFormaPago.Text = txtDescripcion.Text
    sstabFormas.Tab = 2
    cmdSave.Enabled = True
    cboTipoCargoBancario.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Error GoTo NotificaError
    
    Dim vllngPersonaGraba As Long
    Dim llngError As Long
    Dim vlstrsql As String
    
    If fblnRevisaPermiso(vglngNumeroLogin, CLng(cintNumOpcionFormasPago), "E") Then
        '-----------------------'
        '   Persona que graba   '
        '-----------------------'
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
        EntornoSIHO.ConeccionSIHO.BeginTrans
        vlstrsql = " DELETE FROM pvformapagotipocargocomision WHERE intformapago = " & txtClave.Text
        pEjecutaSentencia vlstrsql
        
        llngError = 1
        frsEjecuta_SP txtClave.Text, "Sp_PvDelFormaPago", False, llngError
        If llngError = 0 Then
            Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "FORMA DE PAGO", txtClave.Text)
            EntornoSIHO.ConeccionSIHO.CommitTrans
        Else
            'No se puede eliminar la información, ya ha sido utilizada.
            MsgBox SIHOMsg(771), vbOKOnly + vbCritical, "Mensaje"
            EntornoSIHO.ConeccionSIHO.RollbackTrans
        End If
        
        txtClave.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Sub cmdEnd_Click()
On Error GoTo NotificaError
    
    grdFormas.Row = grdFormas.Rows - 1
    pMuestraForma grdFormas.TextMatrix(grdFormas.Row, cintColIdForma)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdIncluirComision_Click()
    Dim vlstrSentencia As String
    Dim rsComision As New ADODB.Recordset
    Dim vlintComision As Double
    Dim vlintIvaComision As Double
        
    If cboTipoCargoBancario.ListIndex = -1 Or cboTipoCargoBancario.ListIndex = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        cboTipoCargoBancario.SetFocus
        Exit Sub
    End If
    If cboComisionBancaria.ListIndex = -1 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        cboComisionBancaria.SetFocus
        Exit Sub
    End If
    If Trim(grdComisiones.TextMatrix(1, 1)) = "" Then
        grdComisiones.Row = 1
    Else
        If fblnExisteComision() Then
            cboComisionBancaria.SetFocus
            Exit Sub
        End If
        grdComisiones.Rows = grdComisiones.Rows + 1
        grdComisiones.Row = grdComisiones.Rows - 1
    End If
    
    vlintComision = 0
    vlintIvaComision = 0
    vlstrSentencia = "select * from pvcomisionbancaria where smicvecomision = " & cboComisionBancaria.ItemData(cboComisionBancaria.ListIndex)
    Set rsComision = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    If rsComision.RecordCount > 0 Then
        vlintComision = rsComision!mnycomision
        vlintIvaComision = rsComision!smyIVA
    End If
    
    With grdComisiones
        .TextMatrix(.Row, 1) = txtClave.Text
        .TextMatrix(.Row, 2) = cboTipoCargoBancario.ItemData(cboTipoCargoBancario.ListIndex)
        .TextMatrix(.Row, 3) = cboComisionBancaria.ItemData(cboComisionBancaria.ListIndex)
        .TextMatrix(.Row, 4) = cboTipoCargoBancario.List(cboTipoCargoBancario.ListIndex)
        .TextMatrix(.Row, 5) = cboComisionBancaria.List(cboComisionBancaria.ListIndex)
        .TextMatrix(.Row, 6) = FormatPercent(vlintComision / 100, 2)
        .TextMatrix(.Row, 7) = FormatPercent(vlintIvaComision / 100, 2)
        .TextMatrix(.Row, 8) = IIf(.Row = 1, "*", "")
        .Col = 8
        .CellFontBold = True
        .CellFontSize = 10
    End With
    cboTipoCargoBancario.SetFocus
End Sub

Public Function fblnExisteComision() As Boolean
On Error GoTo NotificaError
    Dim fblnExisteConcepto As Boolean
    Dim vllngContador As Long
    
    fblnExisteComision = False
    vllngContador = 1
    Do While Not fblnExisteConcepto And vllngContador <= grdComisiones.Rows - 1
        If Val(grdComisiones.TextMatrix(vllngContador, 1)) = txtClave.Text And Val(grdComisiones.TextMatrix(vllngContador, 2)) = Trim(cboTipoCargoBancario.ItemData(cboTipoCargoBancario.ListIndex)) Then
            'And Val(grdComisiones.TextMatrix(vllngContador, 3)) = Trim(cboComisionBancaria.ItemData(cboComisionBancaria.ListIndex))
            fblnExisteComision = True
        End If
        vllngContador = vllngContador + 1
    Loop

    If fblnExisteComision Then
        'El tipo de cargo bancario seleccionado ya está registrado para la forma de pago.
        MsgBox SIHOMsg(1337), vbInformation + vbOKOnly, "Mensaje"
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnExisteConcepto"))
    Unload Me
End Function


Private Sub cmdLocate_Click()
On Error GoTo NotificaError
    
    lblnCarga = True
    cboDeptoBusqueda_Click
    sstabFormas.Tab = 1
    cboDeptoBusqueda.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
On Error GoTo NotificaError
    
    If grdFormas.Row < grdFormas.Rows - 1 Then
        grdFormas.Row = grdFormas.Row + 1
    End If
    pMuestraForma grdFormas.TextMatrix(grdFormas.Row, cintColIdForma)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
On Error GoTo NotificaError
    
    Dim vllngPersonaGraba As Long
    Dim lngIdForma As Long
    Dim vlstrsql As String
    Dim strCveMetodoPagoCFDI As String
    Dim vlintcontador As Integer
    Dim tmpTerminal As String
    If fblnDatosValidos() Then
        '-----------------------'
        '   Persona que graba   '
        '-----------------------'
        
        
        
         If chkPinPad.Value <> 0 Then
            
                If Trim(txtImpVoucher.Text) = "" Then
                
                'error sin impresora
                 MsgBox "¡Falta capturar impresora!", vbInformation + vbOKOnly, "Mensaje"
               Exit Sub
                
                End If
               If cboTerminal.ListIndex = -1 Then
               'error sin terminal
                MsgBox "¡Falta seleccionar terminal!", vbInformation + vbOKOnly, "Mensaje"
              Exit Sub
               
               End If
               
            
            
            
            End If
        
        
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
        
            If cboMetodosSATCFDi.ListIndex > -1 Then
                strCveMetodoPagoCFDI = fstrClaveMetodoPago(cboMetodosSATCFDi.ItemData(cboMetodosSATCFDi.ListIndex))
            Else
                strCveMetodoPagoCFDI = " "
            End If

            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            If cboTerminal.ListIndex = -1 Then
                tmpTerminal = ""
            Else
                tmpTerminal = CStr(cboTerminal.ItemData(cboTerminal.ListIndex))
                
               
                
                
            End If
            
            
           
            
            
            
            If Not vlblnConsulta Then
                vgstrParametrosSP = Trim(txtDescripcion.Text) _
                                    & "|" & CStr(chkActiva.Value) _
                                    & "|" & CStr(llngNumCuenta) _
                                    & "|" & CStr(chkFolioReferencia.Value) _
                                    & "|" & CStr(IIf(optMoneda(0).Value, 1, 0)) _
                                    & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                                    & "|" & lstrTipo _
                                    & "|" & strCveMetodoPagoCFDI _
                                    & "|" & CStr(IIf(cboMetodosSAT.ItemData(cboMetodosSAT.ListIndex) > 0, cboMetodosSAT.ItemData(cboMetodosSAT.ListIndex), "")) _
                                    & "|" & IIf(chkPinPad.Value = vbChecked, 1, 0) _
                                    & "|" & Trim(txtImpVoucher.Text) _
                                    & "|" & tmpTerminal
                lngIdForma = 1
                frsEjecuta_SP vgstrParametrosSP, "sp_PvInsFormaPago", True, lngIdForma
                
                txtClave.Text = lngIdForma
                
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "FORMA DE PAGO", txtClave.Text)
            Else
         If Trim(txtImpVoucher.Text) = "" Then
         txtImpVoucher.Text = " "
         End If
                vgstrParametrosSP = Trim(txtDescripcion.Text) _
                                    & "|" & CStr(chkActiva.Value) _
                                    & "|" & CStr(llngNumCuenta) _
                                    & "|" & CStr(chkFolioReferencia.Value) _
                                    & "|" & CStr(IIf(optMoneda(0).Value, 1, 0)) _
                                    & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                                    & "|" & lstrTipo _
                                    & "|" & Trim(txtClave.Text) _
                                    & "|" & strCveMetodoPagoCFDI _
                                    & "|" & CStr(IIf(cboMetodosSAT.ItemData(cboMetodosSAT.ListIndex) > 0, cboMetodosSAT.ItemData(cboMetodosSAT.ListIndex), "")) _
                                    & "|" & IIf(chkPinPad.Value = vbChecked, 1, 0) _
                                    & "|" & Trim(txtImpVoucher.Text) _
                                    & "|" & CStr(IIf(optMoneda(0).Value, 1, 0)) _
                                    & "|" & "1"

                frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdFormaPago"
            
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "FORMA DE PAGO", txtClave.Text)
            End If
        
             '--Guardar comisiones
            vlstrsql = " DELETE FROM pvformapagotipocargocomision WHERE intformapago = " & txtClave.Text
            pEjecutaSentencia vlstrsql
            For vlintcontador = 1 To grdComisiones.Rows - 1
                If grdComisiones.TextMatrix(vlintcontador, 4) <> "" Then
                    With grdComisiones
                        vlstrsql = "INSERT INTO pvformapagotipocargocomision (intformapago, intcvetipocargo, smicvecomision, bitpredeterminado) " & _
                                    "VALUES (" & .TextMatrix(vlintcontador, 1) & "," & .TextMatrix(vlintcontador, 2) & "," & .TextMatrix(vlintcontador, 3) & "," & IIf(.TextMatrix(vlintcontador, 8) = "*", 1, 0) & ")"
                        pEjecutaSentencia vlstrsql
                        'vgstrParametrosSP = vlintclavecomision & "|" & .TextMatrix(vlintContador, 1) & "|" & .TextMatrix(vlintContador, 3)
                        'frsEjecuta_SP vgstrParametrosSP, "Sp_PVINSCOMISIONBANCARIAEMPR"
                    End With
                End If
            Next vlintcontador
        
            EntornoSIHO.ConeccionSIHO.CommitTrans
            txtClave.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub
Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    
    fblnDatosValidos = True
    
    If Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    
    If fblnDatosValidos And llngNumCuenta = 0 And Not optTipo(1).Value And Not optTipo(3).Value Then
        fblnDatosValidos = False
        'Seleccione la cuenta contable.
        MsgBox SIHOMsg(211), vbOKOnly + vbExclamation, "Mensaje"
        mskCuenta.SetFocus
    End If
    
    If fblnDatosValidos And llngNumCuenta <> 0 And Not optTipo(1).Value And Not optTipo(3).Value Then
        If Not fblnCuentaAfectable(mskCuenta.Text, vgintClaveEmpresaContable) Then
            fblnDatosValidos = False
            'La cuenta seleccionada no acepta movimientos.
            MsgBox SIHOMsg(375), vbOKOnly + vbExclamation, "Mensaje"
            mskCuenta.SetFocus
        End If
    End If
    
    If fblnDatosValidos And cboDepartamento.ListIndex < 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        cboDepartamento.SetFocus
    End If
    
    If fblnDatosValidos And Not optMoneda(0).Value And Not optMoneda(1).Value Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        optMoneda(0).SetFocus
    End If
    
    If fblnDatosValidos And Not vlblnConsulta Then
        If fblnExisteForma(txtDescripcion.Text) Then
            fblnDatosValidos = False
            'Este forma de pago ya está registrada.
            MsgBox SIHOMsg(320), vbOKOnly + vbExclamation, "Mensaje"
            txtDescripcion.SetFocus
        End If
    End If
    
    If fblnDatosValidos And optTipo(1).Value And Not vlblnConsulta Then
        'Revisar si ya existe una forma de pago tipo crédito para el departamento, ya que solo se permite una:
        vgstrParametrosSP = "-1|-1|-1|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|-1|C"
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaPago")
        If rs.RecordCount <> 0 Then
            fblnDatosValidos = False
            'No se puede guardar la información, ya existe una forma de pago crédito para este departamento.
            MsgBox SIHOMsg(775), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    
    If fblnDatosValidos And optTipo(3).Value And Not vlblnConsulta Then
        'Revisar si ya existe una forma de pago tipo transferencia para el departamento, ya que solo se permite una:
        vgstrParametrosSP = "-1|-1|" & IIf(optMoneda(0).Value, 1, 0) & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|-1|B"
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormaPago")
        If rs.RecordCount <> 0 Then
            fblnDatosValidos = False
            'No se puede guardar la información, ya existe una forma de pago transferencia bancaria para este departamento.
            MsgBox SIHOMsg(1168), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    
    If fblnDatosValidos Then
        fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, CLng(cintNumOpcionFormasPago), "E")
    End If
    
    If fblnDatosValidos And cboMetodosSATCFDi.ListIndex = -1 Then
        If Not optTipo(1).Value = True Then
            fblnDatosValidos = False
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
            cboMetodosSATCFDi.SetFocus
        End If
    End If

    '- Validar que las monedas del banco y de la forma de pago sean las mismas -'
    If fblnDatosValidos And lblCuentaBanco.Visible Then
        If (optMoneda(0).Value And lintMoneda <> 1) Or (optMoneda(1).Value And lintMoneda <> 0) Then
            fblnDatosValidos = False
            MsgBox "¡La forma de pago y el banco seleccionado no tienen la misma moneda, verifique la información!", vbOKOnly + vbExclamation, "Mensaje"
            optMoneda(0).SetFocus
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function

Private Sub cmdTop_Click()
On Error GoTo NotificaError
    
    grdFormas.Row = 1
    pMuestraForma grdFormas.TextMatrix(grdFormas.Row, cintColIdForma)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError
    
    If rsDepartamento.RecordCount = 0 Then
        MsgBox SIHOMsg(13) + Chr(13) + cboDepartamento.ToolTipText, vbExclamation, "Mensaje"
        Unload Me
        Exit Sub
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            If Me.ActiveControl.Name = "mskCuenta" Then
                If Trim(mskCuenta.ClipText) = "" Then
                    llngNumCuenta = flngBusquedaCuentasContables(False, vgintClaveEmpresaContable)
                    If llngNumCuenta <> 0 Then
                        lblnChange = False
                        mskCuenta.Text = fstrCuentaContable(llngNumCuenta)
                        lblnChange = True
                    End If
                Else
                    lblnChange = False
                    mskCuenta.Mask = ""
                    mskCuenta.Text = fstrCuentaCompleta(mskCuenta.Text)
                    mskCuenta.Mask = vgstrEstructuraCuentaContable
                    lblnChange = True
                    
                    llngNumCuenta = flngNumeroCuenta(mskCuenta.Text, vgintClaveEmpresaContable)
                End If
        
                lblDescripcionCuenta.Caption = fstrDescripcionCuenta(mskCuenta.Text, vgintClaveEmpresaContable)
                
                '(CR) Agregado - Revisar si es cuenta de banco -'
                If fblnEsCuentaBanco(llngNumCuenta) Then
                    lblCuentaBanco.Visible = True
                Else
                    If lblCuentaBanco.Visible Then
                        MsgBox "No se realizarán movimientos al libro de ingresos y egresos de bancos para esta forma de pago.", vbOKOnly + vbExclamation, "Mensaje"
                    End If
                    lblCuentaBanco.Visible = False
                End If
                
                cboDepartamento.SetFocus
            Else
                If Me.ActiveControl.Name = "txtClave" Then
                    If Trim(txtClave.Text) = "" Then
                        pLimpia
                        SendKeys vbTab
                    Else
                        pMuestraForma CLng(txtClave.Text)
                        If vlblnConsulta Then
                            pHabilita 0, 0, 1, 0, 0, 0, 1, 1
                            cmdLocate.SetFocus
                        Else
                            SendKeys vbTab
                        End If
                    End If
                Else
                    SendKeys vbTab
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vgstrParametrosSP = "-1|1|*|" & vgintClaveEmpresaContable
    Set rsDepartamento = frsEjecuta_SP(vgstrParametrosSP, "sp_GnSelDepartamento")
    If rsDepartamento.RecordCount > 0 Then
        pLlenarCboRs cboDepartamento, rsDepartamento, 0, 1
        pLlenarCboRs cboDeptoBusqueda, rsDepartamento, 0, 1
        cboDeptoBusqueda.AddItem "<TODOS>", 0
        cboDeptoBusqueda.ItemData(cboDeptoBusqueda.newIndex) = -1
        cboDeptoBusqueda.ListIndex = 0
    End If
    
    vgstrParametrosSP = "select intcvetipocargo, chrdescripcion from pvtipocargobancario where bitactivo = 1 order by chrdescripcion"
    Set rs = frsRegresaRs(vgstrParametrosSP, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboTipoCargoBancario, rs, 0, 1, 3
    Else
        cboTipoCargoBancario.AddItem "<TODOS>", 0
    End If
    
    vgstrParametrosSP = "select * from pvComisionBancaria where bitactivo = 1 order by chrdescripcion"
    Set rs = frsRegresaRs(vgstrParametrosSP, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboComisionBancaria, rs, 0, 1
    End If
    
    pCargaMetodosSAT
    
    cboMetodosSAT.Enabled = True
    lblMetodoPagoSAT.Enabled = True
    
    cboMetodosSATCFDi.Enabled = True
    lblMetodoPagoCFDi.Enabled = True
    
    pCargaTerminales
    chkPinPad_Click
    pConfiguraGrid
    
    sstabFormas.Tab = 0
    fraTipo.BorderStyle = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

'Private Sub pIncluyeComision(vlstrConcepto As String, _
'                             vldblCantidad As Double, _
'                             vldblDescuento As Double, _
'                             vldblIVA As Double, _
'                             vllngCveConcepto As Long, _
'                             vllngCtaIngresos As Long, _
'                             vllngCtaDescuentos As Long, _
'                             vllngCtaIVA As Long, _
'                             vlstrTipoCargo As String, _
'                             Optional vllngintFolioDocumento As Long, _
'                             Optional vllngCveDepartamento As Long)
'On Error GoTo NotificaError
'
'    If Trim(grdNotas.TextMatrix(1, 2)) = "" Then
'        grdNotas.Row = 1
'    Else
'        grdNotas.Rows = grdNotas.Rows + 1
'        grdNotas.Row = grdNotas.Rows - 1
'    End If
'
'    With grdNotas
'        If chkFacturasPaciente.Value = 0 Then
'            .TextMatrix(.Row, vlintColFactura) = IIf(sstFacturasCreditos.Tab = 0, cboFactura.List(cboFactura.ListIndex), cboCreditosDirectos.List(cboCreditosDirectos.ListIndex))
'        Else
'            .TextMatrix(.Row, vlintColFactura) = cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)
'        End If
'        .TextMatrix(.Row, vlintColDescripcionConcepto) = vlstrConcepto
'        .TextMatrix(.Row, vlintColCantidad) = FormatCurrency(vldblCantidad, 2)
'        .TextMatrix(.Row, vlintColDescuento) = FormatCurrency(vldblDescuento, 2)
'        .TextMatrix(.Row, vlintColIVA) = FormatCurrency(vldblIVA, 2)
'        .TextMatrix(.Row, vlintColCveConcepto) = vllngCveConcepto
'        .TextMatrix(.Row, vlintColCtaIngresos) = vllngCtaIngresos
'        .TextMatrix(.Row, vlintColCtaDescuentos) = vllngCtaDescuentos
'        .TextMatrix(.Row, vlintColCtaIVA) = vllngCtaIVA
'        .TextMatrix(.Row, vlintColTipoCargo) = vlstrTipoCargo
'        .TextMatrix(.Row, vlintColTipoNotaFARCDetalle) = IIf(sstFacturasCreditos.Tab = 0, "FA", "MA")
'        .TextMatrix(.Row, 13) = vllngintFolioDocumento
'        .TextMatrix(.Row, 14) = vllngCveDepartamento
'        .TextMatrix(.Row, vlintColIVANotaSinRedondear) = FormatCurrency(vldblIVA, 15)
'    End With
'
'    vldblSubTotalTemporal = vldblSubTotalTemporal + vldblCantidad
'    txtSubtotal.Text = FormatCurrency(vldblSubTotalTemporal, 2)
'
'    vldblDescuentoTemporal = vldblDescuentoTemporal + vldblDescuento
'    txtDescuentoTot.Text = FormatCurrency(vldblDescuentoTemporal, 2)
'
'    vldblIvaTemporal = vldblIvaTemporal + vldblIVA
'    txtIva.Text = FormatCurrency(vldblIvaTemporal, 2)
'
'    vldTotal = vldblSubTotalTemporal - vldblDescuentoTemporal + vldblIvaTemporal
'    txtTotal.Text = FormatCurrency(vldTotal, 2)
'
'Exit Sub
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIncluyeConcepto"))
'    Unload Me
'End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If sstabFormas.Tab = 0 Then
        If cmdSave.Enabled Or vlblnConsulta Then
            Cancel = True
            ' ¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                txtClave.SetFocus
            End If
        End If
    Else
        Cancel = True
        sstabFormas.Tab = 0
        'txtClave.SetFocus
    End If
End Sub

Private Sub grdComisiones_Click()
    Dim vlintCvePaquete As Integer
    Dim vlIntCont As Integer
    With grdComisiones
        If .Col = 8 And .TextMatrix(.Row, 4) <> "" Then
            .TextMatrix(.Row, 8) = "*"
            .CellFontBold = True
            .CellFontSize = 10
            vlintCvePaquete = .Row
            For vlIntCont = 1 To .Rows - 1
                If vlIntCont <> vlintCvePaquete Then
                    .TextMatrix(vlIntCont, 8) = ""
                End If
            Next vlIntCont
        End If
    End With
End Sub

Private Sub grdComisiones_DblClick()
    Dim vlReubicarPredeterminado As Boolean
    
    If grdComisiones.Rows - 1 = 1 Then
        pLimpiaGrid
        pConfiguraGrid
    Else
        If grdComisiones.TextMatrix(grdComisiones.Row, 8) = "*" Then
            vlReubicarPredeterminado = True
        Else
            vlReubicarPredeterminado = False
        End If
        pBorrarRegMshFGrd grdComisiones, grdComisiones.Row
        If grdComisiones.Rows >= 2 And vlReubicarPredeterminado Then
            grdComisiones.TextMatrix(grdComisiones.Rows - 1, 8) = "*"
            grdComisiones.Col = 8
            grdComisiones.CellFontBold = True
            grdComisiones.CellFontSize = 10
        End If
    End If
End Sub

Private Sub pLimpiaGrid()
On Error GoTo NotificaError

     With grdComisiones
        .Clear
        .Cols = 9
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|||Tipo de cargo bancario|Comisión bancaria|Porcentaje|IVA|Predeterminado"
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridNota"))
    Unload Me
End Sub


Private Sub grdFormas_DblClick()
On Error GoTo NotificaError
    
    If Trim(grdFormas.TextMatrix(grdFormas.Row, cintColIdForma)) <> "" Then
        pMuestraForma grdFormas.TextMatrix(grdFormas.Row, cintColIdForma)
        pHabilita 1, 1, 1, 1, 1, 0, 1, 1
        sstabFormas.Tab = 0
        
        cmdLocate.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdFormas_DblClick"))
End Sub

Private Sub grdFormas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdFormas_DblClick
    End If
End Sub

Private Sub mskCuenta_Change()
    If lblnChange Then
        lblDescripcionCuenta.Caption = ""
        llngNumCuenta = 0
    End If
End Sub

Private Sub mskCuenta_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
    pSelMkTexto mskCuenta
End Sub

Private Sub optMoneda_GotFocus(Index As Integer)
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optMoneda_GotFocus"))
End Sub

Private Function fblnExisteForma(vlstrxConcepto As String) As Boolean
    Dim vlrsAux As ADODB.Recordset

    fblnExisteForma = False
    If frsRegresaRs("SELECT * FROM pvFormaPago WHERE RTRIM(chrDescripcion) = '" & RTrim(txtDescripcion) & "' AND smiDepartamento = " & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))).RecordCount > 0 Then fblnExisteForma = True
End Function

Private Sub optTipo_Click(Index As Integer)
    chkFolioReferencia.Enabled = Index = 2 Or Index = 3 Or Index = 4
    If Not chkFolioReferencia.Enabled Then
        chkFolioReferencia.Value = 0
    End If
    
    optMoneda(0).Enabled = Index <> 1 'And Index <> 3
    optMoneda(1).Enabled = Index <> 1 'And Index <> 3
    lblMoneda.Enabled = Index <> 1 'And Index <> 3
    If Not lblMoneda.Enabled Then
        optMoneda(0).Value = True
    End If
    
    lblCuentaContable.Enabled = Index <> 1 And Index <> 3
    mskCuenta.Enabled = Index <> 1 And Index <> 3
    If Not mskCuenta.Enabled Then
        mskCuenta.Mask = ""
        mskCuenta.Text = ""
        mskCuenta.Mask = vgstrEstructuraCuentaContable
    End If
    cmdComisiones.Enabled = False
    chkPinPad.Value = vbUnchecked
    chkPinPad.Enabled = False
    Select Case Index
        Case 0
            lstrTipo = "E" 'Efectivo
        Case 1
            lstrTipo = "C" 'Crédito
        Case 2
            lstrTipo = "T"  'Tarjeta
            cmdComisiones.Enabled = True
            chkPinPad.Enabled = True
        Case 3
            lstrTipo = "B"  'Transferencia bancaria
        Case 4
            lstrTipo = "H"  'Cheque
    End Select
    '(CR) Agregado - Revisar si es cuenta de banco -'
    lblCuentaBanco.Visible = fblnEsCuentaBanco(llngNumCuenta)
    
    If optTipo(1).Value Then
        cboMetodosSAT.Enabled = False
        cboMetodosSAT.ListIndex = 0
        lblMetodoPagoSAT.Enabled = False
        
        cboMetodosSATCFDi.Enabled = False
        cboMetodosSATCFDi.ListIndex = 0
        lblMetodoPagoCFDi.Enabled = False
    Else
        cboMetodosSAT.Enabled = True
        lblMetodoPagoSAT.Enabled = True
        
        cboMetodosSATCFDi.Enabled = True
        lblMetodoPagoCFDi.Enabled = True
        
    End If
End Sub

Private Sub optTipo_GotFocus(Index As Integer)
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
End Sub

Private Sub sstabFormas_Click(PreviousTab As Integer)
On Error GoTo NotificaError
    
    If sstabFormas.Tab = 1 Then
        grdFormas.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstabFormas_Click"))
End Sub

Private Sub txtClave_GotFocus()
On Error GoTo NotificaError
    
    pLimpia
    pLimpiaGrid
    pHabilita 0, 0, 1, 0, 0, 0, 0, 0
    pSelTextBox txtClave

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_GotFocus"))
End Sub

Private Sub pLimpia()
On Error GoTo NotificaError
    
    lblnChange = True
    vlblnConsulta = False
    
    txtClave.Text = frsRegresaRs("SELECT ISNULL(MAX(intFormaPago), 0) + 1 FROM PvFormaPago").Fields(0)
    txtDescripcion.Text = ""
    'txtDescripcionCFD.Text = ""
    
    optTipo(0).Value = True
    cboMetodosSAT.Enabled = True
    cboMetodosSATCFDi.Enabled = True
    lblMetodoPagoSAT.Enabled = True
    lblMetodoPagoCFDi.Enabled = True
    optTipo_Click 0
    
    chkFolioReferencia.Value = 0
    
    optMoneda(0).Value = True
    
    mskCuenta.Mask = ""
    mskCuenta.Text = ""
    mskCuenta.Mask = vgstrEstructuraCuentaContable
    llngNumCuenta = 0
    
    cboDepartamento.ListIndex = -1

    chkActiva.Value = 1
    
    lblnCarga = False
    cboDeptoBusqueda.ListIndex = 0
    
    lblCuentaBanco.Visible = False

    cboMetodosSAT.ListIndex = 0
    cboMetodosSATCFDi.ListIndex = 0
    
    chkPinPad.Value = vbUnchecked
    chkPinPad_Click
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer, vlb4 As Integer, vlb5 As Integer, vlb6 As Integer, vlb7 As Integer, vlb8 As Integer)
On Error GoTo NotificaError
    
    cmdTop.Enabled = vlb1 = 1
    cmdBack.Enabled = vlb2 = 1
    cmdLocate.Enabled = vlb3 = 1
    cmdNext.Enabled = vlb4 = 1
    cmdEnd.Enabled = vlb5 = 1
    cmdSave.Enabled = vlb6 = 1
    cmdDelete.Enabled = vlb7 = 1
    If optTipo(2).Value Then
        cmdComisiones.Enabled = vlb8 = 1
    Else
        cmdComisiones.Enabled = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_KeyPress"))
End Sub

Private Sub pMuestraForma(vllngxNumero As Long)
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim rsMetodos As New ADODB.Recordset
    
    vgstrParametrosSP = CStr(vllngxNumero) & "|-1|-1|-1|-1|*"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormaPago")
    If rs.RecordCount <> 0 Then
        vlblnConsulta = True
    
        txtClave.Text = str(rs!intFormaPago)
        txtDescripcion.Text = Trim(rs!chrDescripcion)
        
        optTipo(0).Value = rs!chrTipo = "E"
        optTipo(1).Value = rs!chrTipo = "C"
        optTipo(2).Value = rs!chrTipo = "T"
        optTipo(3).Value = rs!chrTipo = "B"
        optTipo(4).Value = rs!chrTipo = "H"
        
        chkFolioReferencia.Value = rs!bitpreguntafolio
        
        optMoneda(0).Value = rs!BITPESOS = 1
        optMoneda(1).Value = rs!BITPESOS = 0
        
        lblnChange = False
        llngNumCuenta = rs!INTCUENTACONTABLE
        mskCuenta.Mask = ""
        If optTipo(1).Value Or optTipo(3).Value Then
            mskCuenta.Text = ""
        Else
            mskCuenta.Text = fstrCuentaContable(rs!INTCUENTACONTABLE)
            lblDescripcionCuenta.Caption = fstrDescripcionCuenta(mskCuenta.Text, vgintClaveEmpresaContable)
        End If
        mskCuenta.Mask = vgstrEstructuraCuentaContable
        lblnChange = True
        
        cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, rs!SMIDEPARTAMENTO)
        
        chkActiva.Value = rs!bitestatusactivo
        
        If IsNull(rs!VCHDESCRIPCIONCFD) Then
            cboMetodosSATCFDi.ListIndex = 0
        Else
        
            cboMetodosSATCFDi.ListIndex = flngLocalizaCbo(cboMetodosSATCFDi, fintIDMetodoPago(rs!VCHDESCRIPCIONCFD))
        End If
                 
        lblCuentaBanco.Visible = fblnEsCuentaBanco(llngNumCuenta) '(CR) Agregado para caso 7442
        cboTipoCargoBancario.Text = "<TODOS>"
        
        cboMetodosSAT.ListIndex = 0
        cboMetodosSAT.ListIndex = fintLocalizaCbo(cboMetodosSAT, IIf(IsNull(rs!INTIDFORMAPAGOSAT), 0, rs!INTIDFORMAPAGOSAT))
        
        If optTipo(1).Value Then
            cboMetodosSAT.Enabled = False
            cboMetodosSAT.ListIndex = 0
            lblMetodoPagoSAT.Enabled = False
            
            cboMetodosSATCFDi.Enabled = False
            cboMetodosSATCFDi.ListIndex = 0
            lblMetodoPagoCFDi.Enabled = False
            
        Else
            cboMetodosSAT.Enabled = True
            lblMetodoPagoSAT.Enabled = True
            
            cboMetodosSATCFDi.Enabled = True
            lblMetodoPagoCFDi.Enabled = True
        End If
        
        chkPinPad.Value = IIf(rs!BITUTILIZARPINPAD = 1, vbChecked, vbUnchecked)
        
        
        cboTerminal.ListIndex = fintLocalizaCbo(cboTerminal, IIf(IsNull(rs!intCveTerminal), "", rs!intCveTerminal))
        txtImpVoucher.Text = IIf(IsNull(rs!VCHIMPRESORAVOUCHER), "", rs!VCHIMPRESORAVOUCHER)
    Else
        pLimpia
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraForma"))
End Sub

Private Sub txtDescripcion_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
    pSelTextBox txtDescripcion

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIniciaCaptura"))
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_GotFocus"))
End Sub

'- (CR) Función agregada para verificar si la cuenta de la forma de pago pertenece a un banco -'
Private Function fblnEsCuentaBanco(llngNumeroCuenta As Long) As Boolean
On Error GoTo NotificaError

    Dim rsCuentaBanco As New ADODB.Recordset
    Dim lstrSentencia As String
    
    lstrSentencia = "SELECT bitEstatusMoneda FROM CpBanco WHERE bitEstatus = 1 AND intNumeroCuenta = " & llngNumeroCuenta
    Set rsCuentaBanco = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    fblnEsCuentaBanco = rsCuentaBanco.RecordCount <> 0
    If fblnEsCuentaBanco Then lintMoneda = rsCuentaBanco!bitestatusmoneda
    rsCuentaBanco.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnEsCuentaBanco"))
End Function

Private Function fstrClaveMetodoPago(intCveMetodoPago As Integer) As String
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select CHRCLAVE from PVMETODOPAGOSATCFDI where INTIDREGISTRO = " & intCveMetodoPago)
    If Not rs.EOF Then
        fstrClaveMetodoPago = rs!CHRCLAVE
    Else
        fstrClaveMetodoPago = ""
    End If
    rs.Close
End Function

Private Function fintIDMetodoPago(strCveMetodoPago As String) As Integer
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select INTIDREGISTRO from PVMETODOPAGOSATCFDI where CHRCLAVE = '" & strCveMetodoPago & "'")
    If Not rs.EOF Then
        fintIDMetodoPago = rs!intIdRegistro
    Else
        fintIDMetodoPago = -1
    End If
    rs.Close
End Function

Private Sub txtImpVoucher_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 1
    pSelTextBox txtImpVoucher
End Sub


