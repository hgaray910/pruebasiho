VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantoCajaChica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de salida de caja chica"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstab 
      Height          =   3225
      Left            =   -45
      TabIndex        =   12
      Top             =   -480
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   5689
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoCajaChica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoCajaChica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmMantoCajaChica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtEstructuraem"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtConceptosDe"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "grdConceptoEmpresas"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtEstructuraem 
         Height          =   285
         Left            =   -69120
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txtConceptosDe 
         Height          =   285
         Left            =   -69240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Height          =   2595
         Left            =   -74910
         TabIndex        =   20
         Top             =   450
         Width           =   8595
         Begin VB.ComboBox cboEmpresaContable 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   750
            Width           =   6465
         End
         Begin VB.Frame Frame6 
            Caption         =   "Cuentas contables"
            Height          =   1100
            Left            =   120
            TabIndex        =   21
            Top             =   1280
            Width           =   8295
            Begin VB.TextBox txtCuentaGasto 
               Height          =   315
               Left            =   3915
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   430
               Width           =   4230
            End
            Begin MSMask.MaskEdBox mskCuentaGasto 
               Height          =   315
               Left            =   1725
               TabIndex        =   26
               ToolTipText     =   "Cuenta de ingresos"
               Top             =   430
               Width           =   2160
               _ExtentX        =   3810
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta de gasto"
               Height          =   195
               Left            =   135
               TabIndex        =   23
               Top             =   490
               Width           =   1170
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   300
            Width           =   690
         End
         Begin VB.Label lblConcepto 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   1830
            TabIndex        =   30
            ToolTipText     =   "Descripción del concepto de entrada/salida de dinero"
            Top             =   240
            Width           =   6465
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Empresa contable"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   810
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1755
         Left            =   165
         TabIndex        =   15
         Top             =   510
         Width           =   8475
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1650
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Departamento"
            Top             =   990
            Width           =   6690
         End
         Begin VB.TextBox txtClave 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1650
            MaxLength       =   5
            TabIndex        =   0
            ToolTipText     =   "Clave "
            Top             =   285
            Width           =   1005
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1650
            MaxLength       =   100
            TabIndex        =   1
            ToolTipText     =   "Descripción "
            Top             =   630
            Width           =   6690
         End
         Begin VB.CheckBox chkActiva 
            Caption         =   "Activo"
            Height          =   200
            Left            =   1650
            TabIndex        =   3
            ToolTipText     =   "Estado"
            Top             =   1440
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   345
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   225
            TabIndex        =   17
            Top             =   690
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   225
            TabIndex        =   16
            Top             =   1443
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2025
         TabIndex        =   14
         Top             =   2280
         Width           =   4650
         Begin VB.CommandButton cmdCuentas 
            Caption         =   "Cuentas contables"
            Height          =   495
            Left            =   3600
            TabIndex        =   19
            Top             =   165
            Width           =   975
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   90
            Picture         =   "frmMantoCajaChica.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Primer registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   585
            Picture         =   "frmMantoCajaChica.frx":0456
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anterior registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1080
            Picture         =   "frmMantoCajaChica.frx":05C8
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Búsqueda"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1590
            Picture         =   "frmMantoCajaChica.frx":073A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Siguiente registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2100
            Picture         =   "frmMantoCajaChica.frx":08AC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Ultimo registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2595
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCajaChica.frx":0D9E
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar"
            Top             =   165
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3090
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCajaChica.frx":10E0
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Borrar"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2580
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   8520
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBusqueda 
            Height          =   2340
            Left            =   60
            TabIndex        =   10
            Top             =   165
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   4128
            _Version        =   393216
            GridColor       =   12632256
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConceptoEmpresas 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
End
Attribute VB_Name = "frmMantoCajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
' Programa para dar mantenimiento a los conceptos de salida de dinero de caja chica
' Fecha de programación: Miércoles 26 de Septiembre de 2006
'-----------------------------------------------------------------------------------

Const cintColId = 1
Const cintColDescripcion = 2
Const cintCols = 3
Const cstrColumnas = "|Clave|Descripción"

Dim llngNumCuenta As Long 'Id. de la cuenta seleccionada
'Dim llngNumCuentaDescuento As Long 'Id. de la cuenta seleccionada

Dim lblnConsulta As Boolean
Dim lblnChange As Boolean 'Para ejecutar o no lo que está en el evento mskCuenta_Change
'Dim lblnChangeDescuento As Boolean 'Para ejecutar o no lo que está en el evento mskCuentaDescuento_Change
Dim llngCveConceptoRecepContado As Long 'Clave del concepto de caja chica usado para recepciones pagadas de contado
Dim vblncuentagasto As Boolean
Dim vlblnentra As Boolean


Private Sub pCarga()
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP("-1|-1", "SP_PVSELCONCEPTOCAJACHICA")
    
    With grdBusqueda
        .Clear
        .Cols = cintCols
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        
        .FormatString = cstrColumnas
        .ColWidth(cintColId) = 1000
        .ColWidth(cintColDescripcion) = 6500
        
        .ColAlignment(cintColId) = flexAlignRightCenter
        .ColAlignment(cintColDescripcion) = flexAlignLeftCenter
        
        .ColAlignmentFixed(cintColId) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDescripcion) = flexAlignCenterCenter
        
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                .TextMatrix(.Rows - 1, cintColId) = rs!intConsecutivo
                .TextMatrix(.Rows - 1, cintColDescripcion) = Trim(rs!VCHDESCRIPCION)
                rs.MoveNext
                .Rows = .Rows + 1
            Loop
            .Rows = .Rows - 1
        End If
    End With
    
End Sub

Private Sub cboDepartamento_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
 
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_GotFocus"))
End Sub

Private Sub cboEmpresaContable_Click()
  Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim intIndex As Integer
    mskCuentaGasto.Mask = ""
    mskCuentaGasto.Text = """"
    txtCuentaGasto.Text = ""
    txtEstructuraem.Text = ""
    If cboEmpresaContable.ListIndex > -1 Then
        
        'strSQL = "select * from CnParametro where tnyClaveEmpresa=" & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
        Set rs = frsSelParametros("CN", cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex), "VCHESTRUCTURACUENTACONTABLE")
        If Not rs.EOF Then
            txtEstructuraem.Text = rs!valor
        End If
        rs.Close
        For intIndex = 0 To grdConceptoEmpresas.Rows - 1
            If CLng(grdConceptoEmpresas.TextMatrix(intIndex, 0)) = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) Then
                
                If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 2)) Then
                    mskCuentaGasto.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 2)))
                    txtCuentaGasto.Text = fstrDescripcionCuenta(mskCuentaGasto.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
              '  If IsNumeric(grdConceptoEmpresas.TextMatrix(intIndex, 3)) Then
              '      MskCuentaDescuentos.Text = fstrCuentaContable(CLng(grdConceptoEmpresas.TextMatrix(intIndex, 3)))
              '      txtCuentaDescuento.Text = fstrDescripcionCuenta(MskCuentaDescuentos.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
              '  End If
            End If
        Next
    End If
    mskCuentaGasto.Mask = txtEstructuraem.Text
    'MskCuentaDescuentos.Mask = txtEstructuraem.Text
  
End Sub



Private Sub chkActiva_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActiva_GotFocus"))
End Sub

Private Sub cmdBack_Click()
    On Error GoTo NotificaError
    
    If grdBusqueda.Row > 1 Then
        grdBusqueda.Row = grdBusqueda.Row - 1
    End If
    pMuestra grdBusqueda.TextMatrix(grdBusqueda.Row, cintColId)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdCuentas_Click()
    vlblnentra = True
    pCargaConceptos
    cboEmpresaContable.ListIndex = -1
    cboEmpresaContable.ListIndex = flngLocalizaCbo(cboEmpresaContable, CStr(vgintClaveEmpresaContable))
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 1113, 368), "C") Then
        cboEmpresaContable.Enabled = True
    Else
        cboEmpresaContable.Enabled = False
    End If
    lblConcepto.Caption = txtDescripcion.Text
    sstab.Tab = 2
    If cboEmpresaContable.Enabled = True Then
        cboEmpresaContable.SetFocus
    Else
        mskCuentaGasto.SetFocus
    End If
End Sub

Private Sub pCargaConceptos()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    If txtConceptosDe.Text <> txtClave.Text Then
        txtConceptosDe.Text = txtClave.Text
        grdConceptoEmpresas.Rows = 0
        strSQL = "select * from PVConceptocajachicaempresa where intnumconcepto = " & txtClave.Text
        Set rs = frsRegresaRs(strSQL)
        If rs.RecordCount <> 0 Then
            vblncuentagasto = True
            Do Until rs.EOF
                grdConceptoEmpresas.AddItem ""
                grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 0) = rs!tnyClaveEmpresa
                grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 1) = rs!intNumConcepto
                grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 2) = IIf(IsNull(rs!intcuentagasto), "", rs!intcuentagasto)
                rs.MoveNext
            Loop
        Else
            vblncuentagasto = False
        End If
        rs.Close
    End If
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    
    Dim lngPersonaGraba As Long
    Dim lngError As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, CLng(cintNumOpcionConceptoCaja), "E") Then
        If lblnConsulta And llngCveConceptoRecepContado <> -1 And llngCveConceptoRecepContado = Val(txtClave.Text) Then
            'El concepto de salida de caja chica está en uso,
            MsgBox SIHOMsg(1432) & " no se puede borrar.", vbOKOnly + vbExclamation, "Mensaje"
        Else
            '--------------------------------------------------------
            ' Persona que graba
            '--------------------------------------------------------
            lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If lngPersonaGraba = 0 Then Exit Sub
            
            lngError = 1
            frsEjecuta_SP txtClave.Text, "sp_PvDelConceptoCajaChica", False, lngError
            
            If lngError = 0 Then
                Call pGuardarLogTransaccion(Me.Name, EnmBorrar, lngPersonaGraba, "CONCEPTO CAJA CHICA", txtClave.Text)
            Else
                'No se puede eliminar la información, ya ha sido utilizada.
                MsgBox SIHOMsg(771), vbOKOnly + vbCritical, "Mensaje"
            End If
            txtClave.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
    
    grdBusqueda.Row = grdBusqueda.Rows - 1
    pMuestra grdBusqueda.TextMatrix(grdBusqueda.Row, cintColId)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    sstab.Tab = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError
    
    If grdBusqueda.Row < grdBusqueda.Rows - 1 Then
        grdBusqueda.Row = grdBusqueda.Row + 1
    End If
    pMuestra grdBusqueda.TextMatrix(grdBusqueda.Row, cintColId)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    Dim lngPersonaGraba As Long
    Dim lngIdConcepto As Long
    
    If fblnDatosValidos() Then
        If lblnConsulta And llngCveConceptoRecepContado <> -1 And llngCveConceptoRecepContado = Val(txtClave.Text) And chkActiva = 0 Then
            'El concepto de salida de caja chica está en uso,
            MsgBox SIHOMsg(1432) & " no se puede cambiar su estado.", vbOKOnly + vbExclamation, "Mensaje"
            chkActiva.Value = 1
        Else
            '--------------------------------------------------------
            ' Persona que graba
            '--------------------------------------------------------
            lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If lngPersonaGraba <> 0 Then
                
                If vblncuentagasto = False Then     'Si no tiene cuenta asignada al concepto
                        'Debe agregar la cuenta de gasto para poder guardar el concepto.
                        MsgBox SIHOMsg(1556), vbOKOnly + vbExclamation, "Mensaje"
                Else
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                    
                    If Not lblnConsulta Then
                        vgstrParametrosSP = Trim(txtDescripcion.Text) & "|" & CStr(chkActiva.Value) & "|" & IIf(cboDepartamento.ListIndex = 0, "", cboDepartamento.ItemData(cboDepartamento.ListIndex))
                        lngIdConcepto = 1
                        frsEjecuta_SP vgstrParametrosSP, "sp_PvInsConceptoCajaChica", False, lngIdConcepto
                    Else
                        vgstrParametrosSP = Trim(txtClave.Text) & "|" & Trim(txtDescripcion.Text) & "|" & CStr(chkActiva.Value) & "|" & IIf(cboDepartamento.ListIndex = 0, "", cboDepartamento.ItemData(cboDepartamento.ListIndex))
                        frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdConceptoCajaChica"
                        lngIdConcepto = txtClave.Text
                    End If
                    pGuardaDetalle lngIdConcepto
                    pGuardarLogTransaccion Me.Name, IIf(lblnConsulta, EnmCambiar, EnmGrabar), lngPersonaGraba, "CONCEPTO CAJA CHICA", txtClave.Text
                    
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    
                    txtClave.SetFocus
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub
Private Sub pGuardaDetalle(vllngidcurrentval As Long)

    Dim intIndex As Integer
    Dim strParametros As String
    For intIndex = 0 To grdConceptoEmpresas.Rows - 1
        If grdConceptoEmpresas.TextMatrix(intIndex, 2) <> "" Then
           strParametros = grdConceptoEmpresas.TextMatrix(intIndex, 0) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 1) & "|" & grdConceptoEmpresas.TextMatrix(intIndex, 2) & "|" & vllngidcurrentval
      End If
        frsEjecuta_SP strParametros, "sp_PvUpdConceptCajaChicaEmpres"
    Next
  
End Sub
Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    
    fblnDatosValidos = True
    
    If fblnDatosValidos Then
        fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, CLng(cintNumOpcionConceptoCaja), "E")
    End If
    If fblnDatosValidos And Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    If fblnDatosValidos And cboDepartamento.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        cboDepartamento.SetFocus
    End If

    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
    
    grdBusqueda.Row = 1
    pMuestra grdBusqueda.TextMatrix(grdBusqueda.Row, cintColId)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            If Me.ActiveControl.Name = "txtClave" Then
                If Trim(txtClave.Text) = "" Then
                    pLimpia
                    SendKeys vbTab
                Else
                    
                    pMuestra CLng(txtClave.Text)
                    
                    If lblnConsulta Then
                        pHabilita 0, 0, 1, 0, 0, 0, 1
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

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim rsDeptoCaja As New ADODB.Recordset
    vlblncuentagasto = False
    vlblnentra = True
    Me.Icon = frmMenuPrincipal.Icon
       
    sstab.Tab = 0
    pCargaCombos
    
    'concepto de caja chica para recepciones de contado
    llngCveConceptoRecepContado = -1
    Set rsDeptoCaja = frsSelParametros("IV", vgintClaveEmpresaContable, "INTCONCEPTOCAJACHICACONTADO")
    If Not rsDeptoCaja.EOF Then
        If Not IsNull(rsDeptoCaja!valor) Then llngCveConceptoRecepContado = CLng(rsDeptoCaja!valor)
    End If
    rsDeptoCaja.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub
Private Sub pCargaCombos()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    strSQL = "select * from CNEmpresaContable where bitActiva <> 0 order by vchNombre"
    Set rs = frsRegresaRs(strSQL)
    If Not rs.EOF Then
        pLlenarCboRs cboEmpresaContable, rs, 0, 1
    End If
    rs.Close
    
    cboDepartamento.Clear
    Set rs = frsEjecuta_SP("-1|1|*|-1", "sp_GnSelDepartamento")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboDepartamento, rs, 0, 1
    End If
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ListIndex = 0
    rs.Close
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    If sstab.Tab = 0 Then
        If cmdSave.Enabled Or lblnConsulta Then
            Cancel = True
            ' ¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                txtClave.SetFocus
            End If
        End If
    Else
        Cancel = True
        sstab.Tab = 0
        txtDescripcion.SetFocus
    End If

End Sub

Private Sub grdBusqueda_DblClick()
    On Error GoTo NotificaError
    
    If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cintColId)) <> "" Then
        pMuestra grdBusqueda.TextMatrix(grdBusqueda.Row, cintColId)
        pHabilita 1, 1, 1, 1, 1, 0, 1
        sstab.Tab = 0
        cmdLocate.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdBusqueda_DblClick"))
End Sub

Private Sub grdBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        grdBusqueda_DblClick
    End If

End Sub

Private Sub mskCuentaGasto_Change()
  pAsignaCuentaImproved mskCuentaGasto, txtCuentaGasto
End Sub
Private Sub pAsignaCuentaImproved(mskCuenta As MaskEdBox, txtCuenta As TextBox)
  Dim rs As New ADODB.Recordset
  txtCuenta.Text = ""
  If cboEmpresaContable.ListIndex > -1 Then
    Set rs = frsRegresaRs("SELECT vchCuentaContable, vchDescripcionCuenta, intNumeroCuenta FROM cnCuenta WHERE vchCuentaContable = '" & mskCuenta.Text & "' AND tnyClaveEmpresa = " & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) & " ORDER BY vchCuentaContable")
    If (rs.State <> adStateClosed) Then
      If rs.RecordCount > 0 Then
        txtCuenta.Text = rs!vchDescripcionCuenta
        vblncuentagasto = True
        Else
        vblncuentagasto = False
      End If
      rs.Close
    End If
   End If
End Sub

Private Sub mskCuentaGasto_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskCuentaGasto

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaGasto_GotFocus"))
    Unload Me
End Sub

Private Sub mskCuentaGasto_KeyPress(KeyAscii As Integer)
  On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaGasto, txtCuentaGasto
    Else
        If KeyAscii = 8 Then vblncuentagasto = False
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaGasto_KeyPress"))
End Sub
Private Sub pAsignaCuenta(mskObject As MaskEdBox, txtObject As TextBox)
    On Error GoTo NotificaError
    
    Dim vllngNumeroCuenta As Long
    Dim vlstrCuentaCompleta As String
    If cboEmpresaContable.ListIndex = -1 Then
        Exit Sub
    End If

    If Trim(mskObject.ClipText) = "" Then
        vllngNumeroCuenta = flngBusquedaCuentasContables(False, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
        If vllngNumeroCuenta <> 0 Then
            mskObject.Text = fstrCuentaContable(vllngNumeroCuenta)
        End If
    End If
    
    vlstrCuentaCompleta = fstrCuentaCompleta(mskObject.Text)
    
    mskObject.Mask = ""
    mskObject.Text = vlstrCuentaCompleta
    mskObject.Mask = txtEstructuraem
    
    vllngNumeroCuenta = flngNumeroCuenta(mskObject.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
    
    If vllngNumeroCuenta <> 0 Then
        txtObject.Text = fstrDescripcionCuenta(mskObject.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
    Else
        mskObject.Mask = ""
        mskObject.Text = ""
        mskObject.Mask = txtEstructuraem
        txtObject.Text = ""
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsignaCuenta"))
    Unload Me
End Sub

Private Sub mskCuentaGasto_LostFocus()
Dim intRow As Integer
    Dim lngCuenta As Long
    intRow = fintLocalizaRow
    If intRow > -1 Then
        If cboEmpresaContable.ListIndex > -1 Then
        lngCuenta = flngNumeroCuenta(mskCuentaGasto.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
            If lngCuenta > 0 Then
                grdConceptoEmpresas.TextMatrix(intRow, 2) = lngCuenta
            Else
                grdConceptoEmpresas.TextMatrix(intRow, 2) = ""
            End If
        End If
    End If
End Sub

Private Sub sstab_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If sstab.Tab = 1 Then
        pCarga
        grdBusqueda.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstab_Click"))
End Sub

Private Sub txtClave_GotFocus()
    On Error GoTo NotificaError
    
    pLimpia
    pHabilita 0, 0, 1, 0, 0, 0, 0
    pSelTextBox txtClave

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_GotFocus"))
End Sub
Private Function fintLocalizaRow()
    Dim intIndex As Integer
    If cboEmpresaContable.ListIndex > -1 Then
        For intIndex = 0 To grdConceptoEmpresas.Rows - 1
            If CLng(grdConceptoEmpresas.TextMatrix(intIndex, 0)) = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) Then
                fintLocalizaRow = intIndex
                Exit Function
            End If
        Next
        grdConceptoEmpresas.AddItem ""
        grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 0) = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
        grdConceptoEmpresas.TextMatrix(grdConceptoEmpresas.Rows - 1, 1) = Trim(txtClave.Text)
        fintLocalizaRow = grdConceptoEmpresas.Rows - 1
        Exit Function
    Else
        fintLocalizaRow = -1
    End If
    fintLocalizaRow = -1
End Function
Private Sub pLimpia()
    On Error GoTo NotificaError
    
    lblnChange = True
    lblnConsulta = False
    
    txtClave.Text = frsRegresaRs("select isnull(max(intConsecutivo),0)+1 from PvConceptoCajaChica").Fields(0)
    txtDescripcion.Text = ""
    cboDepartamento.ListIndex = 0
    
    llngNumCuenta = 0
    
    chkActiva.Value = 1


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer, vlb4 As Integer, vlb5 As Integer, vlb6 As Integer, vlb7 As Integer)
    On Error GoTo NotificaError
    
    cmdTop.Enabled = vlb1 = 1
    cmdBack.Enabled = vlb2 = 1
    cmdLocate.Enabled = vlb3 = 1
    cmdNext.Enabled = vlb4 = 1
    cmdEnd.Enabled = vlb5 = 1
    cmdSave.Enabled = vlb6 = 1
    cmdDelete.Enabled = vlb7 = 1

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


Private Sub pMuestra(vllngxNumero As Long)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    
    
    Set rs = frsEjecuta_SP(CStr(vllngxNumero) & "|" & "-1", "SP_PVSELCONCEPTOCAJACHICA")
    If rs.RecordCount <> 0 Then
        lblnConsulta = True
    
        txtClave.Text = Str(rs!intConsecutivo)
        txtDescripcion.Text = Trim(rs!VCHDESCRIPCION)
        If IsNull(rs!intCveDepartamento) Then
            cboDepartamento.ListIndex = 0
        Else
            cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, rs!intCveDepartamento)
        End If
        pCargaConceptos
        lblnChange = False
      
        lblnChange = True
        
        'lblnChangeDescuento = False
      
        'lblnChangeDescuento = True
        
        chkActiva.Value = rs!bitEstado
    Else
        vblncuentagasto = False
        pLimpia
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraForma"))
End Sub


Private Sub txtDescripcion_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
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


