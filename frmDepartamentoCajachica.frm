VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDepartamentoCajachica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos con caja chica"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   4020
      Left            =   -45
      TabIndex        =   0
      Top             =   -375
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   7091
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmDepartamentoCajachica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Combo1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmDepartamentoCajachica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdDepartamentos"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   1590
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   8415
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Seleccione el departamento"
            Top             =   630
            Width           =   6350
         End
         Begin VB.ComboBox Cboempresacontable 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   6350
         End
         Begin MSMask.MaskEdBox mskCuenta 
            Height          =   315
            Left            =   1890
            TabIndex        =   13
            Top             =   1020
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label lblDescripcionCuenta 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3630
            TabIndex        =   16
            Top             =   1020
            Width           =   4605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta contable"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   690
            Width           =   1005
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa contable"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   6495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   6495
      End
      Begin VB.Frame Frame2 
         Height          =   720
         Left            =   3533
         TabIndex        =   1
         Top             =   2055
         Width           =   1635
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   555
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmDepartamentoCajachica.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   60
            Picture         =   "frmDepartamentoCajachica.frx":037A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   1050
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmDepartamentoCajachica.frx":04EC
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Borrar"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDepartamentos 
         Height          =   2310
         Left            =   -74910
         TabIndex        =   17
         Top             =   420
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   4075
         _Version        =   393216
         GridColor       =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa contable"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Empresa contable"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDepartamentoCajachica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const cintColCveDepto = 1
Const cintColDepartamento = 2
Const cintColCuentaContable = 3
Const cintColDescripcion = 4
Const cintColumnas = 5
Const cstrTitulo = "||Departamento|Cuenta|Descripción"


Dim rs As New ADODB.Recordset 'Varios usos
Dim llngNumCuenta  As Long 'Id. de la cuenta contable seleccionada
Dim lblnChange As Boolean 'Para saber cuando limpiar la descripcion de la cuenta
Dim lblnConsulta As Boolean 'Para saber cuando es una consulta o una alta
Dim lblnCargarDeptos As Boolean 'Para saber cuando se debe volver a cargar el combo de departamentos
Dim llngPersonaGraba As Long 'Persona que guarda o elimina datos
Dim vlblnPrimeraVez As Boolean
Dim vlstrsql As String
Dim vlstrestructuracuenta As String


Private Sub cboDepartamento_Click()
    On Error GoTo NotificaError


    If cboDepartamento.ListIndex <> -1 Then
    
        Set rs = frsEjecuta_SP(cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex), "SP_PVSELCAJACHICA")
        
        lblnChange = False
        
        mskCuenta.Mask = ""
        If rs.RecordCount <> 0 Then
            mskCuenta.Text = rs!Cuenta
            lblDescripcionCuenta.Caption = rs!DescripcionCuenta
            
            pHabilita 1, 0, 1
        
        Else
            mskCuenta.Text = ""
            lblDescripcionCuenta.Caption = ""
            
            pHabilita 1, 0, 0
        End If
        mskCuenta.Mask = vlstrestructuracuenta
        
        lblnChange = True
                
        lblnConsulta = rs.RecordCount = 1
    Else
        mskCuenta.Mask = ""
        mskCuenta.Text = ""
        mskCuenta.Mask = vlstrestructuracuenta
        
        lblnConsulta = False
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_Click"))
End Sub

Private Sub cboDepartamento_GotFocus()
    On Error GoTo NotificaError



    pHabilita 1, 0, 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_GotFocus"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError


    lblnConsulta = False
        
    cboDepartamento.ListIndex = -1
    cboDepartamento.Clear
    If lblnCargarDeptos Then
        lblnCargarDeptos = False
        pCargaDeptos
    End If
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Private Sub cboEmpresaContable_Click()
Dim rsestructura As New ADODB.Recordset
    pLimpia
    'vlstrSQL = "select vchestructuracuentacontable from cnParametro where tnyclaveempresa = " & Cboempresacontable.ItemData(Cboempresacontable.ListIndex)
    Set rsestructura = frsSelParametros("CN", cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex), "VCHESTRUCTURACUENTACONTABLE")
    If Not rsestructura.EOF Then
        vlstrestructuracuenta = rsestructura!Valor
    End If
    pCargaDeptos
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError


    If fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionDeptosCajaChica, "E") Then
    
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        
        If llngPersonaGraba <> 0 Then
        
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            
                '------------------------------------------------------------------
                ' Eliminar el registro del departamento como caja chica:
                '------------------------------------------------------------------
                vgstrParametrosSP = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
                frsEjecuta_SP vgstrParametrosSP, "SP_PVDELCAJACHICA"
                '------------------------------------------------------------------
                ' Registro de transacciones:
                '------------------------------------------------------------------
                pGuardarLogTransaccion Me.Name, EnmBorrar, llngPersonaGraba, "DEPARTAMENTO CAJA CHICA", CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
    
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"

            cboDepartamento.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError

    Dim intContador As Integer
    
    With grdDepartamentos
        .Cols = cintColumnas
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = cstrTitulo
        
        For intContador = 1 To .Cols - 1
            .TextMatrix(1, intContador) = ""
        Next intContador
    
        .ColWidth(0) = 100
        .ColWidth(cintColCveDepto) = 0
        .ColWidth(cintColDepartamento) = 2300
        .ColWidth(cintColCuentaContable) = 1500
        .ColWidth(cintColDescripcion) = 3700
    
        .ColAlignment(cintColDepartamento) = flexAlignLeftCenter
        .ColAlignment(cintColCuentaContable) = flexAlignLeftCenter
        .ColAlignment(cintColDescripcion) = flexAlignLeftCenter
    
        .ColAlignmentFixed(cintColDepartamento) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCuentaContable) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDescripcion) = flexAlignCenterCenter
        
        Set rs = frsEjecuta_SP("-1|" & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex), "SP_PVSELCAJACHICA")
        If rs.RecordCount <> 0 Then
        
            Do While Not rs.EOF
                .TextMatrix(.Rows - 1, cintColCveDepto) = rs!intCveDepartamento
                .TextMatrix(.Rows - 1, cintColDepartamento) = rs!Departamento
                .TextMatrix(.Rows - 1, cintColCuentaContable) = rs!Cuenta
                .TextMatrix(.Rows - 1, cintColDescripcion) = rs!DescripcionCuenta
                .Rows = .Rows + 1
                rs.MoveNext
            Loop
            .Rows = .Rows - 1
        
        End If
    
    End With

    sstab.Tab = 1


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError

    
    If fblnDatosValidos() Then
    
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        
        If llngPersonaGraba <> 0 Then
        
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
                '------------------------------------------------------------------
                ' Registro del departamento como caja chica:
                '------------------------------------------------------------------
                vgstrParametrosSP = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & CStr(llngNumCuenta)
                frsEjecuta_SP vgstrParametrosSP, IIf(Not lblnConsulta, "SP_PVINSDEPARTAMENTOCAJACHICA", "SP_PVUPDDEPARTAMENTOCAJACHICA")
                '------------------------------------------------------------------
                ' Registro de transacciones:
                '------------------------------------------------------------------
                pGuardarLogTransaccion Me.Name, IIf(Not lblnConsulta, EnmGrabar, EnmCambiar), llngPersonaGraba, "DEPARTAMENTO CAJA CHICA", CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
    
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"

            cboDepartamento.SetFocus
        End If
    
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError

    fblnDatosValidos = True
    
    If cboDepartamento.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbOKOnly + vbExclamation, "Mensaje"
        cboDepartamento.SetFocus
    End If
    If fblnDatosValidos And Trim(lblDescripcionCuenta.Caption) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        mskCuenta.SetFocus
    End If
    If fblnDatosValidos Then
        If Not fblnCuentaAfectable(mskCuenta.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)) Then
            fblnDatosValidos = False
            'La cuenta seleccionada no acepta movimientos.
            MsgBox SIHOMsg(375), vbOKOnly + vbExclamation, "Mensaje"
            mskCuenta.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionDeptosCajaChica, "E")
    End If

    

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function


Private Sub Form_Activate()
    On Error GoTo NotificaError


    If Trim(vgstrEstructuraCuentaContable) = "" Then
        'No se encuentra registrado el parámetro de estructura de la cuenta contable.
        MsgBox SIHOMsg(260), vbExclamation + vbOKOnly, "Mensaje"
        Unload Me
    End If
 
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    

    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            If Me.ActiveControl.Name = "mskCuenta" Then
                If Trim(mskCuenta.ClipText) = "" Then
                    llngNumCuenta = flngBusquedaCuentasContables(False, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                    If llngNumCuenta <> 0 Then
                        lblnChange = False
                         
                        mskCuenta.Text = fstrCuentaContable(llngNumCuenta)
                        lblnChange = True
                    End If
                Else
                    lblnChange = False
                    mskCuenta.Mask = ""
                    mskCuenta.Text = fstrCuentaCompleta(mskCuenta.Text)
                    mskCuenta.Mask = vlstrestructuracuenta
                    lblnChange = True
                    
                    llngNumCuenta = flngNumeroCuenta(mskCuenta.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                End If
                lblDescripcionCuenta.Caption = fstrDescripcionCuenta(mskCuenta.Text, cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
                SendKeys vbTab
            Else
                SendKeys vbTab
            End If
        End If
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    
    Me.Icon = frmMenuPrincipal.Icon
    
 

    mskCuenta.Mask = vgstrEstructuraCuentaContable

    lblnChange = True
    lblnConsulta = False
    pCargaEmpresas
    vlblnPrimeraVez = True
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub
Private Sub pCargaEmpresas()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    vgstrParametrosSP = -1
    
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_gnselempresascontable")
    If Not rs.EOF Then
        pLlenarCboRs cboEmpresaContable, rs, 1, 0
    End If
    rs.Close
    cboEmpresaContable.ListIndex = fintLocalizaCbo(cboEmpresaContable, CStr(vgintClaveEmpresaContable)) 'se posiciona en la empresa con el que se dio login
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 1932, 355), "C") Then
        cboEmpresaContable.Enabled = True
    Else
        cboEmpresaContable.Enabled = False
    End If
    

End Sub
Private Sub pCargaDeptos()
    On Error GoTo NotificaError

    
        vgstrParametrosSP = "-1|1|G|" & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_GnSelDepartamento")
        If rs.RecordCount <> 0 Then
            pLlenarCboRs cboDepartamento, rs, 0, 1
        End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDeptos"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError


    If sstab.Tab = 0 Then
        If cmdSave.Enabled Or lblnConsulta Then
            Cancel = 1
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                cboDepartamento.ListIndex = -1
                cboDepartamento.SetFocus
            End If
        End If
    Else
        Cancel = 1
        sstab.Tab = 0
        cboDepartamento.SetFocus
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub grdDepartamentos_DblClick()
    On Error GoTo NotificaError


    If Val(grdDepartamentos.TextMatrix(1, cintColCveDepto)) <> 0 Then
    
        cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, CStr(grdDepartamentos.TextMatrix(grdDepartamentos.Row, cintColCveDepto)))
        
        If cboDepartamento.ListIndex = -1 Then
            cboDepartamento.AddItem grdDepartamentos.TextMatrix(grdDepartamentos.Row, cintColDepartamento)
            cboDepartamento.ItemData(cboDepartamento.NewIndex) = Val(grdDepartamentos.TextMatrix(grdDepartamentos.Row, cintColCveDepto))
            cboDepartamento.ListIndex = cboDepartamento.NewIndex
        End If
                
        lblnChange = False
        
        mskCuenta.Mask = ""
        mskCuenta.Text = grdDepartamentos.TextMatrix(grdDepartamentos.Row, cintColCuentaContable)
        mskCuenta.Mask = vlstrestructuracuenta
        lblDescripcionCuenta.Caption = grdDepartamentos.TextMatrix(grdDepartamentos.Row, cintColDescripcion)
    
        lblnChange = True
        
        pHabilita 1, 0, 1
        
        sstab.Tab = 0
        
        cmdLocate.SetFocus
    
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdDepartamentos_DblClick"))
End Sub

Private Sub grdDepartamentos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        grdDepartamentos_DblClick
    End If
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdDepartamentos_KeyDown"))
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub mskCuenta_Change()
    On Error GoTo NotificaError

    
    If lblnChange Then
        lblDescripcionCuenta.Caption = ""
        llngNumCuenta = 0
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskCuenta_Change"))
End Sub

Private Sub mskCuenta_GotFocus()
    On Error GoTo NotificaError


    pHabilita 0, 1, 0
    pSelMkTexto mskCuenta

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskCuenta_GotFocus"))
End Sub

Private Sub pHabilita(intBuscar As Integer, intGrabar As Integer, intBorrar As Integer)
    On Error GoTo NotificaError


    cmdLocate.Enabled = intBuscar = 1
    cmdSave.Enabled = intGrabar = 1
    cmdDelete.Enabled = intBorrar = 1

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

