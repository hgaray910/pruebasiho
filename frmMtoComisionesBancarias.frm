VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMtoComisionesBancarias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones bancarias"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTObj 
      Height          =   4185
      Left            =   -45
      TabIndex        =   15
      Top             =   -450
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   7382
      _Version        =   393216
      TabHeight       =   661
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Mantenimiento"
      TabPicture(0)   =   "frmMtoComisionesBancarias.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Departamentos que lo utilizan"
      TabPicture(1)   =   "frmMtoComisionesBancarias.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdHBusqueda"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmMtoComisionesBancarias.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74820
         TabIndex        =   22
         Top             =   360
         Width           =   8295
         Begin VB.TextBox TxtDescripcioncuenta 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   2880
            TabIndex        =   28
            ToolTipText     =   "Nombre del proveedor"
            Top             =   3480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox TxtComision2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1995
            TabIndex        =   23
            ToolTipText     =   "Descripción de la comisón"
            Top             =   420
            Width           =   6135
         End
         Begin VB.CommandButton CmdRegresar 
            Caption         =   "Regresar"
            Height          =   495
            Left            =   7155
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Regresar a la pantalla principal"
            Top             =   3120
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton CmdBorrar 
            Height          =   495
            Left            =   6660
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Borrar cuenta contable"
            Top             =   3120
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin MSMask.MaskEdBox MskEdit 
            Height          =   315
            Left            =   480
            TabIndex        =   24
            ToolTipText     =   "Cuenta de ingresos"
            Top             =   3480
            Visible         =   0   'False
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCuentasEmpresa 
            DragIcon        =   "frmMtoComisionesBancarias.frx":0756
            Height          =   2055
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Cuentas contables por empresa"
            Top             =   960
            Width           =   8010
            _ExtentX        =   14129
            _ExtentY        =   3625
            _Version        =   393216
            ForeColor       =   0
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   -2147483643
            GridColor       =   12632256
            GridColorFixed  =   -2147483632
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            ScrollBars      =   2
            FormatString    =   "|tnyCvePiso|vchDescripcion"
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLineWidthBand=   1
            _Band(0).TextStyleBand=   0
         End
         Begin VB.Label lblTelProv 
            Caption         =   "Cuentas contables"
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Comisión"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   420
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2475
         Left            =   190
         TabIndex        =   10
         Top             =   600
         Width           =   8190
         Begin VB.TextBox txtComision 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1635
            MaxLength       =   6
            TabIndex        =   2
            ToolTipText     =   "Porcentaje"
            Top             =   1170
            Width           =   585
         End
         Begin VB.TextBox txtCveComision 
            Height          =   315
            Left            =   1635
            MaxLength       =   8
            TabIndex        =   0
            ToolTipText     =   "Clave"
            Top             =   450
            Width           =   1185
         End
         Begin VB.TextBox txtDescripcion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1635
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   "Descripción "
            Top             =   810
            Width           =   6135
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Comisión bancaria activa"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
            Enabled         =   0   'False
            Height          =   255
            Left            =   1635
            TabIndex        =   4
            ToolTipText     =   "Estado"
            Top             =   1950
            Width           =   2375
         End
         Begin VB.ComboBox cboIvas 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMtoComisionesBancarias.frx":0A60
            Left            =   1635
            List            =   "frmMtoComisionesBancarias.frx":0A62
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "IVA"
            Top             =   1545
            Width           =   2310
         End
         Begin VB.Label Label4 
            Caption         =   "%"
            Height          =   195
            Left            =   2280
            TabIndex        =   21
            Top             =   1230
            Width           =   225
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje "
            Height          =   195
            Left            =   210
            TabIndex        =   20
            Top             =   1230
            Width           =   810
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   210
            TabIndex        =   19
            Top             =   870
            Width           =   840
         End
         Begin VB.Label lblClave 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   510
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IVA"
            Height          =   195
            Left            =   210
            TabIndex        =   17
            Top             =   1605
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   210
            TabIndex        =   16
            Top             =   1965
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Height          =   720
         Left            =   1800
         TabIndex        =   13
         Top             =   3240
         Width           =   4850
         Begin VB.CommandButton cmdCuentascontables 
            Caption         =   "Cuentas contables"
            Height          =   480
            Left            =   3570
            TabIndex        =   14
            ToolTipText     =   "Configurar cuentas contables por empresa"
            Top             =   165
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Enabled         =   0   'False
            Height          =   480
            Left            =   3060
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":0A64
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Borrar el registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Enabled         =   0   'False
            Height          =   480
            Left            =   2550
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":0C06
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Guardar el registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2055
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":0F48
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Ultimo registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":10BA
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Siguiente registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1065
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":122C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Búsqueda"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   570
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":139E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Anterior registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   75
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoComisionesBancarias.frx":1510
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Primer registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
         CausesValidation=   0   'False
         DragIcon        =   "frmMtoComisionesBancarias.frx":1682
         Height          =   3450
         Left            =   -74895
         TabIndex        =   31
         ToolTipText     =   "Doble click para seleccionar una comisión bancaria"
         Top             =   525
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6085
         _Version        =   393216
         ForeColor       =   0
         Rows            =   16
         Cols            =   5
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         MergeCells      =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
      End
   End
End
Attribute VB_Name = "frmMtoComisionesBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Permite forzar la declaración de las variables
Dim vgblnNuevoRegistro As Boolean
Public vllngNumeroOpcion As Long
Dim rsComisionesBancarias As New ADODB.Recordset
Dim vlstrsql As String

Private Sub cboCuentaIVAxPagar_Change()

End Sub

Private Sub cmdBorrar_Click()
'-------------------------------------------------------------------------------------------
'Borra la cuenta contable de la empresa
'-------------------------------------------------------------------------------------------

 On Error GoTo NotificaError
    Dim vlstrMensaje As String
    Dim vlintResultado As Integer
    Dim vlintNumReg As Integer
    Dim llngContador As Long
    Dim contador As Integer


    If (grdCuentasEmpresa.Cols - 1) > 0 Then
        For llngContador = 1 To grdCuentasEmpresa.Rows - 1
            If grdCuentasEmpresa.TextMatrix(llngContador, 0) = "*" Then
                contador = contador + 1
            End If
        Next
        vlintResultado = MsgBox(SIHOMsg("6"), (vbYesNo + vbQuestion), "Mensaje")
        If vlintResultado = vbYes Then
            For llngContador = 1 To grdCuentasEmpresa.Rows - 1
                If grdCuentasEmpresa.TextMatrix(llngContador, 0) = "*" Then
                    grdCuentasEmpresa.TextMatrix(llngContador, 3) = ""
                    grdCuentasEmpresa.TextMatrix(llngContador, 4) = ""
                    grdCuentasEmpresa.TextMatrix(llngContador, 5) = ""
                    grdCuentasEmpresa.TextMatrix(llngContador, 0) = ""
                End If
            Next
        Else
            CmdRegresar.SetFocus
        End If
    End If

Exit Sub
NotificaError:
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBorrar_Click"))
End Sub

Private Sub cmdCuentasContables_Click()
On Error GoTo NotificaError
    
    If Me.txtDescripcion <> "" Then
        TxtComision2.Text = txtDescripcion.Text
        SSTObj.Tab = 2
        CmdBorrar.Enabled = False
        Me.grdCuentasEmpresa.SetFocus
        
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCuentascontables_Click"))
End Sub

Private Sub cmdDelete_Click()
    Dim vlstrSentencia As String
    Dim rsComision As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, 3060, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3060, "C", True) _
        Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "C", True) Then
              
        vlstrSentencia = "Select * from pvformapagotipocargocomision where smicvecomision = " & txtCveComision.Text
        Set rsComision = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)

        If rsComision.RecordCount > 0 Then
            'No se puede eliminar la información, ya ha sido utilizada.
            MsgBox SIHOMsg(771), vbOKOnly + vbCritical, "Mensaje"
            pNuevoRegistro True
            pEnfocaTextBox txtCveComision
        Else
            If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                '-----------------------'
                '   Persona que graba   '
                '-----------------------'
                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                If vllngPersonaGraba <> 0 Then
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                    
                    vlstrSentencia = "Delete from pvcomisionbancariaempresa where smicvecomision = " & txtCveComision.Text
                    pEjecutaSentencia vlstrSentencia
                
                    rsComisionesBancarias.Delete
                    rsComisionesBancarias.Requery
                
                    Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "COMISIONBANCARIA", txtCveComision.Text)
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    
                    pNuevoRegistro True
                    pEnfocaTextBox txtCveComision
                End If
            End If
        End If
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub

Private Sub cmdGrabarRegistro_Click()
Dim vlintContador As Integer
Dim vlintclavecomision As Integer
Dim vlintSeqFil As Integer
Dim vllngPersonaGraba As Long
    
    'Checar el pemiso que le mandan
    ' 603   CC
    ' 2367  SI (CC)
    ' 348   PV          3060
    ' 1120  SI (PV)     3061
'    If fblnRevisaPermiso(vglngNumeroLogin, 603, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 603, "C", True) _
'        Or fblnRevisaPermiso(vglngNumeroLogin, 2367, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2367, "C", True) _
'        Or fblnRevisaPermiso(vglngNumeroLogin, 3060, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3060, "C", True) _
'        Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "C", True) Then
    If fblnRevisaPermiso(vglngNumeroLogin, 3060, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3060, "C", True) _
        Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "C", True) Then
                
        '**********************************************************
        ' Procedimiento para grabar una alta o modificación
        '**********************************************************
       ' If cboCuentas.ListIndex = 0 Then
        'End If

        If RTrim(txtDescripcion.Text) = "" Then
            MsgBox SIHOMsg(2) + Chr(13) + txtDescripcion.ToolTipText, vbExclamation, "Mensaje"
            txtDescripcion.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtComision.Text) Then
            'Porcentaje incorrecto.
            MsgBox SIHOMsg(400), vbExclamation, "Mensaje"
            txtComision.SetFocus
            Exit Sub
        End If
        
        If CDbl(txtComision.Text) > 100 Or CDbl(txtComision.Text) <= 0 Then
            'Porcentaje incorrecto.
            MsgBox SIHOMsg(400), vbExclamation, "Mensaje"
            txtComision.SetFocus
            Exit Sub
        Else
            '-----------------------'
            '   Persona que graba   '
            '-----------------------'
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba <> 0 Then
                '--------------------------------------
                ' Grabar la comision
                '--------------------------------------
                EntornoSIHO.ConeccionSIHO.BeginTrans
                
                If Not cmdBuscar.Enabled Then
                    rsComisionesBancarias.AddNew
                    vlintclavecomision = 0
                Else
                    vlintclavecomision = CDbl(txtCveComision.Text)
                End If
                
                rsComisionesBancarias!chrDescripcion = Trim(txtDescripcion.Text)
                rsComisionesBancarias!smyiva = cboIvas.ItemData(cboIvas.ListIndex)
                rsComisionesBancarias!mnycomision = CDbl(txtComision.Text)
                'rsComisiones!bitAsignada = chkIncluidaDefault.Value
                rsComisionesBancarias!bitactivo = chkActivo.Value
                rsComisionesBancarias.Update
                rsComisionesBancarias.Requery
                
                If vlintclavecomision = 0 Then
                    txtCveComision.Text = CStr(flngObtieneIdentity("sec_PvcomisionBancaria", 0))
                    vlintclavecomision = CDbl(txtCveComision.Text)
                End If
                
                 '--Guardar cuentas contables
                vlstrsql = " DELETE FROM PVCOMISIONBANCARIAEMPRESA WHERE smicvecomision = " & vlintclavecomision
                pEjecutaSentencia vlstrsql
                For vlintSeqFil = 1 To grdCuentasEmpresa.Rows - 1
                    If grdCuentasEmpresa.TextMatrix(vlintSeqFil, 3) <> "" Then
                        With grdCuentasEmpresa
                            vgstrParametrosSP = vlintclavecomision & "|" & .TextMatrix(vlintSeqFil, 1) & "|" & .TextMatrix(vlintSeqFil, 3)
                            frsEjecuta_SP vgstrParametrosSP, "Sp_PVINSCOMISIONBANCARIAEMPR"
                        End With
                    End If
                Next vlintSeqFil
                
                If Not cmdBuscar.Enabled Then
                    txtCveComision.Text = flngObtieneIdentity("SEC_PVCOMISIONBANCARIA", rsComisionesBancarias!smicvecomision)
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "COMISIONBANCARIA", txtCveComision.Text)
                Else
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "COMISIONBANCARIA", txtCveComision.Text)
                End If
                
                EntornoSIHO.ConeccionSIHO.CommitTrans
                
                Call pNuevoRegistro(True)
                txtCveComision.SetFocus
            End If
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub


Private Sub cmdRegresar_Click()
On Error GoTo NotificaError 'Manejo del error

    SSTObj.Tab = 0
    cmdGrabarRegistro.Enabled = True
    cmdGrabarRegistro.SetFocus
    
    Exit Sub
NotificaError:
     Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":CmdRegresar_Click"))
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            If Not (SSTObj.Tab = 2 And MskEdit.Visible) Then
                KeyCode = 7
                Unload Me
            End If
        Case vbKeyReturn
             If Not SSTObj.Tab = 2 Then
                SendKeys vbTab
            End If
    End Select
    
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim rscmdCuentasIvaxPagar As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vlstrSentencia = "select * from PvComisionBancaria"
    Set rsComisionesBancarias = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    Me.Icon = frmMenuPrincipal.Icon
    pCargaIvas
    cboIvas.ListIndex = 0
    pNuevoRegistro True
    vgblnNuevoRegistro = True
    cmdCuentascontables.Enabled = False
    
    SSTObj.Tab = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If SSTObj.Tab <> 0 Then
       
        SSTObj.Tab = 0
        If txtDescripcion.Enabled Then
            txtDescripcion.SetFocus
        Else
            txtCveComision.SetFocus
        End If
        Cancel = True
    
    Else
        If Not vgblnNuevoRegistro Then
            If MsgBox(SIHOMsg(9), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                Call pNuevoRegistro(True)
                pEnfocaTextBox txtCveComision
            End If
            Cancel = True
        Else
            rsComisionesBancarias.Close
        End If
    End If
End Sub


Private Sub pNuevoRegistro(vlblnNuevo As Boolean)
    Dim rscuentasempresas As New ADODB.Recordset
    If SSTObj.Tab = 1 Then Exit Sub
    If vlblnNuevo Then txtCveComision.Text = fintSigNumRs(rsComisionesBancarias, 0)
    txtDescripcion.Text = ""
    txtComision.Text = "0.0"
    txtComision.Enabled = False
    txtDescripcion.Enabled = False
    cboIvas.Enabled = False
    chkActivo.Enabled = False
'    chkIncluidaDefault.Enabled = False
    
    If cboIvas.ListCount > 0 Then
        cboIvas.ListIndex = 0
    Else
       Call MsgBox(SIHOMsg("12") & Chr(13) & "Dato:" & cboIvas.ToolTipText, vbExclamation, "Mensaje")
       Unload Me
       Exit Sub
    End If
    
    
    If rsComisionesBancarias.RecordCount > 0 Then
        pHabilitaBotonModifica (True)
        cmdGrabarRegistro.Enabled = False
        cmdDelete.Enabled = False
    Else
        pHabilitaBotonModifica (False)
    End If
    chkActivo.Value = 1
'    chkIncluidaDefault.Value = 0
    vgblnNuevoRegistro = True
    
         'Se inica y limpia el grid
    Me.TxtComision2 = txtComision.Text
    
    Call pIniciaMshFGrid(grdCuentasEmpresa)
    Call pLimpiaMshFGrid(grdCuentasEmpresa)

    pConfFGrid grdCuentasEmpresa, "|Empresa||Cuenta contable|Descripción"
    vgstrParametrosSP = IIf(Me.txtCveComision.Text = "", -1, txtCveComision.Text) & "|" & 10
    Set rscuentasempresas = frsEjecuta_SP(vgstrParametrosSP, "sp_gnselcuentasempresa")
    If rscuentasempresas.RecordCount > 0 Then
         Call pLlenarMshFGrdRs(grdCuentasEmpresa, rscuentasempresas)
          pConfFGrid grdCuentasEmpresa, "||Empresa||Cuenta contable|Descripción"
    End If
    cmdCuentascontables.Enabled = False
    
End Sub
Public Sub pConfFGrid(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
    'Configura el MSHFlexGrid

    ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas
    

     With ObjGrid
       .ColWidth(0) = 150
       .ColWidth(1) = 0
       .ColWidth(2) = 3000
       .ColWidth(3) = 0
       .ColWidth(4) = 1700
       .ColWidth(5) = 4500
       .ColWidth(6) = 0
       .ScrollBars = flexScrollBarHorizontal
    End With


End Sub
Private Sub pLlenaGrid()
    Dim vlstrSentencia As String
    Dim PvComisionBancaria As New ADODB.Recordset
    Dim vlintContador As Integer
    grdHBusqueda.Clear
    

    vlstrSentencia = "SELECT PvComisionBancaria.smiCveComision, " & _
    "PvComisionBancaria.chrDescripcion, " & _
    "PvComisionBancaria.smyIva, " & _
    "PvComisionBancaria.bitActivo " & _
    "FROM PvComisionBancaria "
    Set PvComisionBancaria = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    If PvComisionBancaria.RecordCount > 0 Then
        Call pLlenarMshFGrdRs(grdHBusqueda, PvComisionBancaria)
        pConfiguraGrid
        With grdHBusqueda
            For vlintContador = 1 To .Rows - 1
                .TextMatrix(vlintContador, 4) = IIf(.TextMatrix(vlintContador, 4) = 1, "Activo", "Inactivo")
                .TextMatrix(vlintContador, 3) = FormatPercent(Format(.TextMatrix(vlintContador, 3), "####0.#0") / 100, 2)
            Next
        End With
    Else
        SSTObj.Tab = 0
        cmdBuscar.SetFocus
    End If
    PvComisionBancaria.Close
End Sub

Private Sub pConfiguraGrid()
    With grdHBusqueda
        .FormatString = "|Clave|Descripción|IVA|Estado"
        .ColWidth(0) = 150 'Fix
        .ColWidth(1) = 700 'Clave
        .ColWidth(2) = 4600 'Descripcion
        .ColWidth(3) = 950  'IvA
        .ColWidth(4) = 1000  'Estatus
        .ColAlignment(2) = flexAlignLeftBottom
        .ColAlignmentFixed(2) = flexAlignLeftBottom
        .ColAlignment(3) = flexAlignRightBottom
        .ColAlignmentFixed(3) = flexAlignCenterBottom
        .ScrollBars = flexScrollBarVertical
    End With
End Sub


Private Sub grdCuentasEmpresa_Click()
On Error GoTo NotificaError
Dim vlblnpermiso As Boolean
    
    
    If grdCuentasEmpresa.Rows > 0 And grdCuentasEmpresa.Col = 4 Then
        If vgintClaveEmpresaContable <> grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 1) Then
            If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 3061, 3060), "C", True) Then
                vlblnpermiso = True
            Else
                vlblnpermiso = False
            End If
        Else
            vlblnpermiso = True
        End If
        If vlblnpermiso Then
            TxtDescripcioncuenta.Text = ""
            MskEdit.Visible = True
            MskEdit.Mask = ""
            MskEdit.Move grdCuentasEmpresa.Left + grdCuentasEmpresa.CellLeft, grdCuentasEmpresa.Top + grdCuentasEmpresa.CellTop, grdCuentasEmpresa.CellWidth - 8, grdCuentasEmpresa.CellHeight - 8
            MskEdit.Text = ""
            MskEdit.Mask = grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 6)
            pEnfocaMkTexto MskEdit
        End If
        
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCuentasEmpresa_Click"))
End Sub

Private Sub grdCuentasEmpresa_DblClick()
On Error GoTo NotificaError

  If grdCuentasEmpresa.Rows > 1 And grdCuentasEmpresa.Row > 0 Then
      
      If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 3061, 3060), "C", True) Then
          If grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 0) = "*" Then
              grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 0) = ""
          ElseIf grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 3) <> "" Then
              grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 0) = "*"
          End If
      End If

      pHabilitaBorrar
        
  End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCuentasEmpresa_DblClick"))

End Sub
Private Sub pHabilitaBorrar()
    On Error GoTo NotificaError
    
    Dim X As Long
    Dim vlblnTermina As Boolean
    
    CmdBorrar.Enabled = False
    
    X = 1
    vlblnTermina = False
    Do While X <= grdCuentasEmpresa.Rows - 1 And Not vlblnTermina
        If Trim(grdCuentasEmpresa.TextMatrix(X, 0)) = "*" Then
            vlblnTermina = True
        End If
        X = X + 1
    Loop
    
    If vlblnTermina Then
        CmdBorrar.Enabled = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaBorrar"))
End Sub
Private Sub grdCuentasEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn And grdCuentasEmpresa.Col = 4 Then
        grdCuentasEmpresa_Click
    End If
        
Exit Sub
NotificaError:
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCuentasEmpresa_KeyDown"))
End Sub

Private Sub grdCuentasEmpresa_KeyPress(KeyAscii As Integer)
Dim vlblnpermiso As Boolean
Dim vlstrcaracter As String
Dim vlintContador As Integer

    
    If grdCuentasEmpresa.Rows > 0 And grdCuentasEmpresa.Col = 4 Then
        If fblnVerificaNumerico(KeyAscii) Then
            If vgintClaveEmpresaContable <> grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 1) Then
                If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 3061, 3060), "C", True) Then
                    vlblnpermiso = True
                Else
                    vlblnpermiso = False
                End If
            Else
                vlblnpermiso = True
            End If
            If vlblnpermiso Then
                TxtDescripcioncuenta.Text = ""
                MskEdit.Move grdCuentasEmpresa.Left + grdCuentasEmpresa.CellLeft, grdCuentasEmpresa.Top + grdCuentasEmpresa.CellTop, grdCuentasEmpresa.CellWidth - 8, grdCuentasEmpresa.CellHeight - 8
                MskEdit.Visible = True
                MskEdit.Mask = ""
                
                vlstrcaracter = Chr(KeyAscii)
                If Trim(vlstrcaracter) <> "" Then
                MskEdit.Text = vlstrcaracter
                vlintContador = Len(vlstrcaracter) + 1
                Do While vlintContador <= Len(vgstrEstructuraCuentaContable)
                    MskEdit.Text = MskEdit.Text + IIf(Mid(grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 6), vlintContador, 1) = "#", " ", ".")
                    vlintContador = vlintContador + 1
                Loop
                MskEdit.Mask = grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 6)
                End If
                
                vlintContador = 1
                MskEdit.SelStart = 0
                Do While vlintContador <= Len(MskEdit.Text)
                    If Mid(MskEdit.Text, vlintContador, 1) <> " " And Mid(MskEdit.Text, vlintContador, 1) <> "." Then
                        MskEdit.SelStart = MskEdit.SelStart + 1
                    Else
                        If vlintContador <> 1 Then
                            If Mid(MskEdit.Text, vlintContador, 1) = "." And Mid(MskEdit.Text, vlintContador - 1, 1) <> " " Then
                                MskEdit.SelStart = MskEdit.SelStart + 1
                            End If
                        End If
                    End If
                    vlintContador = vlintContador + 1
                Loop
                MskEdit.SetFocus
                  
            End If
        End If
    Else
        Me.CmdRegresar.SetFocus
    End If


End Sub

Private Sub grdHBusqueda_DblClick()
    If fintLocalizaPkRs(rsComisionesBancarias, 0, grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)) > 0 Then
        pModificaRegistro
        Call pEnfocaTextBox(txtDescripcion)
        SSTObj.Tab = 0
       txtCveComision_KeyDown 13, 0
    Else
        Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
        Call pEnfocaTextBox(txtCveComision)
    End If
End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdHBusqueda_DblClick
    End If
End Sub

Private Sub MskEdit_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = 13 Then
        pAsignaCuenta MskEdit, TxtDescripcioncuenta, grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 1)
    Else
        If KeyAscii = 27 Then
            grdCuentasEmpresa.SetFocus
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":MskEdit_KeyPress"))
End Sub

Private Sub mskEdit_LostFocus()
On Error GoTo NotificaError
        
        MskEdit.Mask = ""
        MskEdit.Text = ""
        MskEdit.Visible = False
      
        
        Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":MskEdit_LostFocus"))
End Sub
Private Sub pAsignaCuenta(mskobject As MaskEdBox, txtObject As TextBox, intEmpresa As Integer)
    
    Dim vllngNumeroCuenta As Long
    Dim rsCuenta As New ADODB.Recordset
    Dim strSentencia As String

On Error GoTo NotificaError
    
    '|  Si no se especificó la cuenta, mostrará la pantalla de búsqueda
    If Trim(mskobject.ClipText) = "" Then
        'vllngNumeroCuenta = flngBusquedaCuentasContables(False, " ", intEmpresa)
        vllngNumeroCuenta = flngBusquedaCuentasContables(False, intEmpresa)
        If vllngNumeroCuenta <> 0 Then mskobject.Text = fstrCuentaContable(vllngNumeroCuenta)
    Else
        vllngNumeroCuenta = flngNumeroCuenta(mskobject.Text, intEmpresa)
    End If
   
    '|  Si la cuenta si existe
    If vllngNumeroCuenta <> 0 Then
        '|  Valida que la cuenta seleccionada sea de tipo "Pasivo"
        strSentencia = " Select intNumeroCuenta " & _
                       "      , RTRIM(vchCuentaContable) " & _
                       "      , vchclasificacionTipo " & _
                       "      , vchtipo " & _
                       "  From CnCuenta " & _
                       "  Where vchCuentaContable = '" & mskobject.Text & "' " & _
                       "  AND bitestatusmovimientos = 1 " & _
                       "  AND TNYCLAVEEMPRESA = " & intEmpresa
        Set rsCuenta = frsRegresaRs(strSentencia)
        If rsCuenta.RecordCount > 0 Then
            If rsCuenta!vchClasificacionTipo <> "Gasto" And rsCuenta!vchClasificacionTipo <> "Costo" Then   'Or rsCuenta!vchTipo <> "Resultados"
                '|  Seleccione otra cuenta contable!
                MsgBox SIHOMsg(202) & Chr(13) & "de tipo Gasto o Costo ", vbExclamation, "Mensaje"
                mskobject.Mask = ""
                mskobject.Text = ""
                txtObject.Text = ""
                mskobject.Mask = grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 6)
                pEnfocaMkTexto mskobject
                rsCuenta.Close
                Exit Sub
            End If
            txtObject.Text = fstrDescripcionCuenta(mskobject.Text, intEmpresa)
            grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 3) = vllngNumeroCuenta
            grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 4) = mskobject.Text
            grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 5) = txtObject.Text
            pEnfocaMkTexto mskobject
        Else
            MsgBox SIHOMsg(375), vbOKOnly + vbExclamation, "Mensaje"
            mskobject.Mask = ""
            mskobject.Text = ""
            mskobject.Mask = grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 6)
            pEnfocaMkTexto mskobject
        End If

    Else
        '|  ¡La cuenta no existe!
        MsgBox SIHOMsg(67), vbCritical, "Mensaje"
        mskobject.Mask = ""
        mskobject.Text = ""
        mskobject.Mask = grdCuentasEmpresa.TextMatrix(grdCuentasEmpresa.Row, 6)
        txtObject.Text = ""
        pEnfocaMkTexto mskobject
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsignaCuenta"))
    Unload Me
End Sub

Private Sub SSTObj_Click(PreviousTab As Integer)
    If SSTObj.Tab = 1 Then
        pLlenaGrid
        grdHBusqueda.Enabled = True
        grdHBusqueda.SetFocus
    End If
End Sub



Private Sub txtComision_GotFocus()
  pSelTextBox txtComision
End Sub

Private Sub txtCveComision_GotFocus()
    
    pSelTextBox txtCveComision

End Sub

Private Sub txtCveComision_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
'modificar uno que ya existe
'-------------------------------------------------------------------------------------------
    Dim vlintNumero As Integer
    Select Case KeyCode
        Case vbKeyReturn
                If fintSigNumRs(rsComisionesBancarias, 0) = CLng(txtCveComision.Text) Then
                    txtDescripcion.Enabled = True
                    txtComision.Enabled = True
                    cboIvas.Enabled = True
                    chkActivo.Enabled = True
'                    chkIncluidaDefault.Enabled = True
                    vgblnNuevoRegistro = False
                    
                    chkActivo.Value = 1
                    chkActivo.Enabled = False
                    
                    pHabilitaBotonModifica False
                    cmdGrabarRegistro.Enabled = cboIvas.ListCount > 0
                    SSTObj.TabEnabled(1) = False
                Else
                    If fintLocalizaPkRs(rsComisionesBancarias, 0, txtCveComision.Text) > 0 Then
                        pModificaRegistro
                        txtDescripcion.Enabled = True
                        txtComision.Enabled = True
                        cboIvas.Enabled = True
                        chkActivo.Enabled = True
                        pHabilitaBotonModifica (True)
                        chkActivo.Enabled = True
'                        chkIncluidaDefault.Enabled = True
                    Else
                        rsComisionesBancarias.MoveLast
                        Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
                        Call pEnfocaTextBox(txtCveComision)
                        txtCveComision_GotFocus
                    End If
                End If
    End Select
End Sub

Private Sub pModificaRegistro()
    Dim rscuentasempresas As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlintContador As Integer
    '-------------------------------------------------------------------------------------------
    ' Permite realizar la modificación de la descripción de un registro
    '-------------------------------------------------------------------------------------------
    vgblnNuevoRegistro = False
    
    ' ---------------------------------------------
'    txtDescripcion.Enabled = False
'    txtComision.Enabled = False
'    cboIvas.Enabled = False 'cboIvas.ListCount > 0
'
'    chkActivo.Enabled = False
'    chkIncluidaDefault.Enabled = False
    ' ---------------------------------------------
    
    '---------------------------------------
    ' Carga las comisiones
    '---------------------------------------
    txtCveComision.Text = rsComisionesBancarias!smicvecomision
    txtDescripcion.Text = rsComisionesBancarias!chrDescripcion
    txtComision.Text = IIf(IsNull(rsComisionesBancarias!mnycomision), 0, rsComisionesBancarias!mnycomision)
    cboIvas.ListIndex = fintLocalizaCbo(cboIvas, rsComisionesBancarias!smyiva)
'    chkIncluidaDefault.Value = IIf(rsComisionesBancarias!bitAsignada Or rsComisionesBancarias!bitAsignada = 1, 1, 0)
    chkActivo.Value = IIf(rsComisionesBancarias!bitactivo Or rsComisionesBancarias!bitactivo = 1, 1, 0)
    SSTObj.TabEnabled(1) = True
    
    'Se inica y limpia el grid
    Me.TxtComision2.Text = txtDescripcion.Text
    
    Call pIniciaMshFGrid(grdCuentasEmpresa)
    Call pLimpiaMshFGrid(grdCuentasEmpresa)

    pConfFGrid grdCuentasEmpresa, "|Empresa||Cuenta contable|Descripción"
    vgstrParametrosSP = IIf(Me.txtCveComision.Text = "", -1, txtCveComision.Text) & "|" & 10
    Set rscuentasempresas = frsEjecuta_SP(vgstrParametrosSP, "sp_gnselcuentasempresa")
    If rscuentasempresas.RecordCount > 0 Then
         Call pLlenarMshFGrdRs(grdCuentasEmpresa, rscuentasempresas)
          pConfFGrid grdCuentasEmpresa, "||Empresa||Cuenta contable|Descripción"
    End If
    cmdCuentascontables.Enabled = True
    
End Sub

Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    cmdPrimerRegistro.Enabled = vlblnHabilita
    cmdAnteriorRegistro.Enabled = vlblnHabilita
    cmdBuscar.Enabled = vlblnHabilita
    SSTObj.TabEnabled(1) = vlblnHabilita
    cmdSiguienteRegistro.Enabled = vlblnHabilita
    cmdUltimoRegistro.Enabled = vlblnHabilita
    cmdDelete.Enabled = vlblnHabilita
   
    cmdGrabarRegistro.Enabled = vlblnHabilita And cboIvas.ListCount > 0
  
Exit Sub
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaBotonModifica"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtCveComision_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
  pSelTextBox txtDescripcion
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtDescripcion.Text <> "" Then cmdCuentascontables.Enabled = True
    End Select
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAnteriorRegistro_Click()
    Call pPosicionaRegRs(rsComisionesBancarias, "A")
    pModificaRegistro
End Sub

Private Sub cmdPrimerRegistro_Click()
    Call pPosicionaRegRs(rsComisionesBancarias, "I")
    pModificaRegistro
End Sub

Private Sub cmdSiguienteRegistro_Click()
    Call pPosicionaRegRs(rsComisionesBancarias, "S")
    pModificaRegistro
End Sub

Private Sub cmdUltimoRegistro_Click()
    Call pPosicionaRegRs(rsComisionesBancarias, "U")
    pModificaRegistro
End Sub

Private Sub pCargaIvas()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "select relPorcentaje, vchDescripcion || ' (' || ltrim(rtrim(cast(relPorcentaje as char(5)))) || '%)' as  vchDescripcion from CnImpuesto where bitActivo=1"
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboIvas, rs, 0, 1
    End If
    rs.Close
    
End Sub


Private Sub chkActivo_Click()
    If vgblnNuevoRegistro Then
        chkActivo.Value = 1
    End If
End Sub

Private Sub cmdBuscar_Click()
    SSTObj.Tab = 1
End Sub


Private Sub txtComision_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtComision, KeyAscii, 2) Then
       KeyAscii = 7
    End If
End Sub

Private Sub txtDescripcion_LostFocus()
    If txtDescripcion.Text <> "" Then
        cmdCuentascontables.Enabled = True
    End If
End Sub
