VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantoClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de clientes"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabClientes 
      Height          =   8910
      Left            =   -135
      TabIndex        =   25
      Top             =   -555
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   15716
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoClientes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoClientes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   3982
         TabIndex        =   33
         Top             =   7920
         Width           =   1590
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   1035
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoClientes.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Borrar"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   540
            Picture         =   "frmMantoClientes.frx":01DA
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Grabar"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   45
            Picture         =   "frmMantoClientes.frx":034C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Búsqueda"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7230
         Left            =   240
         TabIndex        =   26
         Top             =   615
         Width           =   9345
         Begin VB.TextBox txtPorcentaje 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   10
            ToolTipText     =   "Porcentaje de crédito disponible para envío de advertencia al ingreso"
            Top             =   3180
            Width           =   795
         End
         Begin VB.Frame Frame3 
            Caption         =   "Información predeterminada a mostrar para CFD/CFDi generados a crédito"
            Height          =   1095
            Left            =   2160
            TabIndex        =   42
            Top             =   3600
            Width           =   6975
            Begin VB.TextBox txtDescPago 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Left            =   3000
               MaxLength       =   20
               TabIndex        =   12
               ToolTipText     =   "Número de cuenta, tarjeta o referencia del método de pago predeterminado"
               Top             =   645
               Width           =   3855
            End
            Begin VB.ComboBox cboFormaPago 
               Height          =   315
               ItemData        =   "frmMantoClientes.frx":04BE
               Left            =   3000
               List            =   "frmMantoClientes.frx":04C0
               Style           =   2  'Dropdown List
               TabIndex        =   11
               ToolTipText     =   "Selección del método de pago predeterminado para CFDi"
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Número de cuenta, tarjeta o referencia"
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   705
               Width           =   2895
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Método de pago del SAT para CFDi"
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   300
               Width           =   2895
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   1530
            Left            =   2160
            TabIndex        =   40
            Top             =   4800
            Width           =   7140
            Begin VB.TextBox txtDiasEntrega 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6405
               MaxLength       =   3
               TabIndex        =   15
               ToolTipText     =   "Número de días para que venza el crédito"
               Top             =   600
               Width           =   525
            End
            Begin VB.TextBox txtDiasVencimiento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6405
               MaxLength       =   3
               TabIndex        =   17
               ToolTipText     =   "Número de días para que venza el crédito"
               Top             =   1170
               Width           =   525
            End
            Begin VB.OptionButton optTipoVencimiento 
               Caption         =   "A partir de una fecha automática de pago"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   16
               Top             =   990
               Width           =   3330
            End
            Begin VB.OptionButton optTipoVencimiento 
               Caption         =   "A partir de la fecha de entrega del documento (factura, ticket, etc.)"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   14
               Top             =   420
               Width           =   5160
            End
            Begin VB.OptionButton optTipoVencimiento 
               Caption         =   "A partir de la fecha del crédito"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   13
               Top             =   120
               Width           =   3330
            End
            Begin VB.Label lblDiasEntrega 
               AutoSize        =   -1  'True
               Caption         =   "Número de días posteriores a la fecha de entrega para considerar vencido un crédito"
               Height          =   195
               Left            =   255
               TabIndex        =   41
               Top             =   660
               Width           =   6015
            End
            Begin VB.Label lblDias 
               AutoSize        =   -1  'True
               Caption         =   "Número de días a partir de la fecha del crédito para calcular la fecha de pago"
               Height          =   195
               Left            =   285
               TabIndex        =   23
               Top             =   1230
               Width           =   5490
            End
         End
         Begin MSMask.MaskEdBox txtFechaAsignacion 
            Height          =   315
            Left            =   2160
            TabIndex        =   8
            Top             =   2775
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar ..."
            Height          =   300
            Left            =   2160
            TabIndex        =   38
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Frame fraBusqueda 
            BorderStyle     =   0  'None
            Height          =   1875
            Left            =   3885
            TabIndex        =   35
            Top             =   120
            Width           =   5355
            Begin VB.TextBox txtIniciales 
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   37
               ToolTipText     =   "Iniciales del nombre del cliente"
               Top             =   150
               Width           =   5235
            End
            Begin VB.ListBox lstBusqueda 
               Height          =   1230
               Left            =   0
               TabIndex        =   36
               Top             =   510
               Width           =   5235
            End
         End
         Begin VB.TextBox txtNombreCliente 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2055
            Width           =   6960
         End
         Begin VB.TextBox txtDescripcionCuenta 
            Height          =   315
            Left            =   4335
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2415
            Width           =   4785
         End
         Begin VB.ComboBox cboDepartamentos 
            Height          =   315
            ItemData        =   "frmMantoClientes.frx":04C2
            Left            =   2160
            List            =   "frmMantoClientes.frx":04C4
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Selección del departamento al que pertenece el cliente"
            Top             =   6390
            Width           =   4740
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Activo"
            Height          =   210
            Left            =   2160
            TabIndex        =   19
            ToolTipText     =   "Cliente activo"
            Top             =   6780
            Width           =   810
         End
         Begin VB.TextBox txtLimiteCredito 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5160
            TabIndex        =   9
            ToolTipText     =   "Cantidad de límite de crédito"
            Top             =   2775
            Width           =   1500
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Empresa"
            Height          =   200
            Index           =   4
            Left            =   2160
            TabIndex        =   5
            ToolTipText     =   "Tipo de cliente"
            Top             =   1425
            Width           =   960
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Médico"
            Height          =   200
            Index           =   3
            Left            =   2160
            TabIndex        =   4
            ToolTipText     =   "Tipo de cliente"
            Top             =   1230
            Width           =   870
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Empleado"
            Height          =   200
            Index           =   2
            Left            =   2160
            TabIndex        =   3
            ToolTipText     =   "Tipo de cliente"
            Top             =   1020
            Width           =   1020
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Paciente externo"
            Height          =   200
            Index           =   1
            Left            =   2160
            TabIndex        =   2
            ToolTipText     =   "Tipo de cliente"
            Top             =   810
            Width           =   1500
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Paciente interno"
            Height          =   200
            Index           =   0
            Left            =   2160
            TabIndex        =   1
            ToolTipText     =   "Tipo de cliente"
            Top             =   600
            Width           =   1500
         End
         Begin VB.TextBox txtNumeroCliente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   "Número de cliente"
            Top             =   255
            Width           =   930
         End
         Begin MSMask.MaskEdBox mskCuentaContable 
            Height          =   315
            Left            =   2160
            TabIndex        =   7
            ToolTipText     =   "Cuenta contable del cliente"
            Top             =   2415
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   8040
            TabIndex        =   46
            Top             =   3240
            Width           =   120
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje de crédito disponible para envío de advertencia al ingreso"
            Height          =   195
            Left            =   2160
            TabIndex        =   45
            Top             =   3240
            Width           =   4920
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Calcular el vencimiento"
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   4875
            Width           =   1635
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Departamento a cargo"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   6450
            Width           =   1590
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Límite de crédito"
            Height          =   195
            Left            =   3825
            TabIndex        =   32
            Top             =   2835
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha de otorgamiento de crédito"
            Height          =   375
            Left            =   225
            TabIndex        =   31
            Top             =   2775
            Width           =   1845
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta contable"
            Height          =   195
            Left            =   225
            TabIndex        =   30
            Top             =   2505
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   225
            TabIndex        =   29
            Top             =   2115
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de cliente"
            Height          =   195
            Left            =   225
            TabIndex        =   28
            Top             =   585
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   225
            TabIndex        =   27
            Top             =   315
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmMantoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------
' Programa para dar mantenimiento al catálogo de clientes (CcCliente)
' Fecha de programación: Miércoles 28 de Febrero de 2001
'---------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
'---------------------------------------------------------------------------------

Option Explicit


Public llngNumOpcion As Long 'Número de opción en el modulo
Public lblnTodosClientes As Boolean 'Para que se visualicen todos los cliente

Dim rs As New ADODB.Recordset

Dim vlstrx As String

Dim vlblnConsulta As Boolean


Private Sub cboDepartamentos_GotFocus()
On Error GoTo NotificaError
        
   pHabilita 0, 1, 0
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamentos_GotFocus"))
End Sub

Private Sub cboDepartamentos_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then chkActivo.SetFocus
End Sub

Private Sub cboFormaPago_GotFocus()
pHabilita 0, 1, 0
End Sub

Private Sub cboFormaPago_KeyPress(KeyAscii As Integer)

If cboFormaPago.ListIndex = -1 Then
    txtDescPago.Enabled = False
    txtDescPago.Text = ""
ElseIf cboFormaPago.ListIndex = 0 Then
    txtDescPago.Enabled = False
    txtDescPago.Text = ""
'ElseIf cboFormaPago.ListIndex = 2 Then
'    txtDescPago.Enabled = False
'    txtDescPago.Text = ""
Else
    txtDescPago.Enabled = True
End If

If KeyAscii = 13 And txtDescPago.Enabled = True Then
    txtDescPago.SetFocus
Else
    optTipoVencimiento(0).SetFocus
End If

End Sub


Private Sub cboFormaPago_LostFocus()

If cboFormaPago.ListIndex = -1 Then
    txtDescPago.Enabled = False
    txtDescPago.Text = ""
ElseIf cboFormaPago.ListIndex = 0 Then
    txtDescPago.Enabled = False
    txtDescPago.Text = ""
'ElseIf cboFormaPago.ListIndex = 2 Then
'    txtDescPago.Enabled = False
'    txtDescPago.Text = ""
Else
    txtDescPago.Enabled = True
End If

End Sub


Private Sub chkActivo_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_GotFocus"))
End Sub

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then cmdSave.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_KeyPress"))
End Sub

Private Sub cmdBuscar_Click()
    fraBusqueda.Visible = True
    txtIniciales.Text = ""
    lstBusqueda.Clear
    txtIniciales.SetFocus
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ValidaIntegridad

    Dim strSentencia As String
    Dim lngPersonaGraba As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "C") Then
        lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    
        If lngPersonaGraba = 0 Then Exit Sub
    
        strSentencia = "delete from CcCliente where intNumCliente = " & txtNumeroCliente.Text
        pEjecutaSentencia strSentencia

        pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "CLIENTES", txtNumeroCliente.Text
    
        txtNumeroCliente.SetFocus
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If
        
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    Dim lngNumCliente As Long
    Dim rsMovimientocred As ADODB.Recordset
    Dim vlstrsql As String
    
    Label4.Enabled = True
    mskCuentaContable.Enabled = True
    txtDescripcionCuenta.Enabled = True
    
    lngNumCliente = flngNumCliente(lblnTodosClientes, 0)
        
    If lngNumCliente <> 0 Then
        pMuestraCliente lngNumCliente
        
        
vlstrsql = "Select intnumcuentacontable from ccmovimientocredito where intnumcliente = " & txtNumeroCliente.Text
            Set rsMovimientocred = frsRegresaRs(vlstrsql)
            If rsMovimientocred.RecordCount <> 0 Then
                Label4.Enabled = False
                mskCuentaContable.Enabled = False
                txtDescripcionCuenta.Enabled = False
            End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    Dim rsCcCliente As New ADODB.Recordset
    Dim strSentencia As String
    Dim lngPersonaGraba As Long
    Dim strCveMetodoPagoCFDI As String
    
    'Checar el pemiso que le mandan
    If fblnRevisaPermiso(vglngNumeroLogin, 318, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 318, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 620, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 620, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 1139, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 1139, "C", True) Then
    
      If fblnDatosValidos() Then
        
        Dim vllngNumeroCuenta As Long
        Dim vlintErrorCuenta As Integer
        
        vllngNumeroCuenta = flngNumeroCuenta(mskCuentaContable.Text, vgintClaveEmpresaContable)
        vlintErrorCuenta = fintValidaCuenta(vllngNumeroCuenta)
         
        If (vlintErrorCuenta = 1) Then 'La cuenta seleccionada no acepta movimientos.
          MsgBox SIHOMsg(375), vbOKOnly + vbExclamation, "Mensaje"
          mskCuentaContable.SetFocus
          pSelMkTexto mskCuentaContable
        Else
          lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
          If lngPersonaGraba = 0 Then Exit Sub
        
        
          If cboFormaPago.ListIndex > -1 Then
                strCveMetodoPagoCFDI = fstrClaveMetodoPago(cboFormaPago.ItemData(cboFormaPago.ListIndex))
          Else
                strCveMetodoPagoCFDI = ""
          End If

        
        
          strSentencia = "select CcCliente.* from CcCliente where CcCliente.intNumCliente = " & IIf(vlblnConsulta, txtNumeroCliente.Text, "-1")
          Set rsCcCliente = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)

          With rsCcCliente
              If Not vlblnConsulta Then
                .AddNew
                !intNumReferencia = lstBusqueda.ItemData(lstBusqueda.ListIndex)
                !chrTipoCliente = IIf(optTipoCliente(0).Value, "PI", IIf(optTipoCliente(1).Value, "PE", IIf(optTipoCliente(2).Value, "EM", IIf(optTipoCliente(3).Value, "ME", "CO"))))
              End If
              !intnumcuentacontable = flngNumeroCuenta(mskCuentaContable.Text, vgintClaveEmpresaContable)
              !dtmFechaAsignacion = CDate(txtFechaAsignacion.Text)
              !mnyLimiteCredito = Val(Format(txtLimiteCredito.Text, "############.##"))
              !SMICVEDEPARTAMENTO = cboDepartamentos.ItemData(cboDepartamentos.ListIndex)
              !intTipoVencimiento = IIf(optTipoVencimiento(0).Value, 0, IIf(optTipoVencimiento(1).Value, 1, 2))
              If optTipoVencimiento(0).Value Then
                !smiDiasVencimiento = 0
              End If
              If optTipoVencimiento(1).Value Then
                !smiDiasVencimiento = Val(txtDiasEntrega.Text)
              End If
              If optTipoVencimiento(2).Value Then
                !smiDiasVencimiento = Val(txtDiasVencimiento.Text)
              End If
              !bitactivo = chkActivo.Value
              
              'Nuevos campos para CFD y CFDi
              !VCHNUMCTAPAGO = Trim(txtDescPago.Text)
              !VCHTIPOPAGO = strCveMetodoPagoCFDI
             
              !RELPORCENTAJEADVERTENCIA = Val(Format(txtPorcentaje.Text, "###.##"))
                           
              .Update
              If Not vlblnConsulta Then
                txtNumeroCliente.Text = flngObtieneIdentity("SEC_CCCLIENTE", !INTNUMCLIENTE)
              End If
          End With
          rsCcCliente.Close

          pGuardarLogTransaccion Me.Name, EnmGrabar, lngPersonaGraba, "CLIENTE", txtNumeroCliente.Text
          txtNumeroCliente.SetFocus
        End If
      End If

Else
MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub



Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    If cboDepartamentos.ListCount = 0 Then
        'No existen departamentos registrados.
        MsgBox SIHOMsg(239), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If
    cboDepartamentos.Enabled = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "C")
'cboFormaPago.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If fraBusqueda.Visible Then
            fraBusqueda.Visible = False
            cmdBuscar.SetFocus
        Else
            If cmdSave.Enabled Or vlblnConsulta Then
                ' ¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    txtNumeroCliente.SetFocus
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
    
    Dim rs As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Set rs = frsRegresaRs("select vchDescripcion,smiCveDepartamento from NoDepartamento where nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable)
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboDepartamentos, rs, 1, 0
    End If
    rs.Close
    'Se cargan los diferentes métodos de pago configurados en el catálogo de formas de pago
    
    pLlenarCboSentencia cboFormaPago, "SELECT INTIDREGISTRO, VCHDESCRIPCION FROM PVMETODOPAGOSATCFDI ORDER BY VCHDESCRIPCION", 1, 0
    cboFormaPago.AddItem "<NINGUNA>", 0
    cboFormaPago.ItemData(cboFormaPago.newIndex) = 0

    sstabClientes.Tab = 0

    cboFormaPago.ListIndex = 0
    
    pHabilitaPorcentaje

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub lstBusqueda_DblClick()
    
    If lstBusqueda.ListCount <> 0 Then
        txtNombreCliente.Text = lstBusqueda.List(lstBusqueda.ListIndex)
        mskCuentaContable.SetFocus
    End If
    
End Sub

Private Sub lstBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        lstBusqueda_DblClick
    End If
    
End Sub

Private Sub lstBusqueda_LostFocus()
    
    fraBusqueda.Visible = False

End Sub

Private Sub mskCuentaContable_Change()
  txtDescripcionCuenta.Text = ""
End Sub

Private Sub mskCuentaContable_GotFocus()
    
    pHabilita 0, 1, 0

    pSelMkTexto mskCuentaContable

End Sub

Private Sub mskCuentaContable_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        pAsignaCuenta mskCuentaContable, txtDescripcionCuenta
        txtFechaAsignacion.SetFocus
    End If

End Sub

Private Sub pAsignaCuenta(mskobject As MaskEdBox, txtObject As TextBox)
    On Error GoTo NotificaError
    
    Dim vllngNumeroCuenta As Long
    Dim vlstrCuentaCompleta As String

    If Trim(mskobject.ClipText) = "" Then
        vllngNumeroCuenta = flngBusquedaCuentasContables(False, vgintClaveEmpresaContable)
        
        If vllngNumeroCuenta <> 0 Then
            mskobject.Text = fstrCuentaContable(vllngNumeroCuenta)
        End If
    End If
    
    vlstrCuentaCompleta = fstrCuentaCompleta(mskobject.Text)
    
    mskobject.Mask = ""
    mskobject.Text = vlstrCuentaCompleta
    mskobject.Mask = vgstrEstructuraCuentaContable
    
    vllngNumeroCuenta = flngNumeroCuenta(mskobject.Text, vgintClaveEmpresaContable)
    
    If vllngNumeroCuenta <> 0 Then
        txtObject.Text = fstrDescripcionCuenta(mskobject.Text, vgintClaveEmpresaContable)
    Else
        'No se encontró la cuenta contable.
        MsgBox SIHOMsg(222), vbOKOnly + vbExclamation, "Mensaje"
        mskobject.Mask = ""
        mskobject.Text = ""
        mskobject.Mask = vgstrEstructuraCuentaContable
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsignaCuenta"))
    Unload Me
End Sub

Private Sub optTipoCliente_Click(Index As Integer)
    pColor
    pHabilitaPorcentaje
End Sub

Private Sub pColor()
    Dim intContador As Long
    
    For intContador = 0 To 4
        If optTipoCliente(intContador).Value Then
            optTipoCliente(intContador).ForeColor = &HC00000
        Else
            optTipoCliente(intContador).ForeColor = &H80000012
        End If
    Next intContador

End Sub

Private Sub optTipoCliente_GotFocus(Index As Integer)
    On Error GoTo NotificaError
    
    If vlblnConsulta Then
        If mskCuentaContable.Enabled Then
            mskCuentaContable.SetFocus
        Else
            If txtNombreCliente.Enabled Then
                txtNombreCliente.SetFocus
            Else
                If txtFechaAsignacion.Enabled Then
                    txtFechaAsignacion.SetFocus
                End If
            End If
        End If
    End If
    pHabilita 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoCliente_GotFocus"))
End Sub

Private Sub optTipoCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If

    pHabilitaPorcentaje

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoCliente_KeyPress"))
End Sub


Private Sub optTipoVencimiento_Click(Index As Integer)

    lblDiasEntrega.Enabled = optTipoVencimiento(1).Value
    txtDiasEntrega.Enabled = optTipoVencimiento(1).Value
    txtDiasEntrega.Text = IIf(Not optTipoVencimiento(1).Value, " ", txtDiasEntrega.Text)

    lblDias.Enabled = optTipoVencimiento(2).Value
    txtDiasVencimiento.Enabled = optTipoVencimiento(2).Value
    txtDiasVencimiento.Text = IIf(Not optTipoVencimiento(2).Value, " ", txtDiasVencimiento.Text)
    
End Sub


Private Sub optTipoVencimiento_GotFocus(Index As Integer)

    pHabilita 0, 1, 0

End Sub

Private Sub optTipoVencimiento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If optTipoVencimiento(1).Value Then
            txtDiasEntrega.SetFocus
        Else
            If optTipoVencimiento(2).Value Then
                txtDiasVencimiento.SetFocus
            Else
                If cboDepartamentos.Enabled Then
                    cboDepartamentos.SetFocus
                Else
                    chkActivo.SetFocus
                End If
            End If
        End If
    End If

End Sub



Private Sub txtDescPago_GotFocus()
pHabilita 0, 1, 0

pSelTextBox txtDescPago
End Sub

Private Sub txtDescPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        optTipoVencimiento(0).SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub txtDiasEntrega_GotFocus()

    pHabilita 0, 1, 0
    pSelTextBox txtDiasEntrega

End Sub

Private Sub txtDiasEntrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cboDepartamentos.Enabled Then
            cboDepartamentos.SetFocus
        Else
            chkActivo.SetFocus
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If


End Sub

Private Sub txtDiasVencimiento_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 1, 0
    pSelTextBox txtDiasVencimiento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDiasVencimiento_GotFocus"))
End Sub

Private Sub txtDiasVencimiento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If cboDepartamentos.Enabled Then
            cboDepartamentos.SetFocus
        Else
            chkActivo.SetFocus
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDiasVencimiento_KeyPress"))
End Sub

Private Sub txtFechaAsignacion_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 1, 0
    pSelMkTexto txtFechaAsignacion

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaAsignacion_GotFocus"))
End Sub

Private Sub txtFechaAsignacion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtLimiteCredito.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaAsignacion_KeyPress"))
End Sub

Private Sub txtFechaAsignacion_LostFocus()
    
    If Not IsDate(txtFechaAsignacion.Text) Then
        txtFechaAsignacion.Mask = ""
        txtFechaAsignacion.Text = fdtmServerFecha
        txtFechaAsignacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub txtIniciales_Change()
    lstBusqueda.Clear
    lstBusqueda.Visible = False

    If txtIniciales.Text <> "" Then
    
        If optTipoCliente(0).Value Then
            Set rs = frsEjecuta_SP(Trim(txtIniciales.Text) & "|" & vgintClaveEmpresaContable, "SP_CCSelNombrePacInt")
        Else
            If optTipoCliente(1).Value Then
                Set rs = frsEjecuta_SP(Trim(txtIniciales.Text) & "|" & vgintClaveEmpresaContable, "SP_CCSelNombrePacExt")
            Else
                If optTipoCliente(2).Value Then
                    'Empleados activos
                    Set rs = frsEjecuta_SP(Trim(txtIniciales.Text) & "|" & vgintClaveEmpresaContable, "SP_CCSelNombreEmpleados")
                Else
                    If optTipoCliente(3).Value Then
                        'Médicos activos
                        Set rs = frsEjecuta_SP(Trim(txtIniciales.Text), "SP_CCSelNombreMedicos")
                    Else
                        'Empresas activas
                        Set rs = frsEjecuta_SP(Trim(txtIniciales.Text), "SP_CCSelNombreEmpresas")
                    End If
                End If
            End If
        End If
    
        If rs.State <> adStateClosed Then
            If rs.RecordCount <> 0 Then
                pLlenarListRs lstBusqueda, rs, 1, 0
                lstBusqueda.ListIndex = 0
            End If
        End If
    
        lstBusqueda.Visible = True
    End If

End Sub

Private Sub txtIniciales_GotFocus()
    
    fraBusqueda.Visible = True
    pSelTextBox txtIniciales
    
End Sub

Private Sub txtIniciales_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If lstBusqueda.Visible Then
            lstBusqueda.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub txtLimiteCredito_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 1, 0
    pSelTextBox txtLimiteCredito

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtLimiteCredito_GotFocus"))
End Sub

Private Sub txtLimiteCredito_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not fblnFormatoCantidad(txtLimiteCredito, KeyAscii, 2) Then
       KeyAscii = 7
    Else
        pHabilitaPorcentaje
    
        If KeyAscii = 13 Then SendKeys vbTab
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtLimiteCredito_KeyPress"))
End Sub

Private Sub txtLimiteCredito_LostFocus()
    pHabilitaPorcentaje
End Sub

Private Sub txtNumeroCliente_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 1, 0, 0
    pLimpia
    pSelTextBox txtNumeroCliente

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumeroCliente_GotFocus"))
End Sub
Private Sub pLimpia()
    On Error GoTo NotificaError
        
    vlblnConsulta = False
    cmdBuscar.Enabled = True
    
    Set rs = frsRegresaRs("select NVL(max(intNumCliente),0)+1 from ccCliente")
    If rs.RecordCount <> 0 Then
        txtNumeroCliente.Text = rs.Fields(0)
    Else
        txtNumeroCliente.Text = "1"
    End If
    
    optTipoCliente(0).Value = False
    optTipoCliente(1).Value = False
    optTipoCliente(2).Value = False
    optTipoCliente(3).Value = False
    optTipoCliente(4).Value = False
    pColor
    
    txtIniciales.Text = ""
    lstBusqueda.Clear
    fraBusqueda.Visible = False
    
    txtNombreCliente.Text = ""
    
    mskCuentaContable.Mask = ""
    mskCuentaContable.Text = ""
    mskCuentaContable.Mask = vgstrEstructuraCuentaContable
    
    txtDescripcionCuenta.Text = ""
    
    txtFechaAsignacion.Mask = ""
    txtFechaAsignacion.Text = ""
    txtFechaAsignacion.Mask = "##/##/####"
    
    txtLimiteCredito.Text = ""
    txtPorcentaje.Text = ""
    
    optTipoVencimiento(0).Value = True
    txtDiasVencimiento.Text = ""
    
        
    chkActivo.Value = 1
    
    cboDepartamentos.ListIndex = flngLocalizaCbo(cboDepartamentos, Str(vgintNumeroDepartamento))
    
    'Se restauran los campos de la forma de pago por default
    txtDescPago.Text = ""
    
    cboFormaPago.ListIndex = -1
    
    If cboFormaPago.ListIndex = -1 Then
        txtDescPago.Enabled = False
        txtDescPago.Text = ""
    ElseIf cboFormaPago.ListIndex = 0 Then
        txtDescPago.Enabled = False
        txtDescPago.Text = ""
'    ElseIf cboFormaPago.ListIndex = 2 Then
'        txtDescPago.Enabled = False
'        txtDescPago.Text = ""
    Else
        txtDescPago.Enabled = True
    End If

    pHabilitaPorcentaje
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer)
    On Error GoTo NotificaError
    
    cmdLocate.Enabled = vlb1 = 1
    cmdSave.Enabled = vlb2 = 1
    cmdDelete.Enabled = vlb3 = 1
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
End Sub

Private Sub txtNumeroCliente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    Dim rsMovimientocred As ADODB.Recordset
    Dim vlstrsql As String
    
    Label4.Enabled = True
    mskCuentaContable.Enabled = True
    txtDescripcionCuenta.Enabled = True
    
    If KeyAscii = 13 Then
        If Trim(txtNumeroCliente.Text) = "" Then
            pLimpia
        Else
            pMuestraCliente Val(txtNumeroCliente.Text)
                
            vlstrsql = "Select intnumcuentacontable from ccmovimientocredito where intnumcliente = " & txtNumeroCliente.Text
            Set rsMovimientocred = frsRegresaRs(vlstrsql)
            If rsMovimientocred.RecordCount <> 0 Then
                Label4.Enabled = False
                mskCuentaContable.Enabled = False
                txtDescripcionCuenta.Enabled = False
            End If

        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumeroCliente_KeyPress"))
End Sub

Private Sub pMuestraCliente(vllngxNumero As Long)
On Error GoTo NotificaError
Dim rs As New ADODB.Recordset
    
    vgstrParametrosSP = Str(vllngxNumero) & "|0|*|*|" & CStr(vgintClaveEmpresaContable) & "|0"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelDatosCliente")
    If rs.RecordCount <> 0 Then
    
        If rs!SMICVEDEPARTAMENTO = vgintNumeroDepartamento Or lblnTodosClientes Then
                
            vlblnConsulta = True
                    
            cmdBuscar.Enabled = False
            
            txtNumeroCliente.Text = Str(rs!INTNUMCLIENTE)
            
            optTipoCliente(0).Value = rs!chrTipoCliente = "PI"
            optTipoCliente(1).Value = rs!chrTipoCliente = "PE"
            optTipoCliente(2).Value = rs!chrTipoCliente = "EM"
            optTipoCliente(3).Value = rs!chrTipoCliente = "ME"
            optTipoCliente(4).Value = rs!chrTipoCliente = "CO"
            
            txtNombreCliente.Text = IIf(IsNull(rs!NombreCliente), " ", rs!NombreCliente)
            
            mskCuentaContable.Mask = ""
            mskCuentaContable.Text = fstrCuentaContable(rs!intnumcuentacontable)
            mskCuentaContable.Mask = vgstrEstructuraCuentaContable
            
            txtDescripcionCuenta.Text = fstrDescripcionCuenta(mskCuentaContable.Text, vgintClaveEmpresaContable)
            
            txtFechaAsignacion.Mask = ""
            txtFechaAsignacion.Text = rs!dtmFechaAsignacion
            txtFechaAsignacion.Mask = "##/##/####"
            
            txtLimiteCredito.Text = IIf(rs!mnyLimiteCredito = 0, "", Format(Str(rs!mnyLimiteCredito), "###,###,###,###.00"))
            
            optTipoVencimiento(rs!intTipoVencimiento).Value = True
            
            If rs!intTipoVencimiento = 0 Then
                txtDiasEntrega.Text = ""
                txtDiasVencimiento.Text = ""
            End If
            If rs!intTipoVencimiento = 1 Then
                txtDiasEntrega.Text = Format(rs!smiDiasVencimiento)
                txtDiasVencimiento.Text = ""
            End If
            If rs!intTipoVencimiento = 2 Then
                txtDiasEntrega.Text = ""
                txtDiasVencimiento.Text = Format(rs!smiDiasVencimiento)
            End If
            
            chkActivo.Value = rs!bitactivo
            cboDepartamentos.ListIndex = flngLocalizaCbo(cboDepartamentos, Str(rs!SMICVEDEPARTAMENTO))
            
            'Se carga la información de forma de pago por default
            'cboFormaPago.ListIndex = flngLocalizaCboTxt(cboFormaPago, IIf(IsNull(rs!VCHTIPOPAGO), " ", rs!VCHTIPOPAGO))
            
            If IsNull(rs!VCHTIPOPAGO) Then
                cboFormaPago.ListIndex = 0
            Else
                cboFormaPago.ListIndex = flngLocalizaCbo(cboFormaPago, fintIDMetodoPago(rs!VCHTIPOPAGO))
            End If
            
            If cboFormaPago.ListIndex = -1 Then
                txtDescPago.Enabled = False
            ElseIf cboFormaPago.ListIndex = 0 Then
                txtDescPago.Enabled = False
            Else
                txtDescPago.Enabled = True
            End If
            
            txtDescPago.Text = IIf(IsNull(rs!VCHNUMCTAPAGO), "", rs!VCHNUMCTAPAGO)
            
            If rs!RELPORCENTAJEADVERTENCIA = 0 Then
                txtPorcentaje.Text = FormatNumber("0.00", 2)
            Else
                txtPorcentaje.Text = FormatNumber(Str(rs!RELPORCENTAJEADVERTENCIA), 2)
            End If
            
            pHabilita 1, 0, 1
            cmdLocate.SetFocus
        
            pHabilitaPorcentaje
        
        Else
            'El cliente seleccionado no pertenece a este departamento.
            MsgBox SIHOMsg(646), vbOKOnly + vbInformation, "Mensaje"
            pEnfocaTextBox txtNumeroCliente
        End If
    Else
        optTipoCliente(0).SetFocus
    End If
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraCliente"))
End Sub


Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError
Dim rsClienteRepetido As New ADODB.Recordset
Dim lngReferencia As Long
    
    fblnDatosValidos = True
    
    fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "C")
    If Not fblnDatosValidos Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
    If fblnDatosValidos And Not optTipoCliente(0).Value And Not optTipoCliente(1).Value And Not optTipoCliente(2).Value And Not optTipoCliente(3).Value And Not optTipoCliente(4).Value Then
        fblnDatosValidos = False
        'Seleccione el tipo de cliente.
        MsgBox SIHOMsg(321), vbOKOnly + vbInformation, "Mensaje"
        optTipoCliente(0).SetFocus
    End If
    If fblnDatosValidos And Trim(txtNombreCliente.Text) = "" Then
        fblnDatosValidos = False
        'Seleccione el cliente.
        MsgBox SIHOMsg(322), vbOKOnly + vbInformation, "Mensaje"
        cmdBuscar.SetFocus
    End If
    If fblnDatosValidos And Trim(mskCuentaContable.ClipText) = "" Then
        fblnDatosValidos = False
        'Seleccione la cuenta contable.
        MsgBox SIHOMsg(211), vbOKOnly + vbInformation, "Mensaje"
        mskCuentaContable.SetFocus
    End If
    If fblnDatosValidos Then
        If flngNumeroCuenta(mskCuentaContable.Text, vgintClaveEmpresaContable) = 0 Then
            fblnDatosValidos = False
            'Seleccione la cuenta contable.
            MsgBox SIHOMsg(211), vbOKOnly + vbInformation, "Mensaje"
            mskCuentaContable.SetFocus
        End If
    End If
    If fblnDatosValidos And Trim(txtDescripcionCuenta.Text) = "" Then
        fblnDatosValidos = False
        'Seleccione la cuenta contable.
        MsgBox SIHOMsg(211), vbOKOnly + vbInformation, "Mensaje"
        mskCuentaContable.SetFocus
    End If
    If fblnDatosValidos And optTipoVencimiento(1).Value And Val(txtDiasEntrega.Text) = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDiasEntrega.SetFocus
    End If
    If fblnDatosValidos And optTipoVencimiento(2).Value And Val(txtDiasVencimiento.Text) = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDiasVencimiento.SetFocus
    End If
    If fblnDatosValidos Then
        If Not vlblnConsulta Then
            If lstBusqueda.ListIndex <> -1 Then
                vlstrx = "select count(*) from CcCliente inner join NoDepartamento on CcCliente.smiCveDepartamento = NoDepartamento.smiCveDepartamento and NoDepartamento.tnyClaveEmpresa = " & CStr(vgintClaveEmpresaContable) & " " & _
                     "where intNumReferencia=" + Str(lstBusqueda.ItemData(lstBusqueda.ListIndex)) + " and chrTipoCliente=" + "'" & IIf(optTipoCliente(0).Value, "PI", IIf(optTipoCliente(1).Value, "PE", IIf(optTipoCliente(2).Value, "EM", IIf(optTipoCliente(3).Value, "ME", "CO")))) & "'"
                Set rsClienteRepetido = frsRegresaRs(vlstrx)
                If rsClienteRepetido.Fields(0) <> 0 Then
                    fblnDatosValidos = False
                    'Este cliente ya se encuentra registrado.
                    MsgBox SIHOMsg(323), vbOKOnly + vbInformation, "Mensaje"
                    cmdBuscar.SetFocus
                End If
            Else
                fblnDatosValidos = False
                'Este cliente ya se encuentra registrado.
                MsgBox SIHOMsg(323), vbOKOnly + vbInformation, "Mensaje"
                cmdBuscar.SetFocus
            End If
        ElseIf chkActivo.Value = 1 Then
            
            vlstrx = "Select intNumReferencia From CcCliente Where intNumCliente = " & CStr(Trim(txtNumeroCliente.Text))
            Set rsClienteRepetido = frsRegresaRs(vlstrx)
            If Not rsClienteRepetido.EOF Then
                lngReferencia = rsClienteRepetido!intNumReferencia
            End If
            
            vlstrx = "Select * " & _
                     "  From CcCliente " & _
                     "       inner join NoDepartamento on (CcCliente.smiCveDepartamento = NoDepartamento.smiCveDepartamento and NoDepartamento.tnyClaveEmpresa = " & CStr(vgintClaveEmpresaContable) & ") " & _
                     " Where intNumReferencia = " & Str(lngReferencia) & _
                     "   And chrTipoCliente = '" & IIf(optTipoCliente(0).Value, "PI", IIf(optTipoCliente(1).Value, "PE", IIf(optTipoCliente(2).Value, "EM", IIf(optTipoCliente(3).Value, "ME", "CO")))) & "'" & _
                     "   And bitActivo = 1 " & _
                     "   And intNumCliente <> " & CStr(Trim(txtNumeroCliente.Text))
            Set rsClienteRepetido = frsRegresaRs(vlstrx)
            If Not rsClienteRepetido.EOF Then
                fblnDatosValidos = False
                'Este cliente ya se encuentra registrado.
                MsgBox SIHOMsg(323), vbOKOnly + vbInformation, "Mensaje"
                txtNumeroCliente.SetFocus
            End If
                    
        End If
        
    End If
    
'    'Validaciones de los campos de pago por default
    If Len(Trim(txtDescPago.Text)) > 0 Then
        If fblnDatosValidos And Len(Trim(txtDescPago.Text)) < 4 Then
            fblnDatosValidos = False
            '¡Se deben de indicar almenos 4 dígitos de la referencia del pago por defecto!
            MsgBox "¡Se deben de indicar al menos 4 dígitos de la referencia de la pago predeterminada!", vbOKOnly + vbInformation, "Mensaje"
            txtDescPago.SetFocus
            pSelTextBox txtDescPago
            Exit Function
        End If
    End If

'    'Validaciones de los campos de pago por default
'    Select Case cboFormaPago.ListIndex
'        Case 1
'            If fblnDatosValidos And Len(Trim(txtDescPago.Text)) < 4 Then
'                fblnDatosValidos = False
'                '¡Se deben de indicar almenos 4 dígitos de la referencia del pago por defecto!
'                MsgBox "¡Se deben de indicar al menos 4 dígitos de la referencia de la pago predeterminada!", vbOKOnly + vbInformation, "Mensaje"
'                txtDescPago.SetFocus
'                pSelTextBox txtDescPago
'                Exit Function
'            End If
'        Case 3
'            If fblnDatosValidos And Len(Trim(txtDescPago.Text)) < 4 Then
'                fblnDatosValidos = False
'                '¡Se deben de indicar almenos 4 dígitos de la referencia del pago por defecto!
'                MsgBox "¡Se deben de indicar al menos 4 dígitos de la referencia de la forma de pago predeterminada!", vbOKOnly + vbInformation, "Mensaje"
'                txtDescPago.SetFocus
'                pSelTextBox txtDescPago
'                Exit Function
'            End If
'        Case 4
'            If fblnDatosValidos And Len(Trim(txtDescPago.Text)) < 4 Then
'                fblnDatosValidos = False
'                '¡Se deben de indicar almenos 4 dígitos de la referencia del pago por defecto!
'                MsgBox "¡Se deben de indicar al menos 4 dígitos de la referencia de la forma de pago predeterminada!", vbOKOnly + vbInformation, "Mensaje"
'                txtDescPago.SetFocus
'                pSelTextBox txtDescPago
'                Exit Function
'            End If
'    End Select
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function

Private Function fintValidaCuenta(vlngNumero As Long) As Integer
    '=========================================================================================
    ' Función para validar la cuenta antes de incluirla en el detalle de la póliza
    ' Regresa también en la variable <vlintOrden> si es una cuenta de orden o no
    '=========================================================================================
    
    On Error GoTo NotificaError
    
    Dim rsCuenta As New ADODB.Recordset
    Dim vlstrSentencia As String
   
    ' Valores de regreso (Errores):
    ' 1 = Que la cuenta no acepte movimientos
    ' 2 = Que la fecha de la cuenta sea mayor a la fecha de la póliza
    ' 0 = No hay error
    
    fintValidaCuenta = 0
    
    vlstrSentencia = "select * from CnCuenta where intNumeroCuenta=" & vlngNumero
    Set rsCuenta = frsRegresaRs(vlstrSentencia)
     
    If rsCuenta.RecordCount <> 0 Then
'        If Trim(rsCuenta!vchTipo) = "Orden" Then
'            vlintOrden = 1
'        Else
'            vlintOrden = 0
'        End If
    
        If rsCuenta!bitEstatusMovimientos = 0 Then
            fintValidaCuenta = 1
'        Else
'            If CDate(mskFecha.Text) < rsCuenta!DTMFECHAINICIO Then
'                fintValidaCuenta = 2
'            End If
        End If
    End If

Exit Function
NotificaError:
   Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintValidaCuenta"))
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
        fintIDMetodoPago = rs!INTIDREGISTRO
    Else
        fintIDMetodoPago = -1
    End If
    rs.Close
End Function

Private Sub txtPorcentaje_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 1, 0
    pSelTextBox txtPorcentaje

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPorcentaje_GotFocus"))
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not fblnFormatoCantidad(txtPorcentaje, KeyAscii, 2) Then
        KeyAscii = 7
    Else
        If KeyAscii = 13 Then
            If txtPorcentaje.Text = "." Or txtPorcentaje.Text = "" Or txtPorcentaje.Text = " " Then txtPorcentaje.Text = "0"
            
            txtPorcentaje.Text = FormatNumber(txtPorcentaje.Text, 2)
        
            If CDbl(txtPorcentaje.Text) >= 100 Then
                'Dato incorrecto: El porcentaje debe ser menor a 100%
                MsgBox SIHOMsg(35), vbOKOnly + vbInformation, "Mensaje"
                
                txtPorcentaje.Text = FormatNumber("0.00", 2)
                
                pSelTextBox txtPorcentaje
                
                Exit Sub
            Else
                txtPorcentaje.Text = FormatNumber(txtPorcentaje.Text, 2)
            End If
            
            SendKeys vbTab
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPorcentaje_KeyPress"))
End Sub

Private Sub pHabilitaPorcentaje()
On Error GoTo NotificaError

    If optTipoCliente(4).Value = True And Val(txtLimiteCredito.Text) > 0 Then
        txtPorcentaje.Enabled = True
    Else
        txtPorcentaje.Enabled = False
        txtPorcentaje.Text = FormatNumber("0.00", 2)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaPorcentaje"))
End Sub

Private Sub txtPorcentaje_LostFocus()
    If txtPorcentaje.Text = "." Or txtPorcentaje.Text = "" Or txtPorcentaje.Text = " " Then txtPorcentaje.Text = "0"
    
    txtPorcentaje.Text = FormatNumber(txtPorcentaje.Text, 2)

    If CDbl(txtPorcentaje.Text) >= 100 Then
        'Dato incorrecto: El porcentaje debe ser menor a 100%
        MsgBox SIHOMsg(35), vbOKOnly + vbInformation, "Mensaje"
        
        txtPorcentaje.Text = FormatNumber("0.00", 2)
    Else
        txtPorcentaje.Text = FormatNumber(txtPorcentaje.Text, 2)
    End If
End Sub
