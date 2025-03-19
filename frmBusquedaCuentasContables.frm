VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBusquedaCuentasContables 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Búsqueda de cuentas contables"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInvisible 
      Height          =   585
      Left            =   6690
      TabIndex        =   2
      Top             =   -90
      Visible         =   0   'False
      Width           =   1935
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "Por descripción"
         Height          =   210
         Index           =   0
         Left            =   15
         TabIndex        =   4
         Top             =   120
         Width           =   1530
      End
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "Por cuenta contable"
         Height          =   210
         Index           =   1
         Left            =   15
         TabIndex        =   3
         Top             =   330
         Width           =   1800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCuentasContables 
      Height          =   3345
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Cuentas contables"
      Top             =   675
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5900
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      GridColor       =   -2147483632
      FocusRect       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   315
      Left            =   45
      MaxLength       =   100
      TabIndex        =   0
      ToolTipText     =   "Iniciales de la descripción de la cuenta"
      Top             =   330
      Width           =   8415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "* = Cuenta afectable"
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblFormaBusqueda 
      AutoSize        =   -1  'True
      Caption         =   "Leyenda de búsqueda"
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   1590
   End
End
Attribute VB_Name = "frmBusquedaCuentasContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
' Programa para cargar las cuentas contables de la empresa registrada en parámetros
' de control, carga solo cuentas que permiten movimientos y que están activas, esto
' quiere decir que son del catálogo que actualmente se esta usando
' Fecha de programacion: Noviembre del 2000
'----------------------------------------------------------------------------------

'2005-10-27  = Se filtra la cuenta de cuadre (CnCuenta.vchTipo <> 'Cuadre')
'****************************************************************************************
' Ultimas modificaciones al módulo:
'****************************************************************************************
' Fecha:2 Marzo 2009
' Descripción del cambio: se agrega la búsqueda por nivel
' Autor del cambio: I.S.C Markov Mercado
'****************************************************************************************



Option Explicit

Public vllngNumeroCuenta As Long
Public vlblnTodasCuentas As Boolean
Public vlintLvl As Integer
Public vlintcveempresa As Integer

Dim intNiveles As Integer
Dim strMask() As String
Dim strText() As String



Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    optTipoConsulta(0).Value = True
    pCargaCuentas
    
    pSelTextBox txtBusqueda
    txtBusqueda.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyF4 Then
        If optTipoConsulta(0).Value Then
            optTipoConsulta(1).Value = True
        Else
            optTipoConsulta(0).Value = True
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
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

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    vllngNumeroCuenta = 0
    pLimpiaBusqueda
    pConfiguraGrid
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pLimpiaBusqueda()
    On Error GoTo NotificaError

    grdCuentasContables.Clear
    grdCuentasContables.Rows = 2
    grdCuentasContables.Cols = 3
    
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaBusqueda"))
End Sub


Private Sub grdCuentasContables_DblClick()
    On Error GoTo NotificaError
    
    pRegresaDato
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCuentasContables_DblClick"))

End Sub

Private Sub pRegresaDato()
    On Error GoTo NotificaError
    
    If Trim(grdCuentasContables.TextMatrix(1, 1)) <> "" Then
        vlintcveempresa = 0
        vllngNumeroCuenta = grdCuentasContables.RowData(grdCuentasContables.Row)
        Unload Me
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pRegresaDato"))

End Sub
Private Sub grdCuentasContables_GotFocus()
    On Error GoTo NotificaError
    
    If Trim(grdCuentasContables.TextMatrix(1, 1)) = "" Then
        txtBusqueda.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCuentasContables_GotFocus"))

End Sub

Private Sub grdCuentasContables_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
     
    If KeyCode = vbKeyLeft Then
        txtBusqueda.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCuentasContables_KeyDown"))

End Sub

Private Sub grdCuentasContables_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pRegresaDato
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCuentasContables_KeyPress"))

End Sub


Private Sub optTipoConsulta_Click(Index As Integer)
    On Error GoTo NotificaError
    
    If Index = 0 Then
        lblFormaBusqueda.Caption = "<F4> para buscar por cuenta contable"
    Else
        lblFormaBusqueda.Caption = "<F4> para buscar por descripción"
    End If
    pLimpiaBusqueda
    pConfiguraGrid
    txtBusqueda.Text = ""
    txtBusqueda.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoConsulta_Click"))
End Sub


Private Sub txtBusqueda_Change()
    On Error GoTo NotificaError
    
    pCargaCuentas

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_Change"))
End Sub

Private Sub pCargaCuentas()
    On Error GoTo NotificaError
    
    Dim rsCuentasDescripcion As New ADODB.Recordset
    Dim vlstrsql As String
    Dim vlstrparametro As String
    Dim vlstrAndQuery As String
    Dim vlintcont As Integer
    Dim vlintContlvl As Integer
   ' Dim vlintContnumero As Integer
    
    grdCuentasContables.Visible = False
    
    If optTipoConsulta(0).Value Then
        If Not vlblnTodasCuentas Then
            If vlintcveempresa > 0 Then
                vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,CnCuenta.vchCuentaContable Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%' and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and bitEstatusActiva=1 and CnCuenta.bitEstatusMovimientos=1 order by Campo1"
            Else
                vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,CnCuenta.vchCuentaContable Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%' and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and bitEstatusActiva=1 and CnCuenta.bitEstatusMovimientos=1 order by Campo1"
            End If
        Else
            If vlintcveempresa > 0 Then
                vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,CnCuenta.vchCuentaContable Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%' and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and CnCuenta.bitEstatusActiva=1 order by Campo1"
            Else
                vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,CnCuenta.vchCuentaContable Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%' and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and CnCuenta.bitEstatusActiva=1 order by Campo1"
            End If
        End If
    Else
        If Not vlblnTodasCuentas Then
            If vlintcveempresa > 0 Then
                vlstrsql = "select CnCuenta.vchCuentaContable Campo1,case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where CnCuenta.vchCuentaContable >='" + txtBusqueda.Text + "' and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and CnCuenta.bitEstatusActiva=1 and CnCuenta.bitEstatusMovimientos=1 order by Campo1"
            Else
                vlstrsql = "select CnCuenta.vchCuentaContable Campo1,case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where CnCuenta.vchCuentaContable>='" + txtBusqueda.Text + "' and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and bitEstatusActiva=1 and CnCuenta.bitEstatusMovimientos=1 order by Campo1"
            End If
        Else
            If vlintcveempresa > 0 Then
                vlstrsql = "select CnCuenta.vchCuentaContable Campo1,case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where CnCuenta.vchCuentaContable>='" + txtBusqueda.Text + "' and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and CnCuenta.bitEstatusActiva=1 order by Campo1"
            Else
                vlstrsql = "select CnCuenta.vchCuentaContable Campo1,case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos from CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where CnCuenta.vchCuentaContable>='" + txtBusqueda.Text + "' and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' and CnCuenta.bitEstatusActiva=1 order by Campo1"
            End If
        End If
    End If
    
 If vlintLvl Then
     
     pObtenerFormato
      
  If optTipoConsulta(0).Value Then
     If Not vlblnTodasCuentas Then
           If vlintcveempresa > 0 Then
      
                    vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,Cncuenta.vchCuentaContable Campo2,Cncuenta.intNumeroCuenta,Cncuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva=1 " + _
                                 " and CnCuenta.bitEstatusMovimientos=1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
                          
                     Else
                     vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,Cncuenta.vchCuentaContable Campo2,Cncuenta.intNumeroCuenta,Cncuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva=1 " + _
                                 " and CnCuenta.bitEstatusMovimientos=1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
                     End If
                     
       Else
         
           If vlintcveempresa > 0 Then
      
                    vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,CnCuenta.vchCuentaContable Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva=1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
                          
             Else
                     vlstrsql = "select case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo1,CnCuenta.vchCuentaContable Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where upper(case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end)LIKE'" + txtBusqueda.Text + "%'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva= 1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
           End If
       
       End If  ' TODAS LAS CUENTAS
   Else
   
        If Not vlblnTodasCuentas Then
           If vlintcveempresa > 0 Then
     
                    vlstrsql = "select CnCuenta.vchCuentaContable Campo1, case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2 ,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where vchCuentaContable >='" + txtBusqueda.Text + "'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva=1 " + _
                                 " and CnCuenta.bitEstatusMovimientos=1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
                          
                     Else
                     vlstrsql = "select CnCuenta.vchCuentaContable Campo1, case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where vchCuentaContable >='" + txtBusqueda.Text + "'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva=1 " + _
                                 " and CnCuenta.bitEstatusMovimientos=1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
                     End If
                     
       Else
          If vlintcveempresa > 0 Then
      
                    vlstrsql = "select CnCuenta.vchCuentaContable Campo1, case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2 ,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where vchCuentaContable >='" + txtBusqueda.Text + "'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vlintcveempresa) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva=1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
                          
             Else
                     vlstrsql = "select CnCuenta.vchCuentaContable Campo1 ,case when CnCuentaRenombradaEjercicio.vchDescripcionCuenta is null then CnCuenta.vchDescripcionCuenta else CnCuentaRenombradaEjercicio.vchDescripcionCuenta end Campo2 ,CnCuenta.intNumeroCuenta,CnCuenta.bitEstatusMovimientos" + _
                                 " From CnCuenta left join CnCuentaRenombradaEjercicio on CnCuenta.intNumeroCuenta = CnCuentaRenombradaEjercicio.intNumeroCuenta and TO_DATE(TO_CHAR(cncuentarenombradaejercicio.dtmfechahoracambio,'DD-MM-YYYY'),'DD-MM-YYYY') > sysdate where vchCuentaContable >='" + txtBusqueda.Text + "'" + _
                                 " and CnCuenta.tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable) + _
                                 " and ltrim(rtrim(CnCuenta.vchTipo))<>'Cuadre' " + _
                                 " and CnCuenta.bitEstatusActiva=1 " + _
                                 "and CnCuenta.vchCuentaContable like '"
           End If
       
       End If ' TODAS LAS CUENTAS
       
   End If ' OPCION DE ORDEN
       
       
       For vlintcont = 0 To intNiveles
             
             intNiveles = intNiveles
             
             If vlintcont <= vlintLvl - 1 Then
             
                          vlstrparametro = vlstrparametro + Replace(strText(vlintcont), "0", "_")
                          
                          vlstrAndQuery = "and CnCuenta.vchCuentaContable not like "
                          If intNiveles + 1 > vlintcont + 1 Then
                            vlstrparametro = vlstrparametro + "."
                          End If
                     
              Else
                vlstrparametro = vlstrparametro + strText(vlintcont)
                    If intNiveles + 1 > vlintcont + 1 Then
                           vlstrparametro = vlstrparametro + "."
                     End If
             
             End If
             
             
           
       Next vlintcont
          vlstrsql = vlstrsql + vlstrparametro + "'"
          vlstrparametro = ""
          
          
        
    For vlintContlvl = 0 To vlintLvl - 1
        
            For vlintcont = 0 To intNiveles
             
                If vlintContlvl < vlintLvl - 1 Then
            
                       'remplazar los 0 por guiones del nivel primero al  nivel actual
                        If vlintContlvl >= vlintcont Then
                            vlstrparametro = vlstrparametro + Replace(strText(vlintcont), "0", "_")
                            Else
                            vlstrparametro = vlstrparametro + strText(vlintcont)
                        End If
                                                    
                          If intNiveles + 1 > vlintcont + 1 Then
                            vlstrparametro = vlstrparametro + "."
                          End If
                         
                  Else
              
                   vlstrparametro = vlstrparametro + strText(vlintcont)
                       If intNiveles + 1 > vlintcont + 1 Then
                           vlstrparametro = vlstrparametro + "."
                       End If
             
                  End If
           Next vlintcont
           vlstrAndQuery = " And CnCuenta.vchCuentaContable not like '"
          vlstrsql = vlstrsql + vlstrAndQuery + vlstrparametro + "'"
          vlstrparametro = ""
          
          
               
      Next vlintContlvl
      
      '' ORDENAR DEPENDE DEL TIPO DE BUSQUEDA
      
      If optTipoConsulta(0).Value Then
      vlstrsql = vlstrsql + " order by CnCuenta.vchDescripcionCuenta "
      Else
      vlstrsql = vlstrsql + " order by CnCuenta.vchCuentaContable "
      End If
      
      
    
    End If '' si la busqueda es por nivel
        
    With rsCuentasDescripcion
        .LockType = adLockReadOnly
        .MaxRecords = 100
        .CursorType = adOpenForwardOnly
        .ActiveConnection = EntornoSIHO.ConeccionSIHO
        .Source = vlstrsql
        .Open
    End With
    
    pLimpiaBusqueda
    
    If rsCuentasDescripcion.recordCount <> 0 Then

        Do While Not rsCuentasDescripcion.EOF
            If Trim(grdCuentasContables.TextMatrix(1, 1)) <> "" Then
                grdCuentasContables.Rows = grdCuentasContables.Rows + 1
            End If
            If rsCuentasDescripcion!Bitestatusmovimientos = 1 Then
                grdCuentasContables.TextMatrix(grdCuentasContables.Rows - 1, 0) = "*"
            End If
            grdCuentasContables.TextMatrix(grdCuentasContables.Rows - 1, 1) = rsCuentasDescripcion!Campo1
            grdCuentasContables.TextMatrix(grdCuentasContables.Rows - 1, 2) = rsCuentasDescripcion!Campo2
            grdCuentasContables.RowData(grdCuentasContables.Rows - 1) = rsCuentasDescripcion!intNumeroCuenta
            rsCuentasDescripcion.MoveNext
        Loop
    End If
    pConfiguraGrid
    
    grdCuentasContables.Visible = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaCuentas"))

End Sub
Private Sub pConfiguraGrid()
    On Error GoTo NotificaError
    
    With grdCuentasContables
        .FixedCols = 1
        .FixedRows = 1
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .FormatString = IIf(optTipoConsulta(0).Value, "|Descripción|Cuenta contable", "|Cuenta contable|Descripción")
        .ColWidth(0) = 200
        If optTipoConsulta(0).Value Then
            .ColWidth(1) = 5400
            .ColWidth(2) = 2500
        Else
            .ColWidth(1) = 2500
            .ColWidth(2) = 5400
        End If
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(0) = vbCenter
    End With
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pconfiguraGrid"))
End Sub
Private Sub txtBusqueda_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtBusqueda

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_GotFocus"))

End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyDown Then
        grdCuentasContables.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyDown"))

End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        grdCuentasContables.SetFocus
    Else
        If optTipoConsulta(0).Value Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = Asc(".") Then
                KeyAscii = 7
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyPress"))
End Sub

Private Sub pObtenerFormato()
    On Error GoTo NotificaError
    Dim intControl As Integer
    intNiveles = 0
    ReDim strMask(intNiveles)
    ReDim strText(intNiveles)
   ' Me.cboTipoReporte.AddItem "Reporte a nivel " & intNiveles + 1
    For intControl = 1 To Len(vgstrEstructuraCuentaContable)
        If Mid(vgstrEstructuraCuentaContable, intControl, 1) <> "." Then
            strMask(intNiveles) = strMask(intNiveles) & Mid(vgstrEstructuraCuentaContable, intControl, 1)
            strText(intNiveles) = strText(intNiveles) & "0"
        Else
            intNiveles = intNiveles + 1
            ReDim Preserve strMask(intNiveles)
            ReDim Preserve strText(intNiveles)
           ' Me.cboTipoReporte.AddItem "Reporte a nivel " & intNiveles + 1
        End If
    Next
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pObtenerFormato"))
End Sub
