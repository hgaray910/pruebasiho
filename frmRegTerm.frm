VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmRegTerm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de las terminales"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   14655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mensaje recibido"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   7460
      Width           =   14415
      Begin VB.TextBox txtMensaje 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   40
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   315
         Width           =   14320
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reimpresión"
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
      Height          =   1215
      Left            =   8400
      TabIndex        =   15
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton optReimp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "N. seguimiento"
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
         Index           =   1
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optReimp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Referencia"
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame fraMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   1200
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   11895
      Begin VB.Timer tmrProg 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   8520
         Top             =   480
      End
      Begin MSComctlLib.ProgressBar proEsp 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Espere un momento..."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   11415
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdReg 
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   10821
      _Version        =   393216
      BackColorSel    =   -2147483645
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483643
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      Begin HSFlatControls.MyCombo cboTerm 
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   "MyCombo1"
         Sorted          =   0   'False
         List            =   $"frmRegTerm.frx":0000
         ItemData        =   $"frmRegTerm.frx":001B
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
      Begin MSMask.MaskEdBox mskFin 
         Height          =   375
         Left            =   6000
         TabIndex        =   2
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hasta"
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
         Left            =   5280
         TabIndex        =   9
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Desde"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fecha"
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
         Left            =   480
         TabIndex        =   7
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Terminal"
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
         Left            =   480
         TabIndex        =   6
         Top             =   300
         Width           =   1455
      End
   End
   Begin MyCommandButton.MyButton cmdReImpR 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   9050
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   873
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
      Caption         =   "Reimprimir"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmRegTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents ws As WebSocketWrap.Client
Attribute ws.VB_VarHelpID = -1
Dim blnRespuestaEsperar As Boolean
Dim strRespuestaEsperar As String
Dim blnChange As Boolean
Dim strTmpFecha As String
Dim intTimeout As Integer

Private Sub cboTerm_Click()
    pCargaRegistro
End Sub

Private Function fstrGetPPData(arrDatos() As String, strNombre As String) As String
    On Error GoTo Errs
    Dim intIndex As Integer
    For intIndex = 0 To UBound(arrDatos)
         If strNombre = Split(arrDatos(intIndex), "=")(0) Then
             fstrGetPPData = Replace(Split(arrDatos(intIndex), "=")(1), "_", " ")
             Exit Function
         End If
    Next
Errs:
    fstrGetPPData = ""
End Function

Private Sub pReimprimir(strTipo As String, Modo As Integer, Mensaje As String)
    On Error GoTo Errs
    Dim strResp As String
    Dim tmpData As String
    Dim arrDatos() As String
        
    Dim vgrptReporte As CRAXDRT.Report
    Dim rsDummy As ADODB.Recordset
    Dim alstrParametros(22) As String

    If grdReg.TextMatrix(grdReg.Row, 3) = "VENTA" Or grdReg.TextMatrix(grdReg.Row, 3) = "REIMPRESION REF" Or grdReg.TextMatrix(grdReg.Row, 3) = "REIMPRESION SEG" Or grdReg.TextMatrix(grdReg.Row, 3) = "INIPAGO" Then
        If grdReg.TextMatrix(grdReg.Row, 7) <> "" Then
            pMensajeEsp True
            If Modo = 2 Then
            
            strResp = Mensaje
            Else
            strResp = fstrPinPad(grdReg.TextMatrix(grdReg.Row, 11), grdReg.TextMatrix(grdReg.Row, 8), grdReg.TextMatrix(grdReg.Row, 1), IIf(strTipo = "R", grdReg.TextMatrix(grdReg.Row, 7), grdReg.TextMatrix(grdReg.Row, 15)), grdReg.TextMatrix(grdReg.Row, 12), grdReg.TextMatrix(grdReg.Row, 13), grdReg.TextMatrix(grdReg.Row, 9), grdReg.TextMatrix(grdReg.Row, 10), strTipo, CLng(grdReg.TextMatrix(grdReg.Row, 16)), CLng(grdReg.TextMatrix(grdReg.Row, 17)))
     End If
            If strResp <> "" And grdReg.TextMatrix(grdReg.Row, 14) <> "" Then
                If Modo = 2 Then
                
                arrDatos = Split(strResp, "|")
               Set rsDummy = frsRegresaRs("select sysdate from dual", adLockReadOnly, adOpenForwardOnly)
                 alstrParametros(0) = "Firma;1"
   alstrParametros(1) = "NombreComercio;" & fstrGetPPData(arrDatos, "cadena4")
        alstrParametros(2) = "Mensaje1;" & fstrGetPPData(arrDatos, "cadena5")
        alstrParametros(3) = "Mensaje2;" & fstrGetPPData(arrDatos, "cadena6")
        alstrParametros(4) = "Mensaje3;" & fstrGetPPData(arrDatos, "cadena7")
        alstrParametros(5) = "trn_label;" & fstrGetPPData(arrDatos, "cadena9")
        alstrParametros(6) = "PANTarjeta;" & fstrGetPPData(arrDatos, "cadena13")
        alstrParametros(7) = "Copia;2"
        alstrParametros(8) = "MerchID;" & fstrGetPPData(arrDatos, "cadena15")
        alstrParametros(9) = "tipoVocher;" & fstrGetPPData(arrDatos, "cadena17")
        alstrParametros(10) = "Total;" & fstrGetPPData(arrDatos, "cadena23")
        alstrParametros(11) = "Operacion;" & fstrGetPPData(arrDatos, "cadena24")
        alstrParametros(12) = "REF;" & fstrGetPPData(arrDatos, "cadena25")
        alstrParametros(13) = "ARQC;" & fstrGetPPData(arrDatos, "cadena26")
        alstrParametros(14) = "AID;" & fstrGetPPData(arrDatos, "cadena27")
        alstrParametros(15) = "trn_aprnam;" & fstrGetPPData(arrDatos, "cadena28")
        alstrParametros(16) = "Total;" & fstrGetPPData(arrDatos, "trn_amount")
        alstrParametros(18) = "Fecha;" & fstrGetPPData(arrDatos, "cadena33")
        alstrParametros(19) = "NombreCliente;" & fstrGetPPData(arrDatos, "cadena40")
         alstrParametros(20) = "BarrasCode;" & fstrGetPPData(arrDatos, "cadena48")
          alstrParametros(21) = "Autorizacion;" & fstrGetFirma(fstrGetPPData(arrDatos, "cadena39"))
    pInstanciaReporte vgrptReporte, "voucherSantanderCopia.rpt"
   
                    vgrptReporte.DiscardSavedData
                    fblnAsignaImpresoraReportePorNombre grdReg.TextMatrix(grdReg.Row, 14), vgrptReporte
                    pCargaParameterFields alstrParametros, vgrptReporte
                    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
                    
                    alstrParametros(7) = "Copia;1"
                     alstrParametros(9) = "tipoVocher;" & "-C-L-I-E-N-T-E-"
                    pCargaParameterFields alstrParametros, vgrptReporte
                    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
                
                Else
                
                arrDatos = Split(strResp, "|")
                tmpData = fstrGetPPData(arrDatos, "trn_internal_respcode")
                If tmpData = "-1" Then
                    Set rsDummy = frsRegresaRs("select sysdate from dual", adLockReadOnly, adOpenForwardOnly)
                    tmpData = fstrGetPPData(arrDatos, "trn_qty_pay")
                    'alstrParametros(0) = "Operacion;REIMPRESION VENTA"
                   ' alstrParametros(1) = "mer_legend1;" & fstrGetPPData(arrDatos, "mer_legend1")
                    'alstrParametros(2) = "mer_legend2;" & fstrGetPPData(arrDatos, "mer_legend2")
                    'alstrParametros(3) = "mer_legend3;" & fstrGetPPData(arrDatos, "mer_legend3")
                    'alstrParametros(4) = "trn_external_mer_id;" & fstrGetPPData(arrDatos, "trn_external_mer_id")
                    'alstrParametros(5) = "trn_external_ter_id;" & fstrGetPPData(arrDatos, "trn_external_ter_id")
                    'alstrParametros(6) = "trn_fechaTrans;" & fstrGetPPData(arrDatos, "trn_fechaTrans")
                    'alstrParametros(7) = "Copia;0"
                   ' alstrParametros(8) = "trn_label;" & fstrGetPPData(arrDatos, "trn_label")
                    'alstrParametros(9) = "trn_aprnam;" & fstrGetPPData(arrDatos, "trn_aprnam")
                    'alstrParametros(10) = "trn_emv_cryptogram;" & fstrGetPPData(arrDatos, "trn_emv_cryptogram")
                   ' alstrParametros(11) = "trn_AID;" & fstrGetPPData(arrDatos, "trn_AID")
                    'alstrParametros(12) = "trn_pro_name;" & fstrGetPPData(arrDatos, "trn_pro_name")
                   ' alstrParametros(13) = "trn_aco_id;" & fstrGetPPData(arrDatos, "trn_aco_id")
                   'alstrParametros(14) = "trn_auth_code;" & fstrGetPPData(arrDatos, "trn_auth_code")
                   ' alstrParametros(15) = "trn_id;" & fstrGetPPData(arrDatos, "trn_id")
                   ' alstrParametros(16) = "trn_amount;" & fstrGetPPData(arrDatos, "trn_amount")
                    'alstrParametros(17) = "Compra;" & IIf(tmpData = "1", "COMPRA NORMAL", tmpData & " MESES SIN INTERESES")
                   ' alstrParametros(18) = "trn_fe;" & fstrGetPPData(arrDatos, "trn_fe")
                    'alstrParametros(19) = "trn_internal_ter_id;" & fstrGetPPData(arrDatos, "trn_internal_ter_id")
                    
                    alstrParametros(0) = "Firma;1"
        alstrParametros(1) = "NombreComercio;" & fstrGetPPData(arrDatos, "mer_legend1")
        alstrParametros(2) = "Mensaje1;" & fstrGetPPData(arrDatos, "mer_legend2")
        alstrParametros(3) = "Mensaje2;" & fstrGetPPData(arrDatos, "mer_legend3")
        alstrParametros(4) = "Afiliacion;" & fstrGetPPData(arrDatos, "trn_external_mer_id")
        alstrParametros(5) = "TerminalID;" & fstrGetPPData(arrDatos, "trn_external_ter_id")
        alstrParametros(6) = "Fecha;" & fstrGetPPData(arrDatos, "trn_fechaTrans")
        alstrParametros(7) = "Copia;0"
        'alstrParametros(8) = "trn_label;" & fstrGetPPData(arrDatos, "trn_label")
        'alstrParametros(9) = "trn_aprnam;" & fstrGetPPData(arrDatos, "trn_aprnam")
       ' alstrParametros(10) = "ARQC;" & fstrGetPPData(arrDatos, "trn_emv_cryptogram")
        'alstrParametros(11) = "AID;" & fstrGetPPData(arrDatos, "trn_AID")
        alstrParametros(12) = "Emisor;" & fstrGetPPData(arrDatos, "trn_pro_name")
        alstrParametros(13) = "PANTarjeta;" & fstrGetPPData(arrDatos, "trn_aco_id")
        alstrParametros(14) = "Autorizacion;" & fstrGetPPData(arrDatos, "trn_auth_code")
        alstrParametros(15) = "TC;" & fstrGetPPData(arrDatos, "trn_id")
        alstrParametros(16) = "Total;" & fstrGetPPData(arrDatos, "trn_amount")
        'alstrParametros(17) = "Operacion;" & IIf(tmpData = "1", "COMPRA NORMAL", tmpData & " MESES SIN INTERESES")
        alstrParametros(18) = "trn_fe;" & fstrGetPPData(arrDatos, "trn_fe")
        alstrParametros(19) = "NombreCliente;" & fstrGetPPData(arrDatos, "trn_internal_ter_id")
   
                    pInstanciaReporte vgrptReporte, "voucherR.rpt"
   
                    vgrptReporte.DiscardSavedData
                    fblnAsignaImpresoraReportePorNombre grdReg.TextMatrix(grdReg.Row, 14), vgrptReporte
                    pCargaParameterFields alstrParametros, vgrptReporte
                    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
                    
                    alstrParametros(7) = "Copia;1"
                    pCargaParameterFields alstrParametros, vgrptReporte
                    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
                Else
                    tmpData = fstrGetPPData(arrDatos, "trn_msg_host")
                    MsgBox tmpData, vbExclamation, "Mensaje"
                End If
            
            End If
            pMensajeEsp False
        End If
    End If
   End If
   
    Exit Sub
Errs:
     MsgBox Err.Description, vbExclamation, "Mensaje"
End Sub

Private Sub cmdReImpR_Click()
    If optReimp(1).Value Then
        If grdReg.TextMatrix(grdReg.Row, 15) <> "" Then
            pReimprimir "S", 2, txtMensaje
            pCargaRegistro
        Else
            MsgBox "No se cuenta con el número de seguimiento, intente reimprimir por referencia", vbExclamation, "Mensaje"
        End If
    Else
        If grdReg.TextMatrix(grdReg.Row, 7) <> "" Then
            pReimprimir "R", 2, txtMensaje
            pCargaRegistro
        Else
            MsgBox "No se cuenta con la referencia, intente reimprimir por número de seguimiento", vbExclamation, "Mensaje"
        End If
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    Dim vllngNumeroOpcion As Long
    blnChange = False
    Me.Icon = frmMenuPrincipal.Icon
    mskIni.Mask = ""
    mskIni.Text = fdtmServerFecha
    mskIni.Mask = "##/##/####"
    mskFin.Mask = ""
    mskFin.Text = fdtmServerFecha
    mskFin.Mask = "##/##/####"
    
    Set rsTemp = frsSelParametros("SI", vgintClaveEmpresaContable, "INTTERMINALSTIMEOUT")
    If Not rsTemp.EOF Then
        intTimeout = CInt(rsTemp!Valor)
    Else
        intTimeout = 300
    End If
    rsTemp.Close
    If cgstrModulo = "PV" Then
        vllngNumeroOpcion = 7041
        
    ElseIf cgstrModulo = "SI" Then
        vllngNumeroOpcion = 4164
       
    End If
    
    pCargaTerminales
    cboTerm.ListIndex = 0
    'pConfiguraGrid
End Sub
Private Function fstrGetFirma(strNombre As String) As String
   
  
If strNombre = "AutorizadosinFirma" Then
             fstrGetFirma = "AUTORIZADOS SIN FIRMA"
Else
    If strNombre = "VALIDADOCONFIRMAELECTRONICA" Then
              fstrGetFirma = "VALIDADO CON FIRMA ELECTRONICA"
              Else
              fstrGetFirma = "*"
End If
End If


End Function
Private Sub grdReg_GotFocus()
    grdReg.BackColorSel = &H8000000D
    grdReg.ForeColorSel = &H8000000E
    grdReg.Col = 1
    grdReg.ColSel = 15
    
End Sub

Private Sub grdReg_LostFocus()
     grdReg.BackColorSel = &H80000003
     grdReg.ForeColorSel = &H80000008
End Sub

Private Sub grdReg_RowColChange()
    txtMensaje.Text = grdReg.TextMatrix(grdReg.Row, 0)
End Sub

Private Sub mskFin_GotFocus()
    strTmpFecha = mskFin.Text
    pSelMkTexto mskFin
End Sub

Private Sub mskFin_LostFocus()
    If Not IsDate(mskFin.Text) Then
        mskFin.Mask = ""
        mskFin.Text = fdtmServerFecha
        mskFin.Mask = "##/##/####"
    End If
    blnChange = strTmpFecha <> mskFin.Text
    If blnChange Then
    If CDate(mskIni.Text) < CDate(mskFin.Text) Then
     pCargaRegistro
    Else
    MsgBox "Error la fecha fin debe de ser mayor a la fecha fin", vbExclamation, "Mensaje"
    mskFin.Text = fdtmServerFecha
    End If
    
    pCargaRegistro
    
    End If
    
End Sub

Private Sub mskIni_GotFocus()
    strTmpFecha = mskIni.Text
    pSelMkTexto mskIni
End Sub

Private Sub mskIni_LostFocus()
    If Not IsDate(mskIni.Text) Then
        mskIni.Mask = ""
        mskIni.Text = fdtmServerFecha
        mskIni.Mask = "##/##/####"
    End If
    blnChange = strTmpFecha <> mskIni.Text
    If blnChange Then
    
    If CDate(mskIni.Text) < CDate(mskFin.Text) Then
     pCargaRegistro
    Else
    MsgBox "¡Rango de fechas no válido!", vbExclamation, "Mensaje"
    mskFin.Text = fdtmServerFecha
    End If
    
  
    
    End If
    
End Sub

Private Sub pCargaTerminales()
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select distinct pvterminal.INTCVETERMINAL, pvterminal.VCHNOMBRE from pvformapago inner join pvterminal on pvformapago.INTCVETERMINAL = pvterminal.INTCVETERMINAL order by VCHNOMBRE", adLockReadOnly, adOpenForwardOnly)
    Do Until rs.EOF
        cboTerm.AddItem rs!vchNombre
        cboTerm.ItemData(cboTerm.newIndex) = rs!intCveTerminal
        rs.MoveNext
    Loop
    rs.Close
End Sub


Private Sub pCargaRegistro()
    Dim rs As ADODB.Recordset
    Dim intCveTerminal As Integer
    Dim strFechaIni As String
    Dim strFechaFin As String
    Dim strtmp() As String
    Dim strTransaccion As String
    Dim strMonto As String
    Dim strReferencia As String
    Dim strIdCurr As String
    Dim arrDatos() As String
    Dim tmpData As String
    pConfiguraGrid
    txtMensaje.Text = ""
    intCveTerminal = cboTerm.ItemData(cboTerm.ListIndex)
    strFechaIni = Format(CDate(mskIni.Text), "yyyy-mm-dd")
    strFechaFin = Format(CDate(mskFin.Text), "yyyy-mm-dd")
    
    If intCveTerminal > 0 Then
        Set rs = frsEjecuta_SP(CStr(intCveTerminal) & "|" & strFechaIni & "|" & strFechaFin, "sp_pvSelRegTerm")
        Do Until rs.EOF
        
            strtmp = Split(rs!VCHMESSAGETX, ":")
            If strtmp(0) = "T060S000" Then
                strTransaccion = "VENTA"
                strMonto = FormatCurrency(strtmp(2))
                strIdCurr = strtmp(1)
                If UBound(strtmp) >= 3 Then
                    strReferencia = strtmp(3)
                End If
                
            Else
             If strtmp(0) = "INIPAGO" Then
              strTransaccion = "VENTA"
             strMonto = FormatCurrency(strtmp(2))
             strIdCurr = strtmp(1)
              If UBound(strtmp) >= 3 Then
                    strReferencia = strtmp(3)
                End If
             End If
                If strtmp(0) = "PRINT060" Then
                    strTransaccion = "REIMPRESION REF"
                    strMonto = ""
                    strIdCurr = strtmp(1)
                    If UBound(strtmp) >= 2 Then
                        strReferencia = strtmp(2)
                    End If
                Else
                    If strtmp(0) = "PRINT" Then
                        strTransaccion = "REIMPRESION SEG"
                        strMonto = ""
                        strIdCurr = strtmp(1)
                        strReferencia = ""
                    Else
                        strTransaccion = strtmp(0)
                    End If
                End If
            End If
            grdReg.TextMatrix(grdReg.Rows - 1, 1) = rs!INTPROVIDER
            grdReg.TextMatrix(grdReg.Rows - 1, 2) = rs!DTMDATETX
            grdReg.TextMatrix(grdReg.Rows - 1, 3) = strTransaccion
            grdReg.TextMatrix(grdReg.Rows - 1, 4) = rs!formapago
            grdReg.TextMatrix(grdReg.Rows - 1, 5) = strMonto
            grdReg.TextMatrix(grdReg.Rows - 1, 6) = IIf(rs!BITPESOS = 0, "USD", "MXN")
            grdReg.TextMatrix(grdReg.Rows - 1, 7) = strReferencia
            grdReg.TextMatrix(grdReg.Rows - 1, 8) = strIdCurr
            grdReg.TextMatrix(grdReg.Rows - 1, 9) = rs!VCHUSR
            grdReg.TextMatrix(grdReg.Rows - 1, 10) = rs!VCHPWD
            grdReg.TextMatrix(grdReg.Rows - 1, 11) = rs!VCHURI
            grdReg.TextMatrix(grdReg.Rows - 1, 12) = rs!VCHIP
            grdReg.TextMatrix(grdReg.Rows - 1, 13) = rs!VCHPORT
            grdReg.TextMatrix(grdReg.Rows - 1, 14) = IIf(IsNull(rs!VCHIMPRESORAVOUCHER), "", rs!VCHIMPRESORAVOUCHER)
            If Not IsNull(rs!VCHMESSAGERX) Then
                grdReg.TextMatrix(grdReg.Rows - 1, 0) = rs!VCHMESSAGERX
                arrDatos = Split(rs!VCHMESSAGETX, "|")
                tmpData = fstrGetPPData(arrDatos, "trn_internal_respcode")
                If tmpData = "-1" Then
                    tmpData = fstrGetPPData(arrDatos, "trn_id")
                    grdReg.TextMatrix(grdReg.Rows - 1, 15) = tmpData
                End If
            End If
            grdReg.TextMatrix(grdReg.Rows - 1, 16) = rs!intCveTerminal
            grdReg.TextMatrix(grdReg.Rows - 1, 17) = rs!intFormaPago
            rs.MoveNext
            
            If Not rs.EOF Then
                grdReg.AddItem ""
            End If
        Loop
        grdReg.ColSel = 15
        grdReg_RowColChange
        rs.Close
    End If

End Sub

Private Sub pConfiguraGrid()
    grdReg.Cols = 18
    grdReg.Rows = 2
    grdReg.Clear
    
    grdReg.ColWidth(0) = 100
    grdReg.ColWidth(1) = 0  'Provider
    grdReg.ColWidth(2) = 2480
    grdReg.ColWidth(3) = 1820
    grdReg.ColWidth(4) = 3000
    grdReg.ColWidth(5) = 1820
    grdReg.ColWidth(6) = 940
    grdReg.ColWidth(7) = 2420  'Referencia
    grdReg.ColWidth(8) = 0  'Curr ID
    grdReg.ColWidth(9) = 0  'Usr
    grdReg.ColWidth(10) = 0 'Pwd
    grdReg.ColWidth(11) = 0 'URI
    grdReg.ColWidth(12) = 0 'IP
    grdReg.ColWidth(13) = 0 'Port
    grdReg.ColWidth(14) = 0 'Impresora
    grdReg.ColWidth(15) = 1580 'Num Seguimiento
    grdReg.ColWidth(16) = 0 'Clave terminal
    grdReg.ColWidth(17) = 0 'Clave forma pago
    grdReg.ColAlignment(5) = flexAlignRightCenter
    grdReg.ColAlignment(15) = flexAlignLeftCenter
    grdReg.ColAlignmentFixed(5) = flexAlignRightCenter
    grdReg.TextMatrix(0, 2) = "Fecha"
    grdReg.TextMatrix(0, 3) = "Transacción"
    grdReg.TextMatrix(0, 4) = "Forma de pago"
    grdReg.TextMatrix(0, 5) = "Monto"
    grdReg.TextMatrix(0, 6) = "Moneda"
    grdReg.TextMatrix(0, 7) = "Referencia"
    grdReg.TextMatrix(0, 15) = "N. seguimiento"
End Sub


Private Function fstrPinPad(strUriPinpad As String, strMoneda As String, intPpProvider As Integer, strReferencia As String, strHost As String, strPort As String, strUsr As String, strPwd As String, strTipoReImp As String, lngCve As Long, lngCveFormaPago As Long) As String
    On Error GoTo Errs
    Dim intRespLen As Long
    Dim strReturn As String
    Dim lngIdLog As Long
    Set ws = New WebSocketWrap.Client
    ws.Timeout = intTimeout
    ws.Uri = strUriPinpad & "?host=" & strHost & "&port=" & strPort & "&prov=" & IIf(intPpProvider = 1, "FISERV", "EVO") & "&usr=" & strUsr & "&pwd=" & strPwd
   
    blnRespuestaEsperar = False
    lngIdLog = pRegLog(0, "", lngCve, "N", lngCveFormaPago)
    Select Case intPpProvider
       
        Case 2
            If strTipoReImp = "R" Then
                lngIdLog = pRegLog(lngIdLog, "PRINT060:" & strMoneda & ":" & strReferencia, lngCve, "T", lngCveFormaPago)
                ws.SendMessage "PRINT060:" & strMoneda & ":" & strReferencia
            Else
                lngIdLog = pRegLog(lngIdLog, "PRINT:" & strMoneda & ":" & strReferencia, lngCve, "T", lngCveFormaPago)
                ws.SendMessage "PRINT:" & strMoneda & ":" & strReferencia
            End If
        
    End Select
    
    Do While Not blnRespuestaEsperar
        DoEvents
    Loop
    strReturn = strRespuestaEsperar
    If intPpProvider = 1 Then
        intRespLen = CInt(Mid(strReturn, 13, 4)) - 1
        strReturn = Replace(strReturn, ":", "=")
        strReturn = Mid(strReturn, 18, intRespLen)
    ElseIf intPpProvider = 2 Then
        intRespLen = InStr(strReturn, "}") - 14
        strReturn = Mid(strReturn, 13, intRespLen)
        strReturn = Replace(strReturn, "|Respuesta=", "")
        strReturn = Replace(strReturn, "&", "|")
    End If
    pRegLog lngIdLog, strReturn, lngCve, "R", lngCveFormaPago
    If InStr(strReturn, "Error de socket") > 0 Then
        MsgBox "Error de conexión con el socket:" & vbCrLf & strHost & ":" & strPort, vbExclamation, "Mensaje"
        fstrPinPad = ""
    Else
        fstrPinPad = strReturn
    End If
    
    Exit Function
Errs:
    fstrPinPad = ""
    If InStr(Err.Description, "Error de conexión") > 0 Then
        pRegLog lngIdLog, "Error de conexión con el Web Socket: " & strUriPinpad, lngCve, "R", lngCveFormaPago
        MsgBox "Error de conexión con el Web Socket: " & vbCrLf & strUriPinpad, vbExclamation, "Mensaje"
    Else
        pRegLog lngIdLog, Err.Description, lngCve, "R", lngCveFormaPago
        MsgBox Err.Description, vbExclamation, "Mensaje"
    End If
End Function

Private Sub pMensajeEsp(blnShow As Boolean)
    proEsp.Value = 0
    grdReg.SetFocus
    cmdReImpR.Enabled = Not blnShow
    fraMsg.Visible = blnShow
    tmrProg.Enabled = blnShow
End Sub

Private Sub tmrProg_Timer()
    If proEsp.Value < 100 Then
        proEsp.Value = proEsp.Value + 1
    End If
End Sub

Private Sub ws_Answer(answ As String)
    strRespuestaEsperar = answ
    blnRespuestaEsperar = True

End Sub

Public Function pRegLog(lngId, strMessage, intCveTerminal, strTipo As String, lngCveFormaPago As Long) As Long
     Dim rs As ADODB.Recordset
     pRegLog = 0
     Set rs = frsRegresaRs("select * from PVTerminalLog where intID = " & lngId, adLockOptimistic, adOpenStatic)
     If rs.EOF Then
        rs.AddNew
        rs!intCveTerminal = intCveTerminal
        rs!intCveFormaPago = lngCveFormaPago
        rs.Update
        pRegLog = flngObtieneIdentity("SEC_PVTERMINALLOG", 1)
     Else
        If strTipo = "T" Then
            rs!VCHMESSAGETX = strMessage
            rs!DTMDATETX = Now
        Else
            rs!VCHMESSAGERX = strMessage
            rs!dtmDateRX = Now
        End If
        pRegLog = lngId
        rs.Update
     End If
     rs.Close
End Function

