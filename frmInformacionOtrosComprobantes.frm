VERSION 5.00
Begin VB.Form frmInformacionOtrosComprobantes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInformacion 
      Height          =   2855
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdGuardarInformacion 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   2070
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmInformacionOtrosComprobantes.frx":0000
         TabIndex        =   8
         ToolTipText     =   "Confirmar información"
         Top             =   2170
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtTaxID 
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Identificador del contribuyente extranjero"
         Top             =   600
         Width           =   3560
      End
      Begin VB.TextBox txtNumFact 
         Height          =   315
         Left            =   1320
         MaxLength       =   36
         TabIndex        =   2
         ToolTipText     =   "Clave del comprobante"
         Top             =   240
         Width           =   3555
      End
      Begin VB.TextBox txtSerie 
         Height          =   315
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Serie del comprobante"
         Top             =   240
         Width           =   3560
      End
      Begin VB.TextBox txtNumFol 
         Height          =   315
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Número de folio del comprobante"
         Top             =   600
         Width           =   3560
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   4
         ToolTipText     =   "Monto total del comprobante"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1320
         TabIndex        =   10
         Top             =   1380
         Width           =   1643
         Begin VB.OptionButton optMoneda 
            Caption         =   "Dólares"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   6
            ToolTipText     =   "Comprobante emitido en dólares"
            Top             =   0
            Width           =   833
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Pesos"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Comprobante emitido en pesos"
            Top             =   0
            Value           =   -1  'True
            Width           =   728
         End
      End
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   7
         ToolTipText     =   "Tipo de cambio del comprobante"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblSerieONumFact 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   195
         Left            =   140
         TabIndex        =   15
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblNumFolOTaxID 
         AutoSize        =   -1  'True
         Caption         =   "Folio"
         Height          =   195
         Left            =   140
         TabIndex        =   14
         Top             =   660
         Width           =   330
      End
      Begin VB.Label lblMonto 
         AutoSize        =   -1  'True
         Caption         =   "Monto total"
         Height          =   195
         Left            =   140
         TabIndex        =   13
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label lblTipoCambio 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cambio"
         Height          =   195
         Left            =   140
         TabIndex        =   12
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label lblMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   140
         TabIndex        =   11
         Top             =   1388
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmInformacionOtrosComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------
' Forma para ingresar los datos de los comprobantes de tipo CBB o Extranjeros
'-----------------------------------------------------------------------------------------
Public lblnConsultaInfo As Boolean
Public vlintTipoXML As Integer '3 = CBB, 4 = Extranjero
Public vlintMontoTotal As Double
Public vldblTipoCambio As Double
Public vlblnPesos As Boolean
Dim vldtmFechaServer As Date


Private Sub cmdGuardarInformacion_Click()
    On Error GoTo NotificaError
    
    If vlintTipoXML = 3 Then
        If Trim(txtSerie.Text) = "" Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            txtSerie.SetFocus
            Exit Sub
        End If
        
        If Trim(txtNumFol.Text) = "" Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            txtNumFol.SetFocus
            Exit Sub
        Else
            If Val(txtNumFol.Text) = 0 Then
                '¡No ha ingresado datos!
                MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
                txtNumFol.SetFocus
                Exit Sub
            End If
        End If
    Else
        If Trim(txtNumFact.Text) = "" Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            txtNumFact.SetFocus
            Exit Sub
        End If
        
        If Trim(txtTaxID.Text) = "" Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            txtTaxID.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtMonto.Text) = "" Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtMonto.SetFocus
        Exit Sub
    Else
        If CDbl(txtMonto.Text) = 0 Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            txtMonto.SetFocus
            Exit Sub
        End If
    End If
    
    If optMoneda(1).Value Then
        If Trim(txtTipoCambio.Text) = "" Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            txtTipoCambio.SetFocus
            Exit Sub
        Else
            If CDbl(txtTipoCambio.Text) = 0 Then
                '¡No ha ingresado datos!
                MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
                txtTipoCambio.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    vgintTipoXMLCXP = vlintTipoXML
    vgstrUUIDXMLCXP = ""
    vgstrXMLCXP = ""
    
    If vlintTipoXML = 3 Then
        vgstrSerieCXP = Trim(txtSerie.Text)
        vgstrNumFolioCXP = Trim(txtNumFol.Text)
        vgstrNumFactExtCXP = ""
        vgstrTaxIDExtCXP = ""
    Else
        vgstrNumFactExtCXP = Trim(txtNumFact.Text)
        vgstrTaxIDExtCXP = Trim(txtTaxID.Text)
        vgstrSerieCXP = ""
        vgstrNumFolioCXP = ""
    End If
    
    vgdblMontoXMLCXP = CDbl(Trim(txtMonto.Text))
    vgstrMonedaXMLCXP = IIf(optMoneda(0).Value, "MXN", "USD")
    vgdblTipoCambioXMLCXP = CDbl(IIf(Trim(txtTipoCambio.Text) = "", "1", Trim(txtTipoCambio.Text)))
    
    Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGuardarInformacion_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = 27 Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
        frmInformacionOtrosComprobantes.Caption = IIf(vlintTipoXML = 3, "Información del código de barras bidimensional", "Información del comprobante extranjero")
        lblSerieONumFact.Caption = IIf(vlintTipoXML = 3, "Serie", "Clave")
        lblNumFolOTaxID.Caption = IIf(vlintTipoXML = 3, "Folio", "Tax ID")
            
        txtSerie.Visible = IIf(vlintTipoXML = 3, True, False)
        txtNumFact.Visible = IIf(vlintTipoXML = 4, True, False)
        txtNumFol.Visible = IIf(vlintTipoXML = 3, True, False)
        txtTaxID.Visible = IIf(vlintTipoXML = 4, True, False)
        
        vldtmFechaServer = fdtmServerFecha
    If lblnConsultaInfo = True Then
        cmdGuardarInformacion.Visible = False
        fraInformacion.Enabled = False
        If vlintTipoXML = 3 Then 'CBB
            txtSerie.Text = vgstrSerieCXP
            txtNumFol.Text = vgstrNumFolioCXP
        Else 'Comprobante extranjero
            txtNumFact.Text = vgstrNumFactExtCXP
            txtTaxID.Text = vgstrTaxIDExtCXP
        End If
        txtMonto.Text = FormatCurrency(CDbl(Trim(vgdblMontoXMLCXP)), 2)
        If vgstrMonedaXMLCXP = "MXN" Then
           optMoneda(0).Value = True
           txtTipoCambio.Enabled = False
        Else
           optMoneda(1).Value = True
           txtTipoCambio.Text = FormatCurrency(CDbl(Trim(vgdblTipoCambioXMLCXP)), 2)
        End If
        lblnConsultaInfo = False
    Else
        
        If vgintTipoXMLCXP = 0 Or vgintTipoXMLCXP = 1 Or vgintTipoXMLCXP = 2 Or (vgintTipoXMLCXP = 3 And vlintTipoXML = 4) Or (vgintTipoXMLCXP = 4 And vlintTipoXML = 3) Then
            txtMonto.Text = FormatCurrency(vlintMontoTotal, 2)
            If vlblnPesos Then
                optMoneda(0).Value = True
            Else
                optMoneda(1).Value = True
            End If
            pSeleccionCalculoTipoCambio
        Else
            'Hay comprobante CBB ingresado por lo tanto carga la información
            If vgintTipoXMLCXP = 3 Then
                txtSerie.Text = Trim(vgstrSerieCXP)
                txtNumFol.Text = Trim(vgstrNumFolioCXP)
                txtMonto.Text = FormatCurrency(vgdblMontoXMLCXP, 2)
                If vgstrMonedaXMLCXP = "USD" Then
                    optMoneda(1).Value = True
                Else
                    optMoneda(0).Value = True
                End If
                If optMoneda(0).Value Then
                    txtTipoCambio.Text = ""
                    txtTipoCambio.Enabled = False
                    lblTipoCambio.Enabled = False
                Else
                    txtTipoCambio.Text = FormatCurrency(vgdblTipoCambioXMLCXP, 4)
                    txtTipoCambio.Enabled = True
                    lblTipoCambio.Enabled = True
                End If
            End If
            
            'Hay comprobante extranjero ingresado por lo tanto carga la información
            If vgintTipoXMLCXP = 4 Then
                txtNumFact.Text = Trim(vgstrNumFactExtCXP)
                txtTaxID.Text = Trim(vgstrTaxIDExtCXP)
                txtMonto.Text = FormatCurrency(vgdblMontoXMLCXP, 2)
                If vgstrMonedaXMLCXP = "USD" Then
                    optMoneda(1).Value = True
                Else
                    optMoneda(0).Value = True
                End If
                If optMoneda(0).Value Then
                    txtTipoCambio.Text = ""
                    txtTipoCambio.Enabled = False
                    lblTipoCambio.Enabled = False
                Else
                    txtTipoCambio.Text = FormatCurrency(vgdblTipoCambioXMLCXP, 4)
                    txtTipoCambio.Enabled = True
                    lblTipoCambio.Enabled = True
                End If
            End If
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub optMoneda_Click(Index As Integer)
    pSeleccionCalculoTipoCambio
End Sub

Private Sub optMoneda_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = 13 Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optMoneda_KeyDown"))
End Sub

Private Sub txtMonto_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtMonto
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMonto_GotFocus"))
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = 13 Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMonto_KeyDown"))
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtMonto)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMonto_KeyPress"))
End Sub

Private Sub txtMonto_LostFocus()
    On Error GoTo NotificaError

    txtMonto.Text = FormatCurrency(Val(Format(txtMonto.Text, cstrFormato)), 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMonto_LostFocus"))
End Sub

Private Sub txtNumFact_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtNumFact
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumFact_GotFocus"))
End Sub

Private Sub txtNumFact_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = 13 Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumFact_KeyDown"))
End Sub

Private Sub txtNumFol_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtNumFol
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumFol_GotFocus"))
End Sub

Private Sub txtNumFol_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = 13 Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumFol_KeyDown"))
End Sub

Private Sub txtNumFol_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtSerie_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtSerie
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtSerie_GotFocus"))
End Sub

Private Sub txtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = 13 Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtSerie_KeyDown"))
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtTaxID_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtTaxID
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtTaxID_GotFocus"))
End Sub

Private Sub txtTaxID_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = 13 Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtTaxID_KeyDown"))
End Sub

Private Sub txtTipoCambio_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtTipoCambio
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtTipoCambio_GotFocus"))
End Sub

Private Sub txtTipoCambio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = 13 Then SendKeys vbTab
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtTipoCambio_KeyDown"))
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtTipoCambio)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtTipoCambio_KeyPress"))
End Sub

Private Sub pSeleccionCalculoTipoCambio()
On Error GoTo NotificaError

    If optMoneda(0).Value Then
        txtTipoCambio.Text = ""
        txtTipoCambio.Enabled = False
        lblTipoCambio.Enabled = False
    Else
        If vldblTipoCambio = 0 Then
            txtTipoCambio.Text = FormatCurrency(fdblTipoCambio(vldtmFechaServer, "O"), 4)
        Else
            txtTipoCambio.Text = FormatCurrency(vldblTipoCambio, 4)
        End If
        txtTipoCambio.Enabled = True
        lblTipoCambio.Enabled = True
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSeleccionCalculoTipoCambio"))
End Sub

Private Sub txtTipoCambio_LostFocus()
    On Error GoTo NotificaError

    txtTipoCambio.Text = FormatCurrency(Val(Format(txtTipoCambio.Text, cstrFormato)), 4)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtTipoCambio_LostFocus"))
End Sub
