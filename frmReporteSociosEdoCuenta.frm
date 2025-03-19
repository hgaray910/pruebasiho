VERSION 5.00
Begin VB.Form frmReporteSociosEdoCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de cuenta"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraImprime 
      Height          =   735
      Left            =   3960
      TabIndex        =   22
      Top             =   2040
      Width           =   1155
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   580
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteSociosEdoCuenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprimir"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   80
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteSociosEdoCuenta.frx":03CD
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vista previa"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame fraCliente 
      Height          =   2040
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      Begin VB.TextBox txtClaveUnica 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         ToolTipText     =   "Número de socio"
         Top             =   195
         Width           =   1800
      End
      Begin VB.CheckBox chkBitExtranjero 
         Caption         =   "Extranjero"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3420
         TabIndex        =   3
         ToolTipText     =   "Extranjero"
         Top             =   570
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox txtRFC 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "RFC del socio"
         Top             =   540
         Width           =   1800
      End
      Begin VB.TextBox txtClaveSocio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         TabIndex        =   1
         ToolTipText     =   "Número de socio"
         Top             =   540
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label16 
         Caption         =   "Número interior"
         Height          =   255
         Left            =   2805
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Clave única"
         Height          =   195
         Left            =   225
         TabIndex        =   20
         Top             =   255
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Calle"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   945
         Width           =   345
      End
      Begin VB.Label lblSocio 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   315
         Left            =   3405
         TabIndex        =   18
         ToolTipText     =   "Nombre del socio"
         Top             =   195
         Width           =   5250
      End
      Begin VB.Label lblDomicilio 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1500
         TabIndex        =   17
         ToolTipText     =   "Calle del socio"
         Top             =   900
         Width           =   7155
      End
      Begin VB.Label lblCiudad 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5925
         TabIndex        =   16
         ToolTipText     =   "Ciudad del socio"
         Top             =   1605
         Width           =   2730
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   5265
         TabIndex        =   15
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "R. F. C."
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblTelefono 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4020
         TabIndex        =   13
         ToolTipText     =   "Teléfono del socio"
         Top             =   1605
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   2805
         TabIndex        =   12
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label17 
         Caption         =   "Número exterior"
         Height          =   255
         Left            =   225
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblNumeroExterior 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1500
         TabIndex        =   10
         ToolTipText     =   "Número exterior del socio"
         Top             =   1245
         Width           =   1185
      End
      Begin VB.Label lblNumeroInterior 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4020
         TabIndex        =   9
         ToolTipText     =   "Número interior del socio"
         Top             =   1245
         Width           =   1125
      End
      Begin VB.Label lblCP 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         ToolTipText     =   "Código postal del socio"
         Top             =   1605
         Width           =   1185
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Código postal"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label lblColonia 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5925
         TabIndex        =   6
         ToolTipText     =   "Número exterior del socio"
         Top             =   1245
         Width           =   2730
      End
      Begin VB.Label Label20 
         Caption         =   "Colonia"
         Height          =   255
         Left            =   5265
         TabIndex        =   5
         Top             =   1320
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmReporteSociosEdoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vgrptReporte As CRAXDRT.Report

Private Sub pImprime(vlstrTipo As String)

On Error GoTo NotificaError

    Dim rsSocios As New ADODB.Recordset
    Dim rsDatosSocio As New ADODB.Recordset
    Dim alstrParametros(6) As String
    Dim strParametros As String
    
    pInstanciaReporte vgrptReporte, "rptEdoCuentaSocios.rpt"

    vgrptReporte.DiscardSavedData

    Set rsDatosSocio = frsEjecuta_SP(IIf(txtClaveSocio.Text = "", "0", txtClaveSocio.Text), "SP_SOSELDATOSSOCIO")
    
    If rsDatosSocio.RecordCount > 0 Then
        Set rsSocios = frsEjecuta_SP(IIf(txtClaveSocio.Text = "", "0", txtClaveSocio.Text), "SP_PVRPTEDOCTASOCIOS")
    
        If rsSocios.RecordCount <> 0 Then
    
                alstrParametros(0) = "lblNombreHospital;" & Trim(vgstrNombreHospitalCH)
                alstrParametros(1) = "strNombreSocio;" & IIf(IsNull(rsDatosSocio!Nombre), "", rsDatosSocio!Nombre)
                alstrParametros(2) = "strDireccion;" & IIf(IsNull(rsDatosSocio!DOMICILIOCOMPLETO), "", rsDatosSocio!DOMICILIOCOMPLETO)
                alstrParametros(3) = "strCveSocio;" & IIf(IsNull(rsDatosSocio!VCHCLAVESOCIO), "", rsDatosSocio!VCHCLAVESOCIO)
                alstrParametros(4) = "strHispanidad;" & IIf(IsNull(rsDatosSocio!hispanidad), "", rsDatosSocio!hispanidad)
                alstrParametros(5) = "strCtaContable;" & IIf(IsNull(rsDatosSocio!ctacontable), "", rsDatosSocio!ctacontable)
                alstrParametros(6) = "strRegSBE;" & IIf(IsNull(rsDatosSocio!vchregistrosbe), "", rsDatosSocio!vchregistrosbe)
                pCargaParameterFields alstrParametros, vgrptReporte
    
                pCargaParameterFields alstrParametros, vgrptReporte
                pImprimeReporte vgrptReporte, rsSocios, IIf(vlstrTipo = "P", "P", "I"), "Estado de cuenta"
    
        Else
            'No existe información con esos parámetros
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
    
        rsSocios.Close
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    rsDatosSocio.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

Private Sub cmdPreview_Click()
On Error GoTo NotificaError
    
    pImprime "P"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPreview_Click"))
    Unload Me

End Sub

Private Sub cmdPrint_Click()
On Error GoTo NotificaError
    
    pImprime "I"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPreview_Click"))
    Unload Me

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    
    pLimpiaEncabezado
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '¿Desea abandonar la operación?
    If lblSocio.Caption <> "" Then
        Cancel = True
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            pLimpiaEncabezado
            txtClaveUnica.SetFocus
        End If
    End If
End Sub


Private Sub txtClaveUnica_Change()
    lblSocio.Caption = ""
    lblDomicilio.Caption = ""
    lblCiudad.Caption = ""
    txtRFC.Text = ""
    lblTelefono.Caption = ""
    chkBitExtranjero.Value = vbUnchecked
    lblNumeroExterior.Caption = ""
    lblNumeroInterior.Caption = ""
    lblColonia.Caption = ""
    lblCP.Caption = ""
End Sub

Private Sub txtClaveUnica_GotFocus()
    pSelTextBox txtClaveUnica

End Sub


Private Sub txtClaveUnica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtClaveUnica.Text) = "" Then
            With frmSociosBusqueda
                .vglngClaveSocio = 0
                .Show vbModal, Me
                If .vglngClaveSocio <> 0 Then
                    pLlenaInformacionSocio .vglngClaveSocio
                End If
                Unload frmSociosBusqueda
                'if fblnCanFocus(chkBitExtranjero) Then chkBitExtranjero.SetFocus
                cmdPreview.SetFocus
            End With
        Else
            pLlenaInformacionSocio flngObtieneClaveSocio(txtClaveUnica.Text)
        End If
    End If
End Sub


Private Sub txtClaveUnica_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Function flngObtieneClaveSocio(strClaveSocio As String) As Long
    Dim rsSocio As New ADODB.Recordset
    Dim strParametrosSP As String
    
    strParametrosSP = CStr(strClaveSocio) & "|-1"
    Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_PVSELSOCIOS")
    If rsSocio.RecordCount > 0 Then
        flngObtieneClaveSocio = rsSocio!intcvesocio
    Else
        flngObtieneClaveSocio = -1
    End If
    rsSocio.Close

End Function

Private Sub pLlenaInformacionSocio(lngCveSocio As Long)
    Dim strParametrosSP As String
    Dim rsSocio As New ADODB.Recordset
    Dim rsCargos As New ADODB.Recordset
    Dim intContador As Integer
    
    '|  Consulta la información del socio
    strParametrosSP = CStr(lngCveSocio) & "|T"
    Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_SORPTSELSOCIO")
    '|  Si existe la información del socio
    If rsSocio.RecordCount > 0 Then
        pLimpiaEncabezado
        With rsSocio
            txtClaveSocio.Text = lngCveSocio
            txtClaveUnica.Text = IIf(IsNull(!VCHCLAVESOCIO), "", !VCHCLAVESOCIO)
            lblSocio.Caption = IIf(IsNull(!Nombre), "", !Nombre)
            txtRFC.Text = IIf(IsNull(!RFC), "", !RFC)
            lblDomicilio.Caption = IIf(IsNull(!Domicilio), "", !Domicilio)
            lblNumeroExterior.Caption = IIf(IsNull(!NumeroExterior), "", !NumeroExterior)
            lblNumeroInterior.Caption = IIf(IsNull(!NumeroInterior), "", !NumeroInterior)
            lblCP.Caption = IIf(IsNull(!CP), "", !CP)
            lblTelefono.Caption = IIf(IsNull(!Telefono), "", !Telefono)
            lblColonia.Caption = IIf(IsNull(!Colonia), "", !Colonia)
            lblCiudad.Caption = IIf(IsNull(!DescripcionCiudad), "", !DescripcionCiudad)
            llngCveCiudad = IIf(IsNull(!ClaveCiudad), 0, !ClaveCiudad)
        End With
    Else
        '|  ¡No existe información!
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        pLimpiaEncabezado
    End If
End Sub
Private Sub pLimpiaEncabezado()
    
    fraCliente.Enabled = True
        
    txtClaveUnica.Text = ""
    txtClaveSocio.Text = ""
    
    chkBitExtranjero.Enabled = True
    'lblFolio.ForeColor = llngColorActivas
    

    chkBitExtranjero.Value = vbUnchecked
    
End Sub
