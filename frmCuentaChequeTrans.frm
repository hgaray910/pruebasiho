VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCuentaChequeTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta bancaria y RFC de pagos recibidos"
   ClientHeight    =   8775
   ClientLeft      =   3465
   ClientTop       =   2730
   ClientWidth     =   17460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   17460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBotonera 
      Height          =   705
      Left            =   8433
      TabIndex        =   10
      Top             =   7960
      Width           =   595
      Begin VB.CommandButton cmdGuardar 
         Enabled         =   0   'False
         Height          =   495
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCuentaChequeTrans.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Grabar"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame fraMaestro 
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8000
      Begin VB.ComboBox cboCorteTransfiere 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Selección del corte"
         Top             =   225
         Width           =   7280
      End
      Begin VB.Label lblCorteTransfiere 
         AutoSize        =   -1  'True
         Caption         =   "Corte "
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   278
         Width           =   420
      End
   End
   Begin VB.Frame fraDetalle 
      Height          =   7260
      Left            =   120
      TabIndex        =   2
      Top             =   675
      Width           =   17205
      Begin MSMask.MaskEdBox MskFecha 
         Height          =   225
         Left            =   12960
         TabIndex        =   9
         ToolTipText     =   "Fecha del cheque o transferencia con la cual se realizó el pago"
         Top             =   6840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtRFC 
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   14880
         MaxLength       =   13
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "RFC del emisor del cheque o transferencia"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtCuentaBancaria 
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   14880
         MaxLength       =   30
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cuenta bancaria del emisor del cheque o transferencia"
         Top             =   6240
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cboBancoSAT 
         Height          =   315
         Left            =   14280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Selección del banco origen del pago"
         Top             =   6720
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.ComboBox cboCuentasPrevias 
         Height          =   315
         Left            =   14280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Selección de la cuenta bancaria"
         Top             =   6360
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.TextBox txtBancoExtranjero 
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   12720
         MaxLength       =   150
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "BancoExtranjero"
         ToolTipText     =   "Banco extranjero origen del pago"
         Top             =   6120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBusqueda 
         Height          =   6885
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Movimientos del corte"
         Top             =   240
         Width           =   16960
         _ExtentX        =   29924
         _ExtentY        =   12144
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   315
         ForeColorSel    =   -2147483643
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         HighLight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmCuentaChequeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
' Fecha de inicio de desarrollo:    Septiembre, 2014
' Autor:                            Jesús Valles Torres
'****************************************************************************************

Dim vllngCorte As Long
Dim vllngClaveDepartamento As Long
Dim vlblnLectura As Boolean

'Columnas para el grid
    Const cIntColConsecutivo = 1        'Consecutivo del detalle del corte
    Const cIntColNumCorte = 2           'Numero de corte
    Const cIntColFechaDoc = 3           'Fecha y hora del documento
    Const cIntColTipoDoc = 4            'Tipo de documento emitido
    Const cIntColFolioDoc = 5           'Folio del documento emitido
    Const cIntColEstado = 6             'Estado del documento
    Const cIntColTipoPago = 7           'Tipo de pago recibido (Cheque, Transferencia, Efectivo, Tarjeta)
    Const cIntColFolio = 8              'Folio del pago recibido (Cheque, Transferencia, Efectivo, Tarjeta)
    Const cIntColRazonSocial = 9        'Razón social asociada al documento
    Const cIntColRFC = 10               'RFC asociada al documento
    Const cIntColClaveBancoSAT = 11     'Clave del banco del SAT
    Const cIntColDescBancoSAT = 12      'Descripción del banco del SAT
    Const cIntColCuentaBancaria = 13    'Cuenta bancaria del pago recibido (Cheque o transferencia)
    Const cIntColFechaPago = 14         'Fecha del pago recibido (Cheque o transferencia)
    Const cIntColModificación = 15      'Indica si se cambió la información

Option Explicit

Private Sub pCargaCorte()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset

    cboCorteTransfiere.Clear

    If cgstrModulo = "PV" Or cgstrModulo = "CC" Then
'        Set rs = frsRegresaRs("SELECT intnumcorte, dtmfechahora, dtmfecharegistro FROM PVCORTE WHERE chrtipo = 'P' and dtmfechahora >= TO_DATE('01/01/2014','DD/MM/YYYY') and smidepartamento = " & vgintNumeroDepartamento & " ORDER BY intnumcorte DESC", adLockReadOnly, adOpenForwardOnly)
        
        Set rs = frsRegresaRs("SELECT intnumcorte, dtmfechahora, dtmfecharegistro FROM PVCORTE WHERE dtmfechahora >= TO_DATE('01/01/2014','DD/MM/YYYY') and smidepartamento = " & vgintNumeroDepartamento & " ORDER BY intnumcorte DESC", adLockReadOnly, adOpenForwardOnly)
        fraMaestro.Width = 5265
        cboCorteTransfiere.Width = 4545
    Else
'        Set rs = frsRegresaRs("SELECT intnumcorte, dtmfechahora, dtmfecharegistro, vchdescripcion FROM PVCORTE INNER JOIN NODEPARTAMENTO ON PVCORTE.smidepartamento = NODEPARTAMENTO.smicvedepartamento WHERE PVCORTE.chrtipo = 'P' and dtmfechahora >= TO_DATE('01/01/2014','DD/MM/YYYY') and NODEPARTAMENTO.tnyclaveempresa = " & vgintClaveEmpresaContable & " ORDER BY intnumcorte DESC", adLockReadOnly, adOpenForwardOnly)
        
        Set rs = frsRegresaRs("SELECT intnumcorte, dtmfechahora, dtmfecharegistro, vchdescripcion FROM PVCORTE INNER JOIN NODEPARTAMENTO ON PVCORTE.smidepartamento = NODEPARTAMENTO.smicvedepartamento WHERE dtmfechahora >= TO_DATE('01/01/2014','DD/MM/YYYY') and NODEPARTAMENTO.tnyclaveempresa = " & vgintClaveEmpresaContable & " ORDER BY intnumcorte DESC", adLockReadOnly, adOpenForwardOnly)
        fraMaestro.Width = 8000
        cboCorteTransfiere.Width = 7280
    End If
        
    If rs.RecordCount <> 0 Then
        Do While Not rs.EOF
            If cgstrModulo = "PV" Or cgstrModulo = "CC" Then
                cboCorteTransfiere.AddItem CStr(rs!intnumcorte) & " - " & Format(rs!DTMFECHAHORA, "dd/mmm/yyyy hh:mm") & " - " & IIf(IsNull(rs!dtmFechaRegistro), "ABIERTO", "CERRADO")
            Else
                cboCorteTransfiere.AddItem CStr(rs!intnumcorte) & " - " & Format(rs!DTMFECHAHORA, "dd/mmm/yyyy hh:mm") & " - " & IIf(IsNull(rs!dtmFechaRegistro), "ABIERTO", "CERRADO") & " - " & Trim(rs!VCHDESCRIPCION)
            End If
            cboCorteTransfiere.ItemData(cboCorteTransfiere.newIndex) = rs!intnumcorte
            rs.MoveNext
        Loop
        cboCorteTransfiere.ListIndex = 0
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaCorte"))
End Sub

Private Sub cboBancoSAT_KeyPress(KeyAscii As Integer)
    Dim vlClaveBanco As String
On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If Trim(cboBancoSAT.Text) <> "<BANCO EXTRANJERO>" Then
            If Trim(cboBancoSAT.Text) <> "" Then
            
                If Len(Trim(Str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))) = 3 Then
                    vlClaveBanco = Trim(Str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                Else
                    If Len(Trim(Str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))) = 2 Then
                        vlClaveBanco = "0" & Trim(Str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                    Else
                        vlClaveBanco = "00" & Trim(Str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                    End If
                End If
                
                If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColClaveBancoSAT)) <> vlClaveBanco Then
                    grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColModificación) = "*"
                End If
            
                grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColClaveBancoSAT) = vlClaveBanco

                If InStr(1, cboBancoSAT.Text, " - ") <> 0 Then
                    If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColDescBancoSAT)) <> Trim(Left(cboBancoSAT.Text, InStr(1, cboBancoSAT.Text, " - ") - 1)) Then
                        grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColModificación) = "*"
                    End If
                
                    grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColDescBancoSAT) = Left(cboBancoSAT.Text, InStr(1, cboBancoSAT.Text, " - ") - 1)
                Else
                    If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColDescBancoSAT)) <> Trim(cboBancoSAT.Text) Then
                        grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColModificación) = "*"
                    End If
                
                    grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColDescBancoSAT) = Trim(cboBancoSAT.Text)
                End If
                
                pActualizaTrasCambio grdBusqueda.Row
                cmdGuardar.Enabled = True
            End If
            
            cboBancoSAT.Visible = False
            grdBusqueda.SetFocus
            
            If Trim(cboBancoSAT.Text) = "" Then
                If grdBusqueda.Rows - grdBusqueda.Row > 1 Then
                    grdBusqueda.Row = grdBusqueda.Row + 1
                    grdBusqueda.Col = cIntColRFC
    '                grdBusqueda.Col = cIntColDescBancoSAT
                End If
            Else
                grdBusqueda.Col = cIntColCuentaBancaria
            End If
        Else
            cboBancoSAT.Visible = False
            grdBusqueda.Col = cIntColDescBancoSAT
            
            txtBancoExtranjero.Text = ""
            txtBancoExtranjero.Top = grdBusqueda.Top + grdBusqueda.CellTop + 48
            txtBancoExtranjero.Left = grdBusqueda.Left + grdBusqueda.CellLeft + 24
            txtBancoExtranjero.Height = grdBusqueda.CellHeight - 56
            txtBancoExtranjero.Width = grdBusqueda.CellWidth - 40
            txtBancoExtranjero.Visible = True
            txtBancoExtranjero.SetFocus
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboBancoSAT_KeyPress"))
End Sub

Private Sub cboCorteTransfiere_GotFocus()
'    vllngCorte = cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)
End Sub

Private Sub cboCorteTransfiere_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboCuentasPrevias_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If cboCuentasPrevias.ItemData(cboCuentasPrevias.ListIndex) = 0 Then
            With grdBusqueda
                txtCuentaBancaria.Top = .Top + .CellTop + 48
                txtCuentaBancaria.Left = .Left + .CellLeft + 24
                txtCuentaBancaria.Height = 240
                txtCuentaBancaria.Width = 2760
                txtCuentaBancaria.Text = ""
                txtCuentaBancaria.Visible = True
                txtCuentaBancaria.SetFocus
            End With
        Else
            If Trim(cboCuentasPrevias.Text) <> "" Then
                If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColCuentaBancaria)) <> Trim(cboCuentasPrevias.Text) Then
                    grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColModificación) = "*"
                End If
            
                grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColCuentaBancaria) = Trim(cboCuentasPrevias.Text)
                
                pActualizaTrasCambio grdBusqueda.Row
                cmdGuardar.Enabled = True
            End If
            
            cboCuentasPrevias.Visible = False
            grdBusqueda.SetFocus
                        
            grdBusqueda.Col = cIntColFechaPago
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboCuentasPrevias_KeyPress"))
End Sub

Private Sub cboBancoSAT_LostFocus()
On Error GoTo NotificaError
    
    If txtBancoExtranjero.Visible = False Then
        cboBancoSAT.Visible = False
        grdBusqueda.SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboBancoSAT_LostFocus"))
End Sub

Private Sub cboCorteTransfiere_Click()
On Error GoTo NotificaError
    If vllngCorte <> cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex) Then pConsulta
    vllngCorte = cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboCorteTransfiere_Click"))
End Sub

Private Sub cboCuentasPrevias_LostFocus()
On Error GoTo NotificaError

    cboCuentasPrevias.Visible = False
    If txtCuentaBancaria.Visible Then
        txtCuentaBancaria.SetFocus
    Else
        grdBusqueda.SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboCuentasPrevias_LostFocus"))
End Sub

Private Sub cmdGuardar_Click()
On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long
    Dim vlintSeqFil As Integer
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim vlstrTipo As String
    Dim vlClaveBanco As String
    Dim vlStrMensajeIncompletos As String
    Dim vlStrMensajeIncompletos2 As String

    'Verifica que no existe información incompleta que ya se encuentre semi capturada
    vlStrMensajeIncompletos = ""
    With grdBusqueda
        For vlintSeqFil = 1 To .Rows - 1
            If (Len(Trim(.TextMatrix(vlintSeqFil, cIntColRFC))) = 12 Or Len(Trim(.TextMatrix(vlintSeqFil, cIntColRFC))) = 13) And Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)) <> "" And Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) = "" Then
                '¡No ha ingresado la cuenta bancaria!
                MsgBox SIHOMsg(1289), vbOKOnly + vbExclamation, "Mensaje"
                
                .Row = vlintSeqFil
                .Col = cIntColCuentaBancaria
                .SetFocus
                Exit Sub
            Else
                If (Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Cheque" Or Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Transferencia") And .TextMatrix(vlintSeqFil, cIntColModificación) = "*" And Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)) = "" And Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) = "" And Trim(.TextMatrix(vlintSeqFil, cIntColFechaPago)) = "" Then
                    If InStr(vlStrMensajeIncompletos, Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolio)) & " - " & Trim(.TextMatrix(vlintSeqFil, cIntColTipoDoc)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolioDoc))) = 0 Then
                        vlStrMensajeIncompletos = IIf(vlStrMensajeIncompletos = "", "", vlStrMensajeIncompletos & Chr(13)) & Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolio)) & " - " & Trim(.TextMatrix(vlintSeqFil, cIntColTipoDoc)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolioDoc))
                    End If
                End If
            End If
        Next vlintSeqFil
    End With
    
    'Verifica que no existe información incompleta que ya se encuentre semi capturada
    vlStrMensajeIncompletos2 = ""
    With grdBusqueda
        For vlintSeqFil = 1 To .Rows - 1
            If (Len(Trim(.TextMatrix(vlintSeqFil, cIntColRFC))) = 12 Or Len(Trim(.TextMatrix(vlintSeqFil, cIntColRFC))) = 13) And Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)) <> "" And Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) = "" Then
                '¡No ha ingresado la cuenta bancaria!
                MsgBox SIHOMsg(1289), vbOKOnly + vbExclamation, "Mensaje"
                
                .Row = vlintSeqFil
                .Col = cIntColCuentaBancaria
                .SetFocus
                Exit Sub
            Else
                If (Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Efectivo" Or Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Tarjeta") And .TextMatrix(vlintSeqFil, cIntColModificación) = "*" And Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) = "" Then
                    If InStr(vlStrMensajeIncompletos2, Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolio)) & " - " & Trim(.TextMatrix(vlintSeqFil, cIntColTipoDoc)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolioDoc))) = 0 Then
                        vlStrMensajeIncompletos2 = IIf(vlStrMensajeIncompletos2 = "", "", vlStrMensajeIncompletos2 & Chr(13)) & Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolio)) & " - " & Trim(.TextMatrix(vlintSeqFil, cIntColTipoDoc)) & " " & Trim(.TextMatrix(vlintSeqFil, cIntColFolioDoc))
                    End If
                End If
            End If
        Next vlintSeqFil
    End With

    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
    
        pEjecutaSentencia "DELETE FROM PVCORTECHEQUETRANSCTA WHERE intnumcorte = " & Str(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex))
    
        With grdBusqueda
            For vlintSeqFil = 1 To .Rows - 1
                Select Case Trim(.TextMatrix(vlintSeqFil, cIntColTipoDoc))
                    Case "FACTURA"
                        vlstrTipo = "FA"
                    Case "RECIBO"
                        vlstrTipo = "RE"
                    Case "PAGO CRÉDITO"
                        vlstrTipo = "RC"
                    Case "TICKET"
                        vlstrTipo = "TI"
                    Case "HONORARIO"
                        vlstrTipo = "HO"
                    Case "FONDO FIJO"
                        vlstrTipo = "RI"
                    Case "TRANSFERENCIA DINERO"
                        vlstrTipo = "TR"
                    Case "SALIDA DINERO"
                        vlstrTipo = "SD"
                    Case "SALIDA CAJA CHICA"
                        vlstrTipo = "SC"
                    Case "ENTRADA CAJA CHICA"
                        vlstrTipo = "EC"
                    Case "SALDO CAJA CHICA"
                        vlstrTipo = "SA"
                    Case "CARGOS SOCIOS"
                        vlstrTipo = "SO"
                End Select
            
                Set rs2 = frsRegresaRs("SELECT PVDETALLECORTE.INTNUMCORTE, PVDETALLECORTE.INTCONSECUTIVO FROM PVDETALLECORTE INNER JOIN PVCORTE ON PVCORTE.INTNUMCORTE = PVDETALLECORTE.INTNUMCORTE INNER JOIN PVFORMAPAGO ON PVDETALLECORTE.INTFORMAPAGO = PVFORMAPAGO.INTFORMAPAGO WHERE PVDETALLECORTE.INTNUMCORTE <> " & Str(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)) & " AND TRIM(PVDETALLECORTE.CHRTIPODOCUMENTO) = '" & vlstrTipo & "' AND TRIM(PVDETALLECORTE.CHRFOLIODOCUMENTO) = '" & Trim(.TextMatrix(vlintSeqFil, cIntColFolioDoc)) & "' AND CASE WHEN PVFORMAPAGO.CHRTIPO = 'B' THEN 'Transferencia' WHEN PVFORMAPAGO.CHRTIPO = 'H' THEN 'Cheque' WHEN PVFORMAPAGO.CHRTIPO = 'E' THEN 'Efectivo' WHEN PVFORMAPAGO.CHRTIPO = 'T' THEN 'Tarjeta' END = '" & Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) & "' AND PVDETALLECORTE.INTFOLIOCHEQUE = '" & Trim(.TextMatrix(vlintSeqFil, cIntColFolio)) & "' AND PVCORTE.SMIDEPARTAMENTO = " & vllngClaveDepartamento, adLockReadOnly, adOpenForwardOnly)
                If rs2.RecordCount <> 0 Then
                    pEjecutaSentencia "DELETE FROM PVCORTECHEQUETRANSCTA WHERE intnumcorte = " & rs2!intnumcorte & " AND INTCONSECUTIVODETCORTE = " & rs2!INTCONSECUTIVO
                    
                    If Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Cheque" Or Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Transferencia" Then
                        'Cheque o Transferencia
                        If Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT)) <> "" And Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) <> "" Then
                        
                            If Len(Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))) = 3 Then
                                vlClaveBanco = Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))
                            Else
                                If Len(Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))) = 2 Then
                                    vlClaveBanco = "0" & Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))
                                Else
                                    vlClaveBanco = "00" & Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))
                                End If
                            End If
                                                                        
                            frsEjecuta_SP rs2!intnumcorte & "|" & rs2!INTCONSECUTIVO & "|'" & Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) & "'|'" & vlClaveBanco & "'|'" & Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) & "'|'" & fstrFechaSQL(Trim(.TextMatrix(vlintSeqFil, cIntColFechaPago))) & "'|'" & IIf(vlClaveBanco = "000" And Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)) <> "", Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)), "") & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                        End If
                    Else
                        'Efectivo o tarjeta
                        frsEjecuta_SP rs2!intnumcorte & "|" & rs2!INTCONSECUTIVO & "|'" & Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) & "'|'000'|''|''|''", "SP_PVINSCORTECHEQUETRANSCTA"
                    End If
                End If
                
                If Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Cheque" Or Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = "Transferencia" Then
                    'Cheque o Transferencia
                    If Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT)) <> "" And Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) <> "" Then
                        If Len(Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))) = 3 Then
                            vlClaveBanco = Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))
                        Else
                            If Len(Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))) = 2 Then
                                vlClaveBanco = "0" & Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))
                            Else
                                vlClaveBanco = "00" & Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT))
                            End If
                        End If
                                        
                        frsEjecuta_SP Str(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)) & "|" & Trim(.TextMatrix(vlintSeqFil, cIntColConsecutivo)) & "|'" & Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) & "'|'" & vlClaveBanco & "'|'" & Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) & "'|'" & fstrFechaSQL(Trim(.TextMatrix(vlintSeqFil, cIntColFechaPago))) & "'|'" & IIf(vlClaveBanco = "000" And Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)) <> "", Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)), "") & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                    
                        If Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) <> "" And vlClaveBanco <> "000" Then
                            Set rs = frsRegresaRs("SELECT CHRRFC FROM PVRFCCTABANCOSAT WHERE TRIM(CHRRFC) = '" & Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) & "' AND TRIM(CHRCLAVEBANCOSAT) = '" & vlClaveBanco & "' AND TRIM(VCHCUENTABANCARIA) = '" & Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) & "'", adLockReadOnly, adOpenForwardOnly)
                            If rs.RecordCount = 0 Then
                                pEjecutaSentencia "INSERT INTO PVRFCCTABANCOSAT (CHRRFC,CHRCLAVEBANCOSAT,VCHCUENTABANCARIA) VALUES ('" & Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) & "','" & vlClaveBanco & "','" & Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) & "')"
                            End If
                        End If
                    End If
                Else
                    'Efectivo o tarjeta
                    frsEjecuta_SP Str(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)) & "|" & Trim(.TextMatrix(vlintSeqFil, cIntColConsecutivo)) & "|'" & Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) & "'|'000'|''|''|''", "SP_PVINSCORTECHEQUETRANSCTA"
                End If
            Next vlintSeqFil
        End With
        
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, CInt(vllngPersonaGraba), "CUENTAS BANCARIAS CORTE CAJA", cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex))
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        If vlStrMensajeIncompletos <> "" Or vlStrMensajeIncompletos2 <> "" Then
            If vlStrMensajeIncompletos <> "" Then
                'No fue capturada la información del banco, cuenta bancaria y fecha de los siguientes movimientos, se mostrará el RFC relacionado al movimiento del corte.
                MsgBox SIHOMsg(1346) & Chr(13) & vlStrMensajeIncompletos, vbOKOnly + vbInformation, "Mensaje"
            End If
            If vlStrMensajeIncompletos2 <> "" Then
                'No fue capturado el RFC de los siguientes movimientos, se mostrará el RFC relacionado al movimiento del corte.
                MsgBox SIHOMsg(1365) & Chr(13) & vlStrMensajeIncompletos2, vbOKOnly + vbInformation, "Mensaje"
            End If
        Else
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
        End If

        cmdGuardar.Enabled = False
        
        pConsulta
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGuardar_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If cboBancoSAT.Visible Then
            grdBusqueda.SetFocus
            cboBancoSAT.Visible = False
        Else
            If txtCuentaBancaria.Visible Then
                grdBusqueda.SetFocus
                txtCuentaBancaria.Visible = False
            Else
                If txtRFC.Visible Then
                    grdBusqueda.SetFocus
                    txtRFC.Visible = False
                Else
                    If txtBancoExtranjero.Visible Then
                        grdBusqueda.SetFocus
                        txtBancoExtranjero.Visible = False
                    Else
                        If cboCuentasPrevias.Visible Then
                            grdBusqueda.SetFocus
                            cboCuentasPrevias.Visible = False
                        Else
                            If MskFecha.Visible Then
                                grdBusqueda.SetFocus
                                MskFecha.Visible = False
                            Else
                                If cmdGuardar.Enabled Then
                                    '¿Desea abandonar la operación?
                                    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                                        cmdGuardar.Enabled = False
                                        pConsulta
                                    End If
                                Else
                                    Unload Me
                                End If
                            End If
                        End If
                    End If
                End If
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
    
    If cgstrModulo = "CN" Then
        vlblnLectura = Not (fblnRevisaPermiso(vglngNumeroLogin, 3056, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3056, "C", True))
    Else
        If cgstrModulo = "PV" Then
            vlblnLectura = Not (fblnRevisaPermiso(vglngNumeroLogin, 3053, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3053, "C", True))
        Else
            vlblnLectura = Not (fblnRevisaPermiso(vglngNumeroLogin, 3054, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3054, "C", True))
        End If
    End If
    cmdGuardar.Enabled = False
    
    pCargaCorte
    
    vllngCorte = cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)
    
    grdBusqueda.Rows = 2
    grdBusqueda.Cols = cIntColFechaPago
    pConfiguraGridConsulta
    pConsulta
    
    pCargaBancosSAT
    
    txtCuentaBancaria.Text = ""
    txtRFC.Text = ""
    txtBancoExtranjero.Text = ""
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pConfiguraGridConsulta()
    On Error GoTo NotificaError
    Dim vlintseq As Integer
        
    With grdBusqueda
        .FormatString = "|Consecutivo|Corte|Fecha y hora|Tipo documento|Folio|Estado|Tipo pago|Referencia|Razón social / paciente|RFC|Clave banco|Banco del SAT o extranjero|Cuenta bancaria|Fecha del cheque / transferencia|Modificada"
        
        .ColWidth(0) = 100
        .ColWidth(cIntColConsecutivo) = 0       'Consecutivo
        .ColWidth(cIntColNumCorte) = 0          'Corte
        .ColWidth(cIntColFechaDoc) = 1550       'Fecha y hora
        .ColWidth(cIntColTipoDoc) = 2850        'Tipo documento
        .ColWidth(cIntColFolioDoc) = 1250       'Folio
        .ColWidth(cIntColEstado) = 900          'Estado
        .ColWidth(cIntColTipoPago) = 1075       'Tipo pago
        .ColWidth(cIntColFolio) = 1900          'Referencia
        .ColWidth(cIntColRazonSocial) = 3350    'Razón social
        .ColWidth(cIntColRFC) = 1700            'RFC
        .ColWidth(cIntColClaveBancoSAT) = 0     'Clave banco SAT
        .ColWidth(cIntColDescBancoSAT) = 2020   'Descripcion banco SAT
        .ColWidth(cIntColCuentaBancaria) = 2825 'Cuenta bancaria
        .ColWidth(cIntColFechaPago) = 2480      'Fecha del cheque, transferencia o pago (Cheque o Transferencia)
        .ColWidth(cIntColModificación) = 0      'Modificaciones
        
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColConsecutivo) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColNumCorte) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColFechaDoc) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColTipoDoc) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColTipoPago) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColFolioDoc) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColEstado) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColRazonSocial) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColRFC) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColFolio) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColClaveBancoSAT) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColDescBancoSAT) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColCuentaBancaria) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColFechaPago) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColModificación) = flexAlignCenterCenter
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(cIntColConsecutivo) = flexAlignLeftCenter
        .ColAlignment(cIntColNumCorte) = flexAlignLeftCenter
        .ColAlignment(cIntColFechaDoc) = flexAlignLeftCenter
        .ColAlignment(cIntColTipoDoc) = flexAlignLeftCenter
        .ColAlignment(cIntColTipoPago) = flexAlignLeftCenter
        .ColAlignment(cIntColFolioDoc) = flexAlignLeftCenter
        .ColAlignment(cIntColEstado) = flexAlignLeftCenter
        .ColAlignment(cIntColRazonSocial) = flexAlignLeftCenter
        .ColAlignment(cIntColRFC) = flexAlignLeftCenter
        .ColAlignment(cIntColFolio) = flexAlignLeftCenter
        .ColAlignment(cIntColClaveBancoSAT) = flexAlignLeftCenter
        .ColAlignment(cIntColDescBancoSAT) = flexAlignLeftCenter
        .ColAlignment(cIntColCuentaBancaria) = flexAlignLeftCenter
        .ColAlignment(cIntColFechaPago) = flexAlignLeftCenter
        .ColAlignment(cIntColModificación) = flexAlignCenterCenter
        
        For vlintseq = 1 To .Rows - 1
            .TextMatrix(vlintseq, cIntColFechaDoc) = LCase(Format(.TextMatrix(vlintseq, cIntColFechaDoc), "dd/mmm/yyyy HH:mm"))
            .TextMatrix(vlintseq, cIntColModificación) = ""
        Next vlintseq
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridConsulta"))
End Sub

Private Sub pConsulta()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
        
    If cmdGuardar.Enabled = True And vllngCorte <> 0 Then
        '¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbNo Then
            cboCorteTransfiere.ListIndex = fintLocalizaCbo(cboCorteTransfiere, Str(vllngCorte))
            Exit Sub
        End If
    End If
        
    cmdGuardar.Enabled = False
                                        
    With grdBusqueda
        .Redraw = False
        .Clear
        .Rows = 2
        .Cols = 9
                
        pConfiguraGridConsulta
        Set rs = frsEjecuta_SP(Str(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)), "SP_PVSELCORTEDETALLECHQTRANS")
        If rs.RecordCount <> 0 Then
            pLlenarMshFGrdRs grdBusqueda, rs, 0
            pConfiguraGridConsulta
        End If
        .Redraw = True
        
        pFormatoFechaLargaColumnaGrid grdBusqueda, cIntColFechaPago
        
        vllngClaveDepartamento = 0
        Set rs = frsRegresaRs("select nodepartamento.smicvedepartamento from pvcorte inner join nodepartamento on nodepartamento.smicvedepartamento = pvcorte.smidepartamento inner join cnempresacontable on cnempresacontable.tnyclaveempresa = nodepartamento.tnyclaveempresa where pvcorte.intnumcorte = " & Str(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)), adLockReadOnly, adOpenForwardOnly)
        If rs.RecordCount <> 0 Then
            vllngClaveDepartamento = rs!smicvedepartamento
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConsulta"))
End Sub

Public Sub pFormatoFechaLargaColumnaGrid(grdNombre As MSHFlexGrid, vlintxColumna As Integer, Optional vlstrSigno As String)
'----------------------------------------------------------------------
' Procedimiento para dar formato a la columna del grid que son Fechas
'----------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim X As Long
    
    For X = 1 To grdNombre.Rows - 1
        grdNombre.TextMatrix(X, vlintxColumna) = Format(grdNombre.TextMatrix(X, vlintxColumna), "DD/MMM/YYYY")
    Next X

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFormatoFechaColumnaGrid"))
End Sub

Private Sub pCargaBancosSAT()
    'Llenado del listado de los bancos publicados por el SAT
    On Error GoTo NotificaError

    pLlenarCboSentencia cboBancoSAT, "SELECT chrclave, CASE WHEN TRIM(vchnombrecorto) = TRIM(vchnombrerazonsocial) THEN TRIM(vchnombrecorto) ELSE TRIM(vchnombrecorto) || ' - ' || TRIM(vchnombrerazonsocial) END descripcion FROM CPBANCOSAT WHERE bitactivo = 1 ORDER BY descripcion", 1, 0
    
    cboBancoSAT.AddItem "<BANCO EXTRANJERO>", 0
    cboBancoSAT.ItemData(cboBancoSAT.newIndex) = -1
    cboBancoSAT.ListIndex = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaBancosSAT"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    If cmdGuardar.Enabled Then
        '¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            cmdGuardar.Enabled = False
            pConsulta
        Else
            Cancel = True
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub grdBusqueda_Click()
    On Error GoTo NotificaError
    
    If Not vlblnLectura Then
        If grdBusqueda.MouseRow >= grdBusqueda.FixedCols Then pHabilitaSeleccion
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdBusqueda_Click"))
End Sub

Private Sub pHabilitaSeleccion()
    Dim rs As New ADODB.Recordset
    Dim vlintNuevaCuenta As Integer
    Dim vlClaveBanco As String
    On Error GoTo NotificaError

    With grdBusqueda
        If grdBusqueda.Col = cIntColRFC And Trim(.TextMatrix(.Row, cIntColConsecutivo)) <> "" Then
            txtRFC.Top = .Top + .CellTop + 48
            txtRFC.Left = .Left + .CellLeft + 24
            txtRFC.Height = 240
            txtRFC.Width = .CellWidth - 40
            txtRFC.Text = Trim(.TextMatrix(.Row, cIntColRFC))
            txtRFC.Visible = True
            txtRFC.SetFocus
        Else
            If (Len(Trim(.TextMatrix(.Row, cIntColRFC))) = 12 Or Len(Trim(.TextMatrix(.Row, cIntColRFC))) = 13) And .TextMatrix(.Row, cIntColRFC) <> "ACO560518KW7" And .TextMatrix(.Row, cIntColRFC) <> "AAA01010101AAA" Then
                If Trim(.TextMatrix(.Row, cIntColTipoPago)) = "Cheque" Or Trim(.TextMatrix(.Row, cIntColTipoPago)) = "Transferencia" Then
                    If grdBusqueda.Col = cIntColDescBancoSAT And Trim(.TextMatrix(.Row, cIntColConsecutivo)) <> "" Then
                        If fintLocalizaCbo(cboBancoSAT, Val(Trim(.TextMatrix(.Row, cIntColClaveBancoSAT)))) <> -1 Or Trim(.TextMatrix(.Row, cIntColDescBancoSAT)) = "" Then
                            cboBancoSAT.Top = .Top + .CellTop - 16
                            cboBancoSAT.Left = .Left + .CellLeft - 16
                            cboBancoSAT.Width = 4825
                            cboBancoSAT.ListIndex = fintLocalizaCbo(cboBancoSAT, Val(Trim(.TextMatrix(.Row, cIntColClaveBancoSAT))))
                            If cboBancoSAT.ListIndex = -1 Then cboBancoSAT.ListIndex = 1
                            cboBancoSAT.Visible = True
                            cboBancoSAT.SetFocus
                        Else
                            cboBancoSAT.Visible = False
                            
                            txtBancoExtranjero.Text = Trim(.TextMatrix(.Row, cIntColDescBancoSAT))
                            txtBancoExtranjero.Top = grdBusqueda.Top + grdBusqueda.CellTop + 48
                            txtBancoExtranjero.Left = grdBusqueda.Left + grdBusqueda.CellLeft + 24
                            txtBancoExtranjero.Height = grdBusqueda.CellHeight - 56
                            txtBancoExtranjero.Width = grdBusqueda.CellWidth - 40
                            txtBancoExtranjero.Visible = True
                            txtBancoExtranjero.SetFocus
                        End If
                    Else
                        If grdBusqueda.Col = cIntColCuentaBancaria And Trim(.TextMatrix(.Row, cIntColDescBancoSAT)) <> "" And Trim(.TextMatrix(.Row, cIntColConsecutivo)) <> "" Then
                        
                            If Len(Trim(.TextMatrix(.Row, cIntColClaveBancoSAT))) = 3 Then
                                vlClaveBanco = Trim(.TextMatrix(.Row, cIntColClaveBancoSAT))
                            Else
                                If Len(Trim(.TextMatrix(.Row, cIntColClaveBancoSAT))) = 2 Then
                                    vlClaveBanco = "0" & Trim(.TextMatrix(.Row, cIntColClaveBancoSAT))
                                Else
                                    vlClaveBanco = "00" & Trim(.TextMatrix(.Row, cIntColClaveBancoSAT))
                                End If
                            End If
                        
                            Set rs = frsRegresaRs("SELECT intidRegistro, vchcuentabancaria FROM PVRFCCTABANCOSAT WHERE TRIM(CHRRFC) = '" & Trim(.TextMatrix(.Row, cIntColRFC)) & "' AND TRIM(CHRCLAVEBANCOSAT) = '" & vlClaveBanco & "' ORDER BY vchcuentabancaria", adLockReadOnly, adOpenForwardOnly)
                            If rs.RecordCount <> 0 Then
                                cboCuentasPrevias.Clear
    '                            cboCuentasPrevias.AddItem "<NUEVA>"
                                cboCuentasPrevias.AddItem ""
                                cboCuentasPrevias.ItemData(cboCuentasPrevias.newIndex) = 0
                                
                                Do While Not rs.EOF
                                    cboCuentasPrevias.AddItem rs!vchcuentabancaria
                                    cboCuentasPrevias.ItemData(cboCuentasPrevias.newIndex) = rs!intidRegistro
                                    vlintNuevaCuenta = rs!intidRegistro + 1
                                    rs.MoveNext
                                Loop
                                
                                cboCuentasPrevias.Top = .Top + .CellTop - 16
                                cboCuentasPrevias.Left = .Left + .CellLeft - 16
                                cboCuentasPrevias.Width = 2830
                                cboCuentasPrevias.ListIndex = fintLocalizaCritCbo(cboCuentasPrevias, Trim(.TextMatrix(.Row, cIntColCuentaBancaria)))
                                
                                If cboCuentasPrevias.ListIndex = -1 And Trim(.TextMatrix(.Row, cIntColCuentaBancaria)) <> "" Then
                                    cboCuentasPrevias.AddItem Trim(.TextMatrix(.Row, cIntColCuentaBancaria))
                                    cboCuentasPrevias.ItemData(cboCuentasPrevias.newIndex) = vlintNuevaCuenta
                                    cboCuentasPrevias.ListIndex = fintLocalizaCritCbo(cboCuentasPrevias, Trim(.TextMatrix(.Row, cIntColCuentaBancaria)))
                                End If
                                If cboCuentasPrevias.ListIndex = -1 Then cboCuentasPrevias.ListIndex = 0
                                
                                cboCuentasPrevias.Visible = True
                                cboCuentasPrevias.SetFocus
                            Else
                                txtCuentaBancaria.Top = .Top + .CellTop + 48
                                txtCuentaBancaria.Left = .Left + .CellLeft + 24
                                txtCuentaBancaria.Height = 240
                                txtCuentaBancaria.Width = 2760
                                txtCuentaBancaria.Text = Trim(.TextMatrix(.Row, cIntColCuentaBancaria))
                                txtCuentaBancaria.Visible = True
                                txtCuentaBancaria.SetFocus
                            End If
                        End If
                        
                        If grdBusqueda.Col = cIntColFechaPago And Trim(.TextMatrix(.Row, cIntColDescBancoSAT)) <> "" And Trim(.TextMatrix(.Row, cIntColConsecutivo)) <> "" Then
                            MskFecha.Top = .Top + .CellTop + 48
                            MskFecha.Left = .Left + .CellLeft + 24
                            MskFecha.Height = 240
                            MskFecha.Width = 2350
                            MskFecha.Text = CDate(.TextMatrix(.Row, .Col))
                            MskFecha.Visible = True
                            MskFecha.SetFocus
                        End If
                    End If
                End If
            End If
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaSeleccion"))
End Sub

Private Sub pQuitaBanco()
    On Error GoTo NotificaError

    With grdBusqueda
        If Not ((Len(Trim(.TextMatrix(.Row, cIntColRFC))) = 12 Or Len(Trim(.TextMatrix(.Row, cIntColRFC))) = 13) And .TextMatrix(.Row, cIntColRFC) <> "ACO560518KW7" And .TextMatrix(.Row, cIntColRFC) <> "AAA01010101AAA") Then
            If Trim(.TextMatrix(.Row, cIntColClaveBancoSAT)) <> "" Or Trim(.TextMatrix(.Row, cIntColDescBancoSAT)) <> "" Or Trim(.TextMatrix(.Row, cIntColCuentaBancaria)) <> "" Then
                .TextMatrix(.Row, cIntColModificación) = "*"
            End If
            
            .TextMatrix(.Row, cIntColClaveBancoSAT) = ""
            .TextMatrix(.Row, cIntColDescBancoSAT) = ""
            .TextMatrix(.Row, cIntColCuentaBancaria) = ""

            pActualizaTrasCambio .Row
            cmdGuardar.Enabled = True
        Else
            If grdBusqueda.Col = cIntColDescBancoSAT Then
                If Trim(.TextMatrix(.Row, cIntColClaveBancoSAT)) <> "" Or Trim(.TextMatrix(.Row, cIntColDescBancoSAT)) <> "" Or Trim(.TextMatrix(.Row, cIntColCuentaBancaria)) <> "" Then
                    .TextMatrix(.Row, cIntColModificación) = "*"
                End If
            
                .TextMatrix(.Row, cIntColClaveBancoSAT) = ""
                .TextMatrix(.Row, cIntColDescBancoSAT) = ""
                .TextMatrix(.Row, cIntColCuentaBancaria) = ""
                
                pActualizaTrasCambio .Row
                cmdGuardar.Enabled = True
            Else
                If grdBusqueda.Col = cIntColCuentaBancaria Then
                    If Trim(.TextMatrix(.Row, cIntColCuentaBancaria)) <> "" Then
                        .TextMatrix(.Row, cIntColModificación) = "*"
                    End If
                    
                    .TextMatrix(.Row, cIntColCuentaBancaria) = ""
                    
                    pActualizaTrasCambio .Row
                    cmdGuardar.Enabled = True
                End If
            End If
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pQuitaBanco"))
End Sub

Private Sub grdBusqueda_DblClick()
    Dim vlCol As Integer
    With grdBusqueda
        If .MouseRow < .FixedCols And (.Col = cIntColFechaDoc Or .Col = cIntColTipoDoc Or .Col = cIntColTipoPago Or .Col = cIntColFolioDoc Or .Col = cIntColEstado Or .Col = cIntColRazonSocial Or .Col = cIntColRFC Or .Col = cIntColDescBancoSAT Or .Col = cIntColCuentaBancaria) Then
            .Redraw = False
            vlCol = .MouseCol
            
            .Col = cIntColFechaPago
            .ColSel = vlCol
            .Sort = 5
            
            .Col = vlCol

            .Redraw = True
            If .Enabled Then .SetFocus
        End If
    End With
End Sub

Private Sub grdBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    If Not vlblnLectura Then
        If KeyCode = vbKeyReturn Then pHabilitaSeleccion
        If KeyCode = vbKeyDelete Then pQuitaBanco
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdBusqueda_KeyPress"))
End Sub

Private Sub grdBusqueda_Scroll()
    If Not vlblnLectura Then
        grdBusqueda.SetFocus
        cboBancoSAT.Visible = False
        cboCuentasPrevias.Visible = False
        txtCuentaBancaria.Visible = False
        txtRFC.Visible = False
        txtBancoExtranjero.Visible = False
        MskFecha.Visible = False
    End If
End Sub

Private Sub mskFecha_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto MskFecha

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":MskFecha_GotFocus"))
End Sub

Private Sub MskFecha_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If Not IsDate(MskFecha) Then
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa"
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            MskFecha.Text = CDate(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColFechaPago))
            MskFecha.SetFocus
            Exit Sub
        End If
        
        If Year(CDate(MskFecha.Text)) < 1900 Then
            '¡Fecha no válida!
            MsgBox SIHOMsg(254), vbOKOnly + vbExclamation, "Mensaje"
            MskFecha.Text = CDate(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColFechaPago))
            MskFecha.SetFocus
            Exit Sub
        End If
        
        If CDate(MskFecha.Text) > CDate(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColFechaDoc)) Then
            '¡La fecha debe ser menor o igual a la fecha del documento!
            MsgBox SIHOMsg(1316), vbOKOnly + vbExclamation, "Mensaje"
            MskFecha.Text = CDate(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColFechaPago))
            MskFecha.SetFocus
            Exit Sub
        End If
        
        If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColFechaPago)) <> Format(CDate(MskFecha.Text), "dd/mmm/yyyy") Then
            grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColModificación) = "*"
        End If
        
        grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColFechaPago) = Format(CDate(MskFecha.Text), "dd/mmm/yyyy")
        
        pActualizaTrasCambio grdBusqueda.Row
        cmdGuardar.Enabled = True
        
        MskFecha.Visible = False
        grdBusqueda.SetFocus
        
        If grdBusqueda.Rows - grdBusqueda.Row > 1 Then
            grdBusqueda.Row = grdBusqueda.Row + 1
            grdBusqueda.Col = cIntColRFC
'            grdBusqueda.Col = cIntColDescBancoSAT
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":MskFecha_KeyPress"))
End Sub

Private Sub MskFecha_LostFocus()
On Error GoTo NotificaError
    
    MskFecha.Visible = False
    grdBusqueda.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":MskFecha_LostFocus"))
End Sub

Private Sub txtBancoExtranjero_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtBancoExtranjero

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBancoExtranjero_GotFocus"))
End Sub

Private Sub txtBancoExtranjero_KeyPress(KeyAscii As Integer)
    Dim vlClaveBanco As String
    On Error GoTo NotificaError

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

    If KeyAscii = 13 Then
        If Trim(txtBancoExtranjero.Text) <> "" Then
            If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColClaveBancoSAT)) <> "0" Or Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColDescBancoSAT)) <> Trim(txtBancoExtranjero.Text) Then
                grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColModificación) = "*"
            End If
        
            grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColClaveBancoSAT) = "0"
            grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColDescBancoSAT) = Trim(txtBancoExtranjero.Text)
            
            pActualizaTrasCambio grdBusqueda.Row
            cmdGuardar.Enabled = True
        Else
            pQuitaBanco
        End If
        
        txtBancoExtranjero.Visible = False
        grdBusqueda.SetFocus
        
        If Trim(txtBancoExtranjero.Text) = "" Then
            If grdBusqueda.Rows - grdBusqueda.Row > 1 Then
                grdBusqueda.Row = grdBusqueda.Row + 1
                grdBusqueda.Col = cIntColRFC
'                grdBusqueda.Col = cIntColDescBancoSAT
            End If
        Else
            grdBusqueda.Col = cIntColCuentaBancaria
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtBancoExtranjero_KeyPress"))
End Sub

Private Sub txtBancoExtranjero_LostFocus()
On Error GoTo NotificaError

    txtBancoExtranjero.Visible = False
    grdBusqueda.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtBancoExtranjero_LostFocus"))
End Sub

Private Sub txtCuentaBancaria_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtCuentaBancaria

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuentaBancaria_GotFocus"))
End Sub

Private Sub txtRFC_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtRFC

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtRFC_GotFocus"))
End Sub

Private Sub txtCuentaBancaria_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
    
    If KeyAscii = 13 Then
        If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColCuentaBancaria)) <> Trim(txtCuentaBancaria.Text) Then
            grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColModificación) = "*"
        End If
    
        grdBusqueda.TextMatrix(grdBusqueda.Row, cIntColCuentaBancaria) = Trim(txtCuentaBancaria.Text)
        
        pActualizaTrasCambio grdBusqueda.Row
        cmdGuardar.Enabled = True
        txtCuentaBancaria.Visible = False
        grdBusqueda.SetFocus
                
        grdBusqueda.Col = cIntColFechaPago
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuentaBancaria_KeyPress"))
End Sub

Private Sub txtCuentaBancaria_LostFocus()
On Error GoTo NotificaError

    txtCuentaBancaria.Visible = False
    grdBusqueda.SetFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCuentaBancaria_LostFocus"))
End Sub

Private Sub txtRFC_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
Dim vlstrcaracter As String

    With grdBusqueda
        If KeyAscii <> 8 Then
            If KeyAscii = 13 Then
                
                If Trim(txtRFC.Text) = "" Or (Len(Trim(txtRFC.Text)) = 12 Or Len(Trim(txtRFC.Text)) = 13) Then
                    If Trim(.TextMatrix(.Row, cIntColRFC)) <> Trim(txtRFC.Text) Then
                        .TextMatrix(.Row, cIntColModificación) = "*"
                    End If
                    
                    .TextMatrix(.Row, cIntColRFC) = Trim(txtRFC.Text)
    
                    pActualizaTrasCambio .Row
                    cmdGuardar.Enabled = True
                    txtRFC.Visible = False
                    .SetFocus
                    
                    If (Trim(.TextMatrix(.Row, cIntColTipoPago)) = "Cheque" Or Trim(.TextMatrix(.Row, cIntColTipoPago)) = "Transferencia") And (Len(Trim(.TextMatrix(.Row, cIntColRFC))) = 12 Or Len(Trim(.TextMatrix(.Row, cIntColRFC))) = 13) And .TextMatrix(.Row, cIntColRFC) <> "ACO560518KW7" And .TextMatrix(.Row, cIntColRFC) <> "AAA01010101AAA" Then
                        .Col = cIntColDescBancoSAT
                    Else
                        pQuitaBanco
                        If .Rows - .Row > 1 Then
                            .Row = .Row + 1
                        End If
                    End If
                Else
                    If Trim(txtRFC.Text) <> "" And Len(Trim(txtRFC.Text)) <> 12 And Len(Trim(txtRFC.Text)) <> 13 Then
                        'El RFC ingresado no tiene un tamaño válido, favor de verificar:
                        MsgBox SIHOMsg(1345), vbOKOnly + vbInformation, "Mensaje"
                    End If
                End If
            Else
                vlstrcaracter = fStrRFCValido(Chr(KeyAscii))
                If vlstrcaracter <> "" Then
                    KeyAscii = Asc(UCase(vlstrcaracter))
                Else
                    KeyAscii = 7
                End If
            End If
        End If
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtRFC_KeyPress"))
End Sub

Private Sub txtRFC_LostFocus()
On Error GoTo NotificaError

    txtRFC.Visible = False
    grdBusqueda.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtRFC_LostFocus"))
End Sub

Private Sub pActualizaTrasCambio(vlintRenglon As Integer)
    Dim vlintSeqFil As Integer
    Dim vlstrTipoDoc As String
    Dim vlstrTipoPago As String
    Dim vlstrReferencia As String
    Dim vlstrFolio As String
    Dim vlstrRFC As String
    Dim vlstrClaveBanco As String
    Dim vlstrBanco As String
    Dim vlstrCuenta As String
    Dim vlstrFecha As String
    On Error GoTo NotificaError

    With grdBusqueda
        pFechaChequeTrans vlintRenglon
    
        vlstrTipoDoc = Trim(.TextMatrix(vlintRenglon, cIntColTipoDoc))
        vlstrReferencia = Trim(.TextMatrix(vlintRenglon, cIntColFolioDoc))
        vlstrTipoPago = Trim(.TextMatrix(vlintRenglon, cIntColTipoPago))
        vlstrFolio = Trim(.TextMatrix(vlintRenglon, cIntColFolio))
        
        vlstrRFC = Trim(.TextMatrix(vlintRenglon, cIntColRFC))
        
        If Len(Trim(.TextMatrix(vlintRenglon, cIntColClaveBancoSAT))) = 3 Then
            vlstrClaveBanco = Trim(.TextMatrix(vlintRenglon, cIntColClaveBancoSAT))
        Else
            If Len(Trim(.TextMatrix(vlintRenglon, cIntColClaveBancoSAT))) = 2 Then
                vlstrClaveBanco = "0" & Trim(.TextMatrix(vlintRenglon, cIntColClaveBancoSAT))
            Else
                vlstrClaveBanco = "00" & Trim(.TextMatrix(vlintRenglon, cIntColClaveBancoSAT))
            End If
        End If
        
        vlstrBanco = Trim(.TextMatrix(vlintRenglon, cIntColDescBancoSAT))
        vlstrCuenta = Trim(.TextMatrix(vlintRenglon, cIntColCuentaBancaria))
        vlstrFecha = Trim(.TextMatrix(vlintRenglon, cIntColFechaPago))
    
        For vlintSeqFil = 1 To .Rows - 1
            If Trim(.TextMatrix(vlintSeqFil, cIntColTipoDoc)) = vlstrTipoDoc _
                And Trim(.TextMatrix(vlintSeqFil, cIntColFolioDoc)) = vlstrReferencia _
                    And Trim(.TextMatrix(vlintSeqFil, cIntColTipoPago)) = vlstrTipoPago _
                        And Trim(.TextMatrix(vlintSeqFil, cIntColFolio)) = vlstrFolio Then
                            If Trim(.TextMatrix(vlintSeqFil, cIntColRFC)) <> vlstrRFC Or Trim(.TextMatrix(vlintSeqFil, cIntColClaveBancoSAT)) <> vlstrClaveBanco Or Trim(.TextMatrix(vlintSeqFil, cIntColDescBancoSAT)) <> vlstrBanco Or Trim(.TextMatrix(vlintSeqFil, cIntColCuentaBancaria)) <> vlstrCuenta Or Trim(.TextMatrix(vlintSeqFil, cIntColFechaPago)) <> vlstrFecha Then
                                .TextMatrix(vlintSeqFil, cIntColModificación) = "*"
                            End If

                            .TextMatrix(vlintSeqFil, cIntColRFC) = vlstrRFC
                            .TextMatrix(vlintSeqFil, cIntColClaveBancoSAT) = vlstrClaveBanco
                            .TextMatrix(vlintSeqFil, cIntColDescBancoSAT) = vlstrBanco
                            .TextMatrix(vlintSeqFil, cIntColCuentaBancaria) = vlstrCuenta
                            .TextMatrix(vlintSeqFil, cIntColFechaPago) = vlstrFecha
            End If
        Next vlintSeqFil
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pActualizaTrasCambio"))
End Sub

Private Sub pFechaChequeTrans(vlintRenglon As Integer)
    
    With grdBusqueda
        If Trim(.TextMatrix(vlintRenglon, cIntColDescBancoSAT)) = "" And Trim(.TextMatrix(vlintRenglon, cIntColCuentaBancaria)) = "" Then
            If Trim(.TextMatrix(vlintRenglon, cIntColFechaPago)) <> "" Then
                .TextMatrix(vlintRenglon, cIntColModificación) = "*"
            End If
            
            .TextMatrix(vlintRenglon, cIntColFechaPago) = ""
        Else
            If Trim(.TextMatrix(vlintRenglon, cIntColFechaPago)) = "" Then
                If Trim(.TextMatrix(vlintRenglon, cIntColFechaPago)) <> Format(CDate(.TextMatrix(vlintRenglon, cIntColFechaDoc)), "dd/mmm/yyyy") Then
                    .TextMatrix(vlintRenglon, cIntColModificación) = "*"
                End If
            
                .TextMatrix(vlintRenglon, cIntColFechaPago) = Format(CDate(.TextMatrix(vlintRenglon, cIntColFechaDoc)), "dd/mmm/yyyy")
            End If
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFechaChequeTrans"))
End Sub
