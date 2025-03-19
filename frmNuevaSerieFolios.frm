VERSION 5.00
Begin VB.Form frmNuevaSerieFolios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folios"
   ClientHeight    =   1920
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5895
   Icon            =   "frmNuevaSerieFolios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFolioFinal 
      Height          =   285
      Left            =   4800
      MaxLength       =   9
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtFolioInicial 
      Height          =   285
      Left            =   2880
      MaxLength       =   9
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtIdentificador 
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Folio final"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   735
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Folio inicial"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Identificador"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Introduzca una nueva serie de folios para facturas"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmNuevaSerieFolios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intNumDeptoLoc As Integer
Dim strTipoDocumentoLoc As String
Dim strIdentificadorLoc As String
Dim lngInicioLoc As Long
Dim lngFinalLoc As Long
Dim msgResult As VbMsgBoxResult

Public Function fmsgPedirFolios(strTipoDocumento As String, intNumDepto As Integer, ByRef strIdentificador As String, ByRef lngInicio As Long, ByRef lngFinal As Long) As VbMsgBoxResult
    intNumDeptoLoc = intNumDepto
    strTipoDocumentoLoc = strTipoDocumento
    strIdentificadorLoc = strIdentificador
    lngInicioLoc = lngInicio
    lngFinalLoc = lngFinal
    Me.Show vbModal
    strIdentificador = strIdentificadorLoc
    lngInicio = lngInicioLoc
    lngFinal = lngFinalLoc
    fmsgPedirFolios = msgResult
    Unload Me
End Function

Private Sub cmdCancel_Click()
    msgResult = vbCancel
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If fblnDatosValidos Then
        msgResult = vbOK
        Me.Hide
    End If
End Sub

Private Function fblnDatosValidos() As Boolean
    
    Dim rsFoliosActivos As New ADODB.Recordset
    Dim rsFoliosMismoIdentificador As New ADODB.Recordset
    Dim vlstrTipoDocumento As String
    Dim vlstrIndicador As String
    Dim vlstrSentencia As String
    
    vlstrIndicador = Trim(txtIdentificador.Text)
    vlstrTipoDocumento = strTipoDocumentoLoc

    fblnDatosValidos = True
    
    If Trim(txtFolioInicial.Text) = "" Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        If txtFolioInicial.Enabled And txtFolioInicial.Visible Then
          txtFolioInicial.SetFocus
        Else
          If txtFolioFinal.Enabled And txtFolioFinal.Visible Then txtFolioFinal.SetFocus
        End If
    End If
    If fblnDatosValidos And Trim(txtFolioFinal.Text) = "" Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtFolioFinal.SetFocus
    End If
    If fblnDatosValidos And Val(txtFolioInicial.Text) <= 0 Then
        fblnDatosValidos = False
        'Dato incorrecto: El valor debe ser
        MsgBox SIHOMsg(36) & " mayor a cero", vbOKOnly + vbInformation, "Mensaje"
        If txtFolioInicial.Enabled And txtFolioInicial.Visible Then
          txtFolioInicial.SetFocus
        Else
          If txtFolioFinal.Enabled And txtFolioFinal.Visible Then txtFolioFinal.SetFocus
        End If
    End If
    If fblnDatosValidos And Val(txtFolioFinal.Text) < Val(txtFolioInicial.Text) Then
        fblnDatosValidos = False
        'El folio inicial debe ser menor o igual al final!
        MsgBox SIHOMsg(201), vbOKOnly + vbInformation, "Mensaje"
        If txtFolioInicial.Enabled And txtFolioInicial.Visible Then
          txtFolioInicial.SetFocus
        Else
          If txtFolioFinal.Enabled And txtFolioFinal.Visible Then txtFolioFinal.SetFocus
        End If
    End If

    
    If fblnDatosValidos Then
        vlstrSentencia = "select count(*) as Total from RegistroFolio " & _
        " where trim(chrCveDocumento) = '" & vlstrIndicador & "'" & _
        " and chrTipoDocumento = '" & vlstrTipoDocumento & "'" & _
        " and smiDepartamento <> " & intNumDeptoLoc & _
        " and intNumeroInicial <= " & txtFolioInicial.Text & _
        " and intNumeroFinal >= " & txtFolioInicial.Text
        Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
        If rsFoliosMismoIdentificador!Total <> 0 Then
            fblnDatosValidos = False
            '!Existe duplicidad en los folios!
            MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
            txtIdentificador.SetFocus
        End If
        rsFoliosMismoIdentificador.Close
    End If
    If fblnDatosValidos Then
        vlstrSentencia = "select count(*) as Total from RegistroFolio" & _
        " where trim(chrCveDocumento) = '" & vlstrIndicador & "'" & _
        " and chrTipoDocumento = '" & vlstrTipoDocumento & "'" & _
        " and smiDepartamento <> " & intNumDeptoLoc & _
        " and intNumeroInicial <= " & txtFolioFinal.Text & _
        " and intNumeroFinal >= " & txtFolioFinal.Text
        Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
        If rsFoliosMismoIdentificador!Total <> 0 Then
            fblnDatosValidos = False
            '!Existe duplicidad en los folios!
            MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
            txtIdentificador.SetFocus
        End If
        rsFoliosMismoIdentificador.Close
    End If
    If fblnDatosValidos Then
        vlstrSentencia = "select count(*) as Total from RegistroFolio " & _
        " where trim(chrCveDocumento) = '" & vlstrIndicador & "'" & _
        " and chrTipoDocumento = '" & vlstrTipoDocumento & "'" & _
        " and smiDepartamento = " & intNumDeptoLoc & _
        " and intNumeroInicial <= " & txtFolioInicial.Text & _
        " and intNumeroFinal >= " & txtFolioInicial.Text
        Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
        If rsFoliosMismoIdentificador!Total <> 0 Then
            fblnDatosValidos = False
            '!Existe duplicidad en los folios!
            MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
            txtIdentificador.SetFocus
        End If
        rsFoliosMismoIdentificador.Close
    End If
    If fblnDatosValidos Then
        vlstrSentencia = "select count(*) as Total from RegistroFolio" & _
        " where trim(chrCveDocumento) = '" & vlstrIndicador & "'" & _
        " and chrTipoDocumento = '" & vlstrTipoDocumento & "'" & _
        " and smiDepartamento = " & intNumDeptoLoc & _
        " and intNumeroInicial <= " & txtFolioFinal.Text & _
        " and intNumeroFinal >= " & txtFolioFinal.Text
        Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
        If rsFoliosMismoIdentificador!Total <> 0 Then
            fblnDatosValidos = False
            '!Existe duplicidad en los folios!
            MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
            txtIdentificador.SetFocus
        End If
        rsFoliosMismoIdentificador.Close
    End If
    If fblnDatosValidos Then
        strIdentificadorLoc = vlstrIndicador
        lngInicioLoc = Val(txtFolioInicial.Text)
        lngFinalLoc = Val(txtFolioFinal.Text)
    End If
End Function

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
End Sub

Private Sub txtFolioFinal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        cmdOK.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub txtFolioInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFolioFinal.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub txtIdentificador_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        txtFolioInicial.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
