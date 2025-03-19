VERSION 5.00
Begin VB.Form Frmingresoyfirmas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firma autorizada y requisitos al ingreso"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   75
      TabIndex        =   4
      Top             =   0
      Width           =   9975
      Begin VB.TextBox Txtrequisitos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Requisitos del ingreso"
         Top             =   2610
         Width           =   8175
      End
      Begin VB.TextBox txtempresa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Empresa"
         Top             =   240
         Width           =   8175
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Firma autorizada"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1455
      End
      Begin VB.Image picImagen 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   1680
         Stretch         =   -1  'True
         ToolTipText     =   "Firma de la empresa"
         Top             =   650
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Requisitos al ingreso "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   3
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.TextBox txtRutaImagen 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5040
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtClave 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5040
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Frmingresoyfirmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intcveempresa As Long
Public vchnombreempresa As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then
    Unload Me
End If

End Sub

Private Sub Form_Load()
 Dim vlstrSentencia As String
 Dim rsfirmas As New ADODB.Recordset
 Dim rsrequisitos As New ADODB.Recordset
 Dim stmImagen As New ADODB.Stream
 
 Me.Icon = frmMenuPrincipal.Icon
 vlstrSentencia = "SELECT CCFIRMASEMPRESA.BLBIMAGENFIRMA FROM CCFIRMASEMPRESA WHERE INTCVEEMPRESA=" & intcveempresa
 Set rsfirmas = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
 
    Set picImagen.Picture = LoadPicture("", vbLPLarge, vbLPColor)
    Txtrequisitos.Text = ""
    txtempresa.Text = ""
    txtempresa.Text = vchnombreempresa
    
    If rsfirmas.RecordCount > 0 Then
        stmImagen.Type = adTypeBinary
        stmImagen.Open
        stmImagen.Write rsfirmas!BLBIMAGENFIRMA
        stmImagen.SaveToFile App.Path & "\" & txtClave.Text, adSaveCreateOverWrite
        ' Retorna la imagen a la función
        Set picImagen.Picture = LoadPicture(App.Path & "\" & txtClave.Text, vbLPLarge, vbLPColor)
        txtRutaImagen.Text = App.Path & "\" & txtClave.Text
        vgblnFotoExistente = True
        stmImagen.Close
        
    Else
         vgblnFotoExistente = False
         Set picImagen.Picture = LoadPicture("", vbLPLarge, vbLPColor)
    End If
vlstrSentencia = "SELECT VCHREQUISITOSINGRESO FROM CCEMPRESA WHERE INTCVEEMPRESA=" & intcveempresa
Set rsrequisitos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)

If rsrequisitos.RecordCount > 0 Then
    If rsrequisitos!vchrequisitosingreso <> "" Then
        Txtrequisitos.Text = Trim(rsrequisitos!vchrequisitosingreso)
    Else
        Txtrequisitos.Text = ""
    End If
End If

End Sub

