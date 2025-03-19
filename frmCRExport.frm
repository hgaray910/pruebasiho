VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmCRExport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar a PDF"
   ClientHeight    =   1995
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4545
   Icon            =   "frmCRExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton CancelButton 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Cancelar"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Aceptar"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin HSFlatControls.MyCombo cboDestino 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Style           =   1
      Enabled         =   -1  'True
      Text            =   ""
      Sorted          =   -1  'True
      List            =   $"frmCRExport.frx":000C
      ItemData        =   $"frmCRExport.frx":0025
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Páginas"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3015
      Begin VB.OptionButton optTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Todas"
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
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   260
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optRango 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rango"
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
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   260
         Width           =   1095
      End
      Begin VB.TextBox txtIni 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtFin 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "De:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "A:"
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
         Left            =   1680
         TabIndex        =   9
         Top             =   660
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "&Destino"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmCRExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim m_retValue As VbMsgBoxResult

Public Function ShowForm(Owner As Object, ByRef cryDestino As CRExportDestinationType, ByRef strIni As String, strFin As String) As VbMsgBoxResult
    Me.Show vbModal, Owner
    Select Case cboDestino.ListIndex
        Case 0
            cryDestino = crEDTApplication
        Case 1
            cryDestino = crEDTDiskFile
    End Select
    If IsNumeric(txtIni.Text) And IsNumeric(txtFin.Text) Then
        strIni = txtIni.Text
        strFin = txtFin.Text
    End If
    ShowForm = m_retValue
    Unload Me
End Function
Private Sub CancelButton_Click()
    m_retValue = vbCancel
    Me.Hide
End Sub

Private Sub Form_Load()
    cboDestino.ListIndex = 1
    optTodas_Click
End Sub

Private Sub OKButton_Click()
    m_retValue = vbOK
    Me.Hide
End Sub

Private Sub optRango_Click()
    txtIni.Enabled = True
    txtFin.Enabled = True
    txtIni.BackColor = vbWindowBackground
    txtFin.BackColor = vbWindowBackground
End Sub

Private Sub optTodas_Click()
    txtIni.Text = ""
    txtFin.Text = ""
    txtIni.Enabled = False
    txtFin.Enabled = False
    txtIni.BackColor = vbInactiveBorder
    txtFin.BackColor = vbInactiveBorder
End Sub

Private Sub txtIni_KeyPress(KeyAscii As Integer)
    Call pValidaSoloNumero(KeyAscii)
End Sub
