VERSION 5.00
Begin VB.Form frmSelImpresora 
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   360
      Left            =   1230
      TabIndex        =   3
      Top             =   1125
      Width           =   1695
   End
   Begin VB.ComboBox cboDrvImpresoras 
      Height          =   315
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Width           =   4545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Predeterminar Printer"
      Height          =   360
      Left            =   3150
      TabIndex        =   0
      Top             =   1125
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Drivers de impresoras"
      Height          =   300
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   1950
   End
End
Attribute VB_Name = "frmSelImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim vPrinter As Printer
    Dim vlintCotador As Integer
    For Each vPrinter In Printers
        If vPrinter.DeviceName = cboDrvImpresoras.List(cboDrvImpresoras.ListIndex) Then
            ' Aqui pongo la asignación de la impresora
            ' y jupie!!!!
            Set Printer = vPrinter
        End If
    Next
    
End Sub

Private Sub Command2_Click()
    Printer.ScaleMode = vbCharacters
    Printer.CurrentY = 20
    Printer.CurrentX = 30
    Printer.Print "Nomas es una prueba"
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    Dim vPrinter As Printer
    Dim vlintCotador As Integer
    For Each vPrinter In Printers
        cboDrvImpresoras.AddItem vPrinter.DeviceName
    Next
End Sub
