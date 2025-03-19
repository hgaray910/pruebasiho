VERSION 5.00
Begin VB.Form frmMantoAsignacionCF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos para facturas parciales"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuita 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMantoAsignacionCF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2055
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdAsigna 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMantoAsignacionCF.frx":017A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1545
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.ListBox lstCFAsignados 
      Height          =   3375
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.ListBox lstCFDisponibles 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conceptos de facturación asignados"
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   75
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conceptos de facturación disponibles"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   75
      Width           =   2655
   End
End
Attribute VB_Name = "frmMantoAsignacionCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsigna_Click()
    pGrabaAsignacionCF
End Sub

Private Sub cmdQuita_Click()
    pQuitaAsignacion
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then 'Escape
        Unload Me
    End If
End Sub

Private Sub Form_Load()
        
    Me.Icon = frmMenuPrincipal.Icon
    
    pCargaCFDisponibles
    pCargaCFAsignados
    
End Sub

Private Sub pCargaCFDisponibles()
    Dim vlrsCFDisponible As New ADODB.Recordset
    
    Set vlrsCFDisponible = frsRegresaRs("Select PvConceptoFacturacion.SMICVECONCEPTO, PvConceptoFacturacion.CHRDESCRIPCION From PvConceptoFacturacion Where PvConceptoFacturacion.BITFACTURACIONPARCIAL = 0 Order by 2", adLockOptimistic, adOpenDynamic)
    lstCFDisponibles.Clear
    While Not vlrsCFDisponible.EOF
        lstCFDisponibles.AddItem RTrim(vlrsCFDisponible!chrDescripcion)
        lstCFDisponibles.ItemData(lstCFDisponibles.NewIndex) = vlrsCFDisponible!smiCveConcepto
        vlrsCFDisponible.MoveNext
    Wend
    If lstCFDisponibles.ListCount > 1 Then lstCFDisponibles.ListIndex = 0
    vlrsCFDisponible.Close
End Sub

Private Sub pCargaCFAsignados()
    Dim vlrsCFAsignados As New ADODB.Recordset
    Dim vlintCont As Integer
    
    Set vlrsCFAsignados = frsRegresaRs("Select PvConceptoFacturacion.SMICVECONCEPTO, PvConceptoFacturacion.CHRDESCRIPCION, PvConceptoFacturacion.BITFACTURACIONPARCIAL From PvConceptoFacturacion Where PvConceptoFacturacion.BITFACTURACIONPARCIAL <> 0 Order by 3 desc, 2", adLockOptimistic, adOpenDynamic)
    lstCFAsignados.Clear
    vlintCont = 0
    While Not vlrsCFAsignados.EOF
        lstCFAsignados.AddItem RTrim(vlrsCFAsignados!chrDescripcion)
        lstCFAsignados.ItemData(lstCFAsignados.NewIndex) = vlrsCFAsignados!smiCveConcepto
        If vlrsCFAsignados!BITFACTURACIONPARCIAL = 2 Then
            lstCFAsignados.List(vlintCont) = "* " & lstCFAsignados.List(vlintCont)
        End If
        vlintCont = vlintCont + 1
        vlrsCFAsignados.MoveNext
    Wend
    If lstCFAsignados.ListCount > 1 Then lstCFAsignados.ListIndex = 0
    vlrsCFAsignados.Close
End Sub


Private Sub pGrabaAsignacionCF()
    Dim vlstrSentencia As String

    If lstCFDisponibles.ListIndex > -1 Then
        vlstrSentencia = "Update PvConceptoFacturacion Set PvConceptoFacturacion.BITFACTURACIONPARCIAL = " & IIf(fblnExistePredeterminado, 1, 2) & " Where PvConceptoFacturacion.SMICVECONCEPTO = " & lstCFDisponibles.ItemData(lstCFDisponibles.ListIndex)
        pEjecutaSentencia vlstrSentencia
        pCargaCFAsignados
        pCargaCFDisponibles
    End If
End Sub

Private Sub pQuitaAsignacion()
    Dim vlstrSentencia As String

    If lstCFAsignados.ListIndex > -1 Then
        If Mid(lstCFAsignados.List(lstCFAsignados.ListIndex), 1, 1) = "*" Then
            If MsgBox("¿Esta seguro que desea quitar el concepto de facturación predeterminado?", vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                vlstrSentencia = "Update PvConceptoFacturacion Set PvConceptoFacturacion.BITFACTURACIONPARCIAL = 0 Where PvConceptoFacturacion.SMICVECONCEPTO = " & lstCFAsignados.ItemData(lstCFAsignados.ListIndex)
                pEjecutaSentencia vlstrSentencia
                pCargaCFDisponibles
                pCargaCFAsignados
            End If
        Else
            vlstrSentencia = "Update PvConceptoFacturacion Set PvConceptoFacturacion.BITFACTURACIONPARCIAL = 0 Where PvConceptoFacturacion.SMICVECONCEPTO = " & lstCFAsignados.ItemData(lstCFAsignados.ListIndex)
            pEjecutaSentencia vlstrSentencia
            pCargaCFDisponibles
            pCargaCFAsignados
        End If
    End If
End Sub


Private Function fblnExistePredeterminado() As Boolean
    Dim vlintCont As Integer
    
    fblnExistePredeterminado = False
    For vlintCont = 0 To lstCFAsignados.ListCount - 1
        If Mid(lstCFAsignados.List(vlintCont), 1, 1) = "*" Then
            fblnExistePredeterminado = True
            Exit Function
        End If
    Next
    
End Function

Private Sub lstCFAsignados_DblClick()
    Dim vlstrSentencia As String
    
    vlstrSentencia = "Update PvConceptoFacturacion Set PvConceptoFacturacion.BITFACTURACIONPARCIAL = 1 Where PvConceptoFacturacion.BITFACTURACIONPARCIAL <> 0"
    pEjecutaSentencia vlstrSentencia
    vlstrSentencia = "Update PvConceptoFacturacion Set PvConceptoFacturacion.BITFACTURACIONPARCIAL = 2 Where PvConceptoFacturacion.SMICVECONCEPTO = " & lstCFAsignados.ItemData(lstCFAsignados.ListIndex)
    pEjecutaSentencia vlstrSentencia
    pCargaCFAsignados

End Sub
