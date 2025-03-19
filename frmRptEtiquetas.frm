VERSION 5.00
Begin VB.Form frmRptEtiquetas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraImprime 
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   2410
      Width           =   1155
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   580
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptEtiquetas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   80
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptEtiquetas.frx":03CD
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Vista previa"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hispanidad"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton opthispanidad 
         Caption         =   "Todos"
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton opthispanidad 
         Caption         =   "Viuda(o)"
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton opthispanidad 
         Caption         =   "Nieto de Español"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton opthispanidad 
         Caption         =   "Hijo de Español"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton opthispanidad 
         Caption         =   "Español"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRptEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vginttipo As Integer '0 = Etiquetas 1= Etiquetas Sepomex
Private vgrptReporte As CRAXDRT.Report

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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
    Unload Me
End Sub
Private Sub pImprime(vlstrTipo As String)

On Error GoTo NotificaError

    Dim rsSocios As New ADODB.Recordset
    Dim alstrParametros(4) As String
    Dim strParametros As String
    Dim vParametroHispanidad As String
   
If optHispanidad(0).Value = True Then
vParametroHispanidad = "ES"
ElseIf optHispanidad(1).Value = True Then
vParametroHispanidad = "HE"
ElseIf optHispanidad(2).Value = True Then
vParametroHispanidad = "NE"
ElseIf optHispanidad(3).Value = True Then
vParametroHispanidad = "VE"
ElseIf optHispanidad(4).Value = True Then
vParametroHispanidad = " "
End If


If Me.vginttipo = 0 Then 'SI ES 0 ENTONCES SE REFIERE A UN REPORTE SOLO DE ETIQUETAS

 pInstanciaReporte vgrptReporte, "rptEtiquetasCorreoOrdinario.rpt"
        
    vgrptReporte.DiscardSavedData
    
    strParametros = vParametroHispanidad
    
    Set rsSocios = frsEjecuta_SP(strParametros, "Sp_SORPTETQCORREOORDINARIO")
    
    If rsSocios.RecordCount <> 0 Then
            
           pImprimeReporte vgrptReporte, rsSocios, IIf(vlstrTipo = "P", "P", "I"), "Etiquetas para correo"
    
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    rsSocios.Close

Exit Sub



Else ' SI ES IGUAL A 1 SE REFIERE A UN REPORTE DE ETIQUETAS SEPOMEX

pInstanciaReporte vgrptReporte, "rptEtiquetasCorreoSepomex.rpt"
        
    vgrptReporte.DiscardSavedData
    
    strParametros = vParametroHispanidad
    
    Set rsSocios = frsEjecuta_SP(strParametros, "Sp_SORPTETIQUETASEPOMEX")
    
    If rsSocios.RecordCount <> 0 Then
            
           pImprimeReporte vgrptReporte, rsSocios, IIf(vlstrTipo = "P", "P", "I"), "Etiquetas SEPOMEX"
    
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    rsSocios.Close

Exit Sub


End If
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))

   End Sub
   
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    Select Case KeyCode
        
        Case 27
                
                    Unload Me
                    
        Case 13
            
                SendKeys vbTab
                        
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
End Sub
