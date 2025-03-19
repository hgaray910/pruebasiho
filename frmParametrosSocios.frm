VERSION 5.00
Begin VB.Form frmParametrosSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de socios"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   1973
      TabIndex        =   3
      Top             =   1220
      Width           =   615
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000000&
         Picture         =   "frmParametrosSocios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Grabar parámetros"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.ComboBox cboFactura 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Formato de facturas"
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Formato de factura para membresías"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmParametrosSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFactura_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    Dim strSentencia As String
    Dim vllngPersonaGraba As Long
    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
     If Not fblnRevisaPermiso(vglngNumeroLogin, 2492, "E") Then Exit Sub
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    '|  Elimina el formato configurado actualmente, en caso de que exista
    strSentencia = "DELETE " & _
                   "  FROM PVDOCUMENTODEPARTAMENTO " & _
                   " WHERE SMIDEPARTAMENTO = " & CStr(vgintNumeroDepartamento) & _
                   "   AND INTNUMTIPOFORMATO = 9 " & _
                   "   AND CHRTIPOPACIENTE = 'S'"
    pEjecutaSentencia strSentencia
    
    '|  Si se seleccionó algún formato lo inserta
    If cboFactura.ListIndex > 0 Then
        vgstrParametrosSP = CStr(vgintNumeroDepartamento) & "|" & Str(cboFactura.ItemData(cboFactura.ListIndex)) & "|9|S"
        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSDOCUMENTODEPARTAMENTO"
    End If
    
    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "PARÁMETROS DEL SOCIO", CStr(vgintNumeroDepartamento))
    
    '------------------------------------------------------------------
    ' Darle COMMIT a la TRANSACTION
    '------------------------------------------------------------------
    EntornoSIHO.ConeccionSIHO.CommitTrans
    cboFactura.SetFocus
    cmdSave.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys vbTab
        Case vbKeyEscape
            Unload Me
    End Select

End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    pCargaFormatos
    cmdSave.Enabled = False
End Sub

Private Sub pCargaFormatos()
    Dim strSentencia As String
    Dim rsFormatos As New ADODB.Recordset
    
    strSentencia = "select intNumeroFormato,vchDescripcion from Formato where intNumeroTipoFormato=9"
    Set rsFormatos = frsRegresaRs(strSentencia)
    If rsFormatos.RecordCount <> 0 Then
        pLlenarCboRs cboFactura, rsFormatos, 0, 1, 4
    End If
    rsFormatos.Close
    
    cboFactura.ListIndex = 0
    strSentencia = "SELECT intNumFormato " & _
                   "  FROM PvDocumentoDepartamento " & _
                   " WHERE SMIDEPARTAMENTO = " & CStr(vgintNumeroDepartamento) & _
                   "   AND INTNUMTIPOFORMATO = 9 " & _
                   "   AND CHRTIPOPACIENTE = 'S'"
    Set rsFormatos = frsRegresaRs(strSentencia)
    If rsFormatos.RecordCount > 0 Then
        cboFactura.ListIndex = fintLocalizaCbo(cboFactura, rsFormatos!intNumFormato)
    End If
    rsFormatos.Close
    

End Sub


