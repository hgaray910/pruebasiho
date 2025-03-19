VERSION 5.00
Begin VB.Form frmMotivosCancelacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motivos de cancelación de facturas"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CboFolioUUID 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Folio Fiscal a Sustituir"
      Top             =   660
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   2360
      TabIndex        =   3
      Top             =   1090
      Width           =   1700
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   80
         TabIndex        =   2
         ToolTipText     =   "Confirmar la selección del motivo de cancelación del comprobante"
         Top             =   170
         Width           =   1560
      End
   End
   Begin VB.ComboBox CboMotivoCancelacion 
      Height          =   315
      Left            =   242
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Seleccionar el motivo de cancelación del comprobante"
      Top             =   180
      Width           =   5888
   End
End
Attribute VB_Name = "frmMotivosCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnActivaUUID As Boolean
Public intNumCliente As Integer
Public strdtmFechaHora As Date
Option Explicit

Private Sub CboMotivoCancelacion_Click()
    If CboMotivoCancelacion.ItemData(CboMotivoCancelacion.ListIndex) = 1 Then
        CboMotivoCancelacion.Top = 180
        CboFolioUUID.Visible = True
        'CboFolioUUID.Text = ""
        CboFolioUUID.ListIndex = -1
        CboFolioUUID.Top = 660
    Else
        CboMotivoCancelacion.Top = 480
        If CboFolioUUID.Visible Then
            'CboFolioUUID.Text = ""
            CboFolioUUID.ListIndex = -1
            CboFolioUUID.Visible = False
        End If
    End If
End Sub

Private Sub CboMotivoCancelacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdAceptar.Enabled Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
If CboMotivoCancelacion.ListIndex <> -1 Then
    vgMotivoCancelacion = ""
    vgMotivoCancelacion = "0" & CboMotivoCancelacion.ItemData(CboMotivoCancelacion.ListIndex)
    If CboFolioUUID.Visible Then
        If CboFolioUUID.ListIndex = -1 Then
           vgMotivoCancelacion = ""
           MsgBox "¡No ha seleccionado el folio fiscal para ser cancelado!", vbOKOnly + vbInformation, "Mensaje"
           CboFolioUUID.SetFocus
           Exit Sub
        End If
      vgstrFolioFiscalSustituye = ""
      vgstrFolioFiscalSustituye = Trim(Mid(CboFolioUUID.Text, InStr(CboFolioUUID.Text, "|") + 1))
    End If
    Unload Me
End If
End Sub

Private Sub cmdAceptar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = 27 Then
            vgMotivoCancelacion = ""
            vgstrFolioFiscalSustituye = ""
            Unload Me
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
Me.Icon = frmMenuPrincipal.Icon
    On Error GoTo NotificaError
       
        vgMotivoCancelacion = ""
        vgstrFolioFiscalSustituye = ""
        If pMotivoCancelacion(CboMotivoCancelacion, blnActivaUUID, CboFolioUUID, intNumCliente, strdtmFechaHora) = False Then
            CboMotivoCancelacion.Enabled = False
            cmdAceptar.Enabled = False
        End If
        
            CboMotivoCancelacion.Top = 480
            CboFolioUUID.Visible = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub
