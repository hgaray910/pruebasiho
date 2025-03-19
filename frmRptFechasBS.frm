VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptFechasBS 
   Caption         =   "Concentrado de actividades"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmBotonera 
      Height          =   705
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1170
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptFechasBS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir el reporte"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   105
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptFechasBS.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Vista preliminar del reporte"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rango de fechas"
      Height          =   810
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3870
      Begin MSComCtl2.DTPicker dtmFecFin 
         Height          =   330
         Left            =   2385
         TabIndex        =   1
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   38058
      End
      Begin MSComCtl2.DTPicker dtmFecIni 
         Height          =   330
         Left            =   615
         TabIndex        =   2
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   38058
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "A"
         Height          =   195
         Left            =   2160
         TabIndex        =   4
         Top             =   390
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   390
         Width           =   210
      End
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   495
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRptFechasBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
  
    If KeyAscii = 27 Then
              Unload Me
    ElseIf KeyAscii = 13 Then
      pFocusNextControl Me, ActiveControl.TabIndex
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
  dtmFecIni = CDate("01/" & CStr(Month(fdtmServerFecha)) & "/" & CStr(Year(fdtmServerFecha)))
  dtmFecFin = DateAdd("d", -1, CDate("01/" & CStr(Month(fdtmServerFecha) + 1) & "/" & CStr(Year(fdtmServerFecha))))
End Sub
Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Sub pImprime(pstrDestino As String)
  With cryReport
    If .Status = 0 Then MsgBox "No ha terminado de imprimir"
    .Reset
    .WindowState = crptMaximized
    .WindowBorderStyle = crptFixedSingle
    .ReportFileName = App.Path & "\Trabajo social\" & vgstrBaseDatosUtilizada & "\rptActividades.rpt"
    .StoredProcParam(0) = Format(dtmFecIni, "DD/MM/YYYY")
    .StoredProcParam(1) = Format(dtmFecFin, "DD/MM/YYYY")
    .ParameterFields(0) = "Empresa;" & Trim(vgstrNombreHospitalCH) & ";TRUE"
    .Destination = IIf(pstrDestino = "I", crptToPrinter, crptToWindow)
    .Connect = fstrConexionCrystal
    .PrintReport
    .ReportFileName = ""
    If .LastErrorNumber > 0 Then MsgBox "OCURRIO UN ERROR EN EL REPORTEADOR(CR). " & vbCrLf & .LastErrorString, vbExclamation, "Mensaje"
  End With
End Sub


