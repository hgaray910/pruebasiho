VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReportesTrazabilidadVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trazabilidad de medicamentos en ventas a público"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   680
      Left            =   2280
      TabIndex        =   15
      Top             =   1920
      Width           =   1100
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   50
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReportesTrazabilidadVentas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Vista preliminar de la consulta"
         Top             =   140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   545
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReportesTrazabilidadVentas.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5415
      Begin VB.Frame Frame3 
         Caption         =   "Opciones de reporte"
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   5175
         Begin VB.CheckBox chkTodas 
            Caption         =   "Todos"
            Height          =   255
            Left            =   4080
            TabIndex        =   1
            Top             =   510
            Width           =   975
         End
         Begin VB.TextBox txtEtiqueta 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3885
         End
         Begin VB.Label lblDepto 
            Caption         =   "Etiqueta"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rango de fechas de la trazabilidad del medicamento"
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   5175
         Begin MSComCtl2.DTPicker dtpFechaInicio 
            Height          =   315
            Left            =   480
            TabIndex        =   2
            Top             =   300
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   129957891
            CurrentDate     =   38740
         End
         Begin MSComCtl2.DTPicker dtpFechaFin 
            Height          =   315
            Left            =   3000
            TabIndex        =   3
            Top             =   300
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   129957891
            CurrentDate     =   38740
         End
         Begin VB.Label Label2 
            Caption         =   "Del"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "al"
            Height          =   255
            Left            =   2640
            TabIndex        =   11
            Top             =   360
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Siguiente"
         Default         =   -1  'True
         Height          =   255
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReportesTrazabilidadVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vgrptReporte As CRAXDRT.Report

Private Sub chkTodas_Click()
    If chkTodas.Value Then
        txtEtiqueta.Text = "<TODOS>"
        txtEtiqueta.Enabled = False
    Else
        txtEtiqueta.Text = ""
        txtEtiqueta.Enabled = True
    End If
End Sub

Private Sub cmdImprimir_Click()
     pImprime "I"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSiguiente_Click()
    SendKeys vbTab
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub dtpFechaFin_GotFocus()
    dtpFechaFin.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub dtpFechaFin_LostFocus()
    dtpFechaFin.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub dtpFechaInicio_GotFocus()
    dtpFechaInicio.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub dtpFechaInicio_LostFocus()
    dtpFechaInicio.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub Form_Load()
    Dim dtmHoy As Date
    Me.Icon = frmMenuPrincipal.Icon
    pInstanciaReporte vgrptReporte, "rptTrazabilidadVentaPublico.rpt"
    dtmHoy = fdtmServerFecha
    Me.dtpFechaInicio.Value = dtmHoy
    Me.dtpFechaFin.Value = dtmHoy
End Sub

Private Sub pImprime(strDestino As String)
    Dim rs As New ADODB.Recordset
    Dim alstrParametros(2) As String
    Dim strTitulo As String
    Dim intTipo As Integer
    Dim strParametros As String
    
    Const strHoraInicial = " 00:00:00"
    Const strHoraFinal = " 23:59:59"
    
    If dtpFechaInicio.Value > dtpFechaFin.Value Then
        MsgBox SIHOMsg(64), vbExclamation, "Mensaje"
        dtpFechaInicio.SetFocus
    Else
        strTitulo = "Trazabilidad de medicamentos en ventas a público"
        intTipo = 1
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "Empresa;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(1) = "FechaI;" & UCase(Format(dtpFechaInicio.Value, "dd/MMM/yyyy"))
        alstrParametros(2) = "FechaF;" & UCase(Format(dtpFechaFin.Value, "dd/MMM/yyyy"))
        
        strParametros = IIf(chkTodas.Value, "-1", txtEtiqueta.Text) & "|" & Format(dtpFechaInicio.Value, "dd/MM/yyyy") & strHoraInicial & "|" & Format(dtpFechaFin.Value, "dd/MM/yyyy") & strHoraFinal
        
        Set rs = frsEjecuta_SP(strParametros, "SP_RPTVENTAPUBLICOLOTES")
        If Not rs.EOF Then
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rs, strDestino, strTitulo
        Else
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
            If chkTodas.Value = 0 Then
                txtEtiqueta.SetFocus
            End If
        End If
        rs.Close
    End If
End Sub

Private Sub txtEtiqueta_KeyPress(KeyAscii As Integer)
    ' Verificar si la tecla presionada es un número (0-9) o la tecla Backspace (8)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        ' Si no es un número o la tecla Backspace, cancelar la entrada
        KeyAscii = 0
    End If
End Sub
