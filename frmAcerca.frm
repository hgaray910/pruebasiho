VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmAcerca 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Acerca de Sistema Integral Hospitalario"
   ClientHeight    =   8265
   ClientLeft      =   2295
   ClientTop       =   1560
   ClientWidth     =   8085
   ClipControls    =   0   'False
   Icon            =   "frmAcerca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton cmdManual 
      Height          =   675
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Manual del módulo"
      Top             =   7320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1191
      ForeColor       =   14710016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      AppearanceThemes=   2
      BorderColor     =   16777215
      BorderDrawEvent =   1
      TransparentColor=   16777215
      Caption         =   "Manual del módulo"
      DepthEvent      =   1
      ForeColorDisabled=   14710016
      ForeColorOver   =   14710016
      ForeColorFocus  =   14710016
      ForeColorDown   =   14710016
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000015&
      X1              =   8070
      X2              =   8070
      Y1              =   0
      Y2              =   8260
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000015&
      X1              =   0
      X2              =   8070
      Y1              =   8250
      Y2              =   8250
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000015&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8250
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      X1              =   0
      X2              =   8070
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   300
      Picture         =   "frmAcerca.frx":000C
      Top             =   840
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   6720
      Picture         =   "frmAcerca.frx":6050
      Top             =   0
      Width           =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D3D3D3&
      X1              =   360
      X2              =   7680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D3D3D3&
      X1              =   360
      X2              =   7680
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Casa de Software"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   225
      Left            =   1080
      TabIndex        =   18
      Top             =   3615
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chihuahua, Chih."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   1080
      TabIndex        =   17
      Top             =   3855
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. +52 (614) 413 4748"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   900
      TabIndex        =   16
      Top             =   4080
      Width           =   1725
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nextel ID: 62*330740*1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   900
      TabIndex        =   15
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Label lblHospital 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   5760
      Width           =   4185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.sihosys.com"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   3270
      TabIndex        =   13
      Top             =   4320
      Width           =   1410
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nextel ID: 62*330740*5"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   5310
      TabIndex        =   12
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. +52 (55) 5514-6000"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   5325
      TabIndex        =   11
      Top             =   4080
      Width           =   1740
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "México, D.F."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   5790
      TabIndex        =   10
      Top             =   3855
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina Comercial"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   225
      Left            =   5460
      TabIndex        =   9
      Top             =   3615
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro Público del Derecho de Autor No. 03-2007-091110233500-01"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   225
      Left            =   1260
      TabIndex        =   8
      Top             =   3120
      Width           =   5490
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de información para hospitales y servicios clínicos"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   1680
      TabIndex        =   7
      Top             =   3360
      Width           =   4620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   405
      TabIndex        =   6
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label lblUbicacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00111111&
      Height          =   195
      Left            =   1095
      TabIndex        =   4
      Top             =   3780
      Width           =   45
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comercialización de Sistemas de Informática, S.A. de C.V."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   225
      Left            =   1680
      TabIndex        =   3
      Top             =   2880
      Width           =   4590
   End
   Begin VB.Label lblautorizacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Se autoriza el uso de este sistema a"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   5760
      Width           =   3060
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   240
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAcerca.frx":C094
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00111111&
      Height          =   840
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   7485
   End
End
Attribute VB_Name = "frmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Admisión
'| Nombre del Formulario    : frmAcerca
'-------------------------------------------------------------------------------------
'| Objetivo: Muestra información de los datos de la empresa que desarrolla el software
'|y que empresa tiene derecho de uso del mismo
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Nery Lozano - Luis Astudillo
'| Autor                    : Nery Lozano - Luis Astudillo
'| Fecha de Creación        : 11/Diciembre/1999
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit

Private Declare Function shellexecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdManual_Click()
    Dim vlstrRutaManual As String   'Ruta de reportes configurada en siparametro
    Dim rsDatos As ADODB.Recordset  'Recordset
    Dim vlstrSentencia As String    'Query para sacar la ruta de reportes
    Dim vlstrNombreManual As String 'Nombre del manual del módulo
    
    vlstrSentencia = "Select vchValor " & _
                    "from SIPARAMETRO Where vchNombre='VCHRUTAREPORTES'"

    Set rsDatos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    With rsDatos
        If rsDatos.RecordCount > 0 Then
           vlstrRutaManual = IIf(IsNull(!vchvalor), "", !vchvalor)
        End If
    End With
    rsDatos.Close
        
' * * * * * * * * * * * * *  Banco de sangre  * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "BS" Then
    'Banco de sangre
    vlstrNombreManual = "Manual Banco de sangre.pdf"
    End If
    
' * * * * * * * *  Supervisión y estadísticas * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "SE" Then
    vlstrNombreManual = "Manual Administración y Estadísticas.pdf"
    End If
    
' * * * * * * * *  * * * * Sistemas * * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "SI" Then
    vlstrNombreManual = "Manual Sistemas.pdf"
    End If
    
' * * * * * * * *  * * * * Trabajo social * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "TS" Then
    vlstrNombreManual = "Manual Trabajo Social.pdf"
    End If
        
' * * * * * * * *  * * * * Cargos * * * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "CA" Then
    vlstrNombreManual = "Manual Cargos.pdf"
    End If
    
' * * * * * * * *  * * * * Dietología * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "DI" Then
    vlstrNombreManual = "Manual Dietología.pdf"
    End If
    
' * * * * * * * *  * * * * Nómina * * * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "NO" Then
    vlstrNombreManual = "Manual Nómina.pdf"
    End If
    
' * * * * * * * *  * * * * Contabilidad * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "CN" Then
    vlstrNombreManual = "Manual Contabilidad.pdf"
    End If
    
' * * * * * * * *  * * * * Cuentas por cobrar ó crédito * * * * * * * * * * * * * * *
    If cgstrModulo = "CC" Then
    vlstrNombreManual = "Manual Cuentas por cobrar.pdf"
    End If
    
' * * * * * * * *  * * * * Caja * * * * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "PV" Then
    vlstrNombreManual = "Manual Caja.pdf"
    End If
    
' * * * * * * * *  * * * * Expediente * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "EX" Then
    vlstrNombreManual = "Manual Expediente Clínico.pdf"
    End If
    
' * * * * * * * *  * * * * Cuentas por pagar * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "CP" Then
    vlstrNombreManual = "Manual Cuentas por pagar.pdf"
    End If
    
' * * * * * * * *  * * * * Almacén * * * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "IV" Then
    vlstrNombreManual = "Manual Almacén.pdf"
    End If
    
' * * * * * * * *  * * * * Compras * * * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "CO" Then
    vlstrNombreManual = "Manual Compras.pdf"
    End If
    
' * * * * * * * * * * * * Imagenología ó servicios auxiliares* * * * * * * * * * * * *
    If cgstrModulo = "IM" Then
    vlstrNombreManual = "Manual Servicios auxiliares.pdf"
    End If
    
' * * * * * * * *  * * * * Laboratorio * * * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "LA" Then
    vlstrNombreManual = "Manual Laboratorio.pdf"
    End If

' * * * * * * * * * * * * * * * * Admisión * * * * * * * * * * * * * * * * * * * * * *
    If cgstrModulo = "AD" Then
    vlstrNombreManual = "Manual Admisión.pdf"
    End If
    
    shellexecute Me.hwnd, "open", vlstrRutaManual & "\" & vlstrNombreManual, "", "", 4

End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyEscape
    Unload Me
  End Select
End Sub

Private Sub Form_Load()
  lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
  lblHospital.Caption = Trim(vgstrNombreHospitalCH)
  lblHospital.BorderStyle = 0
End Sub

