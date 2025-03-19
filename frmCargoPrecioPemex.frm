VERSION 5.00
Begin VB.Form frmCargoPrecioPemex 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de precios"
   ClientHeight    =   2640
   ClientLeft      =   150
   ClientTop       =   375
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPrecio 
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
      Left            =   6720
      MaxLength       =   12
      TabIndex        =   3
      ToolTipText     =   "Precio del artículo"
      Top             =   840
      Width           =   2325
   End
   Begin VB.Frame frmGrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   640
      Left            =   4440
      TabIndex        =   7
      Top             =   1890
      Width           =   560
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   495
         Left            =   25
         Picture         =   "frmCargoPrecioPemex.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Guardar el Registro"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame FraCaptura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtDescripcionLgaArt 
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
         Left            =   1800
         TabIndex        =   1
         ToolTipText     =   "Descripción del artículo"
         Top             =   240
         Width           =   7125
      End
      Begin VB.TextBox txtpreciounidosis 
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
         Left            =   6600
         MaxLength       =   12
         TabIndex        =   6
         ToolTipText     =   "Precio por unidosis"
         Top             =   1200
         Width           =   2325
      End
      Begin VB.TextBox txtCantidad 
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Cantidad de unidosis que contiene el artículo"
         Top             =   1200
         Width           =   2325
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1800
         MaxLength       =   18
         TabIndex        =   2
         ToolTipText     =   "Código del artículo"
         Top             =   720
         Width           =   2325
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Artículo cargo"
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
         Height          =   255
         Left            =   180
         TabIndex        =   12
         ToolTipText     =   "Nombre del articulo"
         Top             =   290
         Width           =   1620
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Precio unidosis"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   1280
         Width           =   1770
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Precio"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   800
         Width           =   1680
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cantidad"
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
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1280
         Width           =   1545
      End
      Begin VB.Label lblLab 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Código"
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
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   770
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmCargoPrecioPemex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : prjCaja
'| Nombre del Formulario    : frmCargoPrecioPemex
'-------------------------------------------------------------------------------------
'| Objetivo: Realizar la captura de datos a los precios por cago de pemex
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        :
'| Autor                    :
'| Fecha de Creación        : 01/feB/2024
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
Option Explicit

Private Declare Function shellexecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_NORMAL = 1


' Parametros iniciales de la forma
Public vlstrChrCveArticulo As String
Public vlstrArtDescripcion As String
Public vlstrCodigo As String
Public vlintPrecio As Double
Public vlintCantidad As String
Public vlintPreciounidosis As Double


Dim vlstrsql As String

Dim vlIntCont As Long

Dim alLotexArt() As varLotes
Dim vlstrStyle, vlstrResponse, MyString

Dim rs As New ADODB.Recordset


Private Sub cmdGrabarRegistro_Click()
On Error GoTo NotificaError

    Dim intCont, lngPersonaGraba As Long
    Dim strSentencia As String
    Dim lngVal As Long
   ' Dim lngVal As Integer
    'validar si la toma no ha sido aplicada aun, esto en el caso de que se tengan dos pantallas abiertas con la misma toma
    Dim ObjRs As New ADODB.Recordset
    Dim idetiqueta As String
    Dim VSTRPARAMETRO As String


    If fblnDatosValidos Then
           
        lngVal = 1
           

        lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If lngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans


            VSTRPARAMETRO = ""
            
            'es registro nuevo
            If Trim(vlstrChrCveArticulo) = "" Then
            
            VSTRPARAMETRO = Trim(txtCodigo.Text) & "'|" & _
                             Trim(txtDescripcionLgaArt.Text) & "'|" & _
                             Trim(txtPrecio.Text) & "'|" & _
                             Val(txtCantidad.Text) & "'|" & _
                             Trim(txtpreciounidosis)
                             
              frsEjecuta_SP VSTRPARAMETRO, "SP_PVINSLISTAPRECIOPEMEX", True
            
            
            Else ' registro existente
            
            VSTRPARAMETRO = Trim(vlstrChrCveArticulo) & "'|" & _
                            Trim(txtCodigo.Text) & "'|" & _
                             Trim(txtDescripcionLgaArt.Text) & "'|" & _
                             Trim(txtPrecio.Text) & "'|" & _
                             Val(txtCantidad.Text) & "'|" & _
                             Trim(txtpreciounidosis)
                             
              frsEjecuta_SP VSTRPARAMETRO, "SP_PVUPDLISTAPRECIOPEMEX", True
            End If
                          
               

            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, lngPersonaGraba, "SISTEMA ETIQUETADO", CStr(vgintNumeroDepartamento))
            EntornoSIHO.ConeccionSIHO.CommitTrans
                
              vlstrArtDescripcion = ""
     vlstrCodigo = ""
     vlintPrecio = 0
     vlintCantidad = 0
    vlintPreciounidosis = 0
            '|La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                   

                   
            Unload Me
                   
    
         End If

    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError

      
    
    If vlstrArtDescripcion = "" Then
        txtDescripcionLgaArt.Text = ""
        txtCodigo.Text = ""
        txtPrecio.Text = ""
        txtCantidad.Text = ""
        txtpreciounidosis.Text = ""
        'txtDescripcionLgaArt.SetFocus
    Else
        txtDescripcionLgaArt = vlstrArtDescripcion
        txtCodigo = vlstrCodigo
        txtPrecio = vlintPrecio
        txtCantidad = vlintCantidad
        txtpreciounidosis = vlintPreciounidosis
        'txtDescripcionLgaArt.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub



Private Sub Form_Initialize()
        'txtDescripcionLgaArt.SetFocus
        'txtCodigo.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        'If Val(txtCantRecep.Text) <> 0 And txtCantRecep.Text <> "" Then
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg("17"), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                vgblnCapturoLoteYCaduc = False
                Unload Me
            Else
                'pEnfocaTextBox txtFechaElaboracion
                txtDescripcionLgaArt.SetFocus
            End If
'        Else
'            vgblnCapturoLoteYCaduc = False
'            Unload Me
'        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError

    
    Dim rsAux As New ADODB.Recordset
    Dim blnSinViaAdministracion As Boolean
    Dim intcontador As Integer
    fblnDatosValidos = True
   
   
   
    If Trim(Replace(txtDescripcionLgaArt.Text, vbCrLf, "")) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
        txtDescripcionLgaArt.Text = Trim(Replace(txtDescripcionLgaArt.Text, vbCrLf, ""))
        txtDescripcionLgaArt.SetFocus
        Exit Function
    End If
    If Trim(Replace(txtCodigo.Text, vbCrLf, "")) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
        txtCodigo.Text = Trim(Replace(txtCodigo.Text, vbCrLf, ""))
        txtCodigo.SetFocus
        Exit Function
    End If
    If Trim(Replace(txtPrecio.Text, vbCrLf, "")) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
        'txtFormaFarmaceutica.Text = Trim(Replace(FormaFarmaceutica.Text, vbCrLf, ""))
        txtPrecio.SetFocus
        Exit Function
    End If

   
    If Trim(Replace(txtCantidad.Text, vbCrLf, "")) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
        txtCantidad.Text = Trim(Replace(txtCantidad.Text, vbCrLf, ""))
        txtCantidad.SetFocus
        Exit Function
    End If
    If Trim(Replace(txtpreciounidosis.Text, vbCrLf, "")) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
        txtpreciounidosis.Text = Trim(Replace(txtpreciounidosis.Text, vbCrLf, ""))
        txtpreciounidosis.SetFocus
        Exit Function
    End If
    
    

    
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function












Public Sub pValidaSoloNumero(KeyAscii As Integer)

  On Error GoTo NotificaError
  ' Solo permite números
    
    Select Case KeyAscii
        Case Asc(vbCr)
            KeyAscii = 0
        Case 8
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " :modProcedimientos " & ":pValidaSoloNumero"))
End Sub




Private Sub Form_Load()
 'txtDescripcionLgaArt = vlstrArtDescripcion
    Me.Icon = frmMenuPrincipal.Icon
'   txtCodigo.SetFocus
'     If vlstrArtDescripcion = "" Then
'        txtDescripcionLgaArt.Text = ""
'        txtCodigo.Text = ""
'        txtPrecio.Text = ""
'        txtCantidad.Text = ""
'        txtpreciounidosis.Text = ""
'        'txtDescripcionLgaArt.SetFocus
'    Else
'        txtDescripcionLgaArt = vlstrArtDescripcion
'        txtCodigo = vlstrCodigo
'        txtPrecio = vlintPrecio
'        txtCantidad = vlintCantidad
'        txtpreciounidosis = vlintPreciounidosis
'       ' txtDescripcionLgaArt.SetFocus
'    End If

End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtpreciounidosis.SetFocus
    End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
     If Not IsNumeric(Chr$(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
    
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPrecio.SetFocus
    End If
End Sub



Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
     Validar txtCodigo.Text, KeyAscii
End Sub

Private Sub txtDescripcionLgaArt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCodigo.SetFocus
    End If
End Sub

Private Sub txtDescripcionLgaArt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCantidad.SetFocus
    End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    Validar txtPrecio.Text, KeyAscii
End Sub

Private Sub txtpreciounidosis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdGrabarRegistro.SetFocus
    End If
End Sub


 
Public Sub Validar(DatosActuales As String, Caracter As Integer)

  If Caracter = 8 Then Exit Sub

  If InStr("0123456789", Chr$(Caracter)) Then Exit Sub

  If Caracter = 46 And InStr(DatosActuales, ".") = 0 Then Exit Sub

  Caracter = 0
End Sub

Private Sub txtpreciounidosis_KeyPress(KeyAscii As Integer)
    Validar txtpreciounidosis.Text, KeyAscii
End Sub
