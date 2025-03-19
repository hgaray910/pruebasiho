VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmMantoCiudad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de ciudades"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   3825
      Left            =   -10
      TabIndex        =   12
      Top             =   -10
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   6747
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483643
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoCiudad.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoCiudad.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1930
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   6195
         Begin VB.TextBox txtNumero 
            Alignment       =   1  'Right Justify
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
            Left            =   1170
            MaxLength       =   5
            TabIndex        =   0
            ToolTipText     =   "Número"
            Top             =   300
            Width           =   840
         End
         Begin VB.TextBox txtDescripcion 
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
            Left            =   1170
            MaxLength       =   75
            TabIndex        =   1
            ToolTipText     =   "Nombre"
            Top             =   700
            Width           =   4830
         End
         Begin VB.CheckBox chkActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Activo"
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
            Height          =   250
            Left            =   1170
            TabIndex        =   3
            ToolTipText     =   "Activo / inactivo"
            Top             =   1560
            Width           =   1530
         End
         Begin HSFlatControls.MyCombo cboEstado 
            Height          =   375
            Left            =   1170
            TabIndex        =   2
            Top             =   1110
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   ""
            ItemData        =   ""
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Número"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   190
            TabIndex        =   14
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   190
            TabIndex        =   15
            Top             =   760
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   190
            TabIndex        =   18
            Top             =   1170
            Width           =   660
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2940
         Left            =   -74920
         TabIndex        =   17
         Top             =   -60
         Width           =   6300
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCiudades 
            Height          =   2700
            Left            =   0
            TabIndex        =   11
            ToolTipText     =   "Ciudades"
            Top             =   165
            Width           =   6270
            _ExtentX        =   11060
            _ExtentY        =   4763
            _Version        =   393216
            ForeColor       =   0
            Rows            =   0
            FixedRows       =   0
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            AllowBigSelection=   0   'False
            GridLinesFixed  =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   1060
         TabIndex        =   16
         Top             =   1920
         Width           =   4320
         Begin MyCommandButton.MyButton cmdTop 
            Height          =   600
            Left            =   60
            TabIndex        =   4
            ToolTipText     =   "Primer registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmMantoCiudad.frx":0038
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoCiudad.frx":09BA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBack 
            Height          =   600
            Left            =   660
            TabIndex        =   5
            ToolTipText     =   "Anterior registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmMantoCiudad.frx":133C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoCiudad.frx":1CBE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdLocate 
            Height          =   600
            Left            =   1260
            TabIndex        =   6
            ToolTipText     =   "Búsqueda"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmMantoCiudad.frx":2640
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoCiudad.frx":2FC4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdNext 
            Height          =   600
            Left            =   1860
            TabIndex        =   7
            ToolTipText     =   "Siguiente registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmMantoCiudad.frx":3948
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoCiudad.frx":42CA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdEnd 
            Height          =   600
            Left            =   2460
            TabIndex        =   8
            ToolTipText     =   "Último registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmMantoCiudad.frx":4C4C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoCiudad.frx":55CE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSave 
            Height          =   600
            Left            =   3060
            TabIndex        =   9
            ToolTipText     =   "Grabar"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmMantoCiudad.frx":5F50
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoCiudad.frx":68D4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   3660
            TabIndex        =   10
            ToolTipText     =   "Eliminar registro"
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmMantoCiudad.frx":7258
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoCiudad.frx":7BDA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMantoCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Programa para dar mantenimiento al catálogo de estados
' Fecha de inicio de desarrollo: Septiembre, 2003
'------------------------------------------------------------------------------
' Ultimas modificaciones al módulo, especificar:
' Fecha: 17/Abril/2013
' Descripcion del cambio:
' 1.- Se agregó búsqueda de ciudad al teclear iniciales.
' 2.- Se agregó soporte para el scroll del grid
'------------------------------------------------------------------------------

Option Explicit

Public vllngNumeroOpcionCiudad As Long      'Número de opción del módulo que usa el catálogo
Public vllngNumeroOpcionEstado As Long      'Número de opción del módulo que usa el catálogo
Public vllngNumeroOpcionPais As Long        'Número de opción del módulo que usa el catálogo

Dim rsCiudad As New ADODB.Recordset
Dim vlblnConsulta As Boolean
Dim vlstrSentencia As String

Private Sub cboEstado_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstado_GotFocus"))
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    Dim vllngUltimaClave As Long
    
    If KeyAscii = 13 Then
        If cboEstado.ListIndex <> -1 Then
            If cboEstado.ItemData(cboEstado.ListIndex) = 0 Then
                frmMantoEstado.vllngNumeroOpcionEstado = vllngNumeroOpcionEstado
                frmMantoEstado.vllngNumeroOpcionPais = vllngNumeroOpcionPais
                frmMantoEstado.Show vbModal, Me
                
                pLlenaEstado
            
                vllngUltimaClave = frsRegresaRs("select isnull(max(intCveEstado),0) from Estado").Fields(0)
                cboEstado.ListIndex = flngLocalizaCbo_new(cboEstado, STR(vllngUltimaClave))
            Else
                If vlblnConsulta Then
                    chkActivo.SetFocus
                Else
                    cmdSave.SetFocus
                End If
            End If
        Else
            If vlblnConsulta Then
                chkActivo.SetFocus
            Else
                cmdSave.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstado_KeyPress"))
End Sub

Private Sub chkActivo_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_GotFocus"))
End Sub

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_KeyPress"))
End Sub

Private Sub cmdBack_Click()
On Error GoTo NotificaError
    
    If Not rsCiudad.BOF Then
        rsCiudad.MovePrevious
    End If
    If rsCiudad.BOF Then
        rsCiudad.MoveNext
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ValidaIntegridad
Dim sql As String
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCiudad, "C") Then
        
        'EntornoSIHO.ConeccionSIHO.BeginTrans
        
        rsCiudad.Delete
        pEjecutaSentencia "delete from ciudad where intcveciudad = " & txtNumero.Text
        
        'EntornoSIHO.ConeccionSIHO.CommitTrans
        
        'rsCiudad.Update
        
        pllenaCiudades
        
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "CIUDAD", txtNumero.Text)
        txtNumero.SetFocus
        MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
        pLimpia
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If
    
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        'EntornoSIHO.ConeccionSIHO.RollbackTrans

        pllenaCiudades
        pLimpia
    End If
End Sub

Private Sub cmdEnd_Click()
On Error GoTo NotificaError
    
    rsCiudad.MoveLast
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
On Error GoTo NotificaError
    
    SSTab.Tab = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
On Error GoTo NotificaError
    
    If Not rsCiudad.EOF Then
        rsCiudad.MoveNext
    End If
    If rsCiudad.EOF Then
        rsCiudad.MovePrevious
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    Dim vlblnAlta As Boolean

On Error GoTo NotificaError
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCiudad, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionCiudad, "C", True) Then
        If fblnDatosValidos() Then
            With rsCiudad
                If Not vlblnConsulta Then
                    .AddNew
                    vlblnAlta = True
                End If
                !vchDescripcion = Trim(txtDescripcion.Text)
                !INTCVEESTADO = cboEstado.ItemData(cboEstado.ListIndex)
                !BITACTIVA = IIf(chkActivo.Value = 1, 1, 0)
                .Update
                If vlblnAlta Then
                    txtNumero.Text = flngObtieneIdentity("SEC_CIUDAD", rsCiudad!intCveCiudad)
                    'txtNumero.Text = !intCveCiudad
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "CIUDAD", txtNumero.Text)
                    vlblnAlta = False
                Else
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "CIUDAD", txtNumero.Text)
                End If
            End With
            rsCiudad.Requery
            txtNumero.SetFocus
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
        End If
         
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub cmdTop_Click()
On Error GoTo NotificaError
    
    rsCiudad.MoveFirst
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        'If SSTab.Tab = 1 Then
        '    vlblnConsulta = False
        '    SSTab.Tab = 0
        'Else
        '    If vlblnConsulta Or cmdSave.Enabled Then
        '        '¿Desea abandonar la operación?
        '        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
        '            txtNumero.SetFocus
        '        End If
        '    Else
        '        Unload Me
        '    End If
        'End If
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    
    'Color de Tab
    SetStyle SSTab.hwnd, 0
    SetSolidColor SSTab.hwnd, 16777215
    SSTabSubclass SSTab.hwnd
    
    pllenaCiudades
    
    vllngNumeroOpcionCiudad = IIf(cgstrModulo = "SI", 1177, 726)
        
    pLlenaEstado
    
    SSTab.Tab = 0
    Me.Icon = frmMenuPrincipal.Icon

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pLlenaEstado()
On Error GoTo NotificaError
    
    Dim rsEstado As New ADODB.Recordset

    vlstrSentencia = "SELECT intCveEstado, vchDescripcion FROM Estado WHERE bitActivo = 1"
    Set rsEstado = frsRegresaRs(vlstrSentencia)
    
    pLlenarCboRs_new cboEstado, rsEstado, 0, 1, 1
    cboEstado.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaEstado"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError
    
    Select Case SSTab.Tab
        Case 0
            If vlblnConsulta Or cmdSave.Enabled Then
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                    Cancel = True
                    Call pEnfocaTextBox(txtNumero)
                Else
                    Cancel = True
                End If
            End If
        Case 1
            Cancel = True
            SSTab.Tab = 0
            Call pEnfocaTextBox(txtNumero)
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
End Sub

'-----------------------------------------------------------------------------'
' Refresca el grdCiudades y asigna bajo que columna se va a hacer la búsqueda '
'-----------------------------------------------------------------------------'
Private Sub grdCiudades_Click()
On Error GoTo NotificaError

    If grdCiudades.Rows = 0 Then Exit Sub
    grdCiudades.Refresh
    vgintColLoc = grdCiudades.Col
    vgstrAcumTextoBusqueda = ""
    grdCiudades.Col = vgintColLoc
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCiudades_Click"))
End Sub

Private Sub grdCiudades_DblClick()
On Error GoTo NotificaError
 Dim vllngColumnaActual As Long
 
With grdCiudades
            If (.MouseRow = 0) And (.MouseCol > 0) Then
                .Col = .MouseCol
                
                vllngColumnaActual = .Col
            
                vgintColOrd = .Col
                
                'Escoge el Tipo de Ordenamiento
                If vgintTipoOrd = 1 Then
                    vgintTipoOrd = 2
                Else
                    vgintTipoOrd = 1
                End If
                pOrdColMshFGrid grdCiudades, vgintTipoOrd
                pDesSelMshFGrid grdCiudades
                
                .Col = vllngColumnaActual
                .Row = 1
            Else
            
              If fintLocalizaPkRs(rsCiudad, 0, STR(grdCiudades.RowData(grdCiudades.Row))) <> 0 Then
                pMuestra
                SSTab.Tab = 0
                pHabilita 1, 1, 1, 1, 1, 0, 1
                cmdLocate.SetFocus
              Else
                 Unload Me
              End If
               
            End If
    End With
    
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCiudades_DblClick"))
End Sub

Private Sub grdCiudades_GotFocus()
    HookForm grdCiudades 'Permite utilizar el scroll en el grid
End Sub

Private Sub grdCiudades_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdCiudades_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCiudades_KeyDown"))
End Sub

Private Sub grdCiudades_KeyPress(KeyAscii As Integer)
'----------------------------------------------------------------------'
' Caso 7404: Evento que verifica si se presionó una tecla              '
' de la A-Z, a-z, 0-9, á,é,í,ó,ú,ñ,Ñ, se presionó la barra espaciadora '
' Realizando la búsqueda de un criterio dentro del grdCiudades         '
'----------------------------------------------------------------------'
 On Error GoTo NotificaError
 
    If grdCiudades.FixedRows = 0 Then Exit Sub
    Call pSelCriterioMshFGrid(grdCiudades, vgintColLoc, KeyAscii)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCiudades_KeyPress"))
End Sub

Private Sub grdCiudades_LostFocus()
    UnHookForm grdCiudades 'Libera el grid de las funciones del scroll
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
On Error GoTo NotificaError
    
    If SSTab.Tab = 0 Then
        If Not vlblnConsulta Then
            txtNumero.SetFocus
        End If
    End If
    
    If SSTab.Tab = 1 Then
        grdCiudades.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstab_Click"))
End Sub

Private Sub txtDescripcion_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescripcion

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_GotFocus"))
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cboEstado.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))
End Sub

Private Sub txtNumero_GotFocus()
On Error GoTo NotificaError
    
    pLimpia
    pSelTextBox txtNumero

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumero_GotFocus"))
End Sub

Private Sub pLimpia()
On Error GoTo NotificaError
    
    vlblnConsulta = False
    
    txtNumero.Text = flngSiguiente()
    txtDescripcion.Text = ""
    cboEstado.ListIndex = 0
    chkActivo.Value = 1
    
    grdCiudades.Clear
    grdCiudades.Rows = 2
    grdCiudades.Cols = 5
    pGrid
    
    If rsCiudad.RecordCount = 0 Then
        pHabilita 0, 0, 0, 0, 0, 0, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, 0
        pLlenarMshFGrdRs grdCiudades, rsCiudad, 0
        pGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Function flngSiguiente() As Long
On Error GoTo NotificaError
    
    Dim rsSiguienteNumero As New ADODB.Recordset
    
    vlstrSentencia = "select isnull(max(intCveCiudad),0)+1 from Ciudad"
    Set rsSiguienteNumero = frsRegresaRs(vlstrSentencia)
    
    flngSiguiente = rsSiguienteNumero.Fields(0)

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngSiguiente"))
End Function

Private Sub pGrid()
On Error GoTo NotificaError
    
    With grdCiudades
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Número|Nombre"
        .ColWidth(0) = 100
        .ColWidth(1) = 1000     'Numero
        .ColWidth(2) = 4800     'Nombre
        .ColWidth(3) = 0        'Pais
        .ColWidth(4) = 0        'Activo
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGridImpuestos"))
End Sub

Private Sub pHabilita(vlbln1 As Integer, vlbln2 As Integer, vlbln3 As Integer, vlbln4 As Integer, vlbln5 As Integer, vlbln6 As Integer, vlbln7 As Integer)
On Error GoTo NotificaError
    
    If vlbln1 = 1 Then
        cmdTop.Enabled = True
    Else
        cmdTop.Enabled = False
    End If
    If vlbln2 = 1 Then
        cmdBack.Enabled = True
    Else
        cmdBack.Enabled = False
    End If
    If vlbln3 = 1 Then
        cmdLocate.Enabled = True
    Else
        cmdLocate.Enabled = False
    End If
    If vlbln4 = 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
    If vlbln5 = 1 Then
        cmdEnd.Enabled = True
    Else
        cmdEnd.Enabled = False
    End If
    If vlbln6 = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    If vlbln7 = 1 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(txtNumero.Text) = "" Then
            txtNumero.Text = flngSiguiente()
        End If
        
        If fintLocalizaPkRs(rsCiudad, 0, txtNumero.Text) = 0 Then
            txtNumero.Text = flngSiguiente()
        Else
            pMuestra
        End If
        
        If vlblnConsulta Then
            pHabilita 1, 1, 1, 1, 1, 0, 1
            cmdTop.SetFocus
        Else
            pHabilita 0, 0, 0, 0, 0, 1, 0
            txtDescripcion.SetFocus
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumero_KeyPress"))
End Sub

Private Sub pMuestra()
On Error GoTo NotificaError
    
    vlblnConsulta = True
    
    txtNumero.Text = rsCiudad!intCveCiudad
    txtDescripcion.Text = Trim(rsCiudad!vchDescripcion)
    cboEstado.ListIndex = flngLocalizaCbo_new(cboEstado, STR(rsCiudad!INTCVEESTADO))
    chkActivo.Value = IIf(rsCiudad!BITACTIVA, 1, 0)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra"))
End Sub

Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError
    
    fblnDatosValidos = True
    
    If Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    
    If fblnDatosValidos And cboEstado.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboEstado.SetFocus
    End If
    
    If fblnDatosValidos Then
        If cboEstado.ItemData(cboEstado.ListIndex) = 0 Then
            fblnDatosValidos = False
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            cboEstado.SetFocus
        End If
    End If
    
    If fblnDatosValidos Then
        vlstrSentencia = "select count(*) from Ciudad where vchDescripcion = '" & Trim(txtDescripcion.Text) & "'  and intCveEstado = " & cboEstado.ItemData(cboEstado.ListIndex)
        If frsRegresaRs(vlstrSentencia).Fields(0) <> 0 Then
            fblnDatosValidos = False
            'Existe información con el mismo contenido
            MsgBox SIHOMsg(19), vbOKOnly + vbInformation, "Mensaje"
            txtDescripcion.SetFocus
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function

Private Sub pllenaCiudades()
On Error GoTo NotificaError
Dim vlstrSentencia As String
    
    '-----------------------'
    ' Recordsets tipo tabla '
    '-----------------------'
    vlstrSentencia = "SELECT * FROM Ciudad ORDER BY intCveCiudad"
    Set rsCiudad = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenaCiudades"))
End Sub
