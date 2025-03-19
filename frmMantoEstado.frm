VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmMantoEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de estados"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstab 
      Height          =   3945
      Left            =   -10
      TabIndex        =   13
      Top             =   -10
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   6959
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoEstado.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoEstado.frx":001C
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
         Height          =   2320
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   7095
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
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   0
            ToolTipText     =   "Número"
            Top             =   300
            Width           =   825
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
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   1
            ToolTipText     =   "Nombre"
            Top             =   700
            Width           =   4800
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
            Left            =   2160
            TabIndex        =   4
            ToolTipText     =   "Activo / inactivo"
            Top             =   1960
            Width           =   930
         End
         Begin VB.TextBox txtNombreAbreviado 
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
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   2
            ToolTipText     =   "Nombre abreviado"
            Top             =   1110
            Width           =   1485
         End
         Begin HSFlatControls.MyCombo cboPais 
            Height          =   375
            Left            =   2160
            TabIndex        =   3
            ToolTipText     =   "Pais"
            Top             =   1515
            Width           =   4815
            _ExtentX        =   8493
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
            Left            =   195
            TabIndex        =   15
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
            Left            =   195
            TabIndex        =   16
            Top             =   760
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "País"
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
            Left            =   195
            TabIndex        =   20
            Top             =   1580
            Width           =   375
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000005&
            Caption         =   "Nombre abreviado"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   195
            TabIndex        =   19
            Top             =   1170
            Width           =   1905
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
         Height          =   3420
         Left            =   -75000
         TabIndex        =   18
         Top             =   -40
         Width           =   7270
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdEstados 
            Height          =   3030
            Left            =   75
            TabIndex        =   12
            ToolTipText     =   "Estados"
            Top             =   150
            Width           =   7180
            _ExtentX        =   12674
            _ExtentY        =   5345
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
            GridLinesFixed  =   1
            SelectionMode   =   1
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
         Left            =   1530
         TabIndex        =   17
         Top             =   2280
         Width           =   4320
         Begin MyCommandButton.MyButton cmdTop 
            Height          =   600
            Left            =   60
            TabIndex        =   5
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
            Picture         =   "frmMantoEstado.frx":0038
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoEstado.frx":09BA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBack 
            Height          =   600
            Left            =   660
            TabIndex        =   6
            ToolTipText     =   "Registro anterior"
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
            Picture         =   "frmMantoEstado.frx":133C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoEstado.frx":1CBE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdLocate 
            Height          =   600
            Left            =   1260
            TabIndex        =   7
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
            Picture         =   "frmMantoEstado.frx":2640
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoEstado.frx":2FC4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdNext 
            Height          =   600
            Left            =   1860
            TabIndex        =   8
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
            Picture         =   "frmMantoEstado.frx":3948
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoEstado.frx":42CA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdEnd 
            Height          =   600
            Left            =   2460
            TabIndex        =   9
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
            Picture         =   "frmMantoEstado.frx":4C4C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoEstado.frx":55CE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSave 
            Height          =   600
            Left            =   3060
            TabIndex        =   10
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
            Picture         =   "frmMantoEstado.frx":5F50
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoEstado.frx":68D4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   3660
            TabIndex        =   11
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
            Picture         =   "frmMantoEstado.frx":7258
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoEstado.frx":7BDA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMantoEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Programa para dar mantenimiento al catálogo de estados
' Fecha de inicio de desarrollo: Septiembre, 2003
'------------------------------------------------------------------------------
' Ultimas modificaciones al módulo, especificar:
' Fecha:
' Descripcion del cambio:
'------------------------------------------------------------------------------

Option Explicit

Public vllngNumeroOpcionEstado As Long          'Número de opción del módulo que usa el catálogo
Public vllngNumeroOpcionPais As Long            'Número de opción del módulo que usa el catálogo


Dim rsEstado As New ADODB.Recordset

Dim vlblnConsulta As Boolean

Dim vlstrSentencia As String



Private Sub cboPais_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboPais_GotFocus"))
End Sub

Private Sub cboPais_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    Dim vllngUltimaClave As Long
    
    If KeyAscii = 13 Then
        If cboPais.ListIndex <> -1 Then
            If cboPais.ItemData(cboPais.ListIndex) = 0 Then
            
                frmMantoPais.vllngNumeroOpcionPais = vllngNumeroOpcionPais
                frmMantoPais.Show vbModal, Me
                
                pLlenaPais
            
                vllngUltimaClave = frsRegresaRs("select isnull(max(intCvePais),0) from Pais").Fields(0)
                cboPais.ListIndex = flngLocalizaCbo_new(cboPais, STR(vllngUltimaClave))
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboPais_KeyPress"))
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
    
    If Not rsEstado.BOF Then
        rsEstado.MovePrevious
    End If
    If rsEstado.BOF Then
        rsEstado.MoveNext
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ValidaIntegridad
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionEstado, "C") Then
        rsEstado.Delete
        rsEstado.Update
        
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "ESTADO", txtNumero.Text)
        txtNumero.SetFocus
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If
    
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
    
    rsEstado.MoveLast
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    sstab.Tab = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError
    
    If Not rsEstado.EOF Then
        rsEstado.MoveNext
    End If
    If rsEstado.EOF Then
        rsEstado.MovePrevious
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionEstado, "E", True) Then
        If fblnDatosValidos() Then
            With rsEstado
                If Not vlblnConsulta Then
                    .AddNew
                End If
                !vchDescripcion = Trim(txtDescripcion.Text)
                !vchNombreAbreviado = IIf(Trim(txtNombreAbreviado.Text) = "", " ", Trim(txtNombreAbreviado.Text))
                !INTCVEPAIS = cboPais.ItemData(cboPais.ListIndex)
                !bitactivo = IIf(chkActivo.Value = 1, 1, 0)
                .Update
                If Not vlblnConsulta Then
                    txtNumero.Text = flngObtieneIdentity("SEC_ESTADO", rsEstado!INTCVEESTADO)
                End If
                pGuardarLogTransaccion Me.Name, IIf(vlblnConsulta, EnmCambiar, EnmGrabar), vglngNumeroLogin, "ESTADO", txtNumero.Text
            End With
            rsEstado.Requery
            txtNumero.SetFocus
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
    
    rsEstado.MoveFirst
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If sstab.Tab = 1 Then
            vlblnConsulta = False
            sstab.Tab = 0
        Else
            If vlblnConsulta Or cmdSave.Enabled Then
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    txtNumero.SetFocus
                End If
            Else
                Unload Me
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    'Color de Tab
    SetStyle sstab.hwnd, 0
    SetSolidColor sstab.hwnd, 16777215
    SSTabSubclass sstab.hwnd
    
    
    
    '---------------------------------------------------------
    ' Recordsets tipo tabla
    '---------------------------------------------------------
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlstrSentencia = "select * from Estado order by intCveEstado"
    Set rsEstado = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    pLlenaPais
    
    sstab.Tab = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pLlenaPais()
    On Error GoTo NotificaError
    
    Dim rsPais As New ADODB.Recordset

    vlstrSentencia = "select intCvePais, vchDescripcion from Pais where bitActivo = 1"
    Set rsPais = frsRegresaRs(vlstrSentencia)
    
    pLlenarCboRs_new cboPais, rsPais, 0, 1, 1
    cboPais.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaNacionalidad"))
End Sub
Private Sub grdEstados_DblClick()
    On Error GoTo NotificaError
    Dim vllngColumnaActual As Long
    
      With grdEstados
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
                pOrdColMshFGrid grdEstados, vgintTipoOrd
                pDesSelMshFGrid grdEstados
                
                .Col = vllngColumnaActual
                .Row = 1
            Else
                If fintLocalizaPkRs(rsEstado, 0, STR(grdEstados.RowData(grdEstados.Row))) <> 0 Then
                    pMuestra
                    sstab.Tab = 0
                    pHabilita 1, 1, 1, 1, 1, 0, 1
                    cmdLocate.SetFocus
                Else
                    Unload Me
                End If
            End If
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEstados_DblClick"))
End Sub

Private Sub grdEstados_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If fintLocalizaPkRs(rsEstado, 0, STR(grdEstados.RowData(grdEstados.Row))) <> 0 Then
            pMuestra
            sstab.Tab = 0
            pHabilita 1, 1, 1, 1, 1, 0, 1
            cmdLocate.SetFocus
        Else
            Unload Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdEstados_KeyPress"))
End Sub

Private Sub sstab_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If sstab.Tab = 0 Then
        If Not vlblnConsulta Then
            If txtNumero.Enabled Then
                txtNumero.SetFocus
            End If
        End If
    End If
    If sstab.Tab = 1 Then
        grdEstados.SetFocus
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
        txtNombreAbreviado.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))
End Sub

Private Sub txtNombreAbreviado_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtNombreAbreviado
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNombreAbreviado_GotFocus"))
End Sub

Private Sub txtNombreAbreviado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cboPais.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNombreAbreviado_KeyDown"))
End Sub

Private Sub txtNombreAbreviado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNombreAbreviado_KeyPress"))
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
    txtNombreAbreviado.Text = ""
    cboPais.ListIndex = 0
    chkActivo.Value = 1
    
    grdEstados.Clear
    grdEstados.Rows = 2
    grdEstados.Cols = 6
    pGrid
    
    If rsEstado.RecordCount = 0 Then
        pHabilita 0, 0, 0, 0, 0, 0, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, 0
        pLlenarMshFGrdRs grdEstados, rsEstado, 0
        pGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Function flngSiguiente() As Long
    On Error GoTo NotificaError
    
    Dim rsSiguienteNumero As New ADODB.Recordset
    
    vlstrSentencia = "select isnull(max(intCveEstado),0)+1 from Estado"
    Set rsSiguienteNumero = frsRegresaRs(vlstrSentencia)
    
    flngSiguiente = rsSiguienteNumero.Fields(0)


Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngSiguiente"))
End Function

Private Sub pGrid()
    On Error GoTo NotificaError
    
    With grdEstados
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Número|Nombre"
        .ColWidth(0) = 100
        .ColWidth(1) = 1000     'Numero
        .ColWidth(2) = 5750     'Nombre
        .ColWidth(3) = 0        'Pais
        .ColWidth(4) = 0        'Activo
        .ColWidth(5) = 0        'Nombre abreviado
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
        
        If fintLocalizaPkRs(rsEstado, 0, txtNumero.Text) = 0 Then
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
    
    txtNumero.Text = rsEstado!INTCVEESTADO
    txtDescripcion.Text = Trim(rsEstado!vchDescripcion)
    txtNombreAbreviado.Text = IIf(IsNull(rsEstado!vchNombreAbreviado), " ", Trim(rsEstado!vchNombreAbreviado))
    cboPais.ListIndex = flngLocalizaCbo_new(cboPais, STR(rsEstado!INTCVEPAIS))
    chkActivo.Value = IIf(rsEstado!bitactivo Or rsEstado!bitactivo = 1, 1, 0)
    
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
    If fblnDatosValidos And cboPais.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboPais.SetFocus
    End If
    If fblnDatosValidos Then
        If cboPais.ItemData(cboPais.ListIndex) = 0 Then
            fblnDatosValidos = False
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            cboPais.SetFocus
        End If
    End If
    
    If fblnDatosValidos Then
        
        vlstrSentencia = "select count(*) from Estado where vchDescripcion = '" & Trim(txtDescripcion.Text) & "'  and intCveEstado <> " & IIf(vlblnConsulta, txtNumero.Text, "0")
        
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


