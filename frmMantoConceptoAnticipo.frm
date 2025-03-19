VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmMantoConceptoAnticipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de atención médica"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabConceptos 
      Height          =   3825
      Left            =   -10
      TabIndex        =   13
      Top             =   -10
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   6747
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoConceptoAnticipo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoConceptoAnticipo.frx":001C
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
         Height          =   1920
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   7275
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
            Left            =   1560
            MaxLength       =   3
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
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   1
            ToolTipText     =   "Descripción"
            Top             =   700
            Width           =   5505
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
            Height          =   270
            Left            =   1560
            TabIndex        =   4
            Top             =   1560
            Width           =   1530
         End
         Begin VB.TextBox txtAnticipo 
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
            Left            =   5520
            MaxLength       =   30
            TabIndex        =   3
            ToolTipText     =   "Anticipo sugerido"
            Top             =   1110
            Width           =   1545
         End
         Begin HSFlatControls.MyCombo cboTratamiento 
            Height          =   375
            Left            =   1560
            TabIndex        =   2
            ToolTipText     =   "Tipo de tratamiento"
            Top             =   1110
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   -1  'True
            List            =   $"frmMantoConceptoAnticipo.frx":0038
            ItemData        =   $"frmMantoConceptoAnticipo.frx":0050
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
            Caption         =   "Descripción"
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
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tratamiento"
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
            TabIndex        =   17
            Top             =   1170
            Width           =   1170
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Anticipo sugerido"
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
            Left            =   3720
            TabIndex        =   20
            Top             =   1170
            Width           =   1695
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
         Height          =   2750
         Left            =   -74940
         TabIndex        =   19
         Top             =   60
         Width           =   7400
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConceptos 
            Height          =   2700
            Left            =   30
            TabIndex        =   12
            ToolTipText     =   "Conceptos"
            Top             =   45
            Width           =   7340
            _ExtentX        =   12938
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
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
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
         Left            =   1600
         TabIndex        =   18
         Top             =   1900
         Width           =   4320
         Begin MyCommandButton.MyButton cmdTop 
            Height          =   600
            Left            =   60
            TabIndex        =   5
            ToolTipText     =   "Primer concepto"
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
            Picture         =   "frmMantoConceptoAnticipo.frx":005A
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoConceptoAnticipo.frx":09DC
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBack 
            Height          =   600
            Left            =   660
            TabIndex        =   6
            ToolTipText     =   "Anterior concepto"
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
            Picture         =   "frmMantoConceptoAnticipo.frx":135E
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoConceptoAnticipo.frx":1CE0
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
            Picture         =   "frmMantoConceptoAnticipo.frx":2662
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoConceptoAnticipo.frx":2FE6
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdNext 
            Height          =   600
            Left            =   1860
            TabIndex        =   8
            ToolTipText     =   "Siguiente concepto"
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
            Picture         =   "frmMantoConceptoAnticipo.frx":396A
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoConceptoAnticipo.frx":42EC
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdEnd 
            Height          =   600
            Left            =   2460
            TabIndex        =   9
            ToolTipText     =   "Último concepto"
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
            Picture         =   "frmMantoConceptoAnticipo.frx":4C6E
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoConceptoAnticipo.frx":55F0
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
            Picture         =   "frmMantoConceptoAnticipo.frx":5F72
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoConceptoAnticipo.frx":68F6
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   3660
            TabIndex        =   11
            ToolTipText     =   "Eliminar concepto"
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
            Picture         =   "frmMantoConceptoAnticipo.frx":727A
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoConceptoAnticipo.frx":7BFC
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMantoConceptoAnticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Programa para dar mantenimiento al catálogo de conceptos de atención médica
' Fecha de inicio de desarrollo: 06/Agosto/2003
' Autor:                         Rosenda Hernández Anaya
'------------------------------------------------------------------------------
' Ultimas modificaciones al módulo, especificar:
' Fecha:
' Descripcion del cambio:
'------------------------------------------------------------------------------

Option Explicit

Dim rsAdConceptoAnticipo  As New ADODB.Recordset

Dim vlblnConsulta As Boolean

Dim vlstrx As String

Private Sub cboTratamiento_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTratamiento_GotFocus"))
End Sub

Private Sub cboTratamiento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtAnticipo.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTratamiento_KeyPress"))
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
    
    If Not rsAdConceptoAnticipo.BOF Then
        rsAdConceptoAnticipo.MovePrevious
    End If
    If rsAdConceptoAnticipo.BOF Then
        rsAdConceptoAnticipo.MoveNext
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ValidaIntegridad
    
    rsAdConceptoAnticipo.Delete
    rsAdConceptoAnticipo.Update
    
    Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "CONCEPTO DE ATENCION", txtNumero.Text)
    txtNumero.SetFocus
    
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
    
    rsAdConceptoAnticipo.MoveLast
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    sstabConceptos.Tab = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError
    
    If Not rsAdConceptoAnticipo.EOF Then
        rsAdConceptoAnticipo.MoveNext
    End If
    If rsAdConceptoAnticipo.EOF Then
        rsAdConceptoAnticipo.MovePrevious
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
Dim vlblnAlta As Boolean
    
    If fblnDatosValidos() Then
        With rsAdConceptoAnticipo
            If Not vlblnConsulta Then
                .AddNew
                vlblnAlta = True
            End If
            !chrDescripcion = Trim(txtDescripcion.Text)
            !chrTratamiento = cboTratamiento.List(cboTratamiento.ListIndex)
            !mnyAnticipoSugerido = Val(Format(txtAnticipo.Text, "#################.00"))
            !bitactivo = IIf(chkActivo.Value = 1, 1, 0)
            .Update
            If vlblnAlta Then
                txtNumero.Text = flngObtieneIdentity("SEC_ADCONCEPTOANTICIPO", CStr(rsAdConceptoAnticipo!intConsecutivo))
                'txtNumero.Text = !intConsecutivo
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "CONCEPTO DE ATENCION", txtNumero.Text)
                vlblnAlta = False
            Else
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "CONCEPTO DE ATENCION", txtNumero.Text)
            End If
        End With
        rsAdConceptoAnticipo.Requery
        txtNumero.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
    
    rsAdConceptoAnticipo.MoveFirst
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    
    vgstrNombreForm = Me.Name
    
    If KeyAscii = 27 Then
        If sstabConceptos.Tab = 1 Then
            vlblnConsulta = False
            sstabConceptos.Tab = 0
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
    SetStyle sstabConceptos.hwnd, 0
    SetSolidColor sstabConceptos.hwnd, 16777215
    SSTabSubclass sstabConceptos.hwnd
    
    
    Me.Icon = frmMenuPrincipal.Icon
    '---------------------------------------------------------
    ' Recordsets tipo tabla
    '---------------------------------------------------------
    vlstrx = "select * from AdConceptoAnticipo order by intConsecutivo"
    Set rsAdConceptoAnticipo = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
    
    sstabConceptos.Tab = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub grdConceptos_DblClick()
    On Error GoTo NotificaError
    
    If fintLocalizaPkRs(rsAdConceptoAnticipo, 0, Str(grdConceptos.RowData(grdConceptos.Row))) <> 0 Then
        pMuestra
        sstabConceptos.Tab = 0
        pHabilita 1, 1, 1, 1, 1, 0, 1
        cmdLocate.SetFocus
    Else
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdConceptos _DblClick"))
End Sub

Private Sub grdConceptos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If fintLocalizaPkRs(rsAdConceptoAnticipo, 0, Str(grdConceptos.RowData(grdConceptos.Row))) <> 0 Then
            pMuestra
            sstabConceptos.Tab = 0
            pHabilita 1, 1, 1, 1, 1, 0, 1
            cmdLocate.SetFocus
        Else
            Unload Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdConceptos _KeyPress"))
End Sub

Private Sub sstabConceptos_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If sstabConceptos.Tab = 0 Then
        If Not vlblnConsulta Then
            txtNumero.SetFocus
        End If
    End If
    If sstabConceptos.Tab = 1 Then
        grdConceptos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstabConceptos_Click"))
End Sub

Private Sub txtAnticipo_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    txtAnticipo.Text = Val(Format(txtAnticipo.Text, "#############.00"))
    pSelTextBox txtAnticipo

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtAnticipo_GotFocus"))
End Sub

Private Sub txtAnticipo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    
    If KeyCode = vbKeyReturn Then
        If vlblnConsulta Then
            chkActivo.SetFocus
        Else
            cmdSave.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtAnticipo_KeyDown"))
End Sub

Private Sub txtAnticipo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not fblnFormatoCantidad(txtAnticipo, KeyAscii, 2) Then
       KeyAscii = 7
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtAnticipo_KeyPress"))
End Sub

Private Sub txtAnticipo_LostFocus()
    On Error GoTo NotificaError
    
    txtAnticipo.Text = FormatCurrency(Str(Val(Format(txtAnticipo.Text, "###############.00"))))


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtAnticipo_LostFocus"))
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
        cboTratamiento.SetFocus
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
    cboTratamiento.ListIndex = 0
    txtAnticipo.Text = ""
    chkActivo.Value = 1
    
    grdConceptos.Clear
    grdConceptos.Rows = 2
    grdConceptos.Cols = 6
    pGrid
    
    If rsAdConceptoAnticipo.RecordCount = 0 Then
        pHabilita 0, 0, 0, 0, 0, 0, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, 0
        pLlenarMshFGrdRs grdConceptos, rsAdConceptoAnticipo, 0
        pGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Function flngSiguiente() As Long
    On Error GoTo NotificaError
    
    Dim rsSiguienteNumero As New ADODB.Recordset
    
    vlstrx = "select isnull(max(intConsecutivo),0)+1 from AdConceptoAnticipo"
    Set rsSiguienteNumero = frsRegresaRs(vlstrx)
    
    flngSiguiente = rsSiguienteNumero.Fields(0)


Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngSiguiente"))
End Function

Private Sub pGrid()
    On Error GoTo NotificaError
    
    With grdConceptos
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Número|Descripción|Tratamiento|Anticipo|Estado"
        .ColWidth(0) = 100
        .ColWidth(1) = 1000     'Clave empresa
        .ColWidth(2) = 4000     'Nombre
        .ColWidth(3) = 1500     'Tratamiento
        .ColWidth(4) = 0        'Anticipo
        .ColWidth(5) = 0        'Estado
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
        
        If fintLocalizaPkRs(rsAdConceptoAnticipo, 0, txtNumero.Text) = 0 Then
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
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Then
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
    
    txtNumero.Text = rsAdConceptoAnticipo!intConsecutivo
    txtDescripcion.Text = Trim(rsAdConceptoAnticipo!chrDescripcion)
    cboTratamiento.ListIndex = fintLocalizaCritCbo_new(cboTratamiento, Trim(rsAdConceptoAnticipo!chrTratamiento))
    txtAnticipo.Text = FormatCurrency(rsAdConceptoAnticipo!mnyAnticipoSugerido, 2)
    chkActivo.Value = IIf(rsAdConceptoAnticipo!bitactivo Or rsAdConceptoAnticipo!bitactivo = 1, 1, 0)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra"))
End Sub




Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    Dim rsDatoDuplicado As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    fblnDatosValidos = True
    
    If Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    End If
        
    If fblnDatosValidos Then
        vlstrSentencia = "select * from AdConceptoAnticipo where ltrim(rtrim(chrDescripcion)) = '" & Trim(txtDescripcion.Text) & "' and intConsecutivo <>" & Trim(txtNumero.Text)
        Set rsDatoDuplicado = frsRegresaRs(vlstrSentencia)
        If rsDatoDuplicado.RecordCount <> 0 Then
            fblnDatosValidos = False
            'Este dato ya está registrado.
            MsgBox SIHOMsg(404), vbOKOnly + vbInformation, "Mensaje"
            txtDescripcion.SetFocus
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function


