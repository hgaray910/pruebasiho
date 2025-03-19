VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmMantoPais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de países"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstab 
      Height          =   4035
      Left            =   -10
      TabIndex        =   14
      Top             =   -10
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   7117
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoPais.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoPais.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
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
         Height          =   2260
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   8120
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
            Left            =   1170
            MaxLength       =   20
            TabIndex        =   1
            ToolTipText     =   "Nombre"
            Top             =   700
            Width           =   6750
         End
         Begin VB.TextBox txtArea 
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
            MaxLength       =   3
            TabIndex        =   2
            ToolTipText     =   "Area"
            Top             =   1110
            Width           =   795
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
            Left            =   1170
            TabIndex        =   5
            ToolTipText     =   "Activo / inactivo"
            Top             =   1920
            Width           =   1530
         End
         Begin VB.TextBox txtIDPais 
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
            Left            =   7320
            MaxLength       =   2
            TabIndex        =   4
            ToolTipText     =   "ID del país, para efecto del DIOT, ver lista del SAT para identificar el ID del país"
            Top             =   1110
            Width           =   600
         End
         Begin HSFlatControls.MyCombo cboCatalogoSat 
            Height          =   375
            Left            =   2280
            TabIndex        =   23
            ToolTipText     =   "País correspondiente al catálogo del SAT"
            Top             =   1515
            Width           =   5640
            _ExtentX        =   9948
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
         Begin HSFlatControls.MyCombo cboNacionalidad 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            ToolTipText     =   "Nacionalidad"
            Top             =   1110
            Width           =   2655
            _ExtentX        =   4683
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
            TabIndex        =   16
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
            TabIndex        =   17
            Top             =   760
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Área"
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
            TabIndex        =   18
            Top             =   1170
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nacionalidad"
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
            Left            =   2040
            TabIndex        =   21
            Top             =   1170
            Width           =   1335
         End
         Begin VB.Label lbIDPais 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Id del país"
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
            Left            =   6240
            TabIndex        =   22
            ToolTipText     =   "ID del país, usado para efecto del DIOT"
            Top             =   1170
            Width           =   990
         End
         Begin VB.Label lblCatalogoSat 
            BackColor       =   &H80000005&
            Caption         =   "País en catálogo SAT"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   24
            Top             =   1575
            Width           =   2055
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
         Height          =   3540
         Left            =   -75000
         TabIndex        =   20
         Top             =   0
         Width           =   8300
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPaises 
            Height          =   2985
            Left            =   90
            TabIndex        =   13
            ToolTipText     =   "Países"
            Top             =   120
            Width           =   8190
            _ExtentX        =   14446
            _ExtentY        =   5265
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
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
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
         Left            =   2040
         TabIndex        =   19
         Top             =   2230
         Width           =   4320
         Begin MyCommandButton.MyButton cmdTop 
            Height          =   600
            Left            =   60
            TabIndex        =   6
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
            Picture         =   "frmMantoPais.frx":0038
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoPais.frx":09BA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBack 
            Height          =   600
            Left            =   660
            TabIndex        =   7
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
            Picture         =   "frmMantoPais.frx":133C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoPais.frx":1CBE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdLocate 
            Height          =   600
            Left            =   1260
            TabIndex        =   8
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
            Picture         =   "frmMantoPais.frx":2640
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoPais.frx":2FC4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdNext 
            Height          =   600
            Left            =   1860
            TabIndex        =   9
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
            Picture         =   "frmMantoPais.frx":3948
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoPais.frx":42CA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdEnd 
            Height          =   600
            Left            =   2460
            TabIndex        =   10
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
            Picture         =   "frmMantoPais.frx":4C4C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoPais.frx":55CE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSave 
            Height          =   600
            Left            =   3060
            TabIndex        =   11
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
            Picture         =   "frmMantoPais.frx":5F50
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoPais.frx":68D4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   3660
            TabIndex        =   12
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
            Picture         =   "frmMantoPais.frx":7258
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoPais.frx":7BDA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMantoPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Programa para dar mantenimiento al catálogo de países
' Fecha de inicio de desarrollo: Septiembre, 2003
'------------------------------------------------------------------------------
' Ultimas modificaciones al módulo, especificar:
' Fecha:
' Descripcion del cambio:
'------------------------------------------------------------------------------

Option Explicit

Public vllngNumeroOpcionPais As Long        'Número de opción del módulo que usa el catálogo

Dim rsPais As New ADODB.Recordset

Dim vlblnConsulta As Boolean

Dim vlstrSentencia As String

Private Sub cboCatalogoSat_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboCatalogoSat_GotFocus"))
End Sub

Private Sub cboNacionalidad_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboNacionalidad_GotFocus"))
End Sub

Private Sub cboNacionalidad_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    Dim vllngUltimaClave As Long
    
    If KeyAscii = 13 Then
        If cboNacionalidad.ListIndex <> -1 Then
            If cboNacionalidad.ItemData(cboNacionalidad.ListIndex) = 0 Then
                frmMantoTresCampos.vllngCveCatalogo = 33
                frmMantoTresCampos.vlblnVisualizarCboCatalogos = False
                frmMantoTresCampos.Show vbModal, Me
            
                pLlenaNacionalidad
            
                vllngUltimaClave = frsRegresaRs("select isnull(max(intCveNacionalidad),0) from Nacionalidad").Fields(0)
                cboNacionalidad.ListIndex = flngLocalizaCbo_new(cboNacionalidad, STR(vllngUltimaClave))
            Else
                
                    SendKeys vbTab
                
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboNacionalidad_KeyPress"))
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
    
    If Not rsPais.BOF Then
        rsPais.MovePrevious
    End If
    If rsPais.BOF Then
        rsPais.MoveNext
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ValidaIntegridad
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionPais, "E") Then
        rsPais.Delete
        rsPais.Update
        
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "PAIS", txtNumero.Text)
        txtNumero.SetFocus
    End If
    
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
    
    rsPais.MoveLast
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
    
    If Not rsPais.EOF Then
        rsPais.MoveNext
    End If
    If rsPais.EOF Then
        rsPais.MovePrevious
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionPais, "E") Then
        If fblnDatosValidos() Then
            With rsPais
                If Not vlblnConsulta Then
                    .AddNew
                End If
                If Not vlblnConsulta Then
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "PAIS", txtNumero.Text)
                Else
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "PAIS", txtNumero.Text)
                End If
                !vchDescripcion = Trim(txtDescripcion.Text)
                !vchArea = IIf(Trim(txtArea.Text) = "", "  ", Trim(txtArea.Text))
                !INTCVENACIONALIDAD = cboNacionalidad.ItemData(cboNacionalidad.ListIndex)
                !bitactivo = IIf(chkActivo.Value = 1, 1, 0)
                !CHRIDPAIS = txtIDPais.Text
                If cboCatalogoSat.ListIndex > -1 Then
                    !INTCVECATALOGOSAT = cboCatalogoSat.ItemData(cboCatalogoSat.ListIndex)
                End If
                .Update
                If Not vlblnConsulta Then
                    txtNumero.Text = flngObtieneIdentity("SEC_PAIS", CStr(rsPais!INTCVEPAIS))
                    'txtNumero.Text = !intCvePais
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "PAIS", txtNumero.Text)
                Else
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "PAIS", txtNumero.Text)
                End If
            End With
            rsPais.Requery
            txtNumero.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
    
    rsPais.MoveFirst
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
    
    
    
    Me.Icon = frmMenuPrincipal.Icon
    '---------------------------------------------------------
    ' Recordsets tipo tabla
    '---------------------------------------------------------
    
    vlstrSentencia = "select * from Pais order by intCvePais"
    Set rsPais = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    pLlenaNacionalidad
    pLlenaCatalogoSat
    
    sstab.Tab = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pLlenaNacionalidad()
    On Error GoTo NotificaError
    
    Dim rsNacionalidad As New ADODB.Recordset

    vlstrSentencia = "select intCveNacionalidad, vchDescripcion from Nacionalidad where bitActiva = 1"
    Set rsNacionalidad = frsRegresaRs(vlstrSentencia)
    
    pLlenarCboRs_new cboNacionalidad, rsNacionalidad, 0, 1, 1
    cboNacionalidad.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaNacionalidad"))
End Sub


Private Sub grdPaises_DblClick()
    On Error GoTo NotificaError
 
    Dim vllngColumnaActual As Long
    
    With grdPaises
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
            pOrdColMshFGrid grdPaises, vgintTipoOrd
            pDesSelMshFGrid grdPaises
            
            .Col = vllngColumnaActual
            .Row = 1
        Else
            If fintLocalizaPkRs(rsPais, 0, STR(grdPaises.RowData(grdPaises.Row))) <> 0 Then
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPaises_DblClick"))
End Sub

Private Sub grdPaises_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If fintLocalizaPkRs(rsPais, 0, STR(grdPaises.RowData(grdPaises.Row))) <> 0 Then
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPaises_KeyPress"))
End Sub

Private Sub sstab_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If sstab.Tab = 0 Then
        If Not vlblnConsulta Then
            txtNumero.SetFocus
        End If
    End If
    If sstab.Tab = 1 Then
        grdPaises.SetFocus
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
        txtArea.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))
End Sub



Private Sub txtIDPais_Click()
    
    pSelTextBox txtIDPais
    
End Sub

Private Sub txtIDPais_GotFocus()
    
    pSelTextBox txtIDPais
    
End Sub

Private Sub txtIDPais_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
    If IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Or KeyAscii = 32 Then KeyAscii = 7
    
        If KeyAscii = 13 Then
            
            SendKeys vbTab
        
        End If
    
        
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
    txtArea.Text = ""
    cboNacionalidad.ListIndex = 0
    cboCatalogoSat.ListIndex = -1
    chkActivo.Value = 1
    txtIDPais.Text = ""
    
    grdPaises.Clear
    grdPaises.Rows = 2
    grdPaises.Cols = 6
    pGrid
    
    If rsPais.RecordCount = 0 Then
        pHabilita 0, 0, 0, 0, 0, 0, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, 0
        pLlenarMshFGrdRs grdPaises, rsPais, 0
        pGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Function flngSiguiente() As Long
    On Error GoTo NotificaError
    
    Dim rsSiguienteNumero As New ADODB.Recordset
    
    vlstrSentencia = "select isnull(max(intCvePais),0)+1 from Pais"
    Set rsSiguienteNumero = frsRegresaRs(vlstrSentencia)
    
    flngSiguiente = rsSiguienteNumero.Fields(0)


Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngSiguiente"))
End Function

Private Sub pGrid()
    On Error GoTo NotificaError
    
    With grdPaises
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Número|Nombre||||ID del país"
        .ColWidth(0) = 100
        .ColWidth(1) = 1000     'Numero
        .ColWidth(2) = 5500     'Nombre
        .ColWidth(3) = 0        'Area
        .ColWidth(4) = 0        'Nacionalidad
        .ColWidth(5) = 0        'Activo
        .ColWidth(6) = 1200
        
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
        
        If fintLocalizaPkRs(rsPais, 0, txtNumero.Text) = 0 Then
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
    
    txtNumero.Text = rsPais!INTCVEPAIS
    txtDescripcion.Text = Trim(rsPais!vchDescripcion)
    txtArea.Text = Trim(rsPais!vchArea)
    cboNacionalidad.ListIndex = flngLocalizaCbo_new(cboNacionalidad, STR(rsPais!INTCVENACIONALIDAD))
    txtIDPais.Text = IIf(IsNull(rsPais!CHRIDPAIS), "", rsPais!CHRIDPAIS)
    chkActivo.Value = IIf(rsPais!bitactivo Or rsPais!bitactivo = 1, 1, 0)
    ' Cargar codigo del sat
    'cboCatalogoSat.ListIndex = flngLocalizaCbo_new(cboCatalogoSat, Str(rsPais!INCVECATALOGOSAT))
    cboCatalogoSat.ListIndex = flngLocalizaCbo_new(cboCatalogoSat, IIf(IsNull(rsPais!INTCVECATALOGOSAT), "", rsPais!INTCVECATALOGOSAT))
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra"))
End Sub


Private Sub txtArea_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtArea

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtArea_GotFocus"))
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Or KeyAscii = 32 Then KeyAscii = 7
    
    If KeyAscii = 13 Then
        cboNacionalidad.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtArea_KeyPress"))
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
    If fblnDatosValidos And cboNacionalidad.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboNacionalidad.SetFocus
    End If
    If fblnDatosValidos Then
        
        vlstrSentencia = "select count(*) from Pais where vchDescripcion = '" & Trim(txtDescripcion.Text) & "'  and intCvePais <> " & IIf(vlblnConsulta, txtNumero.Text, "0")
        
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

Private Sub pLlenaCatalogoSat()
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset

    Set rs = frsCatalogoSAT("c_Pais")
        
    pLlenarCboRs_new cboCatalogoSat, rs, 0, 1
    'cboCatalogoSat.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCatalogoSat"))
End Sub

