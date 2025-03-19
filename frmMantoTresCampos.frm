VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmMantoTresCampos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nombre del catálogo"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstab 
      Height          =   3825
      Left            =   -10
      TabIndex        =   10
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
      TabPicture(0)   =   "frmMantoTresCampos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkVista"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoTresCampos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkVista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Catálogo central"
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
         Left            =   5520
         TabIndex        =   18
         Top             =   1170
         Visible         =   0   'False
         Width           =   2295
      End
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
         Height          =   1550
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   7935
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
            Left            =   1530
            MaxLength       =   5
            TabIndex        =   0
            ToolTipText     =   "Número "
            Top             =   300
            Width           =   810
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
            Left            =   1530
            MaxLength       =   20
            MultiLine       =   -1  'True
            TabIndex        =   1
            ToolTipText     =   "Descripción"
            Top             =   700
            Width           =   6195
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
            Left            =   1560
            TabIndex        =   2
            Top             =   1170
            Width           =   1605
         End
         Begin HSFlatControls.MyCombo cboCatalogo 
            Height          =   375
            Left            =   3720
            TabIndex        =   17
            ToolTipText     =   "Selección del catálogo para dar mantenimiento"
            Top             =   300
            Width           =   4005
            _ExtentX        =   7064
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
            TabIndex        =   12
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
            TabIndex        =   13
            Top             =   760
            Width           =   1125
         End
         Begin VB.Label lblCatalogos 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Catálogos"
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
            Left            =   2520
            TabIndex        =   16
            Top             =   360
            Width           =   1035
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
         Height          =   3015
         Left            =   -75000
         TabIndex        =   15
         Top             =   0
         Width           =   8120
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConsulta 
            Height          =   2300
            Left            =   75
            TabIndex        =   9
            Top             =   105
            Width           =   8010
            _ExtentX        =   14129
            _ExtentY        =   4048
            _Version        =   393216
            ForeColor       =   0
            Rows            =   0
            FixedRows       =   0
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            AllowBigSelection=   0   'False
            HighLight       =   0
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            MergeCells      =   1
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
         Left            =   2040
         TabIndex        =   14
         Top             =   1500
         Width           =   4320
         Begin MyCommandButton.MyButton cmdTop 
            Height          =   600
            Left            =   60
            TabIndex        =   3
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
            Picture         =   "frmMantoTresCampos.frx":0038
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTresCampos.frx":09BA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBack 
            Height          =   600
            Left            =   660
            TabIndex        =   4
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
            Picture         =   "frmMantoTresCampos.frx":133C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTresCampos.frx":1CBE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdLocate 
            Height          =   600
            Left            =   1260
            TabIndex        =   19
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
            Picture         =   "frmMantoTresCampos.frx":2640
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTresCampos.frx":2FC4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdNext 
            Height          =   600
            Left            =   1860
            TabIndex        =   5
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
            Picture         =   "frmMantoTresCampos.frx":3948
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTresCampos.frx":42CA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdEnd 
            Height          =   600
            Left            =   2460
            TabIndex        =   6
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
            Picture         =   "frmMantoTresCampos.frx":4C4C
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTresCampos.frx":55CE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSave 
            Height          =   600
            Left            =   3060
            TabIndex        =   7
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
            Picture         =   "frmMantoTresCampos.frx":5F50
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTresCampos.frx":68D4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   3660
            TabIndex        =   8
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
            Picture         =   "frmMantoTresCampos.frx":7258
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTresCampos.frx":7BDA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMantoTresCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------
' Programa para dar mantenimiento a tablas de tres campos: clave, descripción y estado.
' La lista de los catálgos a los que se les dará mantenimiento se carga de lo que está
' asignado en SiCatalogoModulo según el contenido de <cgstrModulo>
' Fecha de inicio de desarrollo: 06/Agosto/2003
'--------------------------------------------------------------------------------------

Option Explicit

Public vllngCveCatalogo As Long                 'Consecutivo de la tabla SiCatalogo a la
                                                'que se le dará mantenimiento
                                                'Cuando este parámetro es cero la lista
                                                'se posiciona en el primer catálogo

Public vlblnVisualizarCboCatalogos As Boolean   'Indica si se visualizará la lista de catálogos
Public vlstrModulo As String                    'Indica si se cargará en el combo de cátalogos Todos(Variable vacia), Expediente (EX), Nómina (NO)


Dim rsTabla As New ADODB.Recordset              'Tabla a la que se le da mantenimiento
Dim rs As New ADODB.Recordset                   'Acceso a datos
    
Dim vlblnConsulta As Boolean                    'Bandera para detectar cuando es consulta o alta
Dim vlblnFaltanCatalogos As Boolean             'Cuando no se encontraron catálos en SiCatalogoModulo
Dim vlblnFaltaCatalogoAsignado As Boolean       'Cuando el catálogo al que se le da mantenimiento no se encuentra asignado al módulo

Dim vlstrSentencia As String                    'Para sentencias SQL
Dim vlstrNombreTabla As String                  'Nombre de la tabla
Dim vlstrNombreCampoClave As String             'Nombre del campo que tiene la clave
Dim vlstrNombreCampoDescripcion As String       'Nombre del campo para la descripción
Dim vlstrNombreCampoEstado As String            'Nombre del campo para el estado del registro
Dim vllngNumeroOpcion As Long                   'Número de opción para el módulo
Dim blnCatalogoCentralizado As Boolean
Dim blnCambiandoCatalogo As Boolean
Dim blnBotonCatalogo As Boolean 'Bandera para validar si hay documentos en el catálogo de documentos solicitados a los médicos
Dim blnValidaGuardarCambios As Boolean 'Bandera para validar si se habilita el botón para guardar cambios


Private Function fValidaDescripcionDuplicada() As Boolean
Dim rsDuplicada As New ADODB.Recordset
Dim vlstrsql As String
    
    vlstrsql = "SELECT * FROM " & vlstrNombreTabla & " WHERE " & vlstrNombreCampoDescripcion & " = '" & Trim(txtDescripcion.Text) & "'"
    
    Set rsDuplicada = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    If rsDuplicada.RecordCount <> 0 Then
        fValidaDescripcionDuplicada = True
        MsgBox Me.Caption & " duplicada!", vbExclamation, "Mensaje"
        txtDescripcion.SetFocus
    Else
        fValidaDescripcionDuplicada = False
    End If
    
    rsDuplicada.Close
    
End Function

Private Sub cboCatalogo_Click()
On Error GoTo NotificaError
    
    If cboCatalogo.ListIndex <> -1 Then
        
        vlstrSentencia = "select SiCatalogo.*,SiCatalogoModulo.* from SiCatalogo inner join SiCatalogoModulo on SiCatalogo.intCveCatalogo = SiCatalogoModulo.intCveCatalogo where SiCatalogoModulo.chrModulo = '" & Trim(cgstrModulo) & "' and SiCatalogo.intCveCatalogo = " & Str(cboCatalogo.ItemData(cboCatalogo.ListIndex))
        Set rs = frsRegresaRs(vlstrSentencia)
        
        If rs.RecordCount <> 0 Then
        
            vlstrNombreTabla = Trim(rs!chrTabla)
            vlstrNombreCampoClave = Trim(rs!chrCampoClave)
            vlstrNombreCampoDescripcion = Trim(rs!chrCampoDescripcion)
            vlstrNombreCampoEstado = Trim(rs!chrCampoEstado)
            vllngNumeroOpcion = Trim(rs!intNumeroOpcion)
            
            Me.Caption = Trim(rs!chrDescripcion)
            blnValidaGuardarCambios = False
            If Trim(rs!chrDescripcion) = "Documentos de expedientes de médicos" Then
                blnBotonCatalogo = True
            Else
                blnBotonCatalogo = False
            End If
                        
            ' Recordsets tipo tabla
            vlstrSentencia = "select * from " & vlstrNombreTabla & " order by " & vlstrNombreCampoClave
            Set rsTabla = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            
            txtDescripcion.MaxLength = rsTabla.Fields(1).DefinedSize
            
            If rsTabla.Fields(1).DefinedSize > 50 Then
                frmMantoTresCampos.Height = 3450
                txtDescripcion.Height = 945
                chkActivo.Top = 1680
                Frame1.Height = 2100
                Frame2.Top = 2040
                Frame3.Height = 3015
                grdConsulta.Height = 2790
                Me.chkVista.Top = 2190
            Else
                frmMantoTresCampos.Height = 2910
                txtDescripcion.Height = 315
                chkActivo.Top = 1210
                Frame1.Top = 0
                Frame1.Height = 1550
                Frame2.Top = 1500
                Frame3.Height = 2535
                grdConsulta.Height = 2300
                Me.chkVista.Top = 1560
            End If
            
            pLimpia
            
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboCatalogo_Click"))
End Sub

Private Sub cboCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        txtNumero.SetFocus
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboCatalogo_KeyDown"))
End Sub

Private Sub chkActivo_Click()
        txtDescripcion_GotFocus
End Sub

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If cmdSave.Enabled And cmdSave.Visible Then
            cmdSave.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_KeyPress"))
End Sub



Private Sub chkVista_Click()
    
    If Not blnCambiandoCatalogo Then
        If Me.chkVista.Value = vbChecked Then
            Set rsTabla = frsEjecuta_SP("VM" & UCase(vlstrNombreTabla), "sp_siRegsNuevosVista", False)
        Else
            Set rsTabla = frsRegresaRs("select * from " & vlstrNombreTabla & " order by " & vlstrNombreCampoClave, adLockOptimistic, adOpenDynamic)
        End If
        pLimpia
    End If
    
End Sub

Private Sub cmdBack_Click()
On Error GoTo NotificaError
    
    If Not rsTabla.BOF Then
        rsTabla.MovePrevious
    End If
    If rsTabla.BOF Then
        rsTabla.MoveNext
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub


Private Sub cmdDelete_Click()
On Error GoTo ValidaIntegridad
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C") Then
        'If MsgBox(SIHOMsg(6), vbYesNo + vbCritical, "Mensaje") = vbYes Then
           If blnBotonCatalogo Then
                vlstrSentencia = "SELECT COUNT(*) FROM HODOCUMENTOMEDICO A "
                vlstrSentencia = vlstrSentencia & "INNER JOIN HODOCUMENTOSOLICITAMEDICO B ON A.INTCLAVE = B.INTCVEDOCUMENTO "
                vlstrSentencia = vlstrSentencia & "Where A.INTCLAVE = " & txtNumero.Text
                Set rs = frsRegresaRs(vlstrSentencia)
                If rs(0) > 0 Then
                    MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
                    Exit Sub
                End If
            End If
            rsTabla.Delete
            rsTabla.Update
            Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, Me.Caption, txtNumero.Text)
            txtNumero.SetFocus
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If
    
Exit Sub
ValidaIntegridad:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdEnd_Click()
On Error GoTo NotificaError
    
    rsTabla.MoveLast
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
    
    If Not rsTabla.EOF Then
        rsTabla.MoveNext
    End If
    If rsTabla.EOF Then
        rsTabla.MovePrevious
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
On Error GoTo NotificaError
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then
        
        If fblnDatosValidos() Then
            
            If Me.chkVista.Value = vbChecked And Me.chkVista.Visible Then
                On Error GoTo UpdateCentErr
                EntornoSIHO.ConeccionSIHO.Execute ("insert into " & vlstrNombreTabla & " (" & vlstrNombreCampoClave & ", " & vlstrNombreCampoDescripcion & ", " & vlstrNombreCampoEstado & ") values(" & Me.txtNumero.Text & ", '" & Trim(txtDescripcion.Text) & "', " & IIf(chkActivo.Value = 1, 1, 0) & ")")
                On Error GoTo NotificaError
                chkVista_Click
            Else
                With rsTabla
                    If Not vlblnConsulta Then
                        .AddNew
                    End If
                    .Fields(1) = Trim(txtDescripcion.Text)
                    .Fields(2) = IIf(chkActivo.Value = 1, 1, 0)
                    .Update
                
                    If Not vlblnConsulta Then
                        txtNumero.Text = flngObtieneIdentity("SEC_" + IIf(Len(vlstrNombreTabla) > 26, Mid(vlstrNombreTabla, 1, 26), vlstrNombreTabla), 0)
                        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, Me.Caption, txtNumero.Text)
                    Else
                        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, Me.Caption, txtNumero.Text)
                    End If
                    .Requery
                End With
            End If
            
            txtNumero.SetFocus
        End If
    
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
    Exit Sub
UpdateCentErr:
    MsgBox SIHOMsg(649), , "Mensaje"
End Sub

Private Sub cmdTop_Click()
On Error GoTo NotificaError
    
    rsTabla.MoveFirst
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError
    
    vgstrNombreForm = Me.Name
    
    If vlblnFaltanCatalogos Then
        'El módulo no tiene catálogos asignados para usar esta opción de mantenimiento.
        MsgBox SIHOMsg(567), vbExclamation + vbOKOnly, "Mensaje"
        Unload Me
    End If
    If vlblnFaltaCatalogoAsignado Then
        'El catálogo no se encuentra asignado al módulo.
        MsgBox SIHOMsg(568), vbExclamation + vbOKOnly, "Mensaje"
        Unload Me
    End If
    
    cboCatalogo_Click
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
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
    
    Me.Icon = frmMenuPrincipal.Icon
    
    'Color de Tab
    SetStyle sstab.hwnd, 0
    SetSolidColor sstab.hwnd, 16777215
    SSTabSubclass sstab.hwnd
    

    Set rs = frsEjecuta_SP(IIf(vlstrModulo = "", Trim(cgstrModulo), vlstrModulo), "Sp_GnSelCatalogosModulo")
    If rs.RecordCount <> 0 Then
        
        pLlenarCboRs_new cboCatalogo, rs, 0, 1
        
        If vllngCveCatalogo <> 0 Then
            cboCatalogo.ListIndex = flngLocalizaCbo_new(cboCatalogo, Str(vllngCveCatalogo))
            If cboCatalogo.ListIndex = -1 Then
                vlblnFaltaCatalogoAsignado = True
            End If
        Else
            cboCatalogo.ListIndex = 0
        End If
        
        cboCatalogo.Visible = vlblnVisualizarCboCatalogos
        lblCatalogos.Visible = vlblnVisualizarCboCatalogos
        
    Else
        vlblnFaltanCatalogos = True
    End If
    
    sstab.Tab = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub grdConsulta_DblClick()
On Error GoTo NotificaError
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Integer
    
'*** Código para ordenar datos de columna seleccionada ***************************
    If grdConsulta.MouseRow = 0 Then
        vgintColOrdAnt = vgintColOrd 'Guarda la columna de ordenación anterior
        vgintColOrd = grdConsulta.Col  'Configura la columna a ordenar
        'Escoge el Tipo de Ordenamiento
        If vgintTipoOrd = 1 Then
             vgintTipoOrd = 2
        Else
                vgintTipoOrd = 1
        End If
        grdConsulta.FocusRect = flexFocusNone
        Call pOrdColMshFGrid(grdConsulta, vgintTipoOrd)
        Call pDesSelMshFGrid(grdConsulta)
        grdConsulta.FocusRect = flexFocusHeavy
        Exit Sub
    End If
'************************************************************************
    
    
    If fintLocalizaPkRs(rsTabla, 0, Str(grdConsulta.RowData(grdConsulta.Row))) <> 0 Then
        pMuestra
        sstab.Tab = 0
        If chkVista.Value = vbChecked Then
            pHabilita 0, 0, 0, 0, 0, 1, 0
            cmdSave.SetFocus
        Else
            pHabilita 1, 1, 1, 1, 1, 0, 1
            cmdLocate.SetFocus
        End If
    Else
        Unload Me
    End If
'.sort
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdConsulta_DblClick"))
End Sub

Private Sub grdConsulta_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If fintLocalizaPkRs(rsTabla, 0, Str(grdConsulta.RowData(grdConsulta.Row))) <> 0 Then
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdConsulta_KeyPress"))
End Sub

Private Sub sstab_Click(PreviousTab As Integer)
On Error GoTo NotificaError
    
    If sstab.Tab = 0 Then
        If Not vlblnConsulta Then
            txtNumero.SetFocus
        End If
    End If
    If sstab.Tab = 1 Then
        grdConsulta.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstab_Click"))
End Sub

Private Sub txtCveSecundaria_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_GotFocus"))
End Sub

Private Sub txtCveSecundaria_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    KeyAscii = UCase(KeyAscii)
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_GotFocus"))
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
        KeyAscii = 7
        chkActivo.SetFocus
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
    
    txtNumero.Text = flngSigNumRs(rsTabla, 0)
    txtDescripcion.Text = ""
    chkActivo.Value = 1
    
    grdConsulta.Clear
    grdConsulta.Rows = 2
    grdConsulta.Cols = 4
    pGrid
    
    rsTabla.Requery
    If rsTabla.RecordCount = 0 Then
        pHabilita 0, 0, 0, 0, 0, 0, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, 0
        pLlenarMshFGrdRs grdConsulta, rsTabla, 0
        pGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Function flngSiguiente() As Long
On Error GoTo NotificaError
Dim rsSiguienteNumero As New ADODB.Recordset
    
    vlstrSentencia = "select isnull(max(" & vlstrNombreCampoClave & "),0)+1 from " & vlstrNombreTabla
    Set rsSiguienteNumero = frsRegresaRs(vlstrSentencia)
    
    flngSiguiente = rsSiguienteNumero.Fields(0)

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngSiguiente"))
End Function

Private Sub pGrid()
On Error GoTo NotificaError
    
    With grdConsulta
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Número|Descripción"
        .ColWidth(0) = 100
        .ColWidth(1) = 1000     'Clave empresa
        .ColWidth(2) = 6000     'Nombre
        .ColWidth(3) = 0        'Activo/inactivo
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGrid"))
End Sub

Private Sub pHabilita(vlbln1 As Integer, vlbln2 As Integer, vlbln3 As Integer, vlbln4 As Integer, vlbln5 As Integer, vlbln6 As Integer, vlbln7 As Integer)
On Error GoTo NotificaError
    
    cmdTop.Enabled = vlbln1 = 1
    cmdBack.Enabled = vlbln2 = 1
    cmdLocate.Enabled = vlbln3 = 1
    cmdNext.Enabled = vlbln4 = 1
    cmdEnd.Enabled = vlbln5 = 1
    cmdSave.Enabled = vlbln6 = 1
    cmdDelete.Enabled = vlbln7 = 1
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(txtNumero.Text) = "" Then
            txtNumero.Text = flngSigNumRs(rsTabla, 0)
        End If
        
        If fintLocalizaPkRs(rsTabla, 0, txtNumero.Text) = 0 Then
            txtNumero.Text = flngSigNumRs(rsTabla, 0)
        Else
            pMuestra
        End If
        
        If vlblnConsulta Then
            pHabilita 1, 1, 1, 1, 1, 0, 1
            cmdTop.SetFocus
        ElseIf Not blnCatalogoCentralizado Then
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
    
    txtNumero.Text = rsTabla.Fields(0)
    txtDescripcion.Text = rsTabla.Fields(1)
    chkActivo.Value = IIf(rsTabla.Fields(2) Or rsTabla.Fields(2) = 1, 1, 0)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra"))
End Sub

Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError
Dim rsDatoDuplicado As New ADODB.Recordset
    
    fblnDatosValidos = True
    
    If Not vlblnConsulta Then
        If fValidaDescripcionDuplicada Then
            fblnDatosValidos = False
            Exit Function
        End If
    End If
    
    If Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    
    If fblnDatosValidos And Not vlblnConsulta Then
        vlstrSentencia = "select * from " & vlstrNombreTabla & " where ltrim(rtrim(" & vlstrNombreCampoDescripcion & ")) = '" & Trim(txtDescripcion.Text) & "' and " & vlstrNombreCampoClave & "<>" & Trim(txtNumero.Text)
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


