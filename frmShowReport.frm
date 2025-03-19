VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F593C05F-1783-4E85-9D91-D7A3A5D8D58C}#10.0#0"; "ReporterControls.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmShowReport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmEspere 
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
      Height          =   3015
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   4935
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Espere ..."
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
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3855
      Left            =   80
      ScaleHeight     =   3825
      ScaleWidth      =   6510
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   760
      Width           =   6540
      Begin MSComCtl2.FlatScrollBar VScroll1 
         Height          =   3820
         Left            =   6260
         TabIndex        =   15
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   6747
         _Version        =   393216
         Orientation     =   1179648
         SmallChange     =   415
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   6255
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   6255
         Begin VB.TextBox Text 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   2320
            TabIndex        =   9
            Top             =   -420
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.ComboBox Combo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   2320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   -420
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.CheckBox Check 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
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
            Height          =   390
            Index           =   0
            Left            =   2320
            TabIndex        =   6
            Top             =   -420
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3735
         End
         Begin ReporterControls.OptionList OptionList 
            Height          =   315
            Index           =   0
            Left            =   2320
            TabIndex        =   5
            Top             =   -420
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker DTPicker 
            Height          =   315
            Index           =   0
            Left            =   2320
            TabIndex        =   7
            Top             =   -420
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Format          =   116654081
            CurrentDate     =   38072
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
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
            Height          =   300
            Index           =   0
            Left            =   200
            TabIndex        =   10
            Top             =   -420
            Visible         =   0   'False
            Width           =   45
         End
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5880
      Top             =   4680
   End
   Begin HSFlatControls.MyCombo cboReporte 
      Height          =   375
      Left            =   80
      TabIndex        =   0
      Top             =   360
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   661
      Style           =   1
      Enabled         =   -1  'True
      Text            =   "cboReporte"
      Sorted          =   0   'False
      List            =   $"frmShowReport.frx":0000
      ItemData        =   $"frmShowReport.frx":000E
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
   Begin VB.Frame Frame6 
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
      Left            =   2680
      TabIndex        =   11
      Top             =   4530
      Width           =   1320
      Begin MyCommandButton.MyButton cmdPrint 
         Height          =   600
         Left            =   660
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
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
         Picture         =   "frmShowReport.frx":0015
         BackColorDown   =   -2147483643
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmShowReport.frx":0999
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPreview 
         Height          =   600
         Left            =   60
         TabIndex        =   13
         ToolTipText     =   "Vista previa"
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
         Picture         =   "frmShowReport.frx":131B
         BackColorDown   =   -2147483643
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmShowReport.frx":1C9F
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Reporte"
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
      Left            =   90
      TabIndex        =   14
      Top             =   100
      Width           =   6555
   End
End
Attribute VB_Name = "frmShowReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrReportes() As String
Dim strArchivo As String
Dim lngIdReporte As Long
Dim arrParametros() As String
Dim arrValues() As String
Dim colControls As Collection
Dim lngText As Long
Dim lngCombo As Long
Dim lngPicker As Long
Dim lngCheck As Long
Dim lngOption As Long
Dim lngLabel As Long
Dim lngTop As Long
Dim blnActivate As Boolean
Dim vgrptReporte As CRAXDRT.report
Dim strOriginalSelectionFormula As String

Private Sub pReporte(vlstrDestino As String)
    On Error GoTo NotificaError
    Dim intIndex As Integer
    Dim intOrden As Integer
    Dim strSelFormula As String
    Dim ctlComboBox As ComboBox
    Dim ctlCheckBox As CheckBox
    Dim ctlOptionList As OptionList
    Dim ctlDTPicker As DTPicker
    Dim ctlTextBox As TextBox
    'Crystal 10
    Dim strParametros As String
    Dim vlaryParametros() As String
    Dim rsInfoRep As ADODB.Recordset
    Dim rsReporte As New ADODB.Recordset
    Dim intTipoFuente As Byte
    Dim strComando As String
    Dim strQuery As String
    Me.MousePointer = 11
    strParametros = ""
    vgrptReporte.DiscardSavedData
    'CryReporte.Reset
    'CryReporte.WindowState = crptMaximized
    'CryReporte.WindowBorderStyle = crptFixedDouble
    'CryReporte.SelectionFormula = ""
    'CryReporte.Connect = fstrConexionCrystal()
    'CryReporte.ReportFileName = strArchivo
    'CryReporte.WindowTitle = Me.cboReporte.Text
    For intIndex = 0 To UBound(arrParametros, 2)
        If IsNumeric(arrParametros(5, intIndex)) Then
            intOrden = CInt(arrParametros(5, intIndex))
            Select Case arrParametros(1, intIndex)
                Case "Stored Procedure"
                    Select Case arrParametros(3, intIndex)
                        Case "Lista - ComboBox", "Consulta - ComboBox"
                            Set ctlComboBox = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                'CryReporte.StoredProcParam(intOrden) = ctlComboBox.ItemData(ctlComboBox.ListIndex)
                                strParametros = fAgregaParametro(strParametros, ctlComboBox.ItemData(ctlComboBox.ListIndex))
                            Else
                                'CryReporte.StoredProcParam(intOrden) = ctlComboBox.Text
                                strParametros = fAgregaParametro(strParametros, ctlComboBox.Text)
                            End If
                        Case "Lista - OptionButton"
                            Set ctlOptionList = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                'CryReporte.StoredProcParam(intOrden) = ctlOptionList.LngValue
                                strParametros = fAgregaParametro(strParametros, ctlOptionList.LngValue)
                            Else
                                'CryReporte.StoredProcParam(intOrden) = ctlOptionList.StrValue
                                strParametros = fAgregaParametro(strParametros, ctlOptionList.StrValue)
                            End If
                        Case "Dos valores - CheckBox"
                            Set ctlCheckBox = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                If ctlCheckBox.Value = vbChecked Then
                                    'CryReporte.StoredProcParam(intOrden) = CInt(fObtieneValores(intIndex, 1))
                                    strParametros = fAgregaParametro(strParametros, CInt(fObtieneValores(intIndex, 1)))
                                Else
                                    'CryReporte.StoredProcParam(intOrden) = CInt(fObtieneValores(intIndex, 2))
                                    strParametros = fAgregaParametro(strParametros, CInt(fObtieneValores(intIndex, 2)))
                                End If
                            Else
                                If ctlCheckBox.Value = vbChecked Then
                                    'CryReporte.StoredProcParam(intOrden) = fObtieneValores(intIndex, 1)
                                    strParametros = fAgregaParametro(strParametros, fObtieneValores(intIndex, 1))
                                Else
                                    'CryReporte.StoredProcParam(intOrden) = fObtieneValores(intIndex, 2)
                                    strParametros = fAgregaParametro(strParametros, fObtieneValores(intIndex, 2))
                                End If
                            End If
                        Case "Variable"
                            If fObtieneValores(intIndex, 1) = "(Establecer valor)" Then
                                'CryReporte.StoredProcParam(intOrden) = fObtieneValores(intIndex, 2)
                                strParametros = fAgregaParametro(strParametros, fObtieneValores(intIndex, 2))
                            Else
                                'CryReporte.StoredProcParam(intOrden) = GetValue(fObtieneValores(intIndex, 1))
                                strParametros = fAgregaParametro(strParametros, GetValue(fObtieneValores(intIndex, 1)))
                            End If
                        Case "Fecha ('yyyy-mm-dd') - DTPicker", "Fecha Hora ('yyyy-...') - DTPicker"
                            Set ctlDTPicker = colControls.Item("K" & intIndex)
                            'CryReporte.StoredProcParam(intOrden) = CStr(Format(ctlDTPicker.Value, fObtieneValores(intIndex, 1)))
                            strParametros = fAgregaParametro(strParametros, CStr(Format(ctlDTPicker.Value, fObtieneValores(intIndex, 1))))
                        Case "Fecha - DTPicker", "Fecha Hora - DTPicker"
                            Set ctlDTPicker = colControls.Item("K" & intIndex)
                            'CryReporte.StoredProcParam(intOrden) = ctlDTPicker.Value
                            strParametros = fAgregaParametro(strParametros, fstrFechaSQL(ctlDTPicker.Value))
                        Case "Libre - TextBox"
                            Set ctlTextBox = colControls.Item("K" & intIndex)
                            'CryReporte.StoredProcParam(intOrden) = ctlTextBox.Text
                            strParametros = fAgregaParametro(strParametros, ctlTextBox.Text)
                    End Select
                Case "Parameter Field"
                    ReDim Preserve vlaryParametros(intOrden)
                    Select Case arrParametros(3, intIndex)
                        Case "Fecha - DTPicker", "Fecha Hora - DTPicker"
                            Set ctlDTPicker = colControls.Item("K" & intIndex)
                            'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";Date(" & Year(ctlDTPicker) & "," & Month(ctlDTPicker) & "," & Day(ctlDTPicker) & ");TRUE"
                            vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & Format(ctlDTPicker.Value, "dd-mmm-yyyy") & ";Date"
                        Case "Lista - ComboBox", "Consulta - ComboBox"
                            Set ctlComboBox = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & ctlComboBox.ItemData(ctlComboBox.ListIndex) & ";TRUE"
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & ctlComboBox.ItemData(ctlComboBox.ListIndex) & ";Number"
                            Else
                                'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & ctlComboBox.Text & ";TRUE"
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & ctlComboBox.Text & ";String"
                            End If
                        Case "Lista - OptionButton"
                            Set ctlOptionList = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & ctlOptionList.LngValue & ";TRUE"
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & ctlOptionList.LngValue & ";Number"
                            Else
                                'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & ctlOptionList.StrValue & ";TRUE"
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & ctlOptionList.StrValue & ";String"
                            End If
                        Case "Dos valores - CheckBox"
                            Set ctlCheckBox = colControls.Item("K" & intIndex)
                            If ctlCheckBox.Value = vbChecked Then
                                'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & fObtieneValores(intIndex, 1) & ";TRUE"
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & fObtieneValores(intIndex, 1)
                                Else
                                'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & fObtieneValores(intIndex, 2) & ";TRUE"
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & fObtieneValores(intIndex, 2)
                            End If
                        Case "Variable"
                            If fObtieneValores(intIndex, 1) = "(Establecer valor)" Then
                                'CryReporte.ParameterFields(intOrden) = fObtieneValores(intIndex, 2)
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & fObtieneValores(intIndex, 2)
                            Else
                                'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & GetValue(fObtieneValores(intIndex, 1)) & ";TRUE"
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & GetValue(fObtieneValores(intIndex, 1))
                            End If
                        Case "Fecha ('yyyy-mm-dd') - DTPicker", "Fecha Hora ('yyyy-...') - DTPicker"
                            Set ctlDTPicker = colControls.Item("K" & intIndex)
                            'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & CStr(Format(ctlDTPicker.Value, fObtieneValores(intIndex, 1))) & ";TRUE"
                            vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & CStr(Format(ctlDTPicker.Value, fObtieneValores(intIndex, 1))) & ";String"
                        Case "Libre - TextBox"
                            Set ctlTextBox = colControls.Item("K" & intIndex)
                            'CryReporte.ParameterFields(intOrden) = arrParametros(4, intIndex) & ";" & ctlTextBox.Text & ";TRUE"
                            If arrParametros(2, intIndex) = "Entero" Then
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & ctlTextBox.Text & ";Number"
                            ElseIf arrParametros(2, intIndex) = "Numérico" Then
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & ctlTextBox.Text & ";Number"
                            Else
                                vlaryParametros(intOrden) = arrParametros(4, intIndex) & ";" & ctlTextBox.Text & ";String"
                            End If
                    End Select
                    pCargaParameterFields vlaryParametros, vgrptReporte
                Case "Selection Formula"
                    Select Case arrParametros(3, intIndex)
                        Case "Lista - ComboBox", "Consulta - ComboBox"
                            Set ctlComboBox = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                strSelFormula = arrParametros(4, intIndex) & " " & ctlComboBox.ItemData(ctlComboBox.ListIndex)
                            Else
                                strSelFormula = arrParametros(4, intIndex) & " '" & ctlComboBox.Text & "'"
                            End If
                        Case "Lista - OptionButton"
                            Set ctlOptionList = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                strSelFormula = arrParametros(4, intIndex) & " " & ctlOptionList.LngValue
                            Else
                                strSelFormula = arrParametros(4, intIndex) & " '" & ctlOptionList.StrValue & "'"
                            End If
                        Case "Dos valores - CheckBox"
                            Set ctlCheckBox = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "Entero" Then
                                If ctlCheckBox.Value = vbChecked Then
                                    strSelFormula = arrParametros(4, intIndex) & " " & fObtieneValores(intIndex, 1)
                                Else
                                    strSelFormula = arrParametros(4, intIndex) & " " & fObtieneValores(intIndex, 2)
                                End If
                            Else
                                If ctlCheckBox.Value = vbChecked Then
                                    strSelFormula = arrParametros(4, intIndex) & " '" & fObtieneValores(intIndex, 1) & "'"
                                Else
                                    strSelFormula = arrParametros(4, intIndex) & " '" & fObtieneValores(intIndex, 2) & "'"
                                End If
                            End If
                        Case "Variable"
                            If arrParametros(2, intIndex) = "String" Then
                                If fObtieneValores(intIndex, 1) = "(Establecer valor)" Then
                                    strSelFormula = arrParametros(4, intIndex) & " '" & fObtieneValores(intIndex, 2) & "'"
                                Else
                                    strSelFormula = arrParametros(4, intIndex) & " '" & GetValue(fObtieneValores(intIndex, 1)) & "'"
                                End If
                            Else
                                If fObtieneValores(intIndex, 1) = "(Establecer valor)" Then
                                    strSelFormula = arrParametros(4, intIndex) & " " & fObtieneValores(intIndex, 2)
                                Else
                                    strSelFormula = arrParametros(4, intIndex) & " " & GetValue(fObtieneValores(intIndex, 1))
                                End If
                            End If
                        Case "Libre - TextBox"
                            Set ctlTextBox = colControls.Item("K" & intIndex)
                            If arrParametros(2, intIndex) = "String" Then
                                strSelFormula = arrParametros(4, intIndex) & " '" & ctlTextBox.Text & "'"
                            Else
                                strSelFormula = arrParametros(4, intIndex) & " " & ctlTextBox.Text
                            End If
                        Case "Fecha - DTPicker", "Fecha Hora - DTPicker"
                            Set ctlDTPicker = colControls.Item("K" & intIndex)
                            strSelFormula = arrParametros(4, intIndex) & " Date(" & CStr(Format(ctlDTPicker.Value, "yyyy,mm,dd")) & ")"
                        Case "Fecha ('yyyy-mm-dd') - DTPicker", "Fecha Hora ('yyyy-...') - DTPicker"
                            strSelFormula = arrParametros(4, intIndex) & " " & CStr(Format(ctlDTPicker.Value, fObtieneValores(intIndex, 1)))
                    End Select
                    'If CryReporte.SelectionFormula = "" Then
                    If strOriginalSelectionFormula = "" Then
                        'CryReporte.SelectionFormula = strSelFormula
                        vgrptReporte.RecordSelectionFormula = strSelFormula
                    Else
                        'CryReporte.SelectionFormula = CryReporte.SelectionFormula & " and " & strSelFormula
                        vgrptReporte.RecordSelectionFormula = strOriginalSelectionFormula & " and " & strSelFormula
                    End If
            End Select
        End If
    Next
    Set rsInfoRep = frsRegresaRs("select * from SeReportes where intIdReporte = " & Me.cboReporte.ItemData(Me.cboReporte.ListIndex))
    If Not rsInfoRep.EOF Then
        intTipoFuente = rsInfoRep!intTipoFuente
        strComando = IIf(IsNull(rsInfoRep!vchComando), "", rsInfoRep!vchComando)
    End If
    rsInfoRep.Close
'    If vlstrDestino = "P" Then
'        CryReporte.Destination = crptToWindow
'    Else
'        CryReporte.Destination = crptToPrinter
'    End If
    'CryReporte.Action = 1
    On Error GoTo BadQuery
    If intTipoFuente = 1 Then
        If strComando <> "" Then
            Set rsReporte = frsEjecuta_SP(strParametros, strComando)
        End If
    Else
        If strComando <> "" Then
            strQuery = fGeneraConsulta(strParametros, strComando)
            Set rsReporte = frsRegresaRs(strQuery)
        End If
    End If
    On Error GoTo NotificaError
    pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, Me.cboReporte.Text
    Me.MousePointer = 0
    Exit Sub
NotificaError:
    Me.MousePointer = 0
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pReporte"))
    Exit Sub
BadQuery:
    MsgBox SIHOMsg(628), vbInformation, "Mensaje"
End Sub

Private Sub cboReporte_Click()
    On Error GoTo NotificaError
    ReDim arrValues(2, 0)
    ReDim arrParametros(5, 0)
    Me.Timer1.Enabled = True
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboReporte_Click"))
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo NotificaError
    If Me.cboReporte.ListIndex > 0 Then
        pReporte "P"
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPreview_Click"))
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo NotificaError
    If Me.cboReporte.ListIndex > 0 Then
        pReporte "I"
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrint_Click"))
End Sub

Private Sub DTPicker_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":DTPicker_KeyDown"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    FrmEspere.Visible = False
    If blnActivate Then
        Me.cboReporte.SetFocus
        blnActivate = False
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    lngText = 0
    lngCombo = 0
    lngPicker = 0
    lngCheck = 0
    lngLabel = 0
    lngOption = 0
    lngTop = 0
    Set colControls = New Collection
    pObtenerReportes
    blnActivate = True
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pObtenerReportes()
    On Error GoTo NotificaError
    Dim rsReportes As ADODB.Recordset
    Dim lngIndex As Long
    Set rsReportes = frsRegresaRs("select sereportes.* from sereportes inner join sereportesmodulos on sereportes.INTIDREPORTE = sereportesmodulos.INTIDREPORTE where sereportesmodulos.INTIDMODULO = " & vgintNumeroModulo & " and bitActivo <> 0 order by vchDescripcion", adLockReadOnly, adOpenForwardOnly)
    lngIndex = 1
    Do Until rsReportes.EOF
        Me.cboReporte.AddItem rsReportes!vchDescripcion
        Me.cboReporte.ItemData(Me.cboReporte.NewIndex) = rsReportes!intIdReporte
        ReDim Preserve arrReportes(1, 1 To lngIndex)
        arrReportes(0, lngIndex) = rsReportes!intIdReporte
        arrReportes(1, lngIndex) = rsReportes!vchNombreArchivo
        lngIndex = lngIndex + 1
        rsReportes.MoveNext
    Loop
    rsReportes.Close
    Me.cboReporte.ListIndex = 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pObtenerReportes"))
End Sub

Private Sub pCargarFiltros()
    On Error GoTo NotificaError
    Dim rsParametros As ADODB.Recordset
    Dim lngIndex As Long
    Set rsParametros = frsRegresaRs("select * from SeParametrosReportes where intIdReporte = " & lngIdReporte & " order by vchClase desc, intOrden", adLockReadOnly, adOpenForwardOnly)
    lngIndex = 0
    Do Until rsParametros.EOF
        ReDim Preserve arrParametros(5, lngIndex)
        arrParametros(0, lngIndex) = rsParametros!intIdParametro
        arrParametros(1, lngIndex) = rsParametros!vchClase
        arrParametros(2, lngIndex) = rsParametros!vchTipo
        arrParametros(3, lngIndex) = rsParametros!vchModo
        arrParametros(4, lngIndex) = IIf(IsNull(rsParametros!vchNombre), "", rsParametros!vchNombre)
        arrParametros(5, lngIndex) = rsParametros!intOrden
        pCargarControl rsParametros!vchModo, rsParametros!vchCaption, lngIndex
        lngIndex = lngIndex + 1
        rsParametros.MoveNext
    Loop
    rsParametros.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargarFiltros"))
End Sub

Private Sub pCargarControl(strModo As String, strTitulo As String, lngNP As Long)
    On Error GoTo NotificaError
    Select Case strModo
            Case "Libre - TextBox"
                Load Me.Text(lngText + 1)
                Me.Text(lngText + 1).Top = lngTop + 100
                Me.Text(lngText + 1).Visible = True
                colControls.Add Me.Text(lngText + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.Text(lngText + 1)
                Load Me.Label(lngLabel + 1)
                Me.Label(lngLabel + 1).Top = lngTop + 150
                Me.Label(lngLabel + 1).Visible = True
                Me.Label(lngLabel + 1).Caption = strTitulo
                lngTop = lngTop + Me.Text(lngText + 1).Height + 80
                lngLabel = lngLabel + 1
                lngText = lngText + 1
            Case "Fecha - DTPicker"
                Load Me.DTPicker(lngPicker + 1)
                Me.DTPicker(lngPicker + 1).Top = lngTop + 100
                Me.DTPicker(lngPicker + 1).Visible = True
                colControls.Add Me.DTPicker(lngPicker + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.DTPicker(lngPicker + 1)
                Load Me.Label(lngLabel + 1)
                Me.Label(lngLabel + 1).Top = lngTop + 150
                Me.Label(lngLabel + 1).Visible = True
                Me.Label(lngLabel + 1).Caption = strTitulo
                lngTop = lngTop + Me.DTPicker(lngPicker + 1).Height + 80
                lngLabel = lngLabel + 1
                lngPicker = lngPicker + 1
            Case "Fecha ('yyyy-mm-dd') - DTPicker"
                Load Me.DTPicker(lngPicker + 1)
                Me.DTPicker(lngPicker + 1).Top = lngTop + 100
                Me.DTPicker(lngPicker + 1).Visible = True
                colControls.Add Me.DTPicker(lngPicker + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.DTPicker(lngPicker + 1)
                Load Me.Label(lngLabel + 1)
                Me.Label(lngLabel + 1).Top = lngTop + 150
                Me.Label(lngLabel + 1).Visible = True
                Me.Label(lngLabel + 1).Caption = strTitulo
                lngTop = lngTop + Me.DTPicker(lngPicker + 1).Height + 80
                lngLabel = lngLabel + 1
                lngPicker = lngPicker + 1
            Case "Fecha Hora ('yyyy-...') - DTPicker", "Fecha Hora - DTPicker"
                Load Me.DTPicker(lngPicker + 1)
                Me.DTPicker(lngPicker + 1).Top = lngTop + 100
                Me.DTPicker(lngPicker + 1).Visible = True
                Me.DTPicker(lngPicker + 1).Format = dtpCustom
                Me.DTPicker(lngPicker + 1).CustomFormat = "dd/MM/yyyy hh:mm:ss tt"
                Me.DTPicker(lngPicker + 1).Width = Me.DTPicker(lngPicker + 1).Width + 200
                colControls.Add Me.DTPicker(lngPicker + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.DTPicker(lngPicker + 1)
                Load Me.Label(lngLabel + 1)
                Me.Label(lngLabel + 1).Top = lngTop + 150
                Me.Label(lngLabel + 1).Visible = True
                Me.Label(lngLabel + 1).Caption = strTitulo
                lngTop = lngTop + Me.DTPicker(lngPicker + 1).Height + 80
                lngLabel = lngLabel + 1
                lngPicker = lngPicker + 1
            Case "Consulta - ComboBox"
                Load Me.Combo(lngCombo + 1)
                Me.Combo(lngCombo + 1).Top = lngTop + 100
                Me.Combo(lngCombo + 1).Visible = True
                colControls.Add Me.Combo(lngCombo + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.Combo(lngCombo + 1)
                Load Me.Label(lngLabel + 1)
                Me.Label(lngLabel + 1).Top = lngTop + 150
                Me.Label(lngLabel + 1).Visible = True
                Me.Label(lngLabel + 1).Caption = strTitulo
                lngTop = lngTop + Me.Combo(lngCombo + 1).Height + 80
                lngLabel = lngLabel + 1
                lngCombo = lngCombo + 1
            Case "Dos valores - CheckBox"
                Load Me.Check(lngCheck + 1)
                Me.Check(lngCheck + 1).Top = lngTop + 100
                Me.Check(lngCheck + 1).Visible = True
                colControls.Add Me.Check(lngCheck + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.Label1
                Me.Check(lngCheck + 1).Caption = strTitulo
                lngTop = lngTop + Me.Check(lngCheck + 1).Height + 80
                lngCheck = lngCheck + 1
            Case "Lista - ComboBox"
                Load Me.Combo(lngCombo + 1)
                Me.Combo(lngCombo + 1).Top = lngTop + 100
                Me.Combo(lngCombo + 1).Visible = True
                colControls.Add Me.Combo(lngCombo + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.Combo(lngCombo + 1)
                Load Me.Label(lngLabel + 1)
                Me.Label(lngLabel + 1).Top = lngTop + 150
                Me.Label(lngLabel + 1).Visible = True
                Me.Label(lngLabel + 1).Caption = strTitulo
                lngTop = lngTop + Me.Combo(lngCombo + 1).Height + 80
                lngLabel = lngLabel + 1
                lngCombo = lngCombo + 1
            Case "Lista - OptionButton"
                Load Me.OptionList(lngOption + 1)
                Me.OptionList(lngOption + 1).Top = lngTop + 120
                Me.OptionList(lngOption + 1).Visible = True
                colControls.Add Me.OptionList(lngOption + 1), "K" & lngNP
                pPonerPropiedades lngNP, Me.OptionList(lngOption + 1)
                Load Me.Label(lngLabel + 1)
                Me.Label(lngLabel + 1).Top = lngTop + 150
                Me.Label(lngLabel + 1).Visible = True
                Me.Label(lngLabel + 1).Caption = strTitulo
                lngTop = lngTop + Me.OptionList(lngOption + 1).Height + 10
                lngLabel = lngLabel + 1
                lngOption = lngOption + 1
            Case "Variable"
                pPonerPropiedades lngNP, Me.Label1
            Case Else
    End Select
    Me.Picture2.Height = lngTop + 200
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargarControl"))
End Sub

Private Sub pLimpiaControles()
    On Error GoTo NotificaError
    Dim ctrIndex As Control
    Dim lngIndex As Long
    Me.Picture2.Height = 0
    For lngIndex = 1 To lngLabel
        Unload Label(lngIndex)
    Next
    For Each ctrIndex In colControls
        Unload ctrIndex
    Next
    Set colControls = New Collection
    lngTop = 0
    lngText = 0
    lngCombo = 0
    lngPicker = 0
    lngCheck = 0
    lngLabel = 0
    lngOption = 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaControles"))
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Picture2_Resize()
    On Error GoTo NotificaError
    If Me.Picture2.Height > Me.Picture1.Height Then
        Me.VScroll1.Visible = True
        Me.VScroll1.Min = 0
        Me.VScroll1.Max = Me.Picture2.Height - Me.Picture1.Height + 50
    Else
        Me.VScroll1.Visible = False
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Picture2_Resize"))
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    If arrParametros(2, fObtieneParametroControl(Me.Text(Index))) = "Entero" Then
        If KeyAscii = 8 Or Chr(KeyAscii) = "-" Or Chr(KeyAscii) = "+" Then Exit Sub
        If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
            KeyAscii = 0
        End If
    End If
    If arrParametros(2, fObtieneParametroControl(Me.Text(Index))) = "Numérico" Then
        If KeyAscii = 8 Or Chr(KeyAscii) = "-" Or Chr(KeyAscii) = "+" Or UCase(Chr(KeyAscii)) = "E" Or Chr(KeyAscii) = "." Then Exit Sub
        If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
            KeyAscii = 0
        End If
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Text_KeyPress"))
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo NoEsLong
    If arrParametros(2, fObtieneParametroControl(Me.Text(Index))) = "Entero" Then
        If Me.Text(Index).Text = "" Then Me.Text(Index).Text = "0": Exit Sub
        Me.Text(Index).Text = CLng(Me.Text(Index).Text)
    End If
    If arrParametros(2, fObtieneParametroControl(Me.Text(Index))) = "Numérico" Then
        If Me.Text(Index).Text = "" Then Me.Text(Index).Text = "0": Exit Sub
        Me.Text(Index).Text = CDbl(Me.Text(Index).Text)
    End If
    Exit Sub
NoEsLong:
    Me.Text(Index).Text = "0"
End Sub

Private Sub Timer1_Timer()
        On Error GoTo NotificaError
    FrmEspere.Visible = True
    Frame6.Enabled = False
    frmShowReport.Refresh
    pLimpiaControles
    strOriginalSelectionFormula = ""
    If Me.cboReporte.ListIndex > 0 Then
        Select Case Mid(arrReportes(1, Me.cboReporte.ListIndex), 2, 1)
            Case ":", "\"
                strArchivo = arrReportes(1, Me.cboReporte.ListIndex)
            Case Else
                strArchivo = App.Path & arrReportes(1, Me.cboReporte.ListIndex)
        End Select
        lngIdReporte = arrReportes(0, Me.cboReporte.ListIndex)
        pLInstanciaReporte vgrptReporte, strArchivo
        strOriginalSelectionFormula = vgrptReporte.RecordSelectionFormula
        pCargarFiltros
    End If
    Me.Timer1.Enabled = False
    FrmEspere.Visible = False
    Frame6.Enabled = True
    Exit Sub
NotificaError:
    Me.Timer1.Enabled = False
    If Err.Number = -2147206460 Or Err.Number = -2147206461 Then
        MsgBox "El Archivo asociado a este reporte no existe, favor de contactar al administrador del sistema", , "Mensaje"
        Me.cboReporte.ListIndex = 0
    Else
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Timer1_Timer"))
    End If
End Sub

Private Sub pPonerPropiedades(ByVal lngNP, ByVal ctlRef As Control)
    On Error GoTo NotificaError
    Dim rsPropiedades As ADODB.Recordset
    Dim rsQueryData As ADODB.Recordset
    Dim lngIdParametro As Long
    Dim ctlOptionList As ReporterControls.OptionList
    Dim ctlComboBox As ComboBox
    Dim ctlTextBox As TextBox
    Dim ctlDTPicker As DTPicker
    lngIdParametro = arrParametros(0, lngNP)
    Set rsPropiedades = frsRegresaRs("select * from SePropiedadesParametros where intIdParametro = " & lngIdParametro & " order by intIdPropiedad", adLockReadOnly, adOpenForwardOnly)
    Select Case arrParametros(3, lngNP)
        Case "Lista - OptionButton"
            Set ctlOptionList = ctlRef
            Do Until rsPropiedades.EOF
                If arrParametros(2, lngNP) = "Entero" Then
                    ctlOptionList.AddOption IIf(IsNull(rsPropiedades!vchDefinicion), "", rsPropiedades!vchDefinicion), rsPropiedades!intItemData, ""
                Else
                    ctlOptionList.AddOption IIf(IsNull(rsPropiedades!vchDefinicion), "", rsPropiedades!vchDefinicion), 0, IIf(IsNull(rsPropiedades!vchStringData), "", rsPropiedades!vchStringData)
                End If
                rsPropiedades.MoveNext
                ctlOptionList.SelectOption 1
            Loop
        Case "Lista - ComboBox"
            Set ctlComboBox = ctlRef
            Do Until rsPropiedades.EOF
                ctlComboBox.AddItem IIf(IsNull(rsPropiedades!vchDefinicion), "", rsPropiedades!vchDefinicion)
                ctlComboBox.ItemData(ctlComboBox.NewIndex) = rsPropiedades!intItemData
                rsPropiedades.MoveNext
                ctlComboBox.ListIndex = 0
            Loop
        Case "Consulta - ComboBox"
            On Error GoTo BadQuery
            Set ctlComboBox = ctlRef
            If Not rsPropiedades.EOF Then
                Set rsQueryData = frsRegresaRs(IIf(IsNull(rsPropiedades!vchDefinicion), "", rsPropiedades!vchDefinicion), adLockReadOnly, adOpenForwardOnly)
                Do Until rsQueryData.EOF
                    ctlComboBox.AddItem rsQueryData.Fields(CLng(rsPropiedades!intFieldNumberList)).Value
                    ctlComboBox.ItemData(ctlComboBox.NewIndex) = rsQueryData.Fields(CLng(rsPropiedades!intItemData)).Value
                    rsQueryData.MoveNext
                    ctlComboBox.ListIndex = 0
                Loop
                rsQueryData.Close
            End If
            On Error GoTo NotificaError
        Case "Dos valores - CheckBox"
            If Not rsPropiedades.EOF Then
                ReDim Preserve arrValues(2, UBound(arrValues, 2))
                arrValues(0, UBound(arrValues, 2)) = lngNP
                If arrParametros(2, lngNP) = "Entero" Then
                    arrValues(1, UBound(arrValues, 2)) = rsPropiedades!intItemData
                    rsPropiedades.MoveNext
                    If Not rsPropiedades.EOF Then
                        arrValues(2, UBound(arrValues, 2)) = rsPropiedades!intItemData
                    End If
                Else
                    arrValues(1, UBound(arrValues, 2)) = IIf(IsNull(rsPropiedades!vchStringData), "", rsPropiedades!vchStringData)
                    rsPropiedades.MoveNext
                    If Not rsPropiedades.EOF Then
                        arrValues(2, UBound(arrValues, 2)) = IIf(IsNull(rsPropiedades!vchStringData), "", rsPropiedades!vchStringData)
                    End If
                End If
                ReDim Preserve arrValues(2, UBound(arrValues, 2) + 1)
            End If
        Case "Variable"
            If Not rsPropiedades.EOF Then
                ReDim Preserve arrValues(2, UBound(arrValues, 2))
                arrValues(0, UBound(arrValues, 2)) = lngNP
                If rsPropiedades!vchDefinicion = "(Establecer valor)" Then
                    arrValues(1, UBound(arrValues, 2)) = "(Establecer valor)"
                    arrValues(2, UBound(arrValues, 2)) = rsPropiedades!vchStringData
                Else
                    arrValues(1, UBound(arrValues, 2)) = Mid(rsPropiedades!vchDefinicion, 1, InStr(1, rsPropiedades!vchDefinicion, "-") - 2)
                End If
                ReDim Preserve arrValues(2, UBound(arrValues, 2) + 1)
            End If
        Case "Libre - TextBox"
            Set ctlTextBox = ctlRef
            If arrParametros(2, lngNP) = "Entero" Or arrParametros(2, lngNP) = "Numérico" Then
                ctlTextBox.Text = "0"
            End If
        Case "Fecha ('yyyy-mm-dd') - DTPicker", "Fecha - DTPicker", "Fecha Hora ('yyyy-...') - DTPicker", "Fecha Hora - DTPicker"
            Set ctlDTPicker = ctlRef
            ctlDTPicker.Value = Date
            If Not rsPropiedades.EOF Then
                ReDim Preserve arrValues(2, UBound(arrValues, 2))
                arrValues(0, UBound(arrValues, 2)) = lngNP
                arrValues(1, UBound(arrValues, 2)) = IIf(IsNull(rsPropiedades!vchDefinicion), "yyyy-mm-dd HH:mm:ss", rsPropiedades!vchDefinicion)
                ReDim Preserve arrValues(2, UBound(arrValues, 2) + 1)
            End If

    End Select
    rsPropiedades.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPonerPropiedades"))
    Exit Sub
BadQuery:
    MsgBox SIHOMsg(628), vbInformation, "Mensaje"
End Sub

Private Sub VScroll1_Change()
    On Error GoTo NotificaError
    Me.Picture2.Top = -Me.VScroll1.Value
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":VScroll1_Change"))
End Sub

Private Function fObtieneParametroControl(ByVal ctlData As Control) As Integer
    On Error GoTo NotificaError
    Dim ctlText As Control
    Dim intIndex As Integer
    For intIndex = 0 To UBound(arrParametros, 2)
        If arrParametros(3, intIndex) <> "Variable" Then
            Set ctlText = colControls.Item("K" & intIndex)
            If ctlText.TabIndex = ctlData.TabIndex Then
                fObtieneParametroControl = intIndex
                Exit Function
            End If
        End If
    Next
    fObtieneParametroControl = -1
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fObtieneParametroControl"))
End Function

Private Function fObtieneValores(ByVal intNP As Integer, ByVal bytPos As Byte) As String
    On Error GoTo NotificaError
    Dim intIndex As Integer
    For intIndex = 0 To UBound(arrValues, 2)
        If CStr(intNP) = arrValues(0, intIndex) Then
            fObtieneValores = arrValues(bytPos, intIndex)
            Exit Function
        End If
    Next
    fObtieneValores = ""
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fObtieneValores"))
End Function

Private Sub pLInstanciaReporte(prptReporte As CRAXDRT.report, pstrNombreReporte As String)
    Dim Appn As CRAXDRT.Application
    Dim Obj  As Object
    Dim Sec As CRAXDRT.Section
    Set Appn = CreateObject("CrystalRunTime.Application")
    Set prptReporte = Appn.OpenReport(pstrNombreReporte)
    For Each Sec In prptReporte.Sections
        For Each Obj In Sec.ReportObjects
            If TypeOf Obj Is CRAXDRT.FieldObject Or TypeOf Obj Is CRAXDRT.TextObject Or TypeOf Obj Is CRAXDRT.SubreportObject Then
                Obj.ConditionFormula(crToolTipTextConditionFormulaType) = "Chr(9)"
            End If
        Next
    Next
    Exit Sub
End Sub

Private Function fAgregaParametro(ByVal strParametros As String, ByVal strNuevo As String) As String
    If strParametros = "" Then
        strParametros = strNuevo
    Else
        strParametros = strParametros & "|" & strNuevo
    End If
    fAgregaParametro = strParametros
End Function

Private Function fGeneraConsulta(ByVal strParametros As String, ByVal strCom As String) As String
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim strCar As String
    Dim intPosIni As Integer
    Dim strQuery As String
    Dim strParametro As String
    intPosIni = 1
    For intCount = 1 To Len(strCom)
        strParametro = ""
        strCar = Mid(strCom, intCount, 1)
        If strCar = "?" Then
            For intCount2 = 1 To Len(strParametros)
                If Mid(strParametros, intCount2, 1) <> "|" Then
                    strParametro = strParametro & Mid(strParametros, intCount2, 1)
                Else
                    strParametros = Mid(strParametros, intCount2 + 1, Len(strParametros))
                    Exit For
                End If
            Next
            strQuery = strQuery & Mid(strCom, intPosIni, intCount - intPosIni) & strParametro
            intPosIni = intCount + 1
        End If
    Next
    strQuery = strQuery & Mid(strCom, intPosIni, intCount - intPosIni)
    fGeneraConsulta = strQuery
End Function

