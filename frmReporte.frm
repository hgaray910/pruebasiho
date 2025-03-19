VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporte 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6735
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16245
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "frmReporte.frx":0000
         EndProperty
      EndProperty
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      lastProp        =   600
      _cx             =   18865
      _cy             =   11880
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuExport 
      Caption         =   "Export"
      Visible         =   0   'False
      Begin VB.Menu mnuPDF 
         Caption         =   "Formato PDF"
      End
      Begin VB.Menu mnuOtros 
         Caption         =   "Otros formatos"
      End
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents objSeccion As CRAXDRT.Section
Attribute objSeccion.VB_VarHelpID = -1
Private WithEvents objSeccion1 As CRAXDRT.Section
Attribute objSeccion1.VB_VarHelpID = -1
Public strRutaCBB As String
Public strRutaCBB1 As String

Private Sub CrystalActiveXReportViewer1_ExportButtonClicked(UseDefault As Boolean)
    If fblnUsarPDFCreator Then
        UseDefault = False
        PopupMenu mnuExport
    End If
End Sub

Private Sub mnuOtros_Click()
     Dim cryRpt As CRAXDRT.Report
     Set cryRpt = Me.CrystalActiveXReportViewer1.ReportSource
     cryRpt.Export True
End Sub

Private Sub mnuPDF_Click()
    pExportaPDF Me.CrystalActiveXReportViewer1.ReportSource
End Sub

Private Sub objSeccion_Format(ByVal pFormattingInfo As Object)
On Error GoTo NotificaError
    Set objSeccion.ReportObjects("CBB").FormattedPicture = LoadPicture(strRutaCBB)
    Exit Sub
NotificaError:
    Unload Me
End Sub

Private Sub objSeccion1_Format(ByVal pFormattingInfo As Object)
On Error GoTo NotificaError
    Set objSeccion1.ReportObjects("CBB1").FormattedPicture = LoadPicture(strRutaCBB1)
    Exit Sub
NotificaError:
    Unload Me
End Sub

Private Sub Form_Activate()
    CrystalActiveXReportViewer1.EnablePopupMenu = False
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    Me.Height = Screen.Height
    Me.Width = Screen.Width
End Sub

Public Sub pImprimeCBB(Reporte As CRAXDRT.Report)
    Set objSeccion = Nothing
    Set objSeccion = Reporte.Sections(fintSeccionCBB(Reporte))
End Sub

Public Sub pImprimeCBB1(Reporte As CRAXDRT.Report)
    Set objSeccion1 = Nothing
    Set objSeccion1 = Reporte.Sections(fintSeccionCBB1(Reporte))
End Sub

Private Sub Form_Resize()
    Me.CrystalActiveXReportViewer1.Top = 0
    Me.CrystalActiveXReportViewer1.Left = 0
    Me.CrystalActiveXReportViewer1.Height = ScaleHeight - (Me.StatusBar1.Height + 10)
    Me.CrystalActiveXReportViewer1.Width = ScaleWidth
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    Dim cryRpt As CRAXDRT.Report
    Dim objPrinter As Printer
    If Panel.Index = 2 Then
        Set cryRpt = Me.CrystalActiveXReportViewer1.ReportSource
        cryRpt.PrinterSetup Me.hWnd
        Me.CrystalActiveXReportViewer1.RefreshEx False
        If cryRpt.PrinterName = "" Then
            On Error Resume Next
            Set objPrinter = GetDefaultPrinter()
            Me.StatusBar1.Panels(2).Text = "Imprimir en: " & objPrinter.DeviceName
            If Err.Number <> 0 Then
                Me.StatusBar1.Panels(2).Text = "No existe impresora en el sistema"
            End If
            Err.Clear
        Else
            Me.StatusBar1.Panels(2).Text = "Imprimir en: " & cryRpt.PrinterName
        End If
    End If
End Sub

Private Sub pExportaPDF(cryRpt As CRAXDRT.Report)
On Error GoTo ErrHandler
    Dim resp As VbMsgBoxResult
    Dim cryDestino As CRExportDestinationType
    Dim strRuta As String
    Dim strIni As String
    Dim strFin As String
    Dim lngLastPage As Long
    
    resp = frmCRExport.ShowForm(Me, cryDestino, strIni, strFin)
    If resp = vbOK Then
    
        If IsNumeric(strIni) And IsNumeric(strFin) Then
            If CLng(strIni) = 0 Or CLng(strFin) = 0 Then
                Err.Raise _
                Number:=-555, _
                Description:="Error al generar PDF", _
                Source:="frmreporte.pExportaPDF"
            End If
            If CLng(strIni) > CLng(strFin) Then
                Err.Raise _
                Number:=-555, _
                Description:="Error al generar PDF", _
                Source:="frmreporte.pExportaPDF"
            End If
            Me.CrystalActiveXReportViewer1.GetLastPageNumber lngLastPage, True
            If CLng(strFin) > lngLastPage Then
                Err.Raise _
                Number:=-555, _
                Description:="Error al generar PDF", _
                Source:="frmreporte.pExportaPDF"
            End If
        End If
    
        strRuta = strRutaArchivo(cryRpt.ReportTitle, cryDestino)
        If (cryDestino = crEDTDiskFile And strRuta <> "") Or cryDestino = crEDTApplication Then
            Me.MousePointer = vbHourglass
            pGeneraPDF cryRpt, strRuta, strIni, strFin
            Me.MousePointer = vbNormal
            If cryDestino = crEDTApplication Then pAbrirArchivo strRuta
        End If
    End If
    Exit Sub
ErrHandler:
    Me.MousePointer = vbNormal
    MsgBox "Error al generar el PDF, verifique los parámetros", vbCritical, "Error"
End Sub

Private Function strRutaArchivo(strNombreArchivo, cryDestino As CRExportDestinationType) As String
On Error GoTo ErrHandler
    Dim strRuta As String
    Dim strNombre As String
    Dim StrFiltro As String
    Dim fso As New FileSystemObject
    
    strNombre = Mid(strNombreArchivo, 1, Len(strNombreArchivo) - 4) & ".pdf"
    StrFiltro = "Portable Document Format(*.pdf)|*.pdf"
  
    If cryDestino = crEDTApplication Then
        strRutaArchivo = fstrTmpPath & strNombre
    Else
        CommonDialog1.Filter = StrFiltro
        CommonDialog1.FileName = strNombre
        CommonDialog1.CancelError = True
        CommonDialog1.ShowSave
        strRutaArchivo = CommonDialog1.FileName
        If fso.FileExists(strRutaArchivo) Then
            If MsgBox("¿Desea reemplazar el archivo existente?", vbQuestion + vbYesNo, "Mensaje") = vbNo Then
                strRutaArchivo = ""
            End If
        End If
    End If
    Exit Function
ErrHandler:
    Err.Clear
    strRutaArchivo = ""
End Function
