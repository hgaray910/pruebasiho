VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPrecioCargosPaquete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Precios de cargos en paquetes"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   8330
      TabIndex        =   19
      Top             =   0
      Width           =   5135
      Begin VB.TextBox TxtMargenArt 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   210
         Width           =   650
      End
      Begin VB.TextBox TxtCalcular 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Nuevo importe"
         Top             =   210
         Width           =   1200
      End
      Begin VB.CommandButton CmdCalcular 
         Caption         =   "Distribuir"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3780
         TabIndex        =   20
         ToolTipText     =   "Distribuir nuevo importe"
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Margen"
         Height          =   255
         Left            =   2510
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo importe"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11965
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Total del paquete"
      Top             =   5520
      Width           =   1500
   End
   Begin VB.TextBox TxtIva 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11965
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Iva del paquete"
      Top             =   5160
      Width           =   1500
   End
   Begin VB.TextBox TxtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11965
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Importe del paquete"
      Top             =   4800
      Width           =   1500
   End
   Begin VB.ComboBox cboPaquetes 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Seleccionar el paquete"
      Top             =   160
      Width           =   7300
   End
   Begin VB.Frame Frame4 
      Caption         =   "Subrogados"
      Height          =   2295
      Left            =   128
      TabIndex        =   12
      Top             =   6000
      Width           =   13345
      Begin VB.TextBox TxtPrecio1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         MaxLength       =   15
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSubrogados 
         Height          =   1875
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Proveedores de subrogados"
         Top             =   240
         Width           =   13090
         _ExtentX        =   23098
         _ExtentY        =   3307
         _Version        =   393216
         Cols            =   6
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         MergeCells      =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.Frame Frame7 
      Height          =   700
      Left            =   6292
      TabIndex        =   11
      Top             =   8350
      Width           =   1080
      Begin VB.CommandButton cmdDelete 
         Height          =   495
         Left            =   540
         Picture         =   "frmPrecioCargosPaquete.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Eliminar información"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   495
         Left            =   52
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPrecioCargosPaquete.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Guardar el registro"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contenido del paquete"
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   13345
      Begin VB.TextBox txtMargen 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   600
         MaxLength       =   15
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         MaxLength       =   15
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPaquetes 
         Height          =   3555
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Detalle del paquete"
         Top             =   240
         Width           =   13090
         _ExtentX        =   23098
         _ExtentY        =   6271
         _Version        =   393216
         Cols            =   12
         GridColor       =   12632256
         MergeCells      =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
      End
   End
   Begin VB.Label lblEtiquetaF10 
      AutoSize        =   -1  'True
      Caption         =   "Presione <F10> para asignar este margen a todos los artículos"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   4410
   End
   Begin VB.Label Label60 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   11040
      TabIndex        =   17
      Top             =   5520
      Width           =   825
   End
   Begin VB.Label Label60 
      Caption         =   "IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   11040
      TabIndex        =   16
      Top             =   5160
      Width           =   765
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   11040
      TabIndex        =   15
      Top             =   4800
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paquete"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   220
      Width           =   600
   End
End
Attribute VB_Name = "frmPrecioCargosPaquete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public llngNumeroOpcionModulo As Long
Dim llngRowActualizar As Long
Dim IntNumCargo As Long
Dim strTipo As String
Dim intCtrl As Integer
Dim vgintColumnaCurrency As Integer
Dim vgintColumnaCurrency1 As Integer
Dim vlblnEditando As Boolean
Dim vldblImporte As Double
Dim vldblIVAIm As Double
Dim vldbltotal As Double
Dim vldblIVA As Long
Dim dblIVAPacienteP As Double
Dim blnGrdPaquete As Boolean
Dim blnTxtprecio As Boolean
Dim lngCveProveedor As Integer
Dim dblTxtPrecio1 As Double
Dim vllngPersonaGraba As Long
Dim dblgrdprecio As Double
Dim dblIvaPrecioProveedor As Double
Dim blnIvaPrecioProveedor As Boolean
Dim intGrdPaqueteR As Integer
Dim intGrdPaqueteC As Integer
Dim dimIntGrdTOp As Integer
Dim EnmCambiar As Integer
Dim vlblnPermitirModificar As Boolean

Dim intGrdPaqueteRPro As Integer
Dim intGrdPaqueteCPro As Integer
Dim blnCalcular As Boolean
Dim dblCalcular As Double

Const intColCosto = 6
Const intColMargenUtilidad = 7
Const intColPrecioUnitario = 8
Const intColImporte = 9
Const intColIVANormal = 10
Const intColTotal = 11
Const intColIVAPrecioProv = 12
Const intColObtenerCargo = 13
'Caso 20471
Dim blnPrecioCero As Boolean

Private Sub cboPaquetes_Click()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlintContador As Integer
    
    vldblIVA = 0
    If cboPaquetes.ListIndex > -1 Then
        vlstrSentencia = " Select pvConceptoFacturacion.SMYIVA From PvPaquete "
        vlstrSentencia = vlstrSentencia & "   Inner Join PvConceptoFacturacion On (pvPaquete.SMICONCEPTOFACTURA = pvConceptoFacturacion.SMICVECONCEPTO)"
        vlstrSentencia = vlstrSentencia & " Where pvPaquete.INTNUMPAQUETE = " & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex))
        vldblIVA = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)!smyIVA
    
        grdPaquetes.Clear
        pConfiguraGrid
        pConfiguraGridSub
        
        pLlenaGrid Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex))
        
        vlblnPermitirModificar = True
        If Not fblnPaqueteEnListas(cboPaquetes.ItemData(cboPaquetes.ListIndex)) Then
            'No se podrán modificar los precios del contenido del paquete porque no se encontró configurado en alguna lista de precios.
            MsgBox SIHOMsg(1593), vbOKOnly + vbInformation, "Mensaje"
            vlblnPermitirModificar = False
            
            grdPaquetes.Clear
            pConfiguraGrid
            pConfiguraGridSub
            
            TxtImporte.Text = Format(0, "$ ###,###,###,###.00")
            TxtIva.Text = Format(0, "$ ###,###,###,###.00")
            TxtTotal.Text = Format(0, "$ ###,###,###,###.00")
            
            cmdGrabarRegistro.Enabled = False
            cmdDelete.Enabled = False
            
        Else
            If blnPrecioCero = True Then
                'Se encontraron precios en cero, favor de revisar.
                MsgBox SIHOMsg(1593), vbOKOnly + vbInformation, "Mensaje"
                blnPrecioCero = False
            End If
            
        End If
        
        For vlintContador = 1 To grdPaquetes.Rows - 1
            grdPaquetes.TextMatrix(vlintContador, intColObtenerCargo) = ""
        Next
        
        For vlintContador = 1 To grdSubrogados.Rows - 1
            grdSubrogados.TextMatrix(vlintContador, 8) = ""
        Next
        
        grdPaquetes.Row = 1
        grdPaquetes.Col = vgintColumnaCurrency
    End If
End Sub

Private Sub cboPaquetes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub CmdCalcular_Click()
Dim i As Integer
Dim X As Integer
Dim dblCorrecto As Double
Dim intRenglon As Integer
Dim intRenglonAjuste As Integer
Dim dblMargen As Double
Dim dblImporteSinArtículos As Double
Dim dblImporteSoloArtículos As Double
Dim dblCosto As Double
Dim dbldiferencia As Double
Dim Intfaltantes As Integer
Dim dblPrecioOtros As Double
Dim blnTipoDiferencia As Boolean
'Caso 20471
Dim dblTotalOtros As Double
Dim dblPorcentajeOtros As Double

Dim intTotalArt As Integer

    intRenglonAjuste = 1
    
    dblCalcular = 0
    'If (CDbl(Replace(Replace(TxtCalcular.Text, "$", ""), ",", "")) > CDbl(Replace(Replace(TxtImporte.Text, "$", ""), ",", ""))) And (CDbl(Replace(TxtMargenArt, "%", "")) > 0) Then
    'Caso 20471 - se agregò esta lìnea
    If (CDbl(Replace(Replace(TxtCalcular.Text, "$", ""), ",", "")) > 0) And (CDbl(Replace(TxtMargenArt, "%", "")) > 0) Then
        If CDbl(Replace(Replace(TxtCalcular.Text, "$", ""), ",", "")) > 0 Then
            'Este se usa si solo hay un cargo en el paquete
            If grdPaquetes.Rows - 1 = 1 Then
                grdPaquetes.TextMatrix(1, intColPrecioUnitario) = TxtCalcular.Text
                
                If grdPaquetes.TextMatrix(i, 3) = "AR" Then
                    grdPaquetes.TextMatrix(1, intColMargenUtilidad) = FormatPercent(CDbl(Replace(grdPaquetes.TextMatrix(i, intColPrecioUnitario), "$", "")) / CDbl(Replace(grdPaquetes.TextMatrix(i, intColCosto), "$", "")) - 1)
                Else
                    grdPaquetes.TextMatrix(1, intColMargenUtilidad) = ""
                End If
                
                grdPaquetes.TextMatrix(1, intColImporte) = TxtCalcular.Text
                grdPaquetes.TextMatrix(1, intColIVANormal) = FormatCurrency(CDbl(vldblIVA / 100) * CDbl(Replace(grdPaquetes.TextMatrix(1, intColImporte), "$", "")), 2)
                grdPaquetes.TextMatrix(1, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(1, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(1, intColIVANormal), "$", "")), 2)
                grdPaquetes.TextMatrix(1, intColObtenerCargo) = 1
                
                TxtImporte.Text = TxtCalcular.Text
                TxtIva.Text = FormatCurrency((vldblIVA / 100) * CDbl(Replace(grdPaquetes.TextMatrix(1, intColImporte), "$", "")), 2)
                TxtTotal.Text = Format(Val(Format(TxtImporte.Text, "")) + Val(Format(TxtIva.Text, "")), "$ ###,###,###,##0.00")
            'Para paquetes con más de un cargo
            Else
                dblImporteSinArtículos = 0
                dblImporteSoloArtículos = 0
                intTotalArt = 0
                dblMargen = CDbl(Replace(TxtMargenArt, "%", ""))
                'Agregue esto
                'Para los que son Artículos y Grupos de Cargos con subtipo Artículo
                For i = 1 To grdPaquetes.Rows - 1
                    dblCosto = 0
                    If (grdPaquetes.TextMatrix(i, 3) = "AR") Or (grdPaquetes.TextMatrix(i, 3) = "GC" And grdPaquetes.TextMatrix(i, 14) = "AR") Then
                        intTotalArt = intTotalArt + 1
                        dblCosto = CDbl(Replace(grdPaquetes.TextMatrix(i, intColCosto), "$", ""))
                        grdPaquetes.TextMatrix(i, intColPrecioUnitario) = FormatCurrency((dblCosto * (dblMargen / 100)) + dblCosto, 2)
                        
                        grdPaquetes.TextMatrix(i, intColMargenUtilidad) = TxtMargenArt.Text
                        grdPaquetes.TextMatrix(i, intColImporte) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColPrecioUnitario), "$", "")) * CInt(Replace(grdPaquetes.TextMatrix(i, 4), "$", "")), 2)
                        grdPaquetes.TextMatrix(i, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")), 2)
                        grdPaquetes.TextMatrix(i, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(i, intColIVANormal), "$", "")), 2)
                        grdPaquetes.TextMatrix(i, intColObtenerCargo) = 1
                        dblImporteSoloArtículos = dblImporteSoloArtículos + CDbl(grdPaquetes.TextMatrix(i, intColImporte))
                                                
                    End If
                    If (CDbl(Replace(Replace(TxtCalcular.Text, "$", ""), ",", "")) <= dblImporteSoloArtículos) Then
                        Call cboPaquetes_Click
                        'No se puede realizar el proceso, el importe capturado no alcanza a cubrir el porcentaje de los artículos y artículos de grupo de cargo
                        MsgBox "No se puede realizar el proceso, el importe capturado no alcanza a cubrir el porcentaje de los artículos y artículos de grupo de cargo", vbCritical, "Mensaje"
                        Exit Sub
                    End If
                Next

                'total que queda para repartir entre los que quedan
                dblImporteSinArtículos = (CDbl(Replace(Replace(TxtCalcular.Text, "$", ""), ",", "")) - dblImporteSoloArtículos)
                
                'Para saber cuantos cargos quedan sin distribuir
                dblTotalOtros = 0
                For i = 1 To grdPaquetes.Rows - 1
                    If (grdPaquetes.TextMatrix(i, 3) <> "AR") Then
                        If (grdPaquetes.TextMatrix(i, 3) = "GC" And grdPaquetes.TextMatrix(i, 14) = "AR") Then
                        Else
                            Intfaltantes = Intfaltantes + 1 'CInt(grdPaquetes.TextMatrix(i, 4))
                            dblTotalOtros = dblTotalOtros + CDbl(Replace(grdPaquetes.TextMatrix(i, 15), "$", ""))
                        End If
                    End If
                Next
                
                
                If Intfaltantes > 0 Then
                    'dblPrecioOtros = dblImporteSinArtículos / Intfaltantes
                    
                    For i = 1 To grdPaquetes.Rows - 1
                        If (grdPaquetes.TextMatrix(i, 3) <> "AR") Then
                            If (grdPaquetes.TextMatrix(i, 3) = "GC" And grdPaquetes.TextMatrix(i, 14) = "AR") Then
                            Else
                                dblPorcentajeOtros = (CDbl(Replace(grdPaquetes.TextMatrix(i, 15), "$", "")) * 100) / dblTotalOtros
                                grdPaquetes.TextMatrix(i, 16) = FormatPercent(dblPorcentajeOtros / 100, 2)
                                
                                grdPaquetes.TextMatrix(i, intColPrecioUnitario) = FormatCurrency(dblImporteSinArtículos * (dblPorcentajeOtros / 100), 2)
                                
                                grdPaquetes.TextMatrix(i, intColImporte) = grdPaquetes.TextMatrix(i, intColPrecioUnitario)
                                'grdPaquetes.TextMatrix(i, intColImporte) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColPrecioUnitario), "$", "")) * CInt(Replace(grdPaquetes.TextMatrix(i, 4), "$", "")), 2)
                                grdPaquetes.TextMatrix(i, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")), 2)
                                grdPaquetes.TextMatrix(i, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(i, intColIVANormal), "$", "")), 2)
                                grdPaquetes.TextMatrix(i, intColObtenerCargo) = 1
                                dblImporteSoloArtículos = dblImporteSoloArtículos + CDbl(grdPaquetes.TextMatrix(i, intColImporte))
                            End If
                        End If
                    Next
                    
                    'If CDbl(Replace(Replace(TxtCalcular.Text, "$", ""), ",", "")) <> dblImporteSoloArtículos Then
                    '    dbldiferencia = Format(Val(Format(dblImporteSoloArtículos, "############.00")) - Val(Format(TxtCalcular.Text, "############.00")), "############.00")
                    '
                    '    If dbldiferencia > 0 Then
                    '        For i = 1 To grdPaquetes.Rows - 1
                    '            If grdPaquetes.TextMatrix(i, 3) <> "AR" Then
                    '                'grdPaquetes.TextMatrix(i, intColPrecioUnitario) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColPrecioUnitario), "$", "")) - dbldiferencia, 2)
                    '
                    '                grdPaquetes.TextMatrix(i, intColImporte) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) - dbldiferencia, 2)
                    '                grdPaquetes.TextMatrix(i, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")), 2)
                    '                grdPaquetes.TextMatrix(i, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(i, intColIVANormal), "$", "")), 2)
                    '                grdPaquetes.TextMatrix(i, intColObtenerCargo) = 1
                    '                dblImporteSoloArtículos = dblImporteSoloArtículos - dbldiferencia
                    '
                    '                Exit For
                    '            End If
                    '        Next
                    '    Else
                    '        For i = 1 To grdPaquetes.Rows - 1
                    '            If grdPaquetes.TextMatrix(i, 3) <> "AR" Then
                    '                'grdPaquetes.TextMatrix(i, intColPrecioUnitario) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColPrecioUnitario), "$", "")) + dbldiferencia, 2)
                    '
                    '                grdPaquetes.TextMatrix(i, intColImporte) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) + dbldiferencia, 2)
                    '                grdPaquetes.TextMatrix(i, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")), 2)
                    '                grdPaquetes.TextMatrix(i, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(i, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(i, intColIVANormal), "$", "")), 2)
                    '                grdPaquetes.TextMatrix(i, intColObtenerCargo) = 1
                    '                dblImporteSoloArtículos = dblImporteSoloArtículos + dbldiferencia
                    '
                    '                Exit For
                    '            End If
                    '        Next
                    '    End If
                    'End If
                End If
                
                TxtImporte.Text = FormatCurrency(dblImporteSoloArtículos, 2)
                TxtIva.Text = FormatCurrency((vldblIVA / 100) * dblImporteSoloArtículos, 2)
                TxtTotal.Text = Format(Val(Format(TxtImporte.Text, "")) + Val(Format(TxtIva.Text, "")), "$ ###,###,###,##0.00")
            
            End If
            cmdGrabarRegistro.Enabled = True
            'Hasta aquí
        End If
        
    Else
        'If (CDbl(Replace(Replace(TxtCalcular.Text, "$", ""), ",", "")) <= CDbl(Replace(Replace(TxtImporte.Text, "$", ""), ",", ""))) Then
            'El importe capturado no puede ser menor al importe del paquete
        '    MsgBox "El importe capturado no puede ser menor o igual al importe del paquete", vbCritical, "Mensaje"
        'Else
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbCritical, "Mensaje"
        'End If
    End If
End Sub

Private Sub cmdDelete_Click()
On Error GoTo NotificaError
    Dim rsDelSubrogados As ADODB.Recordset
    Dim vlstrSentenciaX As String
    Dim intCtrlX As Integer
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
               
        If vllngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            '¿Está seguro de eliminar los datos?
            If MsgBox(SIHOMsg(6), vbYesNo + vbQuestion + vbDefaultButton2, "Mensaje") = vbYes Then
                pEjecutaSentencia " UPDATE PVDETALLEPAQUETE SET MNYPRECIOESPECIFICO = 0 WHERE PVDETALLEPAQUETE.INTNUMPAQUETE = " & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex))
                pEjecutaSentencia " DELETE COSERVICIOSUBROGADOPAQUETE WHERE COSERVICIOSUBROGADOPAQUETE.INTNUMPAQUETE = " & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex))
                grdPaquetes.Refresh
                grdSubrogados.Refresh
                
                grdPaquetes.Clear
                pConfiguraGrid
    
                pConfiguraGridSub
    
                pLlenaGrid Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex))
                  
                '¡Los datos han sido guardados satisfactoriamente!
                MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
            Else
                grdPaquetes.SetFocus
            End If
            
            cmdGrabarRegistro.Enabled = False
            cmdDelete.Enabled = False
                        
            Call pGuardarLogTransaccion(Me.Name, 3, vllngPersonaGraba, "BORRAR PRECIO DEL CARGO EN EL PAQUETE", CStr(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario), "$", "")))
            
            EntornoSIHO.ConeccionSIHO.CommitTrans
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Sub cmdGrabarRegistro_Click()
    Dim vlstrSentencia As String
    Dim rsSubrogados As New ADODB.Recordset         'RS tipo tabla para guardar la fractura
    Dim rsPrecioSub As New ADODB.Recordset
    Dim rsCantSubrogados As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim vstrSentencia As String
    Dim vlintcontador0 As Integer

On Error GoTo NotificaError
 
    For vlintcontador0 = 1 To grdPaquetes.Rows - 1
        If Replace(Trim(grdPaquetes.TextMatrix(vlintcontador0, intColPrecioUnitario)), "$", "") = 0 Then
            'No se pueden guardar precios unitarios en ceros
            MsgBox SIHOMsg(1586) & Chr(13) & grdPaquetes.TextMatrix(vlintcontador0, 2), vbOKOnly + vbExclamation, "Mensaje"
            grdPaquetes.Row = vlintcontador0
            grdPaquetes.Col = intColPrecioUnitario
            grdPaquetes.SetFocus
            Exit Sub
        End If
    Next
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
    
        With grdPaquetes
            For vlintContador = 1 To .Rows - 1
                If grdPaquetes.TextMatrix(vlintContador, intColObtenerCargo) <> "" Then
                    vlstrSentencia = "UPDATE PVDETALLEPAQUETE SET MNYPRECIOESPECIFICO = " & Round(CDbl(Replace(grdPaquetes.TextMatrix(vlintContador, intColPrecioUnitario), "$", "")), 2) & _
                    " WHERE INTNUMPAQUETE = " & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex)) & _
                    " AND INTCVECARGO = " & Trim(grdPaquetes.TextMatrix(vlintContador, 1)) & _
                    " AND CHRTIPOCARGO = '" & Trim(grdPaquetes.TextMatrix(vlintContador, 3)) & "'"
                    
                    pEjecutaSentencia vlstrSentencia
                End If
            Next
        End With
        
        grdPaquetes.Refresh
    
        If TxtImporte.Text <> "" And Replace(TxtImporte.Text, "$", "") > 0 Then
            vlstrSentencia = "UPDATE PVDETALLELISTA SET MNYPRECIO = " & CDbl(Replace(TxtImporte.Text, "$", "")) & _
            " WHERE CHRCVECARGO = " & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex)) & _
            " AND CHRTIPOCARGO = '" & "PA" & "'"
            
            pEjecutaSentencia vlstrSentencia
        End If

        For vlintcontador0 = 1 To grdSubrogados.Rows - 1
            If grdSubrogados.TextMatrix(vlintcontador0, 5) <> "" Then
                vlstrSentencia = " Select * from COSERVICIOSUBROGADOPAQUETE where INTNUMPAQUETE = " & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex)) & _
                            " AND INTCVEPROVEEDOR = " & grdSubrogados.TextMatrix(vlintcontador0, 5) & _
                            " AND INTCVECARGO = " & grdSubrogados.TextMatrix(vlintcontador0, 6)
                Set rsCantSubrogados = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                
                If rsCantSubrogados.RecordCount = 0 Then
                    vlstrSentencia = ""
                    If Val(grdSubrogados.TextMatrix(vlintcontador0, 8)) = 1 Then
                        vstrSentencia = "Insert into COSERVICIOSUBROGADOPAQUETE(INTNUMPAQUETE,INTCVEPROVEEDOR,INTCVECARGO,CHRTIPOSERVICIO,MNYCANTIDAD,INTTIPOACUERDO,MNYPRECIOESPECIFICO) " & _
                                        "values(" & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex)) & "," & grdSubrogados.TextMatrix(vlintcontador0, 5) & "," & grdSubrogados.TextMatrix(vlintcontador0, 6) & ",'" & grdSubrogados.TextMatrix(vlintcontador0, 7) & "'," & grdSubrogados.TextMatrix(vlintcontador0, 2) & "," & IIf(grdSubrogados.TextMatrix(vlintcontador0, 3) = "$", 0, 1) & "," & Val(Format(grdSubrogados.TextMatrix(vlintcontador0, 4), "")) & ") "
                        pEjecutaSentencia vstrSentencia
                    End If
                Else
                    If Val(grdSubrogados.TextMatrix(vlintcontador0, 8)) = 1 Then
                        pEjecutaSentencia " UPDATE COSERVICIOSUBROGADOPAQUETE SET MNYPRECIOESPECIFICO = " & Val(Format(grdSubrogados.TextMatrix(vlintcontador0, 4), "")) & "where INTNUMPAQUETE = " & Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex)) & _
                                          " AND INTCVEPROVEEDOR = " & grdSubrogados.TextMatrix(vlintcontador0, 5) & _
                                          " AND INTCVECARGO = " & grdSubrogados.TextMatrix(vlintcontador0, 6)
                    End If
                End If
            End If
        Next
        
        '¡Los datos han sido guardados satisfactoriamente!
        MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Mensaje"
       
        With grdPaquetes
            For vlintContador = 1 To .Rows - 1
                .TextMatrix(vlintContador, intColObtenerCargo) = ""
            Next
        End With
        
        For vlintContador = 1 To grdSubrogados.Rows - 1
            grdSubrogados.TextMatrix(grdSubrogados.Row, 8) = ""
        Next
        
        grdPaquetes.Clear
        grdSubrogados.Clear
        
        pConfiguraGrid
        pLlenaGrid Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex))
        
        pConfiguraGridSub
        
        grdPaquetes.Row = 1
        grdPaquetes.Col = vgintColumnaCurrency
        grdPaquetes.Refresh
        grdPaquetes.SetFocus
        
        cmdGrabarRegistro.Enabled = False
        cmdDelete.Enabled = False
        
        Call pGuardarLogTransaccion(Me.Name, 2, vllngPersonaGraba, "PRECIO DEL CARGO EN EL PAQUETE", CStr(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario), "$", "")))
        EntornoSIHO.ConeccionSIHO.CommitTrans
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Unload Me
 End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    
    vgintColumnaCurrency = intColPrecioUnitario
    vgintColumnaCurrency1 = 4
    pLlenaCombos
    pConfiguraGrid
    pConfiguraGridSub
    
    TxtImporte.Text = Format(0, "$ ###,###,###,###.00")
    TxtIva.Text = Format(0, "$ ###,###,###,###.00")
    TxtTotal.Text = Format(0, "$ ###,###,###,###.00")
    
    blnGrdPaquete = False
    cmdGrabarRegistro.Enabled = False
    cmdDelete.Enabled = False
    Label2.Enabled = False
    'pLlenaGrid
    'pInicia
    blnCalcular = False
    
    TxtCalcular.Enabled = False
    CmdCalcular.Enabled = False
    Label2.Enabled = False
    TxtMargenArt.Enabled = False
    Label3.Enabled = False
    
End Sub

Private Sub pLlenaCombos()
On Error GoTo NotificaError

    Dim rsAux As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsPaquete As New ADODB.Recordset
    
    vlstrSentencia = "Select * From pvpaquete Where bitactivo = 1 and intnumpaquete in (SELECT PvDetallePaquete.intnumpaquete From PvDetallePaquete) order by chrdescripcion"
    Set rsPaquete = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)
    If rsPaquete.RecordCount > 0 Then
        pLlenarCboRs cboPaquetes, rsPaquete, 0, 1
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCombos"))
End Sub

Private Sub pConfiguraGrid()
On Error GoTo NotificaError

    With grdPaquetes
        .Clear
        .Rows = 2
        .FormatString = "||Descripción|Tipo|Cantidad|Unidad|Costo|Margen|Precio unitario|Importe|IVA|Total|"
        .Cols = 17
        
        .ColWidth(0) = 180                      '
        .ColWidth(1) = 0                        'clave cargo
        .ColWidth(2) = 3700                     'Descripción 5940
        .ColWidth(3) = 400                      'Tipo
        .ColWidth(4) = 700                      'Cantidad
        .ColWidth(5) = 1200                     'Unidad
        .ColWidth(intColCosto) = 1100           'Costo
        .ColWidth(intColMargenUtilidad) = 1100  'Margen Utilidad
        .ColWidth(intColPrecioUnitario) = 1100  'Precio Unitario
        .ColWidth(intColImporte) = 1100         'Importe
        .ColWidth(intColIVANormal) = 1100       'IVA normal
        .ColWidth(intColTotal) = 1100           'Total
        .ColWidth(intColIVAPrecioProv) = 0      'IVAPrecioProv
        .ColWidth(intColObtenerCargo) = 0       'obtener cargo
        .ColWidth(14) = 0                       'GC - tipo
        .ColWidth(15) = 0                       'Otros - Precio
        .ColWidth(16) = 0                       'Otros - Porcentaje
              
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(intColCosto) = flexAlignRightCenter
        .ColAlignment(intColMargenUtilidad) = flexAlignRightCenter
        .ColAlignment(intColPrecioUnitario) = flexAlignRightCenter
        .ColAlignment(intColImporte) = flexAlignRightCenter
        .ColAlignment(intColIVANormal) = flexAlignRightCenter
        .ColAlignment(intColTotal) = flexAlignRightCenter
        .ColAlignment(14) = flexAlignRightCenter
        .ColAlignment(15) = flexAlignRightCenter
        .ColAlignment(16) = flexAlignRightCenter

        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignLeftCenter
        .ColAlignmentFixed(2) = flexAlignLeftCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(intColCosto) = flexAlignCenterCenter
        .ColAlignmentFixed(intColMargenUtilidad) = flexAlignCenterCenter
        .ColAlignmentFixed(intColPrecioUnitario) = flexAlignCenterCenter
        .ColAlignmentFixed(intColImporte) = flexAlignCenterCenter
        .ColAlignmentFixed(intColIVANormal) = flexAlignCenterCenter
        .ColAlignmentFixed(intColTotal) = flexAlignCenterCenter
        .ColAlignmentFixed(14) = flexAlignCenterCenter
        .ColAlignmentFixed(15) = flexAlignCenterCenter
        .ColAlignmentFixed(16) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub pConfiguraGridSub()
On Error GoTo NotificaError

    With grdSubrogados
        .Clear
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Proveedor|Cantidad acuerdo|Tipo|Cantidad paquete||"
        
        .ColWidth(0) = 180      '
        .ColWidth(1) = 10000    'Nombre Proveedor
        .ColWidth(2) = 2100     'Cantidad acuerdo
        .ColWidth(3) = 1000     'Tipo
        .ColWidth(4) = 1900     'Cantidad paquete
        .ColWidth(5) = 0        'cve proveedor
        .ColWidth(6) = 0        'cve servicio
        .ColWidth(7) = 0        'tipo servicio
        .ColWidth(8) = 0        '
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter 'flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridSub"))
End Sub

Private Sub pLlenaGrid(ByVal intPaq As Integer)
    Dim rs As ADODB.Recordset
    Dim rsCostoGrupo As ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsPrecioLista As ADODB.Recordset
    
    blnPrecioCero = False
    
    vldblImporte = 0
    vldblIVAIm = 0
    TxtImporte.Text = ""
    TxtCalcular.Text = ""
    TxtMargenArt.Text = ""
       
    vlstrSentencia = "SELECT " & _
       " Case when DETP.chrTipoCargo = 'AR' then ARTI.vchNombreComercial when DETP.chrTipoCargo = 'ES' then EST.vchNombre when DETP.chrTipoCargo = 'OC' then OTROCO.chrDescripcion when DETP.chrTipoCargo = 'EX' then EXA.chrNombre when DETP.chrTipoCargo = 'GE' then GRUPEX.chrNombre when DETP.chrTipoCargo = 'GC' then GRUPCA.vchNombre Else 'Invalido' end As DESCRIPCION, " & _
       " DETP.CHRTIPOCARGO TIPO, " & _
       " DETP.SMICANTIDAD CANTIDAD, " & _
       " Case DETP.INTDESCUENTOINVENTARIO WHEN 1 THEN ivum.VCHDESCRIPCION WHEN 2 THEN ivua.VCHDESCRIPCION END UNIDAD, " & _
       " DETP.MNYPRECIOESPECIFICO mnyPrecio, DETP.MNYPRECIO mnyPrecioPaq, DETP.INTCVECARGO, DETP.mnyIVA, ARTI.chrCveArticulo, " & _
       " case when PAQ.bitcostobase = 0 then IVARTICULOEMPRESAS.MnyCostoMasAlto else IVARTICULOEMPRESAS.MnyCostoUltEntrada end / (CASE WHEN DETP.INTDESCUENTOINVENTARIO = 1 THEN ARTI.INTCONTENIDO ELSE 1 END) mnycosto, PAQ.bitcostobase, DETP.INTDESCUENTOINVENTARIO, " & _
       " case when DETP.chrTipoCargo = 'GC' Then GRUPCA.chrtipo Else Null End As chrtipo, " & _
       " case when DETP.chrTipoCargo <> 'AR' then (select NVL(MNYPRECIO,0) from PVDETALLELISTA where chrtipocargo = DETP.chrTipoCargo " & _
       "                                     AND intcvelista = (Select pvlistaprecio.INTCVELISTA from pvlistaprecio where pvlistaprecio.bitpredeterminada = 1) " & _
       "                                     AND CHRCVECARGO = DETP.intCveCargo) " & _
       " Else null End As precioOC " & _
       " From PvDetallePaquete DETP " & _
       " LEFT JOIN pvpaquete PAQ ON DETP.intnumpaquete = PAQ.intnumpaquete " & _
       " LEFT JOIN PvOtroConcepto OTROCO ON DETP.intCveCargo = OTROCO.intCveConcepto " & _
       " LEFT JOIN LaExamen EXA ON DETP.intCveCargo = EXA.IntCveExamen " & _
       " LEFT JOIN LaGrupoExamen GRUPEX ON DETP.intCveCargo = GRUPEX.IntCveGrupo " & _
       " LEFT JOIN ImEstudio EST ON DETP.intCveCargo = EST.intCveEstudio " & _
       " LEFT JOIN PvGrupoCargo GRUPCA ON DETP.intCveCargo = GRUPCA.intCveGrupo " & _
       " LEFT JOIN IvArticulo ARTI ON DETP.intCveCargo = ARTI.intIdArticulo " & _
       " LEFT JOIN IVARTICULOEMPRESAS ON DETP.chrtipocargo = 'AR' and ARTI.chrcvearticulo = IVARTICULOEMPRESAS.chrcvearticulo AND IVARTICULOEMPRESAS.tnyclaveempresa = " & vgintClaveEmpresaContable & _
       " LEFT JOIN ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ARTI.intCveUniAlternaVta " & _
       " LEFT JOIN ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ARTI.intCveUniMinimaVta " & _
       " Where DETP.intNumPaquete = " & cboPaquetes.ItemData(cboPaquetes.ListIndex) & _
       " order by DESCRIPCION,tipo "
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    intCtrl = 1
    
    grdPaquetes.Redraw = False
    Do Until rs.EOF
        If intCtrl > 1 Then
            grdPaquetes.AddItem ""
        End If
         
        grdPaquetes.TextMatrix(intCtrl, 1) = IIf(IsNull(rs!intCveCargo), "", rs!intCveCargo)
        grdPaquetes.TextMatrix(intCtrl, 2) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
        grdPaquetes.TextMatrix(intCtrl, 3) = IIf(IsNull(rs!tipo), "", rs!tipo)
        grdPaquetes.TextMatrix(intCtrl, 4) = IIf(IsNull(rs!cantidad), "", rs!cantidad)
        grdPaquetes.TextMatrix(intCtrl, 5) = IIf(IsNull(rs!UNIDAD), "", rs!UNIDAD)
        
        If rs!mnyPrecio > 0 Then
            grdPaquetes.TextMatrix(intCtrl, intColPrecioUnitario) = FormatCurrency(IIf(IsNull(rs!mnyPrecio), "", rs!mnyPrecio), 2)
        Else
            grdPaquetes.TextMatrix(intCtrl, intColPrecioUnitario) = FormatCurrency(IIf(IsNull(rs!mnyPrecioPaq), "", rs!mnyPrecioPaq), 2)
        End If
        
        If rs!mnyPrecio > 0 Then
            If (grdPaquetes.TextMatrix(intCtrl, 3) <> "AR") Then
                If (grdPaquetes.TextMatrix(intCtrl, 3) = "GC" And rs!chrTipo = "AR") Then
                    grdPaquetes.TextMatrix(intCtrl, intColImporte) = FormatCurrency(IIf(IsNull(rs!cantidad), "", rs!cantidad) * IIf(IsNull(rs!mnyPrecio), "", rs!mnyPrecio), 2)
                Else
                    grdPaquetes.TextMatrix(intCtrl, intColImporte) = FormatCurrency(IIf(IsNull(rs!mnyPrecio), "", rs!mnyPrecio), 2)
                End If
            Else
                grdPaquetes.TextMatrix(intCtrl, intColImporte) = FormatCurrency(IIf(IsNull(rs!cantidad), "", rs!cantidad) * IIf(IsNull(rs!mnyPrecio), "", rs!mnyPrecio), 2)
            End If
        Else
            grdPaquetes.TextMatrix(intCtrl, intColImporte) = FormatCurrency(IIf(IsNull(rs!cantidad), "", rs!cantidad) * IIf(IsNull(rs!mnyPrecioPaq), "", rs!mnyPrecioPaq), 2)
        End If
        
        grdPaquetes.TextMatrix(intCtrl, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(intCtrl, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(intCtrl, intColImporte), "$", "")), 2)
        grdPaquetes.TextMatrix(intCtrl, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(intCtrl, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(intCtrl, intColIVANormal), "$", "")), 2) 'total
        
        grdPaquetes.TextMatrix(intCtrl, intColCosto) = FormatCurrency(IIf(IsNull(rs!mnycosto), 0, rs!mnycosto), 2)
        
        If IIf(IsNull(rs!tipo), "", rs!tipo) = "AR" Then
            If (rs!mnycosto) = 0 Then
                If (rs!mnyPrecio) > 0 Then
                    grdPaquetes.TextMatrix(intCtrl, intColMargenUtilidad) = FormatPercent(100, 2)
                Else
                    grdPaquetes.TextMatrix(intCtrl, intColMargenUtilidad) = FormatPercent(0, 2)
                End If
            Else
                If rs!mnyPrecio > 0 Then
                    grdPaquetes.TextMatrix(intCtrl, intColMargenUtilidad) = FormatPercent(Round(IIf(IsNull(rs!mnyPrecio), 0, rs!mnyPrecio) / IIf(IsNull(rs!mnycosto), 1, Round(rs!mnycosto, 2)), 4) - 1)
                Else
                    grdPaquetes.TextMatrix(intCtrl, intColMargenUtilidad) = FormatPercent(Round(IIf(IsNull(rs!mnyPrecioPaq), 0, rs!mnyPrecioPaq) / IIf(IsNull(rs!mnycosto), 1, Round(rs!mnycosto, 2)), 4) - 1)
                End If
            End If
        Else
            grdPaquetes.TextMatrix(intCtrl, intColMargenUtilidad) = ""
        End If
        
        If IIf(IsNull(rs!tipo), "", rs!tipo) = "GC" And (IIf(IsNull(rs!chrTipo), "", rs!chrTipo) = "AR" Or IIf(IsNull(rs!chrTipo), "", rs!chrTipo) = "ME") Then
            vlstrSentencia = "SELECT MAX(CASE WHEN " & IIf(IsNull(rs!bitcostobase), 0, rs!bitcostobase) & " = 0 THEN IVARTICULOEMPRESAS.MNYCOSTOMASALTO ELSE IVARTICULOEMPRESAS.MNYCOSTOULTENTRADA END / (CASE WHEN " & IIf(IsNull(rs!INTDESCUENTOINVENTARIO), 0, rs!INTDESCUENTOINVENTARIO) & " = 1 THEN ARTI.INTCONTENIDO ELSE 1 END)) MNYCOSTO " & _
                " FROM PVGRUPOCARGO MAST" & _
                    " INNER JOIN PVDETALLEGRUPOCARGO DET ON MAST.INTCVEGRUPO = DET.INTCVEGRUPO" & _
                    " INNER JOIN IVARTICULO ARTI ON ARTI.INTIDARTICULO = DET.INTCVECARGO" & _
                    " LEFT JOIN IVARTICULOEMPRESAS ON ARTI.CHRCVEARTICULO = IVARTICULOEMPRESAS.CHRCVEARTICULO AND IVARTICULOEMPRESAS.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable & _
                " WHERE CHRTIPO IN ('AR','ME') AND MAST.INTCVEGRUPO = " & IIf(IsNull(rs!intCveCargo), 0, rs!intCveCargo)
            Set rsCostoGrupo = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            
            grdPaquetes.TextMatrix(intCtrl, intColCosto) = FormatCurrency(IIf(IsNull(rsCostoGrupo!mnycosto), 0, rsCostoGrupo!mnycosto), 2)
            grdPaquetes.TextMatrix(intCtrl, intColMargenUtilidad) = FormatPercent(Round(IIf(IsNull(rs!mnyPrecio), 0, rs!mnyPrecio) / IIf(IsNull(rsCostoGrupo!mnycosto), 0, Round(rsCostoGrupo!mnycosto, 2)), 4) - 1)
        End If
        
        'vldblImporte = vldblImporte + CDbl(FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(intCtrl, intColPrecioUnitario), "$", "")) * CInt(Replace(grdPaquetes.TextMatrix(intCtrl, 4), "$", "")), 2))
        vldblImporte = vldblImporte + CDbl(Replace(grdPaquetes.TextMatrix(intCtrl, intColImporte), "$", ""))
        
        'agregue el GC - tipo
        grdPaquetes.TextMatrix(intCtrl, 14) = IIf(IsNull(rs!chrTipo), "", rs!chrTipo)
        
        'agregue el OC - Precio
        If rs!tipo <> "AR" Then
            grdPaquetes.TextMatrix(intCtrl, 15) = FormatCurrency(IIf(IsNull(rs!PrecioOC), "0", rs!PrecioOC), 2)
            If grdPaquetes.TextMatrix(intCtrl, 15) = "0" Then
                blnPrecioCero = True
            End If
        End If
        
        intCtrl = intCtrl + 1
        rs.MoveNext
    Loop
    grdPaquetes.Redraw = True
    
    TxtImporte.Text = FormatCurrency(vldblImporte, 2)
    TxtIva.Text = FormatCurrency((vldblIVA / 100) * vldblImporte, 2)
    TxtTotal.Text = Format(Val(Format(TxtImporte.Text, "")) + Val(Format(TxtIva.Text, "")), "$ ###,###,###,##0.00")
    
    If vldblImporte > 0 Then
        TxtCalcular.Enabled = True
        CmdCalcular.Enabled = True
        Label2.Enabled = True
        TxtMargenArt.Enabled = True
        Label3.Enabled = True
    Else
        TxtCalcular.Enabled = False
        CmdCalcular.Enabled = False
        Label2.Enabled = False
        TxtMargenArt.Enabled = False
        Label3.Enabled = False
    End If
    
    TxtCalcular.Text = FormatCurrency(0, 2)
    TxtMargenArt.Text = "0.00%"
    
    
    vlstrSentencia = "SELECT SUM(PVDETALLELISTA.MNYPRECIO) TOTAL FROM PVDETALLELISTA " & _
                                "WHERE PVDETALLELISTA.CHRCVECARGO = '" & cboPaquetes.ItemData(cboPaquetes.ListIndex) & "' AND PVDETALLELISTA.CHRTIPOCARGO = 'PA'"
    Set rsPrecioLista = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsPrecioLista.RecordCount > 0 Then
        If rsPrecioLista!Total = 0 And Val(Format(TxtTotal.Text, "")) > 0 Then
            cmdGrabarRegistro.Enabled = True
        Else
            cmdGrabarRegistro.Enabled = False
        End If
    Else
        cmdGrabarRegistro.Enabled = False
    End If
    cmdDelete.Enabled = IIf(Val(Format(TxtTotal.Text, "")) > 0, True, False)
    
End Sub

Private Sub grdPaquetes_Click()

    If grdPaquetes.Rows > 2 And Trim(cboPaquetes.Text) <> "" Then
        If (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "AR" And grdPaquetes.Col = intColMargenUtilidad) _
            Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And grdPaquetes.Col = intColMargenUtilidad And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0) Then
            lblEtiquetaF10.Visible = True
        Else
            lblEtiquetaF10.Visible = False
        End If
    End If

    If Trim(grdPaquetes.TextMatrix(1, 1)) <> "" Then
        IntNumCargo = 0
        strTipo = ""
        dblgrdprecio = 0
        
        IntNumCargo = grdPaquetes.TextMatrix(grdPaquetes.Row, 1)
        strTipo = Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3))
        dblgrdprecio = Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario))
        blnIvaPrecioProveedor = IIf(Val(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColIVAPrecioProv), "$", "")) <> 0, 1, 0)
        intGrdPaqueteR = grdPaquetes.Row
        intGrdPaqueteC = grdPaquetes.Col
        dimIntGrdTOp = grdPaquetes.TopRow
        
        If strTipo <> "AR" Then
            pConfiguraGridSub
            pLlenaGridsubrogados IntNumCargo, strTipo
        Else
            grdSubrogados.Clear
            pConfiguraGridSub
        End If
    End If
End Sub

Private Sub pLlenaGridsubrogados(ByVal intnum As Long, StrtipoX As String)
    Dim rsX As ADODB.Recordset
    Dim vlstrSentenciaX As String
    Dim intCtrlX As Integer
                       
    vlstrSentenciaX = "Select coserviciosubrogado.intcveproveedor, " & _
                       "coproveedor.VCHNOMBRECOMERCIAL Nomproveedor, " & _
                       "COServicioSubrogado.MNYCANTIDAD Cantidad, " & _
                       "CASE COServicioSubrogado.INTTIPOACUERDO " & _
                       "WHEN 0 THEN '$' " & _
                       "WHEN 1 THEN '%' " & _
                       "END TIPOACUERDO, " & _
                       "COSERVICIOSUBROGADOPAQUETE.MNYPRECIOESPECIFICO," & _
                       " COServicioSubrogado.INTCVETIPOSERVICIO," & _
                       " coserviciosubrogado.CHRTIPOSERVICIO " & _
                       "from COServicioSubrogado " & _
                       "LEFT join coproveedor on coproveedor.INTCVEPROVEEDOR = COServicioSubrogado.INTCVEPROVEEDOR " & _
                       "left join COSERVICIOSUBROGADOPAQUETE on COSERVICIOSUBROGADOPAQUETE.INTCVEPROVEEDOR = COServicioSubrogado.INTCVEPROVEEDOR " & _
                            "AND COSERVICIOSUBROGADOPAQUETE.CHRTIPOSERVICIO = coserviciosubrogado.CHRTIPOSERVICIO " & _
                            "AND COSERVICIOSUBROGADOPAQUETE.INTCVECARGO = coserviciosubrogado.INTCVETIPOSERVICIO " & _
                            "And COSERVICIOSUBROGADOPAQUETE.INTNUMPAQUETE  = " & CInt(Trim(cboPaquetes.ItemData(cboPaquetes.ListIndex))) & _
                       " where COServicioSubrogado.INTCVETIPOSERVICIO = " & intnum & _
                       " and coserviciosubrogado.CHRTIPOSERVICIO = '" & StrtipoX & "'" & _
                       " order by coproveedor.VCHNOMBRECOMERCIAL "
    Set rsX = frsRegresaRs(vlstrSentenciaX, adLockOptimistic, adOpenDynamic)
    
    intCtrl = 1
    grdSubrogados.TextMatrix(intCtrl, 1) = ""
    grdSubrogados.TextMatrix(intCtrl, 2) = ""
    grdSubrogados.TextMatrix(intCtrl, 3) = ""
   
    Do Until rsX.EOF
        If intCtrl > 1 Then
            grdSubrogados.AddItem ""
        End If
        
        grdSubrogados.TextMatrix(intCtrl, 1) = IIf(IsNull(rsX!Nomproveedor), "", rsX!Nomproveedor)
        grdSubrogados.TextMatrix(intCtrl, 2) = IIf(IsNull(rsX!cantidad), "", rsX!cantidad)
        grdSubrogados.TextMatrix(intCtrl, 3) = IIf(IsNull(rsX!TIPOACUERDO), "", rsX!TIPOACUERDO)
        grdSubrogados.TextMatrix(intCtrl, 4) = IIf(rsX!MNYPRECIOESPECIFICO = 0 Or IsNull(rsX!MNYPRECIOESPECIFICO), " ", IIf(rsX!TIPOACUERDO = "%", FormatNumber(rsX!MNYPRECIOESPECIFICO, 2), FormatCurrency(rsX!MNYPRECIOESPECIFICO, 2)))
        grdSubrogados.TextMatrix(intCtrl, 5) = IIf(IsNull(rsX!INTCVEPROVEEDOR), "", rsX!INTCVEPROVEEDOR)
        grdSubrogados.TextMatrix(intCtrl, 6) = IIf(IsNull(rsX!INTCVETIPOSERVICIO), "", rsX!INTCVETIPOSERVICIO)
        grdSubrogados.TextMatrix(intCtrl, 7) = IIf(IsNull(rsX!CHRTIPOSERVICIO), "", rsX!CHRTIPOSERVICIO)
        
        intCtrl = intCtrl + 1
        rsX.MoveNext
    Loop
End Sub

Private Sub grdPaquetes_GotFocus()
    If grdPaquetes.Rows > 2 And Trim(cboPaquetes.Text) <> "" Then
        If (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "AR" And grdPaquetes.Col = intColMargenUtilidad) _
            Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And grdPaquetes.Col = intColMargenUtilidad And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0) Then
            lblEtiquetaF10.Visible = True
        Else
            lblEtiquetaF10.Visible = False
        End If
    End If
End Sub

Private Sub grdPaquetes_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlstrMargen As String
    Dim vllngContador As Long
    
    blnGrdPaquete = True

    If grdPaquetes.Rows > 2 And Trim(cboPaquetes.Text) <> "" Then
        If (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "AR" And grdPaquetes.Col = intColMargenUtilidad) _
            Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And grdPaquetes.Col = intColMargenUtilidad And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0) Then
                lblEtiquetaF10.Visible = True
        Else
            lblEtiquetaF10.Visible = False
        End If
    End If

    If grdPaquetes.Col >= 0 Then  ' vgintColumnaCurrency Then
        If KeyCode = vbKeyReturn Then 'para que se edite el contenido de la celda como en excel
            If grdPaquetes.TextMatrix(grdPaquetes.Row, 1) <> "" Then
                If grdPaquetes.Col = vgintColumnaCurrency Then 'Columna que puede ser editada
                    If vlblnPermitirModificar Then
                        pEditarColumna KeyCode, txtPrecio, grdPaquetes
                    End If
                End If
                
                If grdPaquetes.TextMatrix(grdPaquetes.Row, 3) = "AR" Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0) Then
                    If grdPaquetes.Col = intColMargenUtilidad Then 'Columna que puede ser editada
                        If vlblnPermitirModificar Then
                            pEditarColumna KeyCode, txtMargen, grdPaquetes
                        End If
                    End If
                End If
            End If
        Else
            If grdPaquetes.TextMatrix(grdPaquetes.Row, 1) <> "" Then
                Select Case KeyCode
                Case 27   'ESC
                Case 38   'Flecha arriba
                    IntNumCargo = grdPaquetes.TextMatrix(grdPaquetes.Row, 1)
                    strTipo = Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3))
                    dblgrdprecio = Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario))
                    
                    intGrdPaqueteR = grdPaquetes.Row
                    blnIvaPrecioProveedor = IIf(Val(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColIVAPrecioProv), "$", "")) <> 0, 1, 0)
                    
                    If strTipo <> "AR" Then
                        pConfiguraGridSub
                        pLlenaGridsubrogados IntNumCargo, strTipo
                    Else
                        grdSubrogados.Clear
                        pConfiguraGridSub
                    End If
                
                Case 40 ' Flecha abajo
                    IntNumCargo = grdPaquetes.TextMatrix(grdPaquetes.Row, 1)
                    strTipo = Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3))
                    dblgrdprecio = Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario))
                            
                    intGrdPaqueteR = grdPaquetes.Row
                    
                    blnIvaPrecioProveedor = IIf(Val(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColIVAPrecioProv), "$", "")) <> 0, 1, 0)
                    
                    If strTipo <> "AR" Then
                        pConfiguraGridSub
                        pLlenaGridsubrogados IntNumCargo, strTipo
                    Else
                        grdSubrogados.Clear
                        pConfiguraGridSub
                    End If
                End Select
            End If
        End If
    Else
        If KeyCode = vbKeyReturn Then
            grdPaquetes.Col = 0
            grdPaquetes.CellFontBold = True
            grdPaquetes.Col = 1
            If grdPaquetes.Row - 1 < grdPaquetes.Rows Then
                If grdPaquetes.Row = grdPaquetes.Rows - 1 Then
                    grdPaquetes.Row = 1
                Else
                    grdPaquetes.Row = grdPaquetes.Row + 1
                    If grdPaquetes.Row = grdPaquetes.Rows - 1 Then
                        grdPaquetes.Row = 1
                    Else
                        grdPaquetes.Row = grdPaquetes.Row + 1
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = vbKeyF10 Then
        If (grdPaquetes.Col = intColMargenUtilidad) And (grdPaquetes.TextMatrix(grdPaquetes.Row, 3) = "AR" Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0)) Then
        
            '¿Desea aplicar este margen a todos los artículos?
            If MsgBox(SIHOMsg(1671), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                vlstrMargen = grdPaquetes.TextMatrix(grdPaquetes.Row, intColMargenUtilidad)
            
                For vllngContador = 1 To grdPaquetes.Rows - 1
                    If grdPaquetes.TextMatrix(vllngContador, 3) = "AR" Or (Trim(grdPaquetes.TextMatrix(vllngContador, 3)) = "GC" And CDbl(grdPaquetes.TextMatrix(vllngContador, intColCosto)) > 0) Then
                        grdPaquetes.TextMatrix(vllngContador, intColMargenUtilidad) = vlstrMargen
                        If grdPaquetes.TextMatrix(vllngContador, intColMargenUtilidad) = "" Then
                            grdPaquetes.TextMatrix(vllngContador, intColPrecioUnitario) = FormatCurrency(CDbl(grdPaquetes.TextMatrix(vllngContador, intColCosto)))
                        Else
                            grdPaquetes.TextMatrix(vllngContador, intColPrecioUnitario) = FormatCurrency(CDbl(grdPaquetes.TextMatrix(vllngContador, intColCosto)) + (CDbl(grdPaquetes.TextMatrix(vllngContador, intColCosto)) * (CDbl(Replace(grdPaquetes.TextMatrix(vllngContador, intColMargenUtilidad), "%", "")) / 100)))
                        End If
                        
                        grdPaquetes.TextMatrix(vllngContador, intColImporte) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(vllngContador, intColPrecioUnitario), "$", "")) * CInt(Replace(grdPaquetes.TextMatrix(vllngContador, 4), "$", "")), 2)
                        grdPaquetes.TextMatrix(vllngContador, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(vllngContador, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(vllngContador, intColImporte), "$", "")), 2)
                        grdPaquetes.TextMatrix(vllngContador, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(vllngContador, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(vllngContador, intColIVANormal), "$", "")), 2)
                    
                        pReCalculaTotales
                        
                        grdPaquetes.TextMatrix(vllngContador, intColObtenerCargo) = 1
                    End If
                Next
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPaquetes_KeyDown"))
    Unload Me
End Sub

Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError
    
    Dim vlintTexto As Integer
    
    '-------------------------
    'Que se salga cuando si no tiene permiso
    '-------------------------
    If Not fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 3068, 3071), "E") Then Exit Sub
    '-------------------------
    'Que se salga cuando ya esta facturado
    '-------------------------
   
    With txtEdit
        If Val(Format(grid.Text, "############.##")) = 0 Then
            .Text = Replace(grid, "$", "") 'Inicialización del Textbox
        Else
            If blnGrdPaquete Then
                .Text = Replace(grdPaquetes.TextMatrix(grid.Row, grid.Col), "%", "") 'Inicialización del Textbox
            Else
                .Text = Replace(grdSubrogados.TextMatrix(grid.Row, grid.Col), "%", "") 'Inicialización del Textbox
                blnGrdPaquete = True
            End If
        End If
       
        Select Case KeyAscii
            Case 0 To 32
                'Edita el texto de la celda en la que está posicionado
                .SelStart = 0
                .SelLength = 1000
            Case 8, 48 To 57
                ' Reemplaza el texto actual solo si se teclean números
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
            Case 46
                ' Reemplaza el texto actual solo si se teclean números
                .Text = "."
                .SelStart = 1
        End Select
    End With
    
    With grid
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
    
    txtEdit.Visible = True
    txtEdit.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEditarColumna"))
    Unload Me
End Sub

Private Sub grdPaquetes_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If grdPaquetes.Rows > 2 And Trim(cboPaquetes.Text) <> "" Then
        If (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "AR" And grdPaquetes.Col = intColMargenUtilidad) _
            Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And grdPaquetes.Col = intColMargenUtilidad And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0) Then
                lblEtiquetaF10.Visible = True
        Else
            lblEtiquetaF10.Visible = False
        End If
    End If
    
    If grdPaquetes.Col = vgintColumnaCurrency Then 'Columna que puede ser editad
        If grdPaquetes.TextMatrix(grdPaquetes.Row, 1) <> "" Then
            If vlblnPermitirModificar Then
                pEditarColumna KeyAscii, txtPrecio, grdPaquetes
            End If
        End If
    End If
    
    If grdPaquetes.Col = intColMargenUtilidad Then 'Columna que puede ser editad
        If grdPaquetes.TextMatrix(grdPaquetes.Row, 1) <> "" Then
            If grdPaquetes.TextMatrix(grdPaquetes.Row, 3) = "AR" Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0) Then
                If vlblnPermitirModificar Then
                    pEditarColumna KeyAscii, txtMargen, grdPaquetes
                End If
            End If
        End If
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdpaquetes_KeyPress"))
    Unload Me
End Sub

Private Sub grdPaquetes_LostFocus()
    lblEtiquetaF10.Visible = False
End Sub

Private Sub grdPaquetes_Scroll()
    txtPrecio.Visible = False
    txtMargen.Visible = False
End Sub

Private Sub grdSubrogados_Click()
    If Trim(grdSubrogados.TextMatrix(1, 1)) <> "" Then
        lngCveProveedor = 0
        dblTxtPrecio1 = 0
        IntNumCargo = 0
        
        lngCveProveedor = grdSubrogados.TextMatrix(grdSubrogados.Row, 5)
        IntNumCargo = grdSubrogados.TextMatrix(grdSubrogados.Row, 6)
        
        intGrdPaqueteRPro = grdSubrogados.Row
        intGrdPaqueteCPro = grdSubrogados.Col
    End If
End Sub

Private Sub grdSubrogados_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    blnGrdPaquete = False
    If grdSubrogados.Col = vgintColumnaCurrency1 Then
        If grdSubrogados.TextMatrix(grdSubrogados.Row, 1) <> "" Then
            If KeyCode = vbKeyReturn Then 'para que se edite el contenido de la celda como en excel
                pEditarColumna 13, TxtPrecio1, grdSubrogados
            End If
        
            Select Case KeyCode
            Case 38   'Flecha arriba
                intGrdPaqueteRPro = grdSubrogados.Row
            Case 40 ' Flecha abajo
                intGrdPaqueteRPro = grdSubrogados.Row
            End Select
        End If
    Else
        If KeyCode = vbKeyReturn Then
            grdSubrogados.Col = 0
            grdSubrogados.CellFontBold = True
            grdSubrogados.Col = 1
            If grdSubrogados.Row - 1 < grdSubrogados.Rows Then
                If grdSubrogados.Row = grdSubrogados.Rows - 1 Then
                    grdSubrogados.Row = 1
                Else
                    grdSubrogados.Row = grdSubrogados.Row + 1
                    If grdSubrogados.Row = grdSubrogados.Rows - 1 Then
                        grdSubrogados.Row = 1
                    Else
                        grdSubrogados.Row = grdSubrogados.Row + 1
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPaquetes_KeyDown"))
    Unload Me
End Sub

Private Sub grdSubrogados_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If grdSubrogados.Col = vgintColumnaCurrency1 Then
        If grdSubrogados.TextMatrix(grdSubrogados.Row, 1) <> "" Then
            pEditarColumna KeyAscii, TxtPrecio1, grdSubrogados
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdsubrogados_KeyPress"))
    Unload Me
End Sub

Private Sub grdSubrogados_Scroll()
    TxtPrecio1.Visible = False
End Sub

Private Sub TxtCalcular_GotFocus()
    Call pEnfocaTextBox(TxtCalcular)
End Sub

Private Sub TxtCalcular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TxtCalcular.Text = FormatCurrency(TxtCalcular.Text, 2)
        TxtMargenArt.SetFocus
    End If
End Sub

Private Sub TxtCalcular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And InStr(TxtCalcular, ".") > 0 Then
        KeyAscii = 0
        Exit Sub
    'ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 Then
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCalcular_LostFocus()
     TxtCalcular.Text = FormatCurrency(TxtCalcular.Text, 2)
End Sub

Private Sub TxtMargenArt_GotFocus()
    Call pEnfocaTextBox(TxtMargenArt)
End Sub

Private Sub TxtMargenArt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TxtMargenArt.Text = Replace(TxtMargenArt.Text, "%", "") + "%"
        CmdCalcular.SetFocus
    End If
End Sub

Private Sub TxtMargenArt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And InStr(TxtMargenArt, ".") > 0 Then
        KeyAscii = 0
        Exit Sub
    'ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 Then
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    'Para verificar que tecla fue presionada en el textbox
    With grdPaquetes
        Select Case KeyCode
            Case 27   'ESC
                txtPrecio.Visible = False
                .SetFocus
            Case 37   'Flecha Izq
                .SetFocus
                DoEvents
                .Col = .Col - 1
                txtPrecio.Visible = False
                .SetFocus
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    .Row = .Row - 1
                End If
                
                txtPrecio.Visible = False
                .SetFocus
            Case 39   'Flecha para der
                .SetFocus
                DoEvents
                .Col = .Col + 1
                txtPrecio.Visible = False
                .SetFocus
            Case 13 'enter
                If Trim(txtPrecio.Text) <> .TextMatrix(.Row, .Col) Then
                    grdPaquetes.TextMatrix(grdPaquetes.Row, grdPaquetes.Col) = FormatCurrency(txtPrecio.Text, 2)
                    cmdGrabarRegistro.Enabled = True
                    
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario), "$", "")) * CInt(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, 4), "$", "")), 2)
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte), "$", "")), 2)
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColIVANormal), "$", "")), 2)
                
                    If txtPrecio.Text >= 0 Then
                        pReCalculaTotales
                    End If
                    
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColObtenerCargo) = 1

                    If (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "AR") _
                        Or (Trim(grdPaquetes.TextMatrix(grdPaquetes.Row, 3)) = "GC" And CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) > 0) Then
                            grdPaquetes.TextMatrix(grdPaquetes.Row, intColMargenUtilidad) = FormatPercent((CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario)) / CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto))) - 1)
                    Else
                        grdPaquetes.TextMatrix(grdPaquetes.Row, intColMargenUtilidad) = ""
                    End If

                    If .Row < .Rows - 1 Then
                        .Row = .Row + 1
                    End If
                
                    txtPrecio.Visible = False
                    
                    .Col = vgintColumnaCurrency
                    
                    .SetFocus
                    
                    blnTxtprecio = True
                Else
                    txtPrecio.Visible = False
                    .SetFocus
                End If
            Case 40 ' flecha abajo
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
                
                txtPrecio.Visible = False
                .SetFocus
        End Select
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPrecio_KeyDown"))
    Unload Me
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    ' Solo permite números
    If Not fblnFormatoCantidad(txtPrecio, KeyAscii, 6) Then KeyAscii = 7

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPrecio_KeyPress"))
    Unload Me
End Sub

Private Sub txtMargen_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    ' Solo permite números
    If Not fblnFormatoCantidad(txtMargen, KeyAscii, 6) Then KeyAscii = 7

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMargen_KeyPress"))
    Unload Me
End Sub

Private Sub txtPrecio_LostFocus()
    txtPrecio.Visible = False
End Sub

Private Sub txtMargen_LostFocus()
    txtMargen.Visible = False
End Sub

Private Sub TxtPrecio1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    'Para verificar que tecla fue presionada en el textbox
    With grdSubrogados
        Select Case KeyCode
            Case 27   'ESC
                txtPrecio.Visible = False
                .SetFocus
            Case 37   'Flecha Izq
                .SetFocus
                DoEvents
                .Col = .Col - 1
                TxtPrecio1.Visible = False
                .SetFocus
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    .Row = .Row - 1
                End If
                TxtPrecio1.Visible = False
                .SetFocus
            Case 39   'Flecha para der
                .SetFocus
                DoEvents
                TxtPrecio1.Visible = False
                .SetFocus
            Case 13 'enter
                If Trim(TxtPrecio1.Text) <> "" Then
                    If Trim(TxtPrecio1.Text) <> .TextMatrix(.Row, .Col) Then
                        If grdSubrogados.TextMatrix(grdSubrogados.Row, 3) = "%" Then
                            If CLng(TxtPrecio1.Text) <= 100 Then
                                grdSubrogados.TextMatrix(grdSubrogados.Row, grdSubrogados.Col) = IIf(grdSubrogados.TextMatrix(grdSubrogados.Row, 3) = "%", FormatNumber(TxtPrecio1.Text, 2), FormatCurrency(TxtPrecio1.Text, 2))
                            Else
                                'La cantidad es incorrecta.
                                MsgBox SIHOMsg(452) & " Debe ser menor o igual a 100%" & Chr(13) & "Proveedor: " & grdSubrogados.TextMatrix(grdSubrogados.Row, 1), vbOKOnly + vbExclamation, "Mensaje"
                                grdSubrogados.Row = intGrdPaqueteRPro
                                grdSubrogados.Col = vgintColumnaCurrency1
                                grdSubrogados.SetFocus
                                Exit Sub
                            End If
                        Else
                            grdSubrogados.TextMatrix(grdSubrogados.Row, grdSubrogados.Col) = IIf(grdSubrogados.TextMatrix(grdSubrogados.Row, 3) = "%", FormatNumber(TxtPrecio1.Text, 2), FormatCurrency(TxtPrecio1.Text, 2))
                        End If
                    
                        cmdGrabarRegistro.Enabled = True
                        txtPrecio.Visible = False
                        .SetFocus
                        
                        If .Row < .Rows - 1 Then
                            .Row = .Row + 1
                        End If
                        
                        blnTxtprecio = False
                    
                        dblTxtPrecio1 = grdSubrogados.TextMatrix(grdSubrogados.Row, 4)
                        grdSubrogados.TextMatrix(grdSubrogados.Row, 8) = 1
                    Else
                        TxtPrecio1.Visible = False
                        .SetFocus
                    End If
                Else
                    .Row = intGrdPaqueteRPro
                    .Col = 4
                    
                    TxtPrecio1.Visible = False
                End If
            Case 40 ' flecha abajo
                .Row = intGrdPaqueteRPro
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
                
                TxtPrecio1.Visible = False
                .SetFocus
        End Select
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPrecio_KeyDown"))
    Unload Me
End Sub

Private Sub TxtPrecio1_KeyPress(KeyAscii As Integer)
     On Error GoTo NotificaError
    ' Solo permite números
    If Not fblnFormatoCantidad(TxtPrecio1, KeyAscii, 6) Then KeyAscii = 7

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPrecio_KeyPress"))
    Unload Me
End Sub

Private Sub TxtPrecio1_LostFocus()
    TxtPrecio1.Visible = False
End Sub

Private Sub pReCalculaTotales()
Dim vlintContador As Integer

    vldblImporte = 0
    
    TxtImporte.Text = ""
    TxtIva.Text = ""
    TxtTotal.Text = ""
        
    With grdPaquetes
        For vlintContador = 1 To .Rows - 1
            vldblImporte = vldblImporte + CDbl(FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(vlintContador, intColPrecioUnitario), "$", "")) * CInt(Replace(grdPaquetes.TextMatrix(vlintContador, 4), "$", "")), 2))
        Next
    End With
            
    TxtImporte.Text = FormatCurrency(vldblImporte, 2)
    TxtIva.Text = FormatCurrency((vldblIVA / 100) * vldblImporte, 2)
    TxtTotal.Text = FormatCurrency(Val(Format(TxtImporte.Text, "")) + Val(Format(TxtIva.Text, "")), 2)
End Sub

Private Function fblnPaqueteEnListas(vllngPaquete As Long) As Boolean
    Dim rs As New ADODB.Recordset
    
    fblnPaqueteEnListas = False
    Set rs = frsRegresaRs("select count(*) Cantidad from pvdetallelista where CHRCVECARGO = " & vllngPaquete & " and CHRTIPOCARGO = 'PA'")
    If rs!cantidad = 0 Or IsNull(rs!cantidad) Then
        fblnPaqueteEnListas = False
    Else
        fblnPaqueteEnListas = True
    End If

    Exit Function
End Function

Private Sub txtMargen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    'Para verificar que tecla fue presionada en el textbox
    With grdPaquetes
        Select Case KeyCode
            Case 27   'ESC
                txtMargen.Visible = False
                .SetFocus
            Case 37   'Flecha Izq
                .SetFocus
                DoEvents
                .Col = .Col - 1
                txtMargen.Visible = False
                .SetFocus
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    .Row = .Row - 1
                End If
                
                txtMargen.Visible = False
                .SetFocus
            Case 39   'Flecha para der
                .SetFocus
                DoEvents
                .Col = .Col + 1
                txtMargen.Visible = False
                .SetFocus
            Case 13 'enter
                If Trim(txtMargen.Text) <> .TextMatrix(.Row, .Col) Then
                                
                    grdPaquetes.TextMatrix(grdPaquetes.Row, grdPaquetes.Col) = ""
                    If Trim(txtMargen.Text) <> "" Then
                        If CDbl(txtMargen.Text) <> 0 Then
                            grdPaquetes.TextMatrix(grdPaquetes.Row, grdPaquetes.Col) = FormatPercent(txtMargen.Text / 100, 2)
                        End If
                    End If
                    
                    cmdGrabarRegistro.Enabled = True
                    
                    If grdPaquetes.TextMatrix(grdPaquetes.Row, grdPaquetes.Col) = "" Then
                        grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario) = FormatCurrency(CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)))
                    Else
                        grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario) = FormatCurrency(CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) + (CDbl(grdPaquetes.TextMatrix(grdPaquetes.Row, intColCosto)) * (CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColMargenUtilidad), "%", "")) / 100)))
                    End If
                    
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColPrecioUnitario), "$", "")) * CInt(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, 4), "$", "")), 2)
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColIVANormal) = FormatCurrency(IIf(CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte), "$", "")) = 0, 0, (vldblIVA / 100)) * CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte), "$", "")), 2)
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColTotal) = FormatCurrency(CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColImporte), "$", "")) + CDbl(Replace(grdPaquetes.TextMatrix(grdPaquetes.Row, intColIVANormal), "$", "")), 2)
                
                    pReCalculaTotales
                    
                    grdPaquetes.TextMatrix(grdPaquetes.Row, intColObtenerCargo) = 1

                    If .Row < .Rows - 1 Then
                        .Row = .Row + 1
                    End If
                
                    txtMargen.Visible = False
                    
                    .Col = intColMargenUtilidad
                    
                    .SetFocus
                Else
                    txtMargen.Visible = False
                    .SetFocus
                End If
            Case 40 ' flecha abajo
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
                
                txtMargen.Visible = False
                .SetFocus
        End Select
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMargen_KeyDown"))
    Unload Me
End Sub
