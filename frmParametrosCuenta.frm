VERSION 5.00
Begin VB.Form frmParametrosCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de parámetros de la cuenta"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNotasDeCredito 
      Caption         =   "Descuentos por notas de crédito"
      Height          =   2565
      Left            =   120
      TabIndex        =   51
      Top             =   2280
      Width           =   8445
      Begin VB.TextBox txtPorcentajeExcedentePorNota 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "100.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   4665
         MaxLength       =   10
         TabIndex        =   14
         ToolTipText     =   "Descuento para las notas de crédito automáticas"
         Top             =   300
         Width           =   1100
      End
      Begin VB.TextBox txtCantidadLimiteExcedente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7185
         MaxLength       =   10
         TabIndex        =   15
         ToolTipText     =   "Cantidad límite"
         Top             =   300
         Width           =   1100
      End
      Begin VB.TextBox txtPorcentajeCopagoPorNota 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4665
         MaxLength       =   10
         TabIndex        =   34
         ToolTipText     =   "Descuento para las notas de crédito automáticas"
         Top             =   2100
         Width           =   1095
      End
      Begin VB.TextBox txtPorcentajeCoasAdicionalPorNota 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4665
         MaxLength       =   10
         TabIndex        =   30
         ToolTipText     =   "Descuento para las notas de crédito automáticas"
         Top             =   1740
         Width           =   1100
      End
      Begin VB.TextBox txtPorcentajeCoaseguroPorNota 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4665
         MaxLength       =   10
         TabIndex        =   22
         ToolTipText     =   "Descuento para las notas de crédito automáticas"
         Top             =   1020
         Width           =   1100
      End
      Begin VB.TextBox txtPorcentajeDeduciblePorNota 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4665
         MaxLength       =   10
         TabIndex        =   18
         ToolTipText     =   "Descuento para las notas de crédito automáticas"
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox txtPorcentajeCoaseguroMPorNota 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4665
         MaxLength       =   10
         TabIndex        =   26
         ToolTipText     =   "Descuento para las notas de crédito automáticas"
         Top             =   1380
         Width           =   1100
      End
      Begin VB.TextBox txtCantidadLimiteCopago 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7185
         MaxLength       =   10
         TabIndex        =   35
         ToolTipText     =   "Cantidad límite"
         Top             =   2100
         Width           =   1100
      End
      Begin VB.TextBox txtCantidadLimiteCoasAdicional 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7185
         MaxLength       =   10
         TabIndex        =   31
         ToolTipText     =   "Cantidad límite"
         Top             =   1740
         Width           =   1100
      End
      Begin VB.TextBox txtCantidadLimiteCoaseguro 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7185
         MaxLength       =   10
         TabIndex        =   23
         ToolTipText     =   "Cantidad límite"
         Top             =   1020
         Width           =   1100
      End
      Begin VB.TextBox txtCantidadLimiteDeducible 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7185
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Cantidad límite"
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox txtCantidadLimiteCoaseguroM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7185
         MaxLength       =   10
         TabIndex        =   27
         ToolTipText     =   "Cantidad límite"
         Top             =   1380
         Width           =   1100
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1560
         TabIndex        =   57
         Top             =   240
         Width           =   2895
         Begin VB.OptionButton optTipoDesctoExcedente 
            Caption         =   "Por cantidad"
            Height          =   220
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            ToolTipText     =   "Por cantidad"
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optTipoDesctoExcedente 
            Caption         =   "Por porcentaje"
            Height          =   220
            Index           =   0
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Por porcentaje"
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1560
         TabIndex        =   56
         Top             =   600
         Width           =   2895
         Begin VB.OptionButton optTipoDesctoDeducible 
            Caption         =   "Por cantidad"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   17
            ToolTipText     =   "Por cantidad"
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optTipoDesctoDeducible 
            Caption         =   "Por porcentaje"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Por porcentaje"
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1560
         TabIndex        =   55
         Top             =   960
         Width           =   2895
         Begin VB.OptionButton optTipoDesctoCoaseguro 
            Caption         =   "Por cantidad"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   21
            ToolTipText     =   "Por cantidad"
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton optTipoDesctoCoaseguro 
            Caption         =   "Por porcentaje"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Por porcentaje"
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1560
         TabIndex        =   54
         Top             =   1320
         Width           =   2895
         Begin VB.OptionButton optTipoDesctoCoaseguroMedico 
            Caption         =   "Por cantidad"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   25
            ToolTipText     =   "Por cantidad"
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optTipoDesctoCoaseguroMedico 
            Caption         =   "Por porcentaje"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Por porcentaje"
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1560
         TabIndex        =   53
         Top             =   1680
         Width           =   2895
         Begin VB.OptionButton optTipoDesctoCoaseguroAdicional 
            Caption         =   "Por cantidad"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   29
            ToolTipText     =   "Por cantidad"
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton optTipoDesctoCoaseguroAdicional 
            Caption         =   "Por porcentaje"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            ToolTipText     =   "Por porcentaje"
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame15 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1560
         TabIndex        =   52
         Top             =   2040
         Width           =   2895
         Begin VB.OptionButton optTipoDesctoCopago 
            Caption         =   "Por cantidad"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   33
            ToolTipText     =   "Por cantidad"
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton optTipoDesctoCopago 
            Caption         =   "Por porcentaje"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "Por porcentaje"
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Label lbPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   0
         Left            =   5820
         TabIndex        =   75
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lbPorcentajeDescuentoExcedente 
         AutoSize        =   -1  'True
         Caption         =   "Excedente"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lbCantidadLimiteExcedente 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad límite"
         Height          =   195
         Left            =   6045
         TabIndex        =   73
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lbPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   4
         Left            =   5820
         TabIndex        =   72
         Top             =   1800
         Width           =   120
      End
      Begin VB.Label lbPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   3
         Left            =   5820
         TabIndex        =   71
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label lbPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   2
         Left            =   5820
         TabIndex        =   70
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label lbPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   1
         Left            =   5820
         TabIndex        =   69
         Top             =   720
         Width           =   120
      End
      Begin VB.Label lbPorcentajeDescuentoCoPago 
         AutoSize        =   -1  'True
         Caption         =   "Copago"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label lbPorcentajeDescuentoCoaAdicional 
         AutoSize        =   -1  'True
         Caption         =   "Coaseguro adicional"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label lbPorcentajeDescuentoCoaseguro 
         AutoSize        =   -1  'True
         Caption         =   "Coaseguro"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lbPorcentajeDescuentoDeducible 
         AutoSize        =   -1  'True
         Caption         =   "Deducible"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   5
         Left            =   5820
         TabIndex        =   64
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Coaseguro médico"
         Height          =   195
         Left            =   105
         TabIndex        =   63
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label lbCantidadLimiteCoPago 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad límite"
         Height          =   195
         Left            =   6045
         TabIndex        =   62
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label lbCantidadLimiteCoaAdicional 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad límite"
         Height          =   195
         Left            =   6045
         TabIndex        =   61
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label lbCantidadLimiteCoaseguro 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad límite"
         Height          =   195
         Left            =   6045
         TabIndex        =   60
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label lbCantidadLimiteDeducible 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad límite"
         Height          =   195
         Left            =   6045
         TabIndex        =   59
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lbCantidadLimiteCoaMedico 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad límite"
         Height          =   195
         Left            =   6045
         TabIndex        =   58
         Top             =   1440
         Width           =   1050
      End
   End
   Begin VB.Frame fraIncluir 
      Caption         =   "Descuento de conceptos de seguro en CFDI de aseguradora"
      Height          =   1575
      Left            =   120
      TabIndex        =   39
      Top             =   6360
      Width           =   8445
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   8175
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   4815
            Begin VB.OptionButton optTipoDesglose 
               Caption         =   "Presentación simple"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   43
               ToolTipText     =   "Presentación simple"
               Top             =   0
               Width           =   1815
            End
            Begin VB.OptionButton optTipoDesglose 
               Caption         =   "Presentación en cuadrícula"
               Height          =   255
               Index           =   1
               Left            =   1840
               TabIndex        =   44
               ToolTipText     =   "Presentación en cuadrícula"
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.CheckBox chkTotCS 
            Caption         =   "Desglosar descuentos comerciales y de seguros en totales de la representación impresa del CFDI"
            Height          =   225
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Desglosar descuentos comerciales y de seguros en totales de la representación impresa del CFDI"
            Top             =   0
            Width           =   7215
         End
      End
      Begin VB.OptionButton optIncCS 
         Caption         =   "Restar de los importes"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "Restar proporcional del monto pagado por conceptos de seguro a los importes"
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optIncCS 
         Caption         =   "Sumar a los descuentos"
         Height          =   225
         Index           =   1
         Left            =   2760
         TabIndex        =   41
         ToolTipText     =   "Sumar proporcional del monto pagado por conceptos de seguro a los descuentos"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame fraParametros 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   48
      Top             =   4920
      Width           =   7815
      Begin VB.CheckBox chkCalcularCargosSeleccionados 
         Caption         =   "Calcular importes de conceptos de factura para seguros con base en los cargos seleccionados para facturar"
         Height          =   435
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   7740
      End
      Begin VB.CheckBox chkDesglosaIVATasaHospi 
         Caption         =   "Desglosar IVA a la tasa del hospital en conceptos de seguro"
         Height          =   199
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   4575
      End
      Begin VB.CheckBox chkCoaseguroPorFactura 
         Caption         =   "Permitir capturar cantidad de coaseguro por factura"
         Height          =   199
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   4095
      End
   End
   Begin VB.Frame fraDesglosaIVA 
      Caption         =   "Conceptos que desglosan IVA"
      Height          =   2085
      Left            =   120
      TabIndex        =   47
      ToolTipText     =   "Conceptos que desglosan IVA"
      Top             =   120
      Width           =   8445
      Begin VB.CheckBox chkDesglosarCopago 
         Caption         =   "Desglosar importes gravado y no gravado"
         Height          =   190
         Left            =   3500
         TabIndex        =   11
         ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
         Top             =   1800
         Width           =   3400
      End
      Begin VB.CheckBox chkDesglosarCoaseguroAdicional 
         Caption         =   "Desglosar importes gravado y no gravado"
         Height          =   190
         Left            =   3500
         TabIndex        =   9
         ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
         Top             =   1500
         Width           =   3400
      End
      Begin VB.CheckBox chkDesglosarCoaseguro 
         Caption         =   "Desglosar importes gravado y no gravado"
         Height          =   190
         Left            =   3500
         TabIndex        =   5
         ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
         Top             =   900
         Width           =   3400
      End
      Begin VB.CheckBox chkDesglosarDeducible 
         Caption         =   "Desglosar importes gravado y no gravado"
         Height          =   190
         Left            =   3500
         TabIndex        =   3
         ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
         Top             =   600
         Width           =   3400
      End
      Begin VB.CheckBox chkDesglosarCoaseguroM 
         Caption         =   "Desglosar importes gravado y no gravado"
         Height          =   190
         Left            =   3500
         TabIndex        =   7
         ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
         Top             =   1200
         Width           =   3400
      End
      Begin VB.CheckBox chkDesglosarIVACoaseguroM 
         Caption         =   "Coaseguro médico"
         Height          =   190
         Left            =   700
         TabIndex        =   6
         ToolTipText     =   "Desglosar el IVA que corresponde"
         Top             =   1200
         Width           =   2000
      End
      Begin VB.CheckBox chkDesglosarIVADeducible 
         Caption         =   "Deducible"
         Height          =   190
         Left            =   700
         TabIndex        =   2
         ToolTipText     =   "Desglosar el IVA que corresponde"
         Top             =   600
         Width           =   2000
      End
      Begin VB.CheckBox chkDesglosarIVACoaseguro 
         Caption         =   "Coaseguro"
         Height          =   190
         Left            =   700
         TabIndex        =   4
         ToolTipText     =   "Desglosar el IVA que corresponde"
         Top             =   900
         Width           =   2000
      End
      Begin VB.CheckBox chkDesglosarIVACoaseguroAdicional 
         Caption         =   "Coaseguro adicional"
         Height          =   190
         Left            =   700
         TabIndex        =   8
         ToolTipText     =   "Desglosar el IVA que corresponde"
         Top             =   1500
         Width           =   2000
      End
      Begin VB.CheckBox chkDesglosarIVACopago 
         Caption         =   "Copago"
         Height          =   190
         Left            =   700
         TabIndex        =   10
         ToolTipText     =   "Desglosar el IVA que corresponde"
         Top             =   1800
         Width           =   2000
      End
      Begin VB.CheckBox chkDesglosarExcedente 
         Caption         =   "Desglosar importes gravado y no gravado"
         Height          =   190
         Left            =   3500
         TabIndex        =   1
         ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
         Top             =   300
         Width           =   3400
      End
      Begin VB.CheckBox chkDesglosarIVAExcedente 
         Caption         =   "Excedente"
         Height          =   190
         Left            =   700
         TabIndex        =   0
         ToolTipText     =   "Desglosar el IVA que corresponde"
         Top             =   300
         Width           =   2000
      End
   End
   Begin VB.Frame Frame2 
      Height          =   670
      Left            =   4027
      TabIndex        =   46
      Top             =   8100
      Width           =   630
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametrosCuenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Grabar los parámetros de la cuenta"
         Top             =   130
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmParametrosCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Esta pantalla se usa para configurar parametros de control de aseguradora a nivel de la cuenta del paciente.
'-----------------------------------------------------------------------------

Option Explicit

Public blnConfiguracionGuardada As Boolean
Public blnHabilita As Boolean

Private Sub chkCalcularCargosSeleccionados_Click()
    If chkCalcularCargosSeleccionados.Value Then
        chkCoaseguroPorFactura.Enabled = True
    Else
        chkCoaseguroPorFactura.Value = 0
        chkCoaseguroPorFactura.Enabled = False
    End If
End Sub

Private Sub chkCalcularCargosSeleccionados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkCoaseguroPorFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosaIVATasaHospi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarCoaseguroM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarIVACoaseguro_Click()
    If chkDesglosarIVACoaseguro.Value = 0 Then
        chkDesglosarCoaseguro.Enabled = False
        chkDesglosarCoaseguro.Value = 0
    Else
        chkDesglosarCoaseguro.Enabled = True
    End If
End Sub

Private Sub chkDesglosarIVACoaseguroAdicional_Click()
    If chkDesglosarIVACoaseguroAdicional.Value = 0 Then
        chkDesglosarCoaseguroAdicional.Enabled = False
        chkDesglosarCoaseguroAdicional.Value = 0
    Else
        chkDesglosarCoaseguroAdicional.Enabled = True
    End If
End Sub

Private Sub chkDesglosarIVACoaseguroM_Click()
    If chkDesglosarIVACoaseguroM.Value = 0 Then
        chkDesglosarCoaseguroM.Enabled = False
        chkDesglosarCoaseguroM.Value = 0
    Else
        chkDesglosarCoaseguroM.Enabled = True
    End If
End Sub

Private Sub chkDesglosarIVACoaseguroM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarIVACopago_Click()
    If chkDesglosarIVACopago.Value = 0 Then
        chkDesglosarCopago.Enabled = False
        chkDesglosarCopago.Value = 0
    Else
        chkDesglosarCopago.Enabled = True
    End If
End Sub

Private Sub chkDesglosarIVADeducible_Click()
    If chkDesglosarIVADeducible.Value = 0 Then
        chkDesglosarDeducible.Enabled = False
        chkDesglosarDeducible.Value = 0
    Else
        chkDesglosarDeducible.Enabled = True
    End If
End Sub

Private Sub chkDesglosarIVAExcedente_Click()
    If chkDesglosarIVAExcedente.Value = 0 Then
        chkDesglosarExcedente.Enabled = False
        chkDesglosarExcedente.Value = 0
    Else
        chkDesglosarExcedente.Enabled = True
    End If
End Sub

Private Sub chkDesglosarCoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarCoaseguroAdicional_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarCopago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarDeducible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarExcedente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarIVACoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarIVACoaseguroAdicional_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarIVACopago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarIVADeducible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDesglosarIVAExcedente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkTotCS_Click()
    If chkTotCS.Value = 1 Then
        optTipoDesglose(0).Enabled = True
        optTipoDesglose(1).Enabled = True
    Else
        optTipoDesglose(0).Enabled = False
        optTipoDesglose(1).Enabled = False
    End If
End Sub

Private Sub chkTotCS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    Dim rsNumeroRegistros As New ADODB.Recordset
    Dim X As Integer
    Dim vllngPersonaGraba As Long
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim vlintErrorCuenta As Integer
    Dim arrCompara(7) As Long
    Dim inti As Long
    Dim intj As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 2447, 2490), "E", True) Then 'And blnHabilita Then
        blnConfiguracionGuardada = True
        Me.Visible = False
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
End Sub

Private Sub Form_Activate()
    chkTotCS_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = 27 Then
        blnConfiguracionGuardada = False
        Me.Visible = False
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon
                
    chkDesglosarIVACopago_Click
    chkDesglosarIVAExcedente_Click
    chkDesglosarIVADeducible_Click
    chkDesglosarIVACoaseguro_Click
    chkDesglosarIVACoaseguroM_Click
    chkDesglosarIVACoaseguroAdicional_Click
    chkCalcularCargosSeleccionados_Click
    
    blnConfiguracionGuardada = False
    
    pHabilita
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub optTipoDesctoCoaseguro_Click(Index As Integer)
    If Index = 0 Then   ' Por porcentaje
        txtPorcentajeCoaseguroPorNota.Text = ".00"
        txtCantidadLimiteCoaseguro.Text = Format(0, "$###,###,###,###.00")
        lbCantidadLimiteCoaseguro.Enabled = False
        lbPorcentaje(2).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoaseguro.Enabled = False
        txtCantidadLimiteCoaseguro.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCoaseguro.Enabled = False
        lbPorcentaje(2).Visible = False
        txtPorcentajeCoaseguroPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
End Sub

Private Sub optTipoDesctoCoaseguro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoCoaseguroAdicional_Click(Index As Integer)
    If Index = 0 Then   ' Por porcentaje
        txtPorcentajeCoasAdicionalPorNota.Text = ".00"
        txtCantidadLimiteCoasAdicional.Text = Format(0, "$###,###,###,###.00")
        lbCantidadLimiteCoaAdicional.Enabled = False
        lbPorcentaje(4).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoaAdicional.Enabled = False
        txtCantidadLimiteCoasAdicional.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCoasAdicional.Enabled = False
        lbPorcentaje(4).Visible = False
        txtPorcentajeCoasAdicionalPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
End Sub

Private Sub optTipoDesctoCoaseguroAdicional_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoCoaseguroMedico_Click(Index As Integer)
    If Index = 0 Then   ' Por porcentaje
        txtPorcentajeCoaseguroMPorNota.Text = ".00"
        txtCantidadLimiteCoaseguroM.Text = Format(0, "$###,###,###,###.00")
        lbCantidadLimiteCoaMedico.Enabled = False
        lbPorcentaje(3).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoaMedico.Enabled = False
        txtCantidadLimiteCoaseguroM.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCoaseguroM.Enabled = False
        lbPorcentaje(3).Visible = False
        txtPorcentajeCoaseguroMPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
End Sub

Private Sub optTipoDesctoCoaseguroMedico_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoCopago_Click(Index As Integer)
    If Index = 0 Then   ' Por porcentaje
        txtPorcentajeCopagoPorNota.Text = ".00"
        txtCantidadLimiteCopago.Text = Format(0, "$###,###,###,###.00")
        lbCantidadLimiteCoPago.Enabled = False
        lbPorcentaje(5).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoPago.Enabled = False
        txtCantidadLimiteCopago.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCopago.Enabled = False
        lbPorcentaje(5).Visible = False
        txtPorcentajeCopagoPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
End Sub

Private Sub optTipoDesctoCopago_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoDeducible_Click(Index As Integer)
    If Index = 0 Then   ' Por porcentaje
        txtPorcentajeDeduciblePorNota.Text = ".00"
        txtCantidadLimiteDeducible.Text = Format(0, "$###,###,###,###.00")
        lbCantidadLimiteDeducible.Enabled = False
        lbPorcentaje(1).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteDeducible.Enabled = False
        txtCantidadLimiteDeducible.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteDeducible.Enabled = False
        lbPorcentaje(1).Visible = False
        txtPorcentajeDeduciblePorNota.Text = Format(0, "$###,###,###,###.00")
    End If
End Sub

Private Sub optTipoDesctoDeducible_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoExcedente_Click(Index As Integer)
    If Index = 0 Then   ' Por porcentaje
        txtPorcentajeExcedentePorNota.Text = ".00"
        txtCantidadLimiteExcedente.Text = Format(0, "$###,###,###,###.00")
        lbCantidadLimiteExcedente.Enabled = False
        lbPorcentaje(0).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteExcedente.Enabled = False
        txtCantidadLimiteExcedente.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteExcedente.Enabled = False
        lbPorcentaje(0).Visible = False
        txtPorcentajeExcedentePorNota.Text = Format(0, "$###,###,###,###.00")
    End If
End Sub

Private Sub optTipoDesctoExcedente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub optIncCS_Click(Index As Integer)

        chkTotCS.Enabled = True

End Sub

Private Sub optIncCS_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub optTipoDesglose_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtCantidadLimiteCoasAdicional_Click()
    pSelTextBox txtCantidadLimiteCoasAdicional
End Sub

Private Sub txtCantidadLimiteCoasAdicional_GotFocus()
    pSelTextBox txtCantidadLimiteCoasAdicional
End Sub

Private Sub txtCantidadLimiteCoasAdicional_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidadLimiteCoasAdicional, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCantidadLimiteCoasAdicional_LostFocus()
    txtCantidadLimiteCoasAdicional.Text = Format(Val(Format(txtCantidadLimiteCoasAdicional.Text, "")), "$###,###,###,###.00")
End Sub

Private Sub txtCantidadLimiteCoaseguro_Click()
    pSelTextBox txtCantidadLimiteCoaseguro
End Sub

Private Sub txtCantidadLimiteCoaseguro_GotFocus()
    pSelTextBox txtCantidadLimiteCoaseguro
End Sub

Private Sub txtCantidadLimiteCoaseguro_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidadLimiteCoaseguro, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCantidadLimiteCoaseguro_LostFocus()
    txtCantidadLimiteCoaseguro.Text = Format(Val(Format(txtCantidadLimiteCoaseguro.Text, "")), "$###,###,###,###.00")
End Sub

Private Sub txtCantidadLimiteCoaseguroM_Click()
    pSelTextBox txtCantidadLimiteCoaseguroM
End Sub

Private Sub txtCantidadLimiteCoaseguroM_GotFocus()
    pSelTextBox txtCantidadLimiteCoaseguroM
End Sub

Private Sub txtCantidadLimiteCoaseguroM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtCantidadLimiteCoaseguroM_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidadLimiteCoaseguroM, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCantidadLimiteCoaseguroM_LostFocus()
    txtCantidadLimiteCoaseguroM.Text = Format(Val(Format(txtCantidadLimiteCoaseguroM.Text, "")), "$###,###,###,###.00")
End Sub

Private Sub txtCantidadLimiteCopago_Click()
    pSelTextBox txtCantidadLimiteCopago
End Sub

Private Sub txtCantidadLimiteCopago_GotFocus()
    pSelTextBox txtCantidadLimiteCopago
End Sub

Private Sub txtCantidadLimiteCopago_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidadLimiteCopago, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCantidadLimiteCopago_LostFocus()
    txtCantidadLimiteCopago.Text = Format(Val(Format(txtCantidadLimiteCopago.Text, "")), "$###,###,###,###.00")
End Sub

Private Sub txtCantidadLimiteDeducible_Click()
    pSelTextBox txtCantidadLimiteDeducible
End Sub

Private Sub txtCantidadLimiteDeducible_GotFocus()
    pSelTextBox txtCantidadLimiteDeducible
End Sub

Private Sub txtCantidadLimiteDeducible_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidadLimiteDeducible, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCantidadLimiteDeducible_LostFocus()
    txtCantidadLimiteDeducible.Text = Format(Val(Format(txtCantidadLimiteDeducible.Text, "")), "$###,###,###,###.00")
End Sub

Private Sub txtCantidadLimiteExcedente_Click()
    pSelTextBox txtCantidadLimiteExcedente
End Sub

Private Sub txtCantidadLimiteExcedente_GotFocus()
    pSelTextBox txtCantidadLimiteExcedente
End Sub

Private Sub txtCantidadLimiteExcedente_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidadLimiteExcedente, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCantidadLimiteExcedente_LostFocus()
    txtCantidadLimiteExcedente.Text = Format(Val(Format(txtCantidadLimiteExcedente.Text, "")), "$###,###,###,###.00")
End Sub

Private Sub txtPorcentajeCoasAdicionalPorNota_GotFocus()
    pSelTextBox txtPorcentajeCoasAdicionalPorNota
End Sub

Private Sub txtPorcentajeCoasAdicionalPorNota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optTipoDesctoCoaseguroAdicional(0).Value Then
            If Val(txtPorcentajeCoasAdicionalPorNota.Text) > 0 Then
                txtCantidadLimiteCoasAdicional.Enabled = True
                lbCantidadLimiteCoaAdicional.Enabled = True
            Else
                txtCantidadLimiteCoasAdicional.Text = 0
                txtCantidadLimiteCoasAdicional.Enabled = False
                lbCantidadLimiteCoaAdicional.Enabled = False
            End If
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtPorcentajeCoasAdicionalPorNota_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtPorcentajeCoasAdicionalPorNota, KeyAscii, 2) Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtPorcentajeCoasAdicionalPorNota_LostFocus()
    If optTipoDesctoCoaseguroAdicional(0).Value Then
        txtPorcentajeCoasAdicionalPorNota.Text = Format(IIf(Val(txtPorcentajeCoasAdicionalPorNota.Text) > 100, "100", Val(txtPorcentajeCoasAdicionalPorNota.Text)), "###.00")
        txtCantidadLimiteCoasAdicional.Text = Format(Val(Format(txtCantidadLimiteCoasAdicional.Text, "")), "$###,###,###,###.00")
    Else
        txtPorcentajeCoasAdicionalPorNota.Text = Format(Val(Format(txtPorcentajeCoasAdicionalPorNota.Text, "")), "$###,###,###,###.00")
    End If
    If optTipoDesctoCoaseguroAdicional(0).Value Then
        If Val(txtPorcentajeCoasAdicionalPorNota.Text) = 0 And txtCantidadLimiteCoasAdicional.Enabled Then
            txtCantidadLimiteCoasAdicional.Enabled = False
            lbCantidadLimiteCoaAdicional.Enabled = False
        End If
    End If
End Sub

Private Sub txtPorcentajeCoaseguroMPorNota_GotFocus()
    pSelTextBox txtPorcentajeCoaseguroMPorNota
End Sub

Private Sub txtPorcentajeCoaseguroMPorNota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optTipoDesctoCoaseguroMedico(0).Value Then
            If Val(txtPorcentajeCoaseguroMPorNota.Text) > 0 Then
                txtCantidadLimiteCoaseguroM.Enabled = True
                lbCantidadLimiteCoaMedico.Enabled = True
            Else
                txtCantidadLimiteCoaseguroM.Text = 0
                txtCantidadLimiteCoaseguroM.Enabled = False
                lbCantidadLimiteCoaMedico.Enabled = False
            End If
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtPorcentajeCoaseguroMPorNota_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtPorcentajeCoaseguroMPorNota, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtPorcentajeCoaseguroMPorNota_LostFocus()
    If optTipoDesctoCoaseguroMedico(0).Value Then
        txtPorcentajeCoaseguroMPorNota.Text = Format(IIf(Val(txtPorcentajeCoaseguroMPorNota.Text) > 100, "100", Val(txtPorcentajeCoaseguroMPorNota.Text)), "###.00")
        txtCantidadLimiteCoaseguroM.Text = Format(Val(Format(txtCantidadLimiteCoaseguroM.Text, "")), "$###,###,###,###.00")
    Else
        txtPorcentajeCoaseguroMPorNota.Text = Format(Val(Format(txtPorcentajeCoaseguroMPorNota.Text, "")), "$###,###,###,###.00")
    End If
    If optTipoDesctoCoaseguroMedico(0).Value Then
        If Val(txtPorcentajeCoaseguroMPorNota.Text) = 0 And txtCantidadLimiteCoaseguroM.Enabled Then
            txtCantidadLimiteCoaseguroM.Enabled = False
            lbCantidadLimiteCoaMedico.Enabled = False
        End If
    End If
End Sub

Private Sub txtPorcentajeCoaseguroPorNota_GotFocus()
    pSelTextBox txtPorcentajeCoaseguroPorNota
End Sub

Private Sub txtPorcentajeCoaseguroPorNota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optTipoDesctoCoaseguro(0).Value Then
            If Val(txtPorcentajeCoaseguroPorNota.Text) > 0 Then
                txtCantidadLimiteCoaseguro.Enabled = True
                lbCantidadLimiteCoaseguro.Enabled = True
            Else
                txtCantidadLimiteCoaseguro.Text = 0
                txtCantidadLimiteCoaseguro.Enabled = False
                lbCantidadLimiteCoaseguro.Enabled = False
            End If
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtPorcentajeCoaseguroPorNota_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtPorcentajeCoaseguroPorNota, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtPorcentajeCoaseguroPorNota_LostFocus()
    If optTipoDesctoCoaseguro(0).Value Then
        txtPorcentajeCoaseguroPorNota.Text = Format(IIf(Val(txtPorcentajeCoaseguroPorNota.Text) > 100, "100", Val(txtPorcentajeCoaseguroPorNota.Text)), "###.00")
        txtCantidadLimiteCoaseguro.Text = Format(Val(Format(txtCantidadLimiteCoaseguro.Text, "")), "$###,###,###,###.00")
    Else
        txtPorcentajeCoaseguroPorNota.Text = Format(Val(Format(txtPorcentajeCoaseguroPorNota.Text, "")), "$###,###,###,###.00")
    End If
    If optTipoDesctoCoaseguro(0).Value Then
        If Val(txtPorcentajeCoaseguroPorNota.Text) = 0 And txtCantidadLimiteCoaseguro.Enabled Then
            txtCantidadLimiteCoaseguro.Enabled = False
            lbCantidadLimiteCoaseguro.Enabled = False
        End If
    End If
End Sub

Private Sub txtPorcentajeCopagoPorNota_GotFocus()
    pSelTextBox txtPorcentajeCopagoPorNota
End Sub

Private Sub txtPorcentajeCopagoPorNota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optTipoDesctoCopago(0).Value Then
            If Val(txtPorcentajeCopagoPorNota.Text) > 0 Then
                txtCantidadLimiteCopago.Enabled = True
                lbCantidadLimiteCoPago.Enabled = True
            Else
                txtCantidadLimiteCopago.Text = 0
                txtCantidadLimiteCopago.Enabled = False
                lbCantidadLimiteCoPago.Enabled = False
            End If
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtPorcentajeCopagoPorNota_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtPorcentajeCopagoPorNota, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtPorcentajeCopagoPorNota_LostFocus()
    If optTipoDesctoCopago(0).Value Then
        txtPorcentajeCopagoPorNota.Text = Format(IIf(Val(txtPorcentajeCopagoPorNota.Text) > 100, "100", Val(txtPorcentajeCopagoPorNota.Text)), "###.00")
        txtCantidadLimiteCopago.Text = Format(Val(Format(txtCantidadLimiteCopago.Text, "")), "$###,###,###,###.00")
    Else
        txtPorcentajeCopagoPorNota.Text = Format(Val(Format(txtPorcentajeCopagoPorNota.Text, "")), "$###,###,###,###.00")
    End If
    If optTipoDesctoCopago(0).Value Then
        If Val(txtPorcentajeCopagoPorNota.Text) = 0 And txtCantidadLimiteCopago.Enabled Then
            txtCantidadLimiteCopago.Enabled = False
            lbCantidadLimiteCoPago.Enabled = False
        End If
    End If
End Sub

Private Sub txtPorcentajeDeduciblePorNota_GotFocus()
    pSelTextBox txtPorcentajeDeduciblePorNota
End Sub

Private Sub txtPorcentajeDeduciblePorNota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optTipoDesctoDeducible(0).Value Then
            If Val(txtPorcentajeDeduciblePorNota.Text) > 0 Then
                txtCantidadLimiteDeducible.Enabled = True
                lbCantidadLimiteDeducible.Enabled = True
            Else
                txtCantidadLimiteDeducible.Text = 0
                txtCantidadLimiteDeducible.Enabled = False
                lbCantidadLimiteDeducible.Enabled = False
            End If
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtPorcentajeDeduciblePorNota_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtPorcentajeDeduciblePorNota, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtPorcentajeDeduciblePorNota_LostFocus()
    If optTipoDesctoDeducible(0).Value Then
        txtPorcentajeDeduciblePorNota.Text = Format(IIf(Val(txtPorcentajeDeduciblePorNota.Text) > 100, "100", Val(txtPorcentajeDeduciblePorNota.Text)), "###.00")
        txtCantidadLimiteDeducible.Text = Format(Val(Format(txtCantidadLimiteDeducible.Text, "")), "$###,###,###,###.00")
    Else
        txtPorcentajeDeduciblePorNota.Text = Format(Val(Format(txtPorcentajeDeduciblePorNota.Text, "")), "$###,###,###,###.00")
    End If
    If optTipoDesctoDeducible(0).Value Then
        If Val(txtPorcentajeDeduciblePorNota.Text) = 0 And txtCantidadLimiteDeducible.Enabled Then
            txtCantidadLimiteDeducible.Enabled = False
            lbCantidadLimiteDeducible.Enabled = False
        End If
    End If
End Sub

Private Sub txtPorcentajeExcedentePorNota_GotFocus()
    pSelTextBox txtPorcentajeExcedentePorNota
End Sub

Private Sub txtPorcentajeExcedentePorNota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optTipoDesctoExcedente(0).Value Then
            If Val(txtPorcentajeExcedentePorNota.Text) > 0 Then
                txtCantidadLimiteExcedente.Enabled = True
                lbCantidadLimiteExcedente.Enabled = True
            Else
                txtCantidadLimiteExcedente.Text = 0
                txtCantidadLimiteExcedente.Enabled = False
                lbCantidadLimiteExcedente.Enabled = False
            End If
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtPorcentajeExcedentePorNota_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtPorcentajeExcedentePorNota, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCantidadLimiteCoasAdicional_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtCantidadLimiteCoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtCantidadLimiteCopago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        chkDesglosaIVATasaHospi.SetFocus
    End If
End Sub

Private Sub txtCantidadLimiteDeducible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtCantidadLimiteExcedente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub pHabilita()
    Dim vlHabilitado As Boolean
    
    vlHabilitado = IIf(fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 2447, 2490), "E", True) And blnHabilita, True, False)
    fraDesglosaIVA.Enabled = vlHabilitado
    fraNotasDeCredito.Enabled = vlHabilitado
    fraParametros.Enabled = vlHabilitado
    'fraIncluir.Enabled = vlHabilitado
    'cmdSave.Enabled = vlHabilitado
End Sub

Private Sub txtPorcentajeExcedentePorNota_LostFocus()
On Error GoTo NotificaError
    
    If optTipoDesctoExcedente(0).Value Then
        txtPorcentajeExcedentePorNota.Text = Format(IIf(Val(txtPorcentajeExcedentePorNota.Text) > 100, "100", Val(txtPorcentajeExcedentePorNota.Text)), "###.00")
        txtCantidadLimiteExcedente.Text = Format(Val(Format(txtCantidadLimiteExcedente.Text, "")), "$###,###,###,###.00")
    Else
        txtPorcentajeExcedentePorNota.Text = Format(Val(Format(txtPorcentajeExcedentePorNota.Text, "")), "$###,###,###,###.00")
    End If
    If optTipoDesctoExcedente(0).Value Then
        If Val(txtPorcentajeExcedentePorNota.Text) = 0 And txtCantidadLimiteExcedente.Enabled = True Then
            txtCantidadLimiteExcedente.Enabled = False
            lbCantidadLimiteExcedente.Enabled = False
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorcentajeExcedentePorNota_LostFocus"))
    
End Sub
