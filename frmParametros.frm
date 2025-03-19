VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros del módulo"
   ClientHeight    =   11280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   685
      Left            =   4320
      TabIndex        =   110
      Top             =   10550
      Width           =   640
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   41
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmParametros.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Grabar los parámetros"
         Top             =   130
         Width           =   540
      End
   End
   Begin TabDlg.SSTab sstPropiedades 
      Height          =   11625
      Left            =   0
      TabIndex        =   111
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   20505
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Cuentas contables"
      TabPicture(0)   =   "frmParametros.frx":0342
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Impresión"
      TabPicture(1)   =   "frmParametros.frx":035E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Venta público"
      TabPicture(2)   =   "frmParametros.frx":037A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sstVentaPublico"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Aseguradoras"
      TabPicture(3)   =   "frmParametros.frx":0396
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraNotasDeCredito"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).Control(2)=   "fraDesglosaIVA"
      Tab(3).Control(3)=   "tmrSetFocus"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Generales"
      TabPicture(4)   =   "frmParametros.frx":03B2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame fraNotasDeCredito 
         Caption         =   "Descuentos por notas de crédito"
         Height          =   2565
         Left            =   -74895
         TabIndex        =   163
         Top             =   7080
         Width           =   9150
         Begin VB.TextBox txtCantidadLimiteCoaseguroM 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7875
            MaxLength       =   10
            TabIndex        =   64
            ToolTipText     =   "Cantidad límite"
            Top             =   1410
            Width           =   1100
         End
         Begin VB.TextBox txtCantidadLimiteDeducible 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7875
            MaxLength       =   10
            TabIndex        =   56
            ToolTipText     =   "Cantidad límite"
            Top             =   660
            Width           =   1100
         End
         Begin VB.TextBox txtCantidadLimiteCoaseguro 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7875
            MaxLength       =   10
            TabIndex        =   60
            ToolTipText     =   "Cantidad límite"
            Top             =   1040
            Width           =   1100
         End
         Begin VB.TextBox txtCantidadLimiteCoasAdicional 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7875
            MaxLength       =   10
            TabIndex        =   68
            ToolTipText     =   "Cantidad límite"
            Top             =   1770
            Width           =   1100
         End
         Begin VB.TextBox txtCantidadLimiteCopago 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7875
            MaxLength       =   10
            TabIndex        =   72
            ToolTipText     =   "Cantidad límite"
            Top             =   2130
            Width           =   1100
         End
         Begin VB.TextBox txtPorcentajeCoaseguroMPorNota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4995
            MaxLength       =   10
            TabIndex        =   63
            ToolTipText     =   "Descuento para las notas de crédito automáticas"
            Top             =   1410
            Width           =   1100
         End
         Begin VB.TextBox txtPorcentajeDeduciblePorNota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4995
            MaxLength       =   10
            TabIndex        =   55
            ToolTipText     =   "Descuento para las notas de crédito automáticas"
            Top             =   660
            Width           =   1100
         End
         Begin VB.TextBox txtPorcentajeCoaseguroPorNota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4995
            MaxLength       =   10
            TabIndex        =   59
            ToolTipText     =   "Descuento para las notas de crédito automáticas"
            Top             =   1040
            Width           =   1100
         End
         Begin VB.TextBox txtPorcentajeCoasAdicionalPorNota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4995
            MaxLength       =   10
            TabIndex        =   67
            ToolTipText     =   "Descuento para las notas de crédito automáticas"
            Top             =   1770
            Width           =   1100
         End
         Begin VB.TextBox txtPorcentajeCopagoPorNota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4995
            MaxLength       =   10
            TabIndex        =   71
            ToolTipText     =   "Descuento para las notas de crédito automáticas"
            Top             =   2130
            Width           =   1100
         End
         Begin VB.TextBox txtCantidadLimiteExcedente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7875
            MaxLength       =   10
            TabIndex        =   52
            ToolTipText     =   "Cantidad límite"
            Top             =   280
            Width           =   1100
         End
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
            Left            =   4995
            MaxLength       =   10
            TabIndex        =   51
            ToolTipText     =   "Descuento para las notas de crédito automáticas"
            Top             =   280
            Width           =   1100
         End
         Begin VB.Frame Frame16 
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1800
            TabIndex        =   169
            Top             =   240
            Width           =   3015
            Begin VB.OptionButton optTipoDesctoExcedente 
               Caption         =   "Por porcentaje"
               Height          =   220
               Index           =   0
               Left            =   120
               TabIndex        =   49
               ToolTipText     =   "Descuento por porcentaje"
               Top             =   120
               Width           =   1335
            End
            Begin VB.OptionButton optTipoDesctoExcedente 
               Caption         =   "Por cantidad"
               Height          =   220
               Index           =   1
               Left            =   1680
               TabIndex        =   50
               ToolTipText     =   "Descuento por cantidad"
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.Frame Frame11 
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1800
            TabIndex        =   168
            Top             =   600
            Width           =   3015
            Begin VB.OptionButton optTipoDesctoDeducible 
               Caption         =   "Por porcentaje"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   53
               ToolTipText     =   "Descuento por porcentaje"
               Top             =   120
               Width           =   1335
            End
            Begin VB.OptionButton optTipoDesctoDeducible 
               Caption         =   "Por cantidad"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   54
               ToolTipText     =   "Descuento por cantidad"
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame Frame12 
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1800
            TabIndex        =   167
            Top             =   960
            Width           =   3015
            Begin VB.OptionButton optTipoDesctoCoaseguro 
               Caption         =   "Por porcentaje"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   57
               ToolTipText     =   "Descuento por porcentaje"
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optTipoDesctoCoaseguro 
               Caption         =   "Por cantidad"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   58
               ToolTipText     =   "Descuento por cantidad"
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.Frame Frame13 
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1800
            TabIndex        =   166
            Top             =   1320
            Width           =   3015
            Begin VB.OptionButton optTipoDesctoCoaseguroMedico 
               Caption         =   "Por porcentaje"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   61
               ToolTipText     =   "Descuento por porcentaje"
               Top             =   120
               Width           =   1335
            End
            Begin VB.OptionButton optTipoDesctoCoaseguroMedico 
               Caption         =   "Por cantidad"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   62
               ToolTipText     =   "Descuento por cantidad"
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame Frame14 
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1800
            TabIndex        =   165
            Top             =   1680
            Width           =   3015
            Begin VB.OptionButton optTipoDesctoCoaseguroAdicional 
               Caption         =   "Por porcentaje"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   65
               ToolTipText     =   "Descuento por porcentaje"
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optTipoDesctoCoaseguroAdicional 
               Caption         =   "Por cantidad"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   66
               ToolTipText     =   "Descuento por cantidad"
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Height          =   400
            Left            =   1800
            TabIndex        =   164
            Top             =   2040
            Width           =   3015
            Begin VB.OptionButton optTipoDesctoCopago 
               Caption         =   "Por porcentaje"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   69
               ToolTipText     =   "Descuento por porcentaje"
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optTipoDesctoCopago 
               Caption         =   "Por cantidad"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   70
               ToolTipText     =   "Descuento por cantidad"
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.Label lbCantidadLimiteCoaMedico 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad límite"
            Height          =   195
            Left            =   6720
            TabIndex        =   187
            Top             =   1470
            Width           =   1050
         End
         Begin VB.Label lbCantidadLimiteDeducible 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad límite"
            Height          =   195
            Left            =   6720
            TabIndex        =   186
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label lbCantidadLimiteCoaseguro 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad límite"
            Height          =   195
            Left            =   6720
            TabIndex        =   185
            Top             =   1100
            Width           =   1050
         End
         Begin VB.Label lbCantidadLimiteCoaAdicional 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad límite"
            Height          =   195
            Left            =   6720
            TabIndex        =   184
            Top             =   1830
            Width           =   1050
         End
         Begin VB.Label lbCantidadLimiteCoPago 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad límite"
            Height          =   195
            Left            =   6720
            TabIndex        =   183
            Top             =   2190
            Width           =   1050
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Coaseguro médico"
            Height          =   195
            Left            =   200
            TabIndex        =   182
            Top             =   1470
            Width           =   1320
         End
         Begin VB.Label lbPorcentaje 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   5
            Left            =   6165
            TabIndex        =   181
            Top             =   2190
            Width           =   120
         End
         Begin VB.Label lbPorcentajeDescuentoDeducible 
            AutoSize        =   -1  'True
            Caption         =   "Deducible"
            Height          =   195
            Left            =   200
            TabIndex        =   180
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lbPorcentajeDescuentoCoaseguro 
            AutoSize        =   -1  'True
            Caption         =   "Coaseguro"
            Height          =   195
            Left            =   200
            TabIndex        =   179
            Top             =   1100
            Width           =   765
         End
         Begin VB.Label lbPorcentajeDescuentoCoaAdicional 
            AutoSize        =   -1  'True
            Caption         =   "Coaseguro adicional"
            Height          =   195
            Left            =   200
            TabIndex        =   178
            Top             =   1830
            Width           =   1440
         End
         Begin VB.Label lbPorcentajeDescuentoCoPago 
            AutoSize        =   -1  'True
            Caption         =   "Copago"
            Height          =   195
            Left            =   200
            TabIndex        =   177
            Top             =   2190
            Width           =   555
         End
         Begin VB.Label lbPorcentaje 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   1
            Left            =   6165
            TabIndex        =   176
            Top             =   720
            Width           =   120
         End
         Begin VB.Label lbPorcentaje 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   2
            Left            =   6165
            TabIndex        =   175
            Top             =   1100
            Width           =   120
         End
         Begin VB.Label lbPorcentaje 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   3
            Left            =   6165
            TabIndex        =   174
            Top             =   1470
            Width           =   120
         End
         Begin VB.Label lbPorcentaje 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   4
            Left            =   6165
            TabIndex        =   173
            Top             =   1830
            Width           =   120
         End
         Begin VB.Label lbCantidadLimiteExcedente 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad límite"
            Height          =   195
            Left            =   6720
            TabIndex        =   172
            Top             =   340
            Width           =   1050
         End
         Begin VB.Label lbPorcentajeDescuentoExcedente 
            AutoSize        =   -1  'True
            Caption         =   "Excedente"
            Height          =   195
            Left            =   200
            TabIndex        =   171
            Top             =   340
            Width           =   765
         End
         Begin VB.Label lbPorcentaje 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   0
            Left            =   6165
            TabIndex        =   170
            Top             =   340
            Width           =   120
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Conceptos de factura para seguros"
         Height          =   4480
         Left            =   -74895
         TabIndex        =   109
         Top             =   350
         Width           =   9150
         Begin VB.Frame Frame10 
            Height          =   720
            Left            =   195
            TabIndex        =   161
            Top             =   3600
            Width           =   8775
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   120
               TabIndex        =   162
               Top             =   360
               Width           =   4815
               Begin VB.OptionButton optTipoDesglose 
                  Caption         =   "Presentación en cuadrícula"
                  Height          =   255
                  Index           =   1
                  Left            =   1840
                  TabIndex        =   36
                  ToolTipText     =   "Presentación en cuadrícula"
                  Top             =   0
                  Width           =   2535
               End
               Begin VB.OptionButton optTipoDesglose 
                  Caption         =   "Presentación simple"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   35
                  ToolTipText     =   "Presentación simple"
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.CheckBox chkTotCS 
               Caption         =   "Desglosar descuentos comerciales y de seguros en totales de la representación impresa del CFDI"
               Enabled         =   0   'False
               Height          =   225
               Left            =   120
               TabIndex        =   34
               ToolTipText     =   "Desglosar descuentos comerciales y de seguros en totales de la representación impresa del CFDI"
               Top             =   0
               Width           =   7180
            End
         End
         Begin VB.OptionButton optIncCS 
            Caption         =   "Sumar a los descuentos"
            Height          =   225
            Index           =   1
            Left            =   6920
            TabIndex        =   33
            ToolTipText     =   "Sumar proporcional del monto pagado por conceptos de seguro a los descuentos"
            Top             =   3240
            Width           =   2175
         End
         Begin VB.OptionButton optIncCS 
            Caption         =   "Restar de los importes"
            Height          =   225
            Index           =   0
            Left            =   4920
            TabIndex        =   32
            ToolTipText     =   "Restar proporcional del monto pagado por conceptos de seguro a los importes"
            Top             =   3240
            Width           =   2055
         End
         Begin VB.ComboBox cboExcedenteIVA 
            Height          =   315
            Left            =   1750
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   30
            ToolTipText     =   "Concepto de factura para copago"
            Top             =   2460
            Width           =   7125
         End
         Begin VB.CheckBox chkCalcularCargosSeleccionados 
            Caption         =   "Calcular importes de conceptos de factura para seguros con base a los cargos seleccionados para facturar"
            Height          =   225
            Left            =   200
            TabIndex        =   31
            Top             =   2880
            Width           =   7980
         End
         Begin VB.ComboBox cboConceptoCoaseguroMedico 
            Height          =   315
            Left            =   1750
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            ToolTipText     =   "Concepto de factura para coaseguro médico"
            Top             =   1380
            Width           =   7125
         End
         Begin VB.ComboBox cboConceptoCoaseguroAdicional 
            Height          =   315
            Left            =   1750
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Concepto de factura para coaseguro adicional"
            Top             =   1740
            Width           =   7125
         End
         Begin VB.ComboBox cboConceptoSumaAsegurada 
            Height          =   315
            Left            =   1750
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Concepto de factura para excedente en suma asegurada"
            Top             =   300
            Width           =   7125
         End
         Begin VB.ComboBox cboConceptoDeducible 
            Height          =   315
            Left            =   1750
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "Concepto de factura para deducible"
            Top             =   660
            Width           =   7125
         End
         Begin VB.ComboBox cboConceptoCoaseguro 
            Height          =   315
            Left            =   1750
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Concepto de factura para coaseguro"
            Top             =   1010
            Width           =   7125
         End
         Begin VB.ComboBox cboConceptoCopago 
            Height          =   315
            Left            =   1750
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            ToolTipText     =   "Concepto de factura para copago"
            Top             =   2100
            Width           =   7125
         End
         Begin VB.Label Label40 
            Caption         =   "Descuento de conceptos de seguro en CFDI de aseguradora"
            Height          =   225
            Left            =   195
            TabIndex        =   160
            Top             =   3240
            Width           =   4575
         End
         Begin VB.Label Label37 
            Caption         =   "Excedente de IVA"
            Height          =   255
            Left            =   200
            TabIndex        =   149
            Top             =   2510
            Width           =   1350
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Coaseguro médico"
            Height          =   195
            Left            =   200
            TabIndex        =   147
            Top             =   1430
            Width           =   1320
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Coaseguro adicional"
            Height          =   195
            Left            =   200
            TabIndex        =   139
            Top             =   1790
            Width           =   1440
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Excedente"
            Height          =   195
            Left            =   200
            TabIndex        =   116
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Deducible"
            Height          =   195
            Left            =   200
            TabIndex        =   115
            Top             =   710
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Coaseguro"
            Height          =   195
            Left            =   200
            TabIndex        =   114
            Top             =   1070
            Width           =   765
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Copago"
            Height          =   195
            Left            =   200
            TabIndex        =   113
            Top             =   2150
            Width           =   555
         End
      End
      Begin VB.Frame fraDesglosaIVA 
         Caption         =   "Conceptos que desglosan IVA"
         Height          =   2085
         Left            =   -74895
         TabIndex        =   148
         ToolTipText     =   "Conceptos que desglosan IVA"
         Top             =   4892
         Width           =   9150
         Begin VB.CheckBox chkDesglosarIVAExcedente 
            Caption         =   "Excedente"
            Height          =   190
            Left            =   700
            TabIndex        =   37
            ToolTipText     =   "Desglosar el IVA que corresponde"
            Top             =   300
            Width           =   2000
         End
         Begin VB.CheckBox chkDesglosarExcedente 
            Caption         =   "Desglosar importes gravado y no gravado"
            Height          =   190
            Left            =   3500
            TabIndex        =   38
            ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
            Top             =   300
            Width           =   3400
         End
         Begin VB.CheckBox chkDesglosarIVACopago 
            Caption         =   "Copago"
            Height          =   190
            Left            =   700
            TabIndex        =   47
            ToolTipText     =   "Desglosar el IVA que corresponde"
            Top             =   1800
            Width           =   2000
         End
         Begin VB.CheckBox chkDesglosarIVACoaseguroAdicional 
            Caption         =   "Coaseguro adicional"
            Height          =   190
            Left            =   700
            TabIndex        =   45
            ToolTipText     =   "Desglosar el IVA que corresponde"
            Top             =   1500
            Width           =   2000
         End
         Begin VB.CheckBox chkDesglosarIVACoaseguro 
            Caption         =   "Coaseguro"
            Height          =   190
            Left            =   700
            TabIndex        =   41
            ToolTipText     =   "Desglosar el IVA que corresponde"
            Top             =   900
            Width           =   2000
         End
         Begin VB.CheckBox chkDesglosarIVADeducible 
            Caption         =   "Deducible"
            Height          =   190
            Left            =   700
            TabIndex        =   39
            ToolTipText     =   "Desglosar el IVA que corresponde"
            Top             =   600
            Width           =   2000
         End
         Begin VB.CheckBox chkDesglosarIVACoaseguroM 
            Caption         =   "Coaseguro médico"
            Height          =   190
            Left            =   700
            TabIndex        =   43
            ToolTipText     =   "Desglosar el IVA que corresponde"
            Top             =   1200
            Width           =   2000
         End
         Begin VB.CheckBox chkDesglosarCoaseguroM 
            Caption         =   "Desglosar importes gravado y no gravado"
            Height          =   190
            Left            =   3500
            TabIndex        =   44
            ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
            Top             =   1200
            Width           =   3400
         End
         Begin VB.CheckBox chkDesglosarDeducible 
            Caption         =   "Desglosar importes gravado y no gravado"
            Height          =   190
            Left            =   3500
            TabIndex        =   40
            ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
            Top             =   600
            Width           =   3400
         End
         Begin VB.CheckBox chkDesglosarCoaseguro 
            Caption         =   "Desglosar importes gravado y no gravado"
            Height          =   190
            Left            =   3500
            TabIndex        =   42
            ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
            Top             =   900
            Width           =   3400
         End
         Begin VB.CheckBox chkDesglosarCoaseguroAdicional 
            Caption         =   "Desglosar importes gravado y no gravado"
            Height          =   190
            Left            =   3500
            TabIndex        =   46
            ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
            Top             =   1500
            Width           =   3400
         End
         Begin VB.CheckBox chkDesglosarCopago 
            Caption         =   "Desglosar importes gravado y no gravado"
            Height          =   190
            Left            =   3500
            TabIndex        =   48
            ToolTipText     =   "Desglosar importes en la impresión de la factura (gravado y no gravado)"
            Top             =   1800
            Width           =   3400
         End
      End
      Begin VB.Timer tmrSetFocus 
         Interval        =   1000
         Left            =   -74880
         Top             =   9840
      End
      Begin VB.Frame Frame7 
         Height          =   10185
         Left            =   120
         TabIndex        =   133
         Top             =   360
         Width           =   9120
         Begin VB.CheckBox chkHonorarioMedicoCredito 
            Caption         =   "Generar cuenta por pagar de honorarios médicos al realizar facturas a crédito"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            ToolTipText     =   "Generar cuenta por pagar de honorarios médicos al realizar facturas a crédito"
            Top             =   3180
            Width           =   8775
         End
         Begin VB.ComboBox cboConceptoFacturacionAsistSocial 
            Height          =   315
            Left            =   5130
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   78
            ToolTipText     =   "Concepto de facturación para factura de asistencia social"
            Top             =   1830
            Width           =   3900
         End
         Begin VB.ComboBox cboUsoCFDIFacturado 
            Height          =   315
            Left            =   5130
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   77
            ToolTipText     =   "Selección del uso de CFDI que se empleará para solicitar el comprobante al médico al facturar honorarios"
            Top             =   1500
            Width           =   3900
         End
         Begin VB.ComboBox cboConceptoEntrada 
            Height          =   315
            Left            =   5130
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   76
            ToolTipText     =   "Selección del concepto de entrada de dinero para devolución de deudores diversos"
            Top             =   1170
            Width           =   3900
         End
         Begin VB.ComboBox cboConceptoHonorariosMedicos 
            Height          =   315
            Left            =   5130
            Style           =   2  'Dropdown List
            TabIndex        =   75
            ToolTipText     =   "Selección del concepto de facturación para honorarios médicos"
            Top             =   840
            Width           =   3900
         End
         Begin VB.CheckBox chkConservarPrecioDescuento 
            Caption         =   "Conservar precio y descuento al excluir cargos"
            Height          =   315
            Left            =   120
            TabIndex        =   80
            ToolTipText     =   "Conservar precio y descuento al excluir cargos"
            Top             =   2385
            Width           =   7455
         End
         Begin VB.CheckBox chkdesgloseIEPS 
            Caption         =   "Desglosar IEPS en impresión de tickets en venta al público"
            Height          =   315
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "Desglosar IEPS en impresión de tickets en venta al público"
            Top             =   2895
            Width           =   7455
         End
         Begin VB.CheckBox chkDesactivarExterno 
            Caption         =   "Desactivar automáticamente los pacientes externos a los que se les facture y cierre la cuenta"
            Height          =   435
            Left            =   120
            TabIndex        =   81
            ToolTipText     =   "Desactivará automáticamente a los pacientes externos a los que se les facture y cierre la cuenta "
            Top             =   2580
            Width           =   7455
         End
         Begin VB.ComboBox cboConceptoParcial 
            Height          =   315
            Left            =   5130
            Style           =   2  'Dropdown List
            TabIndex        =   74
            ToolTipText     =   "Selección del concepto de facturación parcial"
            Top             =   510
            Width           =   3900
         End
         Begin VB.ComboBox cboDepartamentoMsg 
            Height          =   315
            Left            =   5130
            Style           =   2  'Dropdown List
            TabIndex        =   73
            ToolTipText     =   "Selección del departamento con que entrará el usuario"
            Top             =   180
            Width           =   3900
         End
         Begin VB.CheckBox chkCerrarCuentasAut 
            Caption         =   "Cerrar automáticamente las cuentas de externos que excedan el límite máximo de días de apertura"
            Height          =   270
            Left            =   120
            TabIndex        =   79
            Top             =   2175
            Width           =   7455
         End
         Begin VB.Frame frmchecks 
            BorderStyle     =   0  'None
            Height          =   6690
            Left            =   30
            TabIndex        =   152
            Top             =   3450
            Width           =   8990
            Begin VB.CheckBox ChkAuditoriaCargos 
               Caption         =   "Realizar auditoría de cargos"
               Height          =   255
               Left            =   6120
               TabIndex        =   99
               ToolTipText     =   "Realizar auditoría de cargos"
               Top             =   3380
               Width           =   2775
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Trasladar cargos sin actualizar precios y descuentos en forma predeterminada"
               Height          =   195
               Left            =   120
               TabIndex        =   194
               ToolTipText     =   "Trasladar cargos sin actualizar precios y descuentos en forma predeterminada"
               Top             =   0
               Width           =   6015
            End
            Begin VB.CheckBox chkValidacionPMPVentaPublico 
               Caption         =   "Habilitar la validación del precio máximo al público de medicamentos en venta al público"
               Height          =   195
               Left            =   95
               TabIndex        =   193
               ToolTipText     =   "Habilitar la validación del precio máximo al público de medicamentos en venta al público"
               Top             =   4160
               Width           =   6615
            End
            Begin VB.CheckBox chkCapturarMargenSubrogado 
               Caption         =   "Capturar margen subrogado en las listas de precios por cargo"
               Height          =   195
               Left            =   95
               TabIndex        =   101
               ToolTipText     =   "Capturar margen subrogado en las listas de precios por cargo"
               Top             =   3900
               Width           =   6015
            End
            Begin VB.CheckBox chkAbrirCuentaExterna 
               Caption         =   "Abrir cuenta de externo para trasladar el medicamento"
               Height          =   195
               Left            =   95
               TabIndex        =   100
               ToolTipText     =   "Abrir cuenta de externo para trasladar el medicamento"
               Top             =   3640
               Width           =   6015
            End
            Begin VB.CheckBox chkValidaDoble 
               Caption         =   "Solicitar doble contraseña al dar de alta socios"
               Height          =   195
               Left            =   3120
               TabIndex        =   96
               ToolTipText     =   "Solicitar doble contraseña al dar de alta socios"
               Top             =   2880
               Value           =   1  'Checked
               Width           =   3735
            End
            Begin VB.TextBox txtDiasSinRespPresupuesto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5760
               MaxLength       =   3
               TabIndex        =   106
               ToolTipText     =   "Número de días máximo de espera de respuesta en los presupuestos"
               Top             =   5715
               Width           =   630
            End
            Begin VB.CheckBox chkCancelarRecibosOtroDepto 
               Caption         =   "Permitir cancelar entradas y salidas de dinero de otro departamento"
               Height          =   195
               Left            =   95
               TabIndex        =   89
               ToolTipText     =   "Permitir cancelar entradas y salidas de dinero de otro departamento"
               Top             =   1300
               Width           =   7455
            End
            Begin VB.CheckBox chkTrasladarCargos 
               Caption         =   "Trasladar cargos sin actualizar precios y descuentos en forma predeterminada"
               Height          =   195
               Left            =   95
               TabIndex        =   98
               ToolTipText     =   "Trasladar cargos sin actualizar precios y descuentos en forma predeterminada"
               Top             =   3380
               Width           =   6015
            End
            Begin VB.CheckBox chkPermitirCorte 
               Caption         =   "Permitir cerrar el corte de caja chica si existen salidas sin XML relacionados"
               Height          =   195
               Left            =   95
               TabIndex        =   86
               ToolTipText     =   "Indica si se permitirá cerrar el corte de caja chica si existen salidas sin XML relacionados"
               Top             =   520
               Width           =   7455
            End
            Begin VB.CheckBox chkCorteHonorarios 
               Caption         =   "Incluir en el corte entradas de dinero por honorarios médicos"
               Height          =   195
               Left            =   95
               TabIndex        =   85
               ToolTipText     =   "Entrada de dinero en efectivo por concepto de honorarios"
               Top             =   260
               Width           =   7455
            End
            Begin VB.TextBox txtTituloCtasPendFact 
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   107
               Top             =   6315
               Width           =   8870
            End
            Begin VB.TextBox txtIntervaloMsgCargo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               MaxLength       =   2
               TabIndex        =   103
               Top             =   4725
               Width           =   630
            End
            Begin VB.CheckBox chkVerificar 
               Caption         =   "Verificar requisiciones pendientes de surtir y cargos automáticos pendientes de aplicar previo a facturar"
               Height          =   195
               Left            =   95
               TabIndex        =   94
               ToolTipText     =   "Verificar si existen requisiciones pendientes de surtir o si existen cargos automáticos pendientes de aplicar previo a facturar"
               Top             =   2600
               Width           =   8145
            End
            Begin VB.CheckBox Chkcuentacerrada 
               Caption         =   "Permitir devoluciones con cuenta cerrada"
               Height          =   195
               Left            =   95
               TabIndex        =   91
               Top             =   1820
               Width           =   7455
            End
            Begin VB.CheckBox chkRequisicionespendientes 
               Caption         =   "Permitir facturar cuentas con requisiciones pendientes de surtir y cargos automáticos pendientes de aplicar"
               Height          =   195
               Left            =   95
               TabIndex        =   93
               ToolTipText     =   "Permitir facturar cuentas aún con requisiciones pendientes de surtir o con cargos automáticos pendientes de aplicar"
               Top             =   2340
               Width           =   8055
            End
            Begin VB.CheckBox chkFacturaAutomatica 
               Caption         =   "Generar factura automática en venta al público"
               Height          =   195
               Left            =   95
               TabIndex        =   84
               ToolTipText     =   "Generar factura automática en venta al público"
               Top             =   0
               Width           =   7455
            End
            Begin VB.CheckBox chkSocios 
               Caption         =   "Habilitar la administración de socios"
               Height          =   195
               Left            =   95
               TabIndex        =   95
               ToolTipText     =   "Habilitar la administración de socios"
               Top             =   2860
               Width           =   2895
            End
            Begin VB.CheckBox chkPermitirFacturarCargosFueraCatalogo 
               Caption         =   "Permitir facturar con cargos que no existan en el catálogo de cargos por empresa "
               Height          =   195
               Left            =   95
               TabIndex        =   92
               ToolTipText     =   "Permitir facturar con cargos que no existan en el catálogo de cargos por empresa "
               Top             =   2080
               Width           =   7455
            End
            Begin VB.TextBox txtDiasAbrirCuentaInt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5760
               MaxLength       =   3
               TabIndex        =   104
               Top             =   5055
               Width           =   630
            End
            Begin VB.TextBox txtDiasAbrirCuentaExt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5760
               MaxLength       =   3
               TabIndex        =   105
               Top             =   5385
               Width           =   630
            End
            Begin VB.CheckBox chkCerrarCuentas 
               Caption         =   "Permitir cerrar cuentas con requisiciones pendientes de surtir"
               Height          =   195
               Left            =   95
               TabIndex        =   90
               ToolTipText     =   "Permitir cerrar cuentas cuando existen requisiciones pendientes de surtir"
               Top             =   1560
               Width           =   7455
            End
            Begin VB.CheckBox chkCuentaPuenteIngresos 
               Caption         =   "Utilizar cuenta puente en sustitución de cuentas de ingresos al realizar y/o cancelar tickets"
               Height          =   195
               Left            =   95
               TabIndex        =   88
               ToolTipText     =   "Utilizar cuenta puente en sustitución de cuentas de ingresos al realizar y/o cancelar tickets"
               Top             =   1040
               Width           =   7455
            End
            Begin VB.CheckBox chkCuentaPuenteBanco 
               Caption         =   "Utilizar cuenta puente en sustitución de cuentas de banco al cancelar y/o re facturar"
               Height          =   195
               Left            =   95
               TabIndex        =   87
               ToolTipText     =   "Utilizar cuenta puente en sustitución de cuentas de banco al cancelar y/o re facturar"
               Top             =   780
               Width           =   7455
            End
            Begin VB.CheckBox chkSelDeptoCargosDir 
               Caption         =   "Seleccionar departamento para ingresos por cargo directo"
               Height          =   195
               Left            =   95
               TabIndex        =   97
               ToolTipText     =   "Permitir seleccionar el departamento para ingresos por cargo directo"
               Top             =   3120
               Width           =   7455
            End
            Begin MSMask.MaskEdBox mskHoraIniMsgCargo 
               Height          =   315
               Left            =   4680
               TabIndex        =   102
               ToolTipText     =   "Hora en que se  muestra el primer recoratorio. hh:mm"
               Top             =   4395
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   " "
            End
            Begin VB.Label Label43 
               Caption         =   "Días máximo de espera de respuesta en los presupuestos"
               Height          =   285
               Left            =   90
               TabIndex        =   192
               Top             =   5730
               Width           =   5565
            End
            Begin VB.Label Label41 
               Caption         =   "Label41"
               Height          =   345
               Left            =   8460
               TabIndex        =   190
               Top             =   6690
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label Label29 
               Caption         =   "Días máximo para abrir cuentas de externos (usuarios con permiso de escritura)"
               Height          =   285
               Left            =   90
               TabIndex        =   158
               Top             =   5400
               Width           =   5805
            End
            Begin VB.Label Label28 
               Caption         =   "Días máximo para abrir cuentas de internos (usuarios con permiso de escritura)"
               Height          =   255
               Left            =   90
               TabIndex        =   157
               Top             =   5085
               Width           =   5805
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Título para el reporte de ""Cuentas pendientes de facturar"""
               Height          =   195
               Left            =   90
               TabIndex        =   156
               Top             =   6075
               Width           =   4125
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Mostrar aviso de pacientes sin cargo de cuarto cada "
               Height          =   195
               Left            =   90
               TabIndex        =   155
               Top             =   4755
               Width           =   3765
            End
            Begin VB.Label Label30 
               Caption         =   "Hora de inicio del recordatorio de pacientes sin cargo de cuarto "
               Height          =   255
               Left            =   90
               TabIndex        =   154
               Top             =   4440
               Width           =   4500
            End
            Begin VB.Label Label31 
               Caption         =   "minutos"
               Height          =   255
               Left            =   4740
               TabIndex        =   153
               Top             =   4755
               Width           =   600
            End
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Concepto de facturación para factura de asistencia social"
            Height          =   195
            Left            =   120
            TabIndex        =   189
            ToolTipText     =   "Concepto de facturación para factura de asistencia social"
            Top             =   1890
            Width           =   4080
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Uso del CFDI para honorarios médicos facturados"
            Height          =   195
            Left            =   120
            TabIndex        =   188
            Top             =   1560
            Width           =   3510
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Concepto de entrada de dinero para devolución de deudores diversos"
            Height          =   195
            Left            =   120
            TabIndex        =   159
            Top             =   1230
            Width           =   4950
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Concepto de facturación para honorarios médicos"
            Height          =   195
            Left            =   120
            TabIndex        =   151
            Top             =   900
            Width           =   3525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto de facturación parcial"
            Height          =   195
            Left            =   120
            TabIndex        =   137
            Top             =   570
            Width           =   2265
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Departamento encargado del cargo de cuartos"
            Height          =   195
            Left            =   120
            TabIndex        =   134
            Top             =   240
            Width           =   3315
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7605
         Left            =   -74925
         TabIndex        =   123
         Top             =   435
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   13414
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Formatos"
         TabPicture(0)   =   "frmParametros.frx":03CE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Usuarios impresora serial"
         TabPicture(1)   =   "frmParametros.frx":03EA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame9"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame9 
            Height          =   6300
            Left            =   -74880
            TabIndex        =   129
            Top             =   60
            Width           =   8880
            Begin VB.ListBox lstUsuariosAsignados 
               Height          =   2200
               Left            =   2280
               Sorted          =   -1  'True
               TabIndex        =   12
               ToolTipText     =   "Usuarios asignados"
               Top             =   3700
               Width           =   5040
            End
            Begin VB.ListBox lstListaUsuarios 
               Height          =   2200
               Left            =   2280
               Sorted          =   -1  'True
               TabIndex        =   9
               ToolTipText     =   "Usuarios"
               Top             =   520
               Width           =   5040
            End
            Begin VB.Frame freBotones 
               Height          =   720
               Left            =   4095
               TabIndex        =   130
               Top             =   2850
               Width           =   1395
               Begin VB.CommandButton cmdSelecciona 
                  Caption         =   "Excluir"
                  Enabled         =   0   'False
                  Height          =   510
                  Index           =   1
                  Left            =   700
                  MaskColor       =   &H80000014&
                  Picture         =   "frmParametros.frx":0406
                  Style           =   1  'Graphical
                  TabIndex        =   11
                  ToolTipText     =   "Excluir un usuario de la impresora serial."
                  Top             =   150
                  UseMaskColor    =   -1  'True
                  Width           =   630
               End
               Begin VB.CommandButton cmdSelecciona 
                  Caption         =   "Incluir"
                  Height          =   510
                  Index           =   0
                  Left            =   60
                  MaskColor       =   &H80000014&
                  Picture         =   "frmParametros.frx":08F8
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  ToolTipText     =   "Asignar un usuario a la impresora serial."
                  Top             =   150
                  UseMaskColor    =   -1  'True
                  Width           =   630
               End
            End
            Begin VB.Label Label27 
               Caption         =   "Usuarios asignados"
               Height          =   225
               Left            =   405
               TabIndex        =   132
               Top             =   3435
               Width           =   1575
            End
            Begin VB.Label Label26 
               Caption         =   "Lista de usuarios"
               Height          =   225
               Left            =   405
               TabIndex        =   131
               Top             =   255
               Width           =   2175
            End
         End
         Begin VB.Frame Frame3 
            Height          =   7020
            Left            =   360
            TabIndex        =   124
            Top             =   60
            Width           =   8400
            Begin VB.ComboBox cboImpresoraTickets 
               Height          =   315
               Left            =   2460
               Style           =   2  'Dropdown List
               TabIndex        =   2
               ToolTipText     =   "Selección de impresora para los tickets"
               Top             =   1320
               Width           =   5685
            End
            Begin VB.TextBox txtEditCol 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   240
               MaxLength       =   2
               TabIndex        =   150
               Top             =   6240
               Visible         =   0   'False
               Width           =   1365
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCopiasImpresion 
               Height          =   1335
               Left            =   195
               TabIndex        =   8
               Top             =   5520
               Width           =   7920
               _ExtentX        =   13970
               _ExtentY        =   2355
               _Version        =   393216
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.TextBox txtLeyenda3 
               Height          =   495
               Left            =   195
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   7
               ToolTipText     =   "Leyenda3 para factura en dólares"
               Top             =   4890
               Width           =   7920
            End
            Begin VB.TextBox txtLeyenda2 
               Height          =   495
               Left            =   195
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   6
               ToolTipText     =   "Leyenda2 para factura en dólares"
               Top             =   4320
               Width           =   7920
            End
            Begin VB.TextBox txtLeyenda1 
               Height          =   495
               Left            =   195
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   5
               ToolTipText     =   "Leyenda1 para pactura en dólares"
               Top             =   3720
               Width           =   7920
            End
            Begin VB.TextBox txtLeyendaDescuentos 
               Height          =   315
               Left            =   195
               MaxLength       =   20
               TabIndex        =   3
               ToolTipText     =   "Leyenda para descuentos"
               Top             =   2280
               Width           =   7920
            End
            Begin VB.ComboBox cboImpresoras 
               Height          =   315
               Left            =   2460
               Style           =   2  'Dropdown List
               TabIndex        =   1
               ToolTipText     =   "Selección de impresora para las facturas"
               Top             =   860
               Width           =   5685
            End
            Begin VB.TextBox txtLeyendaCliente 
               Height          =   315
               Left            =   195
               MaxLength       =   200
               TabIndex        =   4
               Top             =   3000
               Width           =   7920
            End
            Begin VB.ComboBox cboTickets 
               Height          =   315
               Left            =   2460
               Style           =   2  'Dropdown List
               TabIndex        =   0
               ToolTipText     =   "Selección del formato del ticket"
               Top             =   400
               Width           =   5685
            End
            Begin VB.Label Label42 
               Caption         =   "Impresora para los tickets"
               Height          =   195
               Left            =   195
               TabIndex        =   191
               Top             =   1360
               Width           =   1905
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Leyenda factura en dólares"
               Height          =   195
               Left            =   195
               TabIndex        =   135
               Top             =   3480
               Width           =   1935
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Impresora para las facturas"
               Height          =   195
               Left            =   195
               TabIndex        =   128
               Top             =   900
               Width           =   1905
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Leyenda para descuentos"
               Height          =   195
               Left            =   195
               TabIndex        =   127
               Top             =   2040
               Width           =   1845
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Leyenda de información al cliente"
               Height          =   195
               Left            =   195
               TabIndex        =   126
               Top             =   2760
               Width           =   2370
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Formato para ticket"
               Height          =   195
               Left            =   195
               TabIndex        =   125
               Top             =   440
               Width           =   1365
            End
         End
      End
      Begin TabDlg.SSTab sstVentaPublico 
         Height          =   7005
         Left            =   -74925
         TabIndex        =   117
         Top             =   435
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   12356
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos fiscales factura"
         TabPicture(0)   =   "frmParametros.frx":0DEA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Datos fiscales ticket"
         TabPicture(1)   =   "frmParametros.frx":0E06
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame8"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame8 
            Caption         =   "Departamentos que piden datos fiscales al generar ticket"
            Height          =   6200
            Left            =   -74205
            TabIndex        =   122
            Top             =   200
            Width           =   7400
            Begin VB.ListBox lstDepartamentos 
               Height          =   5460
               Left            =   300
               Style           =   1  'Checkbox
               TabIndex        =   23
               ToolTipText     =   "Seleccione los departamentos para que pidan datos fiscales al generar ticket de venta"
               Top             =   400
               Width           =   6800
            End
         End
         Begin VB.Frame Frame4 
            Height          =   6300
            Left            =   360
            TabIndex        =   118
            Top             =   120
            Width           =   8400
            Begin VB.TextBox txtNumExterior 
               Height          =   315
               Left            =   2160
               MaxLength       =   10
               TabIndex        =   16
               ToolTipText     =   "Número exterior para las facturas de venta al público"
               Top             =   1810
               Width           =   1995
            End
            Begin VB.TextBox txtNumInterior 
               Height          =   315
               Left            =   5760
               MaxLength       =   10
               TabIndex        =   17
               ToolTipText     =   "Número interior para las facturas de venta al público"
               Top             =   1810
               Width           =   1995
            End
            Begin VB.TextBox txtCPPOS 
               Height          =   315
               Left            =   2160
               TabIndex        =   19
               ToolTipText     =   "Código postal para las facturas de venta al público"
               Top             =   2710
               Width           =   1995
            End
            Begin VB.TextBox txtColoniaPOS 
               Height          =   315
               Left            =   2160
               TabIndex        =   18
               ToolTipText     =   "Colonia para las facturas de venta al público"
               Top             =   2240
               Width           =   5600
            End
            Begin VB.TextBox txtNombreFactura 
               Height          =   315
               Left            =   2160
               MaxLength       =   300
               TabIndex        =   13
               ToolTipText     =   "Nombre para las facturas de venta al público"
               Top             =   400
               Width           =   5600
            End
            Begin VB.TextBox txtRFCFactura 
               Height          =   315
               Left            =   2160
               TabIndex        =   14
               ToolTipText     =   "RFC para las facturas de venta al público"
               Top             =   870
               Width           =   2000
            End
            Begin VB.TextBox txtDireccionPOS 
               Height          =   315
               Left            =   2160
               TabIndex        =   15
               ToolTipText     =   "Calle para la factura de venta al público"
               Top             =   1340
               Width           =   5600
            End
            Begin VB.ComboBox cboCiudad 
               Height          =   315
               Left            =   2160
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   20
               ToolTipText     =   "Ciudad para las facturas de venta al público"
               Top             =   3180
               Width           =   5600
            End
            Begin VB.ComboBox cboTipoMedico 
               Height          =   315
               Left            =   2160
               Style           =   2  'Dropdown List
               TabIndex        =   21
               ToolTipText     =   "Tipo de paciente para los médicos en la venta al público"
               Top             =   3645
               Width           =   5595
            End
            Begin VB.ComboBox cboTipoEmpleado 
               Height          =   315
               Left            =   2160
               Style           =   2  'Dropdown List
               TabIndex        =   22
               ToolTipText     =   "Tipo de paciente para los empleados en la venta al público"
               Top             =   4110
               Width           =   5600
            End
            Begin VB.Label Label33 
               Caption         =   "Número exterior"
               Height          =   255
               Left            =   480
               TabIndex        =   146
               Top             =   1845
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "Número interior"
               Height          =   255
               Left            =   4455
               TabIndex        =   145
               Top             =   1845
               Width           =   1200
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Código postal"
               Height          =   195
               Left            =   480
               TabIndex        =   144
               Top             =   2745
               Width           =   960
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Colonia"
               Height          =   195
               Left            =   480
               TabIndex        =   142
               Top             =   2280
               Width           =   525
            End
            Begin VB.Label Label12 
               Caption         =   "Tipo paciente para cuenta de empleados"
               Height          =   405
               Left            =   480
               TabIndex        =   141
               Top             =   4065
               Width           =   1605
            End
            Begin VB.Label Label9 
               Caption         =   "Tipo paciente para cuenta de médicos"
               Height          =   405
               Left            =   480
               TabIndex        =   140
               Top             =   3600
               Width           =   1605
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Ciudad"
               Height          =   255
               Left            =   480
               TabIndex        =   138
               Top             =   3225
               Width           =   1605
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Calle"
               Height          =   195
               Left            =   480
               TabIndex        =   121
               Top             =   1380
               Width           =   345
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "RFC"
               Height          =   255
               Left            =   480
               TabIndex        =   120
               Top             =   915
               Width           =   1605
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Facturar a"
               Height          =   255
               Left            =   480
               TabIndex        =   119
               Top             =   435
               Width           =   1605
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7000
         Left            =   -74910
         TabIndex        =   112
         Top             =   345
         Width           =   9165
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Este tab se muestra invisible"
            Height          =   195
            Left            =   1425
            TabIndex        =   136
            Top             =   1875
            Width           =   1995
         End
      End
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Dirección"
      Height          =   255
      Left            =   2040
      TabIndex        =   143
      Top             =   2805
      Width           =   1605
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Este programa necesita que se encuentre registrado el parámetro general que
' es la estructura de la cuenta contable que se está usando actualmente.
'-----------------------------------------------------------------------------

Option Explicit

Const cintNumTipoFormato = 2

Dim vlstrsql As String
Dim vlstrMascara As String
Dim vllngNumeroCuenta As Long
Dim rsPvParametro As New ADODB.Recordset
Dim rsSiParametro As New ADODB.Recordset  'Concepto de facturación para factura de asistencia social
Dim vgblnNoEditar As Boolean
Dim vlblnEscTxtEditCOl As Boolean
Dim blLicenciaIEPS As Boolean
Dim lblnCargandoParametros As Boolean

Private Sub cboCiudad_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cboTipoMedico.SetFocus
    End If
    
End Sub

Private Sub cboConceptoCoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
End Sub

Private Sub cboConceptoCoaseguroAdicional_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub cboConceptoCoaseguroMedico_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
End Sub

Private Sub cboConceptoCopago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboConceptoDeducible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboConceptoEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboConceptoFacturacionAsistSocial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub


Private Sub cboConceptoParcial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboConceptoHonorariosMedicos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboDepartamentoMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboExcedenteIVA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboImpresoras_GotFocus()
    pEnfocaCbo cboImpresoras
End Sub

Private Sub cboImpresoraTickets_GotFocus()
    pEnfocaCbo cboImpresoraTickets
End Sub

Private Sub cboImpresoraTickets_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtLeyendaDescuentos.SetFocus
    End If
End Sub

Private Sub cboTickets_GotFocus()
    pEnfocaCbo cboTickets
End Sub

Private Sub cboTickets_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cboImpresoras.SetFocus
    End If

End Sub

Private Sub cboTipoEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        cmdSave.SetFocus
    End If
    
End Sub

Private Sub cboTipoMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        cboTipoEmpleado.SetFocus
    End If
    
End Sub

Private Sub cboUsoCFDIFacturado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub
Private Sub chkAbrirCuentaExterna_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkCapturarMargenSubrogado.SetFocus
    End If
End Sub

Private Sub ChkAuditoriaCargos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        chkAbrirCuentaExterna.SetFocus
    End If
End Sub

Private Sub chkCalcularCargosSeleccionados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkCancelarRecibosOtroDepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkCerrarCuentas.SetFocus
    End If
End Sub

Private Sub chkCapturarMargenSubrogado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkCerrarCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkCerrarCuentasAut_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkConservarPrecioDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub Chkcuentacerrada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkPermitirFacturarCargosFueraCatalogo.SetFocus
    End If
End Sub

Private Sub chkCuentaPuenteBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub


Private Sub chkCuentaPuenteIngresos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub


Private Sub chkDesactivarExterno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkDesglosarCoaseguroM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
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

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
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
Private Sub chkdesgloseIEPS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkDetCS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkFacturaAutomatica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkCorteHonorarios.SetFocus
    End If
End Sub

Private Sub chkHonorarioMedicoCredito_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkPermitirCorte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkPermitirFacturarCargosFueraCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkRequisicionespendientes.SetFocus
    End If
End Sub

Private Sub chkRequisicionespendientes_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
End Sub
Private Sub chkCorteHonorarios_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
End Sub

Private Sub chkDesglosarCoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarCoaseguroAdicional_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarCopago_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarDeducible_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarExcedente_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarIVACoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarIVACoaseguroAdicional_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarIVACopago_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarIVADeducible_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkDesglosarIVAExcedente_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub chkSelDeptoCargosDir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkTrasladarCargos.SetFocus
    End If
End Sub

Private Sub chkSocios_Click()
    If chkSocios.Value = 1 Then
        chkValidaDoble.Enabled = True
    Else
        chkValidaDoble.Enabled = False
    End If
End Sub

Private Sub chkSocios_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If chkValidaDoble.Enabled Then
            chkValidaDoble.SetFocus
        Else
            chkSelDeptoCargosDir.SetFocus
        End If
    End If
End Sub

Private Sub chkTotCS_Click()
    If chkTotCS.Value = vbChecked Then
        optTipoDesglose(0).Enabled = True
        optTipoDesglose(1).Enabled = True
    Else
        optTipoDesglose(0).Enabled = False
        optTipoDesglose(1).Enabled = False
    End If
End Sub

Private Sub chkTotCS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkTrasladarCargos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ChkAuditoriaCargos.SetFocus
    End If
End Sub

Private Sub chkValidaDoble_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkSelDeptoCargosDir.SetFocus
    End If
End Sub

Private Sub chkVerificar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkSocios.SetFocus
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    Dim rsNumeroRegistros As New ADODB.Recordset
    Dim X As Integer
    Dim vllngPersonaGraba As Long
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim vllngNumeroCuenta As Long
    Dim vlintErrorCuenta As Integer
    Dim arrCompara(7) As Long
    Dim inti As Long
    Dim intj As Long
    Dim intDias As Integer
    
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 1187, 336), "E", True) Then

        If fstrVerificaHora(mskHoraIniMsgCargo) = "" Then
            MsgBox SIHOMsg(41), vbOKOnly + vbExclamation, "Mensaje"
            mskHoraIniMsgCargo.SetFocus
            pSelMkTexto mskHoraIniMsgCargo
            Exit Sub
        End If
        
        If CInt(txtIntervaloMsgCargo.Text) > 60 Then
           MsgBox SIHOMsg(767), vbOKOnly + vbExclamation, "Mensaje"
           txtIntervaloMsgCargo.SetFocus
           pSelTextBox txtIntervaloMsgCargo
        End If
        
        'Si se deshabilitó el uso de socio se realizan las siguientes validaciones
        If chkSocios.Value = vbUnchecked Then
            'Verifica si hay pacientes de tipo socio pendientes de facturar
            Set rs = frsRegresaRs("Select count(*) as cant from pvcargo pvc inner join expacienteingreso ex inner join sitipoingreso si" & _
                                            " on si.INTCVETIPOINGRESO = ex.INTCVETIPOINGRESO inner join siparametro sip on to_char(ex.INTCVETIPOPACIENTE) = sip.VCHVALOR and sip.VCHNOMBRE = 'INTCVETIPOPACIENTESOCIO'" & _
                                            " on ex.INTNUMCUENTA = pvc.INTMOVPACIENTE and trim(pvc.CHRTIPOPACIENTE) = trim(si.CHRTIPOINGRESO) Where pvc.CHRFOLIOFACTURA Is Null")
            If rs.RecordCount > 0 Then
                If Val(rs!cant) > 0 Then
                    'No se puede deshabilitar la administración de socios, existen cuentas de pacientes pendientes de facturar
                        MsgBox SIHOMsg(1130), vbExclamation, "Mensaje"
                        Exit Sub
                End If
            End If
        End If
                
        ' Persona que graba
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
    
        If cboConceptoSumaAsegurada.ListIndex = -1 Then
            arrCompara(0) = 0
        Else
            arrCompara(0) = cboConceptoSumaAsegurada.ItemData(cboConceptoSumaAsegurada.ListIndex)
        End If
        If cboConceptoDeducible.ListIndex = -1 Then
            arrCompara(1) = 0
        Else
            arrCompara(1) = cboConceptoDeducible.ItemData(cboConceptoDeducible.ListIndex)
        End If
        If cboConceptoCoaseguro.ListIndex = -1 Then
            arrCompara(2) = 0
        Else
            arrCompara(2) = cboConceptoCoaseguro.ItemData(cboConceptoCoaseguro.ListIndex)
        End If
        If cboConceptoCoaseguroAdicional.ListIndex = -1 Then
            arrCompara(3) = 0
        Else
            arrCompara(3) = cboConceptoCoaseguroAdicional.ItemData(cboConceptoCoaseguroAdicional.ListIndex)
        End If
        If cboConceptoCopago.ListIndex = -1 Then
            arrCompara(4) = 0
        Else
            arrCompara(4) = cboConceptoCopago.ItemData(cboConceptoCopago.ListIndex)
        End If
        If cboConceptoCoaseguroMedico.ListIndex = -1 Then
            arrCompara(5) = 0
        Else
            arrCompara(5) = cboConceptoCoaseguroMedico.ItemData(cboConceptoCoaseguroMedico.ListIndex)
        End If
        If cboExcedenteIVA.ListIndex = -1 Then
            arrCompara(6) = 0
        Else
            arrCompara(6) = cboExcedenteIVA.ItemData(cboExcedenteIVA.ListIndex)
        End If
        
        For inti = 0 To 6
            For intj = inti + 1 To 6
                If arrCompara(inti) = arrCompara(intj) And arrCompara(inti) <> 0 Then
                    'No se debe asignar el mismo concepto de facturación a varios conceptos de seguro
                    MsgBox SIHOMsg(1005), vbOKOnly + vbInformation, "Mensaje"
                    Exit Sub
                End If
            Next intj
        Next inti
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        With rsPvParametro
            
            vlstrsql = "Select Count(*) From PvParametro Where tnyClaveEmpresa = " & vgintClaveEmpresaContable
            Set rsNumeroRegistros = frsRegresaRs(vlstrsql)
            
            If rsNumeroRegistros.Fields(0) = 0 Then
                .AddNew
            Else
                .MoveFirst
            End If
            
            !tnyclaveempresa = vgintClaveEmpresaContable
            
            'Impresión
            !vchLeyendaDescuento = txtLeyendaDescuentos.Text
            !vchleyendacliente = txtLeyendaCliente.Text
            !vchLeyenda1 = txtLeyenda1.Text
            !vchLeyenda2 = txtLeyenda2.Text
            !vchLeyenda3 = txtLeyenda3.Text
            
            'Venta al público
            !chrNombreFacturaPOS = Trim(txtNombreFactura.Text)
            !CHRRFCPOS = Trim(txtRFCFactura.Text)
            !chrDireccionPOS = Trim(txtDireccionPOS.Text)
            !vchNumeroExteriorPOS = Trim(txtNumExterior.Text)
            !vchNumeroInteriorPOS = Trim(txtNumInterior.Text)
            !vchColoniaPOS = Trim(txtColoniaPOS.Text)
            !vchCodigoPostalPOS = Trim(txtCPPOS.Text)
            If cboCiudad.ListIndex = -1 Then
                !intCveCiudad = 0
            Else
                !intCveCiudad = cboCiudad.ItemData(cboCiudad.ListIndex)
            End If

            If cboTipoMedico.ListIndex > -1 Then
              !intTipoPacMedico = cboTipoMedico.ItemData(cboTipoMedico.ListIndex)
            End If
            If cboTipoEmpleado.ListIndex > -1 Then
              !intTipoPacEmpleado = cboTipoEmpleado.ItemData(cboTipoEmpleado.ListIndex)
            End If
            
            'Aseguradoras
            !BITCALCULARENBASEACARGOS = chkCalcularCargosSeleccionados.Value
            
            'Excedente
            If cboConceptoSumaAsegurada.ListIndex = -1 Then
                !INTCONCEPTOSUMAASEGURADA = 0
            Else
                !INTCONCEPTOSUMAASEGURADA = cboConceptoSumaAsegurada.ItemData(cboConceptoSumaAsegurada.ListIndex)
            End If
            !INTDESGLOSARIVAEXCEDENTE = chkDesglosarIVAExcedente.Value
            !INTDESGLOSAREXCEDENTE = chkDesglosarExcedente.Value
                        
            'Deducible
            If cboConceptoDeducible.ListIndex = -1 Then
                !INTCONCEPTODEDUCIBLE = 0
            Else
                !INTCONCEPTODEDUCIBLE = cboConceptoDeducible.ItemData(cboConceptoDeducible.ListIndex)
            End If
            !INTDESGLOSARIVADEDUCIBLE = chkDesglosarIVADeducible.Value
            !INTDESGLOSARDEDUCIBLE = chkDesglosarDeducible.Value
                        
            'Coaseguro
            If cboConceptoCoaseguro.ListIndex = -1 Then
                !INTCONCEPTOCOASEGURO = 0
            Else
                !INTCONCEPTOCOASEGURO = cboConceptoCoaseguro.ItemData(cboConceptoCoaseguro.ListIndex)
            End If
            !INTDESGLOSARIVACOASEGURO = chkDesglosarIVACoaseguro.Value
            !INTDESGLOSARCOASEGURO = chkDesglosarCoaseguro.Value
                        
            'Coaseguro médico
            If cboConceptoCoaseguroMedico.ListIndex = -1 Then
                !intConceptoCoaseguroMedico = 0
            Else
                !intConceptoCoaseguroMedico = cboConceptoCoaseguroMedico.ItemData(cboConceptoCoaseguroMedico.ListIndex)
            End If
            !INTDESGLOSARIVACOASEGUROMEDICO = chkDesglosarIVACoaseguroM.Value
            !INTDESGLOSARCOASEGUROMEDICO = chkDesglosarCoaseguroM.Value
            
            'Coaseguro adicional
            If cboConceptoCoaseguroAdicional.ListIndex = -1 Then
                !INTCONCEPTOCOASEGUROADICIONAL = 0
            Else
                !INTCONCEPTOCOASEGUROADICIONAL = cboConceptoCoaseguroAdicional.ItemData(cboConceptoCoaseguroAdicional.ListIndex)
            End If
            !INTDESGLOSARIVACOASEGUROADICIO = chkDesglosarIVACoaseguroAdicional.Value
            !INTDESGLOSARCOASEGUROADICIONAL = chkDesglosarCoaseguroAdicional.Value
            
            'Copago
            If cboConceptoCopago.ListIndex = -1 Then
                !INTCONCEPTOCOPAGO = 0
            Else
                !INTCONCEPTOCOPAGO = cboConceptoCopago.ItemData(cboConceptoCopago.ListIndex)
            End If
            !INTDESGLOSARIVACOPAGO = chkDesglosarIVACopago.Value
            !INTDESGLOSARCOPAGO = chkDesglosarCopago.Value
            
            'Excedente de IVA
            If cboExcedenteIVA.ListIndex = -1 Then
                !intExcedenteIVA = 0
            Else
                !intExcedenteIVA = cboExcedenteIVA.ItemData(cboExcedenteIVA.ListIndex)
            End If
            
            'Generales
            If cboDepartamentoMsg.ListIndex < 0 Then
                !smiCveDepartamentoMsg = 0
            Else
                !smiCveDepartamentoMsg = cboDepartamentoMsg.ItemData(cboDepartamentoMsg.ListIndex)
            End If
            If cboConceptoParcial.ListIndex = -1 Then
                !intCveConceptoParcial = 0
            Else
                !intCveConceptoParcial = cboConceptoParcial.ItemData(cboConceptoParcial.ListIndex)
            End If
            
            If cboConceptoHonorariosMedicos.ListIndex = -1 Then
                !intCveConceptoHonorarioMedico = 0
            Else
                !intCveConceptoHonorarioMedico = cboConceptoHonorariosMedicos.ItemData(cboConceptoHonorariosMedicos.ListIndex)
            End If
            
            If cboUsoCFDIFacturado.ListIndex = -1 Then
                !INTCVEUSOCFDIHONOFACTURADO = 0
            Else
                !INTCVEUSOCFDIHONOFACTURADO = cboUsoCFDIFacturado.ItemData(cboUsoCFDIFacturado.ListIndex)
            End If
            
            'If cboConceptoFacturacionAsistSocial.ListIndex = -1 Then
            '    !INTCVECONCEPTOFACTASISTSOCIAL = 0
            'Else
            '    !INTCVECONCEPTOFACTASISTSOCIAL = cboConceptoFacturacionAsistSocial.ItemData(cboConceptoFacturacionAsistSocial.ListIndex)
            'End If
            
            !bitCerrarCuentasExtAut = IIf(chkCerrarCuentasAut.Value = vbUnchecked, 0, 1)
            !BITDESACTIVARALFACTURAR = IIf(chkDesactivarExterno.Value = vbUnchecked, 0, 1)
            !bitFacturarVentaPublico = chkFacturaAutomatica.Value
            !bitIncluyeHonorarioCorte = chkCorteHonorarios.Value
            !bitCuentaCerrada = Chkcuentacerrada.Value
            !bitfacturarconreqpendiente = chkRequisicionespendientes.Value
            !BITVERIFICARREQUISICIONES = chkVerificar.Value
            !dtmHoraIniMsgCargo = CDate(mskHoraIniMsgCargo.Text)
            !intIntervaloMsgCargo = IIf(txtIntervaloMsgCargo.Text = "", 0, CInt(txtIntervaloMsgCargo.Text))
            !intDiasAbrirCuentasInternos = txtDiasAbrirCuentaInt.Text
            !intDiasAbrirCuentasExternos = txtDiasAbrirCuentaExt.Text
            !vchTituloCtasPendFact = Trim(txtTituloCtasPendFact.Text)
            !BITDESGLOSEIEPSTICKET = IIf(Me.chkdesgloseIEPS.Enabled = True, IIf(Me.chkdesgloseIEPS.Value = vbChecked, 1, 0), 0) '<---------IEPS
            !bitUtilizaCuentaPuenteBanco = chkCuentaPuenteBanco.Value
            !bitUtilizaCuentaPuenteIngresos = chkCuentaPuenteIngresos.Value
            !bitTrasladaCargos = chkTrasladarCargos.Value
            !bitCancelarRecibosOtroDepto = chkCancelarRecibosOtroDepto.Value
            !IntDiasSinRespPresupuesto = txtDiasSinRespPresupuesto.Text
            !BITCAPTURAMARGENSUBROGADO = IIf(chkCapturarMargenSubrogado.Value = vbUnchecked, 0, 1)
            !BitValidacionPMPVentaPublic = IIf(chkValidacionPMPVentaPublico.Value = vbUnchecked, 0, 1)
            .Update
        End With
        
        'caso 19900
        pEjecutaSentencia "Update siparametro set vchvalor = " & ChkAuditoriaCargos.Value & " WHERE VCHNOMBRE = 'BITAUDITORIADECARGOS'"
            
        
        '-----------------
        '««  Impresión  »»
        '-----------------
        vlstrsql = "Delete From PvTicketDepartamento Where smiDepartamento=" + Str(vgintNumeroDepartamento)
        pEjecutaSentencia vlstrsql
        
        If cboTickets.ListIndex > -1 Then
            vlstrsql = "Insert Into PvTicketDepartamento values (" & CStr(vgintNumeroDepartamento) & ", " & CStr(cboTickets.ItemData(cboTickets.ListIndex)) & ")"
            pEjecutaSentencia vlstrsql
        End If
        
        vlstrsql = "Delete From ImpresoraDepartamento where chrTipo = 'FA' and smiCveDepartamento=" + Str(vgintNumeroDepartamento)
        pEjecutaSentencia vlstrsql
        
        vlstrsql = "Delete From ImpresoraDepartamento where chrTipo = 'TI' and smiCveDepartamento=" + Str(vgintNumeroDepartamento)
        pEjecutaSentencia vlstrsql
        
        If cboImpresoras.ListIndex > -1 Then
            vlstrsql = "Insert into ImpresoraDepartamento values (" + Str(vgintNumeroDepartamento) + ",'" + UCase(Trim(cboImpresoras.List(cboImpresoras.ListIndex))) + "','FA')"
            pEjecutaSentencia vlstrsql
        End If
        
        If cboImpresoraTickets.ListIndex > -1 Then
            vlstrsql = "Insert into ImpresoraDepartamento values (" + Str(vgintNumeroDepartamento) + ",'" + UCase(Trim(cboImpresoraTickets.List(cboImpresoraTickets.ListIndex))) + "','TI')"
            pEjecutaSentencia vlstrsql
        End If
        
        vlstrsql = "Delete From PvLoginImpresoraTicket"
        pEjecutaSentencia vlstrsql
        
        For vlintContador = 0 To lstUsuariosAsignados.ListCount - 1
            vlstrsql = "insert into PvLoginImpresoraTicket values (" + Trim(Str(lstUsuariosAsignados.ItemData(vlintContador))) + ")"
            pEjecutaSentencia vlstrsql
        Next vlintContador
        
        '---------------------------
        '««  Copias de Impresión  »»
        '---------------------------
        pEjecutaSentencia "Update siparametro set vchvalor ='" & Me.grdCopiasImpresion.TextMatrix(1, 1) & "' where intcveempresacontable = " & vgintClaveEmpresaContable & " and vchnombre = 'NUMCOPIASFACTURAPACIENTE'"
        pEjecutaSentencia "Update siparametro set vchvalor ='" & Me.grdCopiasImpresion.TextMatrix(2, 1) & "' where intcveempresacontable = " & vgintClaveEmpresaContable & " and vchnombre = 'NUMCOPIASFACTURAEMPRESA'"
        pEjecutaSentencia "Update siparametro set vchvalor ='" & Me.grdCopiasImpresion.TextMatrix(3, 1) & "' where intcveempresacontable = " & vgintClaveEmpresaContable & " and vchnombre = 'NUMCOPIASFACTURADIRECTA'"
        pEjecutaSentencia "Update siparametro set vchvalor ='" & Me.grdCopiasImpresion.TextMatrix(4, 1) & "' where intcveempresacontable = " & vgintClaveEmpresaContable & " and vchnombre = 'NUMCOPIASTICKET'"
        '---------------------
        '««  Venta público  »»
        '---------------------
        vlstrSentencia = "Delete from PvDatosFiscalesDepartamento "
        pEjecutaSentencia vlstrSentencia
        
        For vlintContador = 0 To lstDepartamentos.ListCount - 1
            If lstDepartamentos.Selected(vlintContador) Then
                vlstrSentencia = "insert into PvDatosFiscalesDepartamento (SMIDEPARTAMENTO, CHRESTATUS) values(" & _
                                    Trim(Str(lstDepartamentos.ItemData(vlintContador))) & "," & _
                                    IIf(lstDepartamentos.Selected(vlintContador), "'T'", "'F'") & ")"
                pEjecutaSentencia vlstrSentencia
            End If
        Next
        
        '--------------------------------
        '««  Administración de socios  »»
        '--------------------------------
        vlstrsql = "select vchvalor, VCHSENTENCIA from siparametro where vchnombre like 'BITUTILIZASOCIOS'"
        Set rs = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        rs!VCHVALOR = chkSocios.Value
        rs!vchSentencia = chkValidaDoble.Value
        rs.Update
        rs.Close
        
        'Si se deshabilitó el uso de socio se elimina el paciente configurado de tipo socio
        If chkSocios.Value = vbUnchecked Then
            vlstrSentencia = "Update SIPARAMETRO set vchValor = '0' Where vchNombre = 'INTCVETIPOPACIENTESOCIO'"
            pEjecutaSentencia vlstrSentencia
        End If
        
        
        '|---------------------------------------------------------
        '|  Parámetros registrados en SiParametro
        '|---------------------------------------------------------
        pActualizaParametro vgintClaveEmpresaContable, "BITFACTURARCARGOSFUERACATALOGO", "PV", chkPermitirFacturarCargosFueraCatalogo.Value, "Indica si se permite facturar con cargos que no existan en el catálogo de cargos por empresa"
        pActualizaParametro vgintClaveEmpresaContable, "BITCERRARCUENTAREQUIPENDIENTES", "PV", chkCerrarCuentas.Value, "Indica si se permite cerrar cuentas con requisiciones pendientes de surtir"
        pActualizaParametro vgintClaveEmpresaContable, "BITCUENTAPORPAGARHONORARIOMEDICOCREDITO", "PV", chkHonorarioMedicoCredito.Value, "Indica si se generará la cuenta por pagar al médico por honorario médico al facturar a crédito, 1 = Si genera cuenta por pagar, 0 = No genera cuenta por pagar"
                
        '----------------------------------------------------------------------------------------------------------------
        'Parámetro que indica si se deben conservar los costos y descuentos de los cargos después de hacer un a exclusión
        '----------------------------------------------------------------------------------------------------------------
        pActualizaParametro vgintClaveEmpresaContable, "BITCONSERVARCOSTOSDESCUENTOEXCLUSION", "PV", Me.chkConservarPrecioDescuento.Value, "Indica si se deben conservar los costos y los descuentos de los cargos al momento de realizar una exlcusión"
        'Seleccionar departamento en cargos directos
        pActualizaParametro vgintClaveEmpresaContable, "BITSELECCIONARDEPTOCARGODIRECTO", "PV", chkSelDeptoCargosDir.Value, "Seleccionar departamento para ingresos por cargo directo"
        'Permitir realizar corte de caja chica si existen salidas sin XML relacionados
        pActualizaParametro vgintClaveEmpresaContable, "BITCORTECAJACHICASINXML", "PV", chkPermitirCorte.Value, "Indica si se permitirá realizar el corte de caja chica si existen salidas sin XML relacionados"
        
        If cboConceptoEntrada.ListIndex = -1 Then
            pActualizaParametro vgintClaveEmpresaContable, "INTCONCEPTOENTRADACOMPROBACION", "PV", Null, "Indica el concepto de entrada de dinero usado para comprobación de gastos en entradas y salidas de dinero"
        Else
            pActualizaParametro vgintClaveEmpresaContable, "INTCONCEPTOENTRADACOMPROBACION", "PV", cboConceptoEntrada.ItemData(cboConceptoEntrada.ListIndex), "Indica el concepto de entrada de dinero usado para comprobación de gastos en entradas y salidas de dinero"
        End If
        
        If cboConceptoFacturacionAsistSocial.ListIndex = -1 Then
            pActualizaParametro vgintClaveEmpresaContable, "INTCVECONCEPTOFACTASISTSOCIAL", "PV", 0, "Concepto de facturación para factura de asistencia social"
        Else
            pActualizaParametro vgintClaveEmpresaContable, "INTCVECONCEPTOFACTASISTSOCIAL", "PV", cboConceptoFacturacionAsistSocial.ItemData(cboConceptoFacturacionAsistSocial.ListIndex), "Concepto de facturación para factura de asistencia social"
        End If
        'Abrir cuenta externa
        pActualizaParametro vgintClaveEmpresaContable, "BITGENERAPACIENTEEXTERNO", "PV", chkAbrirCuentaExterna.Value, "Indica si abrirá cuenta de externo para medicamento no aplicado en traslado de cargos"
        
        pActualizaParametro vgintClaveEmpresaContable, "CHRINCLUIRCFDICONCEPTOSSEGURO", "PV", IIf(optIncCS(0).Value, "I", "D"), "Indica cómo se van a incluir los conceptos de seguro en el CFDI de la aseguradora: I = Restado a los importes, D = Sumado a los descuentos"
        pActualizaParametro vgintClaveEmpresaContable, "INTTOTIMPCFDICONCEPTOSSEGURO", "PV", IIf(chkTotCS.Value = vbChecked, IIf(optTipoDesglose(0), "1", "2"), "0"), "Indica si se van a desglosar los descuentos por conceptos de seguro en los totales de la representación impresa del CFDI"
        
        'Descuentos por notas de crédito
        pActualizaParametro vgintClaveEmpresaContable, "VCHTIPODESCTONOTAEXCEDENTE", "PV", IIf(optTipoDesctoExcedente(0).Value = True, "P", "C"), "Indica si el descuento por notas de crédito para el excedente es P = por porcentaje, C = por cantidad"
        pActualizaParametro vgintClaveEmpresaContable, "VCHTIPODESCTONOTADEDUCIBLE", "PV", IIf(optTipoDesctoDeducible(0).Value = True, "P", "C"), "Indica si el descuento por notas de crédito para el deducible es P = por porcentaje, C = por cantidad"
        pActualizaParametro vgintClaveEmpresaContable, "VCHTIPODESCTONOTACOASEGURO", "PV", IIf(optTipoDesctoCoaseguro(0).Value = True, "P", "C"), "Indica si el descuento por notas de crédito para el coaseguro es P = por porcentaje, C = por cantidad"
        pActualizaParametro vgintClaveEmpresaContable, "VCHTIPODESCTONOTACOASEGUROMEDICO", "PV", IIf(optTipoDesctoCoaseguroMedico(0).Value = True, "P", "C"), "Indica si el descuento por notas de crédito para el coaseguro médico es P = por porcentaje, C = por cantidad"
        pActualizaParametro vgintClaveEmpresaContable, "VCHTIPODESCTONOTACOASEGUROADICIONAL", "PV", IIf(optTipoDesctoCoaseguroAdicional(0).Value = True, "P", "C"), "Indica si el descuento por notas de crédito para el coaseguro adicional es P = por porcentaje, C = por cantidad"
        pActualizaParametro vgintClaveEmpresaContable, "VCHTIPODESCTONOTACOPAGO", "PV", IIf(optTipoDesctoCopago(0).Value = True, "P", "C"), "Indica si el descuento por notas de crédito para el copago es P = por porcentaje, C = por cantidad"
        
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTAEXCEDENTE", "PV", Format(txtPorcentajeExcedentePorNota.Text, ""), "Indica el descuento por notas de crédito para el excedente, puede ser porcentaje o cantidad dependiendo del parámetro VCHTIPODESCTONOTAEXCEDENTE"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTADEDUCIBLE", "PV", Format(txtPorcentajeDeduciblePorNota.Text, ""), "Indica el descuento por notas de crédito para el deducible, puede ser porcentaje o cantidad dependiendo del parámetro VCHTIPODESCTONOTADEDUCIBLE"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTACOASEGURO", "PV", Format(txtPorcentajeCoaseguroPorNota.Text, ""), "Indica el descuento por notas de crédito para el coaseguro, puede ser porcentaje o cantidad dependiendo del parámetro VCHTIPODESCTONOTACOASEGURO"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTACOASEGUROMEDICO", "PV", Format(txtPorcentajeCoaseguroMPorNota.Text, ""), "Indica el descuento por notas de crédito para el coaseguro médico, puede ser porcentaje o cantidad dependiendo del parámetro VCHTIPODESCTONOTACOASEGUROMEDICO"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTACOASEGUROADICIONAL", "PV", Format(txtPorcentajeCoasAdicionalPorNota.Text, ""), "Indica el descuento por notas de crédito para el coaseguro adicional, puede ser porcentaje o cantidad dependiendo del parámetro VCHTIPODESCTONOTACOASEGUROADICIONAL"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTACOPAGO", "PV", Format(txtPorcentajeCopagoPorNota.Text, ""), "Indica el descuento por notas de crédito para el copago, puede ser porcentaje o cantidad dependiendo del parámetro VCHTIPODESCTONOTACOPAGO"
        
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTALIMITEEXCEDENTE", "PV", Format(txtCantidadLimiteExcedente.Text, ""), "Indica la cantidad límite por notas de crédito para excedente"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTALIMITEDEDUCIBLE", "PV", Format(txtCantidadLimiteDeducible.Text, ""), "Indica la cantidad límite por notas de crédito para deducible"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTALIMITECOASEGURO", "PV", Format(txtCantidadLimiteCoaseguro.Text, ""), "Indica la cantidad límite por notas de crédito para coaseguro"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTALIMITEMEDICO", "PV", Format(txtCantidadLimiteCoaseguroM.Text, ""), "Indica la cantidad límite por notas de crédito para coaseguro médico"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTALIMITEADICIONAL", "PV", Format(txtCantidadLimiteCoasAdicional.Text, ""), "Indica la cantidad límite por notas de crédito para coaseguro adicional"
        pActualizaParametro vgintClaveEmpresaContable, "NUMDESCTONOTALIMITECOPAGO", "PV", Format(txtCantidadLimiteCopago.Text, ""), "Indica la cantidad límite por notas de crédito para copago"

         
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "PARAMETROS DE CAJA", CStr(vgintNumeroDepartamento))
        
        If chkCerrarCuentasAut = 1 Then
            Set rs = frsRegresaRs("select bitCerrarCuentasExtAut, intDiasAbrirCuentasExternos from PVParametro where tnyclaveempresa = " & vgintClaveEmpresaContable)
            If Not rs.EOF Then
                If Not IsNull(rs!bitCerrarCuentasExtAut) And Not IsNull(rs!intDiasAbrirCuentasExternos) Then
                    If rs!bitCerrarCuentasExtAut <> 0 Then
                        intDias = rs!intDiasAbrirCuentasExternos
                        frsEjecuta_SP intDias & "|" & vglngNumeroLogin, "sp_PVCierreAutomaticoCuentas", True
                    End If
                End If
            End If
        End If
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
        
        Unload Me
        
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError
    
    If Trim(vgstrEstructuraCuentaContable) = "" Then
        'No se encuentra registrado el parámetro de estructura de la cuenta contable.
        MsgBox SIHOMsg(260), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If Me.ActiveControl.Name = "txtEditCol" Then
        Exit Sub
    End If
  
    If vlblnEscTxtEditCOl = True Then
        vlblnEscTxtEditCOl = False
        KeyAscii = 0
        Exit Sub
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
    Dim rs As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon

    vlstrsql = "select * from PvParametro where tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rsPvParametro = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    'Concepto de facturación para factura de asistencia social
    vlstrsql = "select vchvalor from SiParametro where vchnombre like 'INTCVECONCEPTOFACTASISTSOCIAL' and intcveempresacontable = " & vgintClaveEmpresaContable
    Set rsSiParametro = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    
    'Cargar las ciudades:
    vgstrParametrosSP = "-1|-1|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELCIUDAD")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboCiudad, rs, 0, 1
    End If
    
    pCargaConceptoFacturacionAsistSocial
    pCargaUsosCFDI
    pCargaConceptosFactura
    pCargaImpresoras
    pCargaUsuarios
    pCargaFormatosTicket
    pCargaTiposPaciente
    pPreparaIEPS
    
    If rsPvParametro.RecordCount > 0 Then
        pCargaParametros
    End If
    pCargaDepartamentos
    pCargaUtilizaSocios
    pCargaConservarPrecioExclusion '*
    pCargaCantidadImpresiones '*
    sstPropiedades.TabVisible(0) = False
    sstPropiedades.Tab = 1
    sstPropiedades_Click 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pPreparaIEPS()
    
    If Not fblLicenciaIEPS Then ' no hay licencia IEPS ajustamos la pantalla
       'chkdesgloseIEPS.Visible = False
       
       'frmchecks.Top = 2640
       
'       Me.chkFacturaAutomatica.Top = 2400 '2150
'       Me.chkCorteHonorarios.Top = 2745 '2490
'       Me.chkCuentaPuenteBanco.Top = 3090 '2780
'       Me.chkCuentaPuenteIngresos.Top = 3360 '3100
'       Me.chkCerrarCuentas.Top = 3690 '3470
'       Me.Chkcuentacerrada.Top = 4065 '3840
'       Me.chkPermitirFacturarCargosFueraCatalogo.Top = 4425 '4180
'       Me.chkRequisicionespendientes.Top = 4755 '4450
'       Me.chkVerificar.Top = 5520 '5210
'       Me.chkSocios.Top = 5820 '5680
'       Me.chkSelDeptoCargosDir.Top = 6300 '6000
'
'       Me.Label30.Top = 6630 '6330
'       Me.mskHoraIniMsgCargo.Top = 6630 '6270
'
'       Me.Label6.Top = 6945 '6710
'       Me.txtIntervaloMsgCargo.Top = 6945 '6655
'       Me.Label31.Top = 6945 '6710
'
'       Me.Label28.Top = 7335 '7045
'       Me.txtDiasAbrirCuentaInt.Top = 7335 '7005
'
'       Me.Label29.Top = 7695 '7425
'       Me.txtDiasAbrirCuentaExt.Top = 7695 '7385
'
'       Me.Label5.Top = 8070 '7775
'       Me.txtTituloCtasPendFact.Top = 8070 '8025
       
       'Me.Frame7.Height = 9415
       
'       Me.Frame2.Top = 15000
'       Me.sstPropiedades.Height = 15000
'       Me.Height = 15000
       
       'Me.Refresh
       blLicenciaIEPS = False
       
       'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Else
       blLicenciaIEPS = True
        
       Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
End Sub

Private Sub pCargaImpresoras()
Dim vPrinter As Printer
Dim strImpre As String
Dim rsImpre As New ADODB.Recordset
Dim vlImpresoraReg As Boolean
    
    vlImpresoraReg = False
    'agrego la ultima impresora registrada del departamento correspondiente al cboImpresoras
    strImpre = "select CHRNOMBREIMPRESORA from impresoradepartamento where SMICVEDEPARTAMENTO = " & vgintNumeroDepartamento & " and CHRTIPO = 'FA'"
    Set rsImpre = frsRegresaRs(strImpre)
    'si existen impresoras en catalogo
    If rsImpre.RecordCount > 0 Then
        For Each vPrinter In Printers
            If UCase(Trim(rsImpre!chrNombreImpresora)) = UCase(vPrinter.DeviceName) Then vlImpresoraReg = True
            cboImpresoras.AddItem UCase(vPrinter.DeviceName)
            cboImpresoraTickets.AddItem UCase(vPrinter.DeviceName)
        Next
        If vlImpresoraReg = False Then cboImpresoras.AddItem UCase(Trim(rsImpre!chrNombreImpresora))
    Else
        For Each vPrinter In Printers
            cboImpresoras.AddItem UCase(vPrinter.DeviceName)
            cboImpresoraTickets.AddItem UCase(vPrinter.DeviceName)
        Next
    End If
    
    If cboImpresoras.ListCount > 0 Then
        cboImpresoras.ListIndex = 0
        cboImpresoraTickets.ListIndex = 0
    End If
    
End Sub

Private Sub pCargaConceptosFactura()
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset

    '-------------------------------------------
    ' Conceptos para el Excedente en Suma Asegurada
    '-------------------------------------------
    vlstrSentencia = "select smiCveConcepto Clave, chrDescripcion Descrip from pvConceptoFacturacion " & _
    " inner join pvconceptofacturacionempresa on pvconceptofacturacion.smicveconcepto =  pvconceptofacturacionempresa.intcveconceptofactura " & _
    " where pvconceptofacturacion.bitActivo = 1 and pvconceptofacturacionempresa.intcveempresacontable = " & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    pLlenarCboRs cboConceptoSumaAsegurada, rs, 0, 1
    
    'Deducible
    pLlenarCboRs cboConceptoDeducible, rs, 0, 1
    
    'Coaseguro
    pLlenarCboRs cboConceptoCoaseguro, rs, 0, 1
    
    'Coaseguro médico
    pLlenarCboRs cboConceptoCoaseguroMedico, rs, 0, 1
    
    'Coaseguro adicional
    pLlenarCboRs cboConceptoCoaseguroAdicional, rs, 0, 1
    
    'Copago
    pLlenarCboRs cboConceptoCopago, rs, 0, 1
    
    'Excedente de IVA
    pLlenarCboRs cboExcedenteIVA, rs, 0, 1
    
    'Conceptos de facturación parcial
    pLlenarCboRs cboConceptoParcial, rs, 0, 1
    
    'Conceptos de facturación para honorarios médicos
    pLlenarCboRs cboConceptoHonorariosMedicos, rs, 0, 1

End Sub

Private Sub pCargaTiposPaciente()
Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP("ME", "sp_PvSelTipoPacientePOS")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboTipoMedico, rs, 0, 1
    End If
    rs.Close
    
    Set rs = frsEjecuta_SP("EM", "sp_PvSelTipoPacientePOS")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboTipoEmpleado, rs, 0, 1
    End If
    rs.Close
    
End Sub

Private Sub pCargaUsuarios()
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "select intNumeroLogin Clave, rtrim(vchUsuario) Nombre from Login"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    If rs.RecordCount > 0 Then
        cmdSelecciona(0).Enabled = True
    Else
        cmdSelecciona(0).Enabled = False
    End If
    
    pLlenarListRs lstListaUsuarios, rs, 0, 1
    
    vlstrSentencia = "select PvLoginImpresoraTicket.intNumeroLogin Clave, rtrim(vchUsuario) Nombre from PvLoginImpresoraTicket inner join Login on Login.intNumeroLogin = PvLoginImpresoraTicket.intNumeroLogin"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    If rs.RecordCount > 0 Then
        cmdSelecciona(1).Enabled = True
    Else
        cmdSelecciona(1).Enabled = False
    End If
    
    pLlenarListRs lstUsuariosAsignados, rs, 0, 1
    
    rs.Close
    
End Sub

Private Sub pCargaFormatosTicket()
Dim vlstrSentencia As String
Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "Select intCveFormatoTicket, vchDescripcion From pvFormatoTicket"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboTickets, rs, 0, 1
    rs.Close
    
End Sub

Private Sub pCargaParametros()
On Error GoTo NotificaError
Dim vlvrnParametro As Variant
Dim rsCuentasGastosFletes As New ADODB.Recordset
Dim rsNumFormatoFactura As New ADODB.Recordset
Dim rsNumFormatoRecibo As New ADODB.Recordset
Dim rsImpresoraDepartamento As New ADODB.Recordset
Dim rsDepaMsg As New ADODB.Recordset
Dim vlrsPvParametro As New ADODB.Recordset
Dim vlstrSentencia As String
Dim X As Integer
Dim vlstrFormato As String
Dim vlstrFormatoPorc As String
Dim rsEntrada As New ADODB.Recordset
'Caso 19900
Dim rsAuditoriaCargos As New ADODB.Recordset
    
    vlstrFormato = "$###,###,###,###.00"
    vlstrFormatoPorc = "###.00"
    
    ' Leyendas en Dólares
    vlvrnParametro = fRegresaParametro("vchLeyenda1", "PvParametro", 0)
    If vlvrnParametro <> "" Then
        If vlvrnParametro <> 0 Then
            txtLeyenda1.Text = vlvrnParametro
        End If
    End If
    
    ' Título "Cuentas pendientes de facturar"
    txtTituloCtasPendFact.Text = fRegresaParametro("VCHTITULOCTASPENDFACT", "PvParametro", 0)
    
    vlvrnParametro = fRegresaParametro("vchLeyenda2", "PvParametro", 0)
    If vlvrnParametro <> "" Then
        If vlvrnParametro <> 0 Then
            txtLeyenda2.Text = vlvrnParametro
        End If
    End If
    
    vlvrnParametro = fRegresaParametro("vchLeyenda3", "PvParametro", 0)
    If vlvrnParametro <> "" Then
        If vlvrnParametro <> 0 Then
            txtLeyenda3.Text = vlvrnParametro
        End If
    End If
    
    '---------------------------------------------------------------------
    '««  Carga el ticket predeterminado para el departamento del login  »»
    '---------------------------------------------------------------------
    vlstrsql = "Select intCveFormatoTicket From pvTicketDepartamento where smiDepartamento=" & CStr(vgintNumeroDepartamento)
    Set rsNumFormatoFactura = frsRegresaRs(vlstrsql)
    If rsNumFormatoFactura.RecordCount <> 0 Then
        cboTickets.ListIndex = fintLocalizaCbo(cboTickets, rsNumFormatoFactura!intCveFormatoTicket)
    End If
    
    'Datos que salen en una factura del POS
    txtNombreFactura.Text = IIf(IsNull(rsPvParametro!chrNombreFacturaPOS), "", Trim(rsPvParametro!chrNombreFacturaPOS))
    txtRFCFactura.Text = IIf(IsNull(rsPvParametro!CHRRFCPOS), "", Trim(rsPvParametro!CHRRFCPOS))
    txtDireccionPOS.Text = IIf(IsNull(rsPvParametro!chrDireccionPOS), "", Trim(rsPvParametro!chrDireccionPOS))
    txtNumExterior.Text = IIf(IsNull(rsPvParametro!vchNumeroExteriorPOS), "", Trim(rsPvParametro!vchNumeroExteriorPOS))
    txtNumInterior.Text = IIf(IsNull(rsPvParametro!vchNumeroInteriorPOS), "", Trim(rsPvParametro!vchNumeroInteriorPOS))
    txtColoniaPOS.Text = IIf(IsNull(rsPvParametro!vchColoniaPOS), "", Trim(rsPvParametro!vchColoniaPOS))
    txtCPPOS.Text = IIf(IsNull(rsPvParametro!vchCodigoPostalPOS), "", Trim(rsPvParametro!vchCodigoPostalPOS))
    If cboCiudad.ListCount <> 0 Then
        cboCiudad.ListIndex = fintLocalizaCbo(cboCiudad, IIf(IsNull(rsPvParametro!intCveCiudad), 0, rsPvParametro!intCveCiudad))
    End If
    
    'Tipo de paciente para médicos POS
    If cboTipoMedico.ListCount > 0 Then
        cboTipoMedico.ListIndex = fintLocalizaCbo(cboTipoMedico, IIf(IsNull(rsPvParametro!intTipoPacMedico), 0, rsPvParametro!intTipoPacMedico))
    End If
    
    'Tipo de paciente para empleados POS
    If cboTipoEmpleado.ListCount > 0 Then
        cboTipoEmpleado.ListIndex = fintLocalizaCbo(cboTipoEmpleado, IIf(IsNull(rsPvParametro!intTipoPacEmpleado), 0, rsPvParametro!intTipoPacEmpleado))
    End If
   
    'Impresora donde salen las facturas
    vlstrSentencia = "select chrNombreImpresora Impresora from ImpresoraDepartamento where chrTipo = 'FA' and smiCveDepartamento = " & Trim(Str(vgintNumeroDepartamento))
    Set rsImpresoraDepartamento = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With rsImpresoraDepartamento
        If .RecordCount > 0 Then
            cboImpresoras.ListIndex = fintLocalizaCritCbo(cboImpresoras, UCase(Trim(!Impresora)))
        End If
        .Close
    End With
    
    'Impresora donde salen los tickets
    vlstrSentencia = "select chrNombreImpresora Impresora from ImpresoraDepartamento where chrTipo = 'TI' and smiCveDepartamento = " & Trim(Str(vgintNumeroDepartamento))
    Set rsImpresoraDepartamento = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With rsImpresoraDepartamento
        If .RecordCount > 0 Then
            cboImpresoraTickets.ListIndex = fintLocalizaCritCbo(cboImpresoraTickets, UCase(Trim(!Impresora)))
        End If
        .Close
    End With
    
    
    'El concepto de facturación utilizado para el exedente de la suma asegurada
    If cboConceptoSumaAsegurada.ListCount > 0 Then
        cboConceptoSumaAsegurada.ListIndex = fintLocalizaCbo(cboConceptoSumaAsegurada, IIf(IsNull(rsPvParametro!INTCONCEPTOSUMAASEGURADA), 0, rsPvParametro!INTCONCEPTOSUMAASEGURADA))
    End If
    
    'El concepto de facturación utilizado para DEDUCIBLE
    If cboConceptoDeducible.ListCount > 0 Then
        cboConceptoDeducible.ListIndex = fintLocalizaCbo(cboConceptoDeducible, IIf(IsNull(rsPvParametro!INTCONCEPTODEDUCIBLE), 0, rsPvParametro!INTCONCEPTODEDUCIBLE))
    End If
    
    'El concepto de facturación utilizado para COASEGURO
    If cboConceptoCoaseguro.ListCount > 0 Then
        cboConceptoCoaseguro.ListIndex = fintLocalizaCbo(cboConceptoCoaseguro, IIf(IsNull(rsPvParametro!INTCONCEPTOCOASEGURO), 0, rsPvParametro!INTCONCEPTOCOASEGURO))
    End If
    
    'El concepto de facturación utilizado para COASEGURO MËDICO
    If cboConceptoCoaseguroMedico.ListCount > 0 Then
        cboConceptoCoaseguroMedico.ListIndex = fintLocalizaCbo(cboConceptoCoaseguroMedico, IIf(IsNull(rsPvParametro!intConceptoCoaseguroMedico), 0, rsPvParametro!intConceptoCoaseguroMedico))
    End If
    
    'El concepto de facturación utilizado para COASEGURO ADICIONAL
    If cboConceptoCoaseguroAdicional.ListCount > 0 Then
        cboConceptoCoaseguroAdicional.ListIndex = fintLocalizaCbo(cboConceptoCoaseguroAdicional, IIf(IsNull(rsPvParametro!INTCONCEPTOCOASEGUROADICIONAL), 0, rsPvParametro!INTCONCEPTOCOASEGUROADICIONAL))
    End If
    
    'El concepto de facturación utilizado para COPAGO
    If cboConceptoCopago.ListCount > 0 Then
        cboConceptoCopago.ListIndex = fintLocalizaCbo(cboConceptoCopago, IIf(IsNull(rsPvParametro!INTCONCEPTOCOPAGO), 0, rsPvParametro!INTCONCEPTOCOPAGO))
    End If
    
    'El concepto de facturación para el excedente de IVA
    If cboExcedenteIVA.ListCount > 0 Then
        cboExcedenteIVA.ListIndex = fintLocalizaCbo(cboExcedenteIVA, IIf(IsNull(rsPvParametro!intExcedenteIVA), 0, rsPvParametro!intExcedenteIVA))
    End If
    
    chkCalcularCargosSeleccionados.Value = IIf(IsNull(rsPvParametro!BITCALCULARENBASEACARGOS), 0, rsPvParametro!BITCALCULARENBASEACARGOS)
    
    'Desglosar IVA's:
    chkDesglosarIVAExcedente.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARIVAEXCEDENTE), 0, rsPvParametro!INTDESGLOSARIVAEXCEDENTE)
    chkDesglosarIVAExcedente_Click
    chkDesglosarIVADeducible.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARIVADEDUCIBLE), 0, rsPvParametro!INTDESGLOSARIVADEDUCIBLE)
    chkDesglosarIVADeducible_Click
    chkDesglosarIVACoaseguro.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARIVACOASEGURO), 0, rsPvParametro!INTDESGLOSARIVACOASEGURO)
    chkDesglosarIVACoaseguro_Click
    chkDesglosarIVACoaseguroM.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARIVACOASEGUROMEDICO), 0, rsPvParametro!INTDESGLOSARIVACOASEGUROMEDICO)
    chkDesglosarIVACoaseguroM_Click
    chkDesglosarIVACoaseguroAdicional.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARIVACOASEGUROADICIO), 0, rsPvParametro!INTDESGLOSARIVACOASEGUROADICIO)
    chkDesglosarIVACoaseguroAdicional_Click
    chkDesglosarIVACopago.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARIVACOPAGO), 0, rsPvParametro!INTDESGLOSARIVACOPAGO)
    chkDesglosarIVACopago_Click
    
    'Desglosar importes:
    chkDesglosarExcedente.Value = IIf(IsNull(rsPvParametro!INTDESGLOSAREXCEDENTE), 0, rsPvParametro!INTDESGLOSAREXCEDENTE)
    chkDesglosarDeducible.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARDEDUCIBLE), 0, rsPvParametro!INTDESGLOSARDEDUCIBLE)
    chkDesglosarCoaseguro.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARCOASEGURO), 0, rsPvParametro!INTDESGLOSARCOASEGURO)
    chkDesglosarCoaseguroM.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARCOASEGUROMEDICO), 0, rsPvParametro!INTDESGLOSARCOASEGUROMEDICO)
    chkDesglosarCoaseguroAdicional.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARCOASEGUROADICIONAL), 0, rsPvParametro!INTDESGLOSARCOASEGUROADICIONAL)
    chkDesglosarCopago.Value = IIf(IsNull(rsPvParametro!INTDESGLOSARCOPAGO), 0, rsPvParametro!INTDESGLOSARCOPAGO)
      
    'El concepto de facturación parcial
    If cboConceptoParcial.ListCount > 0 Then
        cboConceptoParcial.ListIndex = fintLocalizaCbo(cboConceptoParcial, IIf(IsNull(rsPvParametro!intCveConceptoParcial), 0, rsPvParametro!intCveConceptoParcial))
    End If
    
    'Uso de CFDI facturado
    If cboUsoCFDIFacturado.ListCount > 0 Then
        cboUsoCFDIFacturado.ListIndex = fintLocalizaCbo(cboUsoCFDIFacturado, IIf(IsNull(rsPvParametro!INTCVEUSOCFDIHONOFACTURADO), 0, rsPvParametro!INTCVEUSOCFDIHONOFACTURADO))
    End If
    
    'El Concepto de facturación para factura de asistencia social
    If cboConceptoFacturacionAsistSocial.ListCount > 0 Then
        cboConceptoFacturacionAsistSocial.ListIndex = fintLocalizaCbo(cboConceptoFacturacionAsistSocial, IIf(IsNull(rsSiParametro!VCHVALOR), 0, rsSiParametro!VCHVALOR))
    End If
    
    'El concepto de facturación para honorarios médicos
    If cboConceptoHonorariosMedicos.ListCount > 0 Then
        cboConceptoHonorariosMedicos.ListIndex = fintLocalizaCbo(cboConceptoHonorariosMedicos, IIf(IsNull(rsPvParametro!intCveConceptoHonorarioMedico), 0, rsPvParametro!intCveConceptoHonorarioMedico))
    End If
        
    'El valor de Leyenda descuentos
    txtLeyendaDescuentos.Text = fRegresaParametro("VCHLEYENDADESCUENTO", "PvParametro", 0)
    
    'El valor de Leyenda de informacion al cliente
    txtLeyendaCliente.Text = fRegresaParametro("VCHLEYENDACLIENTE", "PvParametro", 0)
    
    chkCorteHonorarios.Value = fRegresaParametro("bitIncluyeHonorarioCorte", "PvParametro", 0)
    chkVerificar.Value = fRegresaParametro("BITVERIFICARREQUISICIONES", "PvParametro", 0)
    txtDiasAbrirCuentaInt.Text = IIf(IsNull(rsPvParametro!intDiasAbrirCuentasInternos), "0", rsPvParametro!intDiasAbrirCuentasInternos)
    txtDiasAbrirCuentaExt.Text = IIf(IsNull(rsPvParametro!intDiasAbrirCuentasExternos), "0", rsPvParametro!intDiasAbrirCuentasExternos)
    chkCerrarCuentasAut.Value = IIf(IsNull(rsPvParametro!bitCerrarCuentasExtAut), vbUnchecked, IIf(rsPvParametro!bitCerrarCuentasExtAut = 0, vbUnchecked, vbChecked))
    chkDesactivarExterno.Value = IIf(IsNull(rsPvParametro!BITDESACTIVARALFACTURAR), vbUnchecked, IIf(rsPvParametro!BITDESACTIVARALFACTURAR = 0, vbUnchecked, vbChecked))
    Chkcuentacerrada.Value = IIf(IsNull(rsPvParametro!bitCuentaCerrada), vbUnchecked, IIf(rsPvParametro!bitCuentaCerrada = 0, vbUnchecked, vbChecked))
    chkRequisicionespendientes.Value = IIf(IsNull(rsPvParametro!bitfacturarconreqpendiente), vbUnchecked, IIf(rsPvParametro!bitfacturarconreqpendiente = 0, vbUnchecked, vbChecked))
    chkFacturaAutomatica.Value = IIf(IsNull(rsPvParametro!bitFacturarVentaPublico), 0, rsPvParametro!bitFacturarVentaPublico)
    Me.chkdesgloseIEPS.Value = IIf(Me.chkdesgloseIEPS.Enabled = True, fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), 0)
    chkCuentaPuenteBanco.Value = IIf(IsNull(rsPvParametro!bitUtilizaCuentaPuenteBanco), vbUnchecked, IIf(rsPvParametro!bitUtilizaCuentaPuenteBanco = 0, vbUnchecked, vbChecked))
    chkCuentaPuenteIngresos.Value = IIf(IsNull(rsPvParametro!bitUtilizaCuentaPuenteIngresos), vbUnchecked, IIf(rsPvParametro!bitUtilizaCuentaPuenteIngresos = 0, vbUnchecked, vbChecked))
    chkTrasladarCargos.Value = IIf(IsNull(rsPvParametro!bitTrasladaCargos), vbUnchecked, IIf(rsPvParametro!bitTrasladaCargos = 0, vbUnchecked, vbChecked))
    chkCancelarRecibosOtroDepto.Value = IIf(IsNull(rsPvParametro!bitCancelarRecibosOtroDepto), vbUnchecked, IIf(rsPvParametro!bitCancelarRecibosOtroDepto = 0, vbUnchecked, vbChecked))
    txtDiasSinRespPresupuesto.Text = IIf(IsNull(rsPvParametro!IntDiasSinRespPresupuesto), "0", rsPvParametro!IntDiasSinRespPresupuesto)
    chkCapturarMargenSubrogado.Value = IIf(IsNull(rsPvParametro!BITCAPTURAMARGENSUBROGADO), vbUnchecked, IIf(rsPvParametro!BITCAPTURAMARGENSUBROGADO = 0, vbUnchecked, vbChecked))
    chkValidacionPMPVentaPublico.Value = IIf(IsNull(rsPvParametro!BitValidacionPMPVentaPublic), vbUnchecked, IIf(rsPvParametro!BitValidacionPMPVentaPublic = 0, vbUnchecked, vbChecked))
    
    'Caso 19900
    'ChkAuditoriaCargos.Value = IIf(IsNull(rsPvParametro!BITAUDITORIADECARGOS), vbUnchecked, IIf(rsPvParametro!BITAUDITORIADECARGOS = 0, vbUnchecked, vbChecked))
    vlstrSentencia = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITAUDITORIADECARGOS'"
    Set rsAuditoriaCargos = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With rsAuditoriaCargos
        If .RecordCount > 0 Then
            ChkAuditoriaCargos.Value = IIf(IsNull(!VCHVALOR), vbUnchecked, IIf(!VCHVALOR = 0, vbUnchecked, vbChecked))
        End If
        .Close
    End With
    
    
    ' Departamento
    vlstrsql = "select vchDescripcion,smiCveDepartamento from NoDepartamento WHERE BITESTATUS = 1"
    Set rsDepaMsg = frsRegresaRs(vlstrsql)
    If rsDepaMsg.RecordCount <> 0 Then
          pLlenarCboRs cboDepartamentoMsg, rsDepaMsg, 1, 0
    End If
    
    If cboDepartamentoMsg.ListCount > 0 Then
        If IsNull(rsPvParametro!smiCveDepartamentoMsg) Then
            cboDepartamentoMsg.ListIndex = -1
        Else
            cboDepartamentoMsg.ListIndex = fintLocalizaCbo(cboDepartamentoMsg, rsPvParametro!smiCveDepartamentoMsg)
        End If
    End If
    
    If Not IsNull(rsPvParametro!dtmHoraIniMsgCargo) Then
        mskHoraIniMsgCargo.Text = FormatDateTime(rsPvParametro!dtmHoraIniMsgCargo, 4)
    End If
    txtIntervaloMsgCargo.Text = IIf(IsNull(rsPvParametro!intIntervaloMsgCargo), "", rsPvParametro!intIntervaloMsgCargo)
    '|---------------------------------------------------------
    '|  Parámetros registrados en SiParametro
    '|---------------------------------------------------------
    ' Departamento
    vlstrsql = "select trim(chrDescripcion), intNumConcepto from PvConceptoPago WHERE bitestatusActivo = 1 and chrtipo='NO' and bitDesglosaIva = 0"
    cboConceptoEntrada.Clear
    Set rsEntrada = frsRegresaRs(vlstrsql)
    If rsEntrada.RecordCount <> 0 Then
          pLlenarCboRs cboConceptoEntrada, rsEntrada, 1, 0
    End If
    
    lblnCargandoParametros = True
    Set vlrsPvParametro = frsSelParametros("PV", vgintClaveEmpresaContable)
    Do Until vlrsPvParametro.EOF
        Select Case vlrsPvParametro.Fields("Nombre").Value
            Case "BITFACTURARCARGOSFUERACATALOGO"
                chkPermitirFacturarCargosFueraCatalogo.Value = vlrsPvParametro!valor
            Case "BITCUENTAPORPAGARHONORARIOMEDICOCREDITO"
                chkHonorarioMedicoCredito.Value = vlrsPvParametro!valor
            Case "BITCERRARCUENTAREQUIPENDIENTES"
                chkCerrarCuentas.Value = vlrsPvParametro!valor
            Case "BITSELECCIONARDEPTOCARGODIRECTO"
                chkSelDeptoCargosDir.Value = vlrsPvParametro!valor
            Case "BITCORTECAJACHICASINXML"
                chkPermitirCorte.Value = vlrsPvParametro!valor
            Case "INTCONCEPTOENTRADACOMPROBACION"
                If IsNull(vlrsPvParametro!valor) Then
                    cboConceptoEntrada.ListIndex = -1
                Else
                    cboConceptoEntrada.ListIndex = fintLocalizaCbo(cboConceptoEntrada, vlrsPvParametro!valor)
                End If
            Case "CHRINCLUIRCFDICONCEPTOSSEGURO"
                If vlrsPvParametro!valor = "D" Then
                    optIncCS(1).Value = True
                Else
                    optIncCS(0).Value = True
                End If
            Case "INTTOTIMPCFDICONCEPTOSSEGURO"
                chkTotCS.Value = IIf(vlrsPvParametro!valor = "0", vbUnchecked, vbChecked)
                optTipoDesglose(IIf(vlrsPvParametro!valor = "2", 1, 0)).Value = True
            Case "VCHTIPODESCTONOTAEXCEDENTE"
                If IsNull(vlrsPvParametro!valor) Then
                    optTipoDesctoExcedente(0).Value = True
                Else
                    If Trim(vlrsPvParametro!valor) = "C" Then
                        optTipoDesctoExcedente(1).Value = True
                    Else
                        optTipoDesctoExcedente(0).Value = True
                    End If
                End If
            Case "VCHTIPODESCTONOTADEDUCIBLE"
                If IsNull(vlrsPvParametro!valor) Then
                    optTipoDesctoDeducible(0).Value = True
                Else
                    If Trim(vlrsPvParametro!valor) = "C" Then
                        optTipoDesctoDeducible(1).Value = True
                    Else
                        optTipoDesctoDeducible(0).Value = True
                    End If
                End If
            Case "VCHTIPODESCTONOTACOASEGURO"
                If IsNull(vlrsPvParametro!valor) Then
                    optTipoDesctoCoaseguro(0).Value = True
                Else
                    If Trim(vlrsPvParametro!valor) = "C" Then
                        optTipoDesctoCoaseguro(1).Value = True
                    Else
                        optTipoDesctoCoaseguro(0).Value = True
                    End If
                End If
            Case "VCHTIPODESCTONOTACOASEGUROMEDICO"
                If IsNull(vlrsPvParametro!valor) Then
                    optTipoDesctoCoaseguroMedico(0).Value = True
                Else
                    If Trim(vlrsPvParametro!valor) = "C" Then
                        optTipoDesctoCoaseguroMedico(1).Value = True
                    Else
                        optTipoDesctoCoaseguroMedico(0).Value = True
                    End If
                End If
            Case "VCHTIPODESCTONOTACOASEGUROADICIONAL"
                If IsNull(vlrsPvParametro!valor) Then
                    optTipoDesctoCoaseguroAdicional(0).Value = True
                Else
                    If Trim(vlrsPvParametro!valor) = "C" Then
                        optTipoDesctoCoaseguroAdicional(1).Value = True
                    Else
                        optTipoDesctoCoaseguroAdicional(0).Value = True
                    End If
                End If
            Case "VCHTIPODESCTONOTACOPAGO"
                If IsNull(vlrsPvParametro!valor) Then
                    optTipoDesctoCopago(0).Value = True
                Else
                    If Trim(vlrsPvParametro!valor) = "C" Then
                        optTipoDesctoCopago(1).Value = True
                    Else
                        optTipoDesctoCopago(0).Value = True
                    End If
                End If
            Case "NUMDESCTONOTAEXCEDENTE"
                If IsNull(vlrsPvParametro!valor) Then
                    txtPorcentajeExcedentePorNota.Text = Format(0, IIf(optTipoDesctoExcedente(0).Value, vlstrFormatoPorc, vlstrFormato))
                Else
                    txtPorcentajeExcedentePorNota.Text = Format(Val(vlrsPvParametro!valor), IIf(optTipoDesctoExcedente(0).Value, vlstrFormatoPorc, vlstrFormato))
                End If
            Case "NUMDESCTONOTADEDUCIBLE"
                If IsNull(vlrsPvParametro!valor) Then
                    txtPorcentajeDeduciblePorNota.Text = Format(0, IIf(optTipoDesctoDeducible(0).Value, vlstrFormatoPorc, vlstrFormato))
                Else
                    txtPorcentajeDeduciblePorNota.Text = Format(Val(vlrsPvParametro!valor), IIf(optTipoDesctoDeducible(0).Value, vlstrFormatoPorc, vlstrFormato))
                End If
            Case "NUMDESCTONOTACOASEGURO"
                If IsNull(vlrsPvParametro!valor) Then
                    txtPorcentajeCoaseguroPorNota.Text = Format(0, IIf(optTipoDesctoCoaseguro(0).Value, vlstrFormatoPorc, vlstrFormato))
                Else
                    txtPorcentajeCoaseguroPorNota.Text = Format(Val(vlrsPvParametro!valor), IIf(optTipoDesctoCoaseguro(0).Value, vlstrFormatoPorc, vlstrFormato))
                End If
            Case "NUMDESCTONOTACOASEGUROMEDICO"
                If IsNull(vlrsPvParametro!valor) Then
                    txtPorcentajeCoaseguroMPorNota.Text = Format(0, IIf(optTipoDesctoCoaseguroMedico(0).Value, vlstrFormatoPorc, vlstrFormato))
                Else
                    txtPorcentajeCoaseguroMPorNota.Text = Format(Val(vlrsPvParametro!valor), IIf(optTipoDesctoCoaseguroMedico(0).Value, vlstrFormatoPorc, vlstrFormato))
                End If
            Case "NUMDESCTONOTACOASEGUROADICIONAL"
                If IsNull(vlrsPvParametro!valor) Then
                    txtPorcentajeCoasAdicionalPorNota.Text = Format(0, IIf(optTipoDesctoCoaseguroAdicional(0).Value, vlstrFormatoPorc, vlstrFormato))
                Else
                    txtPorcentajeCoasAdicionalPorNota.Text = Format(Val(vlrsPvParametro!valor), IIf(optTipoDesctoCoaseguroAdicional(0).Value, vlstrFormatoPorc, vlstrFormato))
                End If
            Case "NUMDESCTONOTACOPAGO"
                If IsNull(vlrsPvParametro!valor) Then
                    txtPorcentajeCopagoPorNota.Text = Format(0, IIf(optTipoDesctoCopago(0).Value, vlstrFormatoPorc, vlstrFormato))
                Else
                    txtPorcentajeCopagoPorNota.Text = Format(Val(vlrsPvParametro!valor), IIf(optTipoDesctoCopago(0).Value, vlstrFormatoPorc, vlstrFormato))
                End If
            Case "NUMDESCTONOTALIMITEEXCEDENTE"
                If IsNull(vlrsPvParametro!valor) Then
                    txtCantidadLimiteExcedente.Text = Format(0, vlstrFormato)
                Else
                    txtCantidadLimiteExcedente.Text = Format(Val(vlrsPvParametro!valor), vlstrFormato)
                End If
            Case "NUMDESCTONOTALIMITEDEDUCIBLE"
                If IsNull(vlrsPvParametro!valor) Then
                    txtCantidadLimiteDeducible.Text = Format(0, vlstrFormato)
                Else
                    txtCantidadLimiteDeducible.Text = Format(Val(vlrsPvParametro!valor), vlstrFormato)
                End If
            Case "NUMDESCTONOTALIMITECOASEGURO"
                If IsNull(vlrsPvParametro!valor) Then
                    txtCantidadLimiteCoaseguro.Text = Format(0, vlstrFormato)
                Else
                    txtCantidadLimiteCoaseguro.Text = Format(Val(vlrsPvParametro!valor), vlstrFormato)
                End If
            Case "NUMDESCTONOTALIMITEMEDICO"
                If IsNull(vlrsPvParametro!valor) Then
                    txtCantidadLimiteCoaseguroM.Text = Format(0, vlstrFormato)
                Else
                    txtCantidadLimiteCoaseguroM.Text = Format(Val(vlrsPvParametro!valor), vlstrFormato)
                End If
            Case "NUMDESCTONOTALIMITEADICIONAL"
                If IsNull(vlrsPvParametro!valor) Then
                    txtCantidadLimiteCoasAdicional.Text = Format(0, vlstrFormato)
                Else
                    txtCantidadLimiteCoasAdicional.Text = Format(Val(vlrsPvParametro!valor), vlstrFormato)
                End If
            Case "NUMDESCTONOTALIMITECOPAGO"
                If IsNull(vlrsPvParametro!valor) Then
                    txtCantidadLimiteCopago.Text = Format(0, vlstrFormato)
                Else
                    txtCantidadLimiteCopago.Text = Format(Val(vlrsPvParametro!valor), vlstrFormato)
                End If
            Case "BITGENERAPACIENTEEXTERNO"
                chkAbrirCuentaExterna.Value = vlrsPvParametro!valor
        End Select
        vlrsPvParametro.MoveNext
    Loop
    vlrsPvParametro.Close
    If optTipoDesctoExcedente(0).Value Then
        txtCantidadLimiteExcedente.Enabled = Val(Format(txtPorcentajeExcedentePorNota.Text, "")) > 0
        lbCantidadLimiteExcedente.Enabled = Val(Format(txtPorcentajeExcedentePorNota.Text, "")) > 0
    End If
    If optTipoDesctoDeducible(0).Value Then
        txtCantidadLimiteDeducible.Enabled = Val(Format(txtPorcentajeDeduciblePorNota.Text, "")) > 0
        lbCantidadLimiteDeducible.Enabled = Val(Format(txtPorcentajeDeduciblePorNota.Text, "")) > 0
    End If
    If optTipoDesctoCoaseguro(0).Value Then
        txtCantidadLimiteCoaseguro.Enabled = Val(Format(txtPorcentajeCoaseguroPorNota.Text, "")) > 0
        lbCantidadLimiteCoaseguro.Enabled = Val(Format(txtPorcentajeCoaseguroPorNota.Text, "")) > 0
    End If
    If optTipoDesctoCoaseguroMedico(0).Value Then
        txtCantidadLimiteCoaseguroM.Enabled = Val(Format(txtPorcentajeCoaseguroMPorNota.Text, "")) > 0
        lbCantidadLimiteCoaMedico.Enabled = Val(Format(txtPorcentajeCoaseguroMPorNota.Text, "")) > 0
    End If
    If optTipoDesctoCoaseguroAdicional(0).Value Then
        txtCantidadLimiteCoasAdicional.Enabled = Val(Format(txtPorcentajeCoasAdicionalPorNota.Text, "")) > 0
        lbCantidadLimiteCoaAdicional.Enabled = Val(Format(txtPorcentajeCoasAdicionalPorNota.Text, "")) > 0
    End If
    If optTipoDesctoCopago(0).Value Then
        txtCantidadLimiteCopago.Enabled = Val(Format(txtPorcentajeCopagoPorNota.Text, "")) > 0
        lbCantidadLimiteCoPago.Enabled = Val(Format(txtPorcentajeCopagoPorNota.Text, "")) > 0
    End If
    
    lblnCargandoParametros = False
    If Not chkTotCS.Enabled Then
        chkTotCS.Value = vbUnchecked
    End If
    If chkTotCS.Value = vbUnchecked Then
        chkTotCS_Click
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaParametros"))
End Sub

Private Sub pCuentaContable(mskNombre As MaskEdBox, txtNombre As TextBox, objNombre As Object)
On Error GoTo NotificaError
    
    If mskNombre.ClipText = "" Then
        vllngNumeroCuenta = flngBusquedaCuentasContables()
        If vllngNumeroCuenta <> 0 Then
            mskNombre.Text = fstrCuentaContable(vllngNumeroCuenta)
            txtNombre.Text = fstrDescripcionCuenta(fstrCuentaContable(vllngNumeroCuenta), vgintClaveEmpresaContable)
            objNombre.SetFocus
        Else
            mskNombre.SetFocus
        End If
    Else
        
      mskNombre.Mask = ""
      mskNombre.Text = fstrCuentaCompleta(mskNombre.Text)
      mskNombre.Mask = vgstrEstructuraCuentaContable
        
        txtNombre.Text = fstrDescripcionCuenta(mskNombre.Text, vgintClaveEmpresaContable)
        If txtNombre.Text <> "" Then
            objNombre.SetFocus
        Else
          'No se encontró la cuenta contable.
          MsgBox SIHOMsg(222), vbOKOnly + vbExclamation, "Mensaje"
          mskNombre.Mask = ""
          mskNombre.Text = ""
          mskNombre.Mask = vgstrEstructuraCuentaContable
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCuentaContable"))
End Sub
Private Sub pCargaDepartamentos()
On Error GoTo NotificaError
Dim vlstrSentencia As String
Dim rsDepartamento As New ADODB.Recordset
    
    vlstrSentencia = " SELECT D.smiCveDepartamento"
    vlstrSentencia = vlstrSentencia & "     , D.vchDescripcion"
    vlstrSentencia = vlstrSentencia & "     , CASE WHEN DF.chrEstatus = 'T' THEN 1 ELSE 0 END AS Seleccion  "
    vlstrSentencia = vlstrSentencia & " FROM nodepartamento D"
    vlstrSentencia = vlstrSentencia & "   left outer join PvDatosFiscalesDepartamento DF ON D.SMICVEDEPARTAMENTO= DF.SMIDEPARTAMENTO"
    vlstrSentencia = vlstrSentencia & " WHERE D.chrClasificacion <> 'E'"
    vlstrSentencia = vlstrSentencia & "    AND D.bitEstatus = 1"
    vlstrSentencia = vlstrSentencia & "    AND D.bitAtiendePacientes = 0"
    vlstrSentencia = vlstrSentencia & " ORDER BY vchDescripcion"
    
    Set rsDepartamento = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    Do While Not rsDepartamento.EOF
        With lstDepartamentos
            .Visible = False
            .AddItem rsDepartamento!VCHDESCRIPCION
            .ItemData(.newIndex) = rsDepartamento!smicvedepartamento
            .Selected(.newIndex) = rsDepartamento!seleccion = 1
            rsDepartamento.MoveNext
            .Visible = True
        End With
    Loop
    
    If rsDepartamento.RecordCount > 0 Then
        lstDepartamentos.Enabled = True
    End If
    rsDepartamento.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDepartamentos"))
End Sub

Private Sub grdCopiasImpresion_Click()
    If grdCopiasImpresion.MouseRow <> 0 Then ' mientras que no sea el renglon de los titulos
       If grdCopiasImpresion.Col = 1 Then
          pEditarColumna 32, txtEditCol, grdCopiasImpresion
       End If
    End If
End Sub
Private Sub grdCopiasImpresion_GotFocus()
On Error GoTo NotificaError
        If vgblnNoEditar Then Exit Sub
        If grdCopiasImpresion.Col = 1 Then
            pSetCellValueCol grdCopiasImpresion, txtEditCol
        Else
            Exit Sub
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCopiasImpresion_GotFocus"))
End Sub
Private Sub grdCopiasImpresion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If grdCopiasImpresion.Col = 1 Then
        If (KeyCode = vbKeyF2 Or KeyCode = vbKeyReturn) And grdCopiasImpresion.Row <> 0 Then pEditarColumna 13, txtEditCol, grdCopiasImpresion
    Else
        If KeyCode = vbKeyReturn Then
            If grdCopiasImpresion.Row - 1 < grdCopiasImpresion.Rows Then
                If grdCopiasImpresion.Row = grdCopiasImpresion.Rows - 1 Then
                    grdCopiasImpresion.Row = 1
                Else
                    grdCopiasImpresion.Row = grdCopiasImpresion.Row + 1
                End If
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCopiasImpresion_KeyDown"))
End Sub
Private Sub grdCopiasImpresion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If grdCopiasImpresion.Row <> 0 Then
        If grdCopiasImpresion.Col = 1 Then  'Columna que puede ser editada
            pEditarColumna KeyAscii, txtEditCol, grdCopiasImpresion
        Else
            Exit Sub
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCopiasImpresion_KeyPress"))
End Sub
Private Sub grdCopiasImpresion_LeaveCell()
    On Error GoTo NotificaError
        If vgblnNoEditar Then Exit Sub
        grdCopiasImpresion_GotFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCopiasImpresion_LeaveCell"))
End Sub

Private Sub mskHoraIniMsgCargo_GotFocus()
    pSelMkTexto mskHoraIniMsgCargo
End Sub

Private Sub mskHoraIniMsgCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtIntervaloMsgCargo
    End If
    
End Sub

Private Sub mskHoraIniMsgCargo_LostFocus()
    
    If Not IsDate(mskHoraIniMsgCargo.Text) And Trim(mskHoraIniMsgCargo) <> ":" Then
        MsgBox SIHOMsg(41), vbCritical, "Mensaje"
        If fblnCanFocus(mskHoraIniMsgCargo) Then mskHoraIniMsgCargo.SetFocus
    End If
    
End Sub

Private Sub optIncCS_Click(Index As Integer)
        chkTotCS.Enabled = True
End Sub

Private Sub optIncCS_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesglose_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoCoaseguro_Click(Index As Integer)
    
    If Index = 0 Then   ' Por porcentaje
        If Not lblnCargandoParametros Then
            txtPorcentajeCoaseguroPorNota.Text = ".00"
            txtCantidadLimiteCoaseguro.Text = Format(0, "$###,###,###,###.00")
            lbCantidadLimiteCoaseguro.Enabled = False
        End If
        lbPorcentaje(2).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoaseguro.Enabled = False
        txtCantidadLimiteCoaseguro.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCoaseguro.Enabled = False
        lbPorcentaje(2).Visible = False
        If Not lblnCargandoParametros Then txtPorcentajeCoaseguroPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
    
End Sub

Private Sub optTipoDesctoCoaseguro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoCoaseguroAdicional_Click(Index As Integer)
    
    If Index = 0 Then   ' Por porcentaje
        If Not lblnCargandoParametros Then
            txtPorcentajeCoasAdicionalPorNota.Text = ".00"
            txtCantidadLimiteCoasAdicional.Text = Format(0, "$###,###,###,###.00")
            lbCantidadLimiteCoaAdicional.Enabled = False
        End If
        lbPorcentaje(4).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoaAdicional.Enabled = False
        txtCantidadLimiteCoasAdicional.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCoasAdicional.Enabled = False
        lbPorcentaje(4).Visible = False
        If Not lblnCargandoParametros Then txtPorcentajeCoasAdicionalPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
    
End Sub

Private Sub optTipoDesctoCoaseguroAdicional_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoCoaseguroMedico_Click(Index As Integer)
    
    If Index = 0 Then   ' Por porcentaje
        If Not lblnCargandoParametros Then
            txtPorcentajeCoaseguroMPorNota.Text = ".00"
            txtCantidadLimiteCoaseguroM.Text = Format(0, "$###,###,###,###.00")
            lbCantidadLimiteCoaMedico.Enabled = False
        End If
        lbPorcentaje(3).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoaMedico.Enabled = False
        txtCantidadLimiteCoaseguroM.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCoaseguroM.Enabled = False
        lbPorcentaje(3).Visible = False
        If Not lblnCargandoParametros Then txtPorcentajeCoaseguroMPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
    
End Sub

Private Sub optTipoDesctoCoaseguroMedico_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoCopago_Click(Index As Integer)
        
    If Index = 0 Then   ' Por porcentaje
        If Not lblnCargandoParametros Then
            txtPorcentajeCopagoPorNota.Text = ".00"
            txtCantidadLimiteCopago.Text = Format(0, "$###,###,###,###.00")
            lbCantidadLimiteCoPago.Enabled = False
        End If
        lbPorcentaje(5).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteCoPago.Enabled = False
        txtCantidadLimiteCopago.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteCopago.Enabled = False
        lbPorcentaje(5).Visible = False
        If Not lblnCargandoParametros Then txtPorcentajeCopagoPorNota.Text = Format(0, "$###,###,###,###.00")
    End If
    
End Sub

Private Sub optTipoDesctoCopago_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoDeducible_Click(Index As Integer)
    
    If Index = 0 Then   ' Por porcentaje
        If Not lblnCargandoParametros Then
            txtPorcentajeDeduciblePorNota.Text = ".00"
            txtCantidadLimiteDeducible.Text = Format(0, "$###,###,###,###.00")
            lbCantidadLimiteDeducible.Enabled = False
        End If
        lbPorcentaje(1).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteDeducible.Enabled = False
        txtCantidadLimiteDeducible.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteDeducible.Enabled = False
        lbPorcentaje(1).Visible = False
        If Not lblnCargandoParametros Then txtPorcentajeDeduciblePorNota.Text = Format(0, "$###,###,###,###.00")
    End If
    
End Sub

Private Sub optTipoDesctoDeducible_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub optTipoDesctoExcedente_Click(Index As Integer)

    If Index = 0 Then   ' Por porcentaje
        If Not lblnCargandoParametros Then
            txtPorcentajeExcedentePorNota.Text = ".00"
            txtCantidadLimiteExcedente.Text = Format(0, "$###,###,###,###.00")
            lbCantidadLimiteExcedente.Enabled = False
        End If
        lbPorcentaje(0).Visible = True
    Else ' Por cantidad
        lbCantidadLimiteExcedente.Enabled = False
        txtCantidadLimiteExcedente.Text = Format(0, "$###,###,###,###.00")
        txtCantidadLimiteExcedente.Enabled = False
        lbPorcentaje(0).Visible = False
        If Not lblnCargandoParametros Then txtPorcentajeExcedentePorNota.Text = Format(0, "$###,###,###,###.00")
    End If

End Sub

Private Sub optTipoDesctoExcedente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub SSTab1_GotFocus()
    pEnfocaCbo cboTickets
End Sub

Private Sub sstPropiedades_Click(PreviousTab As Integer)

    Select Case sstPropiedades.Tab
    Case 1
        SSTab1.Tab = 0
        If fblnCanFocus(cboTickets) Then cboTickets.SetFocus
        Frame2.Top = 8150
        Me.Height = 9500
    Case 2
        sstVentaPublico.Tab = 0
        If fblnCanFocus(txtNombreFactura) Then txtNombreFactura.SetFocus
        Frame2.Top = 7850
        Me.Height = 9500
    Case 3
        If fblnCanFocus(cboConceptoSumaAsegurada) Then cboConceptoSumaAsegurada.SetFocus
        Frame2.Top = 9650
        Me.Height = 10880
    Case 4
        If fblnCanFocus(cboDepartamentoMsg) Then cboDepartamentoMsg.SetFocus
        Frame2.Top = 10550 '10300
        Me.Height = 11725 '11460
'        Frame2.Top = 10300 ''9970
'        Me.Height = 11460
        If blLicenciaIEPS Then
            chkdesgloseIEPS.Enabled = True
            chkFacturaAutomatica.Enabled = True
           'Frame2.Top = 10455 '10100
           'Me.Height = 11565 '11310
        Else
            chkdesgloseIEPS.Enabled = False
            chkFacturaAutomatica.Enabled = False
           'Frame2.Top = 10555 '9765
           'Me.Height = 11365 '11005
        End If
    End Select

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
End Sub

Private Sub sstVentaPublico_GotFocus()
    txtNombreFactura.SetFocus
End Sub

Private Sub txtCantidadLimiteCoasAdicional_Click()
    pSelTextBox txtCantidadLimiteCoasAdicional
End Sub

Private Sub txtCantidadLimiteCoasAdicional_GotFocus()
    pSelTextBox txtCantidadLimiteCoasAdicional
End Sub

Private Sub txtCantidadLimiteCoasAdicional_KeyPress(KeyAscii As Integer)
    
    If Not fblnFormatoCantidad(txtCantidadLimiteCoasAdicional, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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
    
    If Not fblnFormatoCantidad(txtCantidadLimiteCoaseguro, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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
    
    If Not fblnFormatoCantidad(txtCantidadLimiteCoaseguroM, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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

    If Not fblnFormatoCantidad(txtCantidadLimiteCopago, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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
    
    If Not fblnFormatoCantidad(txtCantidadLimiteDeducible, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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
    
    If Not fblnFormatoCantidad(txtCantidadLimiteExcedente, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
End Sub

Private Sub txtCantidadLimiteExcedente_LostFocus()
    txtCantidadLimiteExcedente.Text = Format(Val(Format(txtCantidadLimiteExcedente.Text, "")), "$###,###,###,###.00")
End Sub

Private Sub txtColoniaPOS_GotFocus()
    pSelTextBox txtColoniaPOS
End Sub

Private Sub txtColoniaPOS_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtCPPOS
    End If
    
End Sub

Private Sub txtColoniaPOS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCPPOS_GotFocus()
    pSelTextBox txtCPPOS
End Sub

Private Sub txtCPPOS_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaCbo cboCiudad
    End If
    
End Sub

Private Sub txtCPPOS_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCP_KeyPress"))
End Sub

Private Sub txtDiasAbrirCuentaExt_GotFocus()
    pSelTextBox txtDiasAbrirCuentaExt
End Sub

Private Sub txtDiasAbrirCuentaExt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDiasSinRespPresupuesto.SetFocus
    End If
End Sub

Private Sub txtDiasAbrirCuentaExt_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDiasAbrirCuentaExt_LostFocus()
    
    If Not IsNumeric(txtDiasAbrirCuentaExt.Text) Then
        txtDiasAbrirCuentaExt.Text = "0"
    End If
    
End Sub

Private Sub txtDiasAbrirCuentaInt_GotFocus()
    pSelTextBox txtDiasAbrirCuentaInt
End Sub

Private Sub txtDiasAbrirCuentaInt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtDiasAbrirCuentaExt
    End If
    
End Sub

Private Sub txtDiasAbrirCuentaInt_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDiasAbrirCuentaInt_LostFocus()
    
    If Not IsNumeric(txtDiasAbrirCuentaInt.Text) Then
        txtDiasAbrirCuentaInt.Text = "0"
    End If
    
End Sub

Private Sub txtDiasSinRespPresupuesto_GotFocus()
    pSelTextBox txtDiasSinRespPresupuesto
End Sub


Private Sub txtDiasSinRespPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTituloCtasPendFact.SetFocus
    End If
End Sub


Private Sub txtDiasSinRespPresupuesto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtDiasSinRespPresupuesto_LostFocus()
    If Not IsNumeric(txtDiasAbrirCuentaExt.Text) Then
        txtDiasAbrirCuentaExt.Text = "0"
    End If
End Sub


Private Sub txtDireccionPOS_GotFocus()
    pSelTextBox txtDireccionPOS
End Sub

Private Sub txtDireccionPOS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtEditCol_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtEditCol_LostFocus()
  pSetCellValueCol grdCopiasImpresion, txtEditCol
End Sub

Private Sub txtIntervaloMsgCargo_GotFocus()
    pSelTextBox txtIntervaloMsgCargo
End Sub

Private Sub txtIntervaloMsgCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtDiasAbrirCuentaInt
    End If
    
End Sub

Private Sub txtIntervaloMsgCargo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtLeyenda1_GotFocus()
    pSelTextBox txtLeyenda1
End Sub

Private Sub txtLeyenda1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLeyenda2_GotFocus()
    pSelTextBox txtLeyenda2
End Sub

Private Sub txtLeyenda2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLeyenda3_GotFocus()
    pSelTextBox txtLeyenda3
End Sub

Private Sub txtLeyenda3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.grdCopiasImpresion.SetFocus
        grdCopiasImpresion.Row = 1
        grdCopiasImpresion.Col = 1
    End If
End Sub
Private Sub txtLeyenda3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtLeyendaCliente_GotFocus()
    pSelTextBox txtLeyendaCliente
End Sub
Private Sub txtLeyendaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtLeyenda1.SetFocus
    End If
End Sub
Private Sub txtLeyendaCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtLeyendaDescuentos_GotFocus()
    pSelTextBox txtLeyendaDescuentos
End Sub
Private Sub txtLeyendaDescuentos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtLeyendaCliente.SetFocus
    End If
End Sub
Private Function fblnCuentaExistente(vlstrCuentaContable As String) As Boolean
On Error GoTo NotificaError
Dim rsCuentaContable As New ADODB.Recordset
    
    fblnCuentaExistente = True
    vlstrsql = "select count(*) from CnCuenta where vchCuentaContable=" + "'" + vlstrCuentaContable + "'" + " and tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable)
    Set rsCuentaContable = frsRegresaRs(vlstrsql)
    If rsCuentaContable.Fields(0) = 0 Then
        fblnCuentaExistente = False
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnCuentaExistente"))
End Function

Private Sub txtLeyendaDescuentos_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombreFactura_GotFocus()
    pSelTextBox txtNombreFactura
End Sub

Private Sub txtNombreFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboConceptoSumaAsegurada_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
End Sub

Private Sub cboImpresoras_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
         cboImpresoraTickets.SetFocus
    End If
    
End Sub

Private Sub txtNombreFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtRFCFactura
    End If
    
End Sub

Private Sub txtNumExterior_GotFocus()
    pSelTextBox txtNumExterior
End Sub

Private Sub txtNumExterior_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtNumInterior
    End If
    
End Sub

Private Sub txtNumInterior_GotFocus()
    pSelTextBox txtNumInterior
End Sub

Private Sub txtNumInterior_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtColoniaPOS
    End If
    
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
    
    If Not fblnFormatoCantidad(txtPorcentajeCoaseguroMPorNota, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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
    
    If Not fblnFormatoCantidad(txtPorcentajeCoaseguroPorNota, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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
        If fblnCanFocus(txtCantidadLimiteCopago) Then
            txtCantidadLimiteCopago.SetFocus
        Else
            If fblnCanFocus(cmdSave) Then cmdSave.SetFocus
        End If
    End If
    
End Sub

Private Sub txtPorcentajeCopagoPorNota_KeyPress(KeyAscii As Integer)
    
    If Not fblnFormatoCantidad(txtPorcentajeCopagoPorNota, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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
    
    If Not fblnFormatoCantidad(txtPorcentajeDeduciblePorNota, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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

    If Not fblnFormatoCantidad(txtPorcentajeExcedentePorNota, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
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

Private Sub txtRFCFactura_GotFocus()
    pSelTextBox txtRFCFactura
End Sub

Private Sub txtRFCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtDireccionPOS
    End If
End Sub

Private Sub txtDireccionPOS_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtNumExterior
    End If
    
End Sub

Private Sub txtLeyenda1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtLeyenda2
    End If
    
End Sub

Private Sub txtLeyenda2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtLeyenda3
    End If
    
End Sub

Private Sub cmdSelecciona_Click(Index As Integer)
On Error GoTo NotificaError
Dim vlintContador As Integer
Dim vlblnElementoExiste As Boolean
Dim vllngCveElemento As Long
    
    If Index = 0 Then
        vlblnElementoExiste = False
        vllngCveElemento = lstListaUsuarios.ItemData(lstListaUsuarios.ListIndex)
        
        For vlintContador = 0 To lstUsuariosAsignados.ListCount - 1
            If lstUsuariosAsignados.ItemData(vlintContador) = vllngCveElemento Then
                vlblnElementoExiste = True
                Exit For
            End If
        Next vlintContador
        
        If vlblnElementoExiste Then
            MsgBox "El usuario ya se encuentra incluido en la lista.", vbExclamation + vbOKOnly, "Mensaje"
        Else
            pSeleccionaLista lstListaUsuarios.ListIndex, lstListaUsuarios, lstUsuariosAsignados, cmdSelecciona(0), cmdSelecciona(1)
        End If
    Else
        If lstUsuariosAsignados.ListIndex <> -1 Then
            pSeleccionaLista lstUsuariosAsignados.ListIndex, lstUsuariosAsignados, lstListaUsuarios, cmdSelecciona(1), cmdSelecciona(0)
        End If
    End If
    
    If Index = 0 Then
        If lstListaUsuarios.ListCount > 0 Then
            lstListaUsuarios.SetFocus
        Else
            cmdSelecciona(1).SetFocus
        End If
    Else
        If lstUsuariosAsignados.ListCount > 0 Then
            lstUsuariosAsignados.SetFocus
        Else
            cmdSelecciona(0).SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSelecciona_Click"))
End Sub

Private Sub lstUsuariosAsignados_DblClick()
    On Error GoTo NotificaError
    
    cmdSelecciona_Click 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstUsuariosAsignados_DblClick"))
End Sub

Private Sub lstUsuariosAsignados_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then lstUsuariosAsignados_DblClick

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstUsuariosAsignados_KeyDown"))
End Sub

Private Sub lstListaUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then lstListaUsuarios_DblClick

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstListaUsuarios_KeyDown"))
End Sub

Private Sub lstListaUsuarios_DblClick()
On Error GoTo NotificaError
    
    cmdSelecciona_Click 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstListaUsuarios_DblClick"))
End Sub

Private Function fintValidaCuenta(vlngNumero As Long) As Integer
    '=========================================================================================
    ' Función para validar la cuenta antes de incluirla en el detalle de la póliza
    ' Regresa también en la variable <vlintOrden> si es una cuenta de orden o no
    '=========================================================================================
 On Error GoTo NotificaError
    Dim rsCuenta As New ADODB.Recordset
    Dim vlstrSentencia As String
   
    ' Valores de regreso (Errores):
    ' 1 = Que la cuenta no acepte movimientos
    ' 2 = Que la fecha de la cuenta sea mayor a la fecha de la póliza
    ' 0 = No hay error
    
    fintValidaCuenta = 0
    
    vlstrSentencia = "select * from CnCuenta where intNumeroCuenta=" & vlngNumero
    Set rsCuenta = frsRegresaRs(vlstrSentencia)
     
    If rsCuenta.RecordCount <> 0 Then
        If rsCuenta!Bitestatusmovimientos = 0 Then
            fintValidaCuenta = 1
        End If
    End If

Exit Function
NotificaError:
   Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintValidaCuenta"))
End Function

Private Sub txtRFCFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTituloCtasPendFact_GotFocus()
    pSelTextBox txtTituloCtasPendFact
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
        cmdSave.SetFocus
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

Private Sub txtTituloCtasPendFact_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSave.SetFocus
End Sub

Private Function fstrVerificaHora(vlstrHora As String)
'Procedimiento para verificar la hora
    On Error GoTo NotificaError
    Dim vldtmHora As Date
    
    If vlstrHora = "  :  " Then
      vldtmHora = fdtmServerHora
    Else
      If CDbl(Mid(vlstrHora, 1, InStr(1, vlstrHora, ":") - 1)) > 23 Then
        vlstrHora = "23" & Mid(vlstrHora, InStr(1, vlstrHora, ":"), Len(vlstrHora))
      End If
      If Mid(vlstrHora, InStr(1, vlstrHora, ":") + 1, Len(vlstrHora)) > 59 Then
        vlstrHora = CDbl(Mid(vlstrHora, 1, InStr(1, vlstrHora, ":") - 1)) & "00"
      End If
      vldtmHora = CDate(vlstrHora)
      fstrVerificaHora = Format(vldtmHora, "hh:mm")
    End If
    
Exit Function
NotificaError:
    If Err.Number = 13 Then
        fstrVerificaHora = ""
        On Error GoTo 0
    End If
End Function
Private Sub pCargaConservarPrecioExclusion()
Dim ObjRS As New ADODB.Recordset
Dim objSTR As String

objSTR = "select vchvalor from siparametro where vchnombre ='BITCONSERVARCOSTOSDESCUENTOEXCLUSION' and INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
Set ObjRS = frsRegresaRs(objSTR, adLockOptimistic)

If ObjRS.RecordCount = 0 Then
   Me.chkConservarPrecioDescuento.Value = vbUnchecked
Else
   Me.chkConservarPrecioDescuento.Value = IIf(ObjRS!VCHVALOR = "1", vbChecked, vbUnchecked)
End If


End Sub
Private Sub pCargaUtilizaSocios()
Dim vlstrsql As String
Dim rsSo As New ADODB.Recordset

    vlstrsql = "select vchvalor, vchnombre, intid, VCHSENTENCIA from siparametro where vchnombre = 'BITUTILIZASOCIOS'"
    Set rsSo = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    With rsSo
        chkSocios.Value = !VCHVALOR
        If !VCHVALOR = 1 Then
            chkValidaDoble.Enabled = True
            If IsNull(!vchSentencia) Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
                !vchSentencia = 1
                .Update
                EntornoSIHO.ConeccionSIHO.CommitTrans
                chkValidaDoble.Value = 1
            Else
                If !vchSentencia = 1 Then
                    chkValidaDoble.Value = 1
                Else
                    chkValidaDoble.Value = 0
                End If
            End If
        Else
            chkValidaDoble.Enabled = False
        End If
        .Close
    End With
    
    vlstrsql = "select * from cccargoexcluido cc where cc.INTTIPOPACIENTE = 1"
    Set rsSo = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    If rsSo.RecordCount > 0 Then
        chkSocios.Enabled = False
    End If
    
End Sub
Private Sub pCargaCantidadImpresiones()
Dim ObjRS As New ADODB.Recordset
Dim objSTR As String

With grdCopiasImpresion
     
     .Clear
     .Rows = 5
     .Cols = 2
     .FixedCols = 1
     .FixedRows = 1
     .FormatString = "Documento|Copias"
     .ColWidth(0) = 4500 'impresión
     .ColWidth(1) = 2600  'copias
     .Col = 0
     .Row = 0
     .CellAlignment = flexAlignCenterCenter
     .Col = 1
     .Row = 0
     .CellAlignment = flexAlignCenterCenter
     .TextMatrix(1, 0) = "Factura del paciente"
     .TextMatrix(2, 0) = "Factura de la empresa"
     .TextMatrix(3, 0) = "Factura directa"
     .TextMatrix(4, 0) = "Tickets"
     .TextMatrix(1, 1) = "2"
     .TextMatrix(2, 1) = "2"
     .TextMatrix(3, 1) = "2"
     .TextMatrix(4, 1) = "2"

     objSTR = "Select vchnombre,vchvalor from siparametro where intcveempresacontable = " & vgintClaveEmpresaContable & " and vchnombre in ('NUMCOPIASFACTURAPACIENTE','NUMCOPIASFACTURAEMPRESA','NUMCOPIASFACTURADIRECTA', 'NUMCOPIASTICKET')"
     Set ObjRS = frsRegresaRs(objSTR, adLockOptimistic)
    If ObjRS.RecordCount > 0 Then
     ObjRS.MoveFirst
     Do While Not ObjRS.EOF
        Select Case ObjRS!vchNombre
        Case "NUMCOPIASFACTURAPACIENTE"
              .TextMatrix(1, 1) = ObjRS!VCHVALOR
        Case "NUMCOPIASFACTURAEMPRESA"
             .TextMatrix(2, 1) = ObjRS!VCHVALOR
        Case "NUMCOPIASFACTURADIRECTA"
             .TextMatrix(3, 1) = ObjRS!VCHVALOR
        Case "NUMCOPIASTICKET"
             .TextMatrix(4, 1) = ObjRS!VCHVALOR
        End Select
     ObjRS.MoveNext
     Loop
    End If
End With
End Sub
Private Sub txtEditCol_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    'Para verificar que tecla fue presionada en el textbox
    With grdCopiasImpresion
        Select Case KeyCode
            Case 27   'ESC
                .SetFocus
                txtEditCol.Visible = False
                KeyCode = 0
                vlblnEscTxtEditCOl = True
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    vgblnNoEditar = True
                    .Row = .Row - 1
                    vgblnNoEditar = False
                End If
                vlblnEscTxtEditCOl = False
            Case 40, 13
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    vgblnNoEditar = True
                    .Row = .Row + 1
                    vgblnNoEditar = False
                End If
                vlblnEscTxtEditCOl = False
        End Select
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtEditCol_KeyDown"))
End Sub
Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError
    Dim vlintTexto As Integer
    Dim blnmuestra As Boolean
    blnmuestra = False
    With txtEdit
    .Text = grid
        Select Case KeyAscii
            Case 0 To 32
                'Edita el texto de la celda en la que está posicionado
                .SelStart = 0
                .SelLength = 1000
                       
                
                blnmuestra = True
                
            Case 8, 48 To 57
                ' Reemplaza el texto actual solo si se teclean números
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
                blnmuestra = True
        End Select
    End With
            If blnmuestra Then
    ' Muestra el textbox en el lugar indicado
    With grid
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
       
    txtEdit.Visible = True
    txtEdit.SetFocus
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEditarColumna"))
End Sub
Private Sub pSetCellValueCol(grid As MSHFlexGrid, txtEdit As TextBox)
    On Error GoTo NotificaError
    If grid.Col = 1 Then
       If txtEditCol.Visible Then
          If txtEditCol.Text <> "" Then
             If IsNumeric(txtEditCol.Text) And Val(txtEditCol.Text) > 0 Then
                grid.Text = Val(txtEditCol.Text)
             End If
          End If
          txtEditCol.Visible = False
          txtEditCol.Text = ""
       End If
    Else
        Exit Sub
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSetCellValueCol"))
End Sub

Private Sub pCargaUsosCFDI()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frsCatalogoSAT("c_UsoCFDI")
    If Not rsTmp.EOF Then
        pLlenarCboRs cboUsoCFDIFacturado, rsTmp, 0, 1
        cboUsoCFDIFacturado.ListIndex = -1
    End If
End Sub

Public Sub pCargaConceptoFacturacionAsistSocial()
On Error GoTo NotificaError
Dim vlstrSentencia As String
Dim rsConcepto As New ADODB.Recordset


    vlstrSentencia = "SELECT smiCveConcepto, chrDescripcion " & _
                     "FROM PvConceptoFacturacion " & _
                     "WHERE bitActivo = 1 " & _
                     "ORDER BY chrDescripcion"
    Set rsConcepto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    pLlenarCboRs cboConceptoFacturacionAsistSocial, rsConcepto, 0, 1
    cboConceptoFacturacionAsistSocial.AddItem " ", 0
    cboConceptoFacturacionAsistSocial.ListIndex = 0

    rsConcepto.Close

    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaConceptoFacturacionAsistSocial"))
End Sub

Public Sub pCargaAsistSocial()
Dim vlstrSentencia As String
Dim rsConcepto As New ADODB.Recordset


    vlstrSentencia = "SELECT smiCveConcepto, chrDescripcion " & _
                     "FROM PvConceptoFacturacion " & _
                     "WHERE bitActivo = 1 " & _
                     "ORDER BY chrDescripcion"
    Set rsConcepto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    pLlenarCboRs cboConceptoFacturacionAsistSocial, rsConcepto, 0, 1
    cboConceptoFacturacionAsistSocial.AddItem " ", 0
    cboConceptoFacturacionAsistSocial.ListIndex = 0

    rsConcepto.Close

    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaAsistSocial"))
End Sub
Private Sub chkValidacionPMPVentaPublico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mskHoraIniMsgCargo.SetFocus
        pSelMkTexto mskHoraIniMsgCargo
    End If
End Sub
