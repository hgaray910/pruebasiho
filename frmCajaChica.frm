VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCajaChica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salidas de caja chica"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   11000
      Left            =   0
      TabIndex        =   93
      Top             =   -360
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   19394
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmCajaChica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDisminucion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraHonorario"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBuscarXML"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SysInfo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraTipo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraBotonera"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraFactura"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraSelXMLCajaChicaFact"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraSelXMLCajaChicaHono"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmCajaChica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraConsulta"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraSelXMLCajaChicaHono 
         Height          =   685
         Left            =   9405
         TabIndex        =   169
         Top             =   6375
         Width           =   2460
         Begin VB.CommandButton cmdBuscarXMLHonorario 
            Height          =   480
            Left            =   1920
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCajaChica.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Comprobante del documento"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   480
         End
         Begin VB.OptionButton optTipoComproCajaChicaHono 
            Caption         =   "CFD / CFDI"
            Height          =   255
            Index           =   0
            Left            =   80
            TabIndex        =   22
            ToolTipText     =   "Comprobante fiscal digital"
            Top             =   130
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optTipoComproCajaChicaHono 
            Caption         =   "CBB"
            Height          =   255
            Index           =   1
            Left            =   80
            TabIndex        =   23
            ToolTipText     =   "Código de barras bidimensional"
            Top             =   400
            Width           =   615
         End
         Begin VB.OptionButton optTipoComproCajaChicaHono 
            Caption         =   "Extranjero"
            Height          =   255
            Index           =   2
            Left            =   760
            TabIndex        =   24
            ToolTipText     =   "Comprobante extranjero"
            Top             =   400
            Width           =   1095
         End
      End
      Begin VB.Frame fraSelXMLCajaChicaFact 
         Height          =   685
         Left            =   9380
         TabIndex        =   168
         Top             =   6840
         Width           =   2460
         Begin VB.OptionButton optTipoComproCajaChicaFact 
            Caption         =   "Extranjero"
            Height          =   255
            Index           =   2
            Left            =   760
            TabIndex        =   73
            ToolTipText     =   "Comprobante extranjero"
            Top             =   400
            Width           =   1095
         End
         Begin VB.OptionButton optTipoComproCajaChicaFact 
            Caption         =   "CBB"
            Height          =   255
            Index           =   1
            Left            =   80
            TabIndex        =   72
            ToolTipText     =   "Código de barras bidimensional"
            Top             =   400
            Width           =   650
         End
         Begin VB.OptionButton optTipoComproCajaChicaFact 
            Caption         =   "CFD / CFDI"
            Height          =   255
            Index           =   0
            Left            =   80
            TabIndex        =   71
            ToolTipText     =   "Comprobante fiscal digital"
            Top             =   130
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CommandButton cmdBuscarXMLFactura 
            Height          =   480
            Left            =   1920
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCajaChica.frx":073A
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Comprobante del documento"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   480
         End
      End
      Begin VB.Frame fraFactura 
         Height          =   5605
         Left            =   135
         TabIndex        =   101
         Top             =   2025
         Width           =   11835
         Begin VB.TextBox txtDescSalida 
            Height          =   715
            Left            =   1920
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   45
            ToolTipText     =   "Descripción de la salida"
            Top             =   3480
            Width           =   4560
         End
         Begin VB.Frame fraDocTicketNota 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   6840
            TabIndex        =   157
            Top             =   120
            Visible         =   0   'False
            Width           =   4935
            Begin VB.TextBox txtTotalTicket 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3285
               TabIndex        =   70
               ToolTipText     =   "Total"
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Left            =   60
               TabIndex        =   158
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame fraDocFlete 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   6840
            TabIndex        =   159
            Top             =   120
            Visible         =   0   'False
            Width           =   4935
            Begin VB.Frame Frame7 
               Height          =   415
               Left            =   1635
               TabIndex        =   163
               Top             =   840
               Width           =   1600
               Begin VB.OptionButton optRetencion 
                  Caption         =   "No"
                  Height          =   195
                  Index           =   1
                  Left            =   840
                  TabIndex        =   69
                  ToolTipText     =   "Sin retención"
                  Top             =   150
                  Width           =   615
               End
               Begin VB.OptionButton optRetencion 
                  Caption         =   "Sí"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   68
                  ToolTipText     =   "Con retención"
                  Top             =   150
                  Width           =   615
               End
            End
            Begin VB.ComboBox cboImpuestoFlete 
               Height          =   315
               Left            =   1630
               Style           =   2  'Dropdown List
               TabIndex        =   65
               ToolTipText     =   "Impuesto"
               Top             =   520
               Width           =   1610
            End
            Begin VB.TextBox txtImporteFlete 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3280
               MaxLength       =   10
               TabIndex        =   64
               ToolTipText     =   "Importe"
               Top             =   120
               Width           =   1560
            End
            Begin VB.Label lblRetencion 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3285
               TabIndex        =   165
               ToolTipText     =   "Retención"
               Top             =   935
               Width           =   1560
            End
            Begin VB.Label lblTotalFlete 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3285
               TabIndex        =   67
               ToolTipText     =   "Total"
               Top             =   1340
               Width           =   1560
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Left            =   60
               TabIndex        =   164
               Top             =   1400
               Width           =   360
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Retención"
               Height          =   195
               Left            =   60
               TabIndex        =   162
               Top             =   995
               Width           =   735
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Impuesto"
               Height          =   195
               Left            =   60
               TabIndex        =   161
               Top             =   580
               Width           =   645
            End
            Begin VB.Label lblImporteIvaFlete 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3285
               TabIndex        =   66
               ToolTipText     =   "Impuesto"
               Top             =   520
               Width           =   1560
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Importe"
               Height          =   195
               Left            =   60
               TabIndex        =   160
               Top             =   180
               Width           =   525
            End
         End
         Begin VB.Frame fraDocFactura 
            BorderStyle     =   0  'None
            Height          =   4725
            Left            =   6780
            TabIndex        =   147
            Top             =   120
            Width           =   5000
            Begin VB.ComboBox cboRetencionISR 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1710
               Style           =   2  'Dropdown List
               TabIndex        =   62
               ToolTipText     =   "Retención de ISR"
               Top             =   3990
               Width           =   1610
            End
            Begin VB.TextBox txtIEPS 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   53
               ToolTipText     =   "Monto del IEPS"
               Top             =   2190
               Width           =   1560
            End
            Begin VB.Frame fraFacturaFlete 
               BorderStyle     =   0  'None
               Height          =   1100
               Left            =   80
               TabIndex        =   170
               Top             =   2880
               Width           =   4935
               Begin VB.ComboBox cboImpuestoFleteFac 
                  Height          =   315
                  Left            =   1640
                  Style           =   2  'Dropdown List
                  TabIndex        =   57
                  ToolTipText     =   "Impuesto del flete"
                  Top             =   345
                  Width           =   1610
               End
               Begin VB.TextBox txtFleteFactura 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3280
                  TabIndex        =   56
                  ToolTipText     =   "Importe del flete"
                  Top             =   0
                  Width           =   1575
               End
               Begin VB.Frame fraRetencionSiNo 
                  Height          =   415
                  Left            =   1640
                  TabIndex        =   171
                  Top             =   640
                  Width           =   1600
                  Begin VB.OptionButton optRetencionFactura 
                     Caption         =   "No"
                     Height          =   195
                     Index           =   1
                     Left            =   840
                     TabIndex        =   60
                     ToolTipText     =   "Sin retención"
                     Top             =   150
                     Width           =   615
                  End
                  Begin VB.OptionButton optRetencionFactura 
                     Caption         =   "Sí"
                     Height          =   195
                     Index           =   0
                     Left            =   120
                     TabIndex        =   59
                     ToolTipText     =   "Con retención"
                     Top             =   150
                     Width           =   495
                  End
               End
               Begin VB.Label lblImpuestoFlete 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   3285
                  TabIndex        =   58
                  ToolTipText     =   "Impuesto del flete"
                  Top             =   345
                  Width           =   1560
               End
               Begin VB.Label lblTituloImpuestoFlete 
                  Caption         =   "Impuesto"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   175
                  Top             =   405
                  Width           =   1245
               End
               Begin VB.Label Label30 
                  AutoSize        =   -1  'True
                  Caption         =   "Flete"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   174
                  Top             =   60
                  Width           =   345
               End
               Begin VB.Label lblRetencionFactura 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   3285
                  TabIndex        =   173
                  ToolTipText     =   "Retención"
                  Top             =   720
                  Width           =   1560
               End
               Begin VB.Label lblRetencionSiNo 
                  AutoSize        =   -1  'True
                  Caption         =   "Retención"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   172
                  Top             =   780
                  Width           =   735
               End
            End
            Begin VB.ComboBox cboImpuesto 
               Height          =   315
               Left            =   1710
               Style           =   2  'Dropdown List
               TabIndex        =   54
               ToolTipText     =   "Impuesto"
               Top             =   2550
               Width           =   1610
            End
            Begin VB.TextBox txtImporteNoGravado 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   48
               ToolTipText     =   "Monto gravado al 0%"
               Top             =   810
               Width           =   1560
            End
            Begin VB.TextBox txtDescuentoNoGravado 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   49
               ToolTipText     =   "Descuento"
               Top             =   1155
               Width           =   1560
            End
            Begin VB.TextBox txtImporteGravado 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   50
               ToolTipText     =   "Monto gravado"
               Top             =   1500
               Width           =   1560
            End
            Begin VB.TextBox txtDescuentoGravado 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   51
               ToolTipText     =   "Descuento"
               Top             =   1845
               Width           =   1560
            End
            Begin VB.CheckBox chkIEPSBaseGravable 
               Caption         =   "Base gravable"
               Height          =   255
               Left            =   1710
               TabIndex        =   52
               ToolTipText     =   "Indica si el IEPS aplica para la base gravable"
               Top             =   2220
               Width           =   1455
            End
            Begin VB.TextBox txtDescuentoExento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   47
               ToolTipText     =   "Descuento"
               Top             =   465
               Width           =   1560
            End
            Begin VB.TextBox txtImporteExento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   46
               ToolTipText     =   "Monto exento"
               Top             =   120
               Width           =   1560
            End
            Begin VB.CheckBox chkRetencionISR 
               Caption         =   "Retención ISR"
               Height          =   480
               Left            =   120
               TabIndex        =   61
               ToolTipText     =   "Retención ISR"
               Top             =   3900
               Width           =   1400
            End
            Begin VB.Label lblRetencionISR 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3360
               TabIndex        =   177
               ToolTipText     =   "Retención de ISR"
               Top             =   3990
               Width           =   1560
            End
            Begin VB.Label lblMontoGravado0 
               AutoSize        =   -1  'True
               Caption         =   "Monto gravado al 0%"
               Height          =   195
               Left            =   120
               TabIndex        =   156
               Top             =   870
               Width           =   1500
            End
            Begin VB.Label lblDescuentoNoGravado 
               AutoSize        =   -1  'True
               Caption         =   "Descuento"
               Height          =   195
               Left            =   120
               TabIndex        =   155
               Top             =   1215
               Width           =   780
            End
            Begin VB.Label lblMontoGravado 
               AutoSize        =   -1  'True
               Caption         =   "Monto gravado"
               Height          =   195
               Left            =   120
               TabIndex        =   154
               Top             =   1560
               Width           =   1080
            End
            Begin VB.Label lblDescuentoGravado 
               AutoSize        =   -1  'True
               Caption         =   "Descuento"
               Height          =   195
               Left            =   120
               TabIndex        =   153
               Top             =   1905
               Width           =   780
            End
            Begin VB.Label lblTituloImpuesto 
               Caption         =   "Impuesto"
               Height          =   195
               Left            =   120
               TabIndex        =   152
               Top             =   2610
               Width           =   1245
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Total factura"
               Height          =   195
               Left            =   120
               TabIndex        =   151
               Top             =   4440
               Width           =   900
            End
            Begin VB.Label lblImporteIVA 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3360
               TabIndex        =   55
               ToolTipText     =   "Impuesto"
               Top             =   2550
               Width           =   1560
            End
            Begin VB.Label lblIeps 
               AutoSize        =   -1  'True
               Caption         =   "IEPS"
               Height          =   195
               Left            =   120
               TabIndex        =   150
               Top             =   2250
               Width           =   360
            End
            Begin VB.Label lblDescuentoExento 
               AutoSize        =   -1  'True
               Caption         =   "Descuento"
               Height          =   195
               Left            =   120
               TabIndex        =   149
               Top             =   525
               Width           =   780
            End
            Begin VB.Label lblMontoExcento 
               AutoSize        =   -1  'True
               Caption         =   "Monto exento"
               Height          =   195
               Left            =   120
               TabIndex        =   148
               Top             =   180
               Width           =   975
            End
            Begin VB.Label lblTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3360
               TabIndex        =   63
               ToolTipText     =   "Total factura"
               Top             =   4365
               Width           =   1560
            End
         End
         Begin VB.TextBox txtFolio 
            Height          =   315
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   38
            ToolTipText     =   "Folio"
            Top             =   2340
            Width           =   1470
         End
         Begin VB.TextBox txtRFC 
            Height          =   315
            Left            =   1920
            MaxLength       =   13
            TabIndex        =   33
            ToolTipText     =   "RFC del Proveedor/Acreedor"
            Top             =   580
            Width           =   2055
         End
         Begin VB.ComboBox cboTipoProveedor 
            Height          =   315
            ItemData        =   "frmCajaChica.frx":0E3C
            Left            =   1920
            List            =   "frmCajaChica.frx":0E49
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Tipo del proveedor"
            Top             =   940
            Width           =   2055
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            ItemData        =   "frmCajaChica.frx":0E6A
            Left            =   1920
            List            =   "frmCajaChica.frx":0E77
            Style           =   2  'Dropdown List
            TabIndex        =   35
            ToolTipText     =   "País del proveedor"
            Top             =   1300
            Width           =   2055
         End
         Begin VB.CheckBox chkXMLrelacionadoFact 
            Caption         =   "Factura con comprobante relacionado"
            Enabled         =   0   'False
            Height          =   480
            Left            =   6900
            TabIndex        =   80
            ToolTipText     =   "Se le ha relacionado satisfactoriamente el comprobante a la factura"
            Top             =   4860
            Width           =   2175
         End
         Begin VB.Frame Frame6 
            Height          =   415
            Left            =   1920
            TabIndex        =   112
            Top             =   2610
            Width           =   2100
            Begin VB.OptionButton optMoneda 
               Caption         =   "Dólares"
               Height          =   190
               Index           =   1
               Left            =   1070
               TabIndex        =   40
               ToolTipText     =   "Dólares"
               Top             =   160
               Width           =   840
            End
            Begin VB.OptionButton optMoneda 
               Caption         =   "Pesos"
               Height          =   190
               Index           =   0
               Left            =   120
               TabIndex        =   39
               ToolTipText     =   "Pesos"
               Top             =   160
               Width           =   735
            End
         End
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   315
            Left            =   1920
            TabIndex        =   37
            ToolTipText     =   "Fecha"
            Top             =   1995
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.ComboBox cboConcepto 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   36
            ToolTipText     =   "Concepto"
            Top             =   1650
            Width           =   4680
         End
         Begin VB.ComboBox cboProveedor 
            Height          =   315
            Left            =   1920
            TabIndex        =   32
            Text            =   "cboProveedor"
            ToolTipText     =   "Proveedor"
            Top             =   230
            Width           =   4680
         End
         Begin VB.Frame Frame2 
            Height          =   415
            Left            =   1920
            TabIndex        =   113
            Top             =   2980
            Width           =   4575
            Begin VB.OptionButton optFlete 
               Caption         =   "Factura de flete"
               Height          =   190
               Left            =   2880
               TabIndex        =   44
               ToolTipText     =   "Factura de flete"
               Top             =   160
               Width           =   1455
            End
            Begin VB.OptionButton optTicket 
               Caption         =   "Ticket"
               Height          =   190
               Left            =   1070
               TabIndex        =   42
               ToolTipText     =   "Ticket"
               Top             =   160
               Width           =   840
            End
            Begin VB.OptionButton optNota 
               Caption         =   "Nota"
               Height          =   190
               Left            =   2050
               TabIndex        =   43
               ToolTipText     =   "Nota"
               Top             =   160
               Width           =   735
            End
            Begin VB.OptionButton optFactura 
               Caption         =   "Factura"
               Height          =   190
               Left            =   120
               TabIndex        =   41
               ToolTipText     =   "Factura"
               Top             =   160
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Descripción de la salida"
            Height          =   195
            Left            =   120
            TabIndex        =   166
            Top             =   3540
            Width           =   1680
         End
         Begin VB.Label Label22 
            Caption         =   "RFC"
            Height          =   195
            Left            =   120
            TabIndex        =   146
            Top             =   640
            Width           =   1530
         End
         Begin VB.Label Label23 
            Caption         =   "Tipo"
            Height          =   195
            Left            =   120
            TabIndex        =   145
            Top             =   1000
            Width           =   495
         End
         Begin VB.Label Label24 
            Caption         =   "País"
            Height          =   195
            Left            =   120
            TabIndex        =   144
            Top             =   1360
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Tipo de documento"
            Height          =   195
            Left            =   120
            TabIndex        =   114
            Top             =   3140
            Width           =   1575
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   300
            Left            =   120
            TabIndex        =   111
            Top             =   2730
            Width           =   1545
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Folio"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   2400
            Width           =   330
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fecha documento"
            Height          =   195
            Left            =   120
            TabIndex        =   104
            Top             =   2060
            Width           =   1650
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Concepto salida"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   1710
            Width           =   1140
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor / acreedor"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   290
            Width           =   1530
         End
      End
      Begin VB.Frame FraBotonera 
         Height          =   705
         Left            =   4215
         TabIndex        =   137
         Top             =   8000
         Width           =   3600
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   45
            Picture         =   "frmCajaChica.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "Primer registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   540
            Picture         =   "frmCajaChica.frx":129A
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "Anterior registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1035
            Picture         =   "frmCajaChica.frx":140C
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1545
            Picture         =   "frmCajaChica.frx":157E
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "Siguiente registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2055
            Picture         =   "frmCajaChica.frx":16F0
            Style           =   1  'Graphical
            TabIndex        =   139
            ToolTipText     =   "Ultimo registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2550
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCajaChica.frx":1BE2
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Grabar"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdCancelar 
            Height          =   495
            Left            =   3045
            Picture         =   "frmCajaChica.frx":1F24
            Style           =   1  'Graphical
            TabIndex        =   138
            ToolTipText     =   "Cancelar salida"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame fraTipo 
         Height          =   720
         Left            =   6330
         TabIndex        =   132
         Top             =   460
         Width           =   5510
         Begin VB.OptionButton optTipo 
            Caption         =   "Factura"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   135
            ToolTipText     =   "Factura"
            Top             =   290
            Width           =   1095
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Honorario"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   134
            ToolTipText     =   "Honorario"
            Top             =   290
            Width           =   1095
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Disminución de fondo"
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   133
            ToolTipText     =   "Disminución de fondo"
            Top             =   290
            Width           =   1935
         End
      End
      Begin SysInfoLib.SysInfo SysInfo 
         Left            =   480
         Top             =   5160
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cmdBuscarXML 
         Left            =   1440
         Top             =   5180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame FraConsulta 
         Height          =   5400
         Left            =   -74895
         TabIndex        =   110
         Top             =   1470
         Width           =   11895
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFacturas 
            Height          =   5205
            Left            =   30
            TabIndex        =   92
            ToolTipText     =   "Salidas de caja chica"
            Top             =   120
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   9181
            _Version        =   393216
            Cols            =   8
            GridColor       =   12632256
            FormatString    =   "|Fecha|Número|Proveedor|Factura|Estado|Registró|Canceló"
            _NumberOfBands  =   1
            _Band(0).Cols   =   8
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   -74895
         TabIndex        =   106
         Top             =   360
         Width           =   11895
         Begin VB.CheckBox chkTodas 
            Caption         =   "Todas"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            ToolTipText     =   "Mostrar todas"
            Top             =   720
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkSinDepositar 
            Caption         =   "Sin depositar"
            Enabled         =   0   'False
            Height          =   255
            Left            =   8880
            TabIndex        =   90
            ToolTipText     =   "Mostrar sin depositar"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkDepositadas 
            Caption         =   "Depositadas"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4560
            TabIndex        =   87
            ToolTipText     =   "Mostrar depositadas"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkPendientes 
            Caption         =   "Pendientes"
            Enabled         =   0   'False
            Height          =   255
            Left            =   6000
            TabIndex        =   88
            ToolTipText     =   "Mostrar pendientes"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkReembolsadas 
            Caption         =   "Reembolsadas"
            Enabled         =   0   'False
            Height          =   255
            Left            =   7440
            TabIndex        =   89
            ToolTipText     =   "Mostrar reembolsadas"
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkCanceladas 
            Caption         =   "Canceladas"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3120
            TabIndex        =   86
            ToolTipText     =   "Mostrar canceladas"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkActivas 
            Caption         =   "Activas"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   85
            ToolTipText     =   "Mostrar activas"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Cargar"
            Height          =   315
            Left            =   10335
            TabIndex        =   91
            ToolTipText     =   "Cargar información"
            Top             =   240
            Width           =   1425
         End
         Begin VB.ComboBox cboProveedorBus 
            Height          =   315
            Left            =   5250
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   240
            Width           =   4995
         End
         Begin MSMask.MaskEdBox mskFechaBusIni 
            Height          =   315
            Left            =   795
            TabIndex        =   81
            ToolTipText     =   "Fecha inicial"
            Top             =   240
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaBusFin 
            Height          =   315
            Left            =   2910
            TabIndex        =   82
            ToolTipText     =   "Fecha final"
            Top             =   240
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label lbProveedor 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   1
            Left            =   4380
            TabIndex        =   109
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2340
            TabIndex        =   108
            Top             =   300
            Width           =   420
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   210
            TabIndex        =   107
            Top             =   300
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1680
         Left            =   135
         TabIndex        =   94
         Top             =   345
         Width           =   11835
         Begin VB.TextBox txtNumero 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1920
            MaxLength       =   9
            TabIndex        =   0
            ToolTipText     =   "Número de factura"
            Top             =   195
            Width           =   1320
         End
         Begin VB.Label lblPersonaCancelaReembolsa 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7830
            TabIndex        =   5
            ToolTipText     =   "Persona que canceló / reembolsó"
            Top             =   1230
            Width           =   3870
         End
         Begin VB.Label lblTituloCancela 
            AutoSize        =   -1  'True
            Caption         =   "Canceló / Reembolsó"
            Height          =   195
            Left            =   6200
            TabIndex        =   100
            Top             =   1290
            Width           =   1545
         End
         Begin VB.Label lblEstado 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7830
            TabIndex        =   4
            ToolTipText     =   "Estado"
            Top             =   885
            Width           =   3870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   6200
            TabIndex        =   99
            Top             =   945
            Width           =   495
         End
         Begin VB.Label lblPersonaRegistra 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   3
            ToolTipText     =   "Persona que registró"
            Top             =   1230
            Width           =   4215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Registró"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   1290
            Width           =   585
         End
         Begin VB.Label lblDepartamento 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   2
            ToolTipText     =   "Departamento"
            Top             =   885
            Width           =   4215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   945
            Width           =   1005
         End
         Begin VB.Label lblFecha 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   1
            ToolTipText     =   "Fecha"
            Top             =   540
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   255
            Width           =   555
         End
      End
      Begin VB.Frame fraHonorario 
         Height          =   4450
         Left            =   135
         TabIndex        =   115
         Top             =   2025
         Width           =   11835
         Begin VB.TextBox txtRFcHono 
            Height          =   315
            Left            =   1920
            MaxLength       =   13
            TabIndex        =   9
            ToolTipText     =   "RFC del médico"
            Top             =   885
            Width           =   2325
         End
         Begin VB.TextBox txtDescSalidaHono 
            Height          =   315
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   14
            ToolTipText     =   "Descripción de la salida"
            Top             =   2400
            Width           =   4680
         End
         Begin VB.OptionButton OptIVAHonorario 
            Caption         =   "Exento"
            Height          =   195
            Index           =   1
            Left            =   7560
            TabIndex        =   17
            ToolTipText     =   "Exento"
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton OptIVAHonorario 
            Caption         =   "IVA"
            Height          =   195
            Index           =   0
            Left            =   6900
            TabIndex        =   16
            ToolTipText     =   "IVA"
            Top             =   945
            Width           =   615
         End
         Begin VB.Frame Frame3 
            Height          =   415
            Left            =   1920
            TabIndex        =   125
            Top             =   1900
            Width           =   2100
            Begin VB.OptionButton optMonedaHonorario 
               Caption         =   "Dólares"
               Height          =   190
               Index           =   1
               Left            =   1070
               TabIndex        =   13
               Top             =   160
               Width           =   975
            End
            Begin VB.OptionButton optMonedaHonorario 
               Caption         =   "Pesos"
               Height          =   190
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   160
               Width           =   975
            End
         End
         Begin VB.CheckBox chkXMLrelacionadoHono 
            Caption         =   "Honorario con comprobante relacionado"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6900
            TabIndex        =   31
            ToolTipText     =   "Se le ha relacionado satisfactoriamente el comprobante al honorario"
            Top             =   2735
            Width           =   2285
         End
         Begin VB.ComboBox cboMedicos 
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Text            =   "cboMedicos"
            ToolTipText     =   "Médico"
            Top             =   540
            Width           =   4680
         End
         Begin VB.TextBox txtCuentaHonorario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4275
            Locked          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Descripción cuenta contable"
            Top             =   195
            Width           =   7425
         End
         Begin VB.ComboBox cboTarifa 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8485
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   "Selección de la tasa de IVA"
            Top             =   1560
            Width           =   1610
         End
         Begin VB.TextBox txtFolioHonorario 
            Height          =   315
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   11
            ToolTipText     =   "Folio"
            Top             =   1575
            Width           =   1470
         End
         Begin VB.TextBox txtMontoHonorario 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10145
            MaxLength       =   14
            TabIndex        =   15
            ToolTipText     =   "Monto del honorario"
            Top             =   540
            Width           =   1560
         End
         Begin VB.ComboBox cboIVAHonorario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8485
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Selección de la tasa de IVA"
            Top             =   885
            Width           =   1610
         End
         Begin VB.CheckBox chkRetencionISRHonorario 
            Caption         =   "Retención de ISR"
            Enabled         =   0   'False
            Height          =   195
            Left            =   6900
            TabIndex        =   19
            ToolTipText     =   "Retención del ISR"
            Top             =   1620
            Width           =   1570
         End
         Begin VB.CheckBox chkRetencionIVAHonorario 
            Caption         =   "Retención de IVA"
            Enabled         =   0   'False
            Height          =   195
            Left            =   6900
            TabIndex        =   21
            ToolTipText     =   "Retención del IVA"
            Top             =   1995
            Width           =   1650
         End
         Begin MSMask.MaskEdBox mskFechaHonorario 
            Height          =   315
            Left            =   1920
            TabIndex        =   10
            ToolTipText     =   "Fecha del honorario"
            Top             =   1230
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCuentaHonorario 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            ToolTipText     =   "Cuenta contable"
            Top             =   195
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label lblRFcHono 
            Caption         =   "RFC"
            Height          =   195
            Left            =   120
            TabIndex        =   176
            Top             =   945
            Width           =   1530
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Descripción de la salida"
            Height          =   195
            Left            =   120
            TabIndex        =   167
            Top             =   2450
            Width           =   1680
         End
         Begin VB.Label lbMedicos 
            AutoSize        =   -1  'True
            Caption         =   "Médico"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   124
            Top             =   600
            Width           =   525
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   122
            Top             =   2050
            Width           =   585
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Folio"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   121
            Top             =   1625
            Width           =   330
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Fecha del recibo"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   120
            Top             =   1290
            Width           =   1185
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta contable"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   119
            Top             =   255
            Width           =   1170
         End
         Begin VB.Label lblIVAHonorario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10145
            TabIndex        =   26
            ToolTipText     =   "IVA del honorario"
            Top             =   885
            Width           =   1560
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Left            =   6900
            TabIndex        =   118
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Subtotal"
            Height          =   195
            Left            =   6900
            TabIndex        =   117
            Top             =   1290
            Width           =   585
         End
         Begin VB.Label lblSubtotalHonorario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10145
            TabIndex        =   27
            ToolTipText     =   "Subtotal"
            Top             =   1230
            Width           =   1560
         End
         Begin VB.Label lblRetencionISRHonorario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10145
            TabIndex        =   28
            ToolTipText     =   "Retención del ISR"
            Top             =   1560
            Width           =   1560
         End
         Begin VB.Label lblRetencionIVAHonorario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10145
            TabIndex        =   29
            ToolTipText     =   "Retención del IVA"
            Top             =   1935
            Width           =   1560
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "Total a pagar"
            Height          =   195
            Left            =   6900
            TabIndex        =   116
            Top             =   2340
            Width           =   945
         End
         Begin VB.Label lblTotalPagarHonorario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10145
            TabIndex        =   30
            ToolTipText     =   "Total a pagar"
            Top             =   2280
            Width           =   1560
         End
      End
      Begin VB.Frame fraDisminucion 
         Height          =   2295
         Left            =   135
         TabIndex        =   126
         Top             =   2040
         Visible         =   0   'False
         Width           =   11835
         Begin VB.TextBox txtTotalDF 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   78
            ToolTipText     =   "Total de la disminución de fondo"
            Top             =   1125
            Width           =   1575
         End
         Begin VB.TextBox txtDescripcionDF 
            Height          =   315
            Left            =   7725
            TabIndex        =   136
            ToolTipText     =   "Descripción de la disminución de fondo"
            Top             =   -840
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Frame Frame5 
            Height          =   415
            Left            =   1920
            TabIndex        =   127
            Top             =   615
            Width           =   2100
            Begin VB.OptionButton OptMonedaDF 
               Caption         =   "Dólares"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   77
               Top             =   120
               Width           =   855
            End
            Begin VB.OptionButton OptMonedaDF 
               Caption         =   "Pesos"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   76
               Top             =   120
               Width           =   735
            End
         End
         Begin MSMask.MaskEdBox MaskDisminucionFecha 
            Height          =   315
            Left            =   1920
            TabIndex        =   75
            ToolTipText     =   "Fecha de la disminución de fondo"
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Total a devolver"
            Height          =   195
            Left            =   120
            TabIndex        =   131
            Top             =   1200
            Width           =   1155
         End
         Begin VB.Label Label28 
            Caption         =   "Fecha del documento"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label Label20 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   6240
            TabIndex        =   129
            Top             =   -780
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   128
            Top             =   765
            Width           =   585
         End
      End
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Fecha del recibo"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   123
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "frmCajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmCajaChica
'-------------------------------------------------------------------------------------

Option Explicit

Private Type typIVA
    lngIdImpuesto As Long
    dblPorcentaje As Double
End Type

Private Type typConcepto
    lngIdConcepto As Long
    lngCtaGasto As Long
End Type

Const cstrFormato = "############.##"

'Búsqueda
Const cintColFecha = 1
Const cintColNumero = 2
Const cintColProveedor = 3
Const cintColFactura = 4
Const cIntColEstado = 5
Const cintColRegistro = 6
Const cintColCancelo = 7
Const cintColumnas = 8
Const cstrTitulos = "|Fecha|Número|Proveedor|Factura|Estado|Registró|Canceló/Reembolsó"

Const clngColorCanceladas = &HC0&
Const clngColorActivas = &H80000012
Const clngColorPendientes = &HC00000        '&H80000012&

Dim arrImpuestosFactura() As typIVA
Dim arrImpuestosHonorario() As typIVA
Dim arrRetencionISR() As typIVA
Dim arrConceptos() As typConcepto
Dim aFormasPago() As FormasPago

Dim rs As New ADODB.Recordset 'Varios usos
Dim rsCajaChica As New ADODB.Recordset 'El registro de la caja chica
Dim rsConsulta As New ADODB.Recordset 'Recordset que trae la información de la factura

Dim llngPersonaGraba As Long
Dim ldblTipoCambio As Double
Dim dblTotal As Double
    
Dim lblnRecargarConceptos As Boolean 'Para saber si se debe cargar de nuevo el combo de conceptos
Dim lblnRecargarImpuestos As Boolean 'Para saber si se debe cargar de nuevo el combo de impuestos
Dim lblnRecargarProveedores As Boolean 'Para saber si se debe cargar de nuevo el combo de proveedores
Dim lblnConsulta As Boolean 'Indica cuando se está realizando una consulta
Dim ldtmFecha As Date 'Fecha actual
Dim llngNumCorte As Long 'Número de corte donde se guarda o cancela
Dim llngNumCorteValidacionImporte As Long 'Número de corte donde se guarda o cancela
Dim llngMensajeCorteValido As Long 'Número de mensaje cuando se revisa si el corte debe cerrarse
Dim llngProveedor As Long
Dim ldblPorcentajeIVA As Double 'Porcentaje de IVA del impuesto seleccionado para el honorario
Dim lConsulta As Boolean

Private Type typTarifaImpuesto
    lngId As Long
    dblPorcentaje As Double
End Type

Dim arrTarifas() As typTarifaImpuesto

Dim ldblTipoCambioOficial As Double 'Para honorarios en dólares

Dim vlblnLicenciaContaElectronica As Boolean

Dim vlintTipoXMLCXP As Integer  '1 = CFDI, 2 = CFD, 3 = CBB, 4 = Extranjero
Dim vlstrUUIDXMLCXP As String
Dim vlstrRFCXMLCXP As String
Dim vldblMontoXMLCXP As Double
Dim vlstrMonedaXMLCXP As String
Dim vldblTipoCambioXMLCXP As Double
Dim vlstrSerieCXP As String
Dim vlstrNumFolioCXP As String
Dim vlstrNumFactExtCXP As String
Dim vlstrTaxIDExtCXP As String
Dim vlstrXMLCXP As String
Dim lstrEstadoFactura As String

Dim rsCpCajaChicaXML As ADODB.Recordset
Dim llngCuentaRetencionFletes As Long   'cuenta para retencion de fletes
Dim ldblPorcentajeRetencionFletes As Double    'Porcentaje de retencion de fletes
Dim llngCuentaFlete As Long             'cuenta para flete
Dim vlintTipoProv As Long

'Registra el movimiento de cancelación en el libro de bancos, si es que se pagó con transferencia -'
Private Sub pCancelaMovimiento(vlintNumPago As Long, vlstrFolio As String, vlStrReferencia As String, vlintCorteMovimiento As Long, vllngCorteActual As Long)
On Error GoTo NotificaError

    Dim rsMovimiento As ADODB.Recordset
    Dim lstrTipoDoc As String, lstrFecha As String
    Dim ldblCantidad As Double
    Dim rs As ADODB.Recordset
    Dim vlstrSentencia As String

'    If vlstrReferencia = "PA" Then  'Pago automático
'            vlstrSentencia = "select distinct  pvcortepoliza.intnumcorte, pvcortepoliza.chrtipodocumento, pvfactura.intconsecutivo " & _
'                     "from pvpagocortepoliza, pvcortepoliza, pvfactura " & _
'                     "where trim(pvpagocortepoliza.chrfoliorecibo) = trim('" & vlstrFolio & "') " & _
'                     "and pvpagocortepoliza.intconsecutivo = pvcortepoliza.intconsecutivo " & _
'                     "and trim(pvfactura.chrfoliofactura) = trim(pvcortepoliza.chrfoliodocumento)"
'        Set rs = frsRegresaRs(vlstrSentencia)
'        If rs.RecordCount > 0 Then
'            vlintCorteMovimiento = rs!intnumcorte
'            vlstrReferencia = rs!chrTipoDocumento
'            vlintNumPago = rs!INTCONSECUTIVO
'        End If
'    End If

    vlstrSentencia = "SELECT MB.intFormaPago, MB.mnyCantidad, MB.mnyTipoCambio, FP.chrTipo, ISNULL(B.tnyNumeroBanco, MB.intCveBanco) AS IdBanco, mb.chrtipomovimiento " & _
                     " FROM PvMovimientoBancoForma MB " & _
                     " INNER JOIN PvFormaPago FP ON MB.intFormaPago = FP.intFormaPago " & _
                     " LEFT  JOIN CpBanco B ON B.intNumeroCuenta = FP.intCuentaContable " & _
                     " WHERE TRIM(MB.chrTipoDocumento) = '" & Trim(vlStrReferencia) & "' AND MB.intNumDocumento = " & vlintNumPago & _
                     " AND MB.intNumCorte = " & vlintCorteMovimiento
    Set rsMovimiento = frsRegresaRs(vlstrSentencia)
    If rsMovimiento.RecordCount > 0 Then
        lstrFecha = fstrFechaSQL(fdtmServerFecha, fdtmServerHora) '- Fecha y hora del movimiento -'

        rsMovimiento.MoveFirst
        Do While Not rsMovimiento.EOF
            If rsMovimiento!chrTipo <> "C" Then
                '- Revisar tipo de forma de pago para determinar movimiento de cancelación -'
                Select Case rsMovimiento!chrTipo
                    Case "E": lstrTipoDoc = "CED" 'Efectivo
                    Case "T": lstrTipoDoc = "CTD" 'Tarjeta de crédito
                    Case "B": lstrTipoDoc = "CRD" 'Transferencia bancaria
                    Case "H": lstrTipoDoc = "CCD" 'Cheque
                End Select

                '- Cantidad negativa para que se tome como abono si se cancela una entrada de dinero, cantidad positiva si se cancela salida de dinero -'
                ldblCantidad = rsMovimiento!MNYCantidad * -1

                '- Guardar información en tabla intermedia -'
                vgstrParametrosSP = vllngCorteActual & "|" & lstrFecha & "|" & rsMovimiento!intFormaPago & "|" & rsMovimiento!IdBanco & "|" & ldblCantidad & "|" & _
                                    IIf(rsMovimiento!mnytipocambio = 0, 1, 0) & "|" & rsMovimiento!mnytipocambio & "|" & lstrTipoDoc & "|" & vlStrReferencia & "|" & _
                                    vlintNumPago & "|" & llngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & lstrFecha & "|" & "1" & "|" & cgstrModulo
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
            End If
            rsMovimiento.MoveNext
        Loop
    End If
    rsMovimiento.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaMovimiento"))
End Sub


'- CASO 7442: Regresa tipo de movimiento según la forma de pago -'
Private Function fstrTipoMovimientoForma(lintCveForma As Integer) As String
On Error GoTo NotificaError

    Dim rsForma As New ADODB.Recordset
    Dim lstrSentencia As String
    
    fstrTipoMovimientoForma = ""
    
    lstrSentencia = "SELECT * FROM PvFormaPago WHERE intFormaPago = " & lintCveForma
    Set rsForma = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsForma.RecordCount > 0 Then
        Select Case rsForma!chrTipo
            Case "E": fstrTipoMovimientoForma = "EDC"
            Case "T": fstrTipoMovimientoForma = "TAD"
            Case "B": fstrTipoMovimientoForma = "TRD"
            Case "H": fstrTipoMovimientoForma = "CDC"
        End Select
    End If
    rsForma.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrTipoMovimientoForma"))
End Function

Private Sub cboConcepto_GotFocus()
    On Error GoTo NotificaError
    pHabilita 0, 0, 0, 0, 0, 1, 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboConcepto_GotFocus"))
End Sub

Private Sub cboImpuesto_Click()
    On Error GoTo NotificaError
    Dim dblImpuesto As Double
    Dim dblImpuestoFlete As Double
    
      If txtFleteFactura.Text <> "" Then
      dblImpuestoFlete = Val(Format(txtFleteFactura.Text, cstrFormato))
      Else
      dblImpuestoFlete = 0
      End If
      
    If Not lblnConsulta Then
        dblImpuesto = 0
        If cboImpuesto.ListIndex <> -1 Then
            dblImpuesto = arrImpuestosFactura(cboImpuesto.ListIndex).dblPorcentaje / 100
        End If
        lblImporteIVA.Caption = FormatCurrency(((Val(Format(txtImporteGravado.Text, cstrFormato)) - Val(Format(txtDescuentoGravado.Text, cstrFormato)) _
                                + IIf(chkIEPSBaseGravable.Value, Val(Format(txtIEPS.Text, cstrFormato)) + dblImpuestoFlete, 0)) * dblImpuesto), 2)
        pCalculaTotal
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboImpuesto_Click"))
End Sub

Private Sub cboImpuesto_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboImpuesto_GotFocus"))
End Sub

Private Sub cboImpuestoFlete_Click()
 On Error GoTo NotificaError

    If cboImpuestoFlete.ListIndex <> -1 And Not lblnConsulta Then
        If cboImpuestoFlete.ListIndex = 0 Then
            lblImporteIvaFlete.Caption = FormatCurrency("0", 2)
        Else
            lblImporteIvaFlete.Caption = FormatCurrency((Val(Format(txtImporteFlete.Text, cstrFormato))) * (arrImpuestosFactura(cboImpuestoFlete.ListIndex - 1).dblPorcentaje / 100), 2)
        End If
    End If
    pCalculaTotalFlete
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboImpuesto_Click"))
End Sub

Private Sub cboImpuestoFlete_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboImpuestoFleteFac_Click()
    On Error GoTo NotificaError
    Dim dblImpuesto As Double
    
    
    If Not lblnConsulta Then
        dblImpuesto = 0
        If cboImpuestoFleteFac.ListIndex <> -1 Then
            dblImpuesto = arrImpuestosFactura(cboImpuestoFleteFac.ListIndex).dblPorcentaje / 100
        End If
        lblImpuestoFlete.Caption = FormatCurrency(Val(Format(txtFleteFactura.Text, cstrFormato)) * dblImpuesto, 2)
        pCalculaTotal
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboImpuestoFleteFac_Click"))
End Sub

Private Sub cboIVAHonorario_Click()
    On Error GoTo NotificaError
    
    If lblnConsulta = False Then
        pCalculaSubtotalHonorario
        pCalculaTotalHonorario
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboIVAHonorario_Click"))
End Sub

Private Sub cboMedicos_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cboMedicos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    If KeyAscii <> 8 Then
    If Len(cboMedicos.Text) > 69 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboMedicos_KeyPress"))
End Sub

Private Sub cboMedicos_LostFocus()
    If cboMedicos.ListIndex <> -1 Then
        Set rs = frsRegresaRs("SELECT VCHRFC FROM COPROVEEDOR WHERE intcveproveedor = " & cboMedicos.ItemData(cboMedicos.ListIndex), adLockOptimistic, adOpenForwardOnly)
        If rs.RecordCount > 0 Then
            txtRFcHono.Text = Trim(Replace(Replace(Replace(rs!vchRFC, "-", ""), "_", ""), " ", ""))
            txtRFcHono.Enabled = False
        End If
    ElseIf cboMedicos.Text <> "" Then
        txtRFcHono.Enabled = True
        txtRFcHono.Text = ""
        If fblnCanFocus(txtRFcHono) Then txtRFcHono.SetFocus
    End If
End Sub

Private Sub cboProveedor_GotFocus()
    On Error GoTo NotificaError
    pHabilita 0, 0, 0, 0, 0, 1, 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboProveedor_GotFocus"))
End Sub

Private Sub cboProveedor_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
Dim rs As ADODB.Recordset
   If KeyAscii <> 8 Then
    If Len(cboProveedor.Text) > 69 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboProveedor_KeyPress"))
End Sub

Private Sub cboProveedor_LostFocus()
 Dim rsTarifaisr As New ADODB.Recordset
    vlintTipoProv = 1
    If cboProveedor.ListIndex <> -1 Then
       frsEjecuta_SP CStr(cboProveedor.ItemData(cboProveedor.ListIndex)), "FN_PROVEEDORISMORAL", True, vlintTipoProv   ' JASM DESCOMENTAR PARA VER SI ES MORAL
    End If
    cboTipoProveedor.Enabled = True
    cboPais.Enabled = True
    
    If cboProveedor.ListIndex <> -1 Then
        Set rs = frsRegresaRs("SELECT VCHRFC, VCHTIPOPROVEEDOR, INTCVEPAIS, NVL(VCHCLAVEREGIMENSAT,0) VCHCLAVEREGIMENSAT,VCHTIPOREGIMEN FROM COPROVEEDOR WHERE intcveproveedor = " & cboProveedor.ItemData(cboProveedor.ListIndex), adLockOptimistic, adOpenForwardOnly)
        If rs.RecordCount > 0 Then 'FISICA
            txtRFC.Text = Trim(Replace(Replace(Replace(rs!vchRFC, "-", ""), "_", ""), " ", ""))
            cboPais.ListIndex = fintLocalizaCbo(cboPais, rs!intCvePais)
            cboTipoProveedor.ListIndex = fintLocalizaCritCbo(cboTipoProveedor, rs!VCHTIPOPROVEEDOR)
            txtRFC.Enabled = False
            cboPais.Enabled = False
            chkRetencionISR.Enabled = False
            cboRetencionISR.Enabled = False
            If rs!VCHCLAVEREGIMENSAT <> 0 And rs!VCHTIPOREGIMEN = "FISICA" Then
                Set rsTarifaisr = frsRegresaRs("SELECT CNT.* FROM CNREGIMENRETENCION CNR INNER JOIN CNTARIFAISR CNT ON CNR.INTIDTARIFA = CNT.INTIDTARIFA WHERE CNR.CHRIDREGIMEN = " & rs!VCHCLAVEREGIMENSAT, adLockOptimistic, adOpenForwardOnly)
                If rsTarifaisr.RecordCount > 0 Then
                    cboRetencionISR.ListIndex = fintLocalizaCbo(cboRetencionISR, rsTarifaisr!intidtarifa)
                    chkRetencionISR.Value = 1
                    'pCalculaTotalRetencionISR
                    pCalculaTotal
                Else
                    chkRetencionISR.Value = 0
                    cboRetencionISR.Enabled = False
                    cboRetencionISR.ListIndex = -1
                    lblRetencionISR.Caption = FormatCurrency("0", 2)
                    pCalculaTotal
                End If
            Else
                chkRetencionISR.Value = 0
                cboRetencionISR.Enabled = False
                cboRetencionISR.ListIndex = -1
                lblRetencionISR.Caption = FormatCurrency("0", 2)
                pCalculaTotal
            End If
        End If
    ElseIf cboProveedor.Text <> "" Then
        txtRFC.Enabled = True
        txtRFC.Text = ""
        chkRetencionISR.Enabled = True
        If FormatCurrency(Val(Format(lblRetencionISR.Caption, cstrFormato)), 2) = 0 Then
            cboRetencionISR.ListIndex = -1
            chkRetencionISR.Value = 0
        Else
            If chkRetencionISR.Value Then
                cboRetencionISR.Enabled = True
            Else
                cboRetencionISR.Enabled = False
            End If
            cboRetencionISR.ListIndex = 0
        End If
        If fblnCanFocus(txtRFC) Then txtRFC.SetFocus
    End If
'    If cboProveedor.ListIndex = -1 Then
'        fraSelXMLCajaChicaFact.Enabled = False
'    End If
End Sub

Private Sub cboRetencionISR_Click()
Dim dblretencion As Double
    If chkRetencionISR.Value And cboProveedor.ListIndex <> 0 Then
            'pCalculaTotalRetencionISR
            pCalculaTotal
    End If
End Sub

Private Sub cboTarifa_Click()
    On Error GoTo NotificaError
    Dim dblMonto As Double
    
    If lblnConsulta = False Then
        If cboTarifa.ListIndex <> -1 Then
            dblMonto = Val(Format(txtMontoHonorario.Text, cstrFormato))
            lblRetencionISRHonorario.Caption = FormatCurrency(dblMonto * arrTarifas(cboTarifa.ListIndex).dblPorcentaje / 100, 2)
            pCalculaTotalHonorario
        Else
            lblRetencionISRHonorario.Caption = FormatCurrency(0, 2)
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTarifa_Click"))
End Sub

Private Sub chkIEPSBaseGravable_Click()
    On Error GoTo NotificaError
    
        If Val(Format(txtIEPS.Text, cstrFormato)) = 0 Or (Val(Format(txtImporteGravado.Text, cstrFormato)) - Val(Format(txtDescuentoGravado.Text, cstrFormato))) <= 0 Then
            chkIEPSBaseGravable.Value = 0
            Exit Sub
        End If
        
        If Not lblTituloImpuesto.Enabled Then
            lblImporteIVA.Caption = FormatCurrency("0", 2)
            cboImpuesto.ListIndex = -1
        Else
            If cboImpuesto.ListIndex <> -1 Then
                cboImpuesto_Click
            End If
        End If
        
        pCalculaTotal
        
        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIEPS_Change"))
End Sub

Private Sub chkRetencionISR_Click()

 Dim rsTarifaisr As New ADODB.Recordset
  
    If Not cboProveedor.ListIndex <> -1 Then
        If Not Val(Format(txtImporteGravado.Text, cstrFormato)) <> 0 And Not Val(Format(txtImporteExento.Text, cstrFormato)) <> 0 And Not Val(Format(txtImporteNoGravado.Text, cstrFormato)) <> 0 Then
            chkRetencionISR.Value = 0
            cboRetencionISR.Enabled = False
            cboRetencionISR.ListIndex = -1
            lblRetencionISR.Caption = FormatCurrency("0", 2)
            Exit Sub
        End If
        
'        If chkRetencionISR Then
'        'Set rsTarifaisr = frsRegresaRs("SELECT CNT.* FROM CNREGIMENRETENCION CNR INNER JOIN CNTARIFAISR CNT ON CNR.INTIDTARIFA = CNT.INTIDTARIFA WHERE CNR.CHRIDREGIMEN = 626", adLockOptimistic, adOpenForwardOnly)
'               ' If rsTarifaisr.RecordCount > 0 Then
'                    'cboRetencionISR.ListIndex = fintLocalizaCbo(cboRetencionISR, rsTarifaisr!intidtarifa)
'                    cboRetencionISR.ListIndex = 0
'                    'pCalculaTotalRetencionISR
'                    pCalculaTotal
''                Else
''                    chkRetencionISR.Value = 0
''                    cboRetencionISR.Enabled = False
''                    cboRetencionISR.ListIndex = -1
''                    lblRetencionISR.Caption = FormatCurrency("0", 2)
''                    pCalculaTotal
''                End If
'        End If
        
        If cboRetencionISR.Enabled Then
            cboRetencionISR.Enabled = False
            cboRetencionISR.ListIndex = -1
            lblRetencionISR.Caption = FormatCurrency("0", 2)
            pCalculaTotal
        Else
            If Not cboProveedor.ListIndex <> -1 Then
                cboRetencionISR.Enabled = True
                cboRetencionISR.ListIndex = 0
                Call pCalculaTotal
            End If
        End If
        
    End If
End Sub

Private Sub chkTodas_Click()
    If chkTodas.Value = vbChecked Then
        chkActivas.Value = vbUnchecked
        chkActivas.Enabled = False
        chkCanceladas.Value = vbUnchecked
        chkCanceladas.Enabled = False
        chkDepositadas.Value = vbUnchecked
        chkDepositadas.Enabled = False
        chkPendientes.Value = vbUnchecked
        chkPendientes.Enabled = False
        chkReembolsadas.Value = vbUnchecked
        chkReembolsadas.Enabled = False
        chkSinDepositar.Value = vbUnchecked
        chkSinDepositar.Enabled = False
    Else
        chkActivas.Enabled = True
        chkCanceladas.Enabled = True
        chkDepositadas.Enabled = True
        chkPendientes.Enabled = True
        chkReembolsadas.Enabled = True
        chkSinDepositar.Enabled = True
    End If
End Sub

Private Sub cmdBack_Click()
    On Error GoTo NotificaError

    If Not rsConsulta.BOF Then
        rsConsulta.MovePrevious
        If rsConsulta.BOF Then
            rsConsulta.MoveNext
        End If
    End If
    pMuestra
    If Trim(rsConsulta!Estado) = "P" Then
        pHabilita 1, 1, 1, 1, 1, 1, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, IIf(Trim(rsConsulta!Estado) = "A", 1, 0)
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBack_Click"))
End Sub

Private Sub cmdBuscarXMLFactura_Click()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vlstrRFC As String
    Dim fsXML As Scripting.FileSystemObject
    Dim vlstrSentencia As String
    
    If vlblnLicenciaContaElectronica Then
        If optTipo(0).Value Then
            If cboProveedor.Text <> "" Then
                ' Que se haya introducido una fecha válida
                If Not IsDate(mskFecha.Text) Then
                    '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                    MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
                    mskFecha.SetFocus
                    Exit Sub
                End If
            
                ' Que la fecha sea menor o igual a la actual
                If CDate(mskFecha.Text) > ldtmFecha Then
                    '¡La fecha debe ser menor o igual a la del sistema!
                    MsgBox SIHOMsg(40), vbExclamation + vbOKOnly, "Mensaje"
                    mskFecha.SetFocus
                    Exit Sub
                End If
                
                If cboProveedor.ListIndex = -1 Then
                    If Trim(txtRFC.Text) = "" Then
                        MsgBox "Favor de ingresar el RFC del proveedor o acreedor.", vbExclamation + vbOKOnly, "Mensaje"
                        If fblnCanFocus(txtRFC) Then txtRFC.SetFocus
                        Exit Sub
                    End If
                End If
            
                If Trim(txtFolio.Text) <> "" Then
                    If (optTipoComproCajaChicaFact(1).Value Or optTipoComproCajaChicaFact(2).Value) And _
                         ((optFactura.Value And CDbl(IIf(Trim(lblTotal.Caption) = "", "0", Trim(lblTotal.Caption))) = 0) Or _
                           optFlete.Value And CDbl(IIf(Trim(lblTotalFlete.Caption) = "", "0", Trim(lblTotalFlete.Caption))) = 0) Then
                        '¡No ha ingresado datos!
                        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
                        If optFactura.Value Then
                            If txtImporteExento.Enabled Then
                                txtImporteExento.SetFocus
                            Else
                                If txtImporteNoGravado.Enabled Then
                                    txtImporteNoGravado.SetFocus
                                Else
                                    If txtImporteGravado.Enabled Then
                                        txtImporteGravado.SetFocus
                                    End If
                                End If
                            End If
                        Else
                            If fblnCanFocus(txtImporteFlete) Then txtImporteFlete.SetFocus
                        End If
                        Exit Sub
                    End If
                                        
                    If cboProveedor.ListIndex = -1 Then
                        vlstrRFC = txtRFC.Text
                    Else
                        vlstrRFC = ""
                        vlstrSentencia = "SELECT vchrfc FROM COPROVEEDOR WHERE intcveproveedor = " & cboProveedor.ItemData(cboProveedor.ListIndex)
                        Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)
                        If rs.RecordCount > 0 Then
                            vlstrRFC = rs!vchRFC
                        End If
                    End If
                    
                    If optTipoComproCajaChicaFact(0).Value Then
                        DoEvents
                        
                        vgintTipoXMLCXP = vlintTipoXMLCXP
                        vgstrUUIDXMLCXP = vlstrUUIDXMLCXP
                        vgstrRFCXMLCXP = vlstrRFCXMLCXP
                        vgdblMontoXMLCXP = vldblMontoXMLCXP
                        vgstrMonedaXMLCXP = vlstrMonedaXMLCXP
                        vgdblTipoCambioXMLCXP = vldblTipoCambioXMLCXP
                        vgstrSerieCXP = vlstrSerieCXP
                        vgstrNumFolioCXP = vlstrNumFolioCXP
                        vgstrNumFactExtCXP = vlstrNumFactExtCXP
                        vgstrTaxIDExtCXP = vlstrTaxIDExtCXP
                        vgstrXMLCXP = vlstrXMLCXP
                        
                        cmdBuscarXML.FileName = ""
                        cmdBuscarXML.Filter = "XML(*.xml)|*.xml"
                        cmdBuscarXML.Flags = cdlOFNHideReadOnly
                        cmdBuscarXML.Flags = cdlOFNFileMustExist
                        cmdBuscarXML.Flags = cdlOFNPathMustExist
                        cmdBuscarXML.ShowOpen
                        If cmdBuscarXML.FileName <> "" Then
                            Set fsXML = New Scripting.FileSystemObject
                            If fsXML.FileExists(cmdBuscarXML.FileName) = True Then
                                If fstrObtieneInformacionXML(cmdBuscarXML.FileName, vlstrRFC, txtFolio.Text, CDate(mskFecha.Text)) Then
                                    vlstrUUIDXMLCXP = vgstrUUIDXMLCXP
                                    vlstrRFCXMLCXP = vgstrRFCXMLCXP
                                    vldblMontoXMLCXP = vgdblMontoXMLCXP
                                    vlstrMonedaXMLCXP = vgstrMonedaXMLCXP
                                    vldblTipoCambioXMLCXP = vgdblTipoCambioXMLCXP
                                    vlstrXMLCXP = vgstrXMLCXP
                                    
                                    vlintTipoXMLCXP = vgintTipoXMLCXP
                                    vlstrSerieCXP = IIf(vgintTipoXMLCXP = 1, "", vgstrSerieCXP)
                                    vlstrNumFolioCXP = IIf(vgintTipoXMLCXP = 1, "", vgstrNumFolioCXP)
                                    vlstrNumFactExtCXP = ""
                                    vlstrTaxIDExtCXP = ""
                                    
                                    chkXMLrelacionadoFact.Value = 2
                                    If lblnConsulta Then
                                        pHabilita 0, 0, 0, 0, 0, 1, 0
                                    End If
                                End If
                            Else
                                MsgBox SIHOMsg(1083), vbOKOnly + vbExclamation, "Mensaje" 'Archivo no encontrado.
                            End If
                        End If
                    Else
                        If optTipoComproCajaChicaFact(1).Value Or optTipoComproCajaChicaFact(2).Value Then
                                    
                            vgintTipoXMLCXP = vlintTipoXMLCXP
                            vgstrUUIDXMLCXP = vlstrUUIDXMLCXP
                            vgstrRFCXMLCXP = vlstrRFCXMLCXP
                            vgdblMontoXMLCXP = vldblMontoXMLCXP
                            vgstrMonedaXMLCXP = vlstrMonedaXMLCXP
                            vgdblTipoCambioXMLCXP = vldblTipoCambioXMLCXP
                            vgstrSerieCXP = vlstrSerieCXP
                            vgstrNumFolioCXP = vlstrNumFolioCXP
                            vgstrNumFactExtCXP = vlstrNumFactExtCXP
                            vgstrTaxIDExtCXP = vlstrTaxIDExtCXP
                            vgstrXMLCXP = vlstrXMLCXP
        
                            frmInformacionOtrosComprobantes.vldblTipoCambio = 0
                            If optFactura.Value Then
                                frmInformacionOtrosComprobantes.vlintMontoTotal = CDbl(IIf(Trim(lblTotal.Caption) = "", "0", Trim(lblTotal.Caption)))
                            Else
                                frmInformacionOtrosComprobantes.vlintMontoTotal = CDbl(IIf(Trim(lblTotalFlete.Caption) = "", "0", Trim(lblTotalFlete.Caption)))
                            End If
                            frmInformacionOtrosComprobantes.vlblnPesos = IIf(optMoneda(0).Value, True, False)
                            frmInformacionOtrosComprobantes.vlintTipoXML = IIf(optTipoComproCajaChicaFact(1).Value, 3, 4)
                            frmInformacionOtrosComprobantes.Show vbModal
                            
                            If vgintTipoXMLCXP = 3 Or vgintTipoXMLCXP = 4 Then
                                vlstrRFCXMLCXP = vgstrRFCXMLCXP
                                vlstrUUIDXMLCXP = ""
                                vldblMontoXMLCXP = vgdblMontoXMLCXP
                                vlstrMonedaXMLCXP = vgstrMonedaXMLCXP
                                vldblTipoCambioXMLCXP = vgdblTipoCambioXMLCXP
                                vlstrXMLCXP = ""
                                
                                vlintTipoXMLCXP = vgintTipoXMLCXP
                                vlstrSerieCXP = IIf(vgintTipoXMLCXP = 3, vgstrSerieCXP, "")
                                vlstrNumFolioCXP = IIf(vgintTipoXMLCXP = 3, vgstrNumFolioCXP, "")
                                vlstrNumFactExtCXP = IIf(vgintTipoXMLCXP = 4, vgstrNumFactExtCXP, "")
                                vlstrTaxIDExtCXP = IIf(vgintTipoXMLCXP = 4, vgstrTaxIDExtCXP, "")
                                
                                chkXMLrelacionadoFact.Value = 2
                                If lblnConsulta Then
                                    pHabilita 0, 0, 0, 0, 0, 1, 0
                                End If
                            End If
                        End If
                    End If
                Else
                    '¡No ha ingresado datos!
                    MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
                    txtFolio.SetFocus
                End If
            Else
                'Seleccione el proveedor
                MsgBox SIHOMsg(206), vbOKOnly + vbInformation, "Mensaje"
                If fblnCanFocus(cboProveedor) Then cboProveedor.SetFocus
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscarXMLFactura_Click"))
End Sub

Private Sub cmdBuscarXMLHonorario_Click()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vlstrRFC As String
    Dim fsXML As Scripting.FileSystemObject
    Dim vlstrSentencia As String
    
    If vlblnLicenciaContaElectronica Then
        If optTipo(1).Value Then
            If cboMedicos.Text <> "" Then
                ' Que se haya introducido una fecha válida
                If Not IsDate(mskFechaHonorario.Text) Then
                    '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                    MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
                    mskFechaHonorario.SetFocus
                    Exit Sub
                End If
            
                ' Que la fecha sea menor o igual a la actual
                If CDate(mskFechaHonorario.Text) > ldtmFecha Then
                    '¡La fecha debe ser menor o igual a la del sistema!
                    MsgBox SIHOMsg(40), vbExclamation + vbOKOnly, "Mensaje"
                    mskFechaHonorario.SetFocus
                    Exit Sub
                End If
            
                If cboMedicos.ListIndex = -1 Then
                    If Trim(txtRFcHono.Text) = "" Then
                        MsgBox "Favor de ingresar el RFC del médico.", vbExclamation + vbOKOnly, "Mensaje"
                        If fblnCanFocus(txtRFcHono) Then txtRFcHono.SetFocus
                        Exit Sub
                    End If
                End If
                
                If Trim(txtFolioHonorario.Text) <> "" Then
                    If (optTipoComproCajaChicaHono(1).Value Or optTipoComproCajaChicaHono(2).Value) And CDbl(IIf(Trim(lblTotalPagarHonorario.Caption) = "", "0", Trim(lblTotalPagarHonorario.Caption))) = 0 Then
                        '¡No ha ingresado datos!
                        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
                        If txtMontoHonorario.Enabled Then
                            txtMontoHonorario.SetFocus
                        End If
                        Exit Sub
                    End If
                                        
                    If cboMedicos.ListIndex = -1 Then
                        vlstrRFC = txtRFcHono.Text
                    Else
                        vlstrRFC = ""
                        vlstrSentencia = "SELECT vchrfc FROM COPROVEEDOR WHERE intcveproveedor = " & cboMedicos.ItemData(cboMedicos.ListIndex)
                        Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)
                        If rs.RecordCount > 0 Then
                            vlstrRFC = rs!vchRFC
                        End If
                    End If
                
                    If optTipoComproCajaChicaHono(0).Value Then
                        DoEvents
                        
                        vgintTipoXMLCXP = vlintTipoXMLCXP
                        vgstrUUIDXMLCXP = vlstrUUIDXMLCXP
                        vgstrRFCXMLCXP = vlstrRFCXMLCXP
                        vgdblMontoXMLCXP = vldblMontoXMLCXP
                        vgstrMonedaXMLCXP = vlstrMonedaXMLCXP
                        vgdblTipoCambioXMLCXP = vldblTipoCambioXMLCXP
                        vgstrSerieCXP = vlstrSerieCXP
                        vgstrNumFolioCXP = vlstrNumFolioCXP
                        vgstrNumFactExtCXP = vlstrNumFactExtCXP
                        vgstrTaxIDExtCXP = vlstrTaxIDExtCXP
                        vgstrXMLCXP = vlstrXMLCXP
                        
                        cmdBuscarXML.FileName = ""
                        cmdBuscarXML.Filter = "XML(*.xml)|*.xml"
                        cmdBuscarXML.Flags = cdlOFNHideReadOnly
                        cmdBuscarXML.Flags = cdlOFNFileMustExist
                        cmdBuscarXML.Flags = cdlOFNPathMustExist
                        cmdBuscarXML.ShowOpen
                        If cmdBuscarXML.FileName <> "" Then
                            Set fsXML = New Scripting.FileSystemObject
                            If fsXML.FileExists(cmdBuscarXML.FileName) = True Then
                                If fstrObtieneInformacionXML(cmdBuscarXML.FileName, vlstrRFC, txtFolioHonorario.Text, CDate(mskFechaHonorario.Text)) Then
                                    vlstrUUIDXMLCXP = vgstrUUIDXMLCXP
                                    vlstrRFCXMLCXP = vgstrRFCXMLCXP
                                    vldblMontoXMLCXP = vgdblMontoXMLCXP
                                    vlstrMonedaXMLCXP = vgstrMonedaXMLCXP
                                    vldblTipoCambioXMLCXP = vgdblTipoCambioXMLCXP
                                    vlstrXMLCXP = vgstrXMLCXP
                                    
                                    vlintTipoXMLCXP = vgintTipoXMLCXP
                                    vlstrSerieCXP = IIf(vgintTipoXMLCXP = 1, "", vgstrSerieCXP)
                                    vlstrNumFolioCXP = IIf(vgintTipoXMLCXP = 1, "", vgstrNumFolioCXP)
                                    vlstrNumFactExtCXP = ""
                                    vlstrTaxIDExtCXP = ""
                                    
                                    chkXMLrelacionadoHono.Value = 2
                                    If lblnConsulta Then
                                        pHabilita 0, 0, 0, 0, 0, 1, 0
                                    End If
                                End If
                            Else
                                MsgBox SIHOMsg(1083), vbOKOnly + vbExclamation, "Mensaje" 'Archivo no encontrado.
                            End If
                        End If
                    Else
                        If optTipoComproCajaChicaHono(1).Value Or optTipoComproCajaChicaHono(2).Value Then
                                    
                            vgintTipoXMLCXP = vlintTipoXMLCXP
                            vgstrUUIDXMLCXP = vlstrUUIDXMLCXP
                            vgstrRFCXMLCXP = vlstrRFCXMLCXP
                            vgdblMontoXMLCXP = vldblMontoXMLCXP
                            vgstrMonedaXMLCXP = vlstrMonedaXMLCXP
                            vgdblTipoCambioXMLCXP = vldblTipoCambioXMLCXP
                            vgstrSerieCXP = vlstrSerieCXP
                            vgstrNumFolioCXP = vlstrNumFolioCXP
                            vgstrNumFactExtCXP = vlstrNumFactExtCXP
                            vgstrTaxIDExtCXP = vlstrTaxIDExtCXP
                            vgstrXMLCXP = vlstrXMLCXP
        
                            frmInformacionOtrosComprobantes.vldblTipoCambio = 0
                            frmInformacionOtrosComprobantes.vlintMontoTotal = CDbl(IIf(Trim(lblTotalPagarHonorario.Caption) = "", "0", Trim(lblTotalPagarHonorario.Caption)))
                            frmInformacionOtrosComprobantes.vlblnPesos = IIf(optMonedaHonorario(0).Value, True, False)
                            frmInformacionOtrosComprobantes.vlintTipoXML = IIf(optTipoComproCajaChicaHono(1).Value, 3, 4)
                            frmInformacionOtrosComprobantes.Show vbModal
                            
                            If vgintTipoXMLCXP = 3 Or vgintTipoXMLCXP = 4 Then
                                vlstrRFCXMLCXP = vgstrRFCXMLCXP
                                vlstrUUIDXMLCXP = ""
                                vldblMontoXMLCXP = vgdblMontoXMLCXP
                                vlstrMonedaXMLCXP = vgstrMonedaXMLCXP
                                vldblTipoCambioXMLCXP = vgdblTipoCambioXMLCXP
                                vlstrXMLCXP = ""
                                
                                vlintTipoXMLCXP = vgintTipoXMLCXP
                                vlstrSerieCXP = IIf(vgintTipoXMLCXP = 3, vgstrSerieCXP, "")
                                vlstrNumFolioCXP = IIf(vgintTipoXMLCXP = 3, vgstrNumFolioCXP, "")
                                vlstrNumFactExtCXP = IIf(vgintTipoXMLCXP = 4, vgstrNumFactExtCXP, "")
                                vlstrTaxIDExtCXP = IIf(vgintTipoXMLCXP = 4, vgstrTaxIDExtCXP, "")
                                
                                chkXMLrelacionadoHono.Value = 2
                                If lblnConsulta Then
                                    pHabilita 0, 0, 0, 0, 0, 1, 0
                                End If
                            End If
                          
                        End If
                    End If
                      If cmdSave.Enabled = True Then cmdSave.SetFocus
                Else
                    '¡No ha ingresado datos!
                    MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
                    txtFolioHonorario.SetFocus
                End If
            Else
                'Seleccione el médico
                MsgBox SIHOMsg(332), vbOKOnly + vbInformation, "Mensaje"
                cboMedicos.SetFocus
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscarXMLHonorario_Click"))
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo NotificaError

    If fblnRevisaPermiso(vglngNumeroLogin, 301, "E") Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If llngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            If optTipo(0).Value = True Or optTipo(1).Value = True Then
                If fblnCancelacionValida() Then 'Esta función cancela la transacción cuando existe algún error
                    '------------------------------------------------------------------
                    ' Cancelar la factura
                    '------------------------------------------------------------------
                    vgstrParametrosSP = CStr(rsConsulta!IdFactura) & "|" & CStr(llngPersonaGraba)
                    frsEjecuta_SP vgstrParametrosSP, "SP_CPUPDCANCELAFACTURACHICA"
                    '------------------------------------------------------------------
                    ' Afectar el corte
                    '------------------------------------------------------------------
                    vgstrParametrosSP = CStr(rsConsulta!IdFactura) & "|" & "SC" & "|" & CStr(rsConsulta!IdCorte)
                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaDoctoCorte")
                    Do While Not rs.EOF
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                        & "|" & CStr(rsConsulta!IdFactura) _
                        & "|" & "SC" _
                        & "|" & CStr(rs!intFormaPago) _
                        & "|" & CStr(rs!mnyCantidadPagada * -1) _
                        & "|" & CStr(rs!mnytipocambio) _
                        & "|" & CStr(rs!intfoliocheque) _
                        & "|" & CStr(rsConsulta!IdCorte)
                        
                        frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                        rs.MoveNext
                    Loop
                    '------------------------------------------------------------------
                    ' Afectar la póliza del corte
                    '------------------------------------------------------------------
                    vgstrParametrosSP = CStr(rsConsulta!IdFactura) & "|" & "SC" & "|" & CStr(rsConsulta!IdCorte)
                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelPolizaDocto")
                    Do While Not rs.EOF
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & CStr(rsConsulta!IdFactura) _
                        & "|" & "SC" _
                        & "|" & CStr(rs!intNumCuenta) _
                        & "|" & CStr(rs!MNYCantidad * IIf(rsConsulta!IdCorte = llngNumCorte, -1, 1)) _
                        & "|" & CStr(IIf(rsConsulta!IdCorte = llngNumCorte, rs!bitcargo, IIf(rs!bitcargo = 1, 0, 1)))
                        
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        rs.MoveNext
                    Loop
                    '------------------------------------------------------------------
                    ' Liberar el corte
                    '------------------------------------------------------------------
                    pLiberaCorte llngNumCorte
                    '------------------------------------------------------------------
                    ' Registro de transacciones
                    '------------------------------------------------------------------
                    pGuardarLogTransaccion Me.Name, EnmCancelacion, llngPersonaGraba, "CAJA CHICA", CStr(rsConsulta!IdFactura)
                
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    
                    'La operación se realizó satisfactoriamente.
                    MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                    If optTipo(0).Value = True Then pLimpia
                    If optTipo(1).Value = True Then pLimpiaHonorarios
                    txtNumero.SetFocus
                End If
           ElseIf optTipo(2).Value = True Then
                If fblnCancelacionValida() Then 'Esta función cancela la transacción cuando existe algún error
                    '------------------------------------------------------------------
                    ' Cancelar la factura
                    '------------------------------------------------------------------
                    vgstrParametrosSP = CStr(rsConsulta!IdFactura) & "|" & CStr(llngPersonaGraba)
                    frsEjecuta_SP vgstrParametrosSP, "SP_CPUPDCANCELAFACTURACHICA"
                    '------------------------------------------------------------------
                    ' Afectar el corte
                    '------------------------------------------------------------------
                    vgstrParametrosSP = CStr(rsConsulta!IdFactura) & "|" & "SC" & "|" & CStr(rsConsulta!IdCorte)
                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaDoctoCorte")
                    Do While Not rs.EOF
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                        & "|" & CStr(rsConsulta!IdFactura) _
                        & "|" & "SC" _
                        & "|" & CStr(rs!intFormaPago) _
                        & "|" & CStr(rs!mnyCantidadPagada * -1) _
                        & "|" & CStr(rs!mnytipocambio) _
                        & "|" & CStr(rs!intfoliocheque) _
                        & "|" & CStr(rsConsulta!IdCorte)
                        
                        frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                        rs.MoveNext
                    Loop
                    
                    '------------------------------------------------------------------
                    ' Afectar la póliza del corte
                    '------------------------------------------------------------------
                    vgstrParametrosSP = CStr(rsConsulta!IdFactura) & "|" & "SC" & "|" & CStr(rsConsulta!IdCorte)
                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelPolizaDocto")
                    Do While Not rs.EOF
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & CStr(rsConsulta!IdFactura) _
                        & "|" & "SC" _
                        & "|" & CStr(rs!intNumCuenta) _
                        & "|" & CStr(rs!MNYCantidad * IIf(rsConsulta!IdCorte = llngNumCorte, -1, 1)) _
                        & "|" & CStr(IIf(rsConsulta!IdCorte = llngNumCorte, rs!bitcargo, IIf(rs!bitcargo = 1, 0, 1)))
                        
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        rs.MoveNext
                    Loop
                                        
                    '------------------------------------------------------------------
                    ' Cancelar el movimiento de la forma de pago
                    '------------------------------------------------------------------
                    pCancelaMovimiento rsConsulta!IdFactura, Trim(rsConsulta!IdFactura), "SC", rsConsulta!IdCorte, llngNumCorte
                    
                    '------------------------------------------------------------------
                    ' Liberar el corte
                    '------------------------------------------------------------------
                    pLiberaCorte llngNumCorte
                    '------------------------------------------------------------------
                    ' Registro de transacciones
                    '------------------------------------------------------------------
                    pGuardarLogTransaccion Me.Name, EnmCancelacion, llngPersonaGraba, "CANCELACIÓN DE DISMINUCIÓN DE FONDO EN CAJA CHICA", CStr(rsConsulta!IdFactura)
                
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    
                    'La operación se realizó satisfactoriamente.
                    MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaDF
                    txtNumero.SetFocus
                End If
           End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCancelar_Click"))
End Sub

Private Function fblnCancelacionValida() As Boolean
    On Error GoTo NotificaError
    Dim lngCorteGrabando As Long

    fblnCancelacionValida = True
    
    ' Que exista un número de corte
    llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
    If llngNumCorte = 0 Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        fblnCancelacionValida = False
        'No se encontró un corte abierto.
        MsgBox SIHOMsg(659), vbExclamation + vbOKOnly, "Mensaje"
    End If

    ' Que tenga el corte esté libre
    If fblnCancelacionValida Then
        lngCorteGrabando = 1
        frsEjecuta_SP CStr(llngNumCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, lngCorteGrabando
        If lngCorteGrabando <> 2 Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            fblnCancelacionValida = False
            'En este momento se está afectando el corte, espere un momento e intente de nuevo.
            MsgBox SIHOMsg(779), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
    
    ' Que exista la factura
    If fblnCancelacionValida Then
        pConsulta Val(txtNumero.Text), ldtmFecha, ldtmFecha, 0, -1, vgintNumeroDepartamento
        If rsConsulta.RecordCount = 0 Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            fblnCancelacionValida = False
            'La información ha cambiado, consulte de nuevo.
            MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If

    ' Que tenga estado activo
    If fblnCancelacionValida Then
        If Trim(rsConsulta!Estado) <> "A" Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            fblnCancelacionValida = False
            'La información ha cambiado, consulte de nuevo.
            MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnCancelacionValida"))
End Function

Private Sub cmdCargar_Click()
    On Error GoTo NotificaError
    Dim intcontador As Integer

    pConfiguraBusqueda
    
    If chkTodas.Value = vbUnchecked And chkActivas.Value = vbUnchecked And chkCanceladas.Value = vbUnchecked And chkDepositadas.Value = vbUnchecked And chkPendientes.Value = vbUnchecked And chkReembolsadas.Value = vbUnchecked And chkSinDepositar.Value = vbUnchecked Then
        MsgBox SIHOMsg(1618), vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    Else
    
    End If
    
    If optTipo(0).Value = True Or optTipo(1).Value = True Then
        pConsulta -1, CDate(mskFechaBusIni.Text), CDate(mskFechaBusFin.Text), 1, cboProveedorBus.ItemData(cboProveedorBus.ListIndex), vgintNumeroDepartamento
    ElseIf optTipo(2).Value = True Then
        pConsulta -1, CDate(mskFechaBusIni.Text), CDate(mskFechaBusFin.Text), 1, -1, vgintNumeroDepartamento
    End If
    With rsConsulta
        If .RecordCount <> 0 Then
            Do While Not .EOF
                grdFacturas.TextMatrix(grdFacturas.Rows - 1, cintColFecha) = Format(!FECHAREGISTRO, "dd/mmm/yyyy")
                grdFacturas.TextMatrix(grdFacturas.Rows - 1, cintColNumero) = !IdFactura
                grdFacturas.TextMatrix(grdFacturas.Rows - 1, cintColProveedor) = IIf(IsNull(!NombreProveedor), "", !NombreProveedor)
                grdFacturas.TextMatrix(grdFacturas.Rows - 1, cintColFactura) = !FolioFactura
                grdFacturas.TextMatrix(grdFacturas.Rows - 1, cIntColEstado) = !EstadoDescripcion
                grdFacturas.TextMatrix(grdFacturas.Rows - 1, cintColRegistro) = !NombreEmpleado
                If Trim(!Estado) = "C" Then
                    grdFacturas.TextMatrix(grdFacturas.Rows - 1, cintColCancelo) = !NombrePersonaCancelo
                Else
                    If Trim(!Estado) = "R" Then
                        grdFacturas.TextMatrix(grdFacturas.Rows - 1, cintColCancelo) = !NombrePersonaReembolso
                    End If
                End If
                
                grdFacturas.Row = grdFacturas.Rows - 1
                For intcontador = 1 To grdFacturas.Cols - 1
                    grdFacturas.Col = intcontador
                    grdFacturas.CellForeColor = IIf(Trim(!Estado) = "C", clngColorCanceladas, IIf(Trim(!Estado) = "P", clngColorPendientes, clngColorActivas))
                Next intcontador
                .MoveNext
                grdFacturas.Rows = grdFacturas.Rows + 1
            Loop
            grdFacturas.Rows = grdFacturas.Rows - 1
            grdFacturas.Col = cintColFecha
            grdFacturas.Row = 1
            grdFacturas.SetFocus
        Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
    
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCargar_Click"))
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError

    rsConsulta.MoveLast
    pMuestra
    If Trim(rsConsulta!Estado) = "P" Then
        pHabilita 1, 1, 1, 1, 1, 1, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, IIf(Trim(rsConsulta!Estado) = "A", 1, 0)
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError

    SSTab.Tab = 1
    If optTipo(0).Value = True Then
        lbProveedor(1).Caption = "Proveedor"
        cboProveedorBus.ToolTipText = "Proveedor"
        cboProveedorBus.Visible = True
        pCargaProveedores
        grdFacturas.Height = 6360
        FraConsulta.Height = 6520
    ElseIf optTipo(1).Value = True Then
        lbProveedor(1).Caption = "Médico"
        cboProveedorBus.ToolTipText = "Médico"
        cboProveedorBus.Visible = True
        pCargaProveedores
        grdFacturas.Height = 4700
        FraConsulta.Height = 4850
    ElseIf optTipo(2).Value = True Then
        lbProveedor(1).Visible = False
        cboProveedorBus.Visible = False
        pCargaProveedores
        grdFacturas.Height = 3450
        FraConsulta.Height = 3600
    End If
  
    cmdCargar_Click
    
    If rsConsulta.RecordCount = 0 Then
        mskFechaBusIni.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError

    If Not rsConsulta.EOF Then
        rsConsulta.MoveNext
        If rsConsulta.EOF Then
            rsConsulta.MovePrevious
        End If
    End If
    pMuestra
    If Trim(rsConsulta!Estado) = "P" Then
        pHabilita 1, 1, 1, 1, 1, 1, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, IIf(Trim(rsConsulta!Estado) = "A", 1, 0)
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError

    Dim lngCorteGrabando  As Long
    Dim dblimporteExento As Double
    Dim dblDescImporteExento As Double
    Dim dblImporteNoGravado As Double
    Dim dblDescImporteNoGravado As Double
    Dim dblimportegravado As Double
    Dim dblDescImporteGravado As Double
    Dim dblImpuesto As Double
    Dim dblTotalDF As Double
    Dim lngCveImpuesto As Long
    Dim lngidfactura As Long
    Dim intcontador As Integer
    Dim strParametros As String
    Dim intNumCuentaHonorario As Long
    Dim intNumCuentaBancaria As Long
    Dim dblMontoExento As Double
    Dim dblMonto As Double
    Dim dblSubTotal As Double
    Dim dblRetencionISR As Double
    Dim dblRetencionIVA As Double
    Dim dblProvisionadoISR As Double
    Dim rsCnParametro As ADODB.Recordset
    Dim rsCuentaCuenta As New ADODB.Recordset
    Dim lnCuentaRetencionIVA As Long
    Dim lnCuentaRetencionISR As Long
    Dim lngCveTarifaISR As Long
    Dim dblIEPS As Double
    Dim vllngNumDetalleCorte As Long
    Dim vlrelivaentrada As Double
    Dim vlintNumeroProceso As Integer
    Dim i As Integer
    Dim dblFlete As Double
    Dim dblRetencionFlete As Double
    Dim rsCorte As New ADODB.Recordset
    Dim llngCorteXml As Long
    Dim lngCveImpuestoFlete As Long
    Dim lngCveRetencionISR As Long
    
    llngNumCorteValidacionImporte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
    If optTipo(0).Value = True Then
        ' FACTURAS
        If lblnConsulta And lstrEstadoFactura = "A" And (optFactura.Value Or optFlete.Value) Then
            llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If llngPersonaGraba = 0 Then
                Exit Sub
            End If
            'si esta activa, se puede asignar el xml
            EntornoSIHO.ConeccionSIHO.BeginTrans
            Set rsCorte = frsRegresaRs("select intNumCorte from CpFacturaCajaChica where intIdFactura = " + Trim(txtNumero.Text))
            If rsCorte.RecordCount <> 0 Then
                llngCorteXml = rsCorte!intNumCorte
            End If
            lngidfactura = Val(txtNumero.Text)
                            vgstrParametrosSP = lngidfactura & "|'A'" _
                            & "|" & CStr(llngPersonaGraba) & "|" & IIf(llngCorteXml = 0, Null, CStr(llngCorteXml)) _
                            & "|" & IIf(IsNull(Trim(vlstrUUIDXMLCXP)), Null, IIf(vlintTipoXMLCXP = 1, Trim(vlstrUUIDXMLCXP), "")) _
                            & "|" & IIf(IsNull(vldblMontoXMLCXP) Or vldblMontoXMLCXP = 0, Null, vldblMontoXMLCXP) _
                            & "|" & IIf(IsNull(Trim(vlstrMonedaXMLCXP)), Null, Trim(vlstrMonedaXMLCXP)) _
                            & "|" & IIf(IsNull(vldblTipoCambioXMLCXP) Or vldblTipoCambioXMLCXP = 0, Null, vldblTipoCambioXMLCXP) _
                            & "|" & IIf(IsNull(Trim(vlintTipoXMLCXP)) Or vlintTipoXMLCXP = 0, Null, Trim(vlintTipoXMLCXP)) _
                            & "|" & IIf(IsNull(Trim(vlstrSerieCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrSerieCXP), "")) _
                            & "|" & IIf(IsNull(Trim(vlstrNumFolioCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrNumFolioCXP), "")) _
                            & "|" & IIf(IsNull(Trim(vlstrNumFactExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrNumFactExtCXP), "")) _
                            & "|" & IIf(IsNull(Trim(vlstrTaxIDExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrTaxIDExtCXP), ""))
                            frsEjecuta_SP vgstrParametrosSP, "SP_CPUPDFACTURACAJACHICA"
            If vlblnLicenciaContaElectronica And optFactura.Value Or optFlete.Value Then
                pEjecutaSentencia "DELETE FROM CPFACTURACAJACHICAXML WHERE INTIDFACTURA = " & lngidfactura
                If Trim(vlstrXMLCXP) <> "" And (vlintTipoXMLCXP = 1 Or vlintTipoXMLCXP = 2) Then
                    With rsCpCajaChicaXML
                        .AddNew
                        !intIdFactura = lngidfactura
                        !CLBXML = vlstrXMLCXP
                        .Update
                    End With
                End If
            End If
            pGuardarLogTransaccion Me.Name, EnmCambiar, llngPersonaGraba, "CAJA CHICA - Asignar XML", CStr(lngidfactura)
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
            txtNumero.SetFocus
        Else
            If fblnDatosValidos() Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
               
                llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
                If llngNumCorte = 0 Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    
                    'No se encontró un corte abierto.
                    MsgBox SIHOMsg(659), vbExclamation + vbOKOnly, "Mensaje"
                Else
                    '------------------------------------------------------------------
                    ' Bloquear el corte
                    '------------------------------------------------------------------
                    lngCorteGrabando = 1
                    frsEjecuta_SP CStr(llngNumCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, lngCorteGrabando
                    If lngCorteGrabando <> 2 Then
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        
                        'En este momento se está afectando el corte, espere un momento e intente de nuevo.
                        MsgBox SIHOMsg(779), vbExclamation + vbOKOnly, "Mensaje"
                    Else
                        '------------------------------------------------------------------
                        ' Insertar la factura
                        '------------------------------------------------------------------
                        lngCveImpuesto = 0
                        If optFactura.Value Then
                            ' facturas
                            dblimporteExento = Val(Format(txtImporteExento.Text, cstrFormato))
                            dblDescImporteExento = Val(Format(txtDescuentoExento.Text, cstrFormato))
                            dblImporteNoGravado = Val(Format(txtImporteNoGravado.Text, cstrFormato))
                            dblDescImporteNoGravado = Val(Format(txtDescuentoNoGravado.Text, cstrFormato))
                            dblimportegravado = Val(Format(txtImporteGravado.Text, cstrFormato))
                            dblDescImporteGravado = Val(Format(txtDescuentoGravado.Text, cstrFormato))
                            dblImpuesto = Val(Format(lblImporteIVA.Caption, cstrFormato)) + Val(Format(lblImpuestoFlete.Caption, cstrFormato))
                            dblTotal = Val(Format(lblTotal.Caption, cstrFormato))
                            dblIEPS = Val(Format(txtIEPS.Text, cstrFormato))
                            dblProvisionadoISR = Val(Format(lblRetencionISR.Caption, cstrFormato))
                            lngCveRetencionISR = 0
                            If cboRetencionISR.ListIndex <> -1 Then
                                lngCveRetencionISR = cboRetencionISR.ItemData(cboRetencionISR.ListIndex)
                            End If
                            dblFlete = Val(Format(txtFleteFactura.Text, cstrFormato))
                            lngCveImpuestoFlete = 0
                            If cboImpuestoFleteFac.ListIndex <> -1 Then
                                lngCveImpuestoFlete = cboImpuestoFleteFac.ItemData(cboImpuestoFleteFac.ListIndex)
                            End If
                            dblRetencionFlete = Val(Format(lblRetencionFactura.Caption, cstrFormato))
                            
                            If dblimportegravado <> 0 Then
                                lngCveImpuesto = cboImpuesto.ItemData(cboImpuesto.ListIndex)
                            End If
                        ElseIf optFlete.Value Then
                            'flete
                            dblTotal = Val(Format(lblTotalFlete.Caption, cstrFormato))
                            dblFlete = Val(Format(txtImporteFlete.Text, cstrFormato))
                            dblImpuesto = Val(Format(lblImporteIvaFlete.Caption, cstrFormato))
                            If dblImpuesto <> 0 Then
                                lngCveImpuesto = cboImpuestoFlete.ItemData(cboImpuestoFlete.ListIndex)
                            End If
                            dblRetencionFlete = IIf(optRetencion(0).Value, Val(Format(lblRetencion.Caption, cstrFormato)), 0)
                        Else
                            'notas y tickets
                            dblTotal = Val(Format(txtTotalTicket.Text, cstrFormato))
                            dblImporteNoGravado = Val(Format(txtTotalTicket.Text, cstrFormato))
                        End If
                                                
                        If lblnConsulta And lstrEstadoFactura = "P" Then
                            lngidfactura = Val(txtNumero.Text)
                            vgstrParametrosSP = lngidfactura & "|'A'" _
                            & "|" & CStr(llngPersonaGraba) & "|" & CStr(llngNumCorte) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrUUIDXMLCXP)), Null, IIf(vlintTipoXMLCXP = 1, Trim(vlstrUUIDXMLCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(vldblMontoXMLCXP) Or vldblMontoXMLCXP = 0, Null, vldblMontoXMLCXP), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrMonedaXMLCXP)), Null, Trim(vlstrMonedaXMLCXP)), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(vldblTipoCambioXMLCXP) Or vldblTipoCambioXMLCXP = 0, Null, vldblTipoCambioXMLCXP), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlintTipoXMLCXP)) Or vlintTipoXMLCXP = 0, Null, Trim(vlintTipoXMLCXP)), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrSerieCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrSerieCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrNumFolioCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrNumFolioCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrNumFactExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrNumFactExtCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrTaxIDExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrTaxIDExtCXP), "")), Null)
                            frsEjecuta_SP vgstrParametrosSP, "SP_CPUPDFACTURACAJACHICA"
                        Else
                            vgstrParametrosSP = fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) _
                            & "|" & CStr(llngProveedor) & "|" & Trim(txtFolio.Text) _
                            & "|" & fstrFechaSQL(mskFecha.Text) & "|" & cboConcepto.ItemData(cboConcepto.ListIndex) _
                            & "|" & CStr(dblimporteExento) & "|" & CStr(dblDescImporteExento) _
                            & "|" & CStr(dblImporteNoGravado) & "|" & CStr(dblDescImporteNoGravado) _
                            & "|" & CStr(dblimportegravado) & "|" & CStr(dblDescImporteGravado) _
                            & "|" & CStr(dblImpuesto) & "|" & CStr(vgintNumeroDepartamento) _
                            & "|" & CStr(llngPersonaGraba) & "|" & CStr(llngNumCorte) _
                            & "|" & "A" & "|" & IIf(optMoneda(0).Value, "1", "0") _
                            & "|" & CStr(lngCveImpuesto) & "|" & CStr(IIf(optMoneda(0).Value, 0, ldblTipoCambio)) _
                            & "|" & IIf(llngProveedor = 0, Trim(cboProveedor.Text), "") & "|" & CStr(IIf(optFactura.Value, "F", IIf(optFlete.Value, "L", IIf(optTicket.Value, "T", "N")))) _
                            & "|" & "" _
                            & "|" & IIf(lngCveRetencionISR = 0, "", lngCveRetencionISR) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrUUIDXMLCXP)), Null, IIf(vlintTipoXMLCXP = 1, Trim(vlstrUUIDXMLCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(vldblMontoXMLCXP) Or vldblMontoXMLCXP = 0, Null, vldblMontoXMLCXP), Null) _
                            & "|" & CStr(dblIEPS) _
                            & "|" & IIf(dblIEPS > 0, CStr(IIf(chkIEPSBaseGravable.Value, 1, 0)), "0") _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrMonedaXMLCXP)), Null, Trim(vlstrMonedaXMLCXP)), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(vldblTipoCambioXMLCXP) Or vldblTipoCambioXMLCXP = 0, Null, vldblTipoCambioXMLCXP), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlintTipoXMLCXP)) Or vlintTipoXMLCXP = 0, Null, Trim(vlintTipoXMLCXP)), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrSerieCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrSerieCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrNumFolioCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrNumFolioCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrNumFactExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrNumFactExtCXP), "")), Null) _
                            & "|" & IIf(optFactura.Value Or optFlete.Value, IIf(IsNull(Trim(vlstrTaxIDExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrTaxIDExtCXP), "")), Null) _
                            & "|" & CStr(txtRFC.Text) & "|" & cboTipoProveedor.Text & "|" & cboPais.ItemData(cboPais.ListIndex)
                            vgstrParametrosSP = vgstrParametrosSP & "|" & dblFlete & "|" & dblRetencionFlete & "|" & txtDescSalida.Text & "|" & dblProvisionadoISR
                                                        
                            lngidfactura = 1
                            frsEjecuta_SP vgstrParametrosSP, "SP_CPINSFACTURACAJACHICA", True, lngidfactura
                            
                            If lngCveImpuestoFlete <> 0 Then
                                pEjecutaSentencia ("insert into CPFACTURACAJACHICAIMPUESTO values (" & lngidfactura & "," & lngCveImpuestoFlete & ",'F')")
                            End If
                        End If
                        
                        If vlblnLicenciaContaElectronica And optFactura.Value Or optFlete.Value Then
                            pEjecutaSentencia "DELETE FROM CPFACTURACAJACHICAXML WHERE INTIDFACTURA = " & lngidfactura
                            
                            If Trim(vlstrXMLCXP) <> "" And (vlintTipoXMLCXP = 1 Or vlintTipoXMLCXP = 2) Then
                                With rsCpCajaChicaXML
                                    .AddNew
                                    !intIdFactura = lngidfactura
                                    !CLBXML = vlstrXMLCXP
                                    .Update
                                End With
                            End If
                        End If
        
                        '------------------------------------------------------------------
                        ' Afectar el corte
                        '------------------------------------------------------------------
                        intcontador = 0
                        Do While intcontador <= UBound(aFormasPago(), 1)
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy"), fdtmServerHora) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(aFormasPago(intcontador).vlintNumFormaPago) _
                            & "|" & CStr(IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, Round(aFormasPago(intcontador).vldblDolares, 2)) * -1) _
                            & "|" & CStr(aFormasPago(intcontador).vldblTipoCambio) _
                            & "|" & IIf(Trim(aFormasPago(intcontador).vlstrFolio) = "", "0", Trim(aFormasPago(intcontador).vlstrFolio)) _
                            & "|" & CStr(llngNumCorte)
                        
                            frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                                                    
                            vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                                                    
                            If Not aFormasPago(intcontador).vlbolEsCredito Then
                                If Trim(aFormasPago(intcontador).vlstrRFC) <> "" And Trim(aFormasPago(intcontador).vlstrBancoSAT) <> "" Then
                                    frsEjecuta_SP llngNumCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(intcontador).vlstrRFC) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(intcontador).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(intcontador).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(intcontador).vldtmFecha))) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                End If
                            End If
                            
                            intcontador = intcontador + 1
                        Loop
                        
                        '------------------------------------------------------------------
                        ' Guardar la póliza
                        '------------------------------------------------------------------
                        ' Cargo a la cuenta del concepto de salida:
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & CStr(lngidfactura) _
                        & "|" & "SC" _
                        & "|" & CStr(arrConceptos(cboConcepto.ListIndex).lngCtaGasto) _
                        & "|" & CStr(Round(IIf(optFactura.Value, _
                                (dblimporteExento + dblImporteNoGravado + dblimportegravado), _
                                 IIf(optFlete.Value, dblFlete, dblTotal)) _
                                * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                        & "|" & "1"
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        
                        ' Cargo al IVA no pagado:
                        If dblImpuesto <> 0 Then
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(glngCtaIVANoPagado) _
                            & "|" & CStr(Round(dblImpuesto * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                            & "|" & "1"
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        End If
                        
                        If optFactura.Value Then
                            ' Cargo al IEPS pagado:
                            If dblIEPS <> 0 Then
                                vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & CStr(lngidfactura) _
                                & "|" & "SC" _
                                & "|" & CStr(glngctaIEPSPagado) _
                                & "|" & CStr(Round(dblIEPS * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                                & "|" & "1"
                                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                            End If
                            
                            ' Agregado para caso 11526
                            ' para obtener la cuenta contable de descuentos sobre compra segun el departamento y tasa de iva
                            ReDim DescuentosporTasaIVA(0)
        
                            ' --Descuento exento--
                            If dblDescImporteExento <> 0 Then
                                vlrelivaentrada = -1
                                vlintNumeroProceso = 1  'Número de proceso del sistema que afecta cuenta de descuentos sobre compra (Salidas de caja chica)
                                pLlenaPolizaDescuento vlintNumeroProceso, vlrelivaentrada, 0, 0, 0, dblDescImporteExento, vgintClaveEmpresaContable
                                If Not vlblnCuentaDescuentoValida Then
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                            End If
                            ' --Descuento no gravado (tasa 0%)--
                            If dblDescImporteNoGravado <> 0 Then
                                vlrelivaentrada = 0
                                vlintNumeroProceso = 1  'Número de proceso del sistema que afecta cuenta de descuentos sobre compra (Salidas de caja chica)
                                pLlenaPolizaDescuento vlintNumeroProceso, vlrelivaentrada, 0, 0, 0, dblDescImporteNoGravado, vgintClaveEmpresaContable
                                If Not vlblnCuentaDescuentoValida Then
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                                
                            End If
                            ' --Descuento gravado--
                            If dblDescImporteGravado <> 0 Then
                                vlrelivaentrada = arrImpuestosFactura(cboImpuesto.ListIndex).dblPorcentaje
                                vlintNumeroProceso = 1  'Número de proceso del sistema que afecta cuenta de descuentos sobre compra (Salidas de caja chica)
                                pLlenaPolizaDescuento vlintNumeroProceso, vlrelivaentrada, 0, 0, 0, dblDescImporteGravado, vgintClaveEmpresaContable
                                If Not vlblnCuentaDescuentoValida Then
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                            End If
                            
                            If dblDescImporteExento <> 0 Or dblDescImporteNoGravado <> 0 Or dblDescImporteGravado <> 0 Then
                                For i = 1 To UBound(DescuentosporTasaIVA())
                                    If Val(Format(DescuentosporTasaIVA(i).ImporteDescuento, "##############.00")) > 0 Then
                                        vgstrParametrosSP = CStr(llngNumCorte) _
                                        & "|" & CStr(lngidfactura) _
                                        & "|" & "SC" _
                                        & "|" & CStr(DescuentosporTasaIVA(i).CuentaDescuento) _
                                        & "|" & CStr(Round(DescuentosporTasaIVA(i).ImporteDescuento * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                                        & "|" & "0"
                                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                                    End If
                                Next i
                            End If
                            
                            'Cargo a la cuenta de flete
                            If dblFlete <> 0 Then
                                vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & CStr(lngidfactura) _
                                & "|" & "SC" _
                                & "|" & CStr(llngCuentaFlete) _
                                & "|" & CStr(Round(dblFlete * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                                & "|" & "1"
                                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                            End If
                            
                            'Abono retencion de flete
                            If dblRetencionFlete <> 0 Then
                                vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & CStr(lngidfactura) _
                                & "|" & "SC" _
                                & "|" & CStr(llngCuentaRetencionFletes) _
                                & "|" & CStr(Round(dblRetencionFlete * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                                & "|" & "0"
                                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                            End If
                            
                            'Abono provisionado de ISR
                            If dblProvisionadoISR <> 0 Then
                                vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & CStr(lngidfactura) _
                                & "|" & "SC" _
                                & "|" & CStr(glngCtaISRprovisionadoResico) _
                                & "|" & CStr(Round(dblProvisionadoISR * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                                & "|" & "0"
                                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                            End If
                            
                            
                        ElseIf optFlete.Value And optRetencion(0).Value Then
                            'retencion de flete
                            If dblRetencionFlete <> 0 Then
                                vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & CStr(lngidfactura) _
                                & "|" & "SC" _
                                & "|" & CStr(llngCuentaRetencionFletes) _
                                & "|" & CStr(Round(dblRetencionFlete * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                                & "|" & "0"
                                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                            End If
                        End If
                                                    
                        ' Abono a la cuenta de la caja chica:
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & CStr(lngidfactura) _
                        & "|" & "SC" _
                        & "|" & CStr(rsCajaChica!INTNUMCUENTACONTABLE) _
                        & "|" & CStr(Round(dblTotal * IIf(optMoneda(0).Value, 1, ldblTipoCambio), 4)) _
                        & "|" & "0"
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        
                        '------------------------------------------------------------------
                        ' Liberar el corte
                        '------------------------------------------------------------------
                        pLiberaCorte llngNumCorte
                    
                        '------------------------------------------------------------------
                        ' Registro de transacciones
                        '------------------------------------------------------------------
                        pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "CAJA CHICA", CStr(lngidfactura)
                        
                        EntornoSIHO.ConeccionSIHO.CommitTrans
                        
                        'La operación se realizó satisfactoriamente.
                        MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                        txtNumero.SetFocus
                    End If
                End If
            End If
        End If
    ElseIf optTipo(1).Value = True Then
        ' HONORARIOS
        If lblnConsulta And lstrEstadoFactura = "A" Then
            llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If llngPersonaGraba = 0 Then
                Exit Sub
            End If
            EntornoSIHO.ConeccionSIHO.BeginTrans
            Set rsCorte = frsRegresaRs("select intNumCorte from CpFacturaCajaChica where intIdFactura = " + Trim(txtNumero.Text))
            If rsCorte.RecordCount <> 0 Then
                llngCorteXml = rsCorte!intNumCorte
            End If
            'si esta activa, se puede asignar el xml
            lngidfactura = Val(txtNumero.Text)
                            vgstrParametrosSP = lngidfactura & "|'A'" _
                            & "|" & CStr(llngPersonaGraba) & "|" & IIf(llngCorteXml = 0, Null, CStr(llngCorteXml)) _
                            & "|" & IIf(IsNull(Trim(vlstrUUIDXMLCXP)), Null, IIf(vlintTipoXMLCXP = 1, Trim(vlstrUUIDXMLCXP), "")) _
                            & "|" & IIf(IsNull(vldblMontoXMLCXP) Or vldblMontoXMLCXP = 0, Null, vldblMontoXMLCXP) _
                            & "|" & IIf(IsNull(Trim(vlstrMonedaXMLCXP)), Null, Trim(vlstrMonedaXMLCXP)) _
                            & "|" & IIf(IsNull(vldblTipoCambioXMLCXP) Or vldblTipoCambioXMLCXP = 0, Null, vldblTipoCambioXMLCXP) _
                            & "|" & IIf(IsNull(Trim(vlintTipoXMLCXP)) Or vlintTipoXMLCXP = 0, Null, Trim(vlintTipoXMLCXP)) _
                            & "|" & IIf(IsNull(Trim(vlstrSerieCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrSerieCXP), "")) _
                            & "|" & IIf(IsNull(Trim(vlstrNumFolioCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrNumFolioCXP), "")) _
                            & "|" & IIf(IsNull(Trim(vlstrNumFactExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrNumFactExtCXP), "")) _
                            & "|" & IIf(IsNull(Trim(vlstrTaxIDExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrTaxIDExtCXP), ""))
                            frsEjecuta_SP vgstrParametrosSP, "SP_CPUPDFACTURACAJACHICA"
            If vlblnLicenciaContaElectronica Then
                pEjecutaSentencia "DELETE FROM CPFACTURACAJACHICAXML WHERE INTIDFACTURA = " & lngidfactura
                If Trim(vlstrXMLCXP) <> "" And (vlintTipoXMLCXP = 1 Or vlintTipoXMLCXP = 2) Then
                    With rsCpCajaChicaXML
                        .AddNew
                        !intIdFactura = lngidfactura
                        !CLBXML = vlstrXMLCXP
                        .Update
                    End With
                End If
            End If
            pGuardarLogTransaccion Me.Name, EnmCambiar, llngPersonaGraba, "CAJA CHICA - Asignar XML", CStr(lngidfactura)
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
            txtNumero.SetFocus
        Else
            If fblnHonorarioValido() Then
                strParametros = mskCuentaHonorario.Text & "|" & CStr(vgintClaveEmpresaContable)
                intNumCuentaHonorario = 1
                frsEjecuta_SP strParametros, "FN_PVSELCUENTA", True, intNumCuentaHonorario
                
                Set rsCnParametro = frsSelParametros("CN", vgintClaveEmpresaContable)
                Do While Not rsCnParametro.EOF
                    If rsCnParametro!Nombre = "INTCTARETENCIONIVA" Then lnCuentaRetencionIVA = rsCnParametro!Valor
                    If rsCnParametro!Nombre = "INTCTARETENCIONISR" Then lnCuentaRetencionISR = rsCnParametro!Valor
                    rsCnParametro.MoveNext
                Loop
                rsCnParametro.Close
                
                EntornoSIHO.ConeccionSIHO.BeginTrans
                llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
                
                If llngNumCorte = 0 Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    
                    'No se encontró un corte abierto.
                    MsgBox SIHOMsg(659), vbExclamation + vbOKOnly, "Mensaje"
                Else
                    '------------------------------------------------------------------
                    ' Bloquear el corte
                    '------------------------------------------------------------------
                    lngCorteGrabando = 1
                    frsEjecuta_SP CStr(llngNumCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, lngCorteGrabando
                    If lngCorteGrabando <> 2 Then
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        
                        'En este momento se está afectando el corte, espere un momento e intente de nuevo.
                        MsgBox SIHOMsg(779), vbExclamation + vbOKOnly, "Mensaje"
                    Else
                        '------------------------------------------------------------------
                        ' Insertar la factura
                        '------------------------------------------------------------------
                        dblMontoExento = IIf(OptIVAHonorario(1).Value = True, Val(Format(txtMontoHonorario.Text, cstrFormato)), 0)
                        dblMonto = Val(Format(txtMontoHonorario.Text, cstrFormato))
                        dblRetencionISR = Val(Format(lblRetencionISRHonorario.Caption, cstrFormato))
                        dblSubTotal = Val(Format(lblSubtotalHonorario.Caption, cstrFormato))
                        dblRetencionIVA = Val(Format(lblRetencionIVAHonorario.Caption, cstrFormato))
                        dblImpuesto = Val(Format(lblIVAHonorario.Caption, cstrFormato))
                        dblTotal = Val(Format(lblTotalPagarHonorario.Caption, cstrFormato))
                        
                        lngCveImpuesto = 0
                        If dblSubTotal <> 0 And cboIVAHonorario.ListIndex <> -1 Then
                            lngCveImpuesto = cboIVAHonorario.ItemData(cboIVAHonorario.ListIndex)
                        Else
                            lngCveImpuesto = 0
                        End If
                        
                        lngCveTarifaISR = 0
                        If dblRetencionISR <> 0 And cboTarifa.ListIndex <> -1 Then
                            lngCveTarifaISR = cboTarifa.ItemData(cboTarifa.ListIndex)
                        Else
                            lngCveTarifaISR = 0
                        End If

                        vgstrParametrosSP = fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) _
                        & "|" & CStr(llngProveedor) & "|" & Trim(txtFolioHonorario.Text) _
                        & "|" & fstrFechaSQL(mskFechaHonorario.Text) & "|" & CStr("") _
                        & "|" & CStr(dblMontoExento) & "|" & "0" _
                        & "|" & CStr(dblMonto) & "|" & CStr(dblRetencionISR) _
                        & "|" & CStr(dblSubTotal) & "|" & CStr(dblRetencionIVA) _
                        & "|" & CStr(dblImpuesto) & "|" & CStr(vgintNumeroDepartamento) _
                        & "|" & CStr(llngPersonaGraba) & "|" & CStr(llngNumCorte) & "|" & "A" _
                        & "|" & IIf(optMonedaHonorario(0).Value, "1", "0") & "|" & CStr(lngCveImpuesto) _
                        & "|" & CStr(IIf(optMonedaHonorario(0).Value, 0, ldblTipoCambioOficial)) _
                        & "|" & IIf(llngProveedor = 0, Trim(cboMedicos.Text), "") & "|" & "H" _
                        & "|" & CStr(intNumCuentaHonorario) & "|" & CStr(lngCveTarifaISR) _
                        & "|" & IIf(IsNull(Trim(vlstrUUIDXMLCXP)), Null, IIf(vlintTipoXMLCXP = 1, Trim(vlstrUUIDXMLCXP), "")) _
                        & "|" & IIf(IsNull(vldblMontoXMLCXP) Or vldblMontoXMLCXP = 0, Null, vldblMontoXMLCXP) _
                        & "|" & Null & "|" & Null _
                        & "|" & IIf(IsNull(Trim(vlstrMonedaXMLCXP)), Null, Trim(vlstrMonedaXMLCXP)) _
                        & "|" & IIf(IsNull(vldblTipoCambioXMLCXP) Or vldblTipoCambioXMLCXP = 0, Null, vldblTipoCambioXMLCXP) _
                        & "|" & IIf(IsNull(Trim(vlintTipoXMLCXP)) Or vlintTipoXMLCXP = 0, Null, Trim(vlintTipoXMLCXP)) _
                        & "|" & IIf(IsNull(Trim(vlstrSerieCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrSerieCXP), "")) _
                        & "|" & IIf(IsNull(Trim(vlstrNumFolioCXP)), Null, IIf(vlintTipoXMLCXP = 2 Or vlintTipoXMLCXP = 3, Trim(vlstrNumFolioCXP), "")) _
                        & "|" & IIf(IsNull(Trim(vlstrNumFactExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrNumFactExtCXP), "")) _
                        & "|" & IIf(IsNull(Trim(vlstrTaxIDExtCXP)), Null, IIf(vlintTipoXMLCXP = 4, Trim(vlstrTaxIDExtCXP), "")) _
                        & "|" & CStr(txtRFcHono.Text) & "|" & Null & "|" & Null
                        vgstrParametrosSP = vgstrParametrosSP & "|" & Null & "|" & Null & "|" & txtDescSalidaHono.Text
                        
                        lngidfactura = 1
                        frsEjecuta_SP vgstrParametrosSP, "SP_CPINSFACTURACAJACHICA", True, lngidfactura
        
                        If vlblnLicenciaContaElectronica Then
                            pEjecutaSentencia "DELETE FROM CPFACTURACAJACHICAXML WHERE INTIDFACTURA = " & lngidfactura
                            
                            If Trim(vlstrXMLCXP) <> "" And (vlintTipoXMLCXP = 1 Or vlintTipoXMLCXP = 2) Then
                                With rsCpCajaChicaXML
                                    .AddNew
                                    !intIdFactura = lngidfactura
                                    !CLBXML = vlstrXMLCXP
                                    .Update
                                End With
                            End If
                        End If
        
                        '------------------------------------------------------------------
                        ' Afectar el corte
                        '------------------------------------------------------------------
                        intcontador = 0
                        Do While intcontador <= UBound(aFormasPago(), 1)
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy"), fdtmServerHora) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(aFormasPago(intcontador).vlintNumFormaPago) _
                            & "|" & CStr(IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, Round(aFormasPago(intcontador).vldblDolares, 2)) * -1) _
                            & "|" & CStr(aFormasPago(intcontador).vldblTipoCambio) _
                            & "|" & IIf(Trim(aFormasPago(intcontador).vlstrFolio) = "", "0", Trim(aFormasPago(intcontador).vlstrFolio)) _
                            & "|" & CStr(llngNumCorte)
                        
                            frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                            
                            vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                                                    
                            If Not aFormasPago(intcontador).vlbolEsCredito Then
                                If Trim(aFormasPago(intcontador).vlstrRFC) <> "" And Trim(aFormasPago(intcontador).vlstrBancoSAT) <> "" Then
                                    frsEjecuta_SP llngNumCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(intcontador).vlstrRFC) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(intcontador).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(intcontador).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(intcontador).vldtmFecha))) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                End If
                            End If
                            
                            intcontador = intcontador + 1
                        Loop
                        
                        '------------------------------------------------------------------
                        ' Guardar la póliza
                        '------------------------------------------------------------------
                        ' Cargo a la cuenta del concepto de salida:
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & CStr(lngidfactura) _
                        & "|" & "SC" _
                        & "|" & CStr(intNumCuentaHonorario) _
                        & "|" & CStr(Round((dblMonto) * IIf(optMonedaHonorario(0).Value, 1, ldblTipoCambioOficial), 4)) _
                        & "|" & "1"
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        
                        ' Cargo al IVA no pagado:
                        If dblImpuesto <> 0 Then
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(glngCtaIVANoPagado) _
                            & "|" & CStr(Round(dblImpuesto * IIf(optMonedaHonorario(0).Value, 1, ldblTipoCambioOficial), 4)) _
                            & "|" & "1"
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        End If
                        
                         ' Abono a la cuenta de Retención del IVA:
    
                        If dblRetencionIVA <> 0 Then
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(lnCuentaRetencionIVA) _
                            & "|" & CStr(Round((dblRetencionIVA) * IIf(optMonedaHonorario(0).Value, 1, ldblTipoCambioOficial), 4)) _
                            & "|" & "0"
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        End If
                        
                        ' Abono a la cuenta de Retención del ISR:
    
                        If dblRetencionISR <> 0 Then
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(lnCuentaRetencionISR) _
                            & "|" & CStr(Round((dblRetencionISR) * IIf(optMonedaHonorario(0).Value, 1, ldblTipoCambioOficial), 4)) _
                            & "|" & "0"
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        End If
                        
                        ' Abono a la cuenta de la caja chica:
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & CStr(lngidfactura) _
                        & "|" & "SC" _
                        & "|" & CStr(rsCajaChica!INTNUMCUENTACONTABLE) _
                        & "|" & CStr(Round(dblTotal * IIf(optMonedaHonorario(0).Value, 1, ldblTipoCambioOficial), 4)) _
                        & "|" & "0"
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        
                        '------------------------------------------------------------------
                        ' Liberar el corte
                        '------------------------------------------------------------------
                        pLiberaCorte llngNumCorte
                    
                        '------------------------------------------------------------------
                        ' Registro de transacciones
                        '------------------------------------------------------------------
                        pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "CAJA CHICA", CStr(lngidfactura)
                        
                        EntornoSIHO.ConeccionSIHO.CommitTrans
                        
                        'La operación se realizó satisfactoriamente.
                        MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                        txtNumero.SetFocus
                        
                        pLimpiaHonorarios
                    End If
                End If
            End If
        End If
    ElseIf optTipo(2).Value = True Then
            If fblnDisminucionValida() Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
                llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
                
                If llngNumCorte = 0 Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    
                    'No se encontró un corte abierto.
                    MsgBox SIHOMsg(659), vbExclamation + vbOKOnly, "Mensaje"
                    Exit Sub
                Else
                    '------------------------------------------------------------------
                    ' Bloquear el corte
                    '------------------------------------------------------------------
                    lngCorteGrabando = 1
                    frsEjecuta_SP CStr(llngNumCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, lngCorteGrabando
                    If lngCorteGrabando <> 2 Then
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        
                        'En este momento se está afectando el corte, espere un momento e intente de nuevo.
                        MsgBox SIHOMsg(779), vbExclamation + vbOKOnly, "Mensaje"
                        Exit Sub
                    Else
                        '------------------------------------------------------------------
                        ' Insertar la disminución de fondo de caja chica
                        '------------------------------------------------------------------
                        vgstrParametrosSP = fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) & "|" & fstrFechaSQL(MaskDisminucionFecha.Text) & "|" & "" & "|" & _
                        CDbl(txtTotalDF.Text) & "|" & CStr(vgintNumeroDepartamento) & "|" & CStr(llngPersonaGraba) & "|" & CStr(llngNumCorte) & "|" & "A" _
                        & "|" & IIf(OptMonedaDF(0).Value, "1", "0") & "|" & CStr(IIf(OptMonedaDF(0).Value, 0, ldblTipoCambioOficial)) & "|" & "D" & "|" & 0
                        lngidfactura = 1
                        frsEjecuta_SP vgstrParametrosSP, "SP_CPINSDFONDOCAJACHICA", True, lngidfactura
                                            
                        '------------------------------------------------------------------
                        ' Afectar el corte
                        '------------------------------------------------------------------
                        intcontador = 0
                        Do While intcontador <= UBound(aFormasPago(), 1)
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy"), fdtmServerHora) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(aFormasPago(intcontador).vlintNumFormaPago) _
                            & "|" & CStr(IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, Round(aFormasPago(intcontador).vldblDolares, 2)) * -1) _
                            & "|" & CStr(aFormasPago(intcontador).vldblTipoCambio) _
                            & "|" & "" & "|" & CStr(llngNumCorte)
                        
                            frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                            
                            vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                                                    
                            If Not aFormasPago(intcontador).vlbolEsCredito Then
                                If Trim(aFormasPago(intcontador).vlstrRFC) <> "" And Trim(aFormasPago(intcontador).vlstrBancoSAT) <> "" Then
                                    frsEjecuta_SP llngNumCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(intcontador).vlstrRFC) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(intcontador).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(intcontador).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(intcontador).vldtmFecha))) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                End If
                            End If
                            
                            intcontador = intcontador + 1
                        Loop
                        
                        '------------------------------------------------------------------
                        ' Guardar la póliza
                        '------------------------------------------------------------------
                                                              
                        ' Cargo a la cuenta del departamento de caja chica:
                        vgstrParametrosSP = CStr(llngNumCorte) _
                        & "|" & CStr(lngidfactura) _
                        & "|" & "SC" _
                        & "|" & CStr(rsCajaChica!INTNUMCUENTACONTABLE) _
                        & "|" & CStr(Round(txtTotalDF.Text * IIf(OptMonedaDF(0).Value, 1, ldblTipoCambioOficial), 4)) _
                        & "|" & "1" & "|" & ""
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        
                        intcontador = 0
                        Do While intcontador <= UBound(aFormasPago(), 1)
                            '------------------------------------------------------------------
                            ' Abono a la cuenta de las cuentas bancarias seleccionadas:
                            Set rsCuentaCuenta = frsRegresaRs("SELECT INTCUENTACONTABLE FROM PVFORMAPAGO WHERE INTFORMAPAGO = " & aFormasPago(intcontador).vlintNumFormaPago, adLockOptimistic, adOpenForwardOnly)
                            If rsCuentaCuenta.RecordCount > 0 Then
                                intNumCuentaBancaria = rsCuentaCuenta!INTCUENTACONTABLE
                            Else
                                intNumCuentaBancaria = 0
                            End If
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & CStr(lngidfactura) _
                            & "|" & "SC" _
                            & "|" & CStr(intNumCuentaBancaria) _
                            & "|" & CStr(Round(txtTotalDF.Text * IIf(OptMonedaDF(0).Value, 1, ldblTipoCambioOficial), 4)) _
                            & "|" & "0" & "|" & ""
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                            
                            '----- Guardar información de la forma de pago en tabla intermedia -----'
                            vgstrParametrosSP = llngNumCorte & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy"), Format(fdtmServerHora, "hh:mm:ss")) & "|" & aFormasPago(intcontador).vlintNumFormaPago & "|" & aFormasPago(intcontador).lngIdBanco & "|" & _
                                                IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, aFormasPago(intcontador).vldblDolares) * -1 & "|" & IIf(aFormasPago(intcontador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(intcontador).vldblTipoCambio & "|" & _
                                                fstrTipoMovimientoForma(aFormasPago(intcontador).vlintNumFormaPago) & "|" & "SC" & "|" & lngidfactura & "|" & llngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy"), fdtmServerHora) & "|" & "1" & "|" & cgstrModulo
                            frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                                                            
                            intcontador = intcontador + 1
                        Loop
                            
                        '------------------------------------------------------------------
                        ' Liberar el corte
                        '------------------------------------------------------------------
                        pLiberaCorte llngNumCorte
                    
                        '------------------------------------------------------------------
                        ' Registro de transacciones
                        '------------------------------------------------------------------
                        pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "DISMINUCIÓN DE FONDO CAJA CHICA", CStr(txtNumero.Text)
                        
                        EntornoSIHO.ConeccionSIHO.CommitTrans
                        'La operación se realizó satisfactoriamente.
                        MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                        pLimpiaDF
                        txtNumero.SetFocus
                    End If
                End If
            End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
End Sub

Private Function fblnDisminucionValida() As Boolean
    On Error GoTo NotificaError
    Dim dblCantidad As Double
    Dim rsFondoActual As ADODB.Recordset
    
    Dim intcontador As Integer
    Dim dblSumatoriaPesos As Double
    Dim dblSumatoriaDolares As Double

fblnDisminucionValida = True
    
    If Not IsDate(MaskDisminucionFecha.Text) Then
        fblnDisminucionValida = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        MaskDisminucionFecha.SetFocus
        Exit Function
    End If
    If CDate(MaskDisminucionFecha.Text) > fdtmServerFecha Then
        fblnDisminucionValida = False
        '¡La fecha debe ser menor o igual a la del sistema!
        MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
        MaskDisminucionFecha.SetFocus
        pSelMkTexto MaskDisminucionFecha
        Exit Function
    End If
    If txtTotalDF.Text < 0 Or txtTotalDF.Text = "" Or txtTotalDF.Text = 0 Then
        fblnDisminucionValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        txtTotalDF.SetFocus
        pSelTextBox txtTotalDF
        dblTotal = txtTotalDF.Text
        Exit Function
    End If
    If OptMonedaDF(0).Value = False And OptMonedaDF(1).Value = False Then
        fblnDisminucionValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        OptMonedaDF(0).SetFocus
        Exit Function
    End If
            
    ' Que se firme correcto
    If fblnDisminucionValida Then
       llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
       fblnDisminucionValida = llngPersonaGraba <> 0
    End If
    
    If fblnDisminucionValida Then
        ldblTipoCambioOficial = fdblTipoCambio(CDate(MaskDisminucionFecha.Text), "O")
        fblnDisminucionValida = fblnFormasPagoPos(aFormasPago(), txtTotalDF.Text, IIf(OptMonedaDF(0).Value, True, False), ldblTipoCambioOficial, False, 0, "")
    End If
    
    'Recorre la formas de pago para saber si las sumatoria de las formas de pago en pesos y en dólares no supera el fondo del corte de caja chica
    If fblnDisminucionValida Then
        intcontador = 0
        dblSumatoriaPesos = 0
        dblSumatoriaDolares = 0
        
        Do While intcontador <= UBound(aFormasPago(), 1)
            If aFormasPago(intcontador).vldblTipoCambio = 0 Then
                dblSumatoriaPesos = dblSumatoriaPesos + aFormasPago(intcontador).vldblCantidad
            Else
                dblSumatoriaDolares = dblSumatoriaDolares + aFormasPago(intcontador).vldblDolares
            End If
            intcontador = intcontador + 1
        Loop
        
        If dblSumatoriaPesos <> 0 Then
            Set rsFondoActual = frsRegresaRs("SELECT NVL(SUM(MNYCANTIDADPAGADA),0) TOTAL FROM PVDETALLECORTE WHERE INTNUMCORTE = " & llngNumCorteValidacionImporte & " AND MNYTIPOCAMBIO = 0", adLockOptimistic, adOpenForwardOnly)
            If dblSumatoriaPesos > rsFondoActual!Total Then
                fblnDisminucionValida = False
                '¡El importe de la salida de dinero en pesos no puede ser mayor que el fondo actual del corte de caja chica en dicha moneda!
                MsgBox SIHOMsg(1447), vbExclamation + vbOKOnly, "Mensaje"
                Exit Function
            End If
        End If
        
        If dblSumatoriaDolares <> 0 Then
            Set rsFondoActual = frsRegresaRs("SELECT NVL(SUM(MNYCANTIDADPAGADA),0) TOTAL FROM PVDETALLECORTE WHERE INTNUMCORTE = " & llngNumCorteValidacionImporte & " AND MNYTIPOCAMBIO <> 0", adLockOptimistic, adOpenForwardOnly)
            If dblSumatoriaDolares > rsFondoActual!Total Then
                fblnDisminucionValida = False
                '¡El importe de la salida de dinero en dólares no puede ser mayor que el fondo actual del corte de caja chica en dicha moneda!
                MsgBox SIHOMsg(1448), vbExclamation + vbOKOnly, "Mensaje"
                Exit Function
            End If
        End If
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDisminucionValida"))
End Function
Private Function fblnHonorarioValido() As Boolean
    On Error GoTo NotificaError
    
    Dim dblCantidad As Double
    Dim vlstrRFC As String
    Dim rsrfc As New ADODB.Recordset
    Dim rsFondoActual As New ADODB.Recordset
    Dim lngCveProveedor As Long
    
    Dim intcontador As Integer
    Dim dblSumatoriaPesos As Double
    Dim dblSumatoriaDolares As Double
    
    fblnHonorarioValido = True
    
    If cboMedicos.Text = "" Then
            fblnHonorarioValido = False
            MsgBox "Seleccione el médico.", vbExclamation + vbOKOnly, "Mensaje"
            cboMedicos.SetFocus
        Exit Function
    End If

    'checar que el RFC del médico sea válido
    If cboMedicos.ListIndex = -1 Then
        If txtRFcHono.Text = "" Then
            fblnHonorarioValido = False
            MsgBox "Favor de ingresar el RFC del médico.", vbOKOnly + vbExclamation, "Mensaje"
            txtRFcHono.SetFocus
        ElseIf cboMedicos.ListIndex = -1 Then
            If Len(txtRFcHono.Text) <> 12 And Len(txtRFcHono.Text) <> 13 Then
                fblnHonorarioValido = False
                MsgBox SIHOMsg(1345), vbExclamation + vbOKOnly, "Mensaje"
                txtRFcHono.SetFocus
            End If
        End If
    End If
    
    If fblnHonorarioValido Then
        llngProveedor = 0
        If Trim(cboMedicos.Text) <> "" Then
            lngCveProveedor = 1
            frsEjecuta_SP UCase(Trim(cboMedicos.Text)), "sp_PvSelCveProveedor", True, lngCveProveedor
            llngProveedor = lngCveProveedor
        End If
    End If
    
    If fblnHonorarioValido And Not IsDate(mskFechaHonorario.Text) Then
        fblnHonorarioValido = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaHonorario.SetFocus
    End If
    If fblnHonorarioValido Then
        If CDate(mskFechaHonorario.Text) > fdtmServerFecha Then
            fblnHonorarioValido = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaHonorario.SetFocus
        End If
    End If
    If fblnHonorarioValido And Not (optMonedaHonorario(0).Value Or optMonedaHonorario(1).Value) Then
        fblnHonorarioValido = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        optMonedaHonorario(0).SetFocus
    End If
    ldblTipoCambioOficial = fdblTipoCambio(CDate(mskFechaHonorario.Text), "O")
    If fblnHonorarioValido And optMonedaHonorario(1).Value Then
        If ldblTipoCambioOficial = 0 Then
            fblnHonorarioValido = False
            'No está registrado el tipo de cambio del día.
            MsgBox SIHOMsg(231), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    If fblnHonorarioValido And Val(Format(txtMontoHonorario.Text, cstrFormato)) = 0 Then
        fblnHonorarioValido = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        txtMontoHonorario.SetFocus
    End If
    If fblnHonorarioValido And OptIVAHonorario(0).Value = True And cboIVAHonorario.ListIndex = -1 Then
        fblnHonorarioValido = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbOKOnly + vbExclamation, "Mensaje"
        cboIVAHonorario.SetFocus
    End If
    If fblnHonorarioValido And chkRetencionISRHonorario.Value = 1 And cboTarifa.ListIndex = -1 Then
        fblnHonorarioValido = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbOKOnly + vbExclamation, "Mensaje"
        cboTarifa.SetFocus
    End If
    If fblnHonorarioValido And mskCuentaHonorario.ClipText = "" Then
        fblnHonorarioValido = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        mskCuentaHonorario.SetFocus
    End If

    ' Que el departamento sea una caja chica
    If fblnHonorarioValido Then
        If rsCajaChica.RecordCount = 0 Then
            fblnHonorarioValido = False
            'Este departamento no es una caja chica.
            MsgBox SIHOMsg(806), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
        
    ' Que se firme correcto
    If fblnHonorarioValido Then
       llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
       fblnHonorarioValido = llngPersonaGraba <> 0
    End If
    
    ' Que se hayan seleccionado formas de pago
    If fblnHonorarioValido Then
        If optMonedaHonorario(1).Value Then
            dblCantidad = Val(Format(lblTotalPagarHonorario.Caption, cstrFormato)) * ldblTipoCambioOficial
        Else
            dblCantidad = Val(Format(lblTotalPagarHonorario.Caption, cstrFormato))
        End If
    
        vlstrRFC = ""
        If cboMedicos.ListIndex <> -1 Then
            Set rsrfc = frsRegresaRs("SELECT vchrfc FROM COPROVEEDOR WHERE intcveproveedor = " & cboMedicos.ItemData(cboMedicos.ListIndex), adLockOptimistic, adOpenForwardOnly)
            If rsrfc.RecordCount > 0 Then
                vlstrRFC = rsrfc!vchRFC
            End If
        End If

        fblnHonorarioValido = fblnFormasPagoPos(aFormasPago(), dblCantidad, True, ldblTipoCambioOficial, False, 0, "", Trim(Replace(Replace(Replace(vlstrRFC, "-", ""), "_", ""), " ", "")))
    End If
    
    'Recorre la formas de pago para saber si las sumatoria de las formas de pago en pesos y en dólares no supera el fondo del corte de caja chica
    If fblnHonorarioValido Then
        intcontador = 0
        dblSumatoriaPesos = 0
        dblSumatoriaDolares = 0
        
        Do While intcontador <= UBound(aFormasPago(), 1)
            If aFormasPago(intcontador).vldblTipoCambio = 0 Then
                dblSumatoriaPesos = dblSumatoriaPesos + aFormasPago(intcontador).vldblCantidad
            Else
                dblSumatoriaDolares = dblSumatoriaDolares + aFormasPago(intcontador).vldblDolares
            End If
            intcontador = intcontador + 1
        Loop
        
        If dblSumatoriaPesos <> 0 Then
            Set rsFondoActual = frsRegresaRs("SELECT NVL(SUM(MNYCANTIDADPAGADA),0) TOTAL FROM PVDETALLECORTE WHERE INTNUMCORTE = " & llngNumCorteValidacionImporte & " AND MNYTIPOCAMBIO = 0", adLockOptimistic, adOpenForwardOnly)
            If dblSumatoriaPesos > rsFondoActual!Total Then
                fblnHonorarioValido = False
                '¡El importe de la salida de dinero en pesos no puede ser mayor que el fondo actual del corte de caja chica en dicha moneda!
                MsgBox SIHOMsg(1447), vbExclamation + vbOKOnly, "Mensaje"
                Exit Function
            End If
        End If
        
        If dblSumatoriaDolares <> 0 Then
            Set rsFondoActual = frsRegresaRs("SELECT NVL(SUM(MNYCANTIDADPAGADA),0) TOTAL FROM PVDETALLECORTE WHERE INTNUMCORTE = " & llngNumCorteValidacionImporte & " AND MNYTIPOCAMBIO <> 0", adLockOptimistic, adOpenForwardOnly)
            If dblSumatoriaDolares > rsFondoActual!Total Then
                fblnHonorarioValido = False
                '¡El importe de la salida de dinero en dólares no puede ser mayor que el fondo actual del corte de caja chica en dicha moneda!
                MsgBox SIHOMsg(1448), vbExclamation + vbOKOnly, "Mensaje"
                Exit Function
            End If
        End If
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnHonorarioValido"))
End Function

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    Dim dblCantidad As Double
    Dim lngCveProveedor As Long
    Dim vlstrRFC As String
    Dim rsrfc As New ADODB.Recordset
    Dim rsFondoActual As New ADODB.Recordset
    
    Dim intcontador As Integer
    Dim dblSumatoriaPesos As Double
    Dim dblSumatoriaDolares As Double
    Dim blnIva As Boolean
    Dim lstrsql As String
    Dim rsCuenta As New ADODB.Recordset
    
    fblnDatosValidos = True
    
    ' Que tenga permisos de guardar
    fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, 1646, "E")
    
    ' Que tenga un proveedor seleccionado o ingresado
    If cboProveedor.Text = "" Then
            fblnDatosValidos = False
            MsgBox SIHOMsg(1456), vbExclamation + vbOKOnly, "Mensaje"
            cboProveedor.SetFocus
        Exit Function
    End If
    
    'Que tenga un RFC o que el tamaño sea valido
    If cboProveedor.ListIndex = -1 And optFactura.Value Then
        If txtRFC.Text = "" Then
            fblnDatosValidos = False
            MsgBox "Favor de ingresar el RFC del proveedor o acreedor.", vbOKOnly + vbExclamation, "Mensaje"
            txtRFC.SetFocus
        ElseIf cboProveedor.ListIndex = -1 Then
            If Len(txtRFC.Text) <> 12 And Len(txtRFC.Text) <> 13 Then
                fblnDatosValidos = False
                MsgBox SIHOMsg(1345), vbExclamation + vbOKOnly, "Mensaje"
                txtRFC.SetFocus
            End If
        End If
    End If

    ' Que si hay importe gravado el IVA sea mayor a CERO
    If fblnDatosValidos And optFactura.Value Then
        If Val(Format(txtImporteGravado.Text, cstrFormato)) > 0 And Val(Format(lblImporteIVA.Caption, cstrFormato)) = 0 Then
            fblnDatosValidos = False
            'Dato incorrecto: El valor debe ser
            MsgBox SIHOMsg(36) & "mayor a cero.", vbExclamation + vbOKOnly, "Mensaje"
            If fblnCanFocus(cboImpuesto) Then cboImpuesto.SetFocus
        End If
    End If
    
    ' Que el departamento sea una caja chica
    If fblnDatosValidos Then
        If rsCajaChica.RecordCount = 0 Then
            fblnDatosValidos = False
            'Este departamento no es una caja chica.
            MsgBox SIHOMsg(806), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If

    If fblnDatosValidos Then
       'llngProveedor = 0
        If Trim(cboProveedor.Text) <> "" Then
          'If CStr(llngProveedor) = "0" Then
             lngCveProveedor = 1
             frsEjecuta_SP UCase(Trim(cboProveedor.Text)), "sp_PvSelCveProveedor", True, lngCveProveedor
             llngProveedor = lngCveProveedor
          'End If
           
        End If
    End If
    
    ' Que se haya elegido un concepto
    If fblnDatosValidos And cboConcepto.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbExclamation + vbOKOnly, "Mensaje"
        cboConcepto.SetFocus
    End If

    ' Que se haya introducido una fecha válida
    If fblnDatosValidos And Not IsDate(mskFecha.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
        mskFecha.SetFocus
    End If

    ' Que la fecha sea menor o igual a la actual
    If fblnDatosValidos Then
        If CDate(mskFecha.Text) > ldtmFecha Then
            fblnDatosValidos = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbExclamation + vbOKOnly, "Mensaje"
            mskFecha.SetFocus
        End If
    End If

    ' Que se haya introducido un folio de factura
    If fblnDatosValidos And llngProveedor > 0 And Trim(txtFolio.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
        txtFolio.SetFocus
    End If
    
    If optFactura.Value Then
        ' Que se escrito algún importe, gravado o no
        If fblnDatosValidos And Val(Format(txtImporteExento.Text, cstrFormato)) = 0 And Val(Format(txtImporteNoGravado.Text, cstrFormato)) = 0 And Val(Format(txtImporteGravado.Text, cstrFormato)) = 0 Then
            fblnDatosValidos = False
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
            txtImporteExento.SetFocus
        End If

        ' Que se haya elegido un impuesto
        If fblnDatosValidos And Val(Format(txtImporteGravado.Text, cstrFormato)) <> 0 And cboImpuesto.ListIndex = -1 Then
            fblnDatosValidos = False
            '¡Dato no válido, seleccione un valor de la lista!
            MsgBox SIHOMsg(3), vbExclamation + vbOKOnly, "Mensaje"
            cboImpuesto.SetFocus
        End If
    
        ' Que el total a pagar sea correcto
        If fblnDatosValidos And Val(Format(lblTotal.Caption, cstrFormato)) < 0 Then
            fblnDatosValidos = False
            'El total a pagar está incorrecto!
            MsgBox SIHOMsg(221), vbExclamation + vbOKOnly, "Mensaje"
            txtImporteExento.SetFocus
        End If
        
        'Cuenta de flete
        If fblnDatosValidos And Val(Format(txtFleteFactura.Text, "###########.0000")) > 0 Then
            lstrsql = "Select CpCuentaFlete.intNumeroCuenta, CnCuenta.vchCuentaContable CuentaFlete " & _
                     "From CpCuentaFlete Inner join CnCuenta On CpCuentaFlete.intNumeroCuenta = CnCuenta.intNumeroCuenta " & _
                     "Where CpCuentaFlete.smiCveDepartamento = -1 Or CpCuentaFlete.smiCveDepartamento = " & vgintNumeroDepartamento
            
            Set rsCuenta = frsRegresaRs(lstrsql)
            If rsCuenta.RecordCount > 0 Then
                If Not fblnCuentaAfectable(Trim(rsCuenta!CuentaFlete), vgintClaveEmpresaContable) Then
                    fblnDatosValidos = False
                    'La cuenta de fletes no está activa.
                    MsgBox SIHOMsg(446), vbOKOnly + vbInformation, "Mensaje"
                Else
                    llngCuentaFlete = rsCuenta!intNumeroCuenta
                End If
            Else
                fblnDatosValidos = False
                '¡No está registrada la cuenta para contabilizar la retención por fletes del departamento!
                MsgBox SIHOMsg(847), vbOKOnly + vbExclamation, "Mensaje"
            End If
        End If
        'cuenta retencion flete
        If fblnDatosValidos And Val(Format(txtFleteFactura.Text, "###########.0000")) > 0 Then
            If optRetencionFactura(0).Value Then
                If ldblPorcentajeRetencionFletes = 0 Then
                    'No se encuentra registrado el porcentaje de retención por fletes.
                    MsgBox SIHOMsg(263), vbOKOnly + vbInformation, "Mensaje"
                    fblnDatosValidos = False
                End If
                If fblnDatosValidos Then
                    If llngCuentaRetencionFletes = 0 Then
                        'No se encuentra registrada la cuenta para retención del pago por fletes.
                        MsgBox SIHOMsg(265), vbOKOnly + vbInformation, "Mensaje"
                        fblnDatosValidos = False
                    ElseIf Not fblnCuentaAfectable(fstrCuentaContable(llngCuentaRetencionFletes), vgintClaveEmpresaContable) Then
                        'La cuenta seleccionada no acepta movimientos.
                        MsgBox SIHOMsg(375) & " " & fstrCuentaContable(llngCuentaRetencionFletes) & " " & fstrDescripcionCuenta(fstrCuentaContable(llngCuentaRetencionFletes), vgintClaveEmpresaContable) & " Cuenta de retención por fletes.", vbOKOnly + vbInformation, "Mensaje"
                        fblnDatosValidos = False
                    End If
                End If
            End If
        End If
        'Cuenta de retención de ISR
        If fblnDatosValidos And Val(Format(lblRetencionISR.Caption, "###########.0000")) > 0 Then
        
        If cboProveedor.ListIndex <> -1 Then
            lngCveProveedor = cboProveedor.ItemData(cboProveedor.ListIndex)
        Else
            lngCveProveedor = -1
        End If
                
        glngCtaISRprovisionadoResico = TraeCuentasISRProv(lngCveProveedor, vgintClaveEmpresaContable, "CO", "P", "P")
        
            If glngCtaISRprovisionadoResico = 0 Then
                fblnDatosValidos = False
                'No está registrada la cuenta para contabilizar el ISR provisionado en el Régimen Simplificado de Confianza.
                MsgBox "No está registrada la cuenta para contabilizar el ISR provisionado en el régimen simplificado de confianza.", vbOKOnly + vbExclamation, "Mensaje"
            Else
                If fblnDatosValidos And Not fblnCuentaAfectable(fstrCuentaContable(glngCtaISRprovisionadoResico), vgintClaveEmpresaContable) Then
                    fblnDatosValidos = False
                    'La cuenta de ISR provisionado en el Régimen Simplificado de Confianza no acepta movimientos.
                    MsgBox "La cuenta de ISR provisionado en el Régimen Simplificado de Confianza no acepta movimientos.", vbOKOnly + vbInformation, "Mensaje"
                End If
            End If
        End If
        
    ElseIf optFlete.Value Then
        If fblnDatosValidos And Val(Format(lblTotalFlete.Caption, cstrFormato)) <= 0 Then
            fblnDatosValidos = False
            'El total a pagar está incorrecto!
            MsgBox SIHOMsg(221), vbExclamation + vbOKOnly, "Mensaje"
            If fblnCanFocus(txtImporteFlete) Then txtImporteFlete.SetFocus
        End If
        If fblnDatosValidos And optRetencion(0).Value Then
            If ldblPorcentajeRetencionFletes = 0 Then
                'No se encuentra registrado el porcentaje de retención por fletes.
                MsgBox SIHOMsg(263), vbOKOnly + vbInformation, "Mensaje"
                fblnDatosValidos = False
            End If
            If fblnDatosValidos Then
                If llngCuentaRetencionFletes = 0 Then
                    'No se encuentra registrada la cuenta para retención del pago por fletes.
                    MsgBox SIHOMsg(265), vbOKOnly + vbInformation, "Mensaje"
                    fblnDatosValidos = False
                ElseIf Not fblnCuentaAfectable(fstrCuentaContable(llngCuentaRetencionFletes), vgintClaveEmpresaContable) Then
                    'La cuenta seleccionada no acepta movimientos.
                    MsgBox SIHOMsg(375) & " " & fstrCuentaContable(llngCuentaRetencionFletes) & " " & fstrDescripcionCuenta(fstrCuentaContable(llngCuentaRetencionFletes), vgintClaveEmpresaContable) & " Cuenta de retención por fletes.", vbOKOnly + vbInformation, "Mensaje"
                    fblnDatosValidos = False
                End If
            End If
        End If
    ElseIf optNota.Value Or optTicket.Value Then
        If fblnDatosValidos And Val(Format(txtTotalTicket.Text, cstrFormato)) <= 0 Then
            fblnDatosValidos = False
            'El total a pagar está incorrecto!
            MsgBox SIHOMsg(221), vbExclamation + vbOKOnly, "Mensaje"
            If fblnCanFocus(txtTotalTicket) Then txtTotalTicket.SetFocus
        End If
    End If
    
    ' Que haya tipo de cambio oficial del dia de la factura
    If fblnDatosValidos Then
        ldblTipoCambio = fdblTipoCambio(CDate(mskFecha.Text), "O")
        If ldblTipoCambio = 0 Then
            fblnDatosValidos = False
            'No está registrado el tipo de cambio del día.
            MsgBox SIHOMsg(231) & " " & UCase(Format(mskFecha.Text, "dd/mmm/yyyy")), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
    
    blnIva = False
    If optFactura.Value Or optFlete.Value Then
        If optFactura.Value And Val(Format(lblImporteIVA.Caption, cstrFormato)) <> 0 Then
            blnIva = True
        ElseIf optFlete.Value And Val(Format(lblImporteIvaFlete.Caption, cstrFormato)) <> 0 Then
            blnIva = True
        End If
    End If
    ' Que esté registrada la cuenta para IVA pagado
    If fblnDatosValidos And blnIva And glngCtaIVAPagado = 0 Then
        fblnDatosValidos = False
        'No se encuentran registradas las cuentas de IVA pagado y no pagado en los parámetros generales del sistema.
        MsgBox SIHOMsg(760), vbExclamation + vbOKOnly, "Mensaje"
    End If
        
    ' Que esté registrada la cuenta para IEPS pagado
    If fblnDatosValidos And optFactura.Value And Val(Format(txtIEPS.Text, cstrFormato)) <> 0 And glngctaIEPSPagado = 0 Then
        fblnDatosValidos = False
        'No se encuentran registradas las cuentas de IVA pagado y no pagado en los parámetros generales del sistema.
        MsgBox SIHOMsg(1347), vbExclamation + vbOKOnly, "Mensaje"
    End If

    ' Que la cuenta de gasto del concepto sea afectable
    If fblnDatosValidos Then
        If Not fblnCuentaAfectable(fstrCuentaContable(arrConceptos(cboConcepto.ListIndex).lngCtaGasto), vgintClaveEmpresaContable) Then
            fblnDatosValidos = False
            'La cuenta seleccionada no acepta movimientos.
            MsgBox SIHOMsg(375) & " " & fstrCuentaContable(arrConceptos(cboConcepto.ListIndex).lngCtaGasto) & " " & fstrDescripcionCuenta(fstrCuentaContable(arrConceptos(cboConcepto.ListIndex).lngCtaGasto), vgintClaveEmpresaContable) & " Cuenta del concepto de salida.", vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If

    ' Que la cuenta de descuento del concepto sea afectable
   ' If fblnDatosValidos And optFactura.Value Then
   '     If Not fblnCuentaAfectable(fstrCuentaContable(arrConceptos(cboConcepto.ListIndex).lngCtaDescuento), vgintClaveEmpresaContable) And (Val(Format(txtDescuentoGravado.Text, cstrFormato)) <> 0 Or Val(Format(txtDescuentoNoGravado.Text, cstrFormato)) <> 0 Or Val(Format(txtDescuentoExento.Text, cstrFormato)) <> 0) Then
   '         fblnDatosValidos = False
   '         'La cuenta seleccionada no acepta movimientos.
   '         MsgBox SIHOMsg(375) & " " & fstrCuentaContable(arrConceptos(cboConcepto.ListIndex).lngCtaDescuento) & " " & fstrDescripcionCuenta(fstrCuentaContable(arrConceptos(cboConcepto.ListIndex).lngCtaDescuento), vgintClaveEmpresaContable) & " Cuenta de descuento del concepto de salida.", vbExclamation + vbOKOnly, "Mensaje"
   '     End If
   ' End If

    ' Que la cuenta para IVA pagado sea afectable
    If fblnDatosValidos Then
        If Not fblnCuentaAfectable(fstrCuentaContable(glngCtaIVAPagado), vgintClaveEmpresaContable) And blnIva Then
            fblnDatosValidos = False
            'La cuenta seleccionada no acepta movimientos.
            MsgBox SIHOMsg(375) & " " & fstrCuentaContable(glngCtaIVAPagado) & " " & fstrDescripcionCuenta(fstrCuentaContable(glngCtaIVAPagado), vgintClaveEmpresaContable) & " Cuenta de la caja chica del departamento.", vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If

    ' Que la cuenta para IEPS pagado sea afectable
    If fblnDatosValidos And optFactura.Value Then
        If Not fblnCuentaAfectable(fstrCuentaContable(glngctaIEPSPagado), vgintClaveEmpresaContable) And Val(Format(txtIEPS.Text, cstrFormato)) <> 0 Then
            fblnDatosValidos = False
            'La cuenta seleccionada no acepta movimientos.
            MsgBox SIHOMsg(375) & " " & fstrCuentaContable(glngctaIEPSPagado) & " " & fstrDescripcionCuenta(fstrCuentaContable(glngctaIEPSPagado), vgintClaveEmpresaContable) & " Cuenta de la caja chica del departamento.", vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If

    ' Que la cuenta de caja chica sea afectable
    If fblnDatosValidos Then
        If Not fblnCuentaAfectable(fstrCuentaContable(rsCajaChica!INTNUMCUENTACONTABLE), vgintClaveEmpresaContable) Then
            fblnDatosValidos = False
            'La cuenta seleccionada no acepta movimientos.
            MsgBox SIHOMsg(375) & " " & fstrCuentaContable(rsCajaChica!INTNUMCUENTACONTABLE) & " " & fstrDescripcionCuenta(fstrCuentaContable(rsCajaChica!INTNUMCUENTACONTABLE), vgintClaveEmpresaContable), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
    
    ' Que se firme correcto
    If fblnDatosValidos Then
       llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
       fblnDatosValidos = llngPersonaGraba <> 0
    End If

    ' Que se hayan seleccionado formas de pago
    If fblnDatosValidos Then
        If optMoneda(1).Value Then
            If optFactura.Value Then
                dblCantidad = Val(Format(lblTotal.Caption, cstrFormato)) * ldblTipoCambio
            ElseIf optFlete.Value Then
                dblCantidad = Val(Format(lblTotalFlete.Caption, cstrFormato)) * ldblTipoCambio
            Else
                dblCantidad = Val(Format(txtTotalTicket, cstrFormato)) * ldblTipoCambio
            End If
        Else
            If optFactura.Value Then
                dblCantidad = Val(Format(lblTotal.Caption, cstrFormato))
            ElseIf optFlete.Value Then
                dblCantidad = Val(Format(lblTotalFlete.Caption, cstrFormato))
            Else
                dblCantidad = Val(Format(txtTotalTicket, cstrFormato))
            End If
        End If
        
        vlstrRFC = ""
        If cboProveedor.ListIndex <> -1 Then
            Set rsrfc = frsRegresaRs("SELECT vchrfc FROM COPROVEEDOR WHERE intcveproveedor = " & cboProveedor.ItemData(cboProveedor.ListIndex), adLockOptimistic, adOpenForwardOnly)
            If rsrfc.RecordCount > 0 Then
                vlstrRFC = rsrfc!vchRFC
            End If
        End If

        fblnDatosValidos = fblnFormasPagoPos(aFormasPago(), dblCantidad, True, ldblTipoCambio, False, 0, "", Trim(Replace(Replace(Replace(vlstrRFC, "-", ""), "_", ""), " ", "")))
    End If
    
    'Recorre la formas de pago para saber si las sumatoria de las formas de pago en pesos y en dólares no supera el fondo del corte de caja chica
    If fblnDatosValidos Then
        intcontador = 0
        dblSumatoriaPesos = 0
        dblSumatoriaDolares = 0
        
        Do While intcontador <= UBound(aFormasPago(), 1)
            If aFormasPago(intcontador).vldblTipoCambio = 0 Then
                dblSumatoriaPesos = dblSumatoriaPesos + aFormasPago(intcontador).vldblCantidad
            Else
                dblSumatoriaDolares = dblSumatoriaDolares + aFormasPago(intcontador).vldblDolares
            End If
            intcontador = intcontador + 1
        Loop
        
        If dblSumatoriaPesos <> 0 Then
            Set rsFondoActual = frsRegresaRs("SELECT NVL(SUM(MNYCANTIDADPAGADA),0) TOTAL FROM PVDETALLECORTE WHERE INTNUMCORTE = " & llngNumCorteValidacionImporte & " AND MNYTIPOCAMBIO = 0", adLockOptimistic, adOpenForwardOnly)
            If dblSumatoriaPesos > rsFondoActual!Total Then
                fblnDatosValidos = False
                '¡El importe de la salida de dinero en pesos no puede ser mayor que el fondo actual del corte de caja chica en dicha moneda!
                MsgBox SIHOMsg(1447), vbExclamation + vbOKOnly, "Mensaje"
                Exit Function
            End If
        End If
        
        If dblSumatoriaDolares <> 0 Then
            Set rsFondoActual = frsRegresaRs("SELECT NVL(SUM(MNYCANTIDADPAGADA),0) TOTAL FROM PVDETALLECORTE WHERE INTNUMCORTE = " & llngNumCorteValidacionImporte & " AND MNYTIPOCAMBIO <> 0", adLockOptimistic, adOpenForwardOnly)
            If dblSumatoriaDolares > rsFondoActual!Total Then
                fblnDatosValidos = False
                '¡El importe de la salida de dinero en dólares no puede ser mayor que el fondo actual del corte de caja chica en dicha moneda!
                MsgBox SIHOMsg(1448), vbExclamation + vbOKOnly, "Mensaje"
                Exit Function
            End If
        End If
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function

Private Sub cmdTop_Click()
    On Error GoTo NotificaError

    rsConsulta.MoveFirst
    pMuestra
    If Trim(rsConsulta!Estado) = "P" Then
        pHabilita 1, 1, 1, 1, 1, 1, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, IIf(Trim(rsConsulta!Estado) = "A", 1, 0)
        cmdLocate.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdTop_Click"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    llngMensajeCorteValido = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
    llngMensajeCorteValido = 0
    If llngMensajeCorteValido <> 0 Then
        'Cierre el corte actual.
        MsgBox SIHOMsg(str(llngMensajeCorteValido)), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
        Exit Sub
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = vbKeyEscape Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            If ActiveControl.Name = "txtNumero" Then
                pConsulta Val(txtNumero.Text), ldtmFecha, ldtmFecha, 0, -1, vgintNumeroDepartamento
                If rsConsulta.RecordCount <> 0 Then
                    pMuestra
                    lConsulta = True
                    If Trim(rsConsulta!Estado) = "P" Then
                        pHabilita 0, 0, 0, 0, 0, 1, 0
                        cmdSave.SetFocus
                    Else
                        pHabilita 0, 0, 1, 0, 0, 0, IIf(Trim(rsConsulta!Estado) = "A", 1, 0)
                    End If
                    If Trim(rsConsulta!Estado) = "A" Then
                        cmdCancelar.SetFocus
                    End If
                Else
                    lConsulta = False
                    If optTipo(0) = True Then optTipo_Click (0)
                    If optTipo(1) = True Then optTipo_Click (1)
                    If optTipo(0) Or optTipo(1) Then SendKeys vbTab
                    If optTipo(2) = True Then optTipo_Click (2)
                    If optTipo(2) = True Then MaskDisminucionFecha.SetFocus
                End If
            ElseIf chkIEPSBaseGravable.Enabled = True And ActiveControl.Name = "chkIEPSBaseGravable" Then
                If fblnCanFocus(cboImpuesto) Then cboImpuesto.SetFocus Else txtFleteFactura.SetFocus
            ElseIf txtIEPS.Enabled = True And ActiveControl.Name = "txtIEPS" Then
                If chkIEPSBaseGravable.Enabled And CDbl(IIf(Trim(txtIEPS.Text) = "", "0", txtIEPS.Text)) > 0 Then
                    chkIEPSBaseGravable.SetFocus
                Else
                    If cboImpuesto.Enabled = True Then cboImpuesto.SetFocus
                    If cboImpuesto.Enabled = False Then cmdSave.SetFocus
                End If
            ElseIf chkRetencionISRHonorario.Enabled = True And ActiveControl.Name = "chkRetencionISRHonorario" Then
                If cboTarifa.Enabled Then
                    cboTarifa.SetFocus
                Else
                    If chkRetencionIVAHonorario.Enabled Then
                        chkRetencionIVAHonorario.SetFocus
                    Else
                        If fraSelXMLCajaChicaHono.Visible Then
                            If optTipoComproCajaChicaHono(0).Value Then optTipoComproCajaChicaHono(0).SetFocus
                            If optTipoComproCajaChicaHono(1).Value Then optTipoComproCajaChicaHono(1).SetFocus
                            If optTipoComproCajaChicaHono(2).Value Then optTipoComproCajaChicaHono(2).SetFocus
                        Else
                            cmdSave.SetFocus
                        End If
                    End If
                End If
            ElseIf cboTarifa.Enabled = True And ActiveControl.Name = "cboTarifa" Then
                If chkRetencionIVAHonorario.Enabled Then
                    chkRetencionIVAHonorario.SetFocus
                Else
                    If fraSelXMLCajaChicaHono.Visible Then
                        If optTipoComproCajaChicaHono(0).Value Then optTipoComproCajaChicaHono(0).SetFocus
                        If optTipoComproCajaChicaHono(1).Value Then optTipoComproCajaChicaHono(1).SetFocus
                        If optTipoComproCajaChicaHono(2).Value Then optTipoComproCajaChicaHono(2).SetFocus
                    Else
                        cmdSave.SetFocus
                    End If
                End If
            ElseIf ActiveControl.Name = "cboImpuestoFlete" Then
                If optRetencion(1).Value Then
                    If fblnCanFocus(optRetencion(1)) Then optRetencion(1).SetFocus
                Else
                    If fblnCanFocus(optRetencion(0)) Then optRetencion(0).SetFocus
                End If
            ElseIf chkRetencionIVAHonorario.Enabled = True And ActiveControl.Name = "chkRetencionIVAHonorario" Then
                If fraSelXMLCajaChicaHono.Visible Then
                    If optTipoComproCajaChicaHono(0).Value Then optTipoComproCajaChicaHono(0).SetFocus
                    If optTipoComproCajaChicaHono(1).Value Then optTipoComproCajaChicaHono(1).SetFocus
                    If optTipoComproCajaChicaHono(2).Value Then optTipoComproCajaChicaHono(2).SetFocus
                Else
                    cmdSave.SetFocus
                End If
            ElseIf optTipo(0).Value = True And ActiveControl.Name = "optTipo" Then cboProveedor.SetFocus
            ElseIf optTipo(1).Value = True And ActiveControl.Name = "optTipo" Then mskCuentaHonorario.SetFocus
            ElseIf optTipo(2).Value = True And ActiveControl.Name = "optTipo" Then
                MaskDisminucionFecha.SetFocus
                pSelMkTexto MaskDisminucionFecha
            ElseIf ActiveControl.Name = "txtTotalTicket" Then cmdSave.SetFocus
            ElseIf ActiveControl.Name = "txtImporteFlete" Then cboImpuestoFlete.SetFocus
            ElseIf ActiveControl.Name = "optRetencion" Then
                If fblnCanFocus(cmdBuscarXMLFactura) Then
                    cmdBuscarXMLFactura.SetFocus
                Else
                    cmdSave.SetFocus
                End If
            ElseIf ActiveControl.Name = "cboImpuestoFleteFac" Then
                If optRetencionFactura(1).Value Then
                    If fblnCanFocus(optRetencionFactura(1)) Then optRetencionFactura(1).SetFocus
                Else
                    If fblnCanFocus(optRetencionFactura(0)) Then optRetencionFactura(0).SetFocus
                End If
            Else
                SendKeys vbTab
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon
    
    vlblnLicenciaContaElectronica = fblnLicenciaContaElectronica

    Set rsCpCajaChicaXML = frsRegresaRs("SELECT * FROM CPFACTURACAJACHICAXML WHERE INTIDFACTURA = -1", adLockOptimistic, adOpenDynamic)
    
    optTipo(0).Value = 1
    
    'Retencion de ISR
    pcargaRetencionISR
    
    'Proveedores
    pCargaProveedores
    pLlenaCombo "Select pais.INTCVEPAIS Cve, pais.VCHDESCRIPCION Nombre From Pais Where Pais.BITACTIVO = 1", cboPais
    
    'Conceptos de salida
    pCargaConceptos
    
    'Impuestos
    pCargaImpuestos
    
    'Parametros para flete
    pCargaParametrosFlete
    
    'Caja chica
    Set rsCajaChica = frsEjecuta_SP(CStr(vgintNumeroDepartamento) & "|" & vgintClaveEmpresaContable, "SP_PVSELCAJACHICA")
    
    SSTab.Tab = 0
        
    fraHonorario.Visible = False
    
    mskCuentaHonorario.Mask = vgstrEstructuraCuentaContable
    
    optMonedaHonorario(0).Value = True
    txtRFC.Enabled = False
    txtRFcHono.Enabled = False
    pLlenaCombos

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub
Private Sub pCargaParametrosFlete()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    
    'No se encuentra registrado el porcentaje de retención por fletes.
    ldblPorcentajeRetencionFletes = 0
    If gdblPorcentajeRetFletes <> 0 Then
        ldblPorcentajeRetencionFletes = gdblPorcentajeRetFletes
    End If
    
    llngCuentaRetencionFletes = 0
    'Cuenta para retención por fletes
    Set rs = frsSelParametros("CN", vgintClaveEmpresaContable, "INTCTARETENCIONFLETES")
    If Not rs.EOF Then
        If Not IsNull(rs!Valor) Then
            llngCuentaRetencionFletes = rs!Valor
        End If
    End If
    rs.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaParametrosFlete"))
End Sub

Private Sub pCargaProveedores()
    Dim strSentencia As String
    On Error GoTo NotificaError

    cboProveedorBus.Clear
    cboProveedor.Clear

    If optTipo(0).Value = True Then
        Set rs = frsEjecuta_SP("-1|-1", "SP_COSELPROVEEDOR")
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                If rs!bitactivo = 1 Then
                    cboProveedor.AddItem rs!VCHNOMBRECOMERCIAL
                    cboProveedor.ItemData(cboProveedor.newIndex) = rs!INTCVEPROVEEDOR
                End If
                cboProveedorBus.AddItem rs!VCHNOMBRECOMERCIAL
                cboProveedorBus.ItemData(cboProveedorBus.newIndex) = rs!INTCVEPROVEEDOR
                rs.MoveNext
            Loop
        End If
    Else
'|       Set rs = frsEjecuta_SP("", "SP_HOMEDICOS")
        strSentencia = "Select Distinct CoProveedor.INTCVEPROVEEDOR Clave, CoProveedor.VCHNOMBRECOMERCIAL Nombre " & _
                                "  From CoProveedor Inner Join HOMedico On (CoProveedor.VCHRFC = HOMedico.VCHRFCMEDICO )  " & _
                                " Order by Nombre"
        Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
       If rs.RecordCount <> 0 Then
            Call pLlenarCboRs(cboProveedorBus, rs, 0, 1, 0)
       End If
    End If
    
    cboProveedorBus.AddItem "<TODOS>", 0
    cboProveedorBus.ItemData(cboProveedorBus.newIndex) = -1
    cboProveedorBus.ListIndex = 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaProveedores"))
End Sub

Private Sub pMuestra()
    On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim rsTarifa As ADODB.Recordset
    Dim rsTarifaId As ADODB.Recordset
    Dim vlstrSentencia As String
    Dim curRetencionProveedor As Currency
        
    fraFactura.Enabled = False
    fraSelXMLCajaChicaFact.Enabled = False
    fraSelXMLCajaChicaHono.Enabled = False
    lblnConsulta = True
    llngProveedor = 0
    With rsConsulta
        txtNumero.Text = !IdFactura
        lblFecha.Caption = Format(!FECHAREGISTRO, "dd/mmm/yyyy")
        lblDepartamento.Caption = !nombreDepartamento
        lblPersonaRegistra.Caption = !NombreEmpleado
        
        lblEstado.ForeColor = IIf(!Estado = "C", clngColorCanceladas, IIf(!Estado = "P", clngColorPendientes, clngColorActivas))
        lblEstado.Caption = !EstadoDescripcion
        lstrEstadoFactura = Trim(!Estado)
        
        lblPersonaCancelaReembolsa.Caption = IIf(!Estado = "C", !NombrePersonaCancelo, !NombrePersonaReembolso)
        If !TipoDocumento = "D" Then optTipo(2).Value = True
        If !TipoDocumento = "H" Then optTipo(1).Value = True
        If !TipoDocumento = "F" Then optTipo(0).Value = True
        If optTipo(0).Value Then
        
            cboProveedor.ListIndex = flngLocalizaCbo(cboProveedor, CStr(!idproveedor))
            llngProveedor = CLng(!idproveedor)
            If cboProveedor.ListIndex = -1 Then cboTipoProveedor.Text = !TIPOPROVEEDOR
            If cboProveedor.ListIndex = -1 Then cboProveedor.Text = !NombreProveedor
            If cboProveedor.ListIndex = -1 Then cboPais.ListIndex = flngLocalizaCbo(cboPais, CStr(!Pais))
            If IsNull(!RFC) Then
                txtRFC.Text = ""
            Else
                txtRFC.Text = !RFC
            End If
            
            cboConcepto.ListIndex = flngLocalizaCbo(cboConcepto, CStr(!IdConcepto))
                
            If cboConcepto.ListIndex = -1 Then
                cboConcepto.AddItem !NombreConcepto
                cboConcepto.ItemData(cboConcepto.newIndex) = !IdConcepto
                cboConcepto.ListIndex = cboConcepto.newIndex
                lblnRecargarConceptos = True
                
                Set rs = frsEjecuta_SP(!IdConcepto & "|-1|" & vgintClaveEmpresaContable & "|-1", "SP_PVSELCONCEPTOCAJACHICACTAS")
                If rs.RecordCount <> 0 Then
                    ReDim Preserve arrConceptos(cboConcepto.ListCount - 1)
                    arrConceptos(cboConcepto.newIndex).lngIdConcepto = rs!intConsecutivo
                    arrConceptos(cboConcepto.newIndex).lngCtaGasto = rs!intNumeroCuenta
                    'arrConceptos(cboConcepto.newIndex).lngCtaDescuento = rs!intCuentaDescuento
                End If
            End If
            
            mskFecha.Mask = ""
            mskFecha.Text = !FechaFactura
            mskFecha.Mask = "##/##/####"
            
            txtFolio.Text = !FolioFactura
            
            optMoneda(0).Value = !Moneda = 1
            optMoneda(1).Value = !Moneda = 0
            
            optFactura.Value = !TipoDocumento = "F"
            optTicket.Value = !TipoDocumento = "T"
            optNota.Value = !TipoDocumento = "N"
            optFlete.Value = !TipoDocumento = "L"
            
            If optFactura.Value Then
                txtImporteExento.Text = FormatCurrency(IIf(IsNull(!ImporteExento), 0, !ImporteExento), 2)
                txtDescuentoExento.Text = FormatCurrency(IIf(IsNull(!DescuentoImporteExento), 0, !DescuentoImporteExento), 2)
                txtImporteNoGravado.Text = FormatCurrency(!ImporteNoGravado, 2)
                txtDescuentoNoGravado.Text = FormatCurrency(!DescuentoImporteNoGravado, 2)
                txtImporteGravado.Text = FormatCurrency(!ImporteGravado, 2)
                txtDescuentoGravado.Text = FormatCurrency(!DescuentoImporteGravado, 2)
                txtIEPS.Text = FormatCurrency(!mnyIeps, 2)
                chkIEPSBaseGravable.Value = IIf(IsNull(!BITIEPSBASEGRAVABLE), 0, !BITIEPSBASEGRAVABLE)
                txtFleteFactura.Text = FormatCurrency(!importeFlete, 2)
                               
                
                lblRetencionISR.Caption = FormatCurrency(txtImporteExento.Text - txtDescuentoExento.Text + txtImporteNoGravado.Text - txtDescuentoNoGravado.Text + txtImporteGravado.Text - txtDescuentoGravado.Text, 2)
                        
                             
                 If fcurVerifRegimen626(!idproveedor) = True Then
                     curRetencionProveedor = fcurObtenerResico(!idproveedor)
                 Else
                    curRetencionProveedor = 0
                 End If
                
                
                lblRetencionISR.Caption = FormatCurrency(lblRetencionISR.Caption * (curRetencionProveedor / 100), 2)
                If !TarifaIsr <> 0 Then
                    chkRetencionISR.Value = 1
                    cboRetencionISR.ListIndex = flngLocalizaCbo(cboRetencionISR, CStr(!TarifaIsr))
                    pCalculaTotalRetencionISR
                Else
                    cboRetencionISR.ListIndex = -1
                    cboRetencionISR.Enabled = False
                    chkRetencionISR.Value = 0
                    chkRetencionISR.Enabled = False
                End If
                If !idproveedor <> 0 Then
                     vlstrSentencia = "SELECT CnRegimenRetencion.intidtarifa, cnTarifaISR.numporcentaje " & _
                        "FROM CoProveedor INNER JOIN CnRegimenRetencion ON trim(CoProveedor.vchclaveregimensat) = trim(CnRegimenRetencion.chridregimen) " & _
                                        "INNER JOIN cnTarifaISR ON CnRegimenRetencion.intidtarifa = cnTarifaISR.intidtarifa " & _
                        "WHERE CoProveedor.intCveProveedor = " & !idproveedor
                    Set rsTarifaId = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rsTarifaId.RecordCount > 0 Then
                        Set rsTarifa = frsEjecuta_SP(rsTarifaId!intidtarifa & "|-1", "SP_CNSELTARIFAISR")
                        If rsTarifa.RecordCount > 0 Then
                            chkRetencionISR.Value = 1
                            cboRetencionISR.ListIndex = flngLocalizaCbo(cboRetencionISR, CStr(rsTarifa!IdTarifa))
                        Else
                            chkRetencionISR.Value = 0
                        End If
                    Else
                        chkRetencionISR.Value = 0
                    End If
                End If
                If !importeFlete > 0 Then
                    optRetencionFactura(0).Value = !retencionFlete > 0
                    optRetencionFactura(1).Value = !retencionFlete = 0
                    cboImpuestoFleteFac.ListIndex = flngLocalizaCbo(cboImpuestoFleteFac, CStr(!idImpuestoFlete))
                    lblRetencionFactura.Caption = FormatCurrency(!retencionFlete, 2)
                    lblImpuestoFlete.Caption = FormatCurrency(!ivaflete, 2)
                Else
                    optRetencionFactura(0).Value = False
                    optRetencionFactura(1).Value = False
                    cboImpuestoFleteFac.ListIndex = -1
                    lblRetencionFactura.Caption = FormatCurrency(0, 2)
                    lblImpuestoFlete.Caption = FormatCurrency(0, 2)
                End If
                
                If Trim(!EstadoDescripcion) = "PENDIENTE" And !ImporteGravado = 0 Then
                    cboImpuesto.ListIndex = -1
                Else
                    cboImpuesto.ListIndex = flngLocalizaCbo(cboImpuesto, CStr(!IdImpuesto))
                    lblImporteIVA.Caption = FormatCurrency(!ImporteIva, 2)
                    If cboImpuesto.ListIndex = -1 And !ImporteGravado <> 0 Then
                        cboImpuesto.AddItem !NombreImpuesto
                        cboImpuesto.ListIndex = cboImpuesto.newIndex
                        lblnRecargarImpuestos = True
                    End If
                End If
                lblTotal.Caption = FormatCurrency(!ImporteNoGravado - !DescuentoImporteNoGravado + !ImporteGravado - !DescuentoImporteGravado + !ImporteExento - !DescuentoImporteExento + !mnyIeps + !importeFlete + !ImporteIva - !retencionFlete - lblRetencionISR.Caption, 2)
            ElseIf optFlete.Value Then
                txtImporteFlete.Text = FormatCurrency(!importeFlete, 2)
                If IsNull(!retencionFlete) Then
                    optRetencion(0).Value = False
                    optRetencion(1).Value = True
                    lblImporteIvaFlete.Caption = FormatCurrency(0, 2)
                Else
                    optRetencion(0).Value = !retencionFlete > 1
                    optRetencion(1).Value = !retencionFlete = 0
                    lblImporteIvaFlete.Caption = FormatCurrency(!retencionFlete, 2)
                End If
                lblRetencion.Caption = FormatCurrency(!retencionFlete, 2)
                cboImpuestoFlete.ListIndex = flngLocalizaCbo(cboImpuestoFlete, CStr(!IdImpuesto))
                lblImporteIvaFlete.Caption = FormatCurrency(!ImporteIva, 2)
                If cboImpuestoFlete.ListIndex = -1 And !ImporteIva <> 0 Then
                    cboImpuestoFlete.AddItem !NombreImpuesto
                    cboImpuestoFlete.ListIndex = cboImpuesto.newIndex
                    lblnRecargarImpuestos = True
                End If
                lblTotalFlete.Caption = FormatCurrency(!importeFlete + !ImporteIva - !retencionFlete, 2)
            Else
                txtTotalTicket.Text = FormatCurrency(!ImporteNoGravado - !DescuentoImporteNoGravado + !ImporteGravado - !DescuentoImporteGravado + !ImporteExento - !DescuentoImporteExento + !mnyIeps + !ImporteIva + !importeFlete, 2)
            End If
            txtDescSalida.Text = ""
            If Not IsNull(!descripcionSalida) Then
                txtDescSalida.Text = Trim(!descripcionSalida)
            End If
            chkXMLrelacionadoFact.Value = IIf(Trim(!CHRTIPOCOMPROBANTE) = "" Or !CHRTIPOCOMPROBANTE = 0 Or IsNull(!CHRTIPOCOMPROBANTE), 0, 2)
            If (!Estado = "P" Or !Estado = "A") And (optFactura.Value Or optFlete.Value) Then
                fraSelXMLCajaChicaFact.Enabled = True
            Else
                fraSelXMLCajaChicaFact.Enabled = False
            End If
        ElseIf optTipo(1).Value Then
            fraHonorario.Enabled = False
            
            Set rs = frsEjecuta_SP(!NumeroCuenta, "sp_CnSelCuentaContable")
            
            
            If rs.RecordCount > 0 Then
                mskCuentaHonorario.Text = rs!vchCuentaContable
                txtCuentaHonorario.Text = Trim(rs!vchDescripcionCuenta)
            End If
                        
            rs.Close
            
            cboMedicos.ListIndex = flngLocalizaCbo(cboMedicos, CStr(!idproveedor))
            If cboMedicos.ListIndex = -1 Then
                cboMedicos.AddItem !NombreProveedor 'cboMedicos.Text = !NombreProveedor
                cboMedicos.ListIndex = cboMedicos.newIndex
            End If
            If IsNull(!RFC) Then
                txtRFcHono.Text = ""
            Else
                txtRFcHono.Text = !RFC
            End If
            
            mskFechaHonorario.Mask = ""
            mskFechaHonorario.Text = !FechaFactura
            mskFechaHonorario.Mask = "##/##/####"
            
            txtFolioHonorario.Text = !FolioFactura
            
            optMonedaHonorario(0).Value = !Moneda = 1
            optMonedaHonorario(1).Value = !Moneda = 0
            
            txtMontoHonorario.Text = FormatCurrency(!ImporteNoGravado, 2)
            lblRetencionISRHonorario.Caption = FormatCurrency(!DescuentoImporteNoGravado, 2)
            lblSubtotalHonorario.Caption = FormatCurrency(!ImporteGravado, 2)
            If IIf(IsNull(!DescuentoImporteGravado), 0, !DescuentoImporteGravado) > 0 Then chkRetencionIVAHonorario.Value = 1
            lblRetencionIVAHonorario.Caption = FormatCurrency(!DescuentoImporteGravado, 2)
            lblIVAHonorario.Caption = FormatCurrency(!ImporteIva, 2)
            
            cboIVAHonorario.ListIndex = flngLocalizaCbo(cboIVAHonorario, CStr(!IdImpuesto))
            
            If IsNull(!ImporteIva) Then
                OptIVAHonorario(0).Value = False
                OptIVAHonorario(1).Value = True
            Else
                If !ImporteIva > 0 Or IIf(IsNull(!ImporteExento), 0, !ImporteExento) = 0 Then
                    OptIVAHonorario(0).Value = True
                    OptIVAHonorario(1).Value = False
                Else
                    OptIVAHonorario(0).Value = False
                    OptIVAHonorario(1).Value = True
                End If
            End If
            
            If !TarifaIsr <> 0 Then chkRetencionISRHonorario.Value = 1 Else chkRetencionISRHonorario.Value = 0
            cboTarifa.ListIndex = flngLocalizaCbo(cboTarifa, CStr(!TarifaIsr))
            
            lblTotalPagarHonorario.Caption = FormatCurrency(!ImporteGravado - !DescuentoImporteGravado - !DescuentoImporteNoGravado, 2)
            
            If cboIVAHonorario.Text <> "" Then
                OptIVAHonorario(0).Value = True
                OptIVAHonorario(1).Value = False
            Else
                OptIVAHonorario(0).Value = False
                OptIVAHonorario(1).Value = True
            End If
            txtDescSalidaHono.Text = ""
            If Not IsNull(!descripcionSalida) Then
                txtDescSalidaHono.Text = Trim(!descripcionSalida)
            End If
            chkXMLrelacionadoHono.Value = IIf(Trim(!CHRTIPOCOMPROBANTE) = "" Or !CHRTIPOCOMPROBANTE = 0 Or IsNull(!CHRTIPOCOMPROBANTE), 0, 2)
            If (!Estado = "A") Then
                fraSelXMLCajaChicaHono.Enabled = True
            Else
                fraSelXMLCajaChicaHono.Enabled = False
            End If
        ElseIf optTipo(2).Value Then
        
            MaskDisminucionFecha.Mask = ""
            MaskDisminucionFecha.Text = !FechaFactura
            MaskDisminucionFecha.Mask = "##/##/####"
        
            OptMonedaDF(0).Value = !Moneda = 1
            OptMonedaDF(1).Value = !Moneda = 0
        
            txtTotalDF.Text = FormatCurrency(!ImporteGravado, 2)
             fraDisminucion.Enabled = False
           
        End If
        
        If optTipo(0).Value Then
            If Trim(!CHRTIPOCOMPROBANTE) = "" Or IsNull(!CHRTIPOCOMPROBANTE) Then
                optTipoComproCajaChicaFact(0).Value = False
                optTipoComproCajaChicaFact(1).Value = False
                optTipoComproCajaChicaFact(2).Value = False
            Else
                If Trim(!CHRTIPOCOMPROBANTE) = "1" Or Trim(!CHRTIPOCOMPROBANTE) = "2" Then
                    optTipoComproCajaChicaFact(0).Value = True
                Else
                    If Trim(!CHRTIPOCOMPROBANTE) = "3" Then
                        optTipoComproCajaChicaFact(1).Value = True
                    Else
                        optTipoComproCajaChicaFact(2).Value = True
                    End If
                End If
            End If
        ElseIf optTipo(1).Value Then
            If Trim(!CHRTIPOCOMPROBANTE) = "" Or IsNull(!CHRTIPOCOMPROBANTE) Then
                optTipoComproCajaChicaHono(0).Value = False
                optTipoComproCajaChicaHono(1).Value = False
                optTipoComproCajaChicaHono(2).Value = False
            Else
                If Trim(!CHRTIPOCOMPROBANTE) = "1" Or Trim(!CHRTIPOCOMPROBANTE) = "2" Then
                    optTipoComproCajaChicaHono(0).Value = True
                Else
                    If Trim(!CHRTIPOCOMPROBANTE) = "3" Then
                        optTipoComproCajaChicaHono(1).Value = True
                    Else
                        optTipoComproCajaChicaHono(2).Value = True
                    End If
                End If
            End If
        End If
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestra"))
End Sub

Private Sub pCalculaTotal()
    On Error GoTo NotificaError

    Dim dblimporteExento As Double
    Dim dbldescuentoExento As Double
    Dim dblImporteNoGravado As Double
    Dim dblDescuentoNoGravado As Double
    Dim dblimportegravado As Double
    Dim dbldescuentogravado As Double
    Dim dblImpuesto As Double
    Dim dblIEPS As Double
    Dim dblFlete As Double
    Dim dblRetencionFlete As Double
    Dim dblImpuestoFlete As Double
    Dim dblRetencionISR As Double
    
    dblimporteExento = Val(Format(txtImporteExento.Text, cstrFormato))
    dbldescuentoExento = Val(Format(txtDescuentoExento.Text, cstrFormato))
    dblImporteNoGravado = Val(Format(txtImporteNoGravado.Text, cstrFormato))
    dblDescuentoNoGravado = Val(Format(txtDescuentoNoGravado.Text, cstrFormato))
    dblimportegravado = Val(Format(txtImporteGravado.Text, cstrFormato))
    dbldescuentogravado = Val(Format(txtDescuentoGravado.Text, cstrFormato))
    dblImpuesto = Val(Format(lblImporteIVA.Caption, cstrFormato))
    dblIEPS = Val(Format(txtIEPS.Text, cstrFormato))
    dblFlete = Val(Format(txtFleteFactura.Text, cstrFormato))
    dblRetencionFlete = Val(Format(lblRetencionFactura.Caption, cstrFormato))
    dblImpuestoFlete = Val(Format(lblImpuestoFlete.Caption, cstrFormato))
    
    pCalculaTotalRetencionISR
    dblRetencionISR = Val(Format(lblRetencionISR.Caption, cstrFormato))
    
    lblTotal.Caption = FormatCurrency(dblimporteExento - dbldescuentoExento + dblImporteNoGravado - dblDescuentoNoGravado + dblimportegravado - dbldescuentogravado + dblIEPS + dblImpuesto + dblFlete + dblImpuestoFlete - dblRetencionFlete - dblRetencionISR, 2)
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculaTotal"))
End Sub

Private Sub pCalculaTotalRetencionISR()
    On Error GoTo NotificaError
    
    Dim rsTarifa As ADODB.Recordset
    Dim dblimporteExento As Double
    Dim dbldescuentoExento As Double
    Dim dblImporteNoGravado As Double
    Dim dblDescuentoNoGravado As Double
    Dim dblimportegravado As Double
    Dim dbldescuentogravado As Double
    Dim dblImpuesto As Double
    Dim dblFlete As Double
    Dim dblRetencionFlete As Double
    Dim dblImpuestoFlete As Double
    Dim curRetencionProveedor As Currency
    
    dblimporteExento = Val(Format(txtImporteExento.Text, cstrFormato))
    dbldescuentoExento = Val(Format(txtDescuentoExento.Text, cstrFormato))
    dblImporteNoGravado = Val(Format(txtImporteNoGravado.Text, cstrFormato))
    dblDescuentoNoGravado = Val(Format(txtDescuentoNoGravado.Text, cstrFormato))
    dblimportegravado = Val(Format(txtImporteGravado.Text, cstrFormato))
    dbldescuentogravado = Val(Format(txtDescuentoGravado.Text, cstrFormato))
    
    curRetencionProveedor = 0
               
      If cboProveedor.Text <> "" And cboProveedor.ListIndex = -1 Then
         If cboRetencionISR.ListIndex <> -1 Then
                curRetencionProveedor = arrRetencionISR(cboRetencionISR.ListIndex).dblPorcentaje
         End If
     End If
    
      If cboProveedor.ListIndex > -1 Then
         If fcurVerifRegimen626(cboProveedor.ItemData(cboProveedor.ListIndex)) = True Then
           curRetencionProveedor = fcurObtenerResico(cboProveedor.ItemData(cboProveedor.ListIndex))
        Else
          curRetencionProveedor = 0
        End If
      End If
           
      If txtFleteFactura.Text <> "" Then
         dblImpuestoFlete = Val(Format(txtFleteFactura.Text, cstrFormato))
      Else
         dblImpuestoFlete = 0
      End If
     
     If curRetencionProveedor <> 0 Then
          If curRetencionProveedor = "1.25" Then
            lblRetencionISR.Caption = FormatCurrency(dblimporteExento - dbldescuentoExento + dblImporteNoGravado - dblDescuentoNoGravado + dblimportegravado - dbldescuentogravado + dblImpuestoFlete)
          Else
            lblRetencionISR.Caption = FormatCurrency(dblimporteExento - dbldescuentoExento + dblImporteNoGravado - dblDescuentoNoGravado + dblimportegravado - dbldescuentogravado)
          End If
     Else
       lblRetencionISR.Caption = FormatCurrency(0)
     End If
    
     If cboRetencionISR.ListIndex <> -1 Then
        lblRetencionISR.Caption = FormatCurrency(lblRetencionISR.Caption * (arrRetencionISR(cboRetencionISR.ListIndex).dblPorcentaje / 100), 2)
     Else
        lblRetencionISR.Caption = FormatCurrency(0, 2)
     End If
     
    If lblRetencionISR = 0 Then
        If cboProveedor.ListIndex = -1 And cboRetencionISR.ListIndex = -1 Then
            chkRetencionISR.Value = 0
            cboRetencionISR.Enabled = False
            'cboRetencionISR.ListIndex = -1
        End If
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculaTotalRetencionISR"))
End Sub
Private Sub pCargaImpuestos()
    On Error GoTo NotificaError

    Dim intcontador As Integer
    cboImpuesto.Clear
    cboImpuestoFlete.Clear
    cboImpuestoFleteFac.Clear
    Set rs = frsEjecuta_SP("-1|1|1", "sp_GnSelImpuesto")

    intcontador = 0
    Do While Not rs.EOF
        ReDim Preserve arrImpuestosFactura(intcontador)
        cboImpuesto.AddItem rs!VCHDESCRIPCION
        cboImpuesto.ItemData(cboImpuesto.newIndex) = rs!smiCveImpuesto
        cboImpuestoFlete.AddItem rs!VCHDESCRIPCION
        cboImpuestoFlete.ItemData(cboImpuestoFlete.newIndex) = rs!smiCveImpuesto
        cboImpuestoFleteFac.AddItem rs!VCHDESCRIPCION
        cboImpuestoFleteFac.ItemData(cboImpuestoFlete.newIndex) = rs!smiCveImpuesto
        arrImpuestosFactura(intcontador).lngIdImpuesto = rs!smiCveImpuesto
        arrImpuestosFactura(intcontador).dblPorcentaje = rs!relPorcentaje
        intcontador = intcontador + 1
        rs.MoveNext
    Loop
    cboImpuestoFlete.AddItem "<NINGUNO>", 0
    cboImpuestoFlete.ListIndex = flngLocalizaCbo(cboImpuestoFlete, CStr(glngCveImpuesto))
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaImpuestos"))
End Sub

Private Sub pCargaConceptos()
    On Error GoTo NotificaError
    Dim intcontador As Integer
    
    cboConcepto.Clear
    
    Set rs = frsEjecuta_SP("-1|1|" & vgintClaveEmpresaContable & "|" & vgintNumeroDepartamento, "SP_PVSELCONCEPTOCAJACHICACTAS")
    intcontador = 0
    
    Do While Not rs.EOF
        ReDim Preserve arrConceptos(intcontador)
        cboConcepto.AddItem rs!VCHDESCRIPCION
        cboConcepto.ItemData(cboConcepto.newIndex) = rs!intConsecutivo
        arrConceptos(intcontador).lngIdConcepto = rs!intConsecutivo
        arrConceptos(intcontador).lngCtaGasto = rs!intNumeroCuenta
        'arrConceptos(intcontador).lngCtaDescuento = rs!intCuentaDescuento
        intcontador = intcontador + 1
        rs.MoveNext
    Loop

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaConceptos"))
End Sub
Private Sub pcargaRetencionISR()
    On Error GoTo NotificaError
    Dim intcontador As Integer
    
    cboRetencionISR.Clear
    cboRetencionISR.Enabled = False
    intcontador = 0
    vgstrParametrosSP = "-1|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CNSELTARIFAISR")
    If rs.RecordCount <> 0 Then
        Call pLlenarCboRs(cboRetencionISR, rs, 0, 1, 0)
        cboRetencionISR.ListIndex = -1
    End If
    Do While Not rs.EOF
        ReDim Preserve arrRetencionISR(intcontador)
            arrRetencionISR(intcontador).lngIdImpuesto = rs!IdTarifa
            arrRetencionISR(intcontador).dblPorcentaje = rs!Porcentaje
        intcontador = intcontador + 1
        rs.MoveNext
    Loop
    
    rs.Close

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaConceptos"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    If SSTab.Tab = 1 Then
        Cancel = 1
        SSTab.Tab = 0
        txtNumero.SetFocus
    Else
        If (cmdSave.Enabled Or cmdCancelar.Enabled Or lblnConsulta) And llngMensajeCorteValido = 0 Then
            Cancel = 1
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                optTipo(0).Value = True
                optTipo_Click 0
                pLimpia
                pHabilita 0, 0, 1, 0, 0, 0, 0
                txtNumero.SetFocus
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub
Private Sub grdFacturas_DblClick()
    On Error GoTo NotificaError
    Dim lblnTermina As Boolean

    If Val(grdFacturas.TextMatrix(grdFacturas.Row, cintColNumero)) <> 0 Then
        lblnTermina = False
        rsConsulta.MoveFirst
        Do While Not rsConsulta.EOF And Not lblnTermina
            If Val(grdFacturas.TextMatrix(grdFacturas.Row, cintColNumero)) = rsConsulta!IdFactura Then
                lblnTermina = True
            Else
                rsConsulta.MoveNext
            End If
        Loop
        If lblnTermina Then
            pMuestra
            lConsulta = True
            If Trim(rsConsulta!Estado) = "P" Then
                pHabilita 1, 1, 1, 1, 1, 1, 0
                cmdSave.SetFocus
            Else
                pHabilita 1, 1, 1, 1, 1, 0, IIf(Trim(rsConsulta!Estado) = "A", 1, 0)
                cmdLocate.SetFocus
            End If
            SSTab.Tab = 0
        Else
            'La información ha cambiado, consulte de nuevo.
            MsgBox SIHOMsg(381), vbOKOnly + vbExclamation, "Mensaje"
            SSTab.Tab = 0
            txtNumero.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdFacturas_DblClick"))
End Sub

Private Sub grdFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        grdFacturas_DblClick
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdFacturas_KeyDown"))
End Sub


Private Sub lblImporteIVA_Change()
    On Error GoTo NotificaError

    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lblImporteIVA_Change"))
End Sub

Private Sub MaskDisminucionFecha_GotFocus()
    pSelMkTexto MaskDisminucionFecha
End Sub

Private Sub MaskDisminucionFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And lConsulta = False Then
        pLimpiaDF
        pHabilita 0, 0, 0, 0, 0, 1, 0
    End If
End Sub

Private Sub mskFecha_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskFecha
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecha_GotFocus"))
End Sub

Private Sub mskFecha_LostFocus()
    On Error GoTo NotificaError

    If Trim(mskFecha.ClipText) = "" Then
        mskFecha.Mask = ""
        mskFecha.Text = ldtmFecha
        mskFecha.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecha_LostFocus"))
End Sub

Private Sub mskFechaBusFin_Change()
    On Error GoTo NotificaError

    cmdCargar.Enabled = IsDate(mskFechaBusIni)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaBusFin_Change"))
End Sub

Private Sub mskFechaBusFin_GotFocus()
    On Error GoTo NotificaError

    pSelMkTexto mskFechaBusFin
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaBusFin_GotFocus"))
End Sub

Private Sub mskFechaBusFin_LostFocus()
    On Error GoTo NotificaError

    If Not IsDate(mskFechaBusFin.Text) Then
        mskFechaBusFin.Mask = ""
        mskFechaBusFin.Text = ldtmFecha
        mskFechaBusFin.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaBusFin_LostFocus"))
End Sub

Private Sub mskFechaBusIni_Change()
    On Error GoTo NotificaError
    cmdCargar.Enabled = IsDate(mskFechaBusIni)
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaBusIni_Change"))
End Sub

Private Sub mskFechaBusIni_GotFocus()
    On Error GoTo NotificaError
        pSelMkTexto mskFechaBusIni
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaBusIni_GotFocus"))
End Sub

Private Sub mskFechaBusIni_LostFocus()
    On Error GoTo NotificaError
    If Not IsDate(mskFechaBusIni.Text) Then
        mskFechaBusIni.Mask = ""
        mskFechaBusIni.Text = ldtmFecha
        mskFechaBusIni.Mask = "##/##/####"
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaBusIni_LostFocus"))
End Sub

Private Sub mskFechaHonorario_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelMkTexto mskFechaHonorario
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaHonorario_GotFocus"))
End Sub

Private Sub mskFechaHonorario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And mskFechaHonorario.Text = "  /  /    " Then
        mskFechaHonorario.Text = fdtmServerFecha
    End If
End Sub

Private Sub mskFechaHonorario_LostFocus()
    On Error GoTo NotificaError

    If Trim(mskFechaHonorario.ClipText) = "" Then
        mskFechaHonorario.Mask = ""
        mskFechaHonorario.Text = ldtmFecha
        mskFechaHonorario.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaHonorario_LostFocus"))
End Sub

Private Sub optFactura_Click()
    pHabilitaBusquedaXML
    pMuestraTipoDocumento
End Sub

Private Sub optFlete_Click()
    pHabilitaBusquedaXML
    pMuestraTipoDocumento
End Sub

Private Sub OptIVAHonorario_Click(Index As Integer)
    On Error GoTo NotificaError

    If OptIVAHonorario(0).Value = True Then
        cboIVAHonorario.Enabled = True
        chkRetencionIVAHonorario.Enabled = True
    End If
    
    If OptIVAHonorario(1).Value = True Then
        cboIVAHonorario.Enabled = False
        chkRetencionIVAHonorario.Enabled = False
    End If
    
    If OptIVAHonorario(0).Value = True Then
        If cboIVAHonorario.ListCount <> 0 Then
            If lblnConsulta = False Then cboIVAHonorario.ListIndex = 0
        End If
    End If
    
    If OptIVAHonorario(1).Value = True Then
        cboIVAHonorario.ListIndex = -1
        
        chkRetencionIVAHonorario.Value = 0
        lblRetencionIVAHonorario.Caption = FormatCurrency(0, 2)
    End If
    
    pCalculaSubtotalHonorario
    pCalculaTotalHonorario
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optIVAHonorario_Click"))
End Sub

Private Sub OptMonedaDF_GotFocus(Index As Integer)
    If Index = 0 Then OptMonedaDF(0).Value = True
    If Index = 1 Then OptMonedaDF(1).Value = True
End Sub

Private Sub OptMonedaDF_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or vbKeyTab Then
    pSelTextBox txtTotalDF
End If
End Sub

Private Sub optNota_Click()
    pHabilitaBusquedaXML
    pMuestraTipoDocumento
End Sub

Private Sub optRetencion_Click(Index As Integer)
    If Not lblnConsulta Then
        pCalculaTotalFlete
        pHabilita 0, 0, 0, 0, 0, 1, 0
    End If
End Sub

Private Sub optRetencionFactura_Click(Index As Integer)

    lblRetencionFactura = FormatCurrency(IIf(optRetencionFactura(0).Value, (Val(Format(txtFleteFactura.Text, cstrFormato)) * ldblPorcentajeRetencionFletes), 0), 2)
    pCalculaTotal
    
End Sub

Private Sub optTicket_Click()
    pHabilitaBusquedaXML
    pMuestraTipoDocumento
End Sub
Private Sub pMuestraTipoDocumento()
    fraDocFactura.Visible = optFactura.Value
    fraDocTicketNota.Visible = optTicket.Value Or optNota.Value
    fraDocFlete.Visible = optFlete.Value
    'notas y tickets
    txtTotalTicket.Text = FormatCurrency("0", 2)
    'flete
    txtImporteFlete.Text = FormatCurrency("0", 2)
    lblImporteIvaFlete.Caption = FormatCurrency("0", 2)
    cboImpuestoFlete.ListIndex = flngLocalizaCbo(cboImpuestoFlete, CStr(glngCveImpuesto))
    lblTotalFlete.Caption = FormatCurrency("0", 2)
    lblRetencion.Caption = FormatCurrency("0", 2)
    'facturas
    txtImporteExento.Text = FormatCurrency("0", 2)
    txtDescuentoExento.Text = FormatCurrency("0", 2)
    txtImporteNoGravado.Text = FormatCurrency("0", 2)
    txtDescuentoNoGravado.Text = FormatCurrency("0", 2)
    txtImporteGravado.Text = FormatCurrency("0", 2)
    txtDescuentoGravado.Text = FormatCurrency("0", 2)
    chkIEPSBaseGravable.Value = vbUnchecked
    txtIEPS.Text = FormatCurrency("0", 2)
    cboImpuesto.ListIndex = -1
    lblImporteIVA.Caption = FormatCurrency("0", 2)
    lblTotal.Caption = FormatCurrency("0", 2)
End Sub
Private Sub optTipo_Click(Index As Integer)
    pCargaProveedores
    
    If optTipo(0) = True Then
        fraFactura.Visible = True
        fraSelXMLCajaChicaFact.Visible = True
        fraSelXMLCajaChicaHono.Visible = False
        fraDisminucion.Visible = False
        fraHonorario.Visible = False
        frmCajaChica.Height = 8500 '7035, Tenia 8130
        pLimpia
        txtNumero.ToolTipText = "Número de factura"
        lblTituloCancela.Caption = "Canceló / Reembolsó"
        FraBotonera.Top = 7650 '6120, tenia 7245
        fraSelXMLCajaChicaHono.Top = 4680
    ElseIf optTipo(1) = True Then
        fraHonorario.Visible = True
        fraHonorario.Height = 3450
        fraFactura.Visible = False
        fraSelXMLCajaChicaFact.Visible = False
        fraSelXMLCajaChicaHono.Visible = True
        fraDisminucion.Visible = False
        frmCajaChica.Height = 6420
        pLimpiaHonorarios
        txtNumero.ToolTipText = "Número de honorario"
        lblTituloCancela.Caption = "Canceló / Reembolsó"
        FraBotonera.Top = 5470 'IIf(vlblnLicenciaContaElectronica = False, 5470, 6120)
    ElseIf optTipo(2) = True Then
        fraDisminucion.Visible = True
        fraSelXMLCajaChicaFact.Visible = False
        fraSelXMLCajaChicaHono.Visible = False
        fraFactura.Visible = False
        fraHonorario.Visible = False
        fraSelXMLCajaChicaFact.Visible = False
        frmCajaChica.Height = 5200
        txtNumero.ToolTipText = "Número del documento"
        pLimpiaDF
        FraBotonera.Top = 4350
        lblTituloCancela.Caption = "Canceló / Depositó"
        pHabilita 0, 0, 1, 0, 0, 0, 0
    End If
End Sub


Private Sub txtDescSalida_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescSalida
End Sub

Private Sub txtDescSalida_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDescSalidaHono_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescSalidaHono
End Sub

Private Sub txtDescSalidaHono_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDescuentoGravado_Change()
    On Error GoTo NotificaError

    If cboImpuesto.ListIndex <> -1 Then
        cboImpuesto_Click
    End If
    
    If chkRetencionISR.Value Then
        pCalculaTotalRetencionISR
    End If
    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoGravado_Change"))
End Sub

Private Sub txtDescuentoGravado_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescuentoGravado

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoGravado_GotFocus"))
End Sub

Private Sub txtDescuentoGravado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtDescuentoGravado)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoGravado_KeyPress"))
End Sub

Private Sub txtDescuentoGravado_LostFocus()
    On Error GoTo NotificaError

    txtDescuentoGravado.Text = FormatCurrency(Val(Format(txtDescuentoGravado.Text, cstrFormato)), 2)
    If Val(Format(txtDescuentoGravado.Text, cstrFormato)) > Val(Format(txtImporteGravado.Text, cstrFormato)) Then
        MsgBox SIHOMsg(925), vbCritical, "Mensaje"
        txtDescuentoGravado.Text = FormatCurrency(0, 2)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoGravado_LostFocus"))
End Sub

Private Sub txtDescuentoNoGravado_Change()
    On Error GoTo NotificaError
    
    If chkRetencionISR.Value Then
        pCalculaTotalRetencionISR
    End If
    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoNoGravado_Change"))
End Sub

Private Sub txtDescuentoExento_Change()
    On Error GoTo NotificaError
    
    If chkRetencionISR.Value Then
        pCalculaTotalRetencionISR
    End If
    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoExento_Change"))
End Sub

Private Sub txtDescuentoNoGravado_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescuentoNoGravado

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoNoGravado_GotFocus"))
End Sub

Private Sub txtDescuentoExento_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescuentoExento

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoExento_GotFocus"))
End Sub

Private Sub txtDescuentoNoGravado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtDescuentoNoGravado)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoNoGravado_KeyPress"))
End Sub

Private Sub txtDescuentoExento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtDescuentoExento)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoExento_KeyPress"))
End Sub

Private Sub txtDescuentoNoGravado_LostFocus()
    On Error GoTo NotificaError
    txtDescuentoNoGravado.Text = FormatCurrency(Val(Format(txtDescuentoNoGravado.Text, cstrFormato)), 2)
    If Val(Format(txtDescuentoNoGravado.Text, cstrFormato)) > Val(Format(txtImporteNoGravado.Text, cstrFormato)) Then
        MsgBox SIHOMsg(925), vbCritical, "Mensaje"
        txtDescuentoNoGravado.Text = FormatCurrency(0, 2)
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoNoGravado_LostFocus"))
End Sub

Private Sub txtDescuentoExento_LostFocus()
    On Error GoTo NotificaError
    txtDescuentoExento.Text = FormatCurrency(Val(Format(txtDescuentoExento.Text, cstrFormato)), 2)
    If Val(Format(txtDescuentoExento.Text, cstrFormato)) > Val(Format(txtImporteExento.Text, cstrFormato)) Then
        MsgBox SIHOMsg(925), vbCritical, "Mensaje"
        txtDescuentoExento.Text = FormatCurrency(0, 2)
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescuentoExento_LostFocus"))
End Sub

Private Sub txtFleteFactura_Change()
On Error GoTo NotificaError

    cboImpuestoFleteFac.Enabled = Val(Format(txtFleteFactura.Text, cstrFormato)) <> 0
    lblRetencionSiNo.Enabled = Val(Format(txtFleteFactura.Text, cstrFormato)) <> 0
    optRetencionFactura(0).Enabled = Val(Format(txtFleteFactura.Text, cstrFormato)) <> 0
    optRetencionFactura(1).Enabled = Val(Format(txtFleteFactura.Text, cstrFormato)) <> 0
    lblRetencionFactura.Enabled = Val(Format(txtFleteFactura.Text, cstrFormato)) <> 0
    lblImpuestoFlete.Enabled = Val(Format(txtFleteFactura.Text, cstrFormato)) <> 0
    lblTituloImpuestoFlete.Enabled = Val(Format(txtFleteFactura.Text, cstrFormato)) <> 0
    
    If Not lblRetencionFactura.Enabled Then
        lblRetencionFactura.Caption = FormatCurrency("0", 2)
        lblImpuestoFlete.Caption = FormatCurrency("0", 2)
        optRetencionFactura(0).Value = False
        optRetencionFactura(1).Value = False
        cboImpuestoFleteFac.ListIndex = -1
    Else
        cboImpuestoFleteFac.ListIndex = flngLocalizaCbo(cboImpuestoFleteFac, CStr(glngCveImpuesto))
    End If
    
    cboImpuestoFleteFac_Click
    'pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFleteFactura_Change"))
End Sub

Private Sub txtFleteFactura_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtFleteFactura
End Sub

Private Sub txtFleteFactura_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtFleteFactura)) Or Not fblnFormatoCantidad(txtFleteFactura, KeyAscii, 2) Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtFleteFactura_LostFocus()
    On Error GoTo NotificaError

    txtFleteFactura.Text = FormatCurrency(Val(Format(txtFleteFactura.Text, cstrFormato)), 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFleteFactura_LostFocus"))
End Sub

Private Sub TxtFolio_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtFolio

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolio_GotFocus"))
End Sub

Private Sub TxtFolio_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolio_KeyPress"))
End Sub

Private Sub txtFolioHonorario_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtFolioHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioHonorario_GotFocus"))
End Sub

Private Sub txtIEPS_Change()
On Error GoTo NotificaError

    If Not lblTituloImpuesto.Enabled Then
        lblImporteIVA.Caption = FormatCurrency("0", 2)
        cboImpuesto.ListIndex = -1
    Else
        If cboImpuesto.ListIndex <> -1 Then
            cboImpuesto_Click
        End If
    End If
    If chkRetencionISR.Value Then
        pCalculaTotalRetencionISR
    End If
    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIEPS_Change"))

End Sub

Private Sub txtIEPS_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtIEPS

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIEPS_GotFocus"))

End Sub

Private Sub txtIEPS_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtIEPS)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIEPS_KeyPress"))
End Sub

Private Sub txtIEPS_LostFocus()

On Error GoTo NotificaError
    
        txtIEPS.Text = FormatCurrency(Val(Format(txtIEPS.Text, cstrFormato)), 2)

        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIEPS_LostFocus"))

End Sub

Private Sub txtImporteFlete_Change()
    On Error GoTo NotificaError

    cboImpuestoFlete_Click
    pCalculaTotalFlete

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteFlete_Change"))
End Sub
Private Sub pCalculaTotalFlete()
    On Error GoTo NotificaError

    Dim dblImporte As Double
    Dim dblImpuesto As Double
    Dim dblRetFlete As Double
    
    dblImporte = Val(Format(txtImporteFlete.Text, cstrFormato))
    dblImpuesto = Val(Format(lblImporteIvaFlete.Caption, cstrFormato))
    dblRetFlete = IIf(optRetencion(0).Value, (dblImporte * ldblPorcentajeRetencionFletes), 0)
    lblRetencion.Caption = FormatCurrency(dblRetFlete, 2)
    lblTotalFlete.Caption = FormatCurrency(dblImporte + dblImpuesto - dblRetFlete, 2)
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculaTotalFlete"))
End Sub

Private Sub txtImporteFlete_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtImporteFlete

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteFlete_GotFocus"))
End Sub

Private Sub txtImporteFlete_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtImporteFlete)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteFlete_KeyPress"))
End Sub

Private Sub txtImporteFlete_LostFocus()
On Error GoTo NotificaError

    txtImporteFlete.Text = FormatCurrency(Val(Format(txtImporteFlete.Text, cstrFormato)), 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteFlete_LostFocus"))
End Sub

Private Sub txtImporteGravado_Change()
    On Error GoTo NotificaError

    lblDescuentoGravado.Enabled = Val(Format(txtImporteGravado.Text, cstrFormato)) <> 0
    txtDescuentoGravado.Enabled = Val(Format(txtImporteGravado.Text, cstrFormato)) <> 0
    If Not lblDescuentoGravado.Enabled Then
        txtDescuentoGravado.Text = FormatCurrency("0", 2)
    End If

    lblTituloImpuesto.Enabled = Val(Format(txtImporteGravado.Text, cstrFormato)) <> 0
    cboImpuesto.Enabled = Val(Format(txtImporteGravado.Text, cstrFormato)) <> 0
    If Not lblTituloImpuesto.Enabled Then
        lblImporteIVA.Caption = FormatCurrency("0", 2)
        cboImpuesto.ListIndex = -1
    Else
        If cboImpuesto.ListIndex <> -1 Then
            cboImpuesto_Click
        End If
    End If
    
    
     If Val(Format(txtImporteNoGravado.Text, cstrFormato)) = 0 And Val(Format(txtImporteGravado.Text, cstrFormato)) = 0 And Val(Format(txtImporteExento.Text, cstrFormato)) = 0 Then
       cboRetencionISR.ListIndex = -1
    End If
'
'    If chkRetencionISR.Value Then
'        pCalculaTotalRetencionISR
'    End If
    
    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteGravado_Change"))
End Sub

Private Sub txtImporteGravado_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtImporteGravado

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteGravado_GotFocus"))
End Sub

Private Sub txtImporteGravado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtImporteGravado)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteGravado_KeyPress"))
End Sub

Private Sub txtImporteGravado_LostFocus()
    On Error GoTo NotificaError

    txtImporteGravado.Text = FormatCurrency(Val(Format(txtImporteGravado.Text, cstrFormato)), 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteGravado_LostFocus"))
End Sub

Private Sub txtImporteNoGravado_Change()
    On Error GoTo NotificaError

    lblDescuentoNoGravado.Enabled = Val(Format(txtImporteNoGravado.Text, cstrFormato)) <> 0
    txtDescuentoNoGravado.Enabled = Val(Format(txtImporteNoGravado.Text, cstrFormato)) <> 0
    
    If Not lblDescuentoNoGravado.Enabled Then
        txtDescuentoNoGravado.Text = FormatCurrency("0", 2)
        cboRetencionISR.Enabled = True
    End If
    
    If Val(Format(txtImporteNoGravado.Text, cstrFormato)) = 0 And Val(Format(txtImporteGravado.Text, cstrFormato)) = 0 And Val(Format(txtImporteExento.Text, cstrFormato)) = 0 Then
       cboRetencionISR.ListIndex = -1
    End If
    
'    If chkRetencionISR.Value Then
'        pCalculaTotalRetencionISR
'    End If
    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteNoGravado_Change"))
End Sub

Private Sub txtImporteExento_Change()
    On Error GoTo NotificaError

    lblDescuentoExento.Enabled = Val(Format(txtImporteExento.Text, cstrFormato)) <> 0
    txtDescuentoExento.Enabled = Val(Format(txtImporteExento.Text, cstrFormato)) <> 0
    If Not lblDescuentoExento.Enabled Then
        txtDescuentoExento.Text = FormatCurrency("0", 2)
    End If
        
    If Val(Format(txtImporteNoGravado.Text, cstrFormato)) = 0 And Val(Format(txtImporteGravado.Text, cstrFormato)) = 0 And Val(Format(txtImporteExento.Text, cstrFormato)) = 0 Then
       cboRetencionISR.ListIndex = -1
    End If
    
'    If chkRetencionISR.Value Then
'        pCalculaTotalRetencionISR
'    End If
    pCalculaTotal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteExento_Change"))
End Sub

Private Sub txtImporteNoGravado_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtImporteNoGravado

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteNoGravado_GotFocus"))
End Sub

Private Sub txtImporteExento_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtImporteExento

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteExento_GotFocus"))
End Sub

Private Sub txtImporteNoGravado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtImporteNoGravado)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteNoGravado_KeyPress"))
End Sub

Private Sub txtImporteExento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtImporteExento)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteExento_KeyPress"))
End Sub

Private Sub txtImporteNoGravado_LostFocus()
    On Error GoTo NotificaError

    txtImporteNoGravado.Text = FormatCurrency(Val(Format(txtImporteNoGravado.Text, cstrFormato)), 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteNoGravado_LostFocus"))
End Sub

Private Sub txtImporteExento_LostFocus()
    On Error GoTo NotificaError

    txtImporteExento.Text = FormatCurrency(Val(Format(txtImporteExento.Text, cstrFormato)), 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteExento_LostFocus"))
End Sub

Private Sub txtMontoHonorario_LostFocus()
    On Error GoTo NotificaError

    txtMontoHonorario.Text = FormatCurrency(Val(Format(txtMontoHonorario.Text, cstrFormato)), 2)
    If OptIVAHonorario(0).Value = False And OptIVAHonorario(1).Value = False And OptIVAHonorario(0).Enabled = True Then
       OptIVAHonorario(0).Value = True
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_LostFocus"))
End Sub

Private Sub txtNumero_GotFocus()
    On Error GoTo NotificaError

    If optTipo(0).Value = True Then
        pLimpia
        pHabilita 0, 0, 1, 0, 0, 0, 0
    ElseIf optTipo(1).Value = True Then
        pLimpiaHonorarios
        pHabilita 0, 0, 1, 0, 0, 0, 0
    ElseIf optTipo(2).Value = True Then
        optTipo_Click (0)
        fraDisminucion.Enabled = True
        pHabilita 0, 0, 1, 0, 0, 0, 0
    End If
       
    pSelTextBox txtNumero

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumero_GotFocus"))
End Sub

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intCancel As Integer)
    On Error GoTo NotificaError

    cmdTop.Enabled = intTop = 1
    cmdBack.Enabled = intBack = 1
    cmdLocate.Enabled = intlocate = 1
    cmdNext.Enabled = intNext = 1
    cmdEnd.Enabled = intEnd = 1
    cmdSave.Enabled = intSave = 1
    cmdCancelar.Enabled = intCancel = 1

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub
Private Sub pLimpiaDF()
On Error GoTo NotificaError

    vlintTipoXMLCXP = 0
    vlstrUUIDXMLCXP = ""
    vlstrRFCXMLCXP = ""
    vldblMontoXMLCXP = 0
    vlstrMonedaXMLCXP = ""
    vldblTipoCambioXMLCXP = 0
    vlstrSerieCXP = ""
    vlstrNumFolioCXP = ""
    vlstrNumFactExtCXP = ""
    vlstrTaxIDExtCXP = ""
    vlstrXMLCXP = ""
    
    txtNumero.Text = frsRegresaRs("SELECT NVL(MAX(INTIDFACTURA),0)+1 FROM CPFACTURACAJACHICA").Fields(0)
    lblFecha.Caption = Format(fdtmServerFecha, "dd/mmm/yyyy")
    lblDepartamento.Caption = vgstrNombreDepartamento
    
    lblPersonaRegistra.Caption = ""
    lblEstado.Caption = ""
    lblPersonaCancelaReembolsa.Caption = ""
    
    OptMonedaDF(0).Value = True
    
    ldtmFecha = fdtmServerFecha
    MaskDisminucionFecha.Mask = ""
    MaskDisminucionFecha.Text = ldtmFecha
    MaskDisminucionFecha.Mask = "##/##/####"
    
    txtTotalDF = FormatCurrency(0)
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    fraDisminucion.Enabled = True
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaDF"))
End Sub
Private Sub pLimpia()
    On Error GoTo NotificaError

    lConsulta = False

    vlintTipoXMLCXP = 0
    vlstrUUIDXMLCXP = ""
    vlstrRFCXMLCXP = ""
    vldblMontoXMLCXP = 0
    vlstrMonedaXMLCXP = ""
    vldblTipoCambioXMLCXP = 0
    vlstrSerieCXP = ""
    vlstrNumFolioCXP = ""
    vlstrNumFactExtCXP = ""
    vlstrTaxIDExtCXP = ""
    vlstrXMLCXP = ""
    
    chkXMLrelacionadoFact.Value = 0
    chkXMLrelacionadoHono.Value = 0

    fraFactura.Enabled = True
    fraSelXMLCajaChicaFact.Enabled = True
    
    lblnConsulta = False
    ldtmFecha = fdtmServerFecha
    
    txtNumero.Text = frsRegresaRs("SELECT NVL(MAX(INTIDFACTURA),0)+1 FROM CPFACTURACAJACHICA").Fields(0)

    lblFecha.Caption = Format(fdtmServerFecha, "dd/mmm/yyyy")
    lblDepartamento.Caption = vgstrNombreDepartamento
    
    lblPersonaRegistra.Caption = ""
    lblEstado.Caption = ""
    lblPersonaCancelaReembolsa.Caption = ""
    
    cboProveedor.ListIndex = -1
    cboProveedor.Text = ""
    cboConcepto.ListIndex = -1
    cboRetencionISR.ListIndex = -1
    txtRFC.Text = ""
    txtRFC.Enabled = False
    cboTipoProveedor.ListIndex = 0
    cboTipoProveedor.Enabled = True
    cboPais.ListIndex = fintLocalizaCbo(cboPais, CStr(vgintCvePaisCH))
    chkRetencionISR.Enabled = True
    chkRetencionISR.Value = 0
    cboRetencionISR.ListIndex = -1
    cboRetencionISR.Enabled = False
    cboPais.Enabled = True
    mskFecha.Mask = ""
    mskFecha.Text = ""
    mskFecha.Mask = "##/##/####"
    
    txtFolio.Text = ""
    
    optMoneda(0).Value = True
    optFactura.Value = True
    
    txtImporteExento.Text = FormatCurrency("0", 2)
    txtDescuentoExento.Text = FormatCurrency("0", 2)
    txtImporteNoGravado.Text = FormatCurrency("0", 2)
    txtDescuentoNoGravado.Text = FormatCurrency("0", 2)
    txtImporteGravado.Text = FormatCurrency("0", 2)
    txtDescuentoGravado.Text = FormatCurrency("0", 2)
    txtIEPS.Text = FormatCurrency("0", 2)
    chkIEPSBaseGravable.Value = False
    txtFleteFactura.Text = FormatCurrency("0", 2)
    cboImpuestoFlete.ListIndex = -1
    lblRetencionFactura.Caption = FormatCurrency("0", 2)
    lblImpuestoFlete.Caption = FormatCurrency("0", 2)
    
    optRetencion(0).Value = False
    optRetencion(1).Value = False
    optRetencionFactura(0).Value = False
    optRetencionFactura(1).Value = False
    
    cboImpuesto.ListIndex = -1
    lblImporteIVA.Caption = FormatCurrency("0", 2)
    lblTotal.Caption = FormatCurrency("0", 2)
    lblRetencionISR = FormatCurrency("0", 2)
    
    cboImpuestoFleteFac.Enabled = False
    lblTituloImpuestoFlete.Enabled = False
    cboImpuestoFleteFac.ListIndex = -1
    lblRetencionSiNo.Enabled = False
    optRetencionFactura(0).Enabled = False
    optRetencionFactura(1).Enabled = False
    lblRetencionFactura.Enabled = False
    lblImpuestoFlete.Enabled = False
    
    If lblnRecargarProveedores Then
        pCargaProveedores
        lblnRecargarProveedores = False
    End If
    If lblnRecargarConceptos Then
        pCargaConceptos
        lblnRecargarConceptos = False
    End If
    If lblnRecargarImpuestos Then
        pCargaImpuestos
        lblnRecargarImpuestos = False
    End If
    
    'Controles de la consulta
    mskFechaBusIni.Mask = ""
    mskFechaBusIni.Text = ldtmFecha
    mskFechaBusIni.Mask = "##/##/####"

    mskFechaBusFin.Mask = ""
    mskFechaBusFin.Text = ldtmFecha
    mskFechaBusFin.Mask = "##/##/####"

    cboProveedorBus.ListIndex = 0

    cmdCancelar.Enabled = False

    pConfiguraBusqueda
    cboTipoProveedor.ListIndex = 2
    pHabilitaBusquedaXML
    txtDescSalida.Text = ""
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Private Sub pConfiguraBusqueda()
    On Error GoTo NotificaError

    Dim intcontador As Integer

    With grdFacturas
        .Cols = cintColumnas
        .Rows = 2
        .FormatString = cstrTitulos
        .FixedCols = 1
        .FixedRows = 1
        
        For intcontador = 1 To .Cols - 1
            .TextMatrix(.Rows - 1, intcontador) = ""
        Next intcontador
        
        .ColWidth(0) = 100
        .ColWidth(cintColFecha) = 1100
        .ColWidth(cintColNumero) = 800
        .ColWidth(cintColProveedor) = IIf(optTipo(2).Value = True, 0, 2800)
        .ColWidth(cintColFactura) = IIf(optTipo(2).Value = True, 0, 1000)
        .ColWidth(cIntColEstado) = 1400
        .ColWidth(cintColRegistro) = 2500
        .ColWidth(cintColCancelo) = 2500
        If optTipo(1).Value Or optTipo(2).Value = True Then
            .TextMatrix(0, cintColProveedor) = "Médico"
            .TextMatrix(0, cintColFactura) = "Folio"
        End If
        If optTipo(2).Value = True Then .TextMatrix(0, cintColCancelo) = "Canceló/Depositó"
        If optTipo(2).Value = True Then .ColWidth(cintColRegistro) = 3570
        If optTipo(2).Value = True Then .ColWidth(cintColCancelo) = 3570
        
        .ColAlignment(cintColFecha) = flexAlignLeftCenter
        .ColAlignment(cintColNumero) = flexAlignRightCenter
        .ColAlignment(cintColProveedor) = flexAlignLeftCenter
        .ColAlignment(cintColFactura) = flexAlignLeftCenter
        .ColAlignment(cIntColEstado) = flexAlignLeftCenter
        .ColAlignment(cintColRegistro) = flexAlignLeftCenter
        .ColAlignment(cintColCancelo) = flexAlignLeftCenter
       
        .ColAlignmentFixed(cintColFecha) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColNumero) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColProveedor) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColFactura) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColEstado) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColRegistro) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCancelo) = flexAlignCenterCenter
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraBusqueda"))
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumero_KeyPress"))
End Sub

Private Sub pConsulta(lngidfactura As Long, dtmFechaIni As Date, dtmFechaFin As Date, bitFiltroFecha As Integer, lngCveProveedor As Long, intCveDepto As Integer)
    On Error GoTo NotificaError

    Dim strFechaIni As String
    Dim strFechaFin As String
    Dim strTipo As String
    
    If bitFiltroFecha = 1 Then
        strFechaIni = fstrFechaSQL(Format(dtmFechaIni, "dd/mm/yyyy"))
        strFechaFin = fstrFechaSQL(Format(dtmFechaFin, "dd/mm/yyyy"))
    Else
        strFechaIni = fstrFechaSQL(Format(fdtmServerFecha, "dd/mm/yyyy"))
        strFechaFin = fstrFechaSQL(Format(fdtmServerFecha, "dd/mm/yyyy"))
    End If
    
    If optTipo(0).Value Then strTipo = "F"
    If optTipo(1).Value Then strTipo = "H"
    If optTipo(2).Value Then strTipo = "D"
    
    vgstrParametrosSP = CStr(lngidfactura) _
    & "|" & strFechaIni _
    & "|" & strFechaFin _
    & "|" & CStr(bitFiltroFecha) _
    & "|" & CStr(lngCveProveedor) _
    & "|" & CStr(intCveDepto) _
    & "|" & vgintClaveEmpresaContable _
    & "|" & strTipo _
    & "|" & "" _
    & "|" & "-1" _
    & "|" & IIf(chkTodas.Value = vbChecked, 1, 0) _
    & "|" & IIf(chkActivas.Value = vbChecked, 1, 0) _
    & "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0) _
    & "|" & IIf(chkDepositadas.Value = vbChecked, 1, 0) _
    & "|" & IIf(chkPendientes.Value = vbChecked, 1, 0) _
    & "|" & IIf(chkReembolsadas.Value = vbChecked, 1, 0) _
    & "|" & IIf(chkSinDepositar.Value = vbChecked, 1, 0)
        
    Set rsConsulta = frsEjecuta_SP(vgstrParametrosSP, "SP_CPSELFACTURACAJACHICA")
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConsulta"))
End Sub

Private Function fblnCuentaGastos(llngCuenta As Long) As Boolean
    On Error GoTo NotificaError
    
    Dim rsResultado As New ADODB.Recordset
    
    Set rsResultado = frsEjecuta_SP(CStr(llngCuenta), "sp_CnSelCuentaContable")
    If rsResultado.RecordCount > 0 Then
        'Se hizo modificación para que tambien se admitan cuentas de tipo "COSTO"
        fblnCuentaGastos = IIf(rsResultado!vchClasificacionTipo = "Gasto", True, IIf(rsResultado!vchClasificacionTipo = "Costo", True, False))
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaGastos"))
End Function

Private Sub mskCuentaHonorario_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim llngNumeroCuenta As Long
    Dim lstrCuentaCompleta As String
    
    If KeyCode = vbKeyReturn Then
        If mskCuentaHonorario.ClipText = "" Then
            llngNumeroCuenta = flngBusquedaCuentasContables()
            txtCuentaHonorario.Text = ""
            If llngNumeroCuenta <> 0 Then
                If fblnCuentaGastos(llngNumeroCuenta) Then
                    lstrCuentaCompleta = fstrCuentaContable(llngNumeroCuenta)
                    lstrCuentaCompleta = fstrCuentaCompleta(lstrCuentaCompleta)
                    mskCuentaHonorario.Mask = ""
                    mskCuentaHonorario.Text = lstrCuentaCompleta
                    mskCuentaHonorario.Mask = vgstrEstructuraCuentaContable
                    txtCuentaHonorario.Text = fstrDescripcionCuenta(fstrCuentaContable(llngNumeroCuenta), vgintClaveEmpresaContable)
                    cboMedicos.SetFocus
                Else
                    'Debe seleccionar una cuenta clasificada como costo o gasto.
                    MsgBox SIHOMsg(1038), vbOKOnly + vbInformation, "Mensaje"
                    txtCuentaHonorario.Text = ""
                    mskCuentaHonorario.SetFocus
                    pSelMkTexto mskCuentaHonorario
                End If
            End If
        Else
            lstrCuentaCompleta = fstrCuentaCompleta(mskCuentaHonorario.Text)
            mskCuentaHonorario.Mask = ""
            mskCuentaHonorario.Text = lstrCuentaCompleta
            mskCuentaHonorario.Mask = vgstrEstructuraCuentaContable

            txtCuentaHonorario.Text = fstrDescripcionCuenta(mskCuentaHonorario.Text, vgintClaveEmpresaContable)

            If txtCuentaHonorario.Text <> "" Then
                If fblnCuentaAfectable(mskCuentaHonorario.Text, vgintClaveEmpresaContable) Then
                    If Not fblnCuentaGastos(flngNumeroCuenta(mskCuentaHonorario.Text, vgintClaveEmpresaContable)) Then
                        'Debe seleccionar una cuenta clasificada como costo o gasto.
                        MsgBox SIHOMsg(1038), vbOKOnly + vbInformation, "Mensaje"
                        txtCuentaHonorario.Text = ""
                        pSelMkTexto mskCuentaHonorario
                    End If
                Else
                    'La cuenta seleccionada no acepta movimientos.
                    MsgBox SIHOMsg(375), vbOKOnly + vbInformation, "Mensaje"
                    txtCuentaHonorario.Text = ""
                    pSelMkTexto mskCuentaHonorario
                End If
            Else
                ' No se encontró la cuenta contable.
                MsgBox SIHOMsg(222), vbOKOnly + vbInformation, "Mensaje"
                pSelMkTexto mskCuentaHonorario
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskCuentaHonorario_KeyDown"))
End Sub
Private Sub pCalculaSubtotalHonorario()
    On Error GoTo NotificaError

    Dim dblMonto As Double
    Dim dblIVA As Double
    
    dblMonto = Val(Format(txtMontoHonorario.Text, cstrFormato))
    
    dblIVA = 0
    ldblPorcentajeIVA = 0
    If OptIVAHonorario(0).Value = True Then
        If cboIVAHonorario.ListIndex <> -1 Then
            ldblPorcentajeIVA = arrImpuestosHonorario(cboIVAHonorario.ListIndex).dblPorcentaje
            dblIVA = dblMonto * ldblPorcentajeIVA / 100
        End If
        
        If chkRetencionIVAHonorario.Value = 1 Then
            chkRetencionIVAHonorario_Click
        End If
    Else
        ldblPorcentajeIVA = 0
        dblIVA = 0
    End If
    
    lblIVAHonorario.Caption = FormatCurrency(dblIVA, 2)
    lblSubtotalHonorario.Caption = FormatCurrency(dblMonto + dblIVA, 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculaSubtotalHonorario"))
End Sub
Private Sub pLlenaCombos()
    Dim strSentencia As String
    Dim rs As ADODB.Recordset
    Dim intcontador As Integer
    On Error GoTo NotificaError

    cboIVAHonorario.Clear

    vgstrParametrosSP = "-1|-1|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_GnSelImpuesto")
    intcontador = 0
    Do While Not rs.EOF
        ReDim Preserve arrImpuestosHonorario(intcontador)
        cboIVAHonorario.AddItem rs!VCHDESCRIPCION
        cboIVAHonorario.ItemData(cboIVAHonorario.newIndex) = rs!smiCveImpuesto
        arrImpuestosHonorario(intcontador).lngIdImpuesto = rs!smiCveImpuesto
        arrImpuestosHonorario(intcontador).dblPorcentaje = rs!relPorcentaje
        intcontador = intcontador + 1
        rs.MoveNext
    Loop
    
    vgstrParametrosSP = "-1|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CNSELTARIFAISR")
    If rs.RecordCount <> 0 Then
        Call pLlenarCboRs(cboTarifa, rs, 0, 1, 0)
    End If
    
    pLlenaComboMedicos
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLlenaCombos"))
End Sub
Private Sub pLlenaComboMedicos()
    Dim strSentencia As String
    
    '|    Set rs = frsEjecuta_SP("", "SP_HOMEDICOS")
    strSentencia = "Select  Distinct CoProveedor.INTCVEPROVEEDOR Clave, CoProveedor.VCHNOMBRECOMERCIAL Nombre " & _
                            "  From CoProveedor Inner Join HOMedico On (CoProveedor.VCHRFC = HOMedico.VCHRFCMEDICO ) " & _
                            " Order by Nombre"
    Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then
        Call pLlenarCboRs(cboMedicos, rs, 0, 1, 0)
        cboMedicos.ListIndex = -1
    End If
    rs.Close
End Sub
Private Sub txtMontoHonorario_Change()
    On Error GoTo NotificaError
    pCalculaSubtotalHonorario
    pCalculaTotalHonorario
    
    OptIVAHonorario(0).Enabled = Val(Format(txtMontoHonorario.Text, cstrFormato)) <> 0
    OptIVAHonorario(1).Enabled = Val(Format(txtMontoHonorario.Text, cstrFormato)) <> 0
    
    cboIVAHonorario.Enabled = Val(Format(txtMontoHonorario.Text, cstrFormato)) <> 0 And OptIVAHonorario(0).Value = True
    
    chkRetencionISRHonorario.Enabled = Val(Format(txtMontoHonorario.Text, cstrFormato)) <> 0
    
    If chkRetencionISRHonorario.Value = 1 Then
        chkRetencionISRHonorario_Click
    End If
    
    If chkRetencionIVAHonorario.Value = 1 Then
        chkRetencionIVAHonorario_Click
    End If
    
    If Val(Format(txtMontoHonorario.Text, cstrFormato)) = 0 Then
        OptIVAHonorario(0).Value = False
        OptIVAHonorario(1).Value = False
        cboIVAHonorario.ListIndex = -1
        chkRetencionISRHonorario.Value = 0
        chkRetencionIVAHonorario.Value = 0
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_Change"))
End Sub
Private Sub chkRetencionIVAHonorario_Click()
    On Error GoTo NotificaError
    
    lblRetencionIVAHonorario.Caption = FormatCurrency(fdblMontoRetencionIVA(), 2)
    pCalculaTotalHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkRetencionIVAHonorario_Click"))
End Sub
Private Sub chkRetencionISRHonorario_Click()
    On Error GoTo NotificaError
    
    cboTarifa.Enabled = chkRetencionISRHonorario.Value = 1
        
    If chkRetencionISRHonorario.Value = 1 Then
        If cboTarifa.ListCount <> 0 Then
            pCargaTarifas
            cboTarifa.ListIndex = 0
        End If
    Else
        cboTarifa.ListIndex = -1
    End If

    pCalculaTotalHonorario
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkRetencionISRHonorario_Click"))
End Sub
Private Sub txtMontoHonorario_GotFocus()
    On Error GoTo NotificaError
       
    pSelTextBox txtMontoHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_GotFocus"))
End Sub
Private Sub txtMontoHonorario_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtMontoHonorario)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_KeyPress"))
End Sub
Private Sub txtMontoHonorario_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Val(Format(txtMontoHonorario.Text, cstrFormato)) = 0 Then
            pEnfocaTextBox txtMontoHonorario
        Else
            txtMontoHonorario.Text = FormatCurrency(Val(Format(txtMontoHonorario.Text, cstrFormato)), 2)
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_KeyDown"))
End Sub
Private Function fdblMontoRetencionIVA() As Double
    On Error GoTo NotificaError
    Dim dblMonto As Double
    
    dblMonto = Val(Format(txtMontoHonorario.Text, cstrFormato))

    fdblMontoRetencionIVA = 0
    If chkRetencionIVAHonorario.Value = 1 Then
        fdblMontoRetencionIVA = (dblMonto * (ldblPorcentajeIVA / 100)) * (gdblPorcentajeRetIVA / 100)
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fdblMontoRetencionIVA"))
End Function
Private Sub pCalculaTotalHonorario()
    On Error GoTo NotificaError

    Dim dblSubTotal As Double
    Dim dblRetencionISR As Double
    Dim dblRetencionIVA As Double
    
    dblSubTotal = Val(Format(lblSubtotalHonorario, cstrFormato))
    dblRetencionISR = Val(Format(lblRetencionISRHonorario, cstrFormato))
    dblRetencionIVA = Val(Format(lblRetencionIVAHonorario, cstrFormato))

    lblTotalPagarHonorario.Caption = FormatCurrency(dblSubTotal - dblRetencionISR - dblRetencionIVA, 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculaTotalHonorario"))
End Sub

Private Sub pCargaTarifas()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim intcontador As Integer
    
    cboTarifa.Clear
    ReDim arrTarifas(0)
    
    vgstrParametrosSP = "-1|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CNSELTARIFAISR")
    If rs.RecordCount <> 0 Then
        intcontador = 0
        Do While Not rs.EOF
            ReDim Preserve arrTarifas(intcontador)
                       
            cboTarifa.AddItem rs!Descripcion
            cboTarifa.ItemData(cboTarifa.newIndex) = rs!IdTarifa
            
            arrTarifas(intcontador).lngId = rs!IdTarifa
            arrTarifas(intcontador).dblPorcentaje = rs!Porcentaje
        
            intcontador = intcontador + 1
        
            rs.MoveNext
        Loop
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaTarifas"))
End Sub
Private Sub pLimpiaHonorarios()
    On Error GoTo NotificaError
    
    pLlenaComboMedicos
    cboMedicos.ListIndex = -1
    
    vlintTipoXMLCXP = 0
    vlstrUUIDXMLCXP = ""
    vlstrRFCXMLCXP = ""
    vldblMontoXMLCXP = 0
    vlstrMonedaXMLCXP = ""
    vldblTipoCambioXMLCXP = 0
    vlstrSerieCXP = ""
    vlstrNumFolioCXP = ""
    vlstrNumFactExtCXP = ""
    vlstrTaxIDExtCXP = ""
    vlstrXMLCXP = ""
    
    chkXMLrelacionadoFact.Value = 0
    chkXMLrelacionadoHono.Value = 0
    
    mskFechaHonorario.Mask = ""
    mskFechaHonorario.Text = fdtmServerFecha
    mskFechaHonorario.Mask = "##/##/####"
    
    txtRFcHono.Text = ""
    txtFolioHonorario.Text = ""
    
    mskCuentaHonorario.Mask = ""
    mskCuentaHonorario.Text = ""
    mskCuentaHonorario.Mask = vgstrEstructuraCuentaContable
    
    txtCuentaHonorario.Text = ""
    
    optMonedaHonorario(0).Value = True
    optMonedaHonorario(1).Value = False
    
    lblEstado.Caption = ""
    txtMontoHonorario.Text = FormatCurrency(0, 2)
    
    OptIVAHonorario(0).Value = False
    OptIVAHonorario(1).Value = False
    cboIVAHonorario.ListIndex = -1
    OptIVAHonorario_Click (0)
    
    lblIVAHonorario.Caption = FormatCurrency(0, 2)
    
    lblSubtotalHonorario.Caption = FormatCurrency(0, 2)
    
    chkRetencionISRHonorario.Value = 0
    chkRetencionISRHonorario_Click
    lblRetencionISRHonorario.Caption = FormatCurrency(0, 2)
    
    chkRetencionIVAHonorario.Value = 0
    lblRetencionIVAHonorario.Caption = FormatCurrency(0, 2)
    
    lblTotalPagarHonorario.Caption = FormatCurrency(0, 2)
    
    fraHonorario.Enabled = True
    
    txtNumero.Text = frsRegresaRs("SELECT NVL(MAX(INTIDFACTURA),0)+1 FROM CPFACTURACAJACHICA").Fields(0)
    txtDescSalidaHono.Text = ""
    
    cmdCancelar.Enabled = False
         
    pHabilitaBusquedaXML
         
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaHonorarios"))
End Sub
Private Sub pHabilitaBusquedaXML()
    On Error GoTo NotificaError
    
    cmdBuscarXMLFactura.Visible = IIf(vlblnLicenciaContaElectronica, True, False)
    cmdBuscarXMLHonorario.Visible = IIf(vlblnLicenciaContaElectronica, True, False)
    fraSelXMLCajaChicaFact.Visible = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value, True, False)
    fraSelXMLCajaChicaHono.Visible = IIf(vlblnLicenciaContaElectronica And optTipo(1).Value, True, False)
    chkXMLrelacionadoFact.Visible = IIf(vlblnLicenciaContaElectronica, True, False)
    chkXMLrelacionadoHono.Visible = IIf(vlblnLicenciaContaElectronica, True, False)
    cmdBuscarXMLFactura.Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value And (optFactura.Value Or optFlete.Value), True, False)
    cmdBuscarXMLHonorario.Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(1).Value, True, False)
    fraSelXMLCajaChicaFact.Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value And (optFactura.Value Or optFlete.Value), True, False)
    fraSelXMLCajaChicaHono.Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(1).Value, True, False)
    optTipoComproCajaChicaFact(0).Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value And (optFactura.Value Or optFlete.Value), True, False)
    optTipoComproCajaChicaFact(1).Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value And (optFactura.Value Or optFlete.Value), True, False)
    optTipoComproCajaChicaFact(2).Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value And (optFactura.Value Or optFlete.Value), True, False)
    optTipoComproCajaChicaHono(0).Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(1).Value, True, False)
    optTipoComproCajaChicaHono(1).Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(1).Value, True, False)
    optTipoComproCajaChicaHono(2).Enabled = IIf(vlblnLicenciaContaElectronica And optTipo(1).Value, True, False)
    chkXMLrelacionadoFact.Value = IIf(vlblnLicenciaContaElectronica, IIf(optTipo(0).Value And (optFactura.Value Or optFlete.Value) And IIf(Trim(vlintTipoXMLCXP) = "", 0, vlintTipoXMLCXP) > 0, 2, 0), 0)
    chkXMLrelacionadoHono.Value = IIf(vlblnLicenciaContaElectronica, IIf(optTipo(1).Value And IIf(Trim(vlintTipoXMLCXP) = "", 0, vlintTipoXMLCXP) > 0, 2, 0), 0)
    
    optTipoComproCajaChicaFact(0).Value = True
    optTipoComproCajaChicaHono(0).Value = True
    
    frmCajaChica.Refresh
'   frmCajaChica.Height = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value, 6495, 5820)
    'frmCajaChica.Height = IIf(vlblnLicenciaContaElectronica, 7050, 6420)
    frmCajaChica.Top = Int((SysInfo.WorkAreaHeight - frmCajaChica.Height) / 2)
    frmCajaChica.Refresh
    
'   FraBotonera.Top = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value, 5520, 4840)
'   fraFactura.Height = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value, 3450, 2730)
'   fraHonorario.Height = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value, 3450, 2730)
'   grdFacturas.Height = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value, 5000, 4340)
'   FraConsulta.Height = IIf(vlblnLicenciaContaElectronica And optTipo(0).Value, 5190, 4530)
    
    'fraBotonera.Top = IIf(vlblnLicenciaContaElectronica, 6120, 5470)
'    fraFactura.Height = IIf(vlblnLicenciaContaElectronica, 4090, 3450)
'   fraSelXMLCajaChicaFact
   ' fraHonorario.Height = IIf(vlblnLicenciaContaElectronica, 4090, 3365)
    grdFacturas.Height = IIf(vlblnLicenciaContaElectronica, 5570, 4930)
    FraConsulta.Height = IIf(vlblnLicenciaContaElectronica, 5760, 5115)
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaBusquedaXML"))
End Sub
Private Sub txtRFC_GotFocus()
        pSelTextBox txtRFC
End Sub
Private Sub txtRFC_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
Dim vlstrcaracter As String
'    If KeyAscii = 13 Then
'        If txtRFC.Text = "" Then
'            MsgBox "Favor de ingresar el RFC del proveedor o acreedor.", vbOKOnly + vbExclamation, "Mensaje"
'            txtRFC.SetFocus
'        ElseIf Len(txtRFC.Text) < 13 Then
'            MsgBox SIHOMsg(1345), vbExclamation + vbOKOnly, "Mensaje"
'            txtRFC.SetFocus
'        End If
'    Else
        If KeyAscii <> 8 Then
            If KeyAscii <> 13 Then
                vlstrcaracter = fStrRFCValido(Chr(KeyAscii))
                If vlstrcaracter <> "" Then
                    KeyAscii = Asc(UCase(vlstrcaracter))
                Else
                    KeyAscii = 7
                End If
            End If
        End If
'    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtRFC_KeyPress"))
End Sub

Private Sub txtRFcHono_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    Dim vlstrcaracter As String

        If KeyAscii <> 8 Then
            If KeyAscii <> 13 Then
                vlstrcaracter = fStrRFCValido(Chr(KeyAscii))
                If vlstrcaracter <> "" Then
                    KeyAscii = Asc(UCase(vlstrcaracter))
                Else
                    KeyAscii = 7
                End If
            End If
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtRFcHono_KeyPress"))
End Sub

Private Sub txtTotalDF_LostFocus()
If txtTotalDF <> "" Then
    txtTotalDF.Text = FormatCurrency(Val(Format(txtTotalDF.Text, cstrFormato)), 2)
End If
End Sub

Private Sub txtTotalDF_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtTotalDF)) Or Not fblnFormatoCantidad(txtTotalDF, KeyAscii, 2) Then
        KeyAscii = 7
    End If
End Sub
Private Sub txtTotalTicket_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtTotalTicket

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtTotalTicket_GotFocus"))
End Sub

Private Sub txtTotalTicket_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtTotalTicket)) Then KeyAscii = 7
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtTotalTicket_KeyPress"))
End Sub
Private Sub txtTotalTicket_LostFocus()
    On Error GoTo NotificaError

    txtTotalTicket.Text = FormatCurrency(Val(Format(txtTotalTicket.Text, cstrFormato)), 2)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtTotalTicket_LostFocus"))
End Sub
Public Function fcurObtenerResico(lngCveProveedor As Long) As Currency
    On Error GoTo NotificaError
    
    Dim rsRetencion As New ADODB.Recordset
    Dim vlstrSentencia As String

    fcurObtenerResico = 0
    vlstrSentencia = "SELECT CnRegimenRetencion.intidtarifa, cnTarifaISR.numporcentaje " & _
                     "FROM CoProveedor INNER JOIN CnRegimenRetencion ON trim(CoProveedor.vchclaveregimensat) = trim(CnRegimenRetencion.chridregimen) " & _
                     "INNER JOIN cnTarifaISR ON CnRegimenRetencion.intidtarifa = cnTarifaISR.intidtarifa " & _
                     "WHERE CoProveedor.vchtiporegimen <> 'MORAL' AND CoProveedor.VCHCLAVEREGIMENSAT = '626' AND CoProveedor.intCveProveedor = " & lngCveProveedor
    
    Set rsRetencion = frsRegresaRs(vlstrSentencia)
    If rsRetencion.RecordCount <> 0 Then
        fcurObtenerResico = rsRetencion!NUMPORCENTAJE
    End If
    rsRetencion.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fcurObtenerResico"))
End Function
Public Function fcurVerifRegimen626(lngCveProveedor As Long) As Boolean
    On Error GoTo NotificaError
    
    Dim rsRetencion As New ADODB.Recordset
    Dim vlstrSentencia As String

    fcurVerifRegimen626 = False
    vlstrSentencia = "SELECT VCHCLAVEREGIMENSAT FROM CoProveedor WHERE VCHCLAVEREGIMENSAT = '626' AND intCveProveedor = " & lngCveProveedor
    
    Set rsRetencion = frsRegresaRs(vlstrSentencia)
    If rsRetencion.RecordCount <> 0 Then
        fcurVerifRegimen626 = True
    End If
    rsRetencion.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fcurVerifRegimen626"))
End Function
Private Function TraeCuentasISRProv(vlProv As Long, vlintEmpresaContable As Integer, vlchrProceso As String, vlchrTipoCuenta As String, vlchrTipoProv As String) As Long
    Dim rsRetencion As New ADODB.Recordset
    Dim strSql As String
    Dim vlblnUpdCveRetencionHonMed As Boolean 'Valida si se realizo el update en el campo intCveRetencion de la tabla Homedico MELR 060423
    Dim rsRetencionHonMed As New ADODB.Recordset 'MELR 070423
    
   If vlchrTipoProv = "P" Then
   
        If vlProv <> -1 Then
      strSql = " SELECT intcuenta " & _
               " FROM cntarifaisrcuentas " & _
               " WHERE tnyclaveempresa =" & CStr(vlintEmpresaContable) & _
               " AND chrtipoproceso = '" & vlchrProceso & "'" & _
               " AND chrtipocuenta = '" & vlchrTipoCuenta & "'" & _
               " AND intidtarifa = " & _
               "     (SELECT intidtarifa " & _
               "      FROM cnregimenretencion " & _
               "      WHERE CHRTIPORETENCION = 'S' " & _
               "            AND chridregimen = (SELECT vchclaveregimensat " & _
               "                                FROM coproveedor " & _
               "                                WHERE intcveproveedor = " & CStr(vlProv) & "))"
        Else
      strSql = " SELECT intcuenta " & _
               " FROM cntarifaisrcuentas " & _
               " WHERE tnyclaveempresa =" & CStr(vlintEmpresaContable) & _
               " AND chrtipoproceso = '" & vlchrProceso & "'" & _
               " AND chrtipocuenta = '" & vlchrTipoCuenta & "'" & _
               " AND intidtarifa = " & cboRetencionISR.ItemData(cboRetencionISR.ListIndex)
        End If
     
   Else
        '--- MELR 060423 ---'
        strSql = " SELECT * from cntarifaisr where intidtarifa = 1 "
        Set rsRetencionHonMed = frsRegresaRs(strSql, adLockOptimistic)
        If rsRetencionHonMed.RecordCount > 0 Then
            strSql = " SELECT INTCVERETENCION FROM homedico " & _
                       " WHERE intcvemedico = " & CStr(vlProv) & ""
            Set rsRetencion = frsRegresaRs(strSql, adLockOptimistic)
            If rsRetencion.RecordCount > 0 Then
                vlblnUpdCveRetencionHonMed = IsNull(rsRetencion!INTCVERETENCION)
                If vlblnUpdCveRetencionHonMed = True Then
                    pEjecutaSentencia "update homedico set INTCVERETENCION = 1 where intcvemedico = " & CStr(vlProv) & ""
                End If
            End If
        End If
        '--- MELR 060423 ---'
      
      ' INTCVERETENCION  == ES EL IDTARIFA DE ISR PARA MEDICOS
      strSql = " SELECT intcuenta " & _
               " FROM cntarifaisrcuentas " & _
               " WHERE tnyclaveempresa = " & CStr(vlintEmpresaContable) & _
               " AND chrtipoproceso = '" & vlchrProceso & "'" & _
               " AND chrtipocuenta = '" & vlchrTipoCuenta & "'" & _
               " AND intidtarifa = " & _
               "                  (SELECT INTCVERETENCION " & _
               "                   FROM homedico " & _
               "                   WHERE intcvemedico = " & CStr(vlProv) & ")"
   End If
      
   Set rsRetencion = frsRegresaRs(strSql, adLockOptimistic)
   If rsRetencion.RecordCount > 0 Then
      TraeCuentasISRProv = rsRetencion!intCuenta
   Else
      TraeCuentasISRProv = 0
   End If
                       
   '--- MELR 060423 ---'
   If vlblnUpdCveRetencionHonMed = True Then
        pEjecutaSentencia "update homedico set INTCVERETENCION = null where intcvemedico = " & CStr(vlProv) & ""
   End If
   '--- MELR 060423 ---'
   
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " :frmProcesoCxP.frm " & ":TraeCuentasISRProv"))
End Function



