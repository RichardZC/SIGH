VERSION 5.00
Begin VB.Form RecibosBoleta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de recibo"
   ClientHeight    =   8430
   ClientLeft      =   6270
   ClientTop       =   1455
   ClientWidth     =   7395
   Icon            =   "RecibosBoleta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RecibosBoleta.frx":000C
   ScaleHeight     =   8430
   ScaleWidth      =   7395
   Begin VB.TextBox txtHora 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   43
      Top             =   1840
      Width           =   515
   End
   Begin VB.PictureBox PicBoletaDetalle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2460
      Left            =   0
      Picture         =   "RecibosBoleta.frx":292B
      ScaleHeight     =   2460
      ScaleWidth      =   7425
      TabIndex        =   35
      Top             =   3130
      Width           =   7425
      Begin VB.PictureBox PicBoletaDetalle2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         Picture         =   "RecibosBoleta.frx":479C
         ScaleHeight     =   2415
         ScaleWidth      =   6975
         TabIndex        =   37
         Top             =   0
         Width           =   6975
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   6150
            TabIndex        =   42
            Top             =   0
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   41
            Top             =   0
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox txtPrecioUnit 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   5460
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   700
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4880
            TabIndex        =   39
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtDescripcion 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   38
            Top             =   0
            Visible         =   0   'False
            Width           =   3940
         End
      End
      Begin VB.VScrollBar vsTratamiento 
         Height          =   2415
         Left            =   6960
         TabIndex        =   36
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox PicBoletaTotales 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      Picture         =   "RecibosBoleta.frx":660D
      ScaleHeight     =   3015
      ScaleWidth      =   7455
      TabIndex        =   23
      Top             =   5520
      Width           =   7455
      Begin VB.TextBox txtObservaciones 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   350
         TabIndex        =   45
         Top             =   380
         Width           =   4680
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   750
         TabIndex        =   44
         Top             =   1480
         Width           =   1335
      End
      Begin VB.TextBox txtDctos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   30
         Top             =   2040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6230
         TabIndex        =   29
         Top             =   170
         Width           =   825
      End
      Begin VB.TextBox txtExonerado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6230
         TabIndex        =   28
         Top             =   730
         Width           =   825
      End
      Begin VB.TextBox txtLetras 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   690
         TabIndex        =   27
         Top             =   1050
         Width           =   4980
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Salir (ESC)"
         DisabledPicture =   "RecibosBoleta.frx":9ED8
         DownPicture     =   "RecibosBoleta.frx":A39C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2520
         Picture         =   "RecibosBoleta.frx":A888
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2040
         Width           =   1065
      End
      Begin VB.CommandButton btnReImprime 
         Caption         =   "ReImprime"
         Height          =   705
         Left            =   3600
         Picture         =   "RecibosBoleta.frx":AD74
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2040
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox TxtAdelantos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6230
         TabIndex        =   24
         Top             =   450
         Width           =   825
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6230
         TabIndex        =   57
         Top             =   1040
         Width           =   825
      End
      Begin VB.Label txtCajero 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6310
         TabIndex        =   56
         Top             =   1480
         Width           =   690
      End
      Begin VB.Label txtTerminal 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1160
         TabIndex        =   46
         Top             =   1790
         Width           =   1335
      End
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   12435
      TabIndex        =   17
      Top             =   5460
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   12435
      TabIndex        =   16
      Top             =   5220
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   12435
      TabIndex        =   15
      Top             =   4965
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   12435
      TabIndex        =   14
      Top             =   4740
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   12435
      TabIndex        =   13
      Top             =   4485
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   12435
      TabIndex        =   12
      Top             =   4260
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   12435
      TabIndex        =   11
      Top             =   4020
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   12435
      TabIndex        =   10
      Top             =   3780
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   12435
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   12435
      TabIndex        =   8
      Top             =   6090
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   12435
      TabIndex        =   7
      Top             =   6420
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   12435
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   12435
      TabIndex        =   5
      Top             =   7020
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   12435
      TabIndex        =   4
      Top             =   7320
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   12435
      TabIndex        =   3
      Top             =   7620
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   12435
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox dev 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   12435
      TabIndex        =   1
      Top             =   8220
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   700
      Left            =   11640
      Picture         =   "RecibosBoleta.frx":B24D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox BoletaCabecera 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      Picture         =   "RecibosBoleta.frx":B726
      ScaleHeight     =   3135
      ScaleWidth      =   7335
      TabIndex        =   21
      Top             =   0
      Width           =   7335
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   7215
         TabIndex        =   34
         Top             =   3360
         Width           =   7215
      End
      Begin VB.TextBox txtTipoPago 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5940
         TabIndex        =   32
         Top             =   1560
         Width           =   940
      End
      Begin VB.Label txtServicio 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1740
         TabIndex        =   55
         Top             =   2455
         Width           =   2205
      End
      Begin VB.Label txtOrden 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6330
         TabIndex        =   54
         Top             =   2175
         Width           =   535
      End
      Begin VB.Label txtCuenta 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4040
         TabIndex        =   53
         Top             =   2180
         Width           =   975
      End
      Begin VB.Label txtHistoriaClinica 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1530
         TabIndex        =   52
         Top             =   2180
         Width           =   1215
      End
      Begin VB.Label txtNroComprobante 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5150
         TabIndex        =   51
         Top             =   1000
         Width           =   1605
      End
      Begin VB.Label txtRuc 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5150
         TabIndex        =   50
         Top             =   380
         Width           =   1605
      End
      Begin VB.Label txtTelefono 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1845
         TabIndex        =   49
         Top             =   1200
         Width           =   2445
      End
      Begin VB.Label txtDireccion 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1845
         TabIndex        =   48
         Top             =   720
         Width           =   2445
      End
      Begin VB.Label txtRazonSocial 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   47
         Top             =   1830
         Width           =   4950
      End
      Begin VB.Label txtNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   455
         Left            =   1845
         TabIndex        =   22
         Top             =   240
         Width           =   2445
      End
   End
   Begin VB.Label lblDev 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Devol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   12435
      TabIndex        =   20
      Top             =   3540
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label LblAnulado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   990
      Left            =   7440
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   6765
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   4785
      Left            =   12480
      TabIndex        =   18
      Top             =   3720
      Width           =   765
   End
End
Attribute VB_Name = "RecibosBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Consulta un Documento de CAJA
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lnIdUsuario As Long
Dim lnNroItemsBoleta As Long
Dim oRsItemsBoleta As New ADODB.Recordset
Dim lnFarmaciaServicio As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasCaja As New ReglasCaja
Dim lcNroBoleta As String, lcNroSerie As String, lnBienFarmacia As Long
Dim ml_lbTienePermisoReimprimeBoleta As Boolean
Dim lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio As Boolean

Property Let lbTienePermisoReimprimeBoleta(lValue As Boolean)
    ml_lbTienePermisoReimprimeBoleta = lValue
    btnReImprime.Visible = ml_lbTienePermisoReimprimeBoleta
    cmdExcel.Visible = ml_lbTienePermisoReimprimeBoleta
End Property

Property Let EsAnulado(lValue As Integer)
  If lValue = 9 Then
    btnReImprime.Visible = False
  Else
    btnReImprime.Visible = True
  End If
End Property

Sub Imprimir(oDOPaciente As doPaciente, oDOAtencion As DOAtencion, oDOComprobantePago As DOCajaComprobantesPago, _
             rsFacturacionProductos As Recordset)
  'Dim rsReporte As New Recordset
  'Dim rsReporte1 As New Recordset
  Dim iFila As Long
  Dim mdTotal  As Double
  Dim oRecibosBoleta As New RecibosBoleta
    Me.txtRazonSocial = oDOComprobantePago.RazonSocial + IIf(IsNull(oDOPaciente.NroHistoriaClinica), "(Particular)", "(HC: " & oDOPaciente.NroHistoriaClinica & ")")
    Me.txtNroComprobante = oDOComprobantePago.nroSerie + " - " + oDOComprobantePago.nrodocumento
    Me.txtFecha = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    Me.txtDctos.Text = "Adelantos: " & Format(oDOComprobantePago.Dctos, "######.#0")
    Me.txtDatos = "Cta: " & Trim(Str(oDOComprobantePago.idCuentaAtencion)) & "    Ord.Pag: " & Trim(Str(rsFacturacionProductos!IdOrden))
    iFila = 1
    mdTotal = 0
    Do While Not rsFacturacionProductos.EOF

        Me.txtCodigo(iFila - 1) = rsFacturacionProductos!Codigo
        Me.txtDescripcion(iFila - 1) = rsFacturacionProductos!NombreProducto
        Me.txtPrecioUnit(iFila - 1) = Format(rsFacturacionProductos!PrecioUnitario, "######.#0")
        Me.txtCantidad(iFila - 1) = rsFacturacionProductos!Cantidad
        mdTotal = mdTotal + rsFacturacionProductos!TotalPorPagar
        Me.txtImporte(iFila - 1) = Format(rsFacturacionProductos!TotalPorPagar, "######.#0")

        rsFacturacionProductos.MoveNext

        If iFila = 8 Then
             Exit Do
        End If

        iFila = iFila + 1
    Loop

    Me.txtTotal = Format(mdTotal, "#######.#0")

Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub

End Sub

Sub LimpiarLineas()
Dim i As Integer

    For i = 0 To 7
        Me.txtCodigo(i) = ""
        Me.txtDescripcion(i) = ""
        Me.txtPrecioUnit(i) = ""
        Me.txtCantidad(i) = ""
        Me.txtImporte(i) = ""
    Next i

End Sub


'***************daniel barrantes**************
'***************Impresion en PANTALLA de una BOLETA
'***************
Sub ImprimirDEBB(nroSerie As String, nroDcto As String, lcFarmaciaServicio As Long, lnIdUsuarioSistema As Long)
Dim rsreporte As New Recordset
Dim mo_PermisosFacturacion As New PermisosFacturacion
Dim iFila As Long: Dim ln As Integer
Dim mdTotal  As Double: Dim lnTotalBoleta As Double
Dim oRecibosBoleta As New RecibosBoleta
Dim lnImporteEXO As Double: Dim lnSubTotal As Double
Dim lnDctos As Double
Dim oReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lbBoletaFarmaciaFormTicket As Boolean
Dim lcDevolucion As String
        lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio = False
        lnIdUsuario = lnIdUsuarioSistema
        lbBoletaFarmaciaFormTicket = False
        If lcFarmaciaServicio = SIGHEntidades.sghbien Then
           Set rsreporte = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(nroSerie, nroDcto)
           
        Else
           Set rsreporte = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(nroSerie, nroDcto)
           If rsreporte.RecordCount > 0 Then
                If Not IsNull(rsreporte.Fields!idCuentaAtencion) And rsreporte.Fields!idCuentaAtencion > 0 Then
                   lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio = True
                End If
           End If
        End If
        'SUNAT
        If lcFarmaciaServicio = SIGHEntidades.sghbien Then
            If IsNull(rsreporte.Fields!FormatoImp2Cinta) Then
                lbBoletaFarmaciaFormTicket = False
            Else
                If rsreporte.Fields!FormatoImp2Cinta = True Then lbBoletaFarmaciaFormTicket = True
            End If
        Else
            If IsNull(rsreporte.Fields!FormatoImpDefaultCinta) Then
                lbBoletaFarmaciaFormTicket = False
            Else
                If rsreporte.Fields!FormatoImpDefaultCinta = True Then lbBoletaFarmaciaFormTicket = True
            End If
        End If
        'SUNAT
        lnNroItemsBoleta = rsreporte.RecordCount
        If lnNroItemsBoleta > 0 Then
            
            wxIdTipoComprobanteDefault = IIf(IsNull(rsreporte.Fields!IdTipoComprobante), 3, rsreporte.Fields!IdTipoComprobante)
            CargaSetup_Caja App.Path & "\archivos", wxIdTipoComprobanteDefault, lbBoletaFarmaciaFormTicket
            
            Set oRsItemsBoleta = rsreporte.Clone
            Select Case rsreporte.Fields!IdEstadoComprobante
            Case 9
               LblAnulado.Visible = True
            Case 6
               LblAnulado.Caption = "Devolución"
               LblAnulado.Visible = True
               LblAnulado.ForeColor = vbGreen
            End Select
            lcNroSerie = nroSerie
            lcNroBoleta = nroDcto
            lnBienFarmacia = lcFarmaciaServicio
            lcDevolucion = IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", "")
            txtTipoPago.Text = lcDevolucion
            Me.txtNombre = lcBuscaParametro.SeleccionaFilaParametro(205)
            Me.txtDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
            Me.txtTelefono = UCase("telf: ") & lcBuscaParametro.SeleccionaFilaParametro(207)
            Me.txtRuc = lcBuscaParametro.SeleccionaFilaParametro(339)
            Me.txtHora = Format$(rsreporte.Fields!FechaCobranza, SIGHEntidades.DevuelveHoraSoloFormato_HM)
            Me.txtTerminal = lcBuscaParametro.RetornaNombreDeServidor
            Me.txtCaja = IIf(IsNull(rsreporte.Fields!nombreCaja), "", rsreporte.Fields!nombreCaja)
            Me.txtObservaciones = (IIf(IsNull(rsreporte.Fields!Observaciones), "", rsreporte.Fields!Observaciones))
            
            Me.txtRazonSocial = Trim(IIf(IsNull(rsreporte.Fields!RazonSocial), "", rsreporte.Fields!RazonSocial))
            Me.txtHistoriaClinica.Caption = IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "" & rsreporte.Fields!NroHistoriaClinica & ")")
            Me.txtNroComprobante.Caption = nroSerie & " - " & nroDcto
            Me.txtFecha.Text = Format(rsreporte.Fields!FechaCobranza, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
            Me.TxtAdelantos.Text = Format(IIf(IsNull(rsreporte.Fields!Adelantos), 0, rsreporte.Fields!Adelantos), "######.#0")
            lnDctos = rsreporte.Fields!Adelantos
            lnImporteEXO = rsreporte.Fields!exoneraciones
            If lcFarmaciaServicio = SIGHEntidades.sghbien Then
               If IsNull(rsreporte.Fields!idPreVenta) Then
                   Me.txtDatos.Text = "Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
                   If lcFarmaciaServicio <> SIGHEntidades.sghbien Then
                      Me.txtOrden.Caption = Trim(Str(IIf(IsNull(rsreporte.Fields!IdOrdenPago), 0, rsreporte.Fields!IdOrdenPago)))
                   End If
                   Me.txtCuenta.Caption = Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
               Else
                   Me.txtCuenta.Caption = Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion))) & "    PreVenta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idPreVenta), 0, rsreporte.Fields!idPreVenta)))
               End If
            Else
                Me.txtDatos.Text = "Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
                Me.txtOrden.Caption = Trim(Str(IIf(IsNull(rsreporte.Fields!IdOrdenPago), 0, rsreporte.Fields!IdOrdenPago)))
                Me.txtCuenta.Caption = Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
            End If
            Me.txtTotal.Caption = Format(rsreporte.Fields!TotalBoleta, "#######.#0")
            lnTotalBoleta = IIf(IsNull(rsreporte.Fields!TotalBoleta), "0", rsreporte.Fields!TotalBoleta)
            txtCajero.Caption = oReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales)
            txtServicio.Caption = oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion))
            iFila = 1
            mdTotal = IIf(IsNull(rsreporte.Fields!TotalBoleta), 0, rsreporte.Fields!TotalBoleta)
            txtLetras.Text = SIGHEntidades.Numlet(SIGHEntidades.DevuelveNumeroSinDecimales(mdTotal)) + " con " + SIGHEntidades.DevuelveSoloDecimales(mdTotal) + "/100   Nuevos Soles"
            lnNroItemsBoleta = 0
            Do While Not rsreporte.EOF
            
                Me.PicBoletaDetalle2.Height = 600 * rsreporte.RecordCount
                vsTratamiento.Max = PicBoletaDetalle2.Height

               
                Load txtCodigo(iFila)
                txtCodigo(iFila).Visible = True
                txtCodigo(iFila).Left = 200
                txtCodigo(iFila).Top = 10 + 330 * (iFila - 1)
                Me.txtCodigo(iFila) = rsreporte.Fields!Codigo
                
                Load txtDescripcion(iFila)
                txtDescripcion(iFila).Visible = True
                txtDescripcion(iFila).Left = 960
                txtDescripcion(iFila).Top = 10 + 330 * (iFila - 1)
                Me.txtDescripcion(iFila) = rsreporte.Fields!NombreProducto
                
                Load txtCantidad(iFila)
                txtCantidad(iFila).Visible = True
                txtCantidad(iFila).Left = 4880
                txtCantidad(iFila).Top = 10 + 330 * (iFila - 1)
                Me.txtCantidad(iFila) = rsreporte.Fields!Cantidad
                
                Load txtPrecioUnit(iFila)
                txtPrecioUnit(iFila).Visible = True
                txtPrecioUnit(iFila).Left = 5450
                txtPrecioUnit(iFila).Top = 10 + 330 * (iFila - 1)
                Me.txtPrecioUnit(iFila) = Format(rsreporte.Fields!PrecioUnitario, "######.#0")
                
                Load txtImporte(iFila)
                txtImporte(iFila).Visible = True
                txtImporte(iFila).Left = 6150
                txtImporte(iFila).Top = 10 + 330 * (iFila - 1)
                lnSubTotal = rsreporte.Fields!TotalPorPagar
                Me.txtImporte(iFila) = Format(lnSubTotal, "######.#0")

                'Me.txtCodigo(iFila - 1) = rsReporte.Fields!Codigo
                'Me.txtDescripcion(iFila - 1) = rsReporte.Fields!NombreProducto
                'Me.txtPrecioUnit(iFila - 1) = Format(rsReporte.Fields!PrecioUnitario, "######.#0")
                'Me.txtCantidad(iFila - 1) = rsReporte.Fields!Cantidad
                'Me.txtImporte(iFila - 1) = Format(lnSubTotal, "######.#0")
                lnNroItemsBoleta = lnNroItemsBoleta + 1
                iFila = iFila + 1

                rsreporte.MoveNext
                If iFila > 17 Then
                  Exit Do
                End If
            Loop
            txtExonerado.Text = Format(lnImporteEXO, "######.#0")
            If txtTipoPago.Text = "" And lnTotalBoleta = 0 And lnImporteEXO > 0 Then
               txtTipoPago.Text = "EXONERADO TOTAL"
            End If
            txtSubTotal.Text = Format(mdTotal + lnImporteEXO + lnDctos, "######.#0")
            'Autorizado para Devoluciones
            If LblAnulado.Visible = False And lnTotalBoleta > 0 And lcDevolucion = "" Then
                Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(lnIdUsuarioSistema)
                If mo_PermisosFacturacion.AutorizarDevoluciones Then
                   For ln = 0 To lnNroItemsBoleta - 1
                       dev(ln).Visible = True
                   Next

                End If
            End If
        End If
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub

End Sub


Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnReImprime_Click()
    If MsgBox("Por favor confirmar, ¿Realmente desea REIMPRIMIR  ?", vbQuestion + vbYesNo, "Estado de Cuenta") = vbNo Then
        Exit Sub
    End If
    Dim oImprimeBoletaContinua As New RptCaja
    If lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio = True Then
       oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE False, "", "", lcNroSerie, lcNroBoleta, True, lnBienFarmacia, True, True
    Else
       oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE False, "", "", lcNroSerie, lcNroBoleta, True, lnBienFarmacia, False, True
    End If
    Set oImprimeBoletaContinua = Nothing
    Unload Me
End Sub

Private Sub cmdExcel_Click()
   Dim oExcel As New RptCaja
   oExcel.ImpresionBoletaEnExcel lcNroSerie, lcNroBoleta, lnBienFarmacia
   Set oExcel = Nothing
End Sub





Private Sub dev_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(txtCantidad(Index).Text) < Val(dev(Index).Text) Then
          MsgBox "La DEVOLUCION debe ser Menor o Igual a la Cantidad de la Boleta", vbExclamation, "Mensaje"
          dev(Index).Text = 0
       Else
          SendKeys "{tab}"
       End If
    End If
End Sub

Private Sub dev_LostFocus(Index As Integer)
    dev_KeyPress Index, 13
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyEscape
             btnCancelar_Click
       End Select
End Sub

Private Sub Form_Load()
    If ml_lbTienePermisoReimprimeBoleta = True Then
       btnReImprime.Visible = True
       cmdExcel.Visible = True
    End If
    'wxParametro288 = lcBuscaParametro.SeleccionaFilaParametro(288)
    'CargaSetup_txtCaja App.Path & "\archivos"
    PicBoletaDetalle2.Top = 0
    PicBoletaDetalle2.Left = 0
    
    With vsTratamiento 'Si vas a utilizar el Vertical
        .Min = 0
        .SmallChange = 90
        .LargeChange = 300
        .Top = 0
        .ZOrder 0
    End With
End Sub

Private Sub vsTratamiento_Change()
    Me.PicBoletaDetalle2.Top = -vsTratamiento.Value
End Sub

Private Sub vsTratamiento_Scroll()
    Me.PicBoletaDetalle2.Top = -vsTratamiento.Value
End Sub
