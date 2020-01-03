VERSION 5.00
Begin VB.Form Recibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de recibo"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   Icon            =   "Recibos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Recibos.frx":000C
   ScaleHeight     =   7845
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRazonSocial 
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
      Left            =   1500
      TabIndex        =   116
      Top             =   1050
      Width           =   4665
   End
   Begin VB.TextBox txtServicio 
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
      Left            =   1485
      TabIndex        =   115
      Top             =   1320
      Width           =   2205
   End
   Begin VB.TextBox txtFecha 
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
      Left            =   4260
      TabIndex        =   114
      Top             =   1290
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   0
      Left            =   1440
      TabIndex        =   113
      Top             =   1740
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   810
      TabIndex        =   112
      Top             =   1740
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   4170
      TabIndex        =   111
      Top             =   1740
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   4800
      TabIndex        =   110
      Top             =   1740
      Width           =   615
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   5430
      TabIndex        =   109
      Top             =   1740
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   810
      TabIndex        =   108
      Top             =   1980
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   810
      TabIndex        =   107
      Top             =   2205
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   810
      TabIndex        =   106
      Top             =   2460
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   4
      Left            =   810
      TabIndex        =   105
      Top             =   2700
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   810
      TabIndex        =   104
      Top             =   2940
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   6
      Left            =   810
      TabIndex        =   103
      Top             =   3180
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   7
      Left            =   810
      TabIndex        =   102
      Top             =   3480
      Width           =   585
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   1
      Left            =   1440
      TabIndex        =   101
      Top             =   1980
      Width           =   2715
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   2
      Left            =   1425
      TabIndex        =   100
      Top             =   2220
      Width           =   2715
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   4
      Left            =   1440
      TabIndex        =   99
      Top             =   2700
      Width           =   2715
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   5
      Left            =   1470
      TabIndex        =   98
      Top             =   2940
      Width           =   2715
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   6
      Left            =   1440
      TabIndex        =   97
      Top             =   3180
      Width           =   2715
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   7
      Left            =   1440
      TabIndex        =   96
      Top             =   3480
      Width           =   2715
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   4170
      TabIndex        =   95
      Top             =   1980
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   4170
      TabIndex        =   94
      Top             =   2220
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   4170
      TabIndex        =   93
      Top             =   2460
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   4
      Left            =   4170
      TabIndex        =   92
      Top             =   2700
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   4170
      TabIndex        =   91
      Top             =   2940
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   6
      Left            =   4170
      TabIndex        =   90
      Top             =   3180
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   7
      Left            =   4170
      TabIndex        =   89
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   4800
      TabIndex        =   88
      Top             =   1980
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   4800
      TabIndex        =   87
      Top             =   2220
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   4800
      TabIndex        =   86
      Top             =   2460
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   4
      Left            =   4800
      TabIndex        =   85
      Top             =   2700
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   4800
      TabIndex        =   84
      Top             =   2940
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   6
      Left            =   4800
      TabIndex        =   83
      Top             =   3180
      Width           =   615
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   7
      Left            =   4800
      TabIndex        =   82
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   5430
      TabIndex        =   81
      Top             =   1980
      Width           =   735
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   5430
      TabIndex        =   80
      Top             =   2220
      Width           =   735
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   5430
      TabIndex        =   79
      Top             =   2460
      Width           =   735
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   4
      Left            =   5430
      TabIndex        =   78
      Top             =   2700
      Width           =   735
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   5430
      TabIndex        =   77
      Top             =   2925
      Width           =   735
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   6
      Left            =   5430
      TabIndex        =   76
      Top             =   3180
      Width           =   735
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   7
      Left            =   5430
      TabIndex        =   75
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   5010
      TabIndex        =   74
      Top             =   6570
      Width           =   1065
   End
   Begin VB.TextBox txtExonerado 
      Alignment       =   1  'Right Justify
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
      Left            =   5010
      TabIndex        =   73
      Top             =   6900
      Width           =   1065
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   5010
      TabIndex        =   72
      Top             =   7230
      Width           =   1065
   End
   Begin VB.TextBox txtNroComprobante 
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
      Left            =   3990
      TabIndex        =   71
      Top             =   0
      Width           =   2205
   End
   Begin VB.TextBox txtDctos 
      Alignment       =   1  'Right Justify
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
      Left            =   1710
      TabIndex        =   70
      Top             =   6930
      Width           =   2295
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   1  'Right Justify
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
      Left            =   735
      TabIndex        =   69
      Top             =   6600
      Width           =   3270
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   3
      Left            =   1440
      TabIndex        =   68
      Top             =   2460
      Width           =   2715
   End
   Begin VB.TextBox txtTipoPago 
      Alignment       =   2  'Center
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
      Left            =   1500
      TabIndex        =   67
      Top             =   15
      Width           =   2145
   End
   Begin VB.TextBox txtCajero 
      Alignment       =   1  'Right Justify
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
      Left            =   735
      TabIndex        =   66
      Top             =   6915
      Width           =   375
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
      Left            =   6195
      TabIndex        =   65
      Top             =   3420
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
      Left            =   6195
      TabIndex        =   64
      Top             =   3180
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
      Left            =   6195
      TabIndex        =   63
      Top             =   2925
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
      Left            =   6195
      TabIndex        =   62
      Top             =   2700
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
      Left            =   6195
      TabIndex        =   61
      Top             =   2445
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
      Left            =   6195
      TabIndex        =   60
      Top             =   2220
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
      Left            =   6195
      TabIndex        =   59
      Top             =   1980
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
      Left            =   6195
      TabIndex        =   58
      Top             =   1740
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtLetras 
      Alignment       =   1  'Right Justify
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
      Left            =   1005
      TabIndex        =   57
      Top             =   7560
      Width           =   5070
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
      Left            =   6195
      TabIndex        =   56
      Top             =   3720
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   8
      Left            =   5430
      TabIndex        =   55
      Top             =   3780
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   8
      Left            =   4800
      TabIndex        =   54
      Top             =   3780
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   8
      Left            =   4170
      TabIndex        =   53
      Top             =   3780
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   8
      Left            =   1440
      TabIndex        =   52
      Top             =   3780
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   8
      Left            =   810
      TabIndex        =   51
      Top             =   3780
      Width           =   585
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
      Left            =   6195
      TabIndex        =   50
      Top             =   4050
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   9
      Left            =   5430
      TabIndex        =   49
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   9
      Left            =   4800
      TabIndex        =   48
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   9
      Left            =   4170
      TabIndex        =   47
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   9
      Left            =   1440
      TabIndex        =   46
      Top             =   4080
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   9
      Left            =   810
      TabIndex        =   45
      Top             =   4080
      Width           =   585
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
      Left            =   6195
      TabIndex        =   44
      Top             =   4380
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   10
      Left            =   5430
      TabIndex        =   43
      Top             =   4410
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   10
      Left            =   4800
      TabIndex        =   42
      Top             =   4410
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   10
      Left            =   4170
      TabIndex        =   41
      Top             =   4410
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   10
      Left            =   1440
      TabIndex        =   40
      Top             =   4410
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   10
      Left            =   810
      TabIndex        =   39
      Top             =   4410
      Width           =   585
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
      Left            =   6195
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   11
      Left            =   5430
      TabIndex        =   37
      Top             =   4710
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   11
      Left            =   4800
      TabIndex        =   36
      Top             =   4710
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   11
      Left            =   4170
      TabIndex        =   35
      Top             =   4710
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   11
      Left            =   1440
      TabIndex        =   34
      Top             =   4710
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   11
      Left            =   810
      TabIndex        =   33
      Top             =   4710
      Width           =   615
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
      Left            =   6195
      TabIndex        =   32
      Top             =   4980
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   12
      Left            =   5430
      TabIndex        =   31
      Top             =   5010
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   12
      Left            =   4800
      TabIndex        =   30
      Top             =   5010
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   12
      Left            =   4170
      TabIndex        =   29
      Top             =   5010
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   12
      Left            =   1440
      TabIndex        =   28
      Top             =   5010
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   12
      Left            =   810
      TabIndex        =   27
      Top             =   5010
      Width           =   615
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
      Left            =   6195
      TabIndex        =   26
      Top             =   5280
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   13
      Left            =   5430
      TabIndex        =   25
      Top             =   5310
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   13
      Left            =   4800
      TabIndex        =   24
      Top             =   5310
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   13
      Left            =   4170
      TabIndex        =   23
      Top             =   5310
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   13
      Left            =   1440
      TabIndex        =   22
      Top             =   5310
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   13
      Left            =   810
      TabIndex        =   21
      Top             =   5310
      Width           =   615
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
      Left            =   6195
      TabIndex        =   20
      Top             =   5580
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   14
      Left            =   5430
      TabIndex        =   19
      Top             =   5610
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   14
      Left            =   4800
      TabIndex        =   18
      Top             =   5610
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   14
      Left            =   4170
      TabIndex        =   17
      Top             =   5610
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   14
      Left            =   1440
      TabIndex        =   16
      Top             =   5610
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   14
      Left            =   810
      TabIndex        =   15
      Top             =   5610
      Width           =   615
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
      Left            =   6195
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   15
      Left            =   5430
      TabIndex        =   13
      Top             =   5910
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   15
      Left            =   4800
      TabIndex        =   12
      Top             =   5910
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   15
      Left            =   4170
      TabIndex        =   11
      Top             =   5910
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   15
      Left            =   1440
      TabIndex        =   10
      Top             =   5910
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   15
      Left            =   810
      TabIndex        =   9
      Top             =   5910
      Width           =   615
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
      Left            =   6195
      TabIndex        =   8
      Top             =   6180
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
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
      Index           =   16
      Left            =   5430
      TabIndex        =   7
      Top             =   6210
      Width           =   735
   End
   Begin VB.TextBox txtPrecioUnit 
      Alignment       =   1  'Right Justify
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
      Index           =   16
      Left            =   4800
      TabIndex        =   6
      Top             =   6210
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Index           =   16
      Left            =   4170
      TabIndex        =   5
      Top             =   6210
      Width           =   615
   End
   Begin VB.TextBox txtDescripcion 
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
      Index           =   16
      Left            =   1440
      TabIndex        =   4
      Top             =   6210
      Width           =   2715
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   16
      Left            =   810
      TabIndex        =   3
      Top             =   6210
      Width           =   615
   End
   Begin VB.CommandButton btnCancelar 
      Cancel          =   -1  'True
      Caption         =   "Salir (ESC)"
      DisabledPicture =   "Recibos.frx":69230
      DownPicture     =   "Recibos.frx":696F4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   0
      Picture         =   "Recibos.frx":69BE0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7170
      Width           =   945
   End
   Begin VB.CommandButton btnReImprime 
      Caption         =   "ReImprime"
      Height          =   700
      Left            =   0
      Picture         =   "Recibos.frx":6A0CC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6450
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   700
      Left            =   0
      Picture         =   "Recibos.frx":6A5A5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   945
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
      Left            =   6195
      TabIndex        =   127
      Top             =   1500
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "....................................................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2220
      TabIndex        =   126
      Top             =   270
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "................................."
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
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2280
      TabIndex        =   125
      Top             =   450
      Width           =   1245
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "...................................."
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
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5580
      TabIndex        =   124
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..........................."
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
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2520
      TabIndex        =   123
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label LblAnulado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   990
      Left            =   90
      TabIndex        =   122
      Top             =   270
      Visible         =   0   'False
      Width           =   6765
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   4785
      Left            =   6210
      TabIndex        =   121
      Top             =   1710
      Width           =   765
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   4785
      Left            =   30
      TabIndex        =   120
      Top             =   1710
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SubTotal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4305
      TabIndex        =   119
      Top             =   6600
      Width           =   630
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Exonerac."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   118
      Top             =   6900
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Boleta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   117
      Top             =   7230
      Width           =   855
   End
End
Attribute VB_Name = "Recibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
  Dim oRecibos As New Recibos
    Me.txtRazonSocial = oDOComprobantePago.RazonSocial + IIf(IsNull(oDOPaciente.NroHistoriaClinica), "(Particular)", "(HC: " & oDOPaciente.NroHistoriaClinica & ")")
    Me.txtNroComprobante = oDOComprobantePago.nroSerie + " - " + oDOComprobantePago.NroDocumento
    Me.txtFecha = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
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
Dim rsReporte As New Recordset
Dim mo_PermisosFacturacion As New PermisosFacturacion
Dim iFila As Long: Dim Ln As Integer
Dim mdTotal  As Double: Dim lnTotalBoleta As Double
Dim oRecibos As New Recibos
Dim lnImporteEXO As Double: Dim lnSubTotal As Double
Dim lnDctos As Double
Dim oReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcDevolucion As String
        lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio = False
        lnIdUsuario = lnIdUsuarioSistema
        If lcFarmaciaServicio = sighentidades.sghbien Then
           Set rsReporte = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(nroSerie, nroDcto)
        Else
           Set rsReporte = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(nroSerie, nroDcto)
           If rsReporte.RecordCount > 0 Then
                If Not IsNull(rsReporte.Fields!idCuentaAtencion) And rsReporte.Fields!idCuentaAtencion > 0 Then
                   lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio = True
                End If
           End If
        End If
        lnNroItemsBoleta = rsReporte.RecordCount
        If lnNroItemsBoleta > 0 Then
            '
            wxIdTipoComprobanteDefault = IIf(IsNull(rsReporte.Fields!IdTipoComprobante), 3, rsReporte.Fields!IdTipoComprobante)
            CargaSetup_Caja App.Path & "\archivos", wxIdTipoComprobanteDefault
            '
            Set oRsItemsBoleta = rsReporte.Clone
            Select Case rsReporte.Fields!IdEstadoComprobante
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
            lcDevolucion = IIf(rsReporte.Fields!IdTipoPago = 2, "DEVOLUCION", "")
            txtTipoPago.Text = lcDevolucion
            Me.txtRazonSocial = Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "", rsReporte.Fields!RazonSocial)) + IIf(IsNull(rsReporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & rsReporte.Fields!NroHistoriaClinica & ")")
            Me.txtNroComprobante = nroSerie & " - " & nroDcto
            Me.txtFecha = Format(rsReporte.Fields!FechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY)
            Me.txtDctos.Text = "Pago a Cta: " & Format(IIf(IsNull(rsReporte.Fields!Adelantos), 0, rsReporte.Fields!Adelantos), "######.#0")
            lnDctos = rsReporte.Fields!Adelantos
            lnImporteEXO = rsReporte.Fields!Exoneraciones
            If lcFarmaciaServicio = sighentidades.sghbien Then
               If IsNull(rsReporte.Fields!idPreVenta) Then
                   Me.txtDatos = "Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))) & "    PreVenta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!IdOrden), 0, rsReporte.Fields!IdOrden)))
               Else
                   Me.txtDatos = "Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))) & "    PreVenta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idPreVenta), 0, rsReporte.Fields!idPreVenta)))
               End If
            Else
                Me.txtDatos = "Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))) & "    Ord.Pag: " & Trim(Str(IIf(IsNull(rsReporte.Fields!IdOrdenPago), 0, rsReporte.Fields!IdOrdenPago)))
            End If
            Me.txtTotal = Format(rsReporte.Fields!TotalBoleta, "#######.#0")
            lnTotalBoleta = IIf(IsNull(rsReporte.Fields!TotalBoleta), "0", rsReporte.Fields!TotalBoleta)
            txtCajero.Text = oReglasCaja.SeleccionaDatosCajero(rsReporte.Fields!IdCajero, sghIniciales)
            txtServicio.Text = oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))
            iFila = 1
            mdTotal = IIf(IsNull(rsReporte.Fields!TotalBoleta), 0, rsReporte.Fields!TotalBoleta)
            txtLetras.Text = "Son: " + sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(mdTotal)) + " con " + sighentidades.DevuelveSoloDecimales(mdTotal) + "/100   Nuevos Soles"
            lnNroItemsBoleta = 0
            Do While Not rsReporte.EOF
                lnSubTotal = rsReporte.Fields!TotalPorPagar
                Me.txtCodigo(iFila - 1) = rsReporte.Fields!Codigo
                Me.txtDescripcion(iFila - 1) = rsReporte.Fields!NombreProducto
                Me.txtPrecioUnit(iFila - 1) = Format(rsReporte.Fields!PrecioUnitario, "######.#0")
                Me.txtCantidad(iFila - 1) = rsReporte.Fields!Cantidad
                Me.txtImporte(iFila - 1) = Format(lnSubTotal, "######.#0")
                lnNroItemsBoleta = lnNroItemsBoleta + 1
                iFila = iFila + 1
                rsReporte.MoveNext
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
                   For Ln = 0 To lnNroItemsBoleta - 1
                       dev(Ln).Visible = True
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
       oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE lcNroSerie, lcNroBoleta, lnBienFarmacia, True
       
    Else
       oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE lcNroSerie, lcNroBoleta, lnBienFarmacia, False
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
    'CargaSetup_Caja App.Path & "\archivos"
    '
    
End Sub

