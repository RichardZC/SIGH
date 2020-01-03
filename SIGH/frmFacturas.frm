VERSION 5.00
Begin VB.Form frmFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Factura"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12585
   Icon            =   "frmFacturas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFacturas.frx":0CCA
   ScaleHeight     =   9450
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton btnReImprime 
      Caption         =   "ReImprime Factura"
      Height          =   700
      Left            =   30
      Picture         =   "frmFacturas.frx":E920
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   7950
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtDDD 
      Alignment       =   2  'Center
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
      Left            =   4680
      TabIndex        =   78
      Top             =   8220
      Width           =   375
   End
   Begin VB.TextBox txtMMM 
      Alignment       =   2  'Center
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
      Left            =   5520
      TabIndex        =   77
      Top             =   8220
      Width           =   1335
   End
   Begin VB.TextBox txtAAA 
      Alignment       =   2  'Center
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
      Left            =   7440
      TabIndex        =   76
      Top             =   8220
      Width           =   375
   End
   Begin VB.TextBox txtAA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   11160
      TabIndex        =   75
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtMM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   10080
      TabIndex        =   74
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtDD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   9120
      TabIndex        =   73
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtRUC 
      Appearance      =   0  'Flat
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
      Left            =   1800
      TabIndex        =   72
      Top             =   3240
      Width           =   2505
   End
   Begin VB.TextBox txtIGV 
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   8160
      Width           =   1665
   End
   Begin VB.CommandButton btnCancelar 
      Cancel          =   -1  'True
      Caption         =   "Salir (ESC)"
      DisabledPicture =   "frmFacturas.frx":EDF9
      DownPicture     =   "frmFacturas.frx":F2BD
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
      Left            =   30
      Picture         =   "frmFacturas.frx":F7A9
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   8760
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualiza Devolución"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   63
      Top             =   8880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.TextBox txtLetras 
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
      Left            =   1680
      TabIndex        =   62
      Top             =   7350
      Width           =   9135
   End
   Begin VB.CommandButton cmdDev 
      Caption         =   "Actualiza Devolución"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   225
      TabIndex        =   8
      Top             =   8805
      Visible         =   0   'False
      Width           =   840
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
      Left            =   11595
      TabIndex        =   0
      Top             =   4500
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
      Left            =   11595
      TabIndex        =   1
      Top             =   4860
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
      Left            =   11595
      TabIndex        =   2
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
      Index           =   3
      Left            =   11595
      TabIndex        =   3
      Top             =   5565
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
      Left            =   11595
      TabIndex        =   4
      Top             =   5940
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
      Left            =   11595
      TabIndex        =   5
      Top             =   6285
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
      Left            =   11595
      TabIndex        =   6
      Top             =   6660
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
      Index           =   7
      Left            =   11595
      TabIndex        =   7
      Top             =   7020
      Visible         =   0   'False
      Width           =   705
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
      Left            =   1320
      TabIndex        =   60
      Top             =   9120
      Width           =   375
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
      Left            =   120
      TabIndex        =   59
      Top             =   270
      Width           =   2145
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
      Left            =   2880
      TabIndex        =   26
      Top             =   5580
      Width           =   6075
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
      Left            =   1920
      TabIndex        =   57
      Top             =   9120
      Width           =   3270
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
      Left            =   4560
      TabIndex        =   56
      Top             =   9120
      Width           =   2295
   End
   Begin VB.TextBox txtNroComprobante 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   8280
      TabIndex        =   55
      Top             =   1680
      Width           =   2205
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   8520
      Width           =   1665
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
      Left            =   8130
      TabIndex        =   53
      Top             =   9180
      Width           =   1065
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   7800
      Width           =   1665
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
      Left            =   10590
      TabIndex        =   51
      Top             =   7080
      Width           =   1095
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
      Left            =   10590
      TabIndex        =   50
      Top             =   6660
      Width           =   1095
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
      Left            =   10590
      TabIndex        =   49
      Top             =   6285
      Width           =   1095
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
      Left            =   10590
      TabIndex        =   48
      Top             =   5940
      Width           =   1095
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
      Left            =   10590
      TabIndex        =   47
      Top             =   5580
      Width           =   1095
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
      Left            =   10590
      TabIndex        =   46
      Top             =   5220
      Width           =   1095
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
      Left            =   10590
      TabIndex        =   45
      Top             =   4860
      Width           =   1095
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
      Left            =   9240
      TabIndex        =   44
      Top             =   7080
      Width           =   855
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
      Left            =   9240
      TabIndex        =   43
      Top             =   6660
      Width           =   855
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
      Left            =   9240
      TabIndex        =   42
      Top             =   6300
      Width           =   855
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
      Left            =   9240
      TabIndex        =   41
      Top             =   5940
      Width           =   855
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
      Left            =   9240
      TabIndex        =   40
      Top             =   5580
      Width           =   855
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
      Left            =   9240
      TabIndex        =   39
      Top             =   5220
      Width           =   855
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
      Left            =   9240
      TabIndex        =   38
      Top             =   4860
      Width           =   855
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   37
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   36
      Top             =   6660
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   35
      Top             =   6300
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   34
      Top             =   5940
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   33
      Top             =   5580
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   32
      Top             =   5220
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   31
      Top             =   4860
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
      Index           =   7
      Left            =   2880
      TabIndex        =   30
      Top             =   7080
      Width           =   6075
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
      Left            =   2880
      TabIndex        =   29
      Top             =   6660
      Width           =   6075
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
      Left            =   2910
      TabIndex        =   28
      Top             =   6300
      Width           =   6075
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
      Left            =   2880
      TabIndex        =   27
      Top             =   5940
      Width           =   6075
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
      Left            =   2865
      TabIndex        =   25
      Top             =   5220
      Width           =   6075
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
      Left            =   2880
      TabIndex        =   24
      Top             =   4860
      Width           =   6075
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   23
      Top             =   7080
      Width           =   705
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   22
      Top             =   6660
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   21
      Top             =   6300
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   20
      Top             =   5940
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   19
      Top             =   5580
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   18
      Top             =   5205
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   17
      Top             =   4860
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
      Index           =   0
      Left            =   10590
      TabIndex        =   16
      Top             =   4500
      Width           =   1095
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
      Left            =   9240
      TabIndex        =   15
      Top             =   4500
      Width           =   855
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
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
      Left            =   1290
      TabIndex        =   14
      Top             =   4500
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2130
      TabIndex        =   13
      Top             =   4500
      Width           =   735
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
      Left            =   2880
      TabIndex        =   12
      Top             =   4500
      Width           =   6075
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
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
      Left            =   9120
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   6720
      TabIndex        =   10
      Top             =   3240
      Width           =   2205
   End
   Begin VB.TextBox txtRazonSocial 
      Appearance      =   0  'Flat
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
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   7065
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   2865
      Left            =   360
      TabIndex        =   70
      Top             =   4440
      Width           =   735
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
      Left            =   7320
      TabIndex        =   68
      Top             =   9180
      Width           =   735
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
      Left            =   5400
      TabIndex        =   58
      Top             =   240
      Visible         =   0   'False
      Width           =   6765
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
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   540
      Width           =   2145
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Left            =   2070
      TabIndex        =   66
      Top             =   0
      Width           =   3375
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
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   990
      Width           =   2145
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "....................................................."
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
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   750
      Width           =   2145
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
      Left            =   11595
      TabIndex        =   61
      Top             =   4260
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Muestra una Factura
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
Dim lcNroBoleta As String, lcNroSerie As String, lnBienFarmacia As Long
Dim ml_lbTienePermisoReimprimeBoleta As Boolean

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

Sub Imprimir(oDOPaciente As doPaciente, oDOAtencion As DOAtencion, oDOComprobantePago As DOCajaComprobantesPago, rsFacturacionProductos As Recordset)
Dim rsReporte As New Recordset
Dim rsReporte1 As New Recordset
Dim iFila As Long
Dim mdTotal  As Double
Dim oRecibos As New RecibosBoleta

    
    txtRazonSocial = oDOComprobantePago.RazonSocial + IIf(IsNull(oDOPaciente.NroHistoriaClinica), "(Particular)", "(HC: " & oDOPaciente.NroHistoriaClinica & ")")
    txtNroComprobante = oDOComprobantePago.nroSerie + " - " + oDOComprobantePago.NroDocumento
    txtRuc = oDOComprobantePago.Ruc
    txtFecha = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
    txtDctos.Text = "Adelantos: " & Format(oDOComprobantePago.Dctos, "0.00")
    txtDatos = "Cta: " & Trim(Str(oDOComprobantePago.idCuentaAtencion)) & "    Ord.Pag: " & Trim(Str(rsFacturacionProductos!IdOrden))
    iFila = 1
    mdTotal = 0
    Do While Not rsFacturacionProductos.EOF
    
        txtCodigo(iFila - 1) = rsFacturacionProductos!Codigo
        txtDescripcion(iFila - 1) = rsFacturacionProductos!NombreProducto
        txtPrecioUnit(iFila - 1) = Format(rsFacturacionProductos!PrecioUnitario, "0.00")
        txtCantidad(iFila - 1) = rsFacturacionProductos!Cantidad
        mdTotal = mdTotal + rsFacturacionProductos!TotalPorPagar
        txtImporte(iFila - 1) = Format(rsFacturacionProductos!TotalPorPagar, "0.00")

        rsFacturacionProductos.MoveNext
        
        If iFila = 8 Then
             Exit Do
        End If
        
        iFila = iFila + 1
    Loop
    txtTotal = Format(mdTotal, "0.00")
        
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


'*************** Adams BONILLA MAGALLANES ***************
'********* Impresion en PANTALLA de una FACTURA *********
Sub ImprimirFactura(nroSerie As String, nroDcto As String, lcFarmaciaServicio As Long, lnIdUsuarioSistema As Long)
  Dim rsReporte As New Recordset
  Dim mo_PermisosFacturacion As New PermisosFacturacion
  Dim iFila As Long: Dim ln As Integer
  Dim mdTotal  As Double: Dim lnTotalBoleta As Double
  Dim oRecibos As New RecibosBoleta
  Dim lnImporteEXO As Double: Dim lnSubTotal As Double
  Dim lnDctos As Double
  Dim oReglasCaja As New SIGHNegocios.ReglasCaja
  Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
  Dim lcDevolucion As String
  Dim IGV As Double
  Dim lnVigv As Double, lnVsubTotal As Double
  
  IGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221)) / 100
  lnIdUsuario = lnIdUsuarioSistema
  lnFarmaciaServicio = lcFarmaciaServicio
  If lcFarmaciaServicio = sighEntidades.sghbien Then
    Set rsReporte = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(nroSerie, nroDcto)
  Else
    Set rsReporte = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(nroSerie, nroDcto)
  End If
  lnNroItemsBoleta = rsReporte.RecordCount
  If lnNroItemsBoleta > 0 Then
    lnVigv = IIf(IsNull(rsReporte.Fields!IGV), 0, rsReporte.Fields!IGV)
    lnVsubTotal = rsReporte.Fields!Subtotal
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
    txtRazonSocial = Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "", rsReporte.Fields!RazonSocial)) + IIf(IsNull(rsReporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & rsReporte.Fields!NroHistoriaClinica & ")")
    txtNroComprobante = nroSerie & " - " & nroDcto
    txtRuc = IIf(IsNull(rsReporte!Ruc), "", rsReporte!Ruc)
    txtFecha = Format(rsReporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY)
    txtDctos.Text = "Pago a Cta: " & Format(IIf(IsNull(rsReporte.Fields!Dctos), 0, rsReporte.Fields!Dctos), "0.00")
    lnDctos = rsReporte.Fields!Dctos
    lnImporteEXO = rsReporte.Fields!Exoneraciones
    If lcFarmaciaServicio = sighEntidades.sghbien Then
      If IsNull(rsReporte.Fields!idPreVenta) Then
        Me.txtDatos = "Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))) & "    PreVenta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!IdOrden), 0, rsReporte.Fields!IdOrden)))
      Else
        Me.txtDatos = "Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))) & "    PreVenta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idPreVenta), 0, rsReporte.Fields!idPreVenta)))
      End If
    Else
      Me.txtDatos = "Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))) & "    Ord.Pag: " & Trim(Str(IIf(IsNull(rsReporte.Fields!IdOrdenPago), 0, rsReporte.Fields!IdOrdenPago)))
    End If
    Me.txtTotal = Format(rsReporte.Fields!TotalBoleta, "0.00")
    lnTotalBoleta = IIf(IsNull(rsReporte.Fields!TotalBoleta), "0", rsReporte.Fields!TotalBoleta)
    txtCajero.Text = oReglasCaja.SeleccionaDatosCajero(rsReporte.Fields!IdCajero, sghIniciales)
    txtServicio.Text = oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion))
    iFila = 1
    mdTotal = IIf(IsNull(rsReporte.Fields!TotalBoleta), 0, rsReporte.Fields!TotalBoleta)
    txtLetras.Text = sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(mdTotal)) + " con " + sighEntidades.DevuelveSoloDecimales(mdTotal) + "/100 "
    lnNroItemsBoleta = 0
    Do While Not rsReporte.EOF
      'If rsReporte.Fields!cantidad > 0 And rsReporte.Fields!precioUnitario > 0 Then
      lnSubTotal = rsReporte.Fields!TotalPorPagar
      Me.txtCodigo(iFila - 1) = rsReporte.Fields!Codigo
      Me.txtDescripcion(iFila - 1) = rsReporte.Fields!NombreProducto
      Me.txtPrecioUnit(iFila - 1) = Format(rsReporte.Fields!PrecioUnitario, "0.00")
      Me.txtCantidad(iFila - 1) = rsReporte.Fields!Cantidad
      Me.txtImporte(iFila - 1) = Format(lnSubTotal, "0.00")
      lnNroItemsBoleta = lnNroItemsBoleta + 1
      iFila = iFila + 1
      'End If
      rsReporte.MoveNext
      If iFila > 17 Then Exit Do
    Loop
    txtExonerado.Text = Format(lnImporteEXO, "0.00")
    If txtTipoPago.Text = "" And lnTotalBoleta = 0 And lnImporteEXO > 0 Then txtTipoPago.Text = "EXONERADO TOTAL"
    txtSubTotal.Text = Format(mdTotal + lnImporteEXO + lnDctos, "0.00")
    Me.txtSubTotal = Format(lnVsubTotal, "0.00")     'Format(mdTotal / (IGV + 1), "0.00")
    Me.txtIGV = Format(lnVigv, "0.00")               'Format(mdTotal * IGV / (IGV + 1), "0.00")
    'Autorizado para Devoluciones
    If LblAnulado.Visible = False And lnTotalBoleta > 0 And lcDevolucion = "" Then
      Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(lnIdUsuarioSistema)
      If mo_PermisosFacturacion.AutorizarDevoluciones Then
        For ln = 0 To lnNroItemsBoleta - 1
          dev(ln).Visible = True
        Next
        cmdDev.Visible = True
        lblDev.Visible = True
      End If
    End If
  End If
  Exit Sub

ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Emisión de Factura"
    Case Else
      MsgBox Err.Description
  End Select
  Exit Sub
End Sub


Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnReImprime_Click()
    If MsgBox("Por favor confirmar, ¿Realmente desea REIMPRIMIR FACTURA ?", vbQuestion + vbYesNo, "Estado de Cuenta") = vbNo Then
        Exit Sub
    End If
    Dim oImprimeFacturaDOS As New RptCaja
    oImprimeFacturaDOS.ImpresionFacturaEnDosTYPE lcNroSerie, lcNroBoleta, lnBienFarmacia
    Set oImprimeFacturaDOS = Nothing

End Sub


Private Sub Command1_Click()
    Set Me.Picture = LoadPicture("")
    Me.PrintForm
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
      
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtFecha_Change()
  Dim D As Date
  D = CDate(txtFecha)
  txtDD.Text = Format(D, "dd")
  txtMM.Text = Format(D, "mm")
  txtAA.Text = Format(D, "yyyy")
  txtDDD.Text = Format(D, "dd")
  txtMMM.Text = Format(D, "mmmm")
  txtAAA.Text = Right(Format(D, "yyyy"), 2)
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set oRsItemsBoleta = Nothing
    Set lcBuscaParametro = Nothing
End Sub
