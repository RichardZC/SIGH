VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FarmVerStock 
   Caption         =   "Stock del Medicamento/Insumo"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "farmVerStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   0
      TabIndex        =   2
      Top             =   45
      Width           =   9915
      Begin UltraGrid.SSUltraGrid grdSaldoItem 
         Height          =   3240
         Left            =   45
         TabIndex        =   4
         Top             =   675
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   5715
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "farmVerStock.frx":0CCA
         Caption         =   ".."
      End
      Begin VB.Label lblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "xxxxxx"
         Height          =   210
         Left            =   8760
         TabIndex        =   5
         Top             =   4065
         Width           =   540
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         Caption         =   "xxxxxx"
         Height          =   210
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   9930
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "farmVerStock.frx":0D06
         DownPicture     =   "farmVerStock.frx":11CA
         Height          =   700
         Left            =   4298
         Picture         =   "farmVerStock.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FarmVerStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte del saldos de un Item por cada Almacen/Farmacia
'        Programado por: Barrantes D
'        Fecha: Mayo 2017
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ml_Producto As String
Dim ml_codigo As String
Dim mo_ReglasFarmacia As New ReglasFarmacia
Dim oRsStockXitem As New Recordset

Property Let codigo(lValue As String)
   ml_codigo = lValue
End Property

Property Let Producto(lValue As String)
   ml_Producto = lValue
End Property


Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub




Private Sub Form_Load()
    Dim lnTotal As Long
    lblProducto.Caption = ml_Producto
    grdSaldoItem.Visible = True
    Set oRsStockXitem = mo_ReglasFarmacia.SaldoDetalladoPorItemSeleccionarPorCodigo(ml_codigo)
    lnTotal = 0
    If oRsStockXitem.RecordCount > 0 Then
       oRsStockXitem.MoveFirst
       Do While Not oRsStockXitem.EOF
            lnTotal = lnTotal + oRsStockXitem!cantidad
            oRsStockXitem.MoveNext
       Loop
       oRsStockXitem.MoveFirst
    End If
    lblStock.Caption = "Stock actual disponible: " & Trim(Str(lnTotal))
    Set Me.grdSaldoItem.DataSource = oRsStockXitem
    mo_Apariencia.ConfigurarFilasBiColores Me.grdSaldoItem, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub grdSaldoItem_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdSaldoItem.Bands(0).Columns("Codigo").Hidden = True
    grdSaldoItem.Bands(0).Columns("Almacen").Width = 7900
    grdSaldoItem.Bands(0).Columns("Cantidad").Width = 1000
    grdSaldoItem.Bands(0).Columns("Cantidad").Header.Caption = "Stock"
    grdSaldoItem.Bands(0).Columns("PrecPond").Hidden = True
    grdSaldoItem.Bands(0).Columns("Importe").Hidden = True
    
End Sub
