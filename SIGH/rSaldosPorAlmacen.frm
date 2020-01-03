VERSION 5.00
Begin VB.Form rSaldosPorAlmacen 
   Caption         =   "Saldos"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "rSaldosPorAlmacen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   6
      Top             =   2250
      Width           =   6360
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rSaldosPorAlmacen.frx":0CCA
         DownPicture     =   "rSaldosPorAlmacen.frx":118E
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
         Left            =   3278
         Picture         =   "rSaldosPorAlmacen.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rSaldosPorAlmacen.frx":1B66
         DownPicture     =   "rSaldosPorAlmacen.frx":1FC6
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
         Left            =   1748
         Picture         =   "rSaldosPorAlmacen.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
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
      Height          =   2205
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6345
      Begin VB.CheckBox chkTodasFarmacias 
         Caption         =   "Todos las Farmacias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox chkExcel 
         Caption         =   "En Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         Picture         =   "rSaldosPorAlmacen.frx":28B0
         TabIndex        =   9
         Top             =   690
         Width           =   1125
      End
      Begin VB.CheckBox chkStkMinimo 
         Caption         =   "Sólo muestra Productos con Saldos menores a su STOCK MINIMO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   1590
         Width           =   5835
      End
      Begin VB.ComboBox cmbAlmacen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2130
         TabIndex        =   0
         Top             =   240
         Width           =   4080
      End
      Begin VB.ComboBox cmbOrden 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "rSaldosPorAlmacen.frx":2BC2
         Left            =   3450
         List            =   "rSaldosPorAlmacen.frx":2BCC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   2745
      End
      Begin VB.CheckBox chkLotes 
         Caption         =   "Se muestra Lotes/F.Vencimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   2985
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
   End
End
Attribute VB_Name = "rSaldosPorAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************daniel barrantes**************
'***************Registro de datos de filtro para el Reporte
'***************Historias Clinicas que pasadas 24 horas no regresan al ARCHIVO
Dim mo_cmbAlmacen As New sighcomun.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ms_MensajeError As String
Dim mo_Teclado As New sighcomun.Teclado
Dim ml_TextoDelFiltro As String
Dim ml_idUsuario As Long
Dim mo_Formulario As New sighcomun.Formulario

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
         Dim oRptClase As New rCrystal
         oRptClase.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
         oRptClase.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
         oRptClase.OrdenadoPor = cmbOrden.ListIndex
         oRptClase.TextoDelFiltro = ml_TextoDelFiltro
         oRptClase.SeMuestraLotes = IIf(chkLotes.Value = 1, True, False)
         oRptClase.StockMinimoMayorAcantidad = IIf(chkStkMinimo.Value = 1, True, False)
         oRptClase.TipoReporte = Me.Name
         oRptClase.Show vbModal
         Set oRptClase = Nothing
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    If chkTodasFarmacias.Value = 1 Then
        ml_TextoDelFiltro = "FILTROS:   Almacén: (Todos)    orden: (" & cmbOrden.Text & ")    " & IIf(chkStkMinimo.Value = 1, "(" & chkStkMinimo.Caption & ")", "")
        mo_cmbAlmacen.BoundText = ""
    Else
        ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")    orden: (" & cmbOrden.Text & ")    " & IIf(chkStkMinimo.Value = 1, "(" & chkStkMinimo.Caption & ")", "")
        If mo_cmbAlmacen.BoundText = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén"
            cmbAlmacen.SetFocus
        End If
    End If
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub








Private Sub chkTodasFarmacias_Click()
   If chkTodasFarmacias.Value = 1 Then
      cmbAlmacen.Visible = False
      chkLotes.Visible = False
      chkStkMinimo.Visible = False
   Else
      cmbAlmacen.Visible = True
      chkLotes.Visible = True
      chkStkMinimo.Visible = True
   End If
End Sub

Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub



Private Sub cmbOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbOrden

End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
End Sub


Private Sub Form_Load()
    cmbOrden.ListIndex = 1
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    '
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacen, False
    End If
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
'           ucListaProductos1.RealizarBusqueda
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbAlmacen = Nothing
    Set mo_Formulario = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub
