VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form rRecetasXusuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recetas por usuario del Sistema"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rRecetasXusuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
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
      Height          =   2445
      Left            =   30
      TabIndex        =   7
      Top             =   45
      Width           =   6780
      Begin VB.CheckBox chkSoloBoletas 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo DOCUMENTOS emitidos por PREVENTAS"
         Height          =   255
         Left            =   2475
         TabIndex        =   17
         Top             =   1815
         Width           =   4200
      End
      Begin VB.ComboBox cmbTipoFinanciamiento 
         Height          =   330
         Left            =   1500
         TabIndex        =   15
         Text            =   "cmbTipoFinanciamiento"
         Top             =   1440
         Width           =   5220
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
         Height          =   315
         Left            =   90
         Picture         =   "rRecetasXusuario.frx":0CCA
         TabIndex        =   14
         Top             =   1800
         Width           =   1605
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   1500
         TabIndex        =   0
         Top             =   240
         Width           =   5220
      End
      Begin VB.ComboBox cmbUsuario 
         Height          =   330
         ItemData        =   "rRecetasXusuario.frx":0FDC
         Left            =   1500
         List            =   "rRecetasXusuario.frx":0FE6
         TabIndex        =   3
         Text            =   "cmbUsuario"
         Top             =   1020
         Width           =   5205
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   630
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   4575
         TabIndex        =   2
         Top             =   630
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2940
         TabIndex        =   12
         Top             =   630
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   5955
         TabIndex        =   13
         Top             =   630
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Producto/Plan"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   1470
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   210
         Left            =   4065
         TabIndex        =   11
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor Farm"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1260
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
      Left            =   15
      TabIndex        =   5
      Top             =   2490
      Width           =   6765
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rRecetasXusuario.frx":1002
         DownPicture     =   "rRecetasXusuario.frx":1462
         Height          =   700
         Left            =   1973
         Picture         =   "rRecetasXusuario.frx":18D7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rRecetasXusuario.frx":1D4C
         DownPicture     =   "rRecetasXusuario.frx":2210
         Height          =   700
         Left            =   3503
         Picture         =   "rRecetasXusuario.frx":26FC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "rRecetasXusuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Recetas por Usuario
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbAlmacen As New sighentidades.ListaDespleglable
Dim mo_cmbUsuario As New sighentidades.ListaDespleglable
Dim mo_cmbTipoFinanciamiento As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Const ml_IdPuntoCarga As Integer = 5
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim rsIdAlmacen As Recordset
Dim ml_idUsuario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
            oRptClaseCry.FechaInicio = CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.FechaFin = CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.idUsuario = Val(mo_cmbUsuario.BoundText)
            oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro & IIf(Me.chkSoloBoletas.Value = 1, " (" & chkSoloBoletas.Caption & ")", "")
            oRptClaseCry.IdTipoFinanciamiento = Val(mo_cmbTipoFinanciamiento.BoundText)
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.SoloBoletas = IIf(Me.chkSoloBoletas.Value = 1, True, False)
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")      F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & IIf(cmbUsuario.Text <> "", ")     Vendedor: " & Trim(cmbUsuario.Text), "") & IIf(Val(mo_cmbTipoFinanciamiento.BoundText) > 0, "      (Producto/plan: " & Trim(cmbTipoFinanciamiento.Text) & ")", "")
    If mo_cmbAlmacen.BoundText = "" Then
        sMensaje = sMensaje + "Por favor elija el Almacén" + Chr(13)
        cmbAlmacen.SetFocus
    End If
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub







Private Sub chkSoloBoletas_Click()
    If Me.chkSoloBoletas.Value = 1 Then
       cmbTipoFinanciamiento.Text = ""
    End If
    
End Sub

Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub








Private Sub cmbTipoFinanciamiento_Click()
         chkSoloBoletas.Value = 0

End Sub

Private Sub cmbUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
      mo_Teclado.RealizarNavegacion KeyCode, cmbUsuario
End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
    Set mo_cmbUsuario.MiComboBox = cmbUsuario
    Set mo_cmbTipoFinanciamiento.MiComboBox = cmbTipoFinanciamiento
End Sub

Sub InicializaFechaHora()
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    txtHrInicio.Text = lcBuscaParametro.SeleccionaFilaParametro(267)
    txtHrFin.Text = lcBuscaParametro.SeleccionaFilaParametro(268)

End Sub
Private Sub Form_Load()
    InicializaFechaHora
    
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idtipoSuministro='01'")
    '
    mo_cmbUsuario.BoundColumn = "IdEmpleado"
    mo_cmbUsuario.ListField = "DEmpleado"
    Set mo_cmbUsuario.RowSource = mo_ReglasComunes.EmpleadosSeleccionarTodos
    '
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_cmbUsuario.BoundText = ml_idUsuario
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacen, False
       mo_Formulario.HabilitarDeshabilitar Me.cmbUsuario, False
    End If
    
    mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
    mo_cmbTipoFinanciamiento.ListField = "Descripcion"
    Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia("")
    
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub






Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.esfecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.esfecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbAlmacen = Nothing
    Set mo_cmbUsuario = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
If Not sighentidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
If Not sighentidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub
