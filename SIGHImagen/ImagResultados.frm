VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ImagResultados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resultados"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "ImagResultados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Informe del Resultado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6345
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   8415
      Begin VB.TextBox txtResultadoFinal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6000
         Left            =   90
         MaxLength       =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   8235
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   0
      TabIndex        =   7
      Top             =   30
      Width           =   8400
      Begin VB.CommandButton cmdBuscaImg 
         Height          =   330
         Left            =   8040
         Picture         =   "ImagResultados.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Muestra una o varias imágenes"
         Top             =   960
         Width           =   300
      End
      Begin VB.TextBox txtRutaImg 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1635
         MaxLength       =   200
         TabIndex        =   13
         Top             =   960
         Width           =   6375
      End
      Begin VB.ComboBox cmbResponsable 
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
         Left            =   4470
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   390
         Width           =   3870
      End
      Begin VB.TextBox txtPaciente 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   75
         TabIndex        =   12
         Top             =   405
         Width           =   4350
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   960
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Paciente"
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
         Left            =   105
         TabIndex        =   11
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "F.Resultado           Ruta de Imágenes"
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
         Left            =   105
         TabIndex        =   9
         Top             =   735
         Width           =   3105
      End
      Begin VB.Label Label7 
         Caption         =   "Realiza Prueba"
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
         Left            =   4485
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   0
      TabIndex        =   3
      Top             =   7800
      Width           =   8415
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         Picture         =   "ImagResultados.frx":1254
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ImagResultados.frx":172D
         DownPicture     =   "ImagResultados.frx":1B8D
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
         Left            =   3210
         Picture         =   "ImagResultados.frx":2002
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   165
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ImagResultados.frx":2477
         DownPicture     =   "ImagResultados.frx":293B
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
         Left            =   4635
         Picture         =   "ImagResultados.frx":2E27
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   165
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ImagResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Ecografía General
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lbEsAgregar As Boolean
Dim ms_Paciente As String
Dim ml_PuntoCarga As Long
Dim oRsResultados As New Recordset
Dim ml_idMovimiento As Long
Dim ml_idProductoCpt As Long
Dim mb_EsResultadoAutomatico As Boolean
Dim ml_Producto As String
Dim mb_SoloEsConsulta As Boolean
Property Let SoloEsConsulta(lValue As Boolean)
    mb_SoloEsConsulta = lValue
End Property

Property Let Producto(lValue As String)
    ml_Producto = lValue
End Property
Property Let EsResultadoAutomatico(lValue As Boolean)
    mb_EsResultadoAutomatico = lValue
End Property


Property Let idProductoCpt(lValue As Long)
    ml_idProductoCpt = lValue
End Property
Property Let idMovimiento(lValue As Long)
    ml_idMovimiento = lValue
End Property

Property Set RsResultados(lValue As Recordset)
    Set oRsResultados = lValue
End Property

Property Let Paciente(lValue As String)
    ms_Paciente = lValue
End Property
Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
End Property


Private Sub btnAceptar_Click()
  On Error GoTo ErrGr
  If ValidaDatosObligatorios Then
    Dim oDoImagMovimientoResultados As New DoImagMovimientoResultados
    Dim oImagMovimientoResultados As New ImagMovimientoResultados
    Dim oConexion As New Connection
    Dim lcMensaje As String
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oImagMovimientoResultados.Conexion = oConexion
    
    If mo_ReglasImagenes.ResultadosImagenesActualizar(oDoImagMovimientoResultados, oImagMovimientoResultados, _
                         lbEsAgregar, txtRutaImg.Text, ml_idMovimiento, ml_idProductoCpt, txtResultadoFinal.Text, _
                         mo_reglasComunes.EmpleadosDevuelveDNI(Val(mo_cmbResponsable.BoundText)), CDate(txtFresultado.Text)) = False Then
       GoTo ErrGr
    End If
'    With oDoImagMovimientoResultados
'         .EquipoRuta = txtRutaImg.Text
'         .idMovimiento = ml_idMovimiento
'         .idProductoCpt = ml_idProductoCpt
'         .IdUsuarioAuditoria = SIGHEntidades.Usuario
'         .Resultado = txtResultadoFinal.Text
'         .ResultadoDNI = mo_reglasComunes.EmpleadosDevuelveDNI(Val(mo_cmbResponsable.BoundText))
'         .ResultadoFecha = CDate(txtFresultado.Text)
'    End With
'    If lbEsAgregar = True Then
'       If oImagMovimientoResultados.Insertar(oDoImagMovimientoResultados) = False Then
'          lcMensaje = oImagMovimientoResultados.MensajeError: GoTo ErrGr
'       End If
'    Else
'       If oImagMovimientoResultados.Modificar(oDoImagMovimientoResultados) = False Then
'          lcMensaje = oImagMovimientoResultados.MensajeError: GoTo ErrGr
'       End If
'    End If

    oConexion.Close
    Set oConexion = Nothing
    Set oDoImagMovimientoResultados = Nothing
    Set oImagMovimientoResultados = Nothing
    Me.Visible = False
  End If
  Exit Sub
ErrGr:
   MsgBox mo_ReglasImagenes.MensajeError
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub btnImprimir_Click()
    Dim oRep As New RayosX
    oRep.ImpresionDelResultado ml_idMovimiento, Me.cmbResponsable.Text, txtFresultado.Text, ml_idProductoCpt
    Set oRep = Nothing

End Sub

Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbResponsable
    AdministrarKeyPreview KeyCode
End Sub

Sub CargaDataCombos()
    '
    Set mo_cmbResponsable.MiComboBox = cmbResponsable
    mo_cmbResponsable.BoundColumn = "idEmpleado"
    mo_cmbResponsable.ListField = "ApNom"
    Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =" & mo_ReglasFarmacia.EmpleadosDevuelveIdCargoSegunPuntoCarga(ml_PuntoCarga))
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub cmdBuscaImg_Click()
     On Error GoTo errcmb
     If Me.txtRutaImg.Text <> "" Then
        Dim oShell As New sighentidades.Shell
        Dim lcArchivoImg As String
        lcArchivoImg = lcBuscaParametro.SeleccionaFilaParametro(567)
        If lcArchivoImg = "S" Then
            '********* visor Radiam Dicom
            lcArchivoImg = sighentidades.SoftwareImagen & " " & txtRutaImg.Text
            oShell.ejecutarComando lcArchivoImg
        ElseIf lcArchivoImg = "W" Then
            '********* internet explorer o google
            oShell.CargarRutaWeb txtRutaImg.Text, Me.hwnd
        End If
      End If
errcmb:
     Set oShell = Nothing
End Sub

Private Sub Form_Load()
    mo_Formulario.HabilitarDeshabilitar txtPaciente, False
    Me.Caption = ml_Producto
    CargaDataCombos
    '
    txtPaciente.Text = ms_Paciente
    txtFresultado.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    
    oRsResultados.Filter = "idProductoCpt=" & ml_idProductoCpt
    If oRsResultados.RecordCount = 0 Then
       lbEsAgregar = True
    Else
       mo_cmbResponsable.BoundText = mo_ReglasFarmacia.EmpleadosDevuelveId(oRsResultados!ResultadoDNI)
       txtFresultado.Text = Format(oRsResultados!ResultadoFecha, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
       txtRutaImg.Text = IIf(IsNull(oRsResultados!EquipoRuta), "", oRsResultados!EquipoRuta)
       txtResultadoFinal.Text = oRsResultados!Resultado
       lbEsAgregar = False
       CargaGridConArchivosImagenes
       If mb_EsResultadoAutomatico = True Then
          Me.Caption = Me.Caption & " (RESULTADO AUTOMATICO)"
          btnAceptar.Visible = False
       End If
       btnImprimir.Visible = True
    End If
    If mb_SoloEsConsulta = True Then
       btnAceptar.Visible = False
       Frame1.Enabled = True
       fraDatosAtencion.Enabled = False
       btnImprimir.Visible = True
       mo_Formulario.HabilitarDeshabilitar cmbResponsable, False
       mo_Formulario.HabilitarDeshabilitar txtFresultado, False
       mo_Formulario.HabilitarDeshabilitar txtRutaImg, False
    End If
End Sub

Sub CargaGridConArchivosImagenes()

End Sub

Private Sub txtFresultado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFresultado
    AdministrarKeyPreview KeyCode
End Sub


Function ValidaDatosObligatorios() As Boolean
   Dim ms_mensaje As String
   ValidaDatosObligatorios = False
   ms_mensaje = ""
   If cmbResponsable.Text = "" Then
      ms_mensaje = ms_mensaje + "Debe elejir al Responsable que realiza la prueba" + Chr(13)
   End If
   If txtResultadoFinal.Text = "" Then
      ms_mensaje = ms_mensaje + "Debe registrar el INFORME DEL RESULTADO" + Chr(13)
   End If
   If ms_mensaje <> "" Then
      MsgBox ms_mensaje, vbInformation, ""
   Else
      ValidaDatosObligatorios = True
   End If
End Function
Private Sub txtResultadoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
