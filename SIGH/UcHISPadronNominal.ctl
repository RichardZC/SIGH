VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcHISPadronNominal 
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16065
   LockControls    =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   16065
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   15855
      Begin VB.CommandButton bntReporte 
         Height          =   795
         Left            =   12855
         Picture         =   "UcHISPadronNominal.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   885
      End
      Begin VB.TextBox txtNombres 
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
         Left            =   8640
         MaxLength       =   30
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtDNI 
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
         Left            =   240
         MaxLength       =   8
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   1920
         MaxLength       =   9
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   11385
         Picture         =   "UcHISPadronNominal.ctx":04D9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   11385
         Picture         =   "UcHISPadronNominal.ctx":30B5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombres"
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
         Left            =   8640
         TabIndex        =   11
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label Label6 
         Caption         =   "DNI                       N° Historia             "
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
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label50 
         Caption         =   "Ap.Paterno                           Ap.Materno"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   4785
      End
   End
   Begin UltraGrid.SSUltraGrid grdPadronNominal 
      Height          =   8400
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   14817
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de Registro Pádron Nominal"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Padrón Nominal "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "UcHISPadronNominal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de padrón nominal ingresado
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_LoginPC As String
'Dim mo_cmbEstablecimiento As New sighentidades.ListaDespleglable
'Dim mo_cmbServicio As New sighentidades.ListaDespleglable
'Dim mo_cmbMes As New sighentidades.ListaDespleglable
'Dim mo_Apariencia As New sighentidades.GridInfragistic

Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_idRegistroSeleccionado As Long
Dim lblNombre As String
Dim ml_IdEstablecimiento As Long
Dim ms_CodigoEstablecimiento As String
Dim ml_idUsuario As Long
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim mo_DatosParametros As New SIGHDatos.Parametros
Dim oRcs_DatosEstablecimiento As New ADODB.Recordset
Dim rsRecordset As ADODB.Recordset

'========================================= PROPIEDADES ==============================
Property Set DataSource(oValue As ADODB.Recordset)
    Set grdPadronNominal.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = grdPadronNominal.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
    IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

Private Sub bntReporte_Click()
    Dim oRsTmp1 As New Recordset
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Set oRsTmp1 = mo_ReglasHIS.PadronNominal_DetalleSeleccionarTodos
    mo_ReglasReportes.ExportarRecordSetAexcel oRsTmp1, "Padron Nominal", "", "", 0, True
End Sub

Private Sub btnBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, btnLimpiar
End Sub

'=========================================== EVENTOS ================================
Private Sub btnBuscar_Click()
    'If Trim(txtNroHistoria.Text) = "" Then
    '    MsgBox "Debe ingresar el número de historia clínica.", vbInformation, "Padron Nominal "
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
Dim oPadronNominal_Detalle As New DoPadronNominal_Detalle
        

'        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
'            UserControl.txtPrimerNombre = "" And UserControl.txtNroHistoria = "") Then
'            MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, Nombres o Nro Historia)", vbInformation, "Filtro de pacientes"
'            Exit Sub
'        End If
        
        oPadronNominal_Detalle.NumDocumento = IIf(UserControl.txtDni = "", 0, Val(UserControl.txtDni))
        oPadronNominal_Detalle.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPadronNominal_Detalle.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPadronNominal_Detalle.Nombres = UserControl.txtNombres
        oPadronNominal_Detalle.IdTipoDoc = 1
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria) Then
           oPadronNominal_Detalle.HistClinica = Val(UserControl.txtNroHistoria)
        End If
        Set grdPadronNominal.DataSource = mo_ReglasHIS.PadronNominalFiltrarNroHisClinica(oPadronNominal_Detalle)
        
        If mo_ReglasHIS.MensajeError <> "" Then
            MsgBox mo_ReglasHIS.MensajeError, vbCritical, "Filtro Padrón Nominal"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdPadronNominal, SIGHEntidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtNroHistoria.Text = ""
        UserControl.txtDni.Text = ""
        UserControl.txtApellidoPaterno.Text = ""
        UserControl.txtApellidoMaterno.Text = ""
        UserControl.txtNombres.Text = ""
End Sub

Private Sub grdPadronNominal_AfterRowActivate()
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdPadronNominal.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = IIf(IsNull(rsRecordset("idpanomdetalle")), -1, rsRecordset("idpanomdetalle"))
End Sub

'Configuracion de Detalle de Atenciones

Private Sub grdPadronNominal_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    Layout.Override.AllowDelete = ssAllowDeleteNo

    grdPadronNominal.Bands(0).Columns("idpanomdetalle").Hidden = True

    grdPadronNominal.Bands(0).Columns("histclinica").Header.Caption = "Historia Clinica"
    grdPadronNominal.Bands(0).Columns("histclinica").Width = 1500
    
    grdPadronNominal.Bands(0).Columns("numdocumento").Header.Caption = "Nro Documento"
    grdPadronNominal.Bands(0).Columns("numdocumento").Width = 1500

    grdPadronNominal.Bands(0).Columns("apellidopaterno").Header.Caption = "Apellido Paterno"
    grdPadronNominal.Bands(0).Columns("apellidopaterno").Width = 2000

    grdPadronNominal.Bands(0).Columns("apellidomaterno").Header.Caption = "Apellido Materno"
    grdPadronNominal.Bands(0).Columns("apellidomaterno").Width = 2000

    grdPadronNominal.Bands(0).Columns("nombres").Header.Caption = "Nombres"
    grdPadronNominal.Bands(0).Columns("nombres").Width = 2000
    
    grdPadronNominal.Bands(0).Columns("fecevaluacion").Header.Caption = "Evaluación"
    grdPadronNominal.Bands(0).Columns("fecevaluacion").Width = 1300

    grdPadronNominal.Bands(0).Columns("idsexo").Header.Caption = "Sexo"
    grdPadronNominal.Bands(0).Columns("idsexo").Width = 1500

    grdPadronNominal.Bands(0).Columns("peso").Header.Caption = "Peso(Kg)"
    grdPadronNominal.Bands(0).Columns("peso").Width = 1000

    grdPadronNominal.Bands(0).Columns("talla").Header.Caption = "Talla(cm)"
    grdPadronNominal.Bands(0).Columns("talla").Width = 1000

    grdPadronNominal.Bands(0).Columns("iddiagnutricional").Header.Caption = "Dx Nutricional"
    grdPadronNominal.Bands(0).Columns("iddiagnutricional").Width = 2500
 

End Sub


'=========================================== METODOS ================================
Public Function inicializar()
  
    'Obtiene los datos del establecimiento del usuario escogido
    Set oRcs_DatosEstablecimiento = mo_ReglasHIS.ObtenerDatosEstablecimientoPorUsuario(ml_idUsuario)
    
    'SE EVALUA SI NO TIENE DATOS CONFIGURADOS
    If oRcs_DatosEstablecimiento.RecordCount = 0 Then
        MsgBox "No tiene configurado el Establecimiento para el Usuario", vbInformation, "HIS "
        Exit Function
    End If
    Do While Not oRcs_DatosEstablecimiento.EOF
        ml_IdEstablecimiento = oRcs_DatosEstablecimiento!IdEstablecimiento
        ms_CodigoEstablecimiento = CStr(oRcs_DatosEstablecimiento!Codigo)
        oRcs_DatosEstablecimiento.MoveNext
    Loop
    
'    CargarComboBoxes
    
End Function

'Public Sub RealizarBusqueda()
'Set grdPadronNominal.DataSource = mo_ReglasHIS.ConsultarRegistroFiltroAtenciones(Val(txtDNI.tex(txtNroHistoria.Text)), _
'                                                                            Val(cmbServicios.ItemData(cmbServicios.ListIndex)), _
'                                                                            UserControl.txtAnio.Text, _
'                                                                            Val(cmbMes.ItemData(cmbMes.ListIndex)))
'
'If mo_ReglasHIS.MensajeError <> "" Then
'    MsgBox mo_ReglasHIS.MensajeError, vbCritical, "Filtro Registros HIS"
'End If
'End Sub

Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombres
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
    Select Case KeyAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc(" "), Asc("Ñ"), Asc("ñ")
    Case Else
    KeyAscii = 0
    End Select
    End If
End Sub

Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
    Select Case KeyAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc(" "), Asc("Ñ"), Asc("ñ")
    Case Else
    KeyAscii = 0
    End Select
    End If
End Sub

Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, btnBuscar
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
    Select Case KeyAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc(" "), Asc("Ñ"), Asc("ñ")
    Case Else
    KeyAscii = 0
    End Select
    End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
   fraBusqueda.Width = UserControl.Width - 110
   UserControl.lblNombre.Width = UserControl.Width
   grdPadronNominal.Width = fraBusqueda.Width
   grdPadronNominal.Height = UserControl.Height - (UserControl.lblNombre.Height + fraBusqueda.Height + 150)
End Sub

'Private Sub CargarComboBoxes()
''mo_cmbEstablecimiento.BoundColumn = "IdEstablecimiento"
''mo_cmbEstablecimiento.ListField = "NombreEstablecimiento"
''Set mo_cmbEstablecimiento.RowSource = ListadoEstablecimientos()
''
''mo_cmbServicio.BoundColumn = "IdServicio"
''mo_cmbServicio.ListField = "Nombre"
''Set mo_cmbServicio.RowSource = mo_ReglasHIS.ListaServiciosPorEstablecimiento(ml_IdEstablecimiento)
''
''mo_cmbMes.BoundColumn = "IdMes"
''mo_cmbMes.ListField = "NombreMes"
''Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses
'
''cmbEstablecimiento.ListIndex = 0
''cmbServicios.ListIndex = 0
''cmbMes.ListIndex = CInt(Month(CDate(mo_DatosParametros.RetornaFechaServidorSQL))) - 1
'
'mo_Apariencia.ConfigurarFilasBiColores grdPadronNominal, sighentidades.GrillaConFilasBicolor
'End Sub

'Private Sub cmbEstablecimiento_Click()
'Set mo_cmbServicio.RowSource = mo_ReglasHIS.ListaServiciosPorEstablecimiento(Val(mo_cmbEstablecimiento.BoundText))
'End Sub

Private Function ListadoEstablecimientos() As Recordset
Dim oTabla As New DOEstablecimiento

'If CInt(mo_DatosParametros.SeleccionaFilaParametro(208)) = ms_CodigoEstablecimiento Then
If CStr(mo_DatosParametros.SeleccionaFilaParametro(208)) = ms_CodigoEstablecimiento Then
    Set oRcs_DatosEstablecimiento = mo_ReglasHIS.ObtenerListaEstablecimientosMR
    oRcs_DatosEstablecimiento.MoveFirst
    Set ListadoEstablecimientos = oRcs_DatosEstablecimiento
Else
    oRcs_DatosEstablecimiento.MoveFirst
    Set ListadoEstablecimientos = oRcs_DatosEstablecimiento
End If
End Function

'Sub AdministrarKeyPreview(KeyCode As Integer)
'    Select Case KeyCode
'     Case vbKeyF6
'         btnBuscar_Click
'    End Select
'End Sub
