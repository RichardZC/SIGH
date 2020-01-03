VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucEstablecimientosLista 
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   10110
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
      Height          =   1485
      Left            =   75
      TabIndex        =   8
      Top             =   510
      Width           =   10005
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7410
         Picture         =   "ucListaEstablecimientos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8775
         Picture         =   "ucListaEstablecimientos.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1275
      End
      Begin VB.ComboBox cmbIdDistrito 
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
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   2565
      End
      Begin VB.ComboBox cmbIdProvincia 
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
         Left            =   2415
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Width           =   2310
      End
      Begin VB.ComboBox cmbIdDepartamento 
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
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1020
         Width           =   2175
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   150
         TabIndex        =   0
         Top             =   450
         Width           =   1275
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   1500
         TabIndex        =   1
         Top             =   450
         Width           =   5850
      End
      Begin VB.Label Label2 
         Caption         =   " Cód.RENAES                                                 Nombre"
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
         TabIndex        =   10
         Top             =   240
         Width           =   7635
      End
      Begin VB.Label Label1 
         Caption         =   "      Departamento                        Provincia                               Distrito"
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
         TabIndex        =   9
         Top             =   780
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdEstablecimientos 
      Height          =   3825
      Left            =   75
      TabIndex        =   7
      Top             =   2055
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6747
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
      Caption         =   "Lista de establecimientos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Establecimientos"
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
      Left            =   15
      TabIndex        =   11
      Top             =   0
      Width           =   10155
   End
End
Attribute VB_Name = "ucEstablecimientosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Establecimientos MINSA
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminReglasCOmunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbIdProvincia As New sighentidades.ListaDespleglable
Dim mo_cmbIdDistrito As New sighentidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)

'JVG - variables de Filtro
Dim ml_NivelMaximo As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdEstablecimientos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdEstablecimientos.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

'JVG - Nivel maximo de Establecimiento
Property Let NivelMaximoEstablecimiento(lValue As Long)
    ml_NivelMaximo = lValue
End Property

'JVG - Descripcion de Establecimiento
Property Let DescripcionEstablecimiento(sValue As String)
    UserControl.txtNombre.Text = sValue
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
        Dim oEstablecimiento As New DOEstablecimiento
        
        If (UserControl.txtCodigo = "" And UserControl.txtNombre = "" And _
            UserControl.cmbIdDepartamento = "" And UserControl.cmbIdProvincia = "" _
            And UserControl.cmbIdDistrito = "") Then
        End If
        'mgaray201503
        oEstablecimiento.Codigo = FormatoCodigoRENAES(UserControl.txtCodigo, GALENHOS) 'UserControl.txtCodigo
        oEstablecimiento.nombre = UserControl.txtNombre
        oEstablecimiento.IdDistrito = Val(mo_cmbIdDistrito.BoundText)
        
        'JVG - Adicion de Nivel maximo de Consulta y Descripcion
        oEstablecimiento.IdTipo = ml_NivelMaximo
        'mgaray201503
        Dim oRs As ADODB.Recordset
        Set oRs = mo_AdminReglasCOmunes.EstablecimientosFiltrar(oEstablecimiento, _
                        Val(mo_cmbIdDepartamento.BoundText), Val(mo_cmbIdProvincia.BoundText))
        
        Set grdEstablecimientos.DataSource = oRs
        
'        Set grdEstablecimientos.DataSource = mo_AdminReglasCOmunes.EstablecimientosFiltrar(oEstablecimiento, _
'                                                    Val(mo_cmbIdDepartamento.BoundText), Val(mo_cmbIdProvincia.BoundText))
        
        If mo_AdminReglasCOmunes.MensajeError <> "" Then
            MsgBox mo_AdminReglasCOmunes.MensajeError, vbInformation, "Filtro Prestamos HC"
        Else
            On Error Resume Next
            If ml_NivelMaximo <> 0 Then
                If oRs.RecordCount = 0 And oEstablecimiento.Codigo <> "" Then
                    Dim oEstablecimientoAux As New DOEstablecimiento
                    Dim oRsAux As ADODB.Recordset
                    Dim sMensaje As String
                    
                    oEstablecimientoAux.Codigo = oEstablecimiento.Codigo
                    oEstablecimientoAux.IdTipo = 0
                    oEstablecimientoAux.nombre = ""
                    oEstablecimientoAux.IdDistrito = 0
                    Set oRsAux = mo_AdminReglasCOmunes.EstablecimientosFiltrar(oEstablecimientoAux, _
                                0, 0)
                    If oRsAux.RecordCount > 0 Then
                        sMensaje = "Codigó de Establecimiento Existe, sin embargo no podra mostrarse debido a:" & _
                                Chr(13) & "- Establecimiento no es del tipo que se esta buscando"
                        If Val(mo_cmbIdDepartamento.BoundText) <> 0 Or _
                                Val(mo_cmbIdProvincia.BoundText) <> 0 Or _
                                oEstablecimiento.IdDistrito <> 0 Then
                                
                            sMensaje = sMensaje & Chr(13) & "- Establecimiento no pertenece a la zona (Departamento/Provincia/Distrito)"
                        End If
                        If oEstablecimiento.nombre <> "" Then
                            sMensaje = sMensaje & Chr(13) & "- Nombre que se especificado para el código no coincide con el registrado"
                        End If
                        
                        sMensaje = sMensaje & Chr(13) & Chr(13) & "Datos Establecimiento encontrado :" & _
                                                Chr(13) & "Nombre : " & oRsAux.Fields!nombre & _
                                                Chr(13) & " Distrito : " & IIf(IsNull(oRsAux.Fields!Distrito), "", oRsAux.Fields!Distrito) & _
                                                Chr(13) & " Provincia: " & IIf(IsNull(oRsAux.Fields!Provincia), "", oRsAux.Fields!Provincia) & _
                                                Chr(13) & "Departamento : " & IIf(IsNull(oRsAux.Fields!Departamento), "", oRsAux.Fields!Departamento)
                                                
                        MsgBox sMensaje, vbInformation, "Busqueda de Establecimientos"
                        
                    End If
                End If
            End If
            grdEstablecimientos.SetFocus
        End If
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtCodigo = ""
        UserControl.txtNombre = ""
        mo_cmbIdDepartamento.BoundText = ""
        mo_cmbIdProvincia.BoundText = ""
        mo_cmbIdDistrito.BoundText = ""
End Sub

Private Sub cmbIdDepartamento_Click()
        
       mo_cmbIdProvincia.BoundColumn = "IdProvincia"
       mo_cmbIdProvincia.ListField = "Nombre"
       On Error Resume Next
       Set mo_cmbIdProvincia.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       
       mo_cmbIdProvincia.BoundText = ""
       mo_cmbIdDistrito.BoundText = ""
        
End Sub
Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDepartamento_LostFocus()
   'If cmbIdDepartamento.Text <> "" Then
   '    mo_cmbIdDepartamento.BoundText = Val(Split(cmbIdDepartamento.Text, " = ")(0))
   'End If
End Sub

Private Sub cmbIdDepartamento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDistrito
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDistrito_LostFocus()
   'If cmbIdDistrito.Text <> "" Then
   '    mo_cmbIdDistrito.BoundText = Val(Split(cmbIdDistrito.Text, " = ")(0))
   'End If
End Sub

Private Sub cmbIdDistrito_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdProvincia_Click()
        
       mo_cmbIdDistrito.BoundColumn = "IdDistrito"
       mo_cmbIdDistrito.ListField = "Nombre"
       Set mo_cmbIdDistrito.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(mo_cmbIdProvincia.BoundText))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Lista de establecimientos"
       End If
       
       mo_cmbIdDistrito.BoundText = ""
       
End Sub

Private Sub cmbIdProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdProvincia
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdProvincia_LostFocus()
   
   'If cmbIdProvincia.Text <> "" Then
   '    mo_cmbIdProvincia.BoundText = Val(Split(cmbIdProvincia.Text, " = ")(0))
   'End If
   
   
End Sub

Private Sub cmbIdProvincia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub grdEstablecimientos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdEstablecimientos.DataSource
    ml_IdRegistroSeleccionado = rsRecordset("IdEstablecimiento")
End Sub


Private Sub grdEstablecimientos_DblClick()
    Dim rsRecordset As ADODB.Recordset
    On Error Resume Next
    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdEstablecimientos.DataSource
    ml_IdRegistroSeleccionado = rsRecordset("IdEstablecimiento")
    RaiseEvent SeleccionaRegistro(ml_IdRegistroSeleccionado)
End Sub

Private Sub grdEstablecimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdEstablecimientos.Bands(0).Columns("IdEstablecimiento").Hidden = True
    'grdEstablecimientos.Bands(0).Columns("IdEnvio").Hidden = True
    
    grdEstablecimientos.Bands(0).Columns("Codigo").Header.Caption = "Codigo"
    grdEstablecimientos.Bands(0).Columns("Codigo").Width = 750
    
    grdEstablecimientos.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdEstablecimientos.Bands(0).Columns("Nombre").Width = 4500
    
    grdEstablecimientos.Bands(0).Columns("Departamento").Header.Caption = "Departamento"
    grdEstablecimientos.Bands(0).Columns("Departamento").Width = 1500
    
    grdEstablecimientos.Bands(0).Columns("Provincia").Header.Caption = "Provincia"
    grdEstablecimientos.Bands(0).Columns("Provincia").Width = 1500
    
    grdEstablecimientos.Bands(0).Columns("Distrito").Header.Caption = "Distrito"
    grdEstablecimientos.Bands(0).Columns("Distrito").Width = 1500

    
End Sub

Public Function Inicializar()
    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdProvincia.MiComboBox = cmbIdProvincia
    Set mo_cmbIdDistrito.MiComboBox = cmbIdDistrito
    mo_Apariencia.ConfigurarFilasBiColores grdEstablecimientos, sighentidades.GrillaConFilasBicolor
End Function

Private Sub grdEstablecimientos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
   If KeyAscii = 13 Then
      grdEstablecimientos_DblClick
   End If
End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 150
   lblNombre.Width = UserControl.Width
   grdEstablecimientos.Width = UserControl.Width - 150
   grdEstablecimientos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 100)
   
End Sub

Public Sub ConfigurarEstablecimientos()
    
    mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
    mo_cmbIdDepartamento.ListField = "Nombre"
    Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosGeograficos.DepartamentosSeleccionarTodos()
    mo_cmbIdDepartamento.BoundText = Trim(Str(Val(Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2))))

End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
        btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub

Public Function DevuelveCodigoRenaes() As String
    DevuelveCodigoRenaes = Trim(txtCodigo.Text)
End Function

Public Function BuscarPorCodigoRenaes(lcCodigoRenaes As String)
    txtCodigo.Text = lcCodigoRenaes
    txtNombre.Text = ""
    mo_cmbIdDistrito.BoundText = ""
    mo_cmbIdDepartamento.BoundText = ""
    mo_cmbIdProvincia.BoundText = ""
    btnBuscar_Click
End Function

Public Function DevuelveCodigoEstablecimiento() As String
    DevuelveCodigoEstablecimiento = Trim(txtCodigo.Text)
End Function

Public Function DevuelveNombreEstablecimiento() As String
    DevuelveNombreEstablecimiento = Trim(txtNombre.Text)
End Function
