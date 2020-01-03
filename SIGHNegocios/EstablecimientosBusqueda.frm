VERSION 5.00
Begin VB.Form EstablecimientosBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "EstablecimientosBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SIGHNegocios.ucEstablecimientosLista ucEstablecimientosLista1 
      Height          =   5145
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   9075
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   2
      Top             =   5160
      Width           =   10425
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         DisabledPicture =   "EstablecimientosBusqueda.frx":0CCA
         DownPicture     =   "EstablecimientosBusqueda.frx":10B3
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
         Left            =   180
         Picture         =   "EstablecimientosBusqueda.frx":14BF
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Agregar"
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EstablecimientosBusqueda.frx":18CB
         DownPicture     =   "EstablecimientosBusqueda.frx":1D8F
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
         Left            =   5475
         Picture         =   "EstablecimientosBusqueda.frx":227B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EstablecimientosBusqueda.frx":2767
         DownPicture     =   "EstablecimientosBusqueda.frx":2BC7
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
         Left            =   3930
         Picture         =   "EstablecimientosBusqueda.frx":303C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "EstablecimientosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Establecimiento MINSA
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------


Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
'JVG - Valores de Descripcion de establecimiento y Nivel del Establecimineto
Dim ml_NivelMaximoEstablecimiento As Long
Dim ms_DescripcionEstablecimiento As String

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucEstablecimientosLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucEstablecimientosLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucEstablecimientosLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucEstablecimientosLista1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

'JVG - Adicion de Filtro de Niveles de Establecimiento Busqueda
Property Let NivelMaximoEstablecimiento(lValue As Long)
    ucEstablecimientosLista1.NivelMaximoEstablecimiento = lValue
End Property

'JVG - Adicion de Filtro de Descripcion de Establecimiento
Property Let DescripcionEstablecimiento(sValue As String)
    ucEstablecimientosLista1.DescripcionEstablecimiento = sValue
    ms_DescripcionEstablecimiento = sValue
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub


Private Sub cmdAgregar_Click()
    Dim lcCodigoRenaes As String
    
    lcCodigoRenaes = Trim(ucEstablecimientosLista1.DevuelveCodigoRenaes)
    
    If lcCodigoRenaes = "" Then
        MsgBox "Tiene que ingresar CODIGO RENAES para buscar en SIS", vbInformation, ""
    Else
        lcCodigoRenaes = sighentidades.FormatoCodigoRENAES(lcCodigoRenaes, SIS)
        
        Dim oBuscaEnSIS As New SIGHNegocios.SisConsumoWeb
        Dim oBuscaEnSUNASA As New SIGHNegocios.SunasaConsumoWeb
        
        Dim rsTmp As New ADODB.Recordset
        'mgaray201503
        Set rsTmp = oBuscaEnSIS.m_eessEleccionarPorCodigoRenaes(lcCodigoRenaes)
        
        
        If rsTmp.RecordCount = 0 Then
            If oBuscaEnSUNASA.HabilitadoParaBusquedaEnWebRenaes = False Then
                MsgBox "Código No Encontrado en la Base de Datos SIS," & _
                        " se recomienda habilitar la busqueda en la web de RENAES para ampliar la efectividad de la busqueda", vbInformation, Me.Caption
                Exit Sub
            End If
            Dim oDomEESs As Dom_eess
    '        If oBuscaEnSIS.ConsultarServicioBuscarEESSxCodigo(lcCodigoRenaes, rsTmp) = True Then
            Set oDomEESs = oBuscaEnSUNASA.ConsultarServicioBuscarEESSxCodigo(lcCodigoRenaes, rsTmp)
            
            If oDomEESs Is Nothing Then
                If oBuscaEnSUNASA.MensajeError = "" Then
                    MsgBox "Código de establecimiento : " & lcCodigoRenaes & ", No se existe en la WEB DEL RENAES", vbInformation, Me.Caption
                Else
                    MsgBox oBuscaEnSUNASA.MensajeError, vbInformation, Me.Caption
                End If
            Else
                If oBuscaEnSUNASA.EsEstablecimientoMinsa(oDomEESs) Then
                    Call grabarEstablecimientoEnSis(oDomEESs)
                    Set rsTmp = oBuscaEnSIS.m_eessEleccionarPorCodigoRenaes(lcCodigoRenaes)
                Else
                    MsgBox "Establecimiento encontrado en WEB de SUNASA, sin embargo " & _
                            "no es un establecimiento perteneciente al MINSA, (" & oDomEESs.pre_nombre & _
                            " - " & oDomEESs.Entidad & ")", _
                            vbInformation, Me.Caption
                End If
            End If
        End If
        
        If rsTmp.RecordCount > 0 Then
            ActualizaDatosEstablecimientoDesdeSIS rsTmp
        End If
        Set rsTmp = Nothing
        Set oBuscaEnSIS = Nothing
    End If
End Sub

Sub ActualizaDatosEstablecimientoDesdeSIS(rsTmp As Recordset)
    Dim oEstablecimientos As New Establecimientos, oDOEstablecimiento As New DOEstablecimiento
    Dim oConexion As New Connection, mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim lcCodigoRenaes As String
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    'mgaray201503
    lcCodigoRenaes = FormatoCodigoRENAES(rsTmp!pre_CodigoRENAES, GALENHOS)  '  Right(rsTmp!pre_codigoRENAES, 5)
    
    Dim oRsNoMinsa As New Recordset
    Set oRsNoMinsa = mo_ReglasComunes.EstablecimientosNoMinsaSeleccionarPorCodigo(lcCodigoRenaes)
    If oRsNoMinsa.RecordCount > 0 Then
        MsgBox "Código RENAES pertence a un establecimiento registrado como No Perteneciente al MINSA (" & oRsNoMinsa.Fields!nombre & ")", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mo_ReglasComunes.EstablecimientosSeleccionarPorCodigo(lcCodigoRenaes, oDOEstablecimiento) = False Then
        Set oEstablecimientos.Conexion = oConexion
        With oDOEstablecimiento
            .Codigo = lcCodigoRenaes
            .IdDistrito = Val(rsTmp!pre_idUbigeo)
            .IdEstablecimiento = mo_ReglasComunes.EstablecimientosDevuelveUltimoID(oConexion) + 1
            'mgaray201503
            .IdTipo = DevuelveTipoEstablecimientoPorIdCategoriaSIS(rsTmp!pre_idCategoriaEESS)
'            If rsTmp!pre_idCategoriaEESS = "01" Or rsTmp!pre_idCategoriaEESS = "02" Or rsTmp!pre_idCategoriaEESS = "12" Then
'               .IdTipo = sghTipoEstablecimiento.PuestoSalud ' 3   'ps
'            ElseIf rsTmp!pre_idCategoriaEESS = "03" Or rsTmp!pre_idCategoriaEESS = "04" Or rsTmp!pre_idCategoriaEESS = "11" Then
'               .IdTipo = sghTipoEstablecimiento.CentroSalud ' 2   'cs
'            Else
'               .IdTipo = sghTipoEstablecimiento.Hospital ' 1   'hosp
'            End If
           ' .IdUsuarioAuditoria = rsTmp!
            .nombre = Left(rsTmp!pre_nombre, 150)
        End With
        If oEstablecimientos.Insertar(oDOEstablecimiento) = False Then
           MsgBox oEstablecimientos.MensajeError, vbInformation, "Establecimientos"
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oEstablecimientos = Nothing
    Set oDOEstablecimiento = Nothing
    Set mo_ReglasComunes = Nothing
    ucEstablecimientosLista1.BuscarPorCodigoRenaes lcCodigoRenaes
End Sub

Private Sub Form_Initialize()
    ucEstablecimientosLista1.ConfigurarEstablecimientos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    Me.ucEstablecimientosLista1.Inicializar
    Me.ucEstablecimientosLista1.Titulo = "Búsqueda de Establecimientos"
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucEstablecimientosLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub ucEstablecimientosLista1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
    If lnIdRegistroSeleccionado > 0 Then
       btnAceptar_Click
    End If
End Sub

Private Function DevuelveTipoEstablecimientoPorIdCategoriaSIS(sIdCategoriaSis As String) As Long
    Select Case sIdCategoriaSis
        Case "01", "02", "12":
            DevuelveTipoEstablecimientoPorIdCategoriaSIS = sghTipoEstablecimiento.PuestoSalud
        Case "03", "04", "11":
            DevuelveTipoEstablecimientoPorIdCategoriaSIS = sghTipoEstablecimiento.CentroSalud
        Case Else
            DevuelveTipoEstablecimientoPorIdCategoriaSIS = sghTipoEstablecimiento.Hospital
    End Select
End Function

Private Function grabarEstablecimientoEnSis(ODom_eess As Dom_eess) As Boolean
On Error GoTo miError
    Dim oConexion As New Connection
    Dim bResultado As Boolean
    Dim oM_eess As New m_eess
    Dim rsTmp As Recordset
    Dim oBuscaEnSIS As New SIGHNegocios.SisConsumoWeb
    
    bResultado = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oM_eess.Conexion = oConexion
    Set rsTmp = oBuscaEnSIS.m_eessEleccionarPorCodigoRenaes(ODom_eess.pre_CodigoRENAES)
    If rsTmp.RecordCount = 0 Then
        bResultado = oM_eess.Insertar(ODom_eess)
    End If
    grabarEstablecimientoEnSis = bResultado
miError:
    If Err Then
        MsgBox Err.Description, vbCritical, Me.Caption
    End If
End Function



