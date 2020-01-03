VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucSolicitudHistoriasLista 
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   ScaleHeight     =   6495
   ScaleWidth      =   10185
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   60
      TabIndex        =   11
      Top             =   570
      Width           =   10035
      Begin VB.CommandButton cmdSinApellidoMaterno 
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
         Left            =   5205
         Picture         =   "ucSolicitudHistoriasLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   315
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
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
         Left            =   3270
         Picture         =   "ucSolicitudHistoriasLista.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1095
         Width           =   315
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8820
         Picture         =   "ucSolicitudHistoriasLista.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1080
         Width           =   1275
      End
      Begin VB.TextBox txtNombreArchivero 
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
         Left            =   4980
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   4815
      End
      Begin VB.TextBox txtCodigoArchivero 
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
         Top             =   540
         Width           =   1515
      End
      Begin VB.TextBox txtPrimerNombre 
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
         Left            =   5580
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1095
         Width           =   1755
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
         Left            =   3675
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1095
         Width           =   1515
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
         Left            =   1755
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1095
         Width           =   1500
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7455
         Picture         =   "ucSolicitudHistoriasLista.ctx":36F0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   1305
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
         Left            =   165
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1095
         Width           =   1515
      End
      Begin MSMask.MaskEdBox txtFechaSolicitud 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   540
         Width           =   1485
         _ExtentX        =   2619
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
      Begin MSMask.MaskEdBox txtFechaRequerida 
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   540
         Width           =   1500
         _ExtentX        =   2646
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
      Begin VB.Label Label1 
         Caption         =   "   Código archivero     Fecha solicitud      Fecha requerida                           Nombre de archivero"
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
         Left            =   60
         TabIndex        =   13
         Top             =   270
         Width           =   9405
      End
      Begin VB.Label Label2 
         Caption         =   "Nº historia clínica       Apellido paterno          Apellido materno          Primer nombre                   "
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
         Left            =   210
         TabIndex        =   12
         Top             =   870
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdPrestamosHC 
      Height          =   4230
      Left            =   60
      TabIndex        =   10
      Top             =   2175
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   7461
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
      Caption         =   "Lista de solicitud de historias"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Solicitud Historia Clínica"
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
      Left            =   45
      TabIndex        =   14
      Top             =   45
      Width           =   10200
   End
End
Attribute VB_Name = "ucSolicitudHistoriasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar Solicitud de Historia
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_IdArchivero  As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdPrestamosHC.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdPrestamosHC.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoBusqueda(lValue As sghTipoBusquedaPrestamoHistoria)
    ml_TipoBusqueda = lValue
End Property
Property Get TipoBusqueda() As sghTipoBusquedaPrestamoHistoria
    TipoBusqueda = ml_TipoBusqueda
End Property
Property Let IdArchivero(lValue As Long)

    ml_IdArchivero = lValue
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(lValue)
    
    UserControl.txtCodigoArchivero = oDOEmpleado.IdEmpleado
    UserControl.txtCodigoArchivero = oDOEmpleado.CodigoPlanilla
    UserControl.txtNombreArchivero = oDOEmpleado.ApellidoPaterno & " " & oDOEmpleado.ApellidoMaterno & " " & oDOEmpleado.Nombres
    
End Property
Property Get IdArchivero() As Long
    IdArchivero = ml_IdArchivero
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        If UserControl.txtCodigoArchivero = "" Then
           MsgBox "Ingrese el CODIGO DEL ARCHIVERO", vbInformation, ""
           Exit Sub
        End If

        Dim oPaciente As New doPaciente
        Dim oHistoria As New DOHistoriaSolicitada
        Dim oDOArchiveroServ As New DOArchiveroServicio
        
        
        
        txtCodigoArchivero_LostFocus
        
        
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtPrimerNombre = "" And UserControl.txtNroHistoria = "") Then
        End If
            
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPaciente.PrimerNombre = UserControl.txtPrimerNombre
        oPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria))
        oHistoria.FechaSolicitud = IIf(UserControl.txtFechaSolicitud = sighentidades.FECHA_VACIA_DMY, 0, UserControl.txtFechaSolicitud)
        oHistoria.FechaRequerida = IIf(UserControl.txtFechaRequerida = sighentidades.FECHA_VACIA_DMY, 0, UserControl.txtFechaRequerida)
        oDOArchiveroServ.IdEmpleado = Val(UserControl.txtCodigoArchivero.Tag)
        
        Set grdPrestamosHC.DataSource = mo_AdminArchivoClinico.HistoriasSolicitadasFiltrar(oPaciente, oHistoria, oDOArchiveroServ)
        If mo_AdminArchivoClinico.MensajeError <> "" Then
            MsgBox mo_AdminArchivoClinico.MensajeError, vbInformation, "Filtro Historias Solicitadas"
            Exit Sub
        End If
        
        Dim rsResultados As Recordset
        Set rsResultados = grdPrestamosHC.DataSource
        
        If Not (rsResultados.EOF And rsResultados.BOF) Then
        Else
            MsgBox "No se encontraron datos", vbInformation, "Búsqueda de historias solicitadas"
            Exit Sub
        End If
        
        'mo_Apariencia.ConfigurarFilasBiColores grdPrestamosHC, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtPrimerNombre = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtFechaSolicitud = sighentidades.FECHA_VACIA_DMY
        UserControl.txtFechaRequerida = sighentidades.FECHA_VACIA_DMY
        UserControl.txtCodigoArchivero.Tag = ""
        UserControl.txtCodigoArchivero = ""
        UserControl.txtNombreArchivero = ""
        
End Sub


Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub grdPrestamosHC_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdPrestamosHC.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdHistoriaSolicitada")
End Sub

Private Sub grdPrestamosHC_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdPrestamosHC.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdHistoriaSolicitada")
    
End Sub


Private Sub grdPrestamosHC_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPrestamosHC.Bands(0).Columns("IdHistoriaSolicitada").Hidden = True
    
    grdPrestamosHC.Bands(0).Columns("HistoriaClinica").Header.Caption = "Nro Historia"
    grdPrestamosHC.Bands(0).Columns("HistoriaClinica").Width = 1000
    
    grdPrestamosHC.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPrestamosHC.Bands(0).Columns("TipoNumeracion").Width = 3000
    
    grdPrestamosHC.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPrestamosHC.Bands(0).Columns("ApellidoPaterno").Width = 1200
    
    grdPrestamosHC.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPrestamosHC.Bands(0).Columns("ApellidoMaterno").Width = 1200
    
    grdPrestamosHC.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPrestamosHC.Bands(0).Columns("PrimerNombre").Width = 1200
    
    grdPrestamosHC.Bands(0).Columns("FechaSolicitud").Header.Caption = "Fecha Sol."
    grdPrestamosHC.Bands(0).Columns("FechaSolicitud").Width = 1750
    
    grdPrestamosHC.Bands(0).Columns("FechaRequerida").Header.Caption = "Fecha Req."
    grdPrestamosHC.Bands(0).Columns("FechaRequerida").Width = 1750
    
    grdPrestamosHC.Bands(0).Columns("Nombre").Header.Caption = "Servicio"
    grdPrestamosHC.Bands(0).Columns("Nombre").Width = 2000

    grdPrestamosHC.Bands(0).Columns("Observacion").Header.Caption = "Observación"
    grdPrestamosHC.Bands(0).Columns("Observacion").Width = 5000
    
End Sub

Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
End Sub


Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtCodigoArchivero_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtCodigoArchivero, txtNombreArchivero
End Sub

Private Sub txtCodigoArchivero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigoArchivero
End Sub

Private Sub txtCodigoArchivero_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtFechaSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaSolicitud
End Sub

Private Sub txtFechaSolicitud_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtFechaRequerida_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaRequerida
End Sub

Private Sub txtFechaRequerida_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
End Sub


Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdPrestamosHC.Width = fraBusqueda.Width
   grdPrestamosHC.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Sub CompletarDatosDeEmpleadoEnElLostFocus(txtCodigoPlanilla As TextBox, txtNombre As TextBox)
Dim oDOEmpleado As New dOEmpleado

        If mo_AdminComun.EmpleadosSeleccionarPorCodigo(txtCodigoPlanilla.Text, oDOEmpleado) Then
            txtCodigoPlanilla.Tag = oDOEmpleado.IdEmpleado
            txtCodigoPlanilla.Text = oDOEmpleado.CodigoPlanilla
            txtNombre = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtCodigoPlanilla.Tag = ""
            txtCodigoPlanilla = ""
            txtNombre = ""
        End If
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdPrestamosHC, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdPrestamosHC, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub inicializar()
    SkinConfigura
End Sub
