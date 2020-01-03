VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Principal1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "SGH para Clínicas"
   ClientHeight    =   8910
   ClientLeft      =   1260
   ClientTop       =   690
   ClientWidth     =   15240
   Icon            =   "Principal1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SGHclinicas.ucPacientesLista ucPacientesLista1 
      Height          =   465
      Left            =   1410
      TabIndex        =   10
      Top             =   285
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   820
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1125
      OleObjectBlob   =   "Principal1.frx":0CCA
      Top             =   6600
   End
   Begin SGHclinicas.ucCajaNotaCredito ucCajaNotaCredito1 
      Height          =   615
      Left            =   1500
      TabIndex        =   8
      Top             =   3645
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.CommandButton cmdFechaHoraServidor 
      BackColor       =   &H00FF0000&
      Height          =   405
      Left            =   12420
      Picture         =   "Principal1.frx":0EFE
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Muestra Hora actual del SERVIDOR"
      Top             =   120
      Width           =   465
   End
   Begin VB.Timer tmrHora 
      Interval        =   20000
      Left            =   10560
      Top             =   5640
   End
   Begin SGHclinicas.ucCatalogoServiciosLista ucCatalogoServiciosLista1 
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   5355
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   5318
      _ExtentY        =   873
   End
   Begin SGHclinicas.ucCajeroLista ucCajeroLista1 
      Height          =   435
      Left            =   5700
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   767
   End
   Begin SGHclinicas.ucCajaLista ucCajaLista1 
      Height          =   375
      Left            =   1125
      TabIndex        =   3
      Top             =   4770
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   8565
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SGHclinicas.ucProcedimientosLista ucProcedimientosLista1 
      Height          =   495
      Left            =   1170
      TabIndex        =   1
      Top             =   1065
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   873
   End
   Begin SGHclinicas.ucEmpleadosLista ucEmpleadosLista1 
      Height          =   585
      Left            =   810
      TabIndex        =   0
      Top             =   2070
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   1032
   End
   Begin VB.PictureBox pctLogo 
      AutoSize        =   -1  'True
      BackColor       =   &H00373842&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   6510
      Left            =   1680
      Picture         =   "Principal1.frx":1340
      ScaleHeight     =   6510
      ScaleWidth      =   10200
      TabIndex        =   6
      Top             =   1890
      Visible         =   0   'False
      Width           =   10200
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   3  'Align Left
      Height          =   8565
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   15108
      ButtonWidth     =   2275
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "VALESC"
            Object.ToolTipText     =   "Registro de Vales de Combustibles"
            ImageIndex      =   3
            Object.Width           =   1500
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Key             =   "RECORRIDO"
            Object.ToolTipText     =   "Registro de Pedidos de Movilidad"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            Key             =   "REPARACIONES"
            Object.ToolTipText     =   "Registro de Gastos por Repuestos y Mano de Obra"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SALIR"
            Key             =   "SALIR"
            Description     =   "salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal1.frx":1713F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal1.frx":17459
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal1.frx":178AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal1.frx":17D01
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal1.frx":18155
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivos 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnuModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCajaApertura 
         Caption         =   "Caja Apertura"
      End
      Begin VB.Menu mnuCajaCierre 
         Caption         =   "Caja Cierre"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "Procesos"
      Begin VB.Menu mnuCpt 
         Caption         =   "Catálogo de Procedimientos"
      End
      Begin VB.Menu mnuEmpleados 
         Caption         =   "Catálogo de Empleados"
      End
      Begin VB.Menu mnuCajasMan 
         Caption         =   "Catálogo de Cajas"
      End
      Begin VB.Menu mnuPacientes 
         Caption         =   "Pacientes"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCaja 
         Caption         =   "Caja Gestión"
      End
      Begin VB.Menu mnuNCredito 
         Caption         =   "Nota de Crédito"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportaSunat 
         Caption         =   "Exporta datos a SUNAT"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuVentas 
         Caption         =   "Ventas"
      End
   End
   Begin VB.Menu mnuSalir1 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Principal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Programa Principal del SIstema, muestra MENU
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
'Option Explicit
Dim ms_ModuloSeleccionado As String
Dim mo_LastControl As Control
Dim mo_LoginForm As Login1
Dim mb_MantenerValoresCitas As Boolean


'Visitas

'Referencias a reglas de negocios
Dim mo_FuenteFinanciamientoDetalle As New SIGHCatalogos.clFuenteFinancDetalle
Dim mo_PartidasDetalle As New SIGHCatalogos.clPartidaDetalle
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja

Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
'Referencia a forms
Dim ml_IdUsuarioAuditoria As Long
Dim mb_LoadingForm As Boolean
Dim mrs_ListItems As New Recordset
Dim ml_ToolbarHeightAdd As Long
Dim mb_abrioCaja As Boolean
Dim lc_NombrePc As String
Dim lbVisualizaListaMedicamentosVencidos As Boolean

'mgaray201503
Dim lbCajeroEmiteSoloServicios As Boolean
Dim mb_UsuarioActualEsCajero As Boolean
Dim moDOCajaGestion As DOCajaGestion
Dim lcOpcionElegida As String

Property Get bAbrioCaja() As Boolean
   bAbrioCaja = mb_abrioCaja
End Property

Property Get oDOCajaGestion() As DOCajaGestion
   Set oDOCajaGestion = moDOCajaGestion
End Property

Property Get bCajeroEmiteSoloServicios() As Boolean
   bCajeroEmiteSoloServicios = lbCajeroEmiteSoloServicios
End Property

Property Get UsuarioActualEsCajero() As Boolean
   UsuarioActualEsCajero = mb_UsuarioActualEsCajero
End Property

'Franco Temporal
Property Get Turno() As Integer
    Dim Hora As Integer
    Hora = Val(Format(Now, "HH"))
    If Hora >= 6 And Hora <= 13 Then
        Turno = 1
    ElseIf Hora >= 14 And Hora <= 21 Then
        Turno = 2
    ElseIf Hora >= 22 Or (Hora >= 0 And Hora <= 5) Then
        Turno = 3
    End If
End Property

Property Set LoginForm(oValue As Login1)
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Set mo_LoginForm = oValue
    ml_IdUsuarioAuditoria = oValue.IdUsuarioAutenticado
    status.Panels(2).Text = "Usuario: " & oValue.NombreUsuarioAutenticado
    status.Panels(3).Text = "Servidor: " & lcBuscaParametro.RetornaNombreDeServidor
    status.Panels(4).Text = "PC: " & lc_NombrePc
    status.Panels(5).Text = lcBuscaParametro.SeleccionaFilaParametro(205)
    status.Panels(6).Text = WxLcVersionSisGalenPlus
    status.Panels(7).Text = lcBuscaParametro.SeleccionaFilaParametro(314) & " " & lcBuscaParametro.RetornaVersionServidorSQLserver
    Set lcBuscaParametro = Nothing
End Property

Private Sub CentrarImagen()
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  If lcBuscaParametro.SeleccionaFilaParametro(282) = "S" Then
     pctLogo.Picture = LoadPicture(App.Path & "\Imagenes\principalcs.jpg")
  Else
     pctLogo.Picture = LoadPicture(App.Path & "\Imagenes\principal.jpg")
  End If
  'Centrar imagen
  Dim to_x As Single
  Dim to_y As Single
  If pctLogo.Picture = 0 Then Exit Sub
  Cls
  to_x = (ScaleWidth - pctLogo.ScaleWidth) / 2
  'to_y = (ScaleHeight - pctLogo.ScaleHeight) / 2
  to_y = 0
  
  Me.PaintPicture pctLogo.Picture, to_x, to_y ', , , , , , &H373842
  Set lcBuscaParametro = Nothing
End Sub

Private Sub Form_Activate()
    CargaSetup_X_PC

End Sub

Private Sub Form_Initialize()
    
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path + "\Imagenes\principal.jpg")
    
    mb_LoadingForm = True
    
    GenerarRecordsetDeListItems
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
On Error Resume Next

    Select Case KeyCode
    Case vbKeyEscape
    
        'WCG 04/06/2006
        Select Case ms_ModuloSeleccionado
        
        'EFGL 14/06/2006
        Case "GestionCaja", "FacturacionProcedimientos", "FacturacionPatologiaClinica", "FacturacionAnatomiaPatologica", "FacturacionImaginologia", "EstadoCuenta"
        'fin EFGL 14/06/2006
        Case Else
            mo_LastControl.Visible = False
        End Select
        
    Case vbKeyF2
    Case vbKeyF6
    
        RealizarBusquedas
    Case vbKeyF7
        LimpiarFiltro
    Case vbKeyF8
    Case vbKeyF9
    Case vbKeyF10
    Case vbKeyF11
    Case vbKeyF12
    
    End Select
       
End Sub
Sub RealizarBusquedas()
    Select Case ms_ModuloSeleccionado
    'MODULO AMBULATORIO
    Case "AdmisionCE"
        'ucCitasLista1
    Case "PacienteCE"
    Case "AtencionesCE"
    Case "InterconsultasCE"
    'MODULO DE CONSULTORIOS DE EMERGENCIA
    Case "PacienteEmerg", "PacienteObservacionEmerg"
    Case "AdmisionConsultorioEmerg"
    
    Case "AtencionesConsultorioEmerg"
        
    Case "InterconsultasConsEmerg"
    
    'MODULO OBSERVACION EMERGENCIA
    Case "AdmisionObservacionEmerg"
        
    Case "InterconsultasObsEmerg"
        
    Case "CamasEmergencia"
        
    
    'MODULO DE HOSPITALIZACION
    Case "PacienteHosp"
        
    Case "AdmisionHospitalizacion"
        
    Case "AtencionesHospitalizacion"
        
    Case "CamasHospitalizacion"
        
    Case "InterconsultasHosp"
        
    
    'MODULO PROGRAMACION MEDICA
       
    'MODULO GENERAL
    Case "Empleado"
        ucEmpleadosLista1.RealizarBusqueda
    Case "Servicios"
    Case "Procedimientos"
        ucProcedimientosLista1.RealizarBusqueda
        
    'MZD Ini 01/06/2005
    'MODULO CAJA
    Case "MovimientosCaja"
        
    'MZD Fin 01/06/2005
    'FIN GENERAL
    'SEGURIDAD
    Case "Roles"
    'mgaray20141009
    Case "AtencionesTriaje":
    'mgaray201411f
    'IMAGENOLOGIA
    Case "ImagTipoModalidadSala":
    Case "ImagSala":
       
    Case "ImagCatalgoServicioDuracion":
    Case "IntegracionSistema"
    End Select

End Sub
Sub LimpiarFiltro()

    Select Case ms_ModuloSeleccionado
    'MODULO AMBULATORIO
    Case "AdmisionCE"
        'ucCitasLista1
    Case "PacienteCE"
    Case "InterconsultasConsEmerg"
        
    
    'MODULO OBSERVACION EMERGENCIA
    Case "AdmisionObservacionEmerg"
    
    'MODULO PROGRAMACION MEDICA
    Case "Programacion"
        
           
    'MODULO GENERAL
    Case "Empleado"
        ucEmpleadosLista1.LimpiarFiltro
    Case "Servicios"
    Case "Procedimientos"
        ucProcedimientosLista1.LimpiarFiltro
     
    'MZD Ini 01/06/2005
    'MODULO CAJA
    Case "MovimientosCaja"
        
    'MZD Fin 01/06/2005
    'FIN GENERAL
    'SEGURIDAD
    Case "Roles"
    'mgaray20141009
    Case "AtencionesTriaje":
    'mgaray201411f
    'IMAGENOLOGIA
    Case "ImagTipoModalidadSala":
    Case "ImagSala":
        
    Case "ImagCatalgoServicioDuracion":
    Case "IntegracionSistema"
    End Select
    
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  Skin1.LoadSkin App.Path & "\" & WxSkin
  Skin1.ApplySkin Me.hwnd
ErrSkin:
End Sub

Private Sub Form_Load()
    SkinConfigura
    ml_ToolbarHeightAdd = 0
    mb_MantenerValoresCitas = False
    lc_NombrePc = sighentidades.RetornaNombrePC
    OcultaBotonXdelFormulario Me.hwnd
    EliminaArchivosOpenOffice
    
End Sub

Private Sub EliminaArchivosOpenOffice()
   Dim Archivo As String, viejo As String
   Dim flag As Boolean
   Dim c As Integer
   On Error GoTo ElimArOP
    flag = True
    viejo = "xxx"
    While (flag = True)
        Archivo = Dir(App.Path & "\plantillas\*.ods")
        If Archivo = "" Or Archivo = viejo Then
            flag = False
        Else
            If InStr("1234567890", Left(Archivo, 1)) > 0 Then
                Kill App.Path & "\plantillas\" & Archivo
            Else
                viejo = Archivo
            End If
        End If
    Wend
ElimArOP:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    GalenhosKillExcelApplication
    End
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    
    mo_LastControl.Top = 0
    mo_LastControl.Left = 600
    mo_LastControl.Width = Me.Width - 100
    mo_LastControl.Height = Me.Height
    
    CentrarImagen
    
    If Me.WindowState <> vbMinimized Then Me.WindowState = vbMaximized
    'debb-hra
    cmdFechaHoraServidor.Top = Me.Top + Me.Height - 1700
    cmdFechaHoraServidor.Left = Me.Left '+ 4700
    
End Sub
Sub ConfigurarPermisosDelItemSeleccionado(lIdUsuario As Long, lIdListItem As Long, sKey As String)



End Sub

Private Sub Form_Terminate()
  mo_AdminSeguridad.LogueaUsuario 0, sighentidades.USUARIO, lc_NombrePc
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mo_AdminSeguridad.LogueaUsuario 0, sighentidades.USUARIO, lc_NombrePc
End Sub



Private Sub mnuAgregar_Click()
    Select Case lcOpcionElegida
    Case "FacturacionCatalogoServicios"
        EdicionCatalogoBaseServicios "ID_Agregar"
    Case "NotaCredito"
        EdicionCajaNotaCredito "ID_Agregar"
    Case "Empleado"
        EdicionEmpleado "ID_Agregar"
    Case "Cajas"
        EdicionCaja "ID_Agregar"
    Case "PacienteCE"
        EdicionPaciente "ID_Agregar", sghConsultaExterna, 101
    End Select
End Sub

Private Sub mnuCaja_Click()
'        lcOpcionElegida = "GestionCaja"
'        If (mb_abrioCaja) Then
'            If mo_LastControl Is ucGestionCaja1 Then
'                mo_LastControl.Visible = True
'                Exit Sub
'            End If
'            mo_LastControl.Visible = False
'            ucGestionCaja1.NombreCajero = status.Panels(2).Text
'            ucGestionCaja1.Visible = True
'            Set mo_LastControl = ucGestionCaja1
'            Exit Sub
'        End If
'
'        ucGestionCaja1.idUsuario = ml_IdUsuarioAuditoria
'        ucGestionCaja1.NombreCajero = status.Panels(2).Text
'        ucGestionCaja1.lnIdTablaLISTBARITEMS = 702
'        ucGestionCaja1.lcNombrePc = lc_NombrePc
'
'
'        ConfigurarControl ucGestionCaja1

End Sub

Private Sub mnuCajaApertura_Click()
    If lcOpcionElegida = "GestionCaja" Then AperturaCaja
End Sub

Private Sub mnuCajaCierre_Click()
    If lcOpcionElegida = "GestionCaja" Then CerrarCaja
End Sub

Private Sub mnuCajasMan_Click()
   lcOpcionElegida = "Cajas"
   ConfigurarControl ucCajaLista1
End Sub

Private Sub mnuCpt_Click()
    lcOpcionElegida = "FacturacionCatalogoServicios"
    ConfigurarControl Me.ucCatalogoServiciosLista1
    
End Sub

Private Sub mnuEliminar_Click()
    Select Case lcOpcionElegida
    Case "FacturacionCatalogoServicios"
        EdicionCatalogoBaseServicios "ID_Eliminar"
    Case "NotaCredito"
        EdicionCajaNotaCredito "ID_Eliminar"
    Case "Empleado"
        EdicionEmpleado "ID_Eliminar"
    Case "Cajas"
        EdicionCaja "ID_Eliminar"
    Case "PacienteCE"
        EdicionPaciente "ID_Eliminar", sghConsultaExterna, 101
    End Select
    
End Sub

Private Sub mnuEmpleados_Click()
   lcOpcionElegida = "Empleado"
   ConfigurarControl ucEmpleadosLista1
End Sub

Private Sub mnuExportaSunat_Click()
    lcOpcionElegida = ""
    rpCajaExportaSunat.Show 1
End Sub

Private Sub mnuModificar_Click()
    Select Case lcOpcionElegida
    Case "FacturacionCatalogoServicios"
        EdicionCatalogoBaseServicios "ID_Modificar"
    Case "NotaCredito"
        EdicionCajaNotaCredito "ID_Modificar"
    Case "Empleado"
        EdicionEmpleado "ID_Modificar"
    Case "Cajas"
        EdicionCaja "ID_Modificar"
    Case "PacienteCE"
        EdicionPaciente "ID_Modificar", sghConsultaExterna, 101
    End Select

End Sub

Private Sub mnuNCredito_Click()
    lcOpcionElegida = "NotaCredito"
    ucCajaNotaCredito1.inicializar
    ConfigurarControl ucCajaNotaCredito1
End Sub

Private Sub mnuPacientes_Click()
   lcOpcionElegida = "PacienteCE"
   Me.ucPacientesLista1.TipoFiltro = sghFiltrarConHistoriasDefinitivas
   ConfigurarControl Me.ucPacientesLista1
End Sub

Private Sub mnuSalir_Click()
    End
End Sub

Private Sub mnuSalir1_Click()
    End
End Sub

Private Sub mnuVentas_Click()
    lcOpcionElegida = ""
    RpRegistroVentas.Show 1
End Sub


Sub ConfigurarControl(oControl As Control)
        
        On Error Resume Next
        
            oControl.inicializar
        
        mo_LastControl.Visible = False
        oControl.Visible = True
       
        
        Set mo_LastControl = oControl
        Form_Resize


End Sub


Private Sub tmrHora_Timer()
  status.Panels(1).Text = ""
End Sub



Sub CierreCtaAtencion()
'        Dim oCierreCtas As New CierreCtaAtencion
'        oCierreCtas.IdUsuario = ml_IdUsuarioAuditoria
'        oCierreCtas.Show 1
'        Unload oCierreCtas

End Sub

Sub EdicionConfiguracionResLab(sToolId As String) 'nuevo Samuel

End Sub


'debb-jamo
Sub EdicionTriaje(sToolId As String)
End Sub

''*******************************INO*************************************
'Sub EdicionTriajeOftalmologico(sToolId As String)
'Dim oTriajeOftalmologico As New SIGHCatalogos.clTriajeOftalomologico
'
'    Dim oRs As New ADODB.Recordset
'
'        Select Case sToolId
'        Case "ID_Agregar":
'           oTriajeOftalmologico.Opcion = sghAgregar
'        Case "ID_Modificar":
'           oTriajeOftalmologico.Opcion = sghModificar
'           oTriajeOftalmologico.idAtencion = ucAtencionesTriajeOftalmologico1.idRegistroSeleccionado
'
'           Set oRs = Me.ucAtencionesTriajeOftalmologico1.DataSource
'            If oRs Is Nothing Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            If oRs.State = 0 Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            If oRs.RecordCount = 0 Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'        Case "ID_Consultar":
'           oTriajeOftalmologico.Opcion = sghConsultar
'           oTriajeOftalmologico.idAtencion = ucAtencionesTriajeOftalmologico1.idRegistroSeleccionado
'           Set oRs = Me.ucAtencionesTriajeOftalmologico1.DataSource
'            If oRs Is Nothing Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            If oRs.State = 0 Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            If oRs.RecordCount = 0 Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'        Case "ID_Eliminar":
'           oTriajeOftalmologico.Opcion = sghEliminar
'           oTriajeOftalmologico.idAtencion = ucAtencionesTriajeOftalmologico1.idRegistroSeleccionado
'           Set oRs = Me.ucAtencionesTriajeOftalmologico1.DataSource
'            If oRs Is Nothing Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            If oRs.State = 0 Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            If oRs.RecordCount = 0 Then
'                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'        End Select
'       oTriajeOftalmologico.idUsuario = ml_IdUsuarioAuditoria
'       oTriajeOftalmologico.lcNombrePc = lc_NombrePc
'       oTriajeOftalmologico.lnIdTablaLISTBARITEMS = 1303
'       oTriajeOftalmologico.MostrarFormulario
'       Set oTriajeOftalmologico = Nothing
'       ucAtencionesTriajeOftalmologico1.RealizarBusqueda
'End Sub
''*******************************INO*************************************


Function SeleccionarOpcion(sToolId As String) As sghOpciones
        
        Select Case sToolId
        Case "ID_Agregar":
            SeleccionarOpcion = sghAgregar
        Case "ID_Modificar":
            SeleccionarOpcion = sghModificar
        Case "ID_Consultar":
            SeleccionarOpcion = sghConsultar
        Case "ID_Eliminar":
            SeleccionarOpcion = sghEliminar
        End Select

End Function

Sub EdicionTurno(sToolId As String)

End Sub

Sub EdicionEmpleado(sToolId As String)
Dim mo_EmpleadoDetalle As New SIGHCatalogos.clEmpleadoDetalle
        
        mo_EmpleadoDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_EmpleadoDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_EmpleadoDetalle.lnIdTablaLISTBARITEMS = 1301
        mo_EmpleadoDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_EmpleadoDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_EmpleadoDetalle.IdEmpleado = Me.ucEmpleadosLista1.idRegistroSeleccionado
            If mo_EmpleadoDetalle.IdEmpleado = -1 Or mo_EmpleadoDetalle.IdEmpleado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        
        mo_EmpleadoDetalle.MostrarFormulario
        Set mo_EmpleadoDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub

Sub EdicionServicio(sToolId As String)
Dim mo_ServicioDetalle As New SIGHProxies.clServicioDetalle

        

End Sub
Sub EdicionEspecialidades(sToolId As String)


End Sub

Sub EdicionMedico(sToolId As String)

End Sub
Sub EdicionAdmisionCE(sToolId As String, lTipoServicio As sghTipoServicio, lnIdTablaLISTBARITEMS As Long)
        

End Sub


Sub EdicionHistoriaClinica(sToolId As String)
End Sub

Sub EdicionMovimientoHistorias(sToolId As String)

End Sub

Sub EdicionSolicitudHistorias(sToolId As String)

End Sub
Sub EdicionArchiveroServicio(sToolId As String)

End Sub

Sub EdicionAdmisionEmergencia(sToolId As String)

End Sub

Sub EdicionAdmisionHospitalizacion(sToolId As String)

End Sub



Sub EdicionPreLiquidacion(sToolId As String)
        

End Sub
Sub EdicionDiagnosticos(sToolId As String)

End Sub
Sub EdicionTiposFinanciamiento(sToolId As String)
End Sub

Sub EdicionFuentesFinanciamiento(sToolId As String)
        
        

End Sub
Sub EdicionPartidaPresupuestal(sToolId As String)
        
        

End Sub





Sub EdicionEstablecimientosNoMinsa(sToolId As String)

End Sub



Sub EdicionFactExamenes(sToolId As String)

End Sub

Sub EdicionFactRecetas(sToolId As String)
End Sub

Sub EdicionCamas(sToolId As String, lbEsEmergencia As Boolean)
End Sub

Sub EdicionCitas(sToolId As String)

End Sub



Sub EdicionProgMedica(sToolId As String)

End Sub

Sub EdicionRoles(sToolId As String)
End Sub

Sub GenerarRecordsetDeListItems()
    
    With mrs_ListItems
          .Fields.Append "IdListItem", adInteger, 4
          .Fields.Append "Clave", adVarChar, 50
          .CursorType = adOpenStatic
          .LockType = adLockOptimistic
          .Open
    End With
    
End Sub



Private Sub ucAdmisionConsEmerg_OnClick(oRecordset As ADODB.Recordset)
    
    ml_ToolbarHeightAdd = 0
    On Error Resume Next
    If Not IsDate(oRecordset!FechaEgresoAdministrativo) Then
        ml_ToolbarHeightAdd = 500
        Select Case oRecordset!IdTipoServicio
        Case 2
'            toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = True
'            toolbar.Tools("ID_EmergenciaAObservacion").Enabled = True
'            toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = True
'            toolbar.Tools("ID_EmergenciaTransferencias").Enabled = True
        Case 4
'            toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = True
'            toolbar.Tools("ID_EmergenciaAObservacion").Enabled = False
'            toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = True
'            toolbar.Tools("ID_EmergenciaTransferencias").Enabled = True
        End Select
    Else
'            toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = False
'            toolbar.Tools("ID_EmergenciaAObservacion").Enabled = False
'            toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = False
'            toolbar.Tools("ID_EmergenciaTransferencias").Enabled = False
    End If

End Sub

Private Sub ucAdmisionHospitalizacion_OnClick(oRecordset As ADODB.Recordset)
    ml_ToolbarHeightAdd = 0
    On Error Resume Next
    
    If Not IsDate(oRecordset!FechaEgresoAdministrativo) Then
'        ml_ToolbarHeightAdd = 500
'        toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = True
'        toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = True
'        toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = True
    Else
'        toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = False
'        toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = False
'        toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = False
    End If
End Sub

Sub EdicionCatalogoBaseBienesInsumos(sToolId As String)

End Sub
Sub EdicionCatalogoBienesInsumos(sToolId As String)

End Sub

Sub EdicionCatalogoBaseServicios(sToolId As String)
Dim mo_CatalogoServiciosDetalle As New SIGHCatalogos.clCatalogoBaseServicDet
    
    mo_CatalogoServiciosDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_CatalogoServiciosDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_CatalogoServiciosDetalle.lnIdTablaLISTBARITEMS = 610
    mo_CatalogoServiciosDetalle.lcNombrePc = lc_NombrePc
    Select Case mo_CatalogoServiciosDetalle.Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
        mo_CatalogoServiciosDetalle.idProducto = Me.ucCatalogoServiciosLista1.idRegistroSeleccionado
        If mo_CatalogoServiciosDetalle.idProducto = -1 Or mo_CatalogoServiciosDetalle.idProducto = 0 Then
            MsgBox "Seleccione un registro", vbInformation, Me.Caption
            Exit Sub
        End If
    End Select

     mo_CatalogoServiciosDetalle.MostrarFormulario
     Set mo_CatalogoServiciosDetalle = Nothing

    Select Case sToolId
    Case "ID_Agregar":
    Case "ID_Modificar":
    Case "ID_Consultar":
    Case "ID_Eliminar":
    End Select

End Sub
Sub EdicionCatalogoServicios(sToolId As String)
Dim mo_CatalogoServiciosDetalle As New SIGHCatalogos.clCatalogoServicioDetalle
    
    mo_CatalogoServiciosDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_CatalogoServiciosDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_CatalogoServiciosDetalle.TipoCatalogo = Me.ucCatalogoServiciosLista1.IdTipoCatalogo
    
    Select Case mo_CatalogoServiciosDetalle.Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
        Exit Sub
    End Select

    mo_CatalogoServiciosDetalle.MostrarFormulario
    Set mo_CatalogoServiciosDetalle = Nothing

    Select Case sToolId
    Case "ID_Agregar":
    Case "ID_Modificar":
    Case "ID_Consultar":
    Case "ID_Eliminar":
    End Select

End Sub


Sub EdicionCentrosCosto(sToolId As String)

End Sub


Sub AperturaCaja()
'Dim oApertura As New AperturaDecaja
'Dim oDOEmpleado As DOEmpleado
'Dim sNombreCajero As String
'Dim oRsPermisos As New Recordset
'Dim lbUsuarioRealizaApertura As Boolean
'        '
'        Set oRsPermisos = mo_AdminSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_IdUsuarioAuditoria)
'        If oRsPermisos.RecordCount > 0 Then
'           Do While Not oRsPermisos.EOF
'              Select Case oRsPermisos.Fields!IdPermiso
'              Case 201    'Caja - Realizar Apertura
'                   lbUsuarioRealizaApertura = True
'              End Select
'              oRsPermisos.MoveNext
'           Loop
'        End If
'        Set oRsPermisos = Nothing
'        '
'        If lbUsuarioRealizaApertura = True Then
'            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(ml_IdUsuarioAuditoria)
'            sNombreCajero = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
'            oApertura.NombreCajero = sNombreCajero
'            oApertura.idUsuario = ml_IdUsuarioAuditoria
'            oApertura.lcNombrePc = lc_NombrePc
'            oApertura.Show 1
'            If oApertura.AperturoCajaOK = True Then
'
'                'debb-15/03/2016 (inicio)
'                If oApertura.IdTurno = 0 Then
'                   MsgBox "Tiene problemas con el TURNO", vbInformation, ""
'                   Exit Sub
'                End If
'                'debb-15/03/2016 (fin)
'
'                mb_abrioCaja = Me.ucGestionCaja1.RealizarAperturaDeCaja(ml_IdUsuarioAuditoria, oApertura.IdCaja, oApertura.IdTurno, oApertura.EmiteSoloServicio)
'
'                '/****************************INO***************************************/
'                mb_abrioCaja = Me.ucGestionDevolucion2.RealizarAperturaDeCaja(ml_IdUsuarioAuditoria, oApertura.IdCaja, oApertura.IdTurno, oApertura.EmiteSoloServicio)
'                '/****************************INO***************************************/
'
'                'mgaray201503
'                Set moDOCajaGestion = New DOCajaGestion
'                moDOCajaGestion.IdCaja = oApertura.IdCaja
'                moDOCajaGestion.IdCajero = oApertura.IdTurno
'                lbCajeroEmiteSoloServicios = oApertura.EmiteSoloServicio
'            End If
'            Unload oApertura
'
'        Else
'            MsgBox "El Usuario no tiene permiso para realizar APERTURA DE CAJA", vbInformation, Me.Caption
'        End If

End Sub

Sub CerrarCaja()
'Dim oRsPermisos As New Recordset
'Dim lbUsuarioRealizaCierre As Boolean
'        '
'        Set oRsPermisos = mo_AdminSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_IdUsuarioAuditoria)
'        If oRsPermisos.RecordCount > 0 Then
'           Do While Not oRsPermisos.EOF
'              Select Case oRsPermisos.Fields!IdPermiso
'              Case 202    'Caja - Realizar Apertura
'                   lbUsuarioRealizaCierre = True
'              End Select
'              oRsPermisos.MoveNext
'           Loop
'        End If
'        Set oRsPermisos = Nothing
'        '
'        If lbUsuarioRealizaCierre = True Then
'
'            If Not mb_abrioCaja Then
'                Exit Sub
'            End If
'            If MsgBox("¿Esta seguro de realizar el CIERRE DE CAJA ?", vbYesNo, Me.Caption) = vbYes Then
'                If ucGestionCaja1.RealizarCierreDeCaja() Then
'                    mb_abrioCaja = False
'                End If
'
'                '/******************************INO*************************************
'                 If ucGestionDevolucion2.RealizarCierreDeCaja() Then
'                    mb_abrioCaja = False
'                End If
'                '/******************************INO*************************************
'
'            Else
'                ucGestionCaja1.MuestraTabEmisionDocumentos (False)
'                mb_abrioCaja = False
'            End If
'        Else
'            MsgBox "El USUARIO no tiene permiso para realizar el  CIERRE"
'        End If
End Sub


Private Sub ucCajeroServicios1_HizoClickEnEscape()
    
    mo_LastControl.Visible = False

End Sub

Private Sub ucEstadoCuenta1_HizoClickEnCancelar()
    mo_LastControl.Visible = False

End Sub

Sub EdicionOrdenesServicio(sToolId As String)

End Sub

Sub EdicionOrdenesServicioPatologiaClinica(sToolId As String)

End Sub

Sub EdicionOrdenesServicioAnatomiaPatologia(sToolId As String)

End Sub


Sub EdicionOrdenesServicioImagenologia(sToolId As String)

End Sub

Sub EdicionOrdenesServicioSalaOperaciones(sToolId As String)

End Sub

Sub EdicionOrdenesServicioFarmacia(sToolId As String)

End Sub

Sub ImprimirParteDiario()
'Dim oRptCaja As New RptCaja
'
'    oRptCaja.IdGestionCaja = Me.ucGestionCaja1.IdGestionCaja
'
'    If oRptCaja.IdGestionCaja <> -1 Then
'        oRptCaja.CrearParteDiario Me.hwnd
'    End If
    
End Sub

Sub ImprimirConsolidadoServicio()
'Dim oRptCaja As New RptCaja
'
'    oRptCaja.IdGestionCaja = Me.ucGestionCaja1.IdGestionCaja
'
'    If oRptCaja.IdGestionCaja <> -1 Then
'        oRptCaja.CrearReporteConsolidadoServicios
'    End If
    
End Sub

Sub EdicionCaja(sToolId As String)
Dim mo_cajaDetalle As New SIGHCatalogos.clCajaDetalle
        
        mo_cajaDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_cajaDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_cajaDetalle.lnIdTablaLISTBARITEMS = 705
        mo_cajaDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_cajaDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_cajaDetalle.IdCaja = Me.ucCajaLista1.idRegistroSeleccionado
            If mo_cajaDetalle.IdCaja = -1 Or mo_cajaDetalle.IdCaja = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_cajaDetalle.MostrarFormulario
        Set mo_cajaDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select
        
End Sub

Sub EdicionInventario(sToolId As String)
End Sub
'**debb2014
Sub EdicionNS(sToolId As String, lbNSsoloParaFarmacia As Boolean)
End Sub
'**debb2014
Sub EdicionNI(sToolId As String, lbNIsoloParaFarmacia As Boolean)
End Sub
Sub EdicionIntervencionS(sToolId As String)
End Sub

Sub EdicionVentas(sToolId As String)
End Sub

Sub EdicionDependenciaExt(sToolId As String)
End Sub

Sub EdicionRayosX(sToolId As String)
End Sub

Sub EdicionImagIngresos(sToolId As String)
End Sub

Sub EdicionImagSalidas(sToolId As String)
End Sub

Sub EdicionImagEcografiaObs(sToolId As String)
End Sub

Sub EdicionImagEcografiaGen(sToolId As String)
        Dim mo_EcogGen As New SIGHImagen.EcogGen
        mo_EcogGen.Opcion = SeleccionarOpcion(sToolId)
        mo_EcogGen.idUsuario = ml_IdUsuarioAuditoria
        mo_EcogGen.lcNombrePc = lc_NombrePc
        mo_EcogGen.lnIdTablaLISTBARITEMS = 1317
        Select Case mo_EcogGen.Opcion
        Case sghAgregar
             If UcImagenesLista1.SeEligioGridBoleta = True Then
                mo_EcogGen.idMovimiento = UcImagenesLista1.idRegistroSeleccionado
                mo_EcogGen.SeEligioGridBoleta = UcImagenesLista1.SeEligioGridBoleta
             End If
        Case sghModificar, sghConsultar, sghEliminar
            If UcImagenesLista1.SeEligioGridBoleta = True Then
            Else
               mo_EcogGen.idMovimiento = UcImagenesLista1.idRegistroSeleccionado
            End If
            If UcImagenesLista1.idRegistroSeleccionado = -1 Or UcImagenesLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_EcogGen.MostrarFormulario
        UcImagenesLista1.RealizarBusqueda
        UcImagenesLista1.SeEligioGridBoleta = False
End Sub

Sub EdicionImagTomografia(sToolId As String)
        Dim mo_tomog As New SIGHImagen.Tomog
        mo_tomog.Opcion = SeleccionarOpcion(sToolId)
        mo_tomog.idUsuario = ml_IdUsuarioAuditoria
        mo_tomog.lnIdTablaLISTBARITEMS = 1319
        mo_tomog.lcNombrePc = lc_NombrePc
        Select Case mo_tomog.Opcion
        Case sghAgregar
             If UcImagenesLista1.SeEligioGridBoleta = True Then
                mo_tomog.idMovimiento = UcImagenesLista1.idRegistroSeleccionado
                mo_tomog.SeEligioGridBoleta = UcImagenesLista1.SeEligioGridBoleta
             End If
        Case sghModificar, sghConsultar, sghEliminar
            mo_tomog.idMovimiento = UcImagenesLista1.idRegistroSeleccionado
            If UcImagenesLista1.idRegistroSeleccionado = -1 Or UcImagenesLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_tomog.MostrarFormulario
        UcImagenesLista1.RealizarBusqueda
        UcImagenesLista1.SeEligioGridBoleta = False
End Sub

Sub EdicionLaboratorio(sToolId As String)
  Dim mo_laboratorio As New SIGHLaboratorio.laboratorio
  mo_laboratorio.Opcion = SeleccionarOpcion(sToolId)
  mo_laboratorio.idUsuario = ml_IdUsuarioAuditoria
  mo_laboratorio.PuntoCarga = 2
  mo_laboratorio.lnIdTablaLISTBARITEMS = 1312
  mo_laboratorio.lcNombrePc = lc_NombrePc
  mo_laboratorio.AreaTrabajo = ucFactOrdenesLaboratorio.AreaTrabajo
  Select Case mo_laboratorio.Opcion
  Case sghAgregar
       If ucFactOrdenesLaboratorio.SeEligioGridBoleta = True Then
          mo_laboratorio.idMovimiento = ucFactOrdenesLaboratorio.idRegistroSeleccionado
          mo_laboratorio.SeEligioGridBoleta = ucFactOrdenesLaboratorio.SeEligioGridBoleta
       End If
  Case sghModificar, sghConsultar, sghEliminar
       If ucFactOrdenesLaboratorio.SeEligioGridBoleta = True Then
       Else
           mo_laboratorio.idMovimiento = ucFactOrdenesLaboratorio.idRegistroSeleccionado
       End If
       If ucFactOrdenesLaboratorio.idRegistroSeleccionado = -1 Or ucFactOrdenesLaboratorio.idRegistroSeleccionado = 0 Then
          MsgBox "Seleccione un registro", vbInformation, Me.Caption
          Exit Sub
       End If
  End Select
  mo_laboratorio.MostrarFormulario
  ucFactOrdenesLaboratorio.RealizarBusqueda
  ucFactOrdenesLaboratorio.SeEligioGridBoleta = False
End Sub

Sub EdicionOrdenesServicioPatologiaClinica_(sToolId As String)

End Sub

'Frank 29042015
Sub EdicionOrdenesServicioAnatomiaPatologia_(sToolId As String)
  Dim mo_laboratorio As New SIGHLaboratorio.laboratorio
  mo_laboratorio.Opcion = SeleccionarOpcion(sToolId)
  mo_laboratorio.idUsuario = ml_IdUsuarioAuditoria
  mo_laboratorio.PuntoCarga = 3
  mo_laboratorio.lnIdTablaLISTBARITEMS = 1312
  mo_laboratorio.lcNombrePc = lc_NombrePc
  mo_laboratorio.AreaTrabajo = ucFacturacionOrdenesPatologia.AreaTrabajo
  Select Case mo_laboratorio.Opcion
  Case sghAgregar
       If ucFacturacionOrdenesPatologia.SeEligioGridBoleta = True Then
          mo_laboratorio.idMovimiento = ucFacturacionOrdenesPatologia.idRegistroSeleccionado
          mo_laboratorio.SeEligioGridBoleta = ucFacturacionOrdenesPatologia.SeEligioGridBoleta
       End If
  Case sghModificar, sghConsultar, sghEliminar
       If ucFacturacionOrdenesPatologia.SeEligioGridBoleta = True Then
       Else
           mo_laboratorio.idMovimiento = ucFacturacionOrdenesPatologia.idRegistroSeleccionado
       End If
       If ucFacturacionOrdenesPatologia.idRegistroSeleccionado = -1 Or ucFacturacionOrdenesPatologia.idRegistroSeleccionado = 0 Then
          MsgBox "Seleccione un registro", vbInformation, Me.Caption
          Exit Sub
       End If
  End Select
  mo_laboratorio.MostrarFormulario
  ucFacturacionOrdenesPatologia.RealizarBusqueda
  ucFacturacionOrdenesPatologia.SeEligioGridBoleta = False
End Sub

Sub EdicionOrdenesBS_(sToolId As String)
  Dim mo_laboratorio As New SIGHLaboratorio.laboratorio
  mo_laboratorio.Opcion = SeleccionarOpcion(sToolId)
  mo_laboratorio.idUsuario = ml_IdUsuarioAuditoria
  mo_laboratorio.PuntoCarga = 11
  mo_laboratorio.lnIdTablaLISTBARITEMS = 1312
  mo_laboratorio.lcNombrePc = lc_NombrePc
  mo_laboratorio.AreaTrabajo = ucFacturacionBS.AreaTrabajo
  Select Case mo_laboratorio.Opcion
  Case sghAgregar
       If ucFacturacionBS.SeEligioGridBoleta = True Then
          mo_laboratorio.idMovimiento = ucFacturacionBS.idRegistroSeleccionado
          mo_laboratorio.SeEligioGridBoleta = ucFacturacionBS.SeEligioGridBoleta
       End If
  Case sghModificar, sghConsultar, sghEliminar
    If ucFacturacionBS.SeEligioGridBoleta = True Then
    Else
       mo_laboratorio.idMovimiento = ucFacturacionBS.idRegistroSeleccionado
    End If
    If ucFacturacionBS.idRegistroSeleccionado = -1 Or ucFacturacionBS.idRegistroSeleccionado = 0 Then
      MsgBox "Seleccione un registro", vbInformation, Me.Caption
      Exit Sub
    End If
  End Select
  mo_laboratorio.MostrarFormulario
  ucFacturacionBS.RealizarBusqueda
End Sub

Sub EdicionResultados(sToolId As String)
  Dim mo_LabIngresos As New SIGHLaboratorio.laboratorio
  mo_LabIngresos.Opcion = SeleccionarOpcion(sToolId)
  mo_LabIngresos.idUsuario = ml_IdUsuarioAuditoria
  Select Case mo_LabIngresos.Opcion
  Case sghAgregar
  Case sghModificar, sghConsultar, sghEliminar
    mo_LabIngresos.idMovimiento = UcLabIngresos1.idRegistroSeleccionado
    If UcLabIngresos1.idRegistroSeleccionado = -1 Or UcLabIngresos1.idRegistroSeleccionado = 0 Then
      MsgBox "Seleccione un registro", vbInformation, Me.Caption
      Exit Sub
    End If
  End Select
  mo_LabIngresos.MostrarFormulario
  UcLabIngresos1.RealizarBusqueda
End Sub

Sub EdicionMuestras(sToolId As String)
  Dim mo_LabSalidas As New SIGHLaboratorio.laboratorio
  mo_LabSalidas.Opcion = SeleccionarOpcion(sToolId)
  mo_LabSalidas.idUsuario = ml_IdUsuarioAuditoria
  Select Case mo_LabSalidas.Opcion
  Case sghAgregar
  Case sghModificar, sghConsultar, sghEliminar
    mo_LabSalidas.idMovimiento = UcLabSalidas1.idRegistroSeleccionado
    If UcLabSalidas1.idRegistroSeleccionado = -1 Or UcLabSalidas1.idRegistroSeleccionado = 0 Then
      MsgBox "Seleccione un registro", vbInformation, Me.Caption
      Exit Sub
    End If
  End Select
  mo_LabSalidas.MostrarFormulario
  UcLabSalidas1.RealizarBusqueda
End Sub

Sub EdicionLabIngresos(sToolId As String)
  Dim mo_LabIngresos As New SIGHLaboratorio.Ingresos
  mo_LabIngresos.Opcion = SeleccionarOpcion(sToolId)
  mo_LabIngresos.idUsuario = ml_IdUsuarioAuditoria
  mo_LabIngresos.idPuntoCarga = UcLabIngresos1.PuntoCarga
  mo_LabIngresos.lnIdTablaLISTBARITEMS = 1313
  mo_LabIngresos.lcNombrePc = lc_NombrePc
  If UcLabIngresos1.PuntoCarga = -1 Or UcLabIngresos1.PuntoCarga = 0 Then
    MsgBox "Escoja un punto de Carga para Agregar/Modificar un registro de Ingreso de Insumos.", vbInformation, Me.Caption
    Exit Sub
  End If
  Select Case mo_LabIngresos.Opcion
  Case sghAgregar
  Case sghModificar, sghConsultar, sghEliminar
    mo_LabIngresos.idMovimiento = UcLabIngresos1.idRegistroSeleccionado
    If UcLabIngresos1.idRegistroSeleccionado = -1 Or UcLabIngresos1.idRegistroSeleccionado = 0 Then
      MsgBox "Seleccione un registro para Modificar Ingreso de Insumos.", vbInformation, Me.Caption
      Exit Sub
    End If
  End Select
  mo_LabIngresos.MostrarFormulario
  UcLabIngresos1.RealizarBusqueda
End Sub

Sub EdicionLabSalidas(sToolId As String)
  Dim mo_LabSalidas As New SIGHLaboratorio.Salidas
  mo_LabSalidas.Opcion = SeleccionarOpcion(sToolId)
  mo_LabSalidas.idUsuario = ml_IdUsuarioAuditoria
  mo_LabSalidas.idPuntoCarga = UcLabSalidas1.PuntoCarga
  mo_LabSalidas.lnIdTablaLISTBARITEMS = 1314
  mo_LabSalidas.lcNombrePc = lc_NombrePc
  If UcLabSalidas1.PuntoCarga = -1 Or UcLabSalidas1.PuntoCarga = 0 Then
    MsgBox "Escoja un punto de Carga para Agregar/Modificar un registro de Salida de Insumos", vbInformation, Me.Caption
    Exit Sub
  End If
  Select Case mo_LabSalidas.Opcion
  Case sghAgregar
  Case sghModificar, sghConsultar, sghEliminar
    mo_LabSalidas.idMovimiento = UcLabSalidas1.idRegistroSeleccionado
    If UcLabSalidas1.idRegistroSeleccionado = -1 Or UcLabSalidas1.idRegistroSeleccionado = 0 Then
      MsgBox "Seleccione un registro para Modificar Salida de Insumos", vbInformation, Me.Caption
      Exit Sub
    End If
  End Select
  mo_LabSalidas.MostrarFormulario
  UcLabSalidas1.RealizarBusqueda
End Sub

Sub EdicionAlojados(sToolId As String)

End Sub


Sub EdicionReembolsos(sToolId As String)
End Sub

Sub EdicionMovimientoFormatoHC(sToolId As String)
End Sub

Sub EdicionConstancias(sToolId As String)
End Sub

Sub EdicionPacExtConSeguro(sToolId As String)
End Sub

'Sub EdicionPacExtParticular(sToolId As String)
'        Dim oFacGeneraCtaPacienteExterno As New FacGeneraCtaPacienteExterno
'        oFacGeneraCtaPacienteExterno.Opcion = SeleccionarOpcion(sToolId)
'        oFacGeneraCtaPacienteExterno.idUsuario = ml_IdUsuarioAuditoria
'        oFacGeneraCtaPacienteExterno.lnIdTablaLISTBARITEMS = 1340
'        oFacGeneraCtaPacienteExterno.lcNombrePc = lc_NombrePc
'        oFacGeneraCtaPacienteExterno.idPuntoCarga = 6  'Consulta externa -admision
'        Select Case oFacGeneraCtaPacienteExterno.Opcion
'        Case sghAgregar
'        Case sghModificar, sghConsultar, sghEliminar
'            If ucPacienteExternos1.IdRegistroSeleccionado = 0 Then
'                MsgBox "Seleccione un registro", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            oFacGeneraCtaPacienteExterno.idAtencion = ucPacienteExternos1.IdRegistroSeleccionado
'        End Select
'        oFacGeneraCtaPacienteExterno.Show 1
'        Unload oFacGeneraCtaPacienteExterno
'
'End Sub

Sub EdicionPaqueteServicio(sToolId As String)
End Sub

Sub EdicionDespachoDonaciones(sToolId As String)
        Dim mo_DespachoDonaciones As New SighFarmacia.DespachoDonaciones
        Dim lcMovimiento As String
        lcMovimiento = Right("0" + Trim(Str(ucFarmDespachoDonaciones1.idRegistroSeleccionado)), 9)
        mo_DespachoDonaciones.Opcion = SeleccionarOpcion(sToolId)
        mo_DespachoDonaciones.idUsuario = ml_IdUsuarioAuditoria
        mo_DespachoDonaciones.lnIdTablaLISTBARITEMS = 1342
        mo_DespachoDonaciones.lcNombrePc = lc_NombrePc
        Select Case mo_DespachoDonaciones.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_DespachoDonaciones.movNumero = lcMovimiento
            If ucFarmDespachoDonaciones1.idRegistroSeleccionado = -1 Or ucFarmDespachoDonaciones1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_DespachoDonaciones.MostrarFormulario
        ucFarmDespachoDonaciones1.RealizarBusqueda
End Sub


'debb-hra
Private Sub cmdFechaHoraServidor_Click()
  'CentrarImagen

  Dim lcBuscaParametro As New SIGHDatos.Parametros
  status.Panels(1).Text = "      " & lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQL1
  'status.Panels(7).Text = lcBuscaParametro.SeleccionaFilaParametro(314) & " " & lcBuscaParametro.RetornaVersionServidorSQLserver
  'status.Panels(7).Width = 3400
  Set lcBuscaParametro = Nothing
End Sub


Sub EdicionHisCE(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
End Sub

Sub EdicionHisDobleDigitacion(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
    If sToolId = "ID_ExportaHIS" Or sToolId = "ID_ExportaURENIS" Then Exit Sub
    Dim oRcsTemp1 As New ADODB.Recordset
    Set oRcsTemp1 = mo_ReglasHIS.ObtenerListaEstablecimientosMR
    If oRcsTemp1.RecordCount = 0 Then
        MsgBox "No ha registrado los establecimientos de la MicroRed", vbExclamation, Me.Caption
        Exit Sub
    End If
    Dim mo_HISDetalle As New SIGHhisDigitacion.MantRegHisCalidad
    mo_HISDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_HISDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_HISDetalle.lcNombrePc = lc_NombrePc
    mo_HISDetalle.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS
    mo_HISDetalle.IdHisDetalle = UcHISCalidad.idRegistroSeleccionado
    
    If mo_HISDetalle.IdHisDetalle = -1 Or mo_HISDetalle.IdHisDetalle = 0 Then
        MsgBox "Seleccione un Registro", vbInformation, Me.Caption
        Exit Sub
    End If
    Select Case mo_HISDetalle.Opcion
    Case sghAgregar
        If UcHISCalidad.Registrado = 1 Then
            MsgBox "Seleccione la opción Modificar(F3) para editar la doble digitación", vbInformation, Me.Caption
            Exit Sub
        End If
        mo_HISDetalle.MostrarFormulario
        UcHISCalidad.CargarListaGenerados
    Case sghModificar, sghConsultar
        If UcHISCalidad.Registrado = -1 Or UcHISCalidad.Registrado = 0 Then
            MsgBox "Seleccione la opción Agregar(F2) para registrar la doble digitación", vbInformation, Me.Caption
            Exit Sub
        End If
        mo_HISDetalle.MostrarFormulario
        UcHISCalidad.CargarListaGenerados
    Case sghEliminar
        MsgBox "No puedes eliminar el registro para la doble digitación", vbInformation, Me.Caption
        Exit Sub
    End Select
    'Frank HIS
    Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
    End Select
End Sub


'JVG - Muestra el formulario de edicion del los Lotes HIS en el sistema
Sub EdicionHisLotesCE(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)

End Sub

Sub EdicionProgramacionHIS(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)

End Sub

Sub EdicionReceta(sToolId As String, lnIdListBarItems As Long, lnIdTipoServicio As Long)
End Sub

'debb 26/7/12
Sub EdicionFua(sToolId As String)
End Sub

Sub EdicionTipoTarifa(sToolId As String)

End Sub


'JVG - Muestra el formulario de edicion los Establecimientos de la MicroRed
Sub EdicionHisEstablecimientos(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)

End Sub

Sub EdicionPadronNominal(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)


End Sub

Sub EdicionMantenedorFarmacia(sToolId As String)
        Dim mo_FarmAlmacen As New SighFarmacia.clAlmacen
         
        mo_FarmAlmacen.Opcion = SeleccionarOpcion(sToolId)
        mo_FarmAlmacen.idUsuario = ml_IdUsuarioAuditoria
        mo_FarmAlmacen.lnIdTablaLISTBARITEMS = 1355
        mo_FarmAlmacen.lcNombrePc = lc_NombrePc
        Select Case mo_FarmAlmacen.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_FarmAlmacen.IdDependenciaExt = ucFarmAlmacenes1.idRegistroSeleccionado
            If ucFarmAlmacenes1.idRegistroSeleccionado = -1 Or ucFarmAlmacenes1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_FarmAlmacen.MostrarFormulario
        ucFarmAlmacenes1.RealizarBusqueda
End Sub

'mgaray201411f
Sub EdicionTipoModalidadSala(sToolId As String)
        
End Sub

Sub EdicionSala(sToolId As String)

        
End Sub

Sub EdicionImagFactCatalogoServiciosDuracion(sToolId As String)
End Sub

Sub EdicionIntegracionSistema(sToolId As String)
End Sub


'debb2014b
Sub EdicionMantenedorHistoricoPrecios(sToolId As String)
        Dim mo_FarmHistPrecio As New SighFarmacia.clFarmHistPrecios
         
        mo_FarmHistPrecio.Opcion = SeleccionarOpcion(sToolId)
        mo_FarmHistPrecio.idUsuario = ml_IdUsuarioAuditoria
        mo_FarmHistPrecio.lnIdTablaLISTBARITEMS = 1359
        mo_FarmHistPrecio.lcNombrePc = lc_NombrePc
        Select Case mo_FarmHistPrecio.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_FarmHistPrecio.IdFarmHistPrecio = ucFarmHpreciosLista1.idRegistroSeleccionado
            If ucFarmHpreciosLista1.idRegistroSeleccionado = -1 Or ucFarmHpreciosLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_FarmHistPrecio.MostrarFormulario
        ucFarmHpreciosLista1.RealizarBusqueda
End Sub

Sub OcultarOpcionesPacticularesMenuPorEstablecimiento()
'toolbar.Index
End Sub

'mgaray201504
Private Function UsuarioEsCajero(mb_UsuarioAccesoGestionCaja As Boolean) As Boolean
    UsuarioEsCajero = False
    
    If mb_UsuarioAccesoGestionCaja = True Then
        Dim oRsPermisos As New Recordset
        Dim lbUsuarioRealizaApertura As Boolean
        
        Set oRsPermisos = mo_AdminSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_IdUsuarioAuditoria)
        If oRsPermisos.RecordCount > 0 Then
           Do While Not oRsPermisos.EOF
              Select Case oRsPermisos.Fields!IdPermiso
              Case 201    'Caja - Realizar Apertura
                   UsuarioEsCajero = True
              End Select
              oRsPermisos.MoveNext
           Loop
           
        End If
        Set oRsPermisos = Nothing
    End If
    
End Function

'FRANK MAYO
Sub EdicionCajaNotaCredito(sToolId As String)
        Dim mo_CajaApruebaNotaCredito As New CajaApruebaNotaCredito
        Dim orsNotasCredito As New Recordset
        mo_CajaApruebaNotaCredito.idUsuario = ml_IdUsuarioAuditoria
        mo_CajaApruebaNotaCredito.Opcion = SeleccionarOpcion(sToolId)
        mo_CajaApruebaNotaCredito.lnIdTablaLISTBARITEMS = 1206
        mo_CajaApruebaNotaCredito.lcNombrePc = lc_NombrePc
        mo_CajaApruebaNotaCredito.idTipoNota = 2 'NOTA CREDITO
        Select Case mo_CajaApruebaNotaCredito.Opcion
        Case sghAgregar
            'mo_AdmisionHospDetalle.TipoServicio = sghEmergenciaConsultorios
        Case sghModificar, sghConsultar, sghEliminar
            Set orsNotasCredito = ucCajaNotaCredito1.DataSource
            If orsNotasCredito.State = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If orsNotasCredito.RecordCount = 0 Then
                MsgBox "No existen registros", vbInformation, Me.Caption
                Exit Sub
            End If
            If ucCajaNotaCredito1.idRegistroSeleccionado = 0 Then
                Exit Sub
            End If
            Set orsNotasCredito = Nothing
            mo_CajaApruebaNotaCredito.idRegistroSeleccionado = Me.ucCajaNotaCredito1.idRegistroSeleccionado
            If mo_CajaApruebaNotaCredito.idRegistroSeleccionado = -1 Or mo_CajaApruebaNotaCredito.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_CajaApruebaNotaCredito.Show 1
End Sub

Private Sub ucHISEstablecimientos_GotFocus()

End Sub

Sub EdicionPaciente(sToolId As String, lTipoServicio As sghTipoServicio, lnIdTablaLISTBARITEMS As Long)
Dim mo_PacienteDetalle As New PacienteDetalle
        
        
        mo_PacienteDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_PacienteDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_PacienteDetalle.TipoServicio = lTipoServicio
        mo_PacienteDetalle.lcNombrePc = lc_NombrePc
        mo_PacienteDetalle.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS

        Select Case mo_PacienteDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_PacienteDetalle.idPaciente = Me.ucPacientesLista1.idRegistroSeleccionado
            If mo_PacienteDetalle.idPaciente = -1 Or mo_PacienteDetalle.idPaciente = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        
        mo_PacienteDetalle.Icon = Me.Icon
        mo_PacienteDetalle.Show 1
        Unload mo_PacienteDetalle

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
            Dim doPaciente As New doPaciente
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub

