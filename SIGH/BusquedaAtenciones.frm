VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form BusquedaAtenciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de atenciones"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   30
      TabIndex        =   3
      Top             =   4410
      Width           =   13755
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "BusquedaAtenciones.frx":0000
         DownPicture     =   "BusquedaAtenciones.frx":0460
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
         Left            =   5505
         Picture         =   "BusquedaAtenciones.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BusquedaAtenciones.frx":0D4A
         DownPicture     =   "BusquedaAtenciones.frx":120E
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
         Left            =   7050
         Picture         =   "BusquedaAtenciones.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdBusqueda 
      Height          =   4365
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   7699
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de atenciones"
   End
End
Attribute VB_Name = "BusquedaAtenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Lista atenciones de un Paciente
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idCuentaAtencion  As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_EstadoCuenta As String
Property Let EstadoCuenta(lValue As String)
    ml_EstadoCuenta = lValue
End Property
Property Get EstadoCuenta() As String
    EstadoCuenta = ml_EstadoCuenta
End Property

Property Set Atenciones(oValue As Recordset)
    Set grdBusqueda.DataSource = oValue
End Property

Property Get Atenciones() As Recordset
End Property

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property

Property Get idCuentaAtencion() As Long
    idCuentaAtencion = ml_idCuentaAtencion
End Property

Private Sub btnAceptar_Click()
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    ml_idCuentaAtencion = 0
    Me.Visible = False
End Sub

Private Sub Form_Load()
    ml_idCuentaAtencion = 0
End Sub


Private Sub grdBusqueda_DblClick()
    Dim rsRecordset As Recordset
    Set rsRecordset = grdBusqueda.DataSource
    ml_idCuentaAtencion = rsRecordset("IdCuentaAtencion")
    ml_EstadoCuenta = rsRecordset("EstadoCuenta")
    btnAceptar_Click
End Sub

Private Sub grdBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        
    
    grdBusqueda.Bands(0).Columns("IdPaciente").Hidden = True
    grdBusqueda.Bands(0).Columns("idTipoNumeracion").Hidden = True
    grdBusqueda.Bands(0).Columns("IdServicioIngreso").Hidden = True
        
    grdBusqueda.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "N° Cuenta"
    grdBusqueda.Bands(0).Columns("IdCuentaAtencion").Width = 800
    
    grdBusqueda.Bands(0).Columns("EstadoCuenta").Header.Caption = "Estado Cuenta"
    grdBusqueda.Bands(0).Columns("EstadoCuenta").Width = 1200
    
    'grdBusqueda.Bands(0).Columns("IdAtencion").Header.Caption = "N° atención"
    'grdBusqueda.Bands(0).Columns("IdAtencion").Width = 1300
    grdBusqueda.Bands(0).Columns("IdAtencion").Hidden = True
    
    grdBusqueda.Bands(0).Columns("FechaIngreso").Header.Caption = "Fec.Ingreso"
    grdBusqueda.Bands(0).Columns("FechaIngreso").Width = 1000
    
'    grdBusqueda.Bands(0).Columns("HoraIngreso").Header.Caption = "Hr Ingreso"
'    grdBusqueda.Bands(0).Columns("HoraIngreso").Width = 600
    grdBusqueda.Bands(0).Columns("HoraIngreso").Hidden = True
        
      
    grdBusqueda.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "N° HC"
    grdBusqueda.Bands(0).Columns("NroHistoriaClinica").Width = 1200
    grdBusqueda.Bands(0).Columns("NroHistoriaClinica").Hidden = True
    
    grdBusqueda.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdBusqueda.Bands(0).Columns("ApellidoPaterno").Width = 1500
    grdBusqueda.Bands(0).Columns("ApellidoPaterno").Hidden = True
    
    grdBusqueda.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdBusqueda.Bands(0).Columns("ApellidoMaterno").Width = 1500
    grdBusqueda.Bands(0).Columns("ApellidoMaterno").Hidden = True
    
    grdBusqueda.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdBusqueda.Bands(0).Columns("PrimerNombre").Width = 1500
    grdBusqueda.Bands(0).Columns("PrimerNombre").Hidden = True

    grdBusqueda.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdBusqueda.Bands(0).Columns("SegundoNombre").Width = 1500
    grdBusqueda.Bands(0).Columns("SegundoNombre").Hidden = True

    grdBusqueda.Bands(0).Columns("ServicioIngreso").Header.Caption = "Servicio Ing"
    grdBusqueda.Bands(0).Columns("ServicioIngreso").Width = 2500
    
    grdBusqueda.Bands(0).Columns("Edad").Header.Caption = "Edad"
    grdBusqueda.Bands(0).Columns("Edad").Width = 600
    
    grdBusqueda.Bands(0).Columns("FechaEgreso").Header.Caption = "F.Egreso"
    grdBusqueda.Bands(0).Columns("FechaEgreso").Width = 1000
    
    grdBusqueda.Bands(0).Columns("HoraEgreso").Header.Caption = "Hr.Egreso"
    grdBusqueda.Bands(0).Columns("HoraEgreso").Width = 700


    grdBusqueda.Bands(0).Columns("Diagnostico").Header.Caption = "Diagnóstico"
    grdBusqueda.Bands(0).Columns("Diagnostico").Width = 5000

    grdBusqueda.Bands(0).Columns("dTipoServicio").Width = 800

    grdBusqueda.Bands(0).Columns("IdEstado").Hidden = True
    grdBusqueda.Bands(0).Columns("IdTipoServicio").Hidden = True
    
    mo_Apariencia.ConfigurarFilasBiColores grdBusqueda, sighentidades.GrillaConFilasBicolor
    
End Sub

Private Sub grdBusqueda_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
   If KeyAscii = 13 Then
      grdBusqueda_DblClick
   End If
End Sub
