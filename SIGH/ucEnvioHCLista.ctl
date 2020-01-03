VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucEnvioHCLista 
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   LockControls    =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   10065
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      Height          =   1515
      Left            =   75
      TabIndex        =   10
      Top             =   540
      Width           =   9975
      Begin VB.TextBox txtNroHistoria 
         Height          =   315
         Left            =   5910
         TabIndex        =   3
         Top             =   450
         Width           =   1845
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8400
         Picture         =   "ucEnvioHCLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   420
         Width           =   585
      End
      Begin VB.TextBox txtApellidoPaterno 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtApellidoMaterno 
         Height          =   315
         Left            =   2070
         TabIndex        =   1
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtPrimerNombre 
         Height          =   315
         Left            =   3990
         TabIndex        =   2
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtIdEnvio 
         Height          =   315
         Left            =   5910
         TabIndex        =   7
         Top             =   1020
         Width           =   1845
      End
      Begin MSMask.MaskEdBox txtFechaSolicitud 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   1020
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaPrestamo 
         Height          =   315
         Left            =   2070
         TabIndex        =   5
         Top             =   1020
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaDevolucion 
         Height          =   315
         Left            =   3990
         TabIndex        =   6
         Top             =   1020
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Apellido paterno                 Apellido materno                Primer nombre                   Nº Historia clínica"
         Height          =   225
         Left            =   180
         TabIndex        =   12
         Top             =   240
         Width           =   7635
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha solicitud                  Fecha prestamo                 Fecha devolución               ID Envío"
         Height          =   225
         Left            =   180
         TabIndex        =   11
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdPrestamosHC 
      Height          =   3765
      Left            =   75
      TabIndex        =   9
      Top             =   2145
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6641
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108864
      Caption         =   "Lista de historias clínicas enviadas"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Envíos de Historia Clínica"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   13
      Top             =   15
      Width           =   10140
   End
End
Attribute VB_Name = "ucEnvioHCLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdPrestamosHC.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdPrestamosHC.DataSource
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
Property Let TipoBusqueda(lValue As sghTipoBusquedaPrestamoHistoria)
    ml_TipoBusqueda = lValue
End Property
Property Get TipoBusqueda() As sghTipoBusquedaPrestamoHistoria
    TipoBusqueda = ml_TipoBusqueda
End Property


Private Sub btnBuscar_Click()
Dim oPaciente As New doPaciente
Dim oPrestamo As New DOPrestamoHistoriaClinica
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtPrimerNombre = "" And UserControl.txtNroHistoria = "") Then
        End If
            
        
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPaciente.PrimerNombre = UserControl.txtPrimerNombre
        oPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
        oPrestamo.FechaSolicitud = IIf(UserControl.txtFechaSolicitud = SIGHComun.FECHA_VACIA_DMY, 0, UserControl.txtFechaSolicitud)
        oPrestamo.FechaPrestamoRequerida = IIf(UserControl.txtFechaPrestamo = SIGHComun.FECHA_VACIA_DMY, 0, UserControl.txtFechaPrestamo)
        oPrestamo.FechaDevolucion = IIf(UserControl.txtFechaDevolucion = SIGHComun.FECHA_VACIA_DMY, 0, UserControl.txtFechaDevolucion)
        oPrestamo.IdEnvio = Val(UserControl.txtIdEnvio)
        
        Select Case ml_TipoBusqueda
        Case sghTodasHistorias
            Set grdPrestamosHC.DataSource = mo_AdminArchivoClinico.PrestamosHistoriaClinicaFiltrar(oPaciente, oPrestamo)
        Case sghHistoriaEnPrestamo
            Set grdPrestamosHC.DataSource = mo_AdminArchivoClinico.PrestamosHistoriaClinicaFiltrarEnviados(oPaciente, oPrestamo)
        End Select
        If mo_AdminArchivoClinico.MensajeError <> "" Then
            MsgBox mo_AdminArchivoClinico.MensajeError, vbCritical, "Filtro PrestamosHC"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdPrestamosHC, SIGHComun.GrillaConFilasBicolor

End Sub

Private Sub grdPrestamosHC_Click()
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdPrestamosHC.DataSource
    On Error Resume Next
    Select Case ml_TipoBusqueda
    Case sghTodasHistorias
        ml_IdRegistroSeleccionado = rsRecordset("IdPrestamo")
    Case sghHistoriaEnPrestamo
        ml_IdRegistroSeleccionado = rsRecordset("IdEnvio")
    End Select
    
End Sub


Private Sub grdPrestamosHC_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdPrestamosHC.Bands(0).Columns("IdPrestamo").Hidden = True
    grdPrestamosHC.Bands(0).Columns("IdEnvio").Hidden = True
    
    grdPrestamosHC.Bands(0).Columns("HistoriaClinica").Header.Caption = "Nro Historia"
    grdPrestamosHC.Bands(0).Columns("HistoriaClinica").Width = 1000
    
    grdPrestamosHC.Bands(0).Columns("Nombres").Header.Caption = "Apellidos y Nombres"
    grdPrestamosHC.Bands(0).Columns("Nombres").Width = 3000
    
    grdPrestamosHC.Bands(0).Columns("FechaSolicitud").Header.Caption = "Fecha Sol."
    grdPrestamosHC.Bands(0).Columns("FechaSolicitud").Width = 1500
    
    grdPrestamosHC.Bands(0).Columns("FechaPrestamo").Header.Caption = "Fecha Prest."
    grdPrestamosHC.Bands(0).Columns("FechaPrestamo").Width = 1500
    
    grdPrestamosHC.Bands(0).Columns("FechaDevolucion").Header.Caption = "Fecha Devol."
    grdPrestamosHC.Bands(0).Columns("FechaDevolucion").Width = 1500

End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdPrestamosHC.Width = fraBusqueda.Width
   grdPrestamosHC.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub









