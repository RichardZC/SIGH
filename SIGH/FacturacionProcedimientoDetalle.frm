VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FacturacionProcedimientoDetalle 
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   1665
   ClientTop       =   750
   ClientWidth     =   11520
   Icon            =   "FacturacionProcedimientoDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   11520
   Begin VB.CommandButton btnBusquedaServicio 
      Caption         =   "..."
      Height          =   315
      Left            =   2790
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2580
      Width           =   345
   End
   Begin VB.CommandButton btnBusquedaProcedimiento 
      Caption         =   ".."
      Height          =   315
      Left            =   2760
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3570
      Width           =   345
   End
   Begin VB.CommandButton btnBusquedaMedico 
      Caption         =   "..."
      Height          =   315
      Left            =   2790
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2970
      Width           =   345
   End
   Begin VB.Frame fraProcedimiento 
      Caption         =   "Procedimientos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   40
      Top             =   1950
      Width           =   11415
      Begin VB.TextBox txtIdServicio 
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
         Left            =   1680
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox lblDescServicio 
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
         Left            =   3150
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   630
         Width           =   5325
      End
      Begin VB.TextBox lblDescMedicoOrdena 
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
         Left            =   3150
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   990
         Width           =   5325
      End
      Begin VB.TextBox txtIdMedicoOrdena 
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
         Left            =   1680
         TabIndex        =   20
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox txtNroOrden 
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
         Left            =   1680
         TabIndex        =   15
         Top             =   210
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtHoraOrden 
         Height          =   315
         Left            =   5310
         TabIndex        =   18
         Top             =   240
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
      Begin MSMask.MaskEdBox txtFechaOrden 
         Height          =   315
         Left            =   3870
         TabIndex        =   17
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
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
      Begin VB.Label Label63 
         Caption         =   "Servicio ordena"
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
         TabIndex        =   29
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label65 
         Caption         =   "Fecha orden"
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
         Left            =   2790
         TabIndex        =   16
         Top             =   270
         Width           =   1530
      End
      Begin VB.Label Label66 
         Caption         =   "Médico ordena"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   26
         Top             =   1020
         Width           =   1350
      End
      Begin VB.Label Label69 
         Caption         =   "Orden Nro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   1260
      End
   End
   Begin VB.Frame fraAddProcedimiento 
      Height          =   945
      Left            =   60
      TabIndex        =   39
      Top             =   3390
      Width           =   11415
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "FacturacionProcedimientoDetalle.frx":0CCA
         DownPicture     =   "FacturacionProcedimientoDetalle.frx":1055
         Height          =   315
         Left            =   2670
         Picture         =   "FacturacionProcedimientoDetalle.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   540
         Width           =   1005
      End
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "FacturacionProcedimientoDetalle.frx":1779
         DownPicture     =   "FacturacionProcedimientoDetalle.frx":1BAB
         Height          =   315
         Left            =   1590
         Picture         =   "FacturacionProcedimientoDetalle.frx":1FDD
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtIdProcedimiento 
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
         Left            =   1650
         TabIndex        =   21
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox lblDescProcedimiento 
         Enabled         =   0   'False
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
         Left            =   3090
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Width           =   8175
      End
      Begin VB.Label Label5 
         Caption         =   "Procedimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   38
      Top             =   7890
      Width           =   11355
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacturacionProcedimientoDetalle.frx":422E
         DownPicture     =   "FacturacionProcedimientoDetalle.frx":468E
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
         Left            =   4305
         Picture         =   "FacturacionProcedimientoDetalle.frx":4B03
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacturacionProcedimientoDetalle.frx":4F78
         DownPicture     =   "FacturacionProcedimientoDetalle.frx":543C
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
         Left            =   5850
         Picture         =   "FacturacionProcedimientoDetalle.frx":5928
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   45
      TabIndex        =   37
      Top             =   885
      Width           =   11430
      Begin VB.TextBox txtIdNroHistoria 
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
         Left            =   3915
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
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
         Left            =   5130
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   3315
      End
      Begin VB.TextBox lblNroCuenta 
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
         Left            =   1695
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   255
         Width           =   1140
      End
      Begin VB.TextBox lblPaciente 
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
         Left            =   1695
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   645
         Width           =   4020
      End
      Begin VB.TextBox lblFechaIngreso 
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
         Left            =   9825
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Width           =   1425
      End
      Begin VB.TextBox lblServicioIngreso 
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
         Left            =   7200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   630
         Width           =   4065
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Cuenta"
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
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label2 
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
         Left            =   165
         TabIndex        =   10
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Ingreso"
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
         Left            =   8535
         TabIndex        =   8
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Nº historia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Servicio Ingreso"
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
         Left            =   5805
         TabIndex        =   12
         Top             =   675
         Width           =   1305
      End
   End
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
      Height          =   885
      Left            =   45
      TabIndex        =   36
      Top             =   0
      Width           =   11430
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   3120
         Picture         =   "FacturacionProcedimientoDetalle.frx":5E14
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   330
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
         Left            =   1695
         TabIndex        =   1
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label50 
         Caption         =   "Nro Historia"
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
         TabIndex        =   0
         Top             =   390
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList lstOpciones 
      Left            =   240
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionProcedimientoDetalle.frx":8A5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionProcedimientoDetalle.frx":8E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionProcedimientoDetalle.frx":934C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionProcedimientoDetalle.frx":9763
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin UltraGrid.SSUltraGrid grdProcedimientos 
      Height          =   3435
      Left            =   60
      TabIndex        =   35
      Top             =   4380
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6059
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
      Caption         =   "Lista de procedimientos"
   End
End
Attribute VB_Name = "FacturacionProcedimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POAtencionesInterconsultas
'        Autor: William Castro Grijalva
'        Fecha: 31/10/2004 09:32:29 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_Diagnosticos As New Collection
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdPreFacturacionProcedimiento As Long
Dim ml_IdTipoServicio As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHComun.ListaDespleglable
Dim mo_ProcedimientoDetalle As New Collection
Dim mo_Procedimiento As New DOAtencionProcedimiento
Dim mrs_Procedimientos As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mrs_ProcedimientosEliminados As New Recordset
Dim mo_FacturacionServicios As New Collection 'WCG20060317 (para los servicios a facturar)
Dim oCuentaAtencion As New DOCuentaAtencion 'WCG20060317 (para los datos de la cuenta del paciente)
Dim ml_IdDepartamento As Long 'WCG20060321 (para diferenciar los procedimientos)



Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdPreFacturacionProcedimiento(lValue As Long)
   ml_IdPreFacturacionProcedimiento = lValue
End Property
Property Get IdPreFacturacionProcedimiento() As Long
   IdPreFacturacionProcedimiento = ml_IdPreFacturacionProcedimiento
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
       
       mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
       mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

End Sub

Private Sub btnAgregarDx_Click()
    'Validamos que no exista el procedimiento en la lista de procedimientos
    If Val(Me.txtIdProcedimiento.Tag) <= 0 Then
        Exit Sub
    End If
    If mrs_Procedimientos.EOF = False Or mrs_Procedimientos.BOF = False Then
        mrs_Procedimientos.MoveFirst
        Do While Not mrs_Procedimientos.EOF
            If mrs_Procedimientos.Fields!IdProcedimiento = Val(Me.txtIdProcedimiento.Tag) Then
                mrs_Procedimientos.MoveFirst
                Exit Sub
            End If
            mrs_Procedimientos.MoveNext
        Loop
        mrs_Procedimientos.MoveFirst
    End If
    
    With mrs_Procedimientos
        .AddNew
        .Fields!IdProcedimiento = Val(Me.txtIdProcedimiento.Tag)
        .Fields!CodigoCPT = Me.txtIdProcedimiento
        .Fields!Descripcion = Me.lblDescProcedimiento
        .Fields!IdMedicoRealiza = 0
        .Fields!NombreMedico = ""
        .Fields!IdServicioRealiza = 0
        .Fields!NombreServicio = ""
        .Fields!FechaRealizacion = 0
        .Fields!HoraRealizacion = ""
        .Fields!IdFacturacionServicio = 0
        .Fields!EstadoRegistro = "A"
    End With
End Sub

Private Sub btnBuscar_Click()


Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion
Dim lIdCuentaAtencionActual As Long
    
    LimpiarDatosDeAtencion
    If (Me.txtNroHistoria) = "" Then
        MsgBox "Ingrese la Historia Clínica a buscar", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Dim rsCuentasAtencion As New ADODB.Recordset
    Dim iCount As Integer

    lIdCuentaAtencionActual = 0
    Set rsCuentasAtencion = mo_AdminCaja.ObtenerCuentasAtencionPorHistoriaClinica(Val(Me.txtNroHistoria))
    iCount = 0
    Do While Not rsCuentasAtencion.EOF
        iCount = iCount + 1
        lIdCuentaAtencionActual = rsCuentasAtencion!IdCuentaAtencion
        rsCuentasAtencion.MoveNext
    Loop
    If iCount > 1 Then
        'Levantamos el formulario para seleccionar la cuenta de atención
        Dim oFrmCuentasAtencion As New CuentasAtencionSeleccionar
        Set oFrmCuentasAtencion.DataSource = rsCuentasAtencion
        oFrmCuentasAtencion.Show vbModal
        If oFrmCuentasAtencion.BotonPresionado = sghCancelar Then
            lIdCuentaAtencionActual = 0
        Else
            lIdCuentaAtencionActual = oFrmCuentasAtencion.IdRegistroSeleccionado
        End If
    End If
    RecuperarDatosCuentaAtencion lIdCuentaAtencionActual


End Sub
Private Sub RecuperarDatosCuentaAtencion(lIdCuentaAtencion As Long)
Dim rsPaciente As New Recordset
Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion
        
    'oDOPaciente.NroHistoriaClinica = Val(Me.cmbNroHistoriaBusqueda.Text)
    oDOCuentaAtencion.IdCuentaAtencion = lIdCuentaAtencion
    
    Screen.MousePointer = vbHourglass
    Set rsPaciente = mo_AdminAdmision.AtencionesFiltrarPacientesParaIngresarProcedimientos(oDOPaciente, oDOCuentaAtencion)
    Screen.MousePointer = vbDefault
    
    'cmbNroHistoriaBusqueda.BoundColumn = ""
    'Set cmbNroHistoriaBusqueda.ListSource = rsPaciente
    
    'Si hay una sola coincidencia
    If rsPaciente.RecordCount = 1 Then
        rsPaciente.MoveFirst
        LimpiarDatosDeAtencion
        
        Me.txtIdNroHistoria.Text = rsPaciente!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsPaciente!IdTipoNumeracion
        Me.lblFechaIngreso = rsPaciente!FechaIngreso
        Me.lblServicioIngreso = rsPaciente!ServicioIngreso
        Me.lblPaciente = rsPaciente!ApellidoPaterno + " " + rsPaciente!ApellidoMaterno + " " + rsPaciente!PrimerNombre + " " + ("" & rsPaciente!SegundoNombre)
        Me.lblNroCuenta = rsPaciente!IdCuentaAtencion
    
        Me.txtNroOrden.SetFocus
    ElseIf rsPaciente.RecordCount > 1 Then
        'cmbNroHistoriaBusqueda.ShowDropDown
        
    ElseIf rsPaciente.RecordCount = 0 Then
        MsgBox "No se encontraron atenciones para el nro de historia o nro de cuenta ingresado", vbInformation, Me.Caption
        LimpiarDatosDeAtencion
    End If
End Sub
Private Sub cmbNroHistoria_Click()
End Sub
Sub LimpiarDatosDeAtencion()
        
        Me.txtIdNroHistoria.Text = ""
        mo_cmbIdTipoGenHistoriaClinica.BoundText = ""
        Me.lblFechaIngreso = ""
        Me.lblServicioIngreso = ""
        Me.lblPaciente = ""
        Me.lblNroCuenta = ""

End Sub
Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long)
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection

    oBusqueda.IdEspecialidad = lIdEspecialidad
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.IdMedico
            lblNombreMedico.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub

Private Sub btnBusquedaMedico_Click()
       CompletarDatosDeMedico txtIdMedicoOrdena, lblDescMedicoOrdena, Val(Me.lblDescServicio.Tag)
End Sub

Private Sub btnBusquedaServicio_Click()
        CompletarDatosDeServicio txtIdServicio, lblDescServicio
End Sub

Private Sub btnQuitarDx_Click()
    Dim doFacturacionServicio As DOFacturacionServicios
    On Error Resume Next
    With mrs_Procedimientos
        If Not .EOF And Not .BOF Then
            If mrs_Procedimientos!IdAtencionProcDetalle <> 0 Then
                'Verificamos que el detalle esté como emitido para poder eliminarse
                Set doFacturacionServicio = mo_AdminFacturacion.FacturacionServiciosSeleccionarPorId(mrs_Procedimientos!IdFacturacionServicio)
                If Not doFacturacionServicio Is Nothing Then
                    If Not (doFacturacionServicio.IdEstadoFacturacion = sghEstadoFacturacion.sghPendientePago And doFacturacionServicio.TotalPorPagar = 0) Then
                        If doFacturacionServicio.IdEstadoFacturacion = sghEstadoFacturacion.sghPendientePago Then
                            MsgBox "No se puede eliminar el item seleccionado por que ya se encuentra en proceso de facturación [Con un importe de S/. " & doFacturacionServicio.TotalPorPagar & " ]", vbExclamation, Me.Caption
                        Else
                            MsgBox "No se puede eliminar el item seleccionado por que ya se encuentra Facturado", vbExclamation, Me.Caption
                        End If
                        Exit Sub
                    End If
                End If
                mrs_ProcedimientosEliminados.AddNew
                mrs_ProcedimientosEliminados!IdAtencionProcDetalle = mrs_Procedimientos!IdAtencionProcDetalle
                mrs_ProcedimientosEliminados!IdFacturacionServicio = mrs_Procedimientos!IdFacturacionServicio
            End If
           .Delete
           .Update
        End If
        .MoveFirst
    End With
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = Me.cmbIdTipoGenHistoriaClinica
End Sub

Private Sub toolProcedimientos_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub txtFechaOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaOrden
End Sub
Private Sub txtFechaOrden_LostFocus()

       If txtFechaOrden <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaOrden, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                 txtFechaOrden = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
        mo_Formulario.MarcarComoVacio txtFechaOrden
End Sub

Private Sub txtFechaOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraOrden_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraOrden
End Sub


Private Sub txtHoraOrden_LostFocus()
    If txtHoraOrden <> SIGHComun.HORA_VACIA_HM Then
         If Not SIGHComun.ValidaHora(txtHoraOrden) Then
             MsgBox "La hora ingresada no es válida", vbInformation, "Datos de paciente"
             txtHoraOrden = SIGHComun.HORA_VACIA_HM
         End If
     End If
   mo_Formulario.MarcarComoVacio txtHoraOrden
End Sub

Private Sub txtHoraOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdMedicoOrdena_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoOrdena
    If KeyCode = vbKeyF1 Then
        btnBusquedaMedico_Click
    End If
End Sub


Private Sub txtIdMedicoOrdena_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoOrdena, lblDescMedicoOrdena
    mo_Formulario.MarcarComoVacio txtIdMedicoOrdena
End Sub

Private Sub txtIdMedicoOrdena_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CompletarDatosDeMedicoEnElLostFocus(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oMedicosEspecialidad As New Collection

    txtMedico = Trim(txtMedico)
    If txtMedico <> "" Then
        Dim oDOEmpleado As New dOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo(CStr(txtMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            txtMedico.Tag = oDoMedico.IdMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
        End If
    End If
    
End Sub

Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicio_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicio, lblDescServicio
    mo_Formulario.MarcarComoVacio txtIdServicio
End Sub

Private Sub txtIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroHistoria_LostFocus()
   mo_Formulario.MarcarComoVacio txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

    mo_Formulario.HabilitarDeshabilitar lblNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar lblPaciente, False
    mo_Formulario.HabilitarDeshabilitar lblFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar lblServicioIngreso, False
    

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
 
    Select Case mi_Opcion
    Case sghAgregar
    Case sghModificar
            Me.fraBusqueda.Enabled = False
    Case sghConsultar
            Me.fraBusqueda.Enabled = False
            Me.fraProcedimiento.Enabled = False
            Me.fraAddProcedimiento.Enabled = False
            Me.grdProcedimientos.Enabled = False
            'WCG comentado por facturacion
            'Me.ucProcedimientoDetalle1.BotonAgregarEnabled = False
            'Me.ucProcedimientoDetalle1.BotonQuitarEnabled = False
            Me.btnAceptar.Enabled = False
            
    Case sghEliminar
            Me.fraBusqueda.Enabled = False
            Me.fraProcedimiento.Enabled = False
            Me.fraAddProcedimiento.Enabled = False
            Me.grdProcedimientos.Enabled = False
            'WCG comentado por facturacion
            'Me.ucProcedimientoDetalle1.BotonAgregarEnabled = False
            'Me.ucProcedimientoDetalle1.BotonQuitarEnabled = False
    
    End Select
 
 'WCG comentado por facturacion
 'Me.ucProcedimientoDetalle1.TipoServicio = ml_IdTipoServicio
 
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar orden de procedimiento"
       Case sghModificar
           Me.Caption = "Modificar orden de procedimiento"
       Case sghConsultar
           Me.Caption = "Consultar orden de procedimiento"
       Case sghEliminar
           Me.Caption = "Eliminar orden de procedimiento"
       End Select
        GenerarRecordsetTemporal
        CargarComboBoxes
        CargarDatosAlFormulario
        mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
        Case vbKeyF6
            btnBuscar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String
   ValidarDatosObligatorios = False
   
    If Me.lblNroCuenta = "" Then
        MsgBox "Seleccione el paciente", vbInformation, Me.Caption
        Exit Function
    End If
    
    If txtNroOrden = "" Then
        MsgBox "Ingrese el nro de orden de procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtIdMedicoOrdena = "" Then
        MsgBox "Ingrese el médico que ordena el procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtFechaOrden = SIGHComun.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de orden del procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtHoraOrden = SIGHComun.HORA_VACIA_HM Then
        MsgBox "Ingrese la hora de orden del procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
    If txtFechaOrden < CDate(Me.lblFechaIngreso) Then
        MsgBox "La fecha de la orden del procedimiento no puede ser menor que la fecha de ingreso de la atención", vbExclamation, Me.Caption
        Exit Function
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

    'WCG Comentado por facturacion
   mo_Procedimiento.IdCuentaAtencion = Val(Me.lblNroCuenta)
   
   CargarProcedimientosAlObjetoDatos mo_Procedimiento, mo_ProcedimientoDetalle
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminFacturacion.AtencionProcedimientosAgregar(mo_Procedimiento, mo_ProcedimientoDetalle)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminFacturacion.AtencionProcedimientosModificar(mo_Procedimiento, mo_ProcedimientoDetalle, mrs_ProcedimientosEliminados)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminFacturacion.AtencionProcedimientosEliminar(mo_Procedimiento, mo_ProcedimientoDetalle)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    
    '1ro
    Dim oDOPreFacturacionProcedimiento As New DOAtencionProcedimiento
    Set oDOPreFacturacionProcedimiento = mo_AdminFacturacion.AtencionProcedimientosSeleccionarPorId(Me.IdPreFacturacionProcedimiento)
    If Not oDOPreFacturacionProcedimiento Is Nothing Then
        CargarDatosDelaAtencion oDOPreFacturacionProcedimiento.IdCuentaAtencion
        mb_ExistenDatos = True
    Else
        mb_ExistenDatos = False
    End If
    
    '2do
    CargarDatosDeDeProcedimientos
   
End Sub
Sub CargarDatosDelaAtencion(lIdCuentaAtencion As Long)
Dim oDOPaciente As New doPaciente
Dim rsPaciente As New ADODB.Recordset
Dim oDOCuentaAtencion As New DOCuentaAtencion
    
    oDOPaciente.NroHistoriaClinica = 0
    oDOCuentaAtencion.IdCuentaAtencion = lIdCuentaAtencion
    Set rsPaciente = mo_AdminAdmision.AtencionesFiltrarPacientesParaIngresarProcedimientos(oDOPaciente, oDOCuentaAtencion)
    
    'Si hay una sola coincidencia
    If Not (rsPaciente.EOF And rsPaciente.BOF) Then
        LimpiarDatosDeAtencion
        Me.txtIdNroHistoria.Text = rsPaciente!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsPaciente!IdTipoNumeracion
        Me.lblFechaIngreso = rsPaciente!FechaIngreso
        Me.lblServicioIngreso = rsPaciente!ServicioIngreso
        Me.lblPaciente = rsPaciente!ApellidoPaterno + " " + rsPaciente!ApellidoMaterno + " " + rsPaciente!PrimerNombre + " " + ("" & rsPaciente!SegundoNombre)
        Me.lblNroCuenta = rsPaciente!IdCuentaAtencion
    End If
    rsPaciente.Close

End Sub
'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
   
End Sub
Sub GenerarRecordsetTemporal()
    
    With mrs_Procedimientos
        .Fields.Append "IdAtencionProcDetalle", adInteger
        .Fields.Append "IdProcedimiento", adInteger
        .Fields.Append "CodigoCPT", adVarChar, 10
        .Fields.Append "Descripcion", adVarChar, 255
        .Fields.Append "FechaRealizacion", adChar, 10
        .Fields.Append "HoraRealizacion", adChar, 5
        .Fields.Append "IdMedicoRealiza", adInteger, , adFldIsNullable
        .Fields.Append "NombreMedico", adVarChar, 100, adFldIsNullable
        .Fields.Append "IdServicioRealiza", adInteger, , adFldIsNullable
        .Fields.Append "NombreServicio", adVarChar, 100, adFldIsNullable
        .Fields.Append "IdFacturacionServicio", adInteger, , adFldIsNullable
        .Fields.Append "EstadoRegistro", adChar, 1
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    'Para los procedimientos eliminados
    With mrs_ProcedimientosEliminados
        .Fields.Append "IdAtencionProcDetalle", adInteger
        .Fields.Append "IdFacturacionServicio", adInteger, , adFldIsNullable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set grdProcedimientos.DataSource = mrs_Procedimientos
    
End Sub

Public Sub CargarDatosDeDeProcedimientos()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDOPreFacturacionProcedimiento As New DOAtencionProcedimiento

    'Carga datos de la cabecera
    Dim rsProcedimiento As New Recordset
    Set oDOPreFacturacionProcedimiento = mo_AdminFacturacion.AtencionProcedimientosSeleccionarPorId(Me.IdPreFacturacionProcedimiento)
    
    If oDOPreFacturacionProcedimiento.IdAtencionProcedimiento = 0 Then
        MsgBox "No existe datos de procedimientos", vbInformation, Me.Caption
        Exit Sub
    End If
    
    txtFechaOrden = oDOPreFacturacionProcedimiento.FechaOrden
    txtHoraOrden = oDOPreFacturacionProcedimiento.HoraOrden
    txtNroOrden = oDOPreFacturacionProcedimiento.NroOrden

    'Completa datos de medico
    If mo_AdminProgramacion.MedicosSeleccionarPorId(oDOPreFacturacionProcedimiento.IdMedicoOrdena, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
        txtIdMedicoOrdena.Text = oDOEmpleado.CodigoPlanilla
        txtIdMedicoOrdena.Tag = oDoMedico.IdMedico
        lblDescMedicoOrdena = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    
    Me.txtIdServicio.Tag = IIf(oDOPreFacturacionProcedimiento.IdServicioOrdena = 0, "", oDOPreFacturacionProcedimiento.IdServicioOrdena)
    Dim oDOServicio As New DOServicio
    If Me.txtIdServicio.Tag <> "" Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oDOPreFacturacionProcedimiento.IdServicioOrdena)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicio.Text = oDOServicio.Codigo
            Me.lblDescServicio = oDOServicio.Nombre
        End If
    End If
  
    Dim rsProcedimientos As New Recordset
    Set rsProcedimientos = mo_AdminFacturacion.AtencionProcedimientoDetalleSeleccionarPorIdPreFacturacionProcedimiento(Me.IdPreFacturacionProcedimiento)
    Do While Not rsProcedimientos.EOF
        With mrs_Procedimientos
            .AddNew
            .Fields!IdAtencionProcDetalle = rsProcedimientos!IdAtencionProcDetalle
            .Fields!IdProcedimiento = rsProcedimientos!IdProcedimiento
            .Fields!CodigoCPT = rsProcedimientos!CodigoCPT
            .Fields!Descripcion = rsProcedimientos!Descripcion
            .Fields!IdMedicoRealiza = rsProcedimientos!IdMedicoRealiza
            .Fields!NombreMedico = rsProcedimientos!NombreMedico
            .Fields!IdServicioRealiza = rsProcedimientos!IdServicioRealiza
            .Fields!NombreServicio = rsProcedimientos!NombreServicio
            .Fields!FechaRealizacion = Format(rsProcedimientos!FechaRealizacion, "dd/mm/yyyy")
            .Fields!HoraRealizacion = rsProcedimientos!HoraRealizacion
            .Fields!IdFacturacionServicio = rsProcedimientos!IdFacturacionServicio
            .Fields!EstadoRegistro = "M"
        End With
        rsProcedimientos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores grdProcedimientos, SIGHComun.GrillaConFilasBicolor
    
End Sub

Sub CargarProcedimientosAlObjetoDatos(oProcedimiento As DOAtencionProcedimiento, oProcedimientoDetalle As Collection)
    Dim oDOProcedimiento As DOProcedimiento
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ProcedimientoS
    '---------------------------------------------------------------------------------
    'Datos de la cabecera
    oProcedimiento.IdAtencionProcedimiento = Me.IdPreFacturacionProcedimiento
    oProcedimiento.IdCuentaAtencion = Val(lblNroCuenta)
    oProcedimiento.IdMedicoOrdena = Val(txtIdMedicoOrdena.Tag)
    oProcedimiento.IdServicioOrdena = Val(Me.txtIdServicio.Tag)
    oProcedimiento.FechaOrden = txtFechaOrden.Text
    oProcedimiento.HoraOrden = txtHoraOrden.Text
    oProcedimiento.NroOrden = txtNroOrden.Text
    oProcedimiento.IdUsuarioAuditoria = ml_IdUsuario
    
    'Datos del detalle
    Dim oFacturacionProcDetalle As DOAtencionProcDetalle
    Dim oFacturacionServicios As DOFacturacionServicios 'WCG20060317
    If Not (mrs_Procedimientos.BOF And mrs_Procedimientos.EOF) Then
        Set oFacturacionProcDetalle = New DOAtencionProcDetalle
        mrs_Procedimientos.MoveFirst
        Do While Not mrs_Procedimientos.EOF
            Set oFacturacionProcDetalle = New DOAtencionProcDetalle
            
            oFacturacionProcDetalle.IdAtencionProcDetalle = mrs_Procedimientos!IdAtencionProcDetalle
            oFacturacionProcDetalle.IdAtencionProcedimiento = Me.IdPreFacturacionProcedimiento
            oFacturacionProcDetalle.FechaRealizacion = IIf(Trim(mrs_Procedimientos!FechaRealizacion) <> "__/__/____" And Trim(mrs_Procedimientos!FechaRealizacion) <> "", mrs_Procedimientos!FechaRealizacion, 0)
            oFacturacionProcDetalle.HoraRealizacion = mrs_Procedimientos!HoraRealizacion
            oFacturacionProcDetalle.IdMedicoRealiza = IIf(IsNull(mrs_Procedimientos!IdMedicoRealiza), 0, mrs_Procedimientos!IdMedicoRealiza)
            oFacturacionProcDetalle.IdProcedimiento = mrs_Procedimientos!IdProcedimiento
            oFacturacionProcDetalle.IdServicioRealiza = IIf(IsNull(mrs_Procedimientos!IdServicioRealiza), 0, mrs_Procedimientos!IdServicioRealiza)
            oFacturacionProcDetalle.IdUsuarioAuditoria = ml_IdUsuario
            oFacturacionProcDetalle.IdFacturacionServicio = IIf(IsNull(mrs_Procedimientos!IdFacturacionServicio), 0, mrs_Procedimientos!IdFacturacionServicio)
            oFacturacionProcDetalle.EstadoRegistro = mrs_Procedimientos!EstadoRegistro
            oProcedimientoDetalle.Add oFacturacionProcDetalle
            mrs_Procedimientos.MoveNext
        Loop
    End If
    
End Sub

Private Sub grdProcedimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdProcedimientos.Bands(0).Columns("IdAtencionProcDetalle").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdProcedimiento").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdFacturacionServicio").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdMedicoRealiza").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdServicioRealiza").Hidden = True
    
    grdProcedimientos.Bands(0).Columns("CodigoCPT").Header.Caption = "CPT"
    grdProcedimientos.Bands(0).Columns("CodigoCPT").Width = 1000
    
    'grdProcedimientos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdProcedimientos.Bands(0).Columns("Descripcion").Width = 10000
    
    'grdProcedimientos.Bands(0).Columns("FechaRealizacion").Header.Caption = "Fecha"
    grdProcedimientos.Bands(0).Columns("FechaRealizacion").Hidden = True
    
    'grdProcedimientos.Bands(0).Columns("HoraRealizacion").Header.Caption = "Hora"
    grdProcedimientos.Bands(0).Columns("HoraRealizacion").Hidden = True
    
    'grdProcedimientos.Bands(0).Columns("NombreMedico").Header.Caption = "Médico"
    grdProcedimientos.Bands(0).Columns("NombreMedico").Hidden = True

    'grdProcedimientos.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
    grdProcedimientos.Bands(0).Columns("NombreServicio").Hidden = True
    grdProcedimientos.Bands(0).Columns("EstadoRegistro").Hidden = True


End Sub
Private Sub btnBusquedaProcedimiento_Click()
Dim oBusqueda As New ProcedimientosBusqueda
Dim oDOProcedimiento As DOProcedimiento
    oBusqueda.IdDiferenciacion = ml_IdDepartamento
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOProcedimiento Is Nothing Then
            If oDOProcedimiento.IdProducto = 0 Then
                MsgBox "El procedimiento ingresado no se encuentra en el catálogo de servicios (Tarifario)", vbInformation, Me.Caption
                txtIdProcedimiento.Tag = ""
                txtIdProcedimiento.Text = ""
                lblDescProcedimiento = ""
            Else
                txtIdProcedimiento.Text = Trim(oDOProcedimiento.CodigoCPT2004)
                txtIdProcedimiento.Tag = oDOProcedimiento.IdProcedimiento
                lblDescProcedimiento = oDOProcedimiento.Descripcion
            End If
        Else
            txtIdProcedimiento.Text = ""
            txtIdProcedimiento.Tag = ""
            lblDescProcedimiento = ""
        End If
    End If
    
End Sub

Private Sub txtIdProcedimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdProcedimiento
End Sub

Private Sub txtIdProcedimiento_LostFocus()

    txtIdProcedimiento.Text = UCase(txtIdProcedimiento.Text)

    If txtIdProcedimiento.Text <> "" Then
        Dim oDOProcedimiento As DOProcedimiento
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorCodigoCPT(txtIdProcedimiento.Text)
        If Not oDOProcedimiento Is Nothing Then
            If oDOProcedimiento.IdProducto = 0 Then
                MsgBox "El procedimiento ingresado no se encuentra en el catálogo de servicios (Tarifario)", vbInformation, Me.Caption
                txtIdProcedimiento.Tag = ""
                txtIdProcedimiento.Text = ""
                lblDescProcedimiento = ""
            Else
                txtIdProcedimiento.Tag = oDOProcedimiento.IdProcedimiento
                lblDescProcedimiento = oDOProcedimiento.Descripcion
            End If
        Else
            txtIdProcedimiento.Tag = ""
            lblDescProcedimiento = ""
        End If
    Else
        txtIdProcedimiento.Tag = ""
        lblDescProcedimiento = ""
    End If
   'mo_Formulario.MarcarComoVacio txtIdProcedimiento
End Sub

Private Sub txtIdProcedimiento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtIdProcedimiento_LostFocus
        btnAgregarDx_Click
        Exit Sub
    End If
    
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New ServiciosBusqueda
Dim oDOServicio As New DOServicio

    oBusqueda.IdTipoServicio = 0
    oBusqueda.HabilitarTipoServicio = True
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Text = oDOServicio.Codigo
            txtIdServicio.Tag = oDOServicio.IdServicio
            lblDescripcionServicio = oDOServicio.Nombre
            lblDescripcionServicio.Tag = oDOServicio.IdEspecialidad
        Else
            txtIdServicio.Text = ""
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
            lblDescripcionServicio.Tag = ""
        End If
    End If

End Sub

Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDOServicio As DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Tag = oDOServicio.IdServicio
            lblDescripcionServicio = oDOServicio.Nombre
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
        End If
   End If

End Sub

Private Sub txtNroOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroOrden
    AdministrarKeyPreview KeyCode
End Sub
Property Let IdDepartamento(lValue As Long)
   ml_IdDepartamento = lValue
End Property
Property Get IdDepartamento() As Long
   IdDepartamento = ml_IdDepartamento
End Property

