VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FarmaciaRecetasDetalle 
   Caption         =   "Registro de Recetas"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "FarmaciaRecetasDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
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
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   11430
      Begin VB.TextBox txtNroCuenta 
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
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   3120
         Picture         =   "FarmaciaRecetasDetalle.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label50 
         Caption         =   "Nro Cuenta"
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
      Left            =   0
      TabIndex        =   40
      Top             =   885
      Width           =   11430
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
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   45
      TabIndex        =   39
      Top             =   7890
      Width           =   11355
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarmaciaRecetasDetalle.frx":3913
         DownPicture     =   "FarmaciaRecetasDetalle.frx":3DD7
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
         Picture         =   "FarmaciaRecetasDetalle.frx":42C3
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarmaciaRecetasDetalle.frx":47AF
         DownPicture     =   "FarmaciaRecetasDetalle.frx":4C0F
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
         Picture         =   "FarmaciaRecetasDetalle.frx":5084
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraAddProcedimiento 
      Height          =   975
      Left            =   15
      TabIndex        =   38
      Top             =   3360
      Width           =   11415
      Begin VB.TextBox txtCantidad 
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
         TabIndex        =   31
         Top             =   540
         Width           =   975
      End
      Begin VB.CommandButton btnBusquedaCatalogoBienesInsumos 
         Caption         =   ".."
         Height          =   315
         Left            =   2700
         TabIndex        =   28
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox lblDescProducto 
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   180
         Width           =   8175
      End
      Begin VB.TextBox txtIdProducto 
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
         TabIndex        =   27
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "FarmaciaRecetasDetalle.frx":54F9
         DownPicture     =   "FarmaciaRecetasDetalle.frx":592B
         Height          =   315
         Left            =   2730
         Picture         =   "FarmaciaRecetasDetalle.frx":5D5D
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   540
         Width           =   1005
      End
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "FarmaciaRecetasDetalle.frx":7FAE
         DownPicture     =   "FarmaciaRecetasDetalle.frx":8339
         Height          =   315
         Left            =   3840
         Picture         =   "FarmaciaRecetasDetalle.frx":86CC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad"
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
         TabIndex        =   30
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label Label5 
         Caption         =   "Medicamento"
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
         TabIndex        =   26
         Top             =   180
         Width           =   1260
      End
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
      Left            =   15
      TabIndex        =   37
      Top             =   1950
      Width           =   11415
      Begin VB.CommandButton btnBusquedaServicio 
         Caption         =   "..."
         Height          =   315
         Left            =   2700
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   960
         Width           =   345
      End
      Begin VB.CommandButton btnBusquedaMedico 
         Caption         =   "..."
         Height          =   315
         Left            =   2700
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   345
      End
      Begin VB.TextBox txtNroReceta 
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
         Top             =   240
         Width           =   975
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
         TabIndex        =   19
         Top             =   600
         Width           =   975
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
         Left            =   3120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   5325
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
         Left            =   3120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   960
         Width           =   5325
      End
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFechaReceta 
         Height          =   315
         Left            =   3950
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
      Begin VB.Label Label69 
         Caption         =   "Orden Receta"
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
         TabIndex        =   18
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label Label65 
         Caption         =   "Fecha Receta"
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
         TabIndex        =   22
         Top             =   990
         Width           =   1425
      End
   End
   Begin MSComctlLib.ImageList lstOpciones 
      Left            =   195
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
            Picture         =   "FarmaciaRecetasDetalle.frx":8A5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FarmaciaRecetasDetalle.frx":8E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FarmaciaRecetasDetalle.frx":934C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FarmaciaRecetasDetalle.frx":9763
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin UltraGrid.SSUltraGrid grdProcedimientos 
      Height          =   3435
      Left            =   0
      TabIndex        =   34
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
      Caption         =   "Lista de medicamentos"
   End
End
Attribute VB_Name = "FarmaciaRecetasDetalle"
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
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_RecetaDetalles As New Collection
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdReceta As Long
'Dim ml_IdTipoServicio As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHComun.ListaDespleglable
'Dim mo_RecetaDetalle As New Collection
Dim mo_Receta As New DOFarmaciaRecetas
Dim mrs_Productos As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mrs_ProductosEliminados As New Recordset

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
Property Let IdReceta(lValue As Long)
   ml_IdReceta = lValue
End Property
Property Get IdReceta() As Long
   IdReceta = ml_IdReceta
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
    If Val(Me.txtIdProducto.Tag) <= 0 Then
        Exit Sub
    End If
    If mrs_Productos.EOF = False Or mrs_Productos.BOF = False Then
        mrs_Productos.MoveFirst
        Do While Not mrs_Productos.EOF
            If mrs_Productos.Fields!IdProducto = Val(Me.txtIdProducto.Tag) Then
                mrs_Productos.MoveFirst
                Exit Sub
            End If
            mrs_Productos.MoveNext
        Loop
        mrs_Productos.MoveFirst
    End If
    
    With mrs_Productos
        .AddNew
        
        .Fields!IdRecetaDetalle = 0
        .Fields!IdReceta = 0
        .Fields!IdFacturacionBienes = 0
        .Fields!IdProducto = Val(Me.txtIdProducto.Tag)
        .Fields!Codigo = Me.txtIdProducto.Text
        .Fields!Descripcion = Me.lblDescProducto
        .Fields!Cantidad = Val(Me.txtCantidad)
        .Fields!EstadoRegistro = "A"
    End With
End Sub

Private Sub btnBuscar_Click()
Dim rsPaciente As New Recordset
Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion
    
    If (Me.txtNroCuenta.Text) = "" Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If
    
    'oDOPaciente.NroHistoriaClinica = Val(Me.cmbNroHistoriaBusqueda.Text)
    oDOCuentaAtencion.IdCuentaAtencion = Val(Me.txtNroCuenta.Text)
    
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
Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection

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
       CompletarDatosDeMedico txtIdMedicoOrdena, lblDescMedicoOrdena
End Sub

Private Sub btnBusquedaServicio_Click()
        CompletarDatosDeServicio txtIdServicio, lblDescServicio
End Sub

Private Sub btnQuitarDx_Click()
    Dim doFacturacionBienesInsumos As doFacturacionBienesInsumos
    On Error Resume Next
    With mrs_Productos
        If Not .EOF And Not .BOF Then
            If mrs_Productos!IdRecetaDetalle <> 0 Then
                'Verificamos que el detalle esté como emitido para poder eliminarse
                Set doFacturacionBienesInsumos = mo_AdminFacturacion.FacturacionBienesInsumosSeleccionarPorId(mrs_Productos!IdFacturacionBienes)
                If Not doFacturacionBienesInsumos Is Nothing Then
                    If doFacturacionBienesInsumos.IdEstadoFacturacion <> 1 Then
                        MsgBox "No se puede eliminar el item seleccionado por que ya se encuentra [" & mo_AdminFacturacion.EstadosFacturacionObtenerDescripcionPorId(doFacturacionBienesInsumos.IdEstadoFacturacion) + "]", vbExclamation, Me.Caption
                        Exit Sub
                    End If
                End If
                mrs_ProductosEliminados.AddNew
                mrs_ProductosEliminados!IdRecetaDetalle = mrs_Productos!IdRecetaDetalle
                mrs_ProductosEliminados!IdFacturacionBienes = mrs_Productos!IdFacturacionBienes
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

Private Sub txtFechaReceta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaReceta
End Sub
Private Sub txtFechaReceta_LostFocus()

    If txtFechaReceta <> SIGHComun.FECHA_VACIA_DMY Then
         If Not EsFecha(txtFechaReceta, "DD/MM/AAAA") Then
             MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
              txtFechaReceta = SIGHComun.FECHA_VACIA_DMY
         End If
     End If
     
     mo_Formulario.MarcarComoVacio txtFechaReceta
End Sub

Private Sub txtFechaReceta_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
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

Private Sub txtNroCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroCuenta
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroCuenta_LostFocus()
   mo_Formulario.MarcarComoVacio txtNroCuenta
End Sub

Private Sub txtNroCuenta_KeyPress(KeyAscii As Integer)
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
           Me.Caption = "Agregar orden de receta"
       Case sghModificar
           Me.Caption = "Modificar orden de receta"
       Case sghConsultar
           Me.Caption = "Consultar orden de receta"
       Case sghEliminar
           Me.Caption = "Eliminar orden de receta"
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
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
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
    
    If txtNroReceta = "" Then
        MsgBox "Ingrese el nro de receta", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtIdMedicoOrdena = "" Then
        MsgBox "Ingrese el médico que ordena la receta", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtFechaReceta = SIGHComun.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de la receta", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
    If txtFechaReceta < CDate(Me.lblFechaIngreso) Then
        MsgBox "La fecha de la receta no puede ser menor que la fecha de ingreso de la atención", vbExclamation, Me.Caption
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
   mo_Receta.IdCuentaAtencion = Val(Me.lblNroCuenta)
   
   CargarProcedimientosAlObjetoDatos mo_Receta, mo_RecetaDetalles
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminFacturacion.FarmaciaRecetasAgregar(mo_Receta, mo_RecetaDetalles)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminFacturacion.FarmaciaRecetasModificar(mo_Receta, mo_RecetaDetalles, mrs_ProductosEliminados)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminFacturacion.FarmaciaRecetasEliminar(mo_Receta, mo_RecetaDetalles)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    
    '1ro
    Dim oDOFarmaciaRecetas As New DOFarmaciaRecetas
    Set oDOFarmaciaRecetas = mo_AdminFacturacion.FarmaciaRecetasSeleccionarPorId(Me.IdReceta)
    If Not oDOFarmaciaRecetas Is Nothing Then
        CargarDatosDelaAtencion oDOFarmaciaRecetas.IdCuentaAtencion
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
    
    With mrs_Productos
        .Fields.Append "IdRecetaDetalle", adInteger
        .Fields.Append "IdReceta", adInteger
        .Fields.Append "IdFacturacionBienes", adInteger
        .Fields.Append "IdProducto", adInteger
        .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
        .Fields.Append "Descripcion", adVarChar, 200, adFldIsNullable
        .Fields.Append "Cantidad", adInteger
        .Fields.Append "EstadoRegistro", adVarChar, 1, adFldIsNullable
        
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    'Para los procedimientos eliminados
    With mrs_ProductosEliminados
        .Fields.Append "IdRecetaDetalle", adInteger
        .Fields.Append "IdFacturacionBienes", adInteger, , adFldIsNullable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set grdProcedimientos.DataSource = mrs_Productos
    
End Sub

Public Sub CargarDatosDeDeProcedimientos()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDOFarmaciaRecetas As New DOFarmaciaRecetas

    'Carga datos de la cabecera
    Dim rsProcedimiento As New Recordset
    Set oDOFarmaciaRecetas = mo_AdminFacturacion.FarmaciaRecetasSeleccionarPorId(Me.IdReceta)
    
    If oDOFarmaciaRecetas.IdReceta = 0 Then
        MsgBox "No existe datos de recetas", vbInformation, Me.Caption
        Exit Sub
    End If
    
    txtFechaReceta = oDOFarmaciaRecetas.FechaReceta
    txtNroReceta = oDOFarmaciaRecetas.NroReceta

    'Completa datos de medico
    If mo_AdminProgramacion.MedicosSeleccionarPorId(oDOFarmaciaRecetas.IdMedicoOrdena, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
        txtIdMedicoOrdena.Text = oDOEmpleado.CodigoPlanilla
        txtIdMedicoOrdena.Tag = oDoMedico.IdMedico
        lblDescMedicoOrdena = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    
    Me.txtIdServicio.Tag = IIf(oDOFarmaciaRecetas.IdServicioOrdena = 0, "", oDOFarmaciaRecetas.IdServicioOrdena)
    Dim oDOServicio As New DOServicio
    If Me.txtIdServicio.Tag <> "" Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oDOFarmaciaRecetas.IdServicioOrdena)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicio.Text = oDOServicio.Codigo
            Me.lblDescServicio = oDOServicio.Nombre
        End If
    End If
  
    Dim rsProcedimientos As New Recordset
    Set rsProcedimientos = mo_AdminFacturacion.FarmaciaRecetasDetalleSeleccionarPorIdReceta(Me.IdReceta)
    Do While Not rsProcedimientos.EOF
        With mrs_Productos
            .AddNew
            
            .Fields!IdRecetaDetalle = rsProcedimientos!IdRecetaDetalle
            .Fields!IdReceta = rsProcedimientos!IdReceta
            .Fields!IdFacturacionBienes = rsProcedimientos!IdFacturacionBienes
            .Fields!IdProducto = rsProcedimientos!IdProducto
            .Fields!Descripcion = rsProcedimientos!NombreProducto
            .Fields!Cantidad = rsProcedimientos!Cantidad
            .Fields!EstadoRegistro = "M"
        End With
        rsProcedimientos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores grdProcedimientos, SIGHComun.GrillaConFilasBicolor
    
End Sub

Sub CargarProcedimientosAlObjetoDatos(oReceta As DOFarmaciaRecetas, oRecetaDetalle As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ProcedimientoS
    '---------------------------------------------------------------------------------
    'Datos de la cabecera
    oReceta.IdReceta = Me.IdReceta
    oReceta.IdCuentaAtencion = Val(lblNroCuenta)
    oReceta.IdMedicoOrdena = Val(txtIdMedicoOrdena.Tag)
    oReceta.IdServicioOrdena = Val(Me.txtIdServicio.Tag)
    oReceta.FechaReceta = txtFechaReceta.Text
    oReceta.NroReceta = txtNroReceta.Text
    oReceta.IdUsuarioAuditoria = ml_IdUsuario
    
    Set oRecetaDetalle = New Collection
    'Datos del detalle
    Dim oFarmaciaRecetasDetalle As DOFarmaciaRecetasDetalle
    If Not (mrs_Productos.BOF And mrs_Productos.EOF) Then
        Set oFarmaciaRecetasDetalle = New DOFarmaciaRecetasDetalle
        mrs_Productos.MoveFirst
        Do While Not mrs_Productos.EOF
            Set oFarmaciaRecetasDetalle = New DOFarmaciaRecetasDetalle
            
            oFarmaciaRecetasDetalle.IdRecetaDetalle = mrs_Productos!IdRecetaDetalle
            oFarmaciaRecetasDetalle.IdReceta = Me.IdReceta
            oFarmaciaRecetasDetalle.IdProducto = mrs_Productos!IdProducto
            oFarmaciaRecetasDetalle.Cantidad = mrs_Productos!Cantidad
            oFarmaciaRecetasDetalle.IdUsuarioAuditoria = ml_IdUsuario
            oFarmaciaRecetasDetalle.IdFacturacionBienes = IIf(IsNull(mrs_Productos!IdFacturacionBienes), 0, mrs_Productos!IdFacturacionBienes)
            oFarmaciaRecetasDetalle.EstadoRegistro = mrs_Productos!EstadoRegistro
            
            oRecetaDetalle.Add oFarmaciaRecetasDetalle
            mrs_Productos.MoveNext
        Loop
    End If
    
End Sub

Private Sub grdProcedimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdProcedimientos.Bands(0).Columns("IdRecetaDetalle").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdReceta").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdFacturacionBienes").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdProducto").Hidden = True
    grdProcedimientos.Bands(0).Columns("EstadoRegistro").Hidden = True
    
    
    grdProcedimientos.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdProcedimientos.Bands(0).Columns("Codigo").Width = 1000
    
    grdProcedimientos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdProcedimientos.Bands(0).Columns("Descripcion").Width = 5000
    
    grdProcedimientos.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    grdProcedimientos.Bands(0).Columns("Cantidad").Width = 1000


End Sub
Private Sub btnBusquedaCatalogoBienesInsumos_Click()
Dim oBusqueda As New BienesInsumosBusqueda
Dim oDOCatalogoBienesInsumos As DOCatalogoBienesInsumos

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOCatalogoBienesInsumos = mo_AdminServiciosComunes.CatalogoBienesInsumosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOCatalogoBienesInsumos Is Nothing Then
            txtIdProducto.Text = oDOCatalogoBienesInsumos.Codigo
            txtIdProducto.Tag = oDOCatalogoBienesInsumos.IdProducto
            lblDescProducto = oDOCatalogoBienesInsumos.Nombre
        Else
            txtIdProducto.Text = ""
            txtIdProducto.Tag = ""
            lblDescProducto = ""
        End If
    End If
    
End Sub

Private Sub txtIdProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdProducto
End Sub

Private Sub txtIdProducto_LostFocus()

    txtIdProducto.Text = UCase(txtIdProducto.Text)

    If txtIdProducto.Text <> "" Then
        Dim oDOProcedimiento As DOProcedimiento
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorCodigoCPT(txtIdProducto.Text)
        If Not oDOProcedimiento Is Nothing Then
            txtIdProducto.Tag = oDOProcedimiento.IdProcedimiento
            lblDescProducto = oDOProcedimiento.Descripcion
        Else
            txtIdProducto.Tag = ""
            lblDescProducto = ""
        End If
    Else
        txtIdProducto.Tag = ""
        lblDescProducto = ""
    End If
   'mo_Formulario.MarcarComoVacio txtIdProducto
End Sub

Private Sub txtIdProducto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtIdProducto_LostFocus
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



