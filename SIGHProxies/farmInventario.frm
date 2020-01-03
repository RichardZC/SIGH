VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form FarmInventario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "farmInventario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7725
      Left            =   15
      TabIndex        =   4
      Top             =   45
      Width           =   15105
      Begin SIGHProxies.ucFarmaciaItemsInventario grdProductos 
         Height          =   3435
         Left            =   105
         TabIndex        =   48
         Top             =   2160
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   6059
      End
      Begin VB.CommandButton CargaInventarioExcel 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7545
         Picture         =   "farmInventario.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Carga C:\tformdetl.XLS (e=codigo,f=lote,g=f_vencimiento,h=saldo,i=registro sanitario <<empieza en FILA=2>> <<Hoja=tformdetl>>"
         Top             =   960
         Width           =   435
      End
      Begin VB.Frame fraDatosHistoria 
         Caption         =   "Datos de Cabecera"
         Height          =   1875
         Left            =   90
         TabIndex        =   20
         Top             =   210
         Width           =   14985
         Begin VB.CommandButton btnCArgaDesdeSismedv2 
            Caption         =   "DBF"
            Enabled         =   0   'False
            Height          =   315
            Left            =   6945
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "ODBC HIS, tener TFORMDETL.DBF, TMOVIMDET.DBF en c:\...\galenhos\archivos"
            Top             =   750
            Width           =   435
         End
         Begin VB.Frame Frame 
            Height          =   1650
            Left            =   8070
            TabIndex        =   44
            Top             =   180
            Width           =   4245
            Begin VB.CommandButton cmdActualizaInventarioTemp 
               Caption         =   "Agrega Inventarios Temporales"
               DisabledPicture =   "farmInventario.frx":110C
               DownPicture     =   "farmInventario.frx":154E
               Height          =   1110
               Left            =   2850
               Picture         =   "farmInventario.frx":1990
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   495
               Width           =   1320
            End
            Begin VB.CheckBox chkInventarioTemp 
               Caption         =   "El registro del inventario es temporal"
               Height          =   330
               Left            =   135
               TabIndex        =   45
               Top             =   135
               Width           =   3960
            End
            Begin UltraGrid.SSUltraGrid grdIventariosTemp 
               Height          =   1110
               Left            =   150
               TabIndex        =   46
               Top             =   510
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   1958
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Inventarios Temporales"
            End
            Begin VB.Line Line 
               BorderColor     =   &H80000004&
               BorderWidth     =   2
               X1              =   45
               X2              =   4485
               Y1              =   465
               Y2              =   465
            End
         End
         Begin VB.ComboBox cmbTipoInventario 
            Height          =   330
            ItemData        =   "farmInventario.frx":1DD2
            Left            =   5955
            List            =   "farmInventario.frx":1DD4
            TabIndex        =   37
            Top             =   1215
            Width           =   1935
         End
         Begin VB.TextBox txtNinventario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   27
            Top             =   330
            Width           =   1035
         End
         Begin VB.TextBox txtEstado 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            MaxLength       =   30
            TabIndex        =   26
            Top             =   330
            Width           =   1635
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1650
            Left            =   12375
            TabIndex        =   22
            Top             =   180
            Width           =   2520
            Begin VB.CommandButton btnCierre 
               Caption         =   "Cierre"
               DisabledPicture =   "farmInventario.frx":1DD6
               DownPicture     =   "farmInventario.frx":2236
               Height          =   945
               Left            =   195
               Picture         =   "farmInventario.frx":26AB
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   615
               Width           =   2205
            End
            Begin MSMask.MaskEdBox txtFcierre 
               Height          =   315
               Left            =   1020
               TabIndex        =   24
               Top             =   210
               Width           =   1350
               _ExtentX        =   2381
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
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "F. Cierre"
               Height          =   210
               Left            =   150
               TabIndex        =   25
               Top             =   240
               Width           =   675
            End
         End
         Begin VB.ComboBox cmbAlmacen 
            Height          =   330
            Left            =   1305
            TabIndex        =   21
            Top             =   735
            Width           =   5610
         End
         Begin MSMask.MaskEdBox txtFmodificacion 
            Height          =   315
            Left            =   1290
            TabIndex        =   28
            Top             =   1200
            Width           =   1350
            _ExtentX        =   2381
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
         Begin MSMask.MaskEdBox txtFregistro 
            Height          =   315
            Left            =   6510
            TabIndex        =   29
            Top             =   330
            Width           =   1350
            _ExtentX        =   2381
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
         Begin VB.Label lblMensaje 
            AutoSize        =   -1  'True
            Caption         =   "Espere...el Sistema está REGENERANDO SALDOS y luego generando SALDOS AUTOMATICOS"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   90
            TabIndex        =   39
            Top             =   1545
            Visible         =   0   'False
            Width           =   7695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo Inventario"
            Height          =   210
            Left            =   4695
            TabIndex        =   38
            Top             =   1245
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   210
            Left            =   120
            TabIndex        =   34
            Top             =   810
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "N° Inventario"
            Height          =   210
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   210
            Left            =   2700
            TabIndex        =   32
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "F.Modificación"
            Height          =   210
            Left            =   120
            TabIndex        =   31
            Top             =   1230
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "F.Registro"
            Height          =   210
            Left            =   5400
            TabIndex        =   30
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame fraDetalleLote 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   9900
         TabIndex        =   5
         Top             =   5550
         Width           =   5145
         Begin VB.CommandButton cmdAdicionarItem 
            Caption         =   "Agregar"
            DisabledPicture =   "farmInventario.frx":2B20
            DownPicture     =   "farmInventario.frx":2F80
            Height          =   930
            Left            =   4155
            Picture         =   "farmInventario.frx":33F5
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1005
            Width           =   930
         End
         Begin VB.ComboBox cmbTipoSalida 
            Height          =   330
            ItemData        =   "farmInventario.frx":3837
            Left            =   1290
            List            =   "farmInventario.frx":3839
            TabIndex        =   12
            Text            =   "cmbTipoSalida"
            Top             =   1470
            Width           =   1620
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4080
            MaxLength       =   30
            TabIndex        =   9
            Top             =   180
            Width           =   1005
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4080
            MaxLength       =   30
            TabIndex        =   13
            Top             =   630
            Width           =   1005
         End
         Begin VB.TextBox txtRegSanitario 
            Height          =   315
            Left            =   1290
            MaxLength       =   50
            TabIndex        =   6
            Top             =   210
            Width           =   1605
         End
         Begin VB.TextBox txtLote 
            Height          =   315
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   7
            Top             =   630
            Width           =   1605
         End
         Begin MSMask.MaskEdBox txtFvencimiento 
            Height          =   315
            Left            =   1290
            TabIndex        =   8
            Top             =   1050
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Threed.SSCommand btnModificar 
            Height          =   465
            Left            =   2925
            TabIndex        =   11
            Top             =   1020
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   820
            _Version        =   262144
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "farmInventario.frx":383B
            Caption         =   "Modificar"
            PictureAlignment=   9
         End
         Begin Threed.SSCommand btnQuitar 
            Height          =   465
            Left            =   2925
            TabIndex        =   14
            Top             =   1470
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   820
            _Version        =   262144
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "farmInventario.frx":67C7
            Caption         =   "Quitar"
            PictureAlignment=   9
            ShapeSize       =   1
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Salida"
            Height          =   210
            Left            =   60
            TabIndex        =   36
            Top             =   1500
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   210
            Left            =   3330
            TabIndex        =   19
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta"
            Height          =   210
            Left            =   3030
            TabIndex        =   18
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Reg.Sanitario"
            Height          =   210
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   210
            Left            =   90
            TabIndex        =   16
            Top             =   660
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "F.Vencimiento"
            Height          =   210
            Left            =   90
            TabIndex        =   15
            Top             =   1080
            Width           =   1170
         End
      End
      Begin UltraGrid.SSUltraGrid grdProductosDetalle 
         Height          =   1905
         Left            =   60
         TabIndex        =   35
         Top             =   5670
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   3360
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle Producto"
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   30
      TabIndex        =   0
      Top             =   7860
      Width           =   15120
      Begin VB.CheckBox chkEnExcel 
         Caption         =   "Reportes en EXCEL"
         Height          =   225
         Left            =   12705
         TabIndex        =   50
         Top             =   405
         Width           =   2190
      End
      Begin VB.CommandButton btnImprimirInvConteo 
         Caption         =   "Conteo"
         Height          =   700
         Left            =   1560
         Picture         =   "farmInventario.frx":8C49
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimirInvDet 
         Caption         =   "Detallado"
         Height          =   700
         Left            =   2880
         Picture         =   "farmInventario.frx":9122
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimirInvGeneral 
         Caption         =   "General"
         Height          =   700
         Left            =   4320
         Picture         =   "farmInventario.frx":95FB
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         Height          =   700
         Left            =   150
         Picture         =   "farmInventario.frx":9AD4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "farmInventario.frx":9FAD
         DownPicture     =   "farmInventario.frx":A40D
         Height          =   700
         Left            =   6173
         Picture         =   "farmInventario.frx":A882
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "farmInventario.frx":ACF7
         DownPicture     =   "farmInventario.frx":B1BB
         Height          =   700
         Left            =   7703
         Picture         =   "farmInventario.frx":B6A7
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FarmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Inventario
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mo_cmbTipoSalida As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoInventario As New SIGHEntidades.ListaDespleglable

Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mrs_ProductosCabecera As New ADODB.Recordset
Dim mrs_ProductosDetalle As New ADODB.Recordset
Dim oRsSaldosAjuste As New ADODB.Recordset
Dim oRsIventarioTmp As New Recordset
Dim oRsItemsUnidosis As New Recordset
Dim lnCodigoProducto As Long
Dim ml_IdInventario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_FarmInventario As New sighComun.DoFarmInventario
Dim ms_MensajeError As String
Dim LdFechaMinimaVencimiento As Date
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcSql As String
Dim mo_mensajeError As String
Dim lcIdTipoSuministro As String
Const lcConstanteMovimientoEntrada As String = "E"
Const LcIdTipoDocumentoNINGUNO As Long = 22
Const lcConstanteMovimientoSalida As String = "S"
Dim lbLaFarmaciaEsUnidosis As Boolean

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let IdInventario(lValue As Long)
   ml_IdInventario = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Sub CargaDatosAlObjetosDeDatos()
    
    Select Case mi_Opcion
    Case sghAgregar
        With mo_FarmInventario
            .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL    'txtFregistro.Text
            .IdAlmacen = mo_cmbAlmacen.BoundText
            .idEstadoInventario = sghEstadoTabla.sghRegistrado    'registrado
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .idTipoInventario = mo_cmbTipoInventario.BoundText
        End With
   Case sghModificar
        With mo_FarmInventario
            .fechaModificacion = lcBuscaParametro.RetornaFechaServidorSQL
            .IdUsuarioAuditoria = ml_idUsuario
        End With
   Case sghEliminar
        With mo_FarmInventario
            .fechaModificacion = lcBuscaParametro.RetornaFechaServidorSQL
            .IdUsuarioAuditoria = ml_idUsuario
            .idEstadoInventario = sghEstadoTabla.sghAnulado      'Anulado
        End With
   End Select
End Sub
Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

    
    
    If btnAceptar.Enabled = False Then
      Exit Sub
    End If
    Select Case mi_Opcion
    Case sghAgregar, sghModificar
           cmbAceptar
    Case sghEliminar
           If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
               CargaDatosAlObjetosDeDatos
               If AnularInventario() Then
                   MsgBox " Se anuló exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Sub cmbAceptar()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If AgregarDatos() Then
                MsgBox " Los datos se agregaron exitosamente INVENTARIO: " & mo_FarmInventario.NumeroInventario, vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo agregar los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ModificarDatos() Then
                MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghEliminar
        If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
            CargaDatosAlObjetosDeDatos
            If AnularInventario() Then
                MsgBox " Se anuló exitosamente", vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   End Select
End Sub

Function AgregarDatos() As Boolean
    Dim lcInventarioTemporal As String
    Dim oReglasCaja As New SIGHNegocios.ReglasCaja
    Dim oConexion As New Connection
    If Me.chkInventarioTemp.Value = 1 Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        lcInventarioTemporal = Left(oReglasCaja.SeleccionaDatosCajeroConexion(Val(SIGHEntidades.Usuario), sghUsuario, oConexion), 4)
    Else
        lcInventarioTemporal = ""
    End If
    Set oConexion = Nothing
    Set oReglasCaja = Nothing
    '
    AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeInventario(mo_FarmInventario, mrs_ProductosCabecera, mrs_ProductosDetalle, _
                                                                 mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcInventarioTemporal)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasFarmacia.ModificaDatosDeInventario(mo_FarmInventario, mrs_ProductosCabecera, mrs_ProductosDetalle, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function AnularInventario() As Boolean
    AnularInventario = mo_ReglasFarmacia.AnulaInventario(mo_FarmInventario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function CierraInventarioCreaNIporAjusteInventario() As Boolean
        Dim mo_farmMovimiento As New DoFarmMovimiento
        Dim mo_farmMovimientoNotaIngreso As New DOfarmMovimientoNotaIngreso
        Dim oDoProveedores As New DoProveedores
        Dim oConexion As New Connection
        Dim oRsTmp1 As New Recordset
        Dim lnTotalDocumento As Double
        Dim lnCantidad As Long, lnPrecio As Double
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        mrs_ProductosDetalle.Filter = "cantidadSobrante>0"
        lnTotalDocumento = 0
        If mrs_ProductosDetalle.RecordCount > 0 Then
            mrs_ProductosDetalle.MoveFirst
            Do While Not mrs_ProductosDetalle.EOF
               Set oRsTmp1 = mo_reglasComunes.CatalogoBienesInsumosSeleccionarXid(mrs_ProductosDetalle.Fields!idProducto, oConexion)
               lnPrecio = oRsTmp1!PrecioCompra
               oRsTmp1.Close
               lnCantidad = mrs_ProductosDetalle.Fields!cantidadSobrante
               lnTotalDocumento = lnTotalDocumento + Round(lnCantidad * lnPrecio, 2)
               mrs_ProductosDetalle.Fields!Cantidad = lnCantidad
               mrs_ProductosDetalle.Fields!precio = lnPrecio
               mrs_ProductosDetalle.Fields!Total = Round(lnCantidad * lnPrecio, 2)
               mrs_ProductosDetalle.Update
               mrs_ProductosDetalle.MoveNext
            Loop
            With mo_farmMovimiento
                .DocumentoIdtipo = 10
                .DocumentoNumero = mo_FarmInventario.NumeroInventario
                .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                .IdAlmacenDestino = Val(mo_cmbAlmacen.BoundText)
                .IdAlmacenOrigen = 0
                .idEstadoMovimiento = sghEstadoTabla.sghRegistrado
                .idTipoConcepto = 20
                .idUsuario = ml_idUsuario
                .IdUsuarioAuditoria = ml_idUsuario
                .MovTipo = lcConstanteMovimientoEntrada
                .Observaciones = ""
                .Total = lnTotalDocumento
            End With
            With mo_farmMovimientoNotaIngreso
                .DocumentoFechaRecepcion = Format(mo_farmMovimiento.fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                .idPaciente = 0
                .IdComprobantePago = 0
                .IdProveedor = 0
                .idTipoCompra = 1
                .idTipoProceso = 1
                .IdUsuarioAuditoria = ml_idUsuario
                .MovTipo = lcConstanteMovimientoEntrada
                .NumeroProceso = ""
                .OrigenFecha = 0
                .OrigenIdTipo = 22
                .oRigenNumero = ""
                .idCuentaAtencion = 0
                .idFuenteFinanciamiento = 0
            End With
            With oDoProveedores
            End With
            oConexion.Close
            CierraInventarioCreaNIporAjusteInventario = mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento, _
                                             mo_farmMovimientoNotaIngreso, oDoProveedores, mrs_ProductosDetalle, 0, _
                                             mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
        Else
            CierraInventarioCreaNIporAjusteInventario = True
        End If
        Set mo_farmMovimiento = Nothing
        Set mo_farmMovimientoNotaIngreso = Nothing
        Set oDoProveedores = Nothing
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
End Function

Function CierraInventarioCreaNSporAjusteInventario() As Boolean
        Dim mo_farmMovimiento As New DoFarmMovimiento
        Dim oConexion As New Connection
        Dim oRsTmp1 As New Recordset
        Dim lnTotalDocumento As Double
        Dim lnCantidad As Long, lnPrecio As Double
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        mrs_ProductosDetalle.Filter = "cantidadFaltante>0"
        lnTotalDocumento = 0
        If mrs_ProductosDetalle.RecordCount > 0 Then
            mrs_ProductosDetalle.MoveFirst
            Do While Not mrs_ProductosDetalle.EOF
               Set oRsTmp1 = mo_reglasComunes.CatalogoBienesInsumosSeleccionarXid(mrs_ProductosDetalle.Fields!idProducto, oConexion)
               lnPrecio = oRsTmp1!PrecioCompra
               oRsTmp1.Close
               lnCantidad = mrs_ProductosDetalle.Fields!cantidadFaltante
               lnTotalDocumento = lnTotalDocumento + Round(lnCantidad * lnPrecio, 2)
               mrs_ProductosDetalle.Fields!Cantidad = lnCantidad
               mrs_ProductosDetalle.Fields!precio = lnPrecio
               mrs_ProductosDetalle.Fields!Total = Round(lnCantidad * lnPrecio, 2)
               mrs_ProductosDetalle.Update
               mrs_ProductosDetalle.MoveNext
            Loop
            With mo_farmMovimiento
                .DocumentoIdtipo = 10
                .DocumentoNumero = mo_FarmInventario.NumeroInventario
                .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                .IdAlmacenDestino = 0
                .IdAlmacenOrigen = Val(mo_cmbAlmacen.BoundText)
                .idEstadoMovimiento = sghEstadoTabla.sghRegistrado
                .idTipoConcepto = 20
                .idUsuario = ml_idUsuario
                .IdUsuarioAuditoria = ml_idUsuario
                .MovTipo = lcConstanteMovimientoSalida
                .Observaciones = ""
                .Total = lnTotalDocumento
            End With
            CierraInventarioCreaNSporAjusteInventario = mo_ReglasFarmacia.AgregaDatosDeNotaSalida(mo_farmMovimiento, _
                                                           mrs_ProductosDetalle, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
        Else
            CierraInventarioCreaNSporAjusteInventario = True
        End If
        Set mo_farmMovimiento = Nothing
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
End Function

Function CierraInventarioActualizaCabeceraInventario() As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInventario As New SIGHDatos.FarmInventario
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oInventario.Conexion = oConexion
    CierraInventarioActualizaCabeceraInventario = oInventario.modificar(mo_FarmInventario)
    oConexion.Close
    Set oConexion = Nothing
    Set oInventario = Nothing
End Function
Function CierraInventario() As Boolean
    If mo_cmbTipoInventario.BoundText = SIGHEntidades.sghInventarioTipo.sghManual Then
        CierraInventario = mo_ReglasFarmacia.CierraInventario(mo_FarmInventario, grdProductos.DevuelveTotal, _
                                                             mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    Else
        If CierraInventarioCreaNIporAjusteInventario = True Then
           If CierraInventarioCreaNSporAjusteInventario = True Then
              If CierraInventarioActualizaCabeceraInventario = True Then
                 CierraInventario = True
              Else
                 MsgBox "!!!....Hubo problemas al CERRAR INVENTARIO.........................................!!" & Chr(13) & _
                        " ELIMINE la NOTA DE SALIDA y NOTA DE INGRESO POR AJUSTE DE INVENTARIO CREADAS AHORITA" & Chr(13) & _
                        "                            y vuelva a CERRAR EL INVENTARIO                          ", vbInformation, "CIERRE DEL INVENTARIO"
              End If
            Else
               MsgBox "!!!....Hubo problemas al generar NOTA DE SALIDA POR AJUSTE DE INVENTARIO...........!!" & Chr(13) & _
                      " ELIMINE la NOTA DE SALIDA y NOTA DE INGRESO POR AJUSTE DE INVENTARIO CREADAS AHORITA" & Chr(13) & _
                      "                           y vuelva a CERRAR EL INVENTARIO                           ", vbInformation, "CIERRE DEL INVENTARIO"
           End If
        Else
           MsgBox "!!!....Hubo problemas al generar NOTA DE INGRESO POR AJUSTE DE INVENTARIO!!" & Chr(13) & _
                  "     ELIMINE la NOTA DE INGRESO POR AJUSTE DE INVENTARIO CREADA AHORITA    " & Chr(13) & _
                  "                      y vuelva a CERRAR EL INVENTARIO                      ", vbInformation, "CIERRE DEL INVENTARIO"
        End If
    End If
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function


Function ValidarDatosObligatorios() As Boolean
   
   ValidarDatosObligatorios = False
   ms_MensajeError = ""
   If cmbAlmacen.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén" + Chr(13)
       cmbAlmacen.SetFocus
   End If
   Set mrs_ProductosCabecera = grdProductos.DevuelveProductos
   mrs_ProductosDetalle.Filter = ""
   If mrs_ProductosCabecera.RecordCount = 0 Then
       ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
   Else
        mrs_ProductosCabecera.MoveFirst
        Do While Not mrs_ProductosCabecera.EOF
           If Trim(mrs_ProductosCabecera.Fields!Codigo) = "" Or Trim(mrs_ProductosCabecera.Fields!nombreProducto) = "" Then
              mrs_ProductosCabecera.Delete
              mrs_ProductosCabecera.Update
           ElseIf sghInventarioTipo.sghManual = Val(mo_cmbTipoInventario.BoundText) And mrs_ProductosCabecera.Fields!Cantidad = 0 Then
              'ms_MensajeError = ms_MensajeError + "El producto " + Trim(mrs_ProductosCabecera.Fields!Codigo) + " - " + Trim(mrs_ProductosCabecera.Fields!nombreProducto) + "  Tiene problemas de CANTIDAD" + Chr(13)
           ElseIf mrs_ProductosCabecera!precio = 0 Then
              ms_MensajeError = ms_MensajeError + "El producto " + Trim(mrs_ProductosCabecera.Fields!Codigo) + " - " + Trim(mrs_ProductosCabecera.Fields!nombreProducto) + "  Tiene problemas de PRECIO" + Chr(13)
           End If
           If mrs_ProductosCabecera!esPaquete = True And (mi_Opcion = sghAgregar Or mi_Opcion = sghModificar) Then
              mrs_ProductosDetalle.Filter = "idProducto = " & mrs_ProductosCabecera!idProducto
              If mrs_ProductosDetalle.RecordCount > 1 Then
                 ms_MensajeError = ms_MensajeError + "El producto " + Trim(mrs_ProductosCabecera.Fields!Codigo) + " - " + Trim(mrs_ProductosCabecera.Fields!nombreProducto) + "  es PAQUETE no puede tener más de un LOTE/F.VENCIMIENTO" + Chr(13)
              Else
                 mrs_ProductosDetalle.MoveFirst
                 If Trim(mrs_ProductosDetalle!lote) <> WxLOTEpaquete Then
                    ms_MensajeError = ms_MensajeError + "El producto " + Trim(mrs_ProductosCabecera.Fields!Codigo) + " - " + Trim(mrs_ProductosCabecera.Fields!nombreProducto) + "  (es un PAQUETE el LOTE debe llamarse " & WxLOTEpaquete & ")" + Chr(13)
                 End If
                 If mrs_ProductosDetalle!fechaVencimiento <> CDate(WxFVENCIMIENTOpaquete) Then
                    ms_MensajeError = ms_MensajeError + "El producto " + Trim(mrs_ProductosCabecera.Fields!Codigo) + " - " + Trim(mrs_ProductosCabecera.Fields!nombreProducto) + "  (es un PAQUETE la FECHA DE VENCIMIENTO debe ser " & WxFVENCIMIENTOpaquete & ")" & Chr(13)
                 End If
                 If Trim(mrs_ProductosDetalle!registroSanitario) <> WxREGSANITARIOpaquete Then
                    ms_MensajeError = ms_MensajeError + "El producto " + Trim(mrs_ProductosCabecera.Fields!Codigo) + " - " + Trim(mrs_ProductosCabecera.Fields!nombreProducto) + "  (es un PAQUETE el REGISTRO SANITARIO debe llamarse " & WxREGSANITARIOpaquete & ")" + Chr(13)
                 End If

              End If
  
           End If

           mrs_ProductosCabecera.MoveNext
        Loop
        'Chequea DETALLE a nivel de  Lotes, ELIMINA los que no tienen CABECERA
        mrs_ProductosDetalle.Filter = ""
        If mrs_ProductosDetalle.RecordCount = 0 Then
            ms_MensajeError = ms_MensajeError + "Por favor Ingrese los LOTES, F.VENCIMIENTO de Productos" + Chr(13)
        Else
             mrs_ProductosDetalle.MoveFirst
             Do While Not mrs_ProductosDetalle.EOF
                mrs_ProductosCabecera.MoveFirst
                mrs_ProductosCabecera.Find "idProducto= " & mrs_ProductosDetalle.Fields!idProducto
                If mrs_ProductosCabecera.EOF Then
                   mrs_ProductosDetalle.Delete
                   mrs_ProductosDetalle.Update
                Else
                    If mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = 3 Then
                          mrs_ProductosCabecera.MoveFirst
                          mrs_ProductosCabecera.Find "idProducto=" & mrs_ProductosDetalle!idProducto
                          ms_MensajeError = ms_MensajeError + "El lote: " & mrs_ProductosDetalle!lote & _
                                           " del producto " + Trim(mrs_ProductosCabecera.Fields!Codigo) & _
                                           " - " & Trim(mrs_ProductosCabecera.Fields!nombreProducto) & _
                                           " sólo puede ser de TIPO SALIDA: Ventas,IntervencionesSanitarias o Donaciones" & Chr(13)
                    End If
                End If
                mrs_ProductosDetalle.MoveNext
             Loop
        End If
        Set grdProductosDetalle.DataSource = Nothing
   End If
   'Es un despacho hacia la FARMACIA UNIDOSIS
   ms_MensajeError = ms_MensajeError & mo_ReglasFarmacia.DevuelveSiSonItemsDeUNIDOSIS(lbLaFarmaciaEsUnidosis, _
                                                         mrs_ProductosCabecera, oRsItemsUnidosis)
   
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function

Private Sub ImprimeDocumento()
    Dim oRptClase As New rCrystal
    oRptClase.MovTipo = ""
    oRptClase.Documento = txtNinventario.Text
    oRptClase.TextoDelFiltro = ""
    oRptClase.Almacen = cmbAlmacen.Text & " (" & Label13.Caption & ": " & Trim(cmbTipoInventario.Text) & ")"
    oRptClase.AlmacenO = ""
    oRptClase.IdAlmacenDestino = mo_cmbAlmacen.BoundText
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = ""
    oRptClase.Importe = 0
    oRptClase.EnArchivoExcel = Me.chkEnExcel.Value
    oRptClase.TipoReporte = "Inventario"
    oRptClase.Show vbModal
    Set oRptClase = Nothing

End Sub



Private Sub btnCancelar_Click()
     Me.Visible = False
     LimpiarVariablesDeMemoria
End Sub

Private Sub btnCArgaDesdeSismedv2_Click()
    If mi_Opcion = sghAgregar And mo_cmbTipoInventario.BoundText = sghInventarioTipo.sghManual And _
                                                        mrs_ProductosDetalle.RecordCount = 0 Then
        Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, lcMensaje As String
        Dim lcCodigo As String, lcFvencimiento As Date, lnSaldo As Long, lcRegSanitario As String, lnPrecioUnitario As Double
        Dim oRsTmp As New Recordset, rs As New Recordset
        Dim oConexion As New Connection
        Dim lnIdProducto As Long, lcLote As String, lnIdTipoSalidaBienInsumo As Long, lcNombreProducto As String
        Dim lbEsNuevo As Boolean
        Dim oConexionFox99 As New Connection
        Dim oRsFox99 As New Recordset
        Dim lcSql As String
        Me.MousePointer = 11
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        
        oConexionFox99.CommandTimeout = 300
        oConexionFox99.Open "DSN=his"
        oConexionFox99.CursorLocation = adUseClient
        oRsFox99.Open "select * from tformdetl where saldo>0   order by     annoMes desc", oConexionFox99, adOpenKeyset, adLockOptimistic
        If oRsFox99.RecordCount > 0 Then
            oRsFox99.MoveFirst
            lcSql = oRsFox99!annoMes
            oRsFox99.Filter = "annoMes=" & lcSql
            Do While Not oRsFox99.EOF
                    lcCodigo = oRsFox99!codigo_med
                    lcLote = oRsFox99!lote
                    lcFvencimiento = oRsFox99!fechVto
                    lnSaldo = oRsFox99!saldo
                    
                    Set oRsTmp = mo_ReglasFacturacion.FactCatalogoBienesInsumosSeleccionarXcodigo(lcCodigo, oConexion)
                    If oRsTmp.RecordCount > 0 Then
                        lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(oRsTmp!idProducto, 3)
                        If lnPrecioUnitario > 0 Then
                            lbEsNuevo = True
                            If mrs_ProductosDetalle.RecordCount > 0 Then
                               mrs_ProductosDetalle.MoveFirst
                               Do While Not mrs_ProductosDetalle.EOF
                                  If mrs_ProductosDetalle!idProducto = oRsTmp.Fields!idProducto And _
                                     Trim(mrs_ProductosDetalle!lote) = lcLote And _
                                     mrs_ProductosDetalle!fechaVencimiento = CDate(lcFvencimiento) Then
                                     lbEsNuevo = False
                                     Exit Do
                                  End If
                                  mrs_ProductosDetalle.MoveNext
                               Loop
                            End If
                            If lbEsNuevo = True Then
                                mrs_ProductosDetalle.AddNew
                                mrs_ProductosDetalle.Fields!idProducto = oRsTmp.Fields!idProducto
                                mrs_ProductosDetalle.Fields!lote = lcLote
                                mrs_ProductosDetalle.Fields!fechaVencimiento = lcFvencimiento
                                mrs_ProductosDetalle.Fields!Cantidad = lnSaldo
                                mrs_ProductosDetalle.Fields!precio = lnPrecioUnitario
                                mrs_ProductosDetalle.Fields!Total = Round(lnSaldo * lnPrecioUnitario, 2)
                                mrs_ProductosDetalle.Fields!registroSanitario = lcRegSanitario
                                mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = IIf(oRsTmp.Fields!idTipoSalidaBienInsumo = 3, 2, oRsTmp.Fields!idTipoSalidaBienInsumo)
                                mrs_ProductosDetalle.Fields!cantidadSaldo = lnSaldo
                                mrs_ProductosDetalle.Fields!cantidadFaltante = 0
                                mrs_ProductosDetalle.Fields!cantidadSobrante = 0
                                mrs_ProductosDetalle.Fields!EsHistoricoSaldo = 0
                                mrs_ProductosDetalle.Fields!nombreProducto = oRsTmp!nombre
                                mrs_ProductosDetalle.Fields!Codigo = lcCodigo
                                mrs_ProductosDetalle.Update
                            Else
                                lcMensaje = lcMensaje & "El código: " & lcCodigo & " Lote: " & lcLote & " F.Vencim: " & lcFvencimiento & " ya existen, solo se registrará el primero" & Chr(13)
                            End If
                        Else
                           lcMensaje = lcMensaje & "El código " & lcCodigo & " NO SE CARGO porque no tiene PRECIO" & Chr(13)
                        End If
                    Else
                        lcMensaje = lcMensaje & "El código " & lcCodigo & " NO EXISTE" & Chr(13)
                    End If
                         
                 
                 oRsFox99.MoveNext
            Loop
       End If
       oRsFox99.Close
       oConexionFox99.Close
       Set oRsFox99 = Nothing
       Set oConexionFox99 = Nothing
        
        
        
        If mrs_ProductosDetalle.RecordCount > 0 Then
            mrs_ProductosDetalle.Sort = "idProducto"
            With rs
                  .Fields.Append "IdProducto", adInteger, 4
                  .Fields.Append "Codigo", adVarChar, 20
                  .Fields.Append "NombreProducto", adChar, 300
                  .Fields.Append "idTipoSalidaBienInsumo", adInteger
                  .Fields.Append "Cantidad", adInteger
                  .Fields.Append "Precio", adDouble
                  .Fields.Append "Total", adDouble
                  .Fields.Append "CantidadSaldo", adInteger
                  .Fields.Append "CantidadFaltante", adInteger
                  .Fields.Append "CantidadSobrante", adInteger
                  .CursorType = adOpenKeyset
                  .LockType = adLockOptimistic
                  .Open
            End With
            mrs_ProductosDetalle.MoveFirst
            Do While Not mrs_ProductosDetalle.EOF
                lnIdProducto = mrs_ProductosDetalle!idProducto
                lnSaldo = 0
                lcCodigo = mrs_ProductosDetalle!Codigo
                lnPrecioUnitario = mrs_ProductosDetalle!precio
                lnIdTipoSalidaBienInsumo = mrs_ProductosDetalle!idTipoSalidaBienInsumo
                lcNombreProducto = mrs_ProductosDetalle!nombreProducto
                Do While Not mrs_ProductosDetalle.EOF And lnIdProducto = mrs_ProductosDetalle!idProducto
                  lnSaldo = lnSaldo + mrs_ProductosDetalle!Cantidad
                  mrs_ProductosDetalle.MoveNext
                  If mrs_ProductosDetalle.EOF Then
                     Exit Do
                  End If
                Loop
                rs.AddNew
                rs!idProducto = lnIdProducto
                rs!Codigo = lcCodigo
                rs!nombreProducto = lcNombreProducto
                rs!Cantidad = lnSaldo
                rs!precio = lnPrecioUnitario
                rs!Total = Round(lnSaldo * lnPrecioUnitario, 2)
                rs!idTipoSalidaBienInsumo = lnIdTipoSalidaBienInsumo
                rs!cantidadSaldo = 0
                rs!cantidadFaltante = 0
                rs!cantidadSobrante = 0
            Loop
            rs.MoveFirst
            Me.grdProductos.CargarItemsALaGrilla rs, True
        End If
        oConexion.Close
        Set oConexion = Nothing
        ActualizaREgistroSAnitario
        MsgBox "Terminó de cargar EXCEL (Ha cargado solo los que tienen PRECIO DE VENTA y QUE EXISTAN LOS CODIGOS DIGEMID" & Chr(13) & Chr(13) & lcMensaje, vbInformation, Me.Caption
    Else
        MsgBox "Solo se usa en el PRIMER INVENTARIO y TIPO_INVENTARIO=MANUAL. Además NO debe haber registrado NINGUN ITEM", vbInformation, Me.Caption
    End If
    Me.MousePointer = 1

End Sub

Private Sub btnCierre_Click()
   If MsgBox("Al CERRAR el Inventario, ya no se podrá modificar  " & Chr(13) & "¿Esta seguro?", vbQuestion + vbYesNo, "") = vbYes Then
        With mo_FarmInventario
            .FechaCierre = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .idEstadoInventario = 2     'Cerrado
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        If CierraInventario() Then
            MsgBox " Se CERRO EL INVENTARIO exitosamente", vbInformation, Me.Caption
            Me.Visible = False
        Else
            MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
        End If
   End If
End Sub

Private Sub btnImprimir_Click()
    ImprimeDocumento
End Sub

Function ValidacionesXlote() As Boolean
    ValidacionesXlote = True
    On Error Resume Next
    ms_MensajeError = ""
    If Me.txtRegSanitario.Text = "" Then
        ms_MensajeError = "Por favor ingrese el REGISTRO SANITARIO"
        txtRegSanitario.SetFocus
    ElseIf Len(Me.txtRegSanitario.Text) < 6 Then
        ms_MensajeError = "El REGISTRO SANITARIO debe ser mayor a 6 caracteres"
        txtRegSanitario.SetFocus
    ElseIf txtLote.Text = "" Then
        ms_MensajeError = "Por favor ingrese el Lote"
        txtLote.SetFocus
    ElseIf txtFvencimiento.Text = SIGHEntidades.FECHA_VACIA_DMY Then
        ms_MensajeError = "Por favor ingrese la Fecha de Vencimiento"
        txtFvencimiento.SetFocus
    ElseIf CDate(txtFvencimiento.Text) <= LdFechaMinimaVencimiento Then
        ms_MensajeError = "La Fecha de Vencimiento debe ser mayor a " & LdFechaMinimaVencimiento
        txtFvencimiento.SetFocus
    ElseIf Val(txtCantidad.Text) <= 0 Then
        ms_MensajeError = "Por favor ingrese la Cantidad"
        txtCantidad.SetFocus
    ElseIf Val(txtPrecio.Text) <= 0 Then
        ms_MensajeError = "Por favor ingrese el Precio" & Chr(13) & "En Fact_Config-->Catálogo Bienes e Insumos"
    ElseIf Trim(Me.cmbTipoSalida.Text) = "" Then
        ms_MensajeError = "Por favor elija el Tipo de Salida"
        Me.cmbTipoSalida.SetFocus
        
    End If
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       ValidacionesXlote = False
    End If
End Function


Private Sub btnImprimirInvConteo_Click()
   'Mariano 11112014
   Dim oRptClase As New rCrytalInventario
    oRptClase.MovTipo = ""
    oRptClase.Documento = txtNinventario.Text
    oRptClase.TextoDelFiltro = ""
    oRptClase.Almacen = cmbAlmacen.Text & " (" & Label13.Caption & ": " & Trim(cmbTipoInventario.Text) & ")"
    oRptClase.AlmacenO = ""
    oRptClase.IdAlmacenDestino = mo_cmbAlmacen.BoundText
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = ""
    oRptClase.Importe = 0
    oRptClase.TipoReporte = "Inventario"
    oRptClase.Rreportes = "InventarioC"
    oRptClase.Show vbModal
    Set oRptClase = Nothing

End Sub

Private Sub btnImprimirInvDet_Click()
    'Mariano 11112014
     Dim oRptClase As New rCrytalInventario
    oRptClase.MovTipo = ""
    oRptClase.Documento = txtNinventario.Text
    oRptClase.TextoDelFiltro = ""
    oRptClase.Almacen = cmbAlmacen.Text & " (" & Label13.Caption & ": " & Trim(cmbTipoInventario.Text) & ")"
    oRptClase.AlmacenO = ""
    oRptClase.IdAlmacenDestino = mo_cmbAlmacen.BoundText
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = ""
    oRptClase.Importe = 0
    oRptClase.TipoReporte = "Inventario"
    oRptClase.Rreportes = "InventarioD"
    oRptClase.Show vbModal
    Set oRptClase = Nothing
End Sub

Private Sub btnImprimirInvGeneral_Click()
    'Mariano 11112014
    Dim oRptClase As New rCrytalInventario
    oRptClase.MovTipo = ""
    oRptClase.Documento = txtNinventario.Text
    oRptClase.TextoDelFiltro = ""
    oRptClase.Almacen = cmbAlmacen.Text & " (" & Label13.Caption & ": " & Trim(cmbTipoInventario.Text) & ")"
    oRptClase.AlmacenO = ""
    oRptClase.IdAlmacenDestino = mo_cmbAlmacen.BoundText
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = ""
    oRptClase.Importe = 0
    oRptClase.TipoReporte = "Inventario"
    oRptClase.Rreportes = "InventarioG"
    oRptClase.Show vbModal
    Set oRptClase = Nothing

End Sub

Private Sub btnModificar_Click()
    If ValidacionesXlote = False Then
       Exit Sub
    End If
    Dim lbContinuar As Boolean
    Dim lnCantidadSaldo As Long, lnCantidadFaltante As Long, lnCantidadSobrante As Long
    lbContinuar = False
    If mrs_ProductosDetalle.RecordCount > 0 Then
        mrs_ProductosDetalle.MoveFirst
        Do While Not mrs_ProductosDetalle.EOF
           If mrs_ProductosDetalle.Fields!idProducto = lnCodigoProducto And Trim(mrs_ProductosDetalle!lote) = Trim(txtLote.Text) And mrs_ProductosDetalle!fechaVencimiento = txtFvencimiento.Text And mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = Me.cmbTipoSalida.ListIndex Then
              lbContinuar = True
              Exit Do
           End If
           mrs_ProductosDetalle.MoveNext
        Loop
    End If
    If lbContinuar = True Then
        With mrs_ProductosDetalle
            lnCantidadSaldo = .Fields!cantidadSaldo
            lnCantidadFaltante = .Fields!cantidadFaltante
            lnCantidadSobrante = .Fields!cantidadSobrante
            ActualizaCantidadesFaltantesYsobrantes Val(mo_cmbTipoInventario.BoundText), Val(txtCantidad.Text), _
                                                   lnCantidadSaldo, lnCantidadFaltante, lnCantidadSobrante
            .Fields!idProducto = lnCodigoProducto
            .Fields!lote = txtLote.Text
            .Fields!fechaVencimiento = txtFvencimiento.Text
            .Fields!Cantidad = Val(txtCantidad.Text)
            .Fields!precio = CDbl(txtPrecio.Text)
            .Fields!Total = Round(Val(txtCantidad.Text) * CDbl(txtPrecio.Text), 2)
            .Fields!registroSanitario = txtRegSanitario.Text
            .Fields!idTipoSalidaBienInsumo = Me.cmbTipoSalida.ListIndex
            .Fields!cantidadFaltante = lnCantidadFaltante
            .Fields!cantidadSobrante = lnCantidadSobrante
            .Update
        End With
        LimpiaDatos
        SumaCantidadesDeLotes
        Set grdProductosDetalle.DataSource = mrs_ProductosDetalle
        grdProductos.SetFocus
    End If

End Sub

Private Sub btnQuitar_Click()
    On Error Resume Next
    With mrs_ProductosDetalle
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
    Set grdProductosDetalle.DataSource = mrs_ProductosDetalle
    LimpiaDatos
    SumaCantidadesDeLotes
    
End Sub






Private Sub CargaInventarioExcel_Click()
    If mi_Opcion = sghAgregar And mo_cmbTipoInventario.BoundText = sghInventarioTipo.sghManual And _
                                                        mrs_ProductosDetalle.RecordCount = 0 Then
        Dim EXL As Excel.Application
        Set EXL = New Excel.Application
        Dim W As Excel.Workbook
        Set W = EXL.Workbooks.Open("c:\tformdetl.xls")
        Dim s As Excel.Worksheet
        Set s = W.Sheets("tformdetl")
        Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, lcMensaje As String
        Dim lcCodigo As String, lcFvencimiento As String, lnSaldo As Long, lcRegSanitario As String, lnPrecioUnitario As Double
        Dim oRsTmp As New Recordset, rs As New Recordset
        Dim oConexion As New Connection
        Dim lnIdProducto As Long, lcLote As String, lnIdTipoSalidaBienInsumo As Long, lcNombreProducto As String
        Dim lbEsNuevo As Boolean
        Me.MousePointer = 11
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        lnFila = 2
        lnFilaFinal = 10000
        lcMensaje = ""
        For lnFor = lnFila To lnFilaFinal
            lcRango = "E" + Trim(Str(lnFor))
            lcCodigo = Right("00000" & Trim(s.Range(lcRango).Value), 5)
            If Val(lcCodigo) = 0 Then
               Exit For
            End If
            lcRango = "F" + Trim(Str(lnFor))
            lcLote = Trim(s.Range(lcRango).Value)
            lcRango = "G" + Trim(Str(lnFor))
            lcFvencimiento = Trim(s.Range(lcRango).Value)
            lcRango = "H" + Trim(Str(lnFor))
            lnSaldo = Val(Trim(s.Range(lcRango).Value))
            lcRango = "I" + Trim(Str(lnFor))
            lcRegSanitario = Trim(s.Range(lcRango).Value)
            Set oRsTmp = mo_ReglasFacturacion.FactCatalogoBienesInsumosSeleccionarXcodigo(lcCodigo, oConexion)
            If oRsTmp.RecordCount > 0 Then
                lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(oRsTmp!idProducto, 3)
                If lnPrecioUnitario > 0 Then
                    lbEsNuevo = True
                    If mrs_ProductosDetalle.RecordCount > 0 Then
                       mrs_ProductosDetalle.MoveFirst
                       Do While Not mrs_ProductosDetalle.EOF
                          If mrs_ProductosDetalle!idProducto = oRsTmp.Fields!idProducto And _
                             Trim(mrs_ProductosDetalle!lote) = lcLote And _
                             mrs_ProductosDetalle!fechaVencimiento = CDate(lcFvencimiento) Then
                             lbEsNuevo = False
                             Exit Do
                          End If
                          mrs_ProductosDetalle.MoveNext
                       Loop
                    End If
                    If lbEsNuevo = True Then
                        mrs_ProductosDetalle.AddNew
                        mrs_ProductosDetalle.Fields!idProducto = oRsTmp.Fields!idProducto
                        mrs_ProductosDetalle.Fields!lote = lcLote
                        mrs_ProductosDetalle.Fields!fechaVencimiento = CDate(lcFvencimiento)
                        mrs_ProductosDetalle.Fields!Cantidad = lnSaldo
                        mrs_ProductosDetalle.Fields!precio = lnPrecioUnitario
                        mrs_ProductosDetalle.Fields!Total = Round(lnSaldo * lnPrecioUnitario, 2)
                        mrs_ProductosDetalle.Fields!registroSanitario = lcRegSanitario
                        mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = IIf(oRsTmp.Fields!idTipoSalidaBienInsumo = 3, 2, oRsTmp.Fields!idTipoSalidaBienInsumo)
                        mrs_ProductosDetalle.Fields!cantidadSaldo = lnSaldo
                        mrs_ProductosDetalle.Fields!cantidadFaltante = 0
                        mrs_ProductosDetalle.Fields!cantidadSobrante = 0
                        mrs_ProductosDetalle.Fields!EsHistoricoSaldo = 0
                        mrs_ProductosDetalle.Fields!nombreProducto = oRsTmp!nombre
                        mrs_ProductosDetalle.Fields!Codigo = lcCodigo
                        mrs_ProductosDetalle.Update
                    Else
                        lcMensaje = lcMensaje & "El código: " & lcCodigo & " Lote: " & lcLote & " F.Vencim: " & lcFvencimiento & " ya existen, solo se registrará el primero" & Chr(13)
                    End If
                Else
                   lcMensaje = lcMensaje & "El código " & lcCodigo & " NO SE CARGO porque no tiene PRECIO" & Chr(13)
                End If
            Else
                lcMensaje = lcMensaje & "El código " & lcCodigo & " NO EXISTE" & Chr(13)
            End If
        Next
        Set s = Nothing
'        W.Save
        W.Close
        Set W = Nothing
        Set EXL = Nothing
        If mrs_ProductosDetalle.RecordCount > 0 Then
            mrs_ProductosDetalle.Sort = "idProducto"
            With rs
                  .Fields.Append "IdProducto", adInteger, 4
                  .Fields.Append "Codigo", adVarChar, 20
                  .Fields.Append "NombreProducto", adChar, 300
                  .Fields.Append "idTipoSalidaBienInsumo", adInteger
                  .Fields.Append "Cantidad", adInteger
                  .Fields.Append "Precio", adDouble
                  .Fields.Append "Total", adDouble
                  .Fields.Append "CantidadSaldo", adInteger
                  .Fields.Append "CantidadFaltante", adInteger
                  .Fields.Append "CantidadSobrante", adInteger
                  .CursorType = adOpenKeyset
                  .LockType = adLockOptimistic
                  .Open
            End With
            mrs_ProductosDetalle.MoveFirst
            Do While Not mrs_ProductosDetalle.EOF
                lnIdProducto = mrs_ProductosDetalle!idProducto
                lnSaldo = 0
                lcCodigo = mrs_ProductosDetalle!Codigo
                lnPrecioUnitario = mrs_ProductosDetalle!precio
                lnIdTipoSalidaBienInsumo = mrs_ProductosDetalle!idTipoSalidaBienInsumo
                lcNombreProducto = mrs_ProductosDetalle!nombreProducto
                Do While Not mrs_ProductosDetalle.EOF And lnIdProducto = mrs_ProductosDetalle!idProducto
                  lnSaldo = lnSaldo + mrs_ProductosDetalle!Cantidad
                  mrs_ProductosDetalle.MoveNext
                  If mrs_ProductosDetalle.EOF Then
                     Exit Do
                  End If
                Loop
                rs.AddNew
                rs!idProducto = lnIdProducto
                rs!Codigo = lcCodigo
                rs!nombreProducto = lcNombreProducto
                rs!Cantidad = lnSaldo
                rs!precio = lnPrecioUnitario
                rs!Total = Round(lnSaldo * lnPrecioUnitario, 2)
                rs!idTipoSalidaBienInsumo = lnIdTipoSalidaBienInsumo
                rs!cantidadSaldo = 0
                rs!cantidadFaltante = 0
                rs!cantidadSobrante = 0
            Loop
            rs.MoveFirst
            Me.grdProductos.CargarItemsALaGrilla rs, True
        End If
        oConexion.Close
        Set oConexion = Nothing
        ActualizaREgistroSAnitario
        MsgBox "Terminó de cargar EXCEL (Ha cargado solo los que tienen PRECIO DE VENTA y QUE EXISTAN LOS CODIGOS DIGEMID" & Chr(13) & Chr(13) & lcMensaje, vbInformation, Me.Caption
    Else
        MsgBox "Solo se usa en el PRIMER INVENTARIO y TIPO_INVENTARIO=MANUAL. Además NO debe haber registrado NINGUN ITEM", vbInformation, Me.Caption
    End If
    Me.MousePointer = 1
End Sub

Sub ActualizaREgistroSAnitario()
On Error GoTo errRSACT
    Dim oConexionFox As New Connection
    Dim oRsFox As New Recordset
    Dim lcSql As String
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"
    oConexionFox.CursorLocation = adUseClient
    oRsFox.Open "select * from tmovimDet order by movFechUlt", oConexionFox, adOpenKeyset, adLockOptimistic
    If oRsFox.RecordCount > 0 And mrs_ProductosDetalle.RecordCount > 0 Then
       mrs_ProductosDetalle.MoveFirst
       Do While Not mrs_ProductosDetalle.EOF
          lcSql = "medcod='" & mrs_ProductosDetalle!Codigo & "' and medLote='" & _
                        Trim(mrs_ProductosDetalle!lote) & "' and medFechVto='" & _
                        Format(mrs_ProductosDetalle!fechaVencimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & "'"
          oRsFox.Filter = lcSql
          If oRsFox.RecordCount > 0 Then
             oRsFox.MoveLast
             If Not IsNull(oRsFox!medRegSan) Then
                mrs_ProductosDetalle!registroSanitario = Trim(oRsFox!medRegSan)
                mrs_ProductosDetalle.Update
             Else
                mrs_ProductosDetalle!registroSanitario = "Sin RS"
                mrs_ProductosDetalle.Update
             End If
          Else
                mrs_ProductosDetalle!registroSanitario = "Sin RS"
                mrs_ProductosDetalle.Update
          
          End If
          mrs_ProductosDetalle.MoveNext
       Loop
    End If
    oRsFox.Close
    oConexionFox.Close
    Set oRsFox = Nothing
    Set oConexionFox = Nothing
errRSACT:
'Resume
End Sub

Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen
End Sub







Private Sub cmdDelI_Click()
    On Error GoTo ErrDel
    oRsSaldosAjuste.Delete
    oRsSaldosAjuste.Update
ErrDel:
End Sub

Private Sub cmdLimpiaLista_Click()
    If MsgBox("Elimina toda la Lista de Medicamentos?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       LimpiaLista
    End If
End Sub

Private Sub cmbAlmacen_LostFocus()
    If Val(mo_cmbAlmacen.BoundText) > 0 And mi_Opcion = sghAgregar Then
       mo_reglasComunes.LlenaDataComboTipoSalidaBienSegunAlmacen Me.cmbTipoSalida, Val(mo_cmbAlmacen.BoundText), lcIdTipoSuministro
       mo_cmbTipoInventario.BoundText = sghInventarioTipo.sghManual
       lbLaFarmaciaEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacen.BoundText))
       InventarioAutomatico
       '
       
    End If
End Sub



'debb-29/12/2016
Sub InventarioAutomatico()
       grdProductos.idTipoInventario = sghInventarioTipo.sghManual
       lblMensaje.Visible = False
       Dim lbContinuar8 As Boolean
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim lcCodigo As String, lcRegistroSanitario As String, lnPrecioVenta As Double, lbPrimerItem As Boolean
       Dim lbEsNuevo As Boolean
       Set oRsTmp1 = mo_ReglasFarmacia.farmMovimientoSeleccionarPorIdAlmacenDestino(Val(mo_cmbAlmacen.BoundText))
       If oRsTmp1.RecordCount > 0 Then
          If MsgBox("Ya hubo un PRIMER INVENTARIO" & Chr(13) & "Esta seguro de generar INVENTARIO AUTOMATICO ?", vbQuestion + vbYesNo, "") = vbYes Then
             Me.MousePointer = 11
             Me.Refresh
             '**************Regenera Saldos *****************
             lblMensaje.Visible = True
             mo_cmbTipoInventario.BoundText = sghInventarioTipo.sghAutomatico
             lbContinuar8 = True
             Set oRsTmp2 = mo_reglasComunes.AuditoriaPorTablaFechas("FarmSaldo", Date, Now)
             If oRsTmp2.RecordCount > 0 Then
                 oRsTmp2.MoveFirst
                 Do While Not oRsTmp2.EOF
                 If InStr(oRsTmp2!Observaciones, mo_cmbAlmacen.BoundText) > 0 Then
    '                If MsgBox("Ya hubo REGENERADO DE SALDOS PARA ESA FARMACIA" & _
    '                           Chr(13) & "REGENERA SALDOS OTRA VEZ ?", vbQuestion + vbYesNo, "") <> vbYes Then
                        lbContinuar8 = False
                        Exit Do
    '                End If
                 End If
                 oRsTmp2.MoveNext
                 Loop
             End If
             oRsTmp2.Close
             If lbContinuar8 = True Then
                    Dim oRegenerarSaldo As New HerrRegeneraSaldos
                    oRegenerarSaldo.EsperaApulsarACEPTAR = True
                    oRegenerarSaldo.idUsuario = ml_idUsuario
                    oRegenerarSaldo.lcNombrePc = mo_lcNombrePc
                    oRegenerarSaldo.IdAlmacenAregenerar = Val(mo_cmbAlmacen.BoundText)
                    oRegenerarSaldo.RegeneraDesdeUltimoMes = True
                    oRegenerarSaldo.FormularioUsadoDesdeOtroFrm = True
                    oRegenerarSaldo.Show 1
                    Set oRegenerarSaldo = Nothing
             End If
             oRsTmp1.Close
             Me.Refresh
             '**************Datos de la tabla FarmInventarioCabecera *****************
             grdProductos.idTipoInventario = sghInventarioTipo.sghAutomatico
             grdProductos.CargaProductosPorIdAlmacen Val(mo_cmbAlmacen.BoundText)
             '**************Datos de la tabla FarmInventarioDetalle *****************
             GenerarRecordsetTemporal
             Dim oConexion As New Connection
             oConexion.CommandTimeout = 300
             oConexion.CursorLocation = adUseClient
             oConexion.Open SIGHEntidades.CadenaConexion
             Set oRsTmp1 = mo_ReglasFarmacia.FarmDevuelveSaldosConSinLotesPorAlmacen(Val(mo_cmbAlmacen.BoundText), 0, 0, 0)
             If oRsTmp1.RecordCount > 0 Then
                lbPrimerItem = True
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                   lcCodigo = oRsTmp1.Fields!Codigo
                   Do While Not oRsTmp1.EOF And lcCodigo = oRsTmp1.Fields!Codigo
                        '
                        lcRegistroSanitario = ""
                        Set oRsTmp2 = mo_ReglasFarmacia.ExportaPreciosSismedRegSant(oRsTmp1.Fields!idProducto, oConexion)
                        If oRsTmp2.RecordCount > 0 Then
                           lcRegistroSanitario = oRsTmp2!registroSanitario
                        End If
                        oRsTmp2.Close
                        '
If oRsTmp1.Fields!idProducto = 720 And Trim(oRsTmp1.Fields!lote) = "111" Then
lnPrecioVenta = 0
End If
                        lnPrecioVenta = 0
                        Set oRsTmp2 = mo_reglasComunes.FactCatalogoBienesInsumosHospXfiltro("idProducto=" & oRsTmp1!idProducto & " and PrecioUnitario>0", oConexion)
                        If oRsTmp2.RecordCount > 0 Then
                           lnPrecioVenta = oRsTmp2!PrecioUnitario
                        End If
                        oRsTmp2.Close
                        '
                        mrs_ProductosDetalle.AddNew
                        mrs_ProductosDetalle.Fields!idProducto = oRsTmp1.Fields!idProducto
                        mrs_ProductosDetalle.Fields!lote = oRsTmp1.Fields!lote
                        mrs_ProductosDetalle.Fields!fechaVencimiento = oRsTmp1.Fields!fechaVencimiento
                        mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = oRsTmp1.Fields!IdTipoSalidaBienInsumoSaldo
                        mrs_ProductosDetalle.Fields!Cantidad = 0
                        mrs_ProductosDetalle.Fields!precio = lnPrecioVenta 'oRsTmp1.Fields!Precio
                        mrs_ProductosDetalle.Fields!Total = 0
                        mrs_ProductosDetalle.Fields!registroSanitario = lcRegistroSanitario
                        mrs_ProductosDetalle!cantidadSobrante = 0
                        mrs_ProductosDetalle!EsHistoricoSaldo = 1
                        mrs_ProductosDetalle!cantidadSaldo = oRsTmp1.Fields!cantidadLote
                        mrs_ProductosDetalle!cantidadFaltante = oRsTmp1.Fields!cantidadLote
                        mrs_ProductosDetalle.Update
                        
'                        lbEsNuevo = True
'                        If lbPrimerItem = True Then
'                            mrs_ProductosDetalle.MoveFirst
'                            Do While Not mrs_ProductosDetalle.EOF
'                               If mrs_ProductosDetalle.Fields!idProducto Then
'                                  lbEsNuevo = False
'                                  Exit Do
'                               End If
'                               mrs_ProductosDetalle.MoveNext
'                            Loop
'                        Else
'                            lbPrimerItem = False
'                        End If
'                        '
'                        If lbEsNuevo = True Then
'                            mrs_ProductosDetalle.AddNew
'                            mrs_ProductosDetalle.Fields!idProducto = oRsTmp1.Fields!idProducto
'                            mrs_ProductosDetalle.Fields!Lote = oRsTmp1.Fields!Lote
'                            mrs_ProductosDetalle.Fields!FechaVencimiento = oRsTmp1.Fields!FechaVencimiento
'                            mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = oRsTmp1.Fields!idTipoSalidaBienInsumo
'                            mrs_ProductosDetalle.Fields!Cantidad = 0
'                            mrs_ProductosDetalle.Fields!Precio = lnPrecioVenta 'oRsTmp1.Fields!Precio
'                            mrs_ProductosDetalle.Fields!total = 0
'                            mrs_ProductosDetalle.Fields!RegistroSanitario = lcRegistroSanitario
'                            mrs_ProductosDetalle!CantidadSobrante = 0
'                            mrs_ProductosDetalle!EsHistoricoSaldo = 1
'                        End If
'                        mrs_ProductosDetalle!CantidadSaldo = mrs_ProductosDetalle!CantidadSaldo + oRsTmp1.Fields!cantidadLote
'                        mrs_ProductosDetalle!CantidadFaltante = mrs_ProductosDetalle!CantidadFaltante + oRsTmp1.Fields!cantidadLote
'                        mrs_ProductosDetalle.Update
                        oRsTmp1.MoveNext
                        If oRsTmp1.EOF Then
                           Exit Do
                        End If
                   Loop
                Loop
             End If
             oConexion.Close
             Set oConexion = Nothing
             
             Me.MousePointer = 1
          Else
             cmbAlmacen.Text = ""
             cmbAlmacen.SetFocus
          End If
       Else
          Me.CargaInventarioExcel.Enabled = True
          Me.btnCArgaDesdeSismedv2.Enabled = True
       End If
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       lblMensaje.Visible = False


End Sub

Private Sub cmdActualizaInventarioTemp_Click()
    If oRsIventarioTmp.RecordCount > 0 Then
       Dim rs As Recordset
       oRsIventarioTmp.MoveFirst
       Do While Not oRsIventarioTmp.EOF
          If oRsIventarioTmp!elegir = True Then
             Set rs = mo_ReglasFarmacia.FarmInventarioCabeceraDevuelveProductosPorId(oRsIventarioTmp!IdInventario)
             rs.Filter = "cantidad>0"
             If rs.RecordCount > 0 Then
                grdProductos.CargaProductosDelInventarioTemporal rs
                '
                'eliminando DETALLE
                rs.MoveFirst
                Do While Not rs.EOF
                    mrs_ProductosDetalle.Filter = "idProducto=" & rs!idProducto
                   If mrs_ProductosDetalle.RecordCount > 0 Then
                       mrs_ProductosDetalle.MoveFirst
                      Do While Not mrs_ProductosDetalle.EOF
                          mrs_ProductosDetalle.Delete
                          mrs_ProductosDetalle.Update
                          mrs_ProductosDetalle.MoveNext
                      Loop
                   End If
                   rs.MoveNext
                Loop
                'agregando DETALLE
                CargaProductosDetalle oRsIventarioTmp!IdInventario, True
                '
             End If
             rs.Close
          End If
          oRsIventarioTmp.MoveNext
       Loop
       oRsIventarioTmp.MoveFirst
       Set rs = Nothing
    End If
End Sub

Private Sub cmdAdicionarItem_Click()
    If ValidacionesXlote = False Then
       Exit Sub
    End If
    If mrs_ProductosDetalle.RecordCount > 0 Then
        mrs_ProductosDetalle.MoveFirst
        Do While Not mrs_ProductosDetalle.EOF
           If mrs_ProductosDetalle.Fields!idProducto = lnCodigoProducto And Trim(mrs_ProductosDetalle!lote) = txtLote.Text And mrs_ProductosDetalle!fechaVencimiento = txtFvencimiento.Text And mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = Me.cmbTipoSalida.ListIndex Then
              MsgBox "Ya existe ese 'Lote/FechaVencimiento/TipoSalida' para el Producto", vbInformation, "Inventario"
              Exit Sub
           End If
           mrs_ProductosDetalle.MoveNext
        Loop
    End If
    Dim lnCantidadSaldo As Long, lnCantidadFaltante As Long, lnCantidadSobrante As Long
    With mrs_ProductosDetalle
        ActualizaCantidadesFaltantesYsobrantes Val(mo_cmbTipoInventario.BoundText), Val(txtCantidad.Text), _
                                               lnCantidadSaldo, lnCantidadFaltante, lnCantidadSobrante
        .AddNew
        .Fields!idProducto = lnCodigoProducto
        .Fields!lote = txtLote.Text
        .Fields!fechaVencimiento = txtFvencimiento.Text
        .Fields!Cantidad = Val(txtCantidad.Text)
        .Fields!precio = CDbl(txtPrecio.Text)
        .Fields!Total = Round(Val(txtCantidad.Text) * CDbl(txtPrecio.Text), 2)
        .Fields!registroSanitario = txtRegSanitario.Text
        .Fields!idTipoSalidaBienInsumo = Me.cmbTipoSalida.ListIndex
        .Fields!idTipoSalidaBienInsumo = Me.cmbTipoSalida.ListIndex
        .Fields!EsHistoricoSaldo = 0
        .Fields!cantidadFaltante = lnCantidadFaltante
        .Fields!cantidadSobrante = lnCantidadSobrante
        .Update
    End With
    LimpiaDatos
    SumaCantidadesDeLotes
    Set grdProductosDetalle.DataSource = mrs_ProductosDetalle
    grdProductos.SetFocus

End Sub

Sub ActualizaCantidadesFaltantesYsobrantes(lnTipoInventario As sghInventarioTipo, lnCantidad As Long, _
                                           lnCantidadSaldo As Long, _
                                           ByRef lnCantidadFaltante As Long, ByRef lnCantidadSobrante As Long)
        lnCantidadFaltante = 0
        lnCantidadSobrante = 0
        If lnTipoInventario = sghInventarioTipo.sghAutomatico Then
               If lnCantidadSaldo < lnCantidad Then
                  lnCantidadFaltante = 0
                  lnCantidadSobrante = lnCantidad - lnCantidadSaldo
               Else
                  lnCantidadFaltante = lnCantidadSaldo - lnCantidad
                  lnCantidadSobrante = 0
               End If
        End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
    Set mo_cmbTipoSalida.MiComboBox = cmbTipoSalida
    Set mo_cmbTipoInventario.MiComboBox = cmbTipoInventario
End Sub


Private Sub Form_Load()
    LdFechaMinimaVencimiento = Date + Val(lcBuscaParametro.SeleccionaFilaParametro(224))
    CargarComboBoxes
    ConfigurarGrdProductos
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Inventario"
    Case sghModificar
        Me.Caption = "Modificar Inventario"
    Case sghConsultar
        Me.Caption = "Consultar Inventario"
    Case sghEliminar
        Me.Caption = "Anular Inventario"
    End Select
    CargarDatosAlFormulario
    InventarioTemporales
End Sub



Sub ConfigurarGrdProductos()
    grdProductos.IdInventario = ml_IdInventario
    grdProductos.Inicializar
    grdProductos.TipoPrecioParaNiNs = 3   'Venta
End Sub


Sub InventarioTemporales()
     chkInventarioTemp.Enabled = False
     grdIventariosTemp.Enabled = False
     cmdActualizaInventarioTemp.Enabled = False
     If mi_Opcion = sghAgregar Then
        chkInventarioTemp.Enabled = True
     Else
        If Val(txtNinventario.Text) > 0 Then
            Dim oRsTmp1 As New Recordset
            grdIventariosTemp.Enabled = True
            cmdActualizaInventarioTemp.Enabled = True
            With oRsIventarioTmp
                .Fields.Append "IdInventario", adInteger, 4
                .Fields.Append "Inventario", adVarChar, 20
                .Fields.Append "Elegir", adBoolean
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open
            End With
            Set oRsTmp1 = mo_ReglasFarmacia.farmInventarioSeleccionarPorIdAlmacen(Val(mo_cmbAlmacen.BoundText))
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                  If Val(oRsTmp1!NumeroInventario) = 0 And oRsTmp1!idEstadoInventario = sghEstadoTabla.sghRegistrado Then
                     oRsIventarioTmp.AddNew
                     oRsIventarioTmp!IdInventario = oRsTmp1!IdInventario
                     oRsIventarioTmp!Inventario = Left(oRsTmp1!NumeroInventario & " " & Format(oRsTmp1!fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM), 20)
                     oRsIventarioTmp!elegir = False
                     oRsIventarioTmp.Update
                  End If
                  oRsTmp1.MoveNext
               Loop
               oRsTmp1.Close
            End If
            'oRsTmp1.Close
            Set grdIventariosTemp.DataSource = oRsIventarioTmp
            If oRsIventarioTmp.RecordCount > 0 Then
               oRsIventarioTmp.MoveFirst
            End If
            mo_Apariencia.ConfigurarFilasBiColores Me.grdIventariosTemp, SIGHEntidades.GrillaConFilasBicolor
            grdIventariosTemp.Caption = ""
            grdIventariosTemp.Bands(0).Columns("idInventario").Hidden = True
            grdIventariosTemp.Bands(0).Columns("Inventario").Width = 1500
            grdIventariosTemp.Bands(0).Columns("Inventario").Activation = ssActivationActivateNoEdit
            grdIventariosTemp.Bands(0).Columns("Inventario").Header.Caption = "Inventario Temporal"
            grdIventariosTemp.Bands(0).Columns("Elegir").Width = 500
            
            Set oRsTmp1 = Nothing
        Else
           chkInventarioTemp.Value = 1
           btnCierre.Enabled = False
        End If
     End If

End Sub

Sub CargarDatosAlFormulario()
     
     mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
     mo_Formulario.HabilitarDeshabilitar Me.txtNinventario, False
     mo_Formulario.HabilitarDeshabilitar Me.txtFcierre, False
     mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
     mo_Formulario.HabilitarDeshabilitar Me.txtPrecio, False
     mo_Formulario.HabilitarDeshabilitar Me.txtFmodificacion, False
     mo_Formulario.HabilitarDeshabilitar Me.cmbTipoInventario, False
     If mi_Opcion = sghConsultar Or mi_Opcion = sghEliminar Then
            cmdAdicionarItem.Enabled = False
            Me.btnModificar.Enabled = False
            Me.btnQuitar.Enabled = False
     End If
     grdProductos.IdPuntoCarga = 5    'Farmacia
     btnCierre.Enabled = False
     GenerarRecordsetTemporal
     Select Case mi_Opcion
     Case sghAgregar
        txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL      'Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
        grdProductos.IdInventario = 0
        
        grdProductos.LimpiarGrilla
        grdProductos.CargaProductosPorIdInventario
     Case sghModificar
        btnCierre.Enabled = True
        cmbAlmacen.Enabled = False
        CargarDatosALosControles
     Case sghConsultar
        CargarDatosALosControles
        cmbAlmacen.Enabled = False
        btnAceptar.Enabled = False
     Case sghEliminar
        CargarDatosALosControles
        cmbAlmacen.Enabled = False
 End Select
End Sub

Sub CargarDatosALosControles()
   mo_FarmInventario.IdInventario = ml_IdInventario
   If Not mo_ReglasFarmacia.FarmInventarioSeleccionarPorId(mo_FarmInventario) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   mo_reglasComunes.LlenaDataComboTipoSalidaBienSegunAlmacen Me.cmbTipoSalida, mo_FarmInventario.IdAlmacen, lcIdTipoSuministro
   CargaDatos
   btnImprimir.Visible = True
End Sub


Sub CargaDatos()
   '**************Datos de la tabla FarmInventario *****************
   lbLaFarmaciaEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(mo_FarmInventario.IdAlmacen)
   txtNinventario.Text = mo_FarmInventario.NumeroInventario
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDelInventario("idEstadoInventario=" & mo_FarmInventario.idEstadoInventario)
   mo_cmbAlmacen.BoundText = mo_FarmInventario.IdAlmacen
   txtFregistro.Text = Format(mo_FarmInventario.fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   If mo_FarmInventario.fechaModificacion <> 0 Then txtFmodificacion.Text = Format(mo_FarmInventario.fechaModificacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   If mo_FarmInventario.FechaCierre <> 0 Then txtFcierre.Text = Format(mo_FarmInventario.FechaCierre, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   ml_IdInventario = mo_FarmInventario.IdInventario
   If mo_FarmInventario.idEstadoInventario = 0 Or mo_FarmInventario.idEstadoInventario = 2 Then
      btnAceptar.Enabled = False
      btnCierre.Enabled = False
   End If
   mo_cmbTipoInventario.BoundText = mo_FarmInventario.idTipoInventario
   '**************Datos de la tabla FarmInventarioCabecera *****************
   grdProductos.idTipoInventario = mo_FarmInventario.idTipoInventario
   grdProductos.IdInventario = ml_IdInventario
   grdProductos.CargaProductosPorIdInventario
   '**************Datos de la tabla FarmInventarioDetalle *****************
   CargaProductosDetalle ml_IdInventario, False
'   Dim oRsTmp As New ADODB.Recordset
'   Set oRsTmp = mo_ReglasFarmacia.FarmInventarioDetalleDevuelveProductosLotesPorId(ml_IdInventario)
'   If oRsTmp.RecordCount > 0 Then
'      oRsTmp.MoveFirst
'
'      Do While Not oRsTmp.EOF
'            mrs_ProductosDetalle.AddNew
'            mrs_ProductosDetalle.Fields!idProducto = oRsTmp.Fields!idProducto
'            mrs_ProductosDetalle.Fields!LOTE = oRsTmp.Fields!LOTE
'            mrs_ProductosDetalle.Fields!FechaVencimiento = oRsTmp.Fields!FechaVencimiento
'            mrs_ProductosDetalle.Fields!Cantidad = oRsTmp.Fields!Cantidad
'            mrs_ProductosDetalle.Fields!Precio = oRsTmp.Fields!Precio
'            mrs_ProductosDetalle.Fields!total = Round(oRsTmp.Fields!Cantidad * oRsTmp.Fields!Precio, 2)
'            mrs_ProductosDetalle.Fields!registroSanitario = IIf(IsNull(oRsTmp.Fields!registroSanitario), "", oRsTmp.Fields!registroSanitario)
'            mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = oRsTmp.Fields!idTipoSalidaBienInsumo
'            mrs_ProductosDetalle.Fields!CantidadSaldo = IIf(IsNull(oRsTmp.Fields!CantidadSaldo), 0, oRsTmp.Fields!CantidadSaldo)
'            mrs_ProductosDetalle.Fields!CantidadFaltante = IIf(IsNull(oRsTmp.Fields!CantidadFaltante), 0, oRsTmp.Fields!CantidadFaltante)
'            mrs_ProductosDetalle.Fields!CantidadSobrante = IIf(IsNull(oRsTmp.Fields!CantidadSobrante), 0, oRsTmp.Fields!CantidadSobrante)
'            mrs_ProductosDetalle.Fields!EsHistoricoSaldo = IIf(IsNull(oRsTmp.Fields!EsHistoricoSaldo), 0, oRsTmp.Fields!EsHistoricoSaldo)
'            mrs_ProductosDetalle.Update
'            oRsTmp.MoveNext
'       Loop
'   End If
'   oRsTmp.Close
'   Set oRsTmp = Nothing
   If mo_FarmInventario.idTipoInventario = SIGHEntidades.sghInventarioTipo.sghAutomatico Then
        btnImprimirInvConteo.Visible = True
        btnImprimirInvDet.Visible = True
        btnImprimirInvGeneral.Visible = True
   End If
   cmbAlmacen_LostFocus
End Sub

Sub CargaProductosDetalle(lnIdInventario As Long, lbSoloMayoresAcero As Boolean)
   Dim oRsTmp As New ADODB.Recordset
   Set oRsTmp = mo_ReglasFarmacia.FarmInventarioDetalleDevuelveProductosLotesPorId(lnIdInventario)
   If lbSoloMayoresAcero = True Then
      oRsTmp.Filter = "cantidad>0"
   End If
   If oRsTmp.RecordCount > 0 Then
      oRsTmp.MoveFirst
      Do While Not oRsTmp.EOF
            mrs_ProductosDetalle.AddNew
            mrs_ProductosDetalle.Fields!idProducto = oRsTmp.Fields!idProducto
            mrs_ProductosDetalle.Fields!lote = oRsTmp.Fields!lote
            mrs_ProductosDetalle.Fields!fechaVencimiento = oRsTmp.Fields!fechaVencimiento
            mrs_ProductosDetalle.Fields!Cantidad = oRsTmp.Fields!Cantidad
            mrs_ProductosDetalle.Fields!precio = oRsTmp.Fields!precio
            mrs_ProductosDetalle.Fields!Total = Round(oRsTmp.Fields!Cantidad * oRsTmp.Fields!precio, 2)
            mrs_ProductosDetalle.Fields!registroSanitario = IIf(IsNull(oRsTmp.Fields!registroSanitario), "", oRsTmp.Fields!registroSanitario)
            mrs_ProductosDetalle.Fields!idTipoSalidaBienInsumo = oRsTmp.Fields!idTipoSalidaBienInsumo
            mrs_ProductosDetalle.Fields!cantidadSaldo = IIf(IsNull(oRsTmp.Fields!cantidadSaldo), 0, oRsTmp.Fields!cantidadSaldo)
            mrs_ProductosDetalle.Fields!cantidadFaltante = IIf(IsNull(oRsTmp.Fields!cantidadFaltante), 0, oRsTmp.Fields!cantidadFaltante)
            mrs_ProductosDetalle.Fields!cantidadSobrante = IIf(IsNull(oRsTmp.Fields!cantidadSobrante), 0, oRsTmp.Fields!cantidadSobrante)
            mrs_ProductosDetalle.Fields!EsHistoricoSaldo = IIf(IsNull(oRsTmp.Fields!EsHistoricoSaldo), 0, oRsTmp.Fields!EsHistoricoSaldo)
            mrs_ProductosDetalle.Update
            oRsTmp.MoveNext
       Loop
   End If
   oRsTmp.Close
   Set oRsTmp = Nothing

End Sub

Sub GenerarRecordsetTemporal()
   If mrs_ProductosDetalle.State = 1 Then Set mrs_ProductosDetalle = Nothing
   With mrs_ProductosDetalle
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "NumeroDocumento", adChar, 20
          .Fields.Append "Lote", adChar, 15
          .Fields.Append "FechaVencimiento", adDate
          .Fields.Append "IdTipoSalidaBienInsumo", adInteger
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "RegistroSanitario", adVarChar, 50
          .Fields.Append "CantidadSaldo", adInteger
          .Fields.Append "CantidadFaltante", adInteger
          .Fields.Append "CantidadSobrante", adInteger
          .Fields.Append "EsHistoricoSaldo", adInteger
          .Fields.Append "nombreProducto", adVarChar, 300
          .Fields.Append "codigo", adVarChar, 20
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    mo_Apariencia.ConfigurarFilasBiColores Me.grdProductosDetalle, SIGHEntidades.GrillaConFilasBicolor
End Sub

Sub CargarComboBoxes()
    '
    Set oRsItemsUnidosis = mo_ReglasFarmacia.farmUnidosisSeleccionarTodos
    '
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    '
    mo_cmbTipoSalida.BoundColumn = "IdTipoSalidaBienInsumo"
    mo_cmbTipoSalida.ListField = "Tipo"
    Set mo_cmbTipoSalida.RowSource = mo_ReglasFarmacia.farmTipoSalidaBienInsumoDevuelveTodos
    '
    mo_cmbTipoInventario.BoundColumn = "idTipoInventario"
    mo_cmbTipoInventario.ListField = "Descripcion"
    Set mo_cmbTipoInventario.RowSource = mo_ReglasFarmacia.farmTipoInventarioSeleccionarTodos
End Sub








Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdProductos_OnClick(oRecordset As ADODB.Recordset)
    If oRecordset.RecordCount = 0 Then Exit Sub
    grdProductosDetalle.Caption = Trim(oRecordset.Fields!Codigo) & " - " & oRecordset.Fields!nombreProducto
    lnCodigoProducto = oRecordset.Fields!idProducto
    txtPrecio.Text = oRecordset.Fields!precio
    mrs_ProductosDetalle.Filter = "idProducto=" & lnCodigoProducto
    Set grdProductosDetalle.DataSource = mrs_ProductosDetalle
    LimpiaDatos

    If oRecordset!esPaquete = True Then
       txtRegSanitario.Text = WxREGSANITARIOpaquete
       txtLote.Text = WxLOTEpaquete
       txtFvencimiento.Text = WxFVENCIMIENTOpaquete
    End If
    mo_Formulario.HabilitarDeshabilitar txtLote, True
    mo_Formulario.HabilitarDeshabilitar txtFvencimiento, True
    mo_Formulario.HabilitarDeshabilitar cmbTipoSalida, True
    If Val(lcIdTipoSuministro) = 2 Then   'donaciones
       Me.cmbTipoSalida.ListIndex = 4
       mo_Formulario.HabilitarDeshabilitar cmbTipoSalida, False
    Else
       If oRecordset.Fields("idTipoSalidaBienInsumo").Value <> sghTipoSalidaItemFarmacia.sghVentaEstrategico Then
           Me.cmbTipoSalida.ListIndex = SIGHEntidades.ElijeSiEsEstrategicoDevuelveId(oRecordset.Fields("idTipoSalidaBienInsumo").Value)
           mo_Formulario.HabilitarDeshabilitar cmbTipoSalida, False
       Else
           Me.cmbTipoSalida.ListIndex = 0
        End If
    End If
    If mrs_ProductosDetalle.RecordCount > 0 Then
       mrs_ProductosDetalle.MoveFirst
    End If
    cmdAdicionarItem.Enabled = True
    Me.btnModificar.Enabled = False
    Me.btnQuitar.Enabled = False
    fraDetalleLote.Caption = "Agregar"
    txtRegSanitario.SetFocus

End Sub


Private Sub grdProductosDetalle_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdProductosDetalle_DblClick()
   If Not mrs_ProductosDetalle.EOF Or Not mrs_ProductosDetalle.BOF Then
        txtLote.Text = mrs_ProductosDetalle.Fields!lote
        txtFvencimiento.Text = mrs_ProductosDetalle.Fields!fechaVencimiento
        txtRegSanitario.Text = mrs_ProductosDetalle.Fields!registroSanitario
        txtPrecio.Text = mrs_ProductosDetalle.Fields!precio
        txtCantidad.Text = mrs_ProductosDetalle.Fields!Cantidad
'        btnAgregar.Enabled = False
        Me.cmbTipoSalida.ListIndex = SIGHEntidades.ElijeSiEsEstrategicoDevuelveId(mrs_ProductosDetalle.Fields("idTipoSalidaBienInsumo").Value)
        mo_Formulario.HabilitarDeshabilitar cmbTipoSalida, False
        mo_Formulario.HabilitarDeshabilitar txtLote, False
        mo_Formulario.HabilitarDeshabilitar txtFvencimiento, False
        cmdAdicionarItem.Enabled = False
        Me.btnModificar.Enabled = True
        Me.btnQuitar.Enabled = IIf(mrs_ProductosDetalle!EsHistoricoSaldo = 1, False, True)
        fraDetalleLote.Caption = "Modificar/Quitar"
   End If
End Sub

Private Sub grdProductosDetalle_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     grdProductosDetalle.Bands(0).Columns("IdProducto").Hidden = True
     grdProductosDetalle.Bands(0).Columns("IdProducto").Activation = ssActivationActivateNoEdit
     '
     On Error Resume Next
     grdProductosDetalle.ValueLists.Add "TipoSalida"
     grdProductosDetalle.ValueLists("TipoSalida").ValueListItems.Add 1, "Ventas"
     grdProductosDetalle.ValueLists("TipoSalida").ValueListItems.Add 2, "IntervSanit"
     grdProductosDetalle.ValueLists("TipoSalida").ValueListItems.Add 3, "Vtas/IntervSanit"
     grdProductosDetalle.ValueLists("TipoSalida").ValueListItems.Add 4, "Donaciones"
     grdProductosDetalle.Bands(0).Columns("idTipoSalidaBienInsumo").ValueList = "TipoSalida"
     grdProductosDetalle.Bands(0).Columns("idTipoSalidaBienInsumo").Style = ssStyleDropDownList
     grdProductosDetalle.Bands(0).Columns("idTipoSalidaBienInsumo").Activation = ssActivationActivateNoEdit
     grdProductosDetalle.Bands(0).Columns("idTipoSalidaBienInsumo").Width = 800
     grdProductosDetalle.Bands(0).Columns("idTipoSalidaBienInsumo").Header.Caption = "Tipo"
     '
     grdProductosDetalle.Bands(0).Columns("Lote").Activation = ssActivationActivateNoEdit
     grdProductosDetalle.Bands(0).Columns("FechaVencimiento").Activation = ssActivationActivateNoEdit
     grdProductosDetalle.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
     grdProductosDetalle.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
     grdProductosDetalle.Bands(0).Columns("Precio").Format = "#0.00"
     grdProductosDetalle.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
     grdProductosDetalle.Bands(0).Columns("Total").Format = "#0.00"
     grdProductosDetalle.Bands(0).Columns("RegistroSanitario").Activation = ssActivationActivateNoEdit
     grdProductosDetalle.Bands(0).Columns("CantidadSaldo").Hidden = True
     grdProductosDetalle.Bands(0).Columns("CantidadFaltante").Hidden = True
     grdProductosDetalle.Bands(0).Columns("CantidadSobrante").Hidden = True
     grdProductosDetalle.Bands(0).Columns("EsHistoricoSaldo").Hidden = True
     grdProductosDetalle.Bands(0).Columns("NumeroDocumento").Hidden = True
End Sub

Private Sub grdProductosDetalle_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdProductosDetalle_DblClick
    End If
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub


Sub LimpiaDatos()
    txtLote.Text = ""
    txtFvencimiento.Text = SIGHEntidades.FECHA_VACIA_DMY
    txtRegSanitario.Text = ""
    txtCantidad.Text = ""
'    btnAgregar.Enabled = True
'    btnQuitar.Enabled = True
End Sub


Private Sub txtFvencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFvencimiento
End Sub

Private Sub txtFvencimiento_LostFocus()
    If txtFvencimiento <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFvencimiento, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFvencimiento = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub





Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtLote
End Sub



Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtRegSanitario_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRegSanitario
End Sub

Sub SumaCantidadesDeLotes()
    Dim lnTotCantidad As Long, lnTotalFaltantes As Long, lnTotalSobrantes As Long, lnTotalSaldos    As Long
    lnTotCantidad = 0: lnTotalFaltantes = 0: lnTotalSobrantes = 0: lnTotalSaldos = 0
    If mrs_ProductosDetalle.RecordCount > 0 Then
       mrs_ProductosDetalle.MoveFirst
       Do While Not mrs_ProductosDetalle.EOF
          lnTotCantidad = lnTotCantidad + mrs_ProductosDetalle.Fields!Cantidad
          lnTotalFaltantes = lnTotalFaltantes + mrs_ProductosDetalle.Fields!cantidadFaltante
          lnTotalSobrantes = lnTotalSobrantes + mrs_ProductosDetalle.Fields!cantidadSobrante
          lnTotalSaldos = lnTotalSaldos + mrs_ProductosDetalle.Fields!cantidadSaldo
          mrs_ProductosDetalle.MoveNext
       Loop
       ActualizaCantidadesFaltantesYsobrantes Val(mo_cmbTipoInventario.BoundText), lnTotCantidad, _
                                              lnTotalSaldos, lnTotalFaltantes, lnTotalSobrantes
       mrs_ProductosDetalle.MoveFirst
    End If
    grdProductos.ActualizaCantidadTotalDeLotes lnTotCantidad, lnCodigoProducto, lnTotalSaldos, _
                                               lnTotalFaltantes, lnTotalSobrantes
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
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

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_Apariencia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbAlmacen = Nothing
    Set mrs_ProductosCabecera = Nothing
    Set mrs_ProductosDetalle = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_FarmInventario = Nothing
End Sub








Sub LimpiaLista()
        If oRsSaldosAjuste.RecordCount > 0 Then
          oRsSaldosAjuste.MoveFirst
          Do While Not oRsSaldosAjuste.EOF
             oRsSaldosAjuste.Delete
             oRsSaldosAjuste.Update
             oRsSaldosAjuste.MoveNext
          Loop
        End If

End Sub




Function AgregaDatosDeNotaSalidaAI(oDoMovimiento As DoFarmMovimiento, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New farmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
    End With
    
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: mo_mensajeError = oMovimiento.MensajeError: GoTo TerminarNS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.Cantidad = oRsDetalleProductos.Fields!Cantidad
       oDoMovimientoDetalle.fechaVencimiento = oRsDetalleProductos.Fields!fechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.lote = oRsDetalleProductos.Fields!lote
       oDoMovimientoDetalle.precio = oRsDetalleProductos.Fields!precio
       oDoMovimientoDetalle.registroSanitario = ""
       oDoMovimientoDetalle.Total = oRsDetalleProductos.Fields!Total
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: mo_mensajeError = oMovimientoDetalle.MensajeError: GoTo TerminarNS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarNS:
    If bProcesoOK Then
        AgregaDatosDeNotaSalidaAI = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeNotaSalidaAI = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function



