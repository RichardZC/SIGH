VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form aForm1 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13860
   Icon            =   "aform1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   13860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Datos de Cabecera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   13755
      Begin VB.ComboBox cmbIdPuntoDeCarga 
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
         Height          =   330
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   720
         Width           =   2805
      End
      Begin VB.ComboBox cmbFechaIngreso 
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
         Height          =   330
         Left            =   5010
         TabIndex        =   17
         Top             =   240
         Width           =   2550
      End
      Begin VB.TextBox txtIdOrden 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   16
         Top             =   270
         Width           =   1455
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   7470
         Picture         =   "aform1.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   1305
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         Height          =   315
         Left            =   2760
         TabIndex        =   14
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   1140
         Width           =   315
      End
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3870
         TabIndex        =   13
         Top             =   1530
         Width           =   3645
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9120
         TabIndex        =   12
         Top             =   1140
         Width           =   4515
      End
      Begin VB.TextBox txtNcuenta 
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1140
         Width           =   1245
      End
      Begin VB.TextBox txtNombrePaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3150
         TabIndex        =   10
         Top             =   1140
         Width           =   4365
      End
      Begin VB.TextBox txtNroOrdenSisSoat 
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
         Left            =   9120
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Frame fraTipoVenta 
         Height          =   525
         Left            =   1380
         TabIndex        =   6
         Top             =   540
         Width           =   6165
         Begin Threed.SSOption optVentas 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   180
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Venta Directa"
            Value           =   -1
         End
         Begin Threed.SSOption optPreventa 
            Height          =   255
            Left            =   4410
            TabIndex        =   8
            Top             =   180
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PreVenta"
         End
      End
      Begin VB.TextBox txtEstado 
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
         Left            =   9120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   270
         Width           =   1845
      End
      Begin MSMask.MaskEdBox txtHentrega 
         Height          =   315
         Left            =   2820
         TabIndex        =   19
         Top             =   2010
         Width           =   780
         _ExtentX        =   1376
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
      Begin MSMask.MaskEdBox txtFentrega 
         Height          =   315
         Left            =   1380
         TabIndex        =   20
         Top             =   2010
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
      Begin MSDataListLib.DataCombo cmbFormaPago 
         Height          =   360
         Left            =   1380
         TabIndex        =   21
         Top             =   1530
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Orden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   32
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblEstadoOrden 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11490
         TabIndex        =   31
         Top             =   510
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Registro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3630
         TabIndex        =   30
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   29
         Top             =   1575
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F.Despacho"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   28
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   27
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "N° Orden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8310
         TabIndex        =   26
         Top             =   1650
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Venta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   25
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Plan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8760
         TabIndex        =   24
         Top             =   1230
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8640
         TabIndex        =   23
         Top             =   780
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8490
         TabIndex        =   22
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   1
      Top             =   7650
      Width           =   13710
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "aform1.frx":3913
         DownPicture     =   "aform1.frx":3DD7
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
         Left            =   6870
         Picture         =   "aform1.frx":42C3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "aform1.frx":47AF
         DownPicture     =   "aform1.frx":4C0F
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
         Left            =   5340
         Picture         =   "aform1.frx":5084
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin Galenhos.ucFacturacionItems ucFacturacionProductos 
      Height          =   4965
      Left            =   30
      TabIndex        =   0
      Top             =   2580
      Width           =   13725
      _extentx        =   24209
      _extenty        =   8758
   End
End
Attribute VB_Name = "aForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ml_PuntoCarga As Long
Dim ml_idOrden As Long
Dim mi_Opcion As sghOpciones
Dim ms_MensajeError As String
Dim ml_IdUsuario As Long
Dim mb_ExistenDatos As Boolean

Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_Apariencia As New sighcomun.GridInfragistic
Dim mo_cmbIdPuntoCarga As New sighcomun.ListaDespleglable
Dim mo_cmbIdEstado As New sighcomun.ListaDespleglable
Dim mo_cmbFechaIngreso As New sighcomun.ListaDespleglable
Dim mo_cmbIdTipoGenHistoriaClinica As New sighcomun.ListaDespleglable
Dim mo_DOFactOrdenServicio As New DOFactOrdenServicio
Dim mo_DOAtencion As New DOAtencion
Dim ml_IdPaciente As Long
Dim ml_IdTipoFinanciamiento As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsFormaPago As New ADODB.Recordset
Dim lbDocumentoYaRegistradoEnSeguros As Boolean



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

Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
End Property

Property Get PuntoCarga() As Long
    PuntoCarga = ml_PuntoCarga
End Property

Property Let IdTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Get IdTipoFinanciamiento() As Long
    IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Property Let IdOrden(lValue As Long)
    ml_idOrden = lValue
End Property

Property Get IdOrden() As Long
    IdOrden = ml_idOrden
End Property


Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    If MsgBox("Los datos se agregaron correctamente" & Chr(13) & " " & Chr(13) & "        ¿ desea IMPRIMIR?       ", vbQuestion + vbYesNo, "Estado de Cuenta") = vbYes Then
                         Impresion
                    End If
                    
                    Me.txtIdOrden = mo_DOFactOrdenServicio.IdOrden
                    'MsgBox "Los datos se agregaron correctamente" + Chr(13) + "Se ha generado la orden número " & mo_DOFactOrdenServicio.IdOrden, vbInformation, Me.Caption
                    'LimpiarFormulario
                    Me.Visible = False
                Else
                    MsgBox "No se pudo agregar los datos", vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    If MsgBox("Desea REIMPRIMIR ??", vbQuestion + vbYesNo, "Estado de Cuenta") = vbYes Then
                       Impresion
                    End If
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos", vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            If MsgBox("¿Realmente desea Eliminar?", vbQuestion + vbYesNo, "Estado de Cuenta") = vbNo Then
                 Exit Sub
            End If
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos"
               End If
           End If
   End Select
        
End Sub

Sub LimpiarFormulario()
    
    Me.ucFacturacionProductos.LimpiarGrilla

End Sub

Function ValidarDatosObligatorios() As Boolean
    ValidarDatosObligatorios = False
    Select Case mi_Opcion
    Case sghAgregar, sghModificar
        Dim rsProductos As Recordset
        Set rsProductos = Me.ucFacturacionProductos.FacturacionProductos
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                If rsProductos!IdProducto = 0 Then
                    MsgBox "Uno de los productos tiene datos imcompletos, por favor verifique", vbInformation, Me.Caption
                    Exit Function
                End If
                rsProductos.MoveNext
            Loop
        End If
        If Me.ucFacturacionProductos.DevuelveTotalPagar <= 0 Then
           MsgBox "El Importe Total es 0.....verifique", vbInformation, Me.Caption
           Exit Function
        End If
    End Select
    ValidarDatosObligatorios = True
End Function

Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With mo_DOFactOrdenServicio
             .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL      'Now
             .FechaModificacion = 0
             If txtHentrega.Text <> "__:__" Then
                .FechaOrden = CDate(txtFentrega.Text & " " & txtHentrega.Text) 'Now
             End If
             .idAtencion = mo_DOAtencion.idAtencion
             .IdPuntoCarga = mo_cmbIdPuntoCarga.BoundText
             .IdUsuarioAuditoria = ml_IdUsuario
             .IdUsuarioCrea = ml_IdUsuario
             .IdUsuarioModifica = 0
             .IdComprobantePago = 0
             .IdEstadoOrden = 1  '1 Registrado 4 Pagado 9 Anulado
             .IdFormaPago = cmbFormaPago.BoundText
        End With
    Case sghModificar
        mo_DOFactOrdenServicio.IdFormaPago = cmbFormaPago.BoundText
        If txtHentrega.Text <> "__:__" Then
           mo_DOFactOrdenServicio.FechaOrden = CDate(txtFentrega.Text & " " & txtHentrega.Text) 'Now
        End If
    End Select
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
    

    
   ValidarReglas = True
End Function
Function AgregarDatos() As Boolean

    AgregarDatos = mo_ReglasFacturacion.FactOrdenServicioAgregar(mo_DOFactOrdenServicio, Me.ucFacturacionProductos.FacturacionProductos, Me.ucFacturacionProductos.ProductosEliminados, ml_IdUsuario)
   
End Function

Function ModificarDatos() As Boolean
    
    ModificarDatos = mo_ReglasFacturacion.FactOrdenServicioModificar(mo_DOFactOrdenServicio, Me.ucFacturacionProductos.FacturacionProductos, Me.ucFacturacionProductos.ProductosEliminados, ml_IdUsuario)

End Function

Function EliminarDatos() As Boolean
    
    EliminarDatos = mo_ReglasFacturacion.FactOrdenServicioEliminar(mo_DOFactOrdenServicio)

End Function

Private Sub btnBuscarPaciente_Click()
    
    If txtNroCuenta.Text <> "" Then
        CargaPorCuentaAtencion
        Exit Sub
    End If
    If txtNroOrdenSisSoat.Text <> "" Then
       CargaOrdenYaRegistradaEnSisSoat
       Exit Sub
    End If


    Dim rsHistorias As New Recordset
    Dim oDOPaciente As New doPaciente
    
    oDOPaciente.NroHistoriaClinica = Val(Me.txtNroHistoriaBusqueda.Text)
    oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda
    oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
    oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
    oDOPaciente.SegundoNombre = Me.txtSegundoNombreBusqueda
    oDOPaciente.IdDocIdentidad = 1
    oDOPaciente.nroDocumento = ""
    
    If (oDOPaciente.ApellidoPaterno + oDOPaciente.ApellidoMaterno + _
    oDOPaciente.PrimerNombre + oDOPaciente.SegundoNombre = "") And _
    (Val(Me.txtNroHistoriaBusqueda.Text) = 0) And _
    (oDOPaciente.nroDocumento = "") Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Set rsHistorias = mo_AdminAdmision.PacientesFiltrar(oDOPaciente)
    Screen.MousePointer = vbDefault
    
    Set grdPacientesEncontrados.DataSource = rsHistorias
        
    'Si hay una sola coincidencia
    If rsHistorias.RecordCount = 1 Then
        
        ml_IdPaciente = rsHistorias!idPaciente
        txtPaciente = rsHistorias!ApellidoPaterno + " " + rsHistorias!ApellidoMaterno + " " + rsHistorias!PrimerNombre
       
        Dim rs As New Recordset
        Set rs = mo_ReglasFacturacion.CuentaAtencionSeleccionarUltimaPorIdPaciente(ml_IdPaciente)
        If rs.RecordCount = 1 Then
            txtNroCuenta.Text = rs!idCuentaAtencion
            Set mo_DOAtencion = mo_ReglasFacturacion.SeleccionarUltimaAtencion(ml_IdPaciente, rs!idCuentaAtencion)
            Set Me.ucFacturacionProductos.Atencion = mo_DOAtencion
            Set mo_cmbFechaIngreso.RowSource = rs
            cmbFechaIngreso.ListIndex = 0
            
            Me.ucFacturacionProductos.LimpiarGrilla
            Me.ucFacturacionProductos.CargaProductosPorIdOrden
        Else
            MsgBox "El paciente no tiene ninguna cuenta en estado Abierto", vbInformation, Me.Caption
            Exit Sub
        End If
        
    ElseIf rsHistorias.RecordCount > 1 Then
        Me.grdPacientesEncontrados.Visible = True
    ElseIf rsHistorias.RecordCount = 0 Then
        
        MsgBox "No se encontró datos", vbInformation, Me.Caption
        Me.grdPacientesEncontrados.Visible = False
        txtNroHistoriaBusqueda.Text = ""
        txtApellidoMaternoBusqueda = ""
        txtPrimerNombreBusqueda = ""
        txtSegundoNombreBusqueda = ""
        txtApellidoPaternoBusqueda = ""
        txtNroCuenta.Text = ""
        
    End If
    
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPacientesEncontrados, sighcomun.GrillaConFilasBicolor
    Screen.MousePointer = vbDefault

End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub Form_Activate()
    If txtNroCuenta.Text <> "" And mi_Opcion = sghAgregar Then
        'Es una cuenta que se AGREGA desde SEGUROS (sis/soat)
        btnBuscarPaciente_Click
        txtHentrega.Text = "__:__"
        txtFentrega.Text = "__/__/____"
        txtHentrega.Enabled = False
        txtFentrega.Enabled = False
        cmbFormaPago.Enabled = False
        btnBuscarPaciente.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    lbDocumentoYaRegistradoEnSeguros = False
    
    Set mo_cmbFechaIngreso.MiComboBox = cmbFechaIngreso
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    
    ConfigurarPuntosDeCarga
    ConfigurarTiposHistoriaClinica
    ConfigurarFechaIngreso
    CargaDataCombos
    
    mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga
    mo_cmbIdTipoGenHistoriaClinica.BoundText = lcBuscaParametro.SeleccionaFilaParametro(212)

    Me.ucFacturacionProductos.IdUsuario = ml_IdUsuario
    Me.ucFacturacionProductos.Inicializar
    Me.ucFacturacionProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    Me.ucFacturacionProductos.TipoProducto = sghServicio
    Me.ucFacturacionProductos.IdPuntoCarga = ml_PuntoCarga

    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Ordenes de Servicio"
    Case sghModificar
        Me.Caption = "Modificar Ordenes de Servicio"
    Case sghConsultar
        Me.Caption = "Consultar Ordenes de Servicio"
    Case sghEliminar
        Me.Caption = "Eliminar Ordenes de Servicio"
    End Select
    
    CargarDatosAlFormulario
End Sub

Sub CargarDatosAlFormulario()

    Me.grdPacientesEncontrados.Left = 210
    Me.grdPacientesEncontrados.Top = 810

 Select Case mi_Opcion
     Case sghAgregar
        txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL   'Format(Now, "HH:MM")
        txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL    'Format(Date, "DD/MM/YYYY")
     Case sghModificar
        Me.fraBusqueda.Enabled = False
        CargarDatosALosControles
     Case sghConsultar
        Me.fraBusqueda.Enabled = False
        CargarDatosALosControles
     Case sghEliminar
        Me.fraBusqueda.Enabled = False
        CargarDatosALosControles
 End Select
End Sub

Sub CargarDatosALosControles()

        'Carga datos de la orden
        Set mo_DOFactOrdenServicio = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorId(Me.IdOrden)
        
        If Not mo_DOFactOrdenServicio Is Nothing Then
             With mo_DOFactOrdenServicio
                mo_cmbIdPuntoCarga.BoundText = mo_DOFactOrdenServicio.IdPuntoCarga
                 Me.txtIdOrden = Me.IdOrden
                 mb_ExistenDatos = True
                 
                 Select Case mo_DOFactOrdenServicio.IdEstadoOrden
                 Case 1
                    lblEstadoOrden = "Estado: INGRESADO"
                 Case 4
                    lblEstadoOrden = "Estado: PAGADO"
                    lblEstadoOrden.ForeColor = &HB82C2F
                    If mi_Opcion = sghEliminar Then
                        MsgBox "Las ordenes PAGADAS no se pueden eliminar", vbInformation, "Facturación"
                        Me.btnAceptar.Enabled = False 'Las ordenes pagadas no se pueden eliminar
                    End If
                 Case 9
                    lblEstadoOrden = "Estado: ANULADO"
                    lblEstadoOrden.ForeColor = &H3448FE
                    MsgBox "Las ordenes ANULADAS no se pueden modificar ni eliminar", vbInformation, "Facturación"
                    Me.btnAceptar.Enabled = False   'Las ordenes anuladas no se pueden modificar ni eliminar
                 End Select
                 txtHentrega.Text = Format(mo_DOFactOrdenServicio.FechaOrden, "HH:MM")
                 txtFentrega.Text = Format(mo_DOFactOrdenServicio.FechaOrden, "DD/MM/YYYY")
                 cmbFormaPago.BoundText = mo_DOFactOrdenServicio.IdFormaPago
             End With
         Else
            mb_ExistenDatos = False
            Exit Sub
         End If
         
        'Cargar datos del paciente y de la atencion
        Set mo_DOAtencion = mo_AdminAdmision.AtencionesSeleccionarPorId(mo_DOFactOrdenServicio.idAtencion)
        Set Me.ucFacturacionProductos.Atencion = mo_DOAtencion
        cmbFechaIngreso.Text = mo_DOAtencion.FechaIngreso
        txtNroCuenta.Text = mo_DOAtencion.idCuentaAtencion

        
        Dim oDOPaciente As New doPaciente
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DOAtencion.idPaciente)
        If Not oDOPaciente Is Nothing Then
            txtPaciente.Text = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre
            mo_cmbIdTipoGenHistoriaClinica.BoundText = oDOPaciente.IdTipoNumeracion
            Me.txtNroHistoriaBusqueda.Text = oDOPaciente.NroHistoriaClinica
        End If
         
        'Cargar datos de los servicios
        Me.ucFacturacionProductos.LimpiarGrilla
        Me.ucFacturacionProductos.IdOrden = Me.IdOrden
        Me.ucFacturacionProductos.IdEstadoOrden = mo_DOFactOrdenServicio.IdEstadoOrden
        Me.ucFacturacionProductos.CargaProductosPorIdOrden
   
        txtHentrega.Enabled = False
        txtFentrega.Enabled = False
   
        Select Case mi_Opcion
        Case sghModificar
        Case sghEliminar
        Case sghConsultar
        End Select
   
   
End Sub


'Private Sub txtNroHistoria_LostFocus()
'Dim oPaciente As New doPaciente
'Dim rsRespuesta As New Recordset
'
'    If txtNroHistoria = "" Then
'        Exit Sub
'    End If
'
'    oPaciente.NroHistoriaClinica = Val(txtNroHistoria)
'    oPaciente.IdTipoNumeracion = Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
'    Set rsRespuesta = mo_AdminAdmision.PacientesFiltrar(oPaciente)
'
'    If rsRespuesta.RecordCount = 0 Then
'        MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
'    ElseIf rsRespuesta.RecordCount = 1 Then
'        ml_IdPaciente = rsRespuesta!IdPaciente
'        txtPaciente = rsRespuesta!ApellidoPaterno + " " + rsRespuesta!ApellidoMaterno + " " + rsRespuesta!PrimerNombre
'
'        Dim rs As New Recordset
'        Set rs = mo_ReglasFacturacion.CuentaAtencionSeleccionarUltimaPorIdPaciente(ml_IdPaciente)
'        If rs.RecordCount = 1 Then
'
'            Set mo_DOAtencion = mo_ReglasFacturacion.SeleccionarUltimaAtencion(ml_IdPaciente, rs!IdCuentaAtencion)
'            Set Me.ucFacturacionProductos.Atencion = mo_DOAtencion
'            Set mo_cmbFechaIngreso.RowSource = rs
'            cmbFechaIngreso.ListIndex = 0
'
'            Me.ucFacturacionProductos.LimpiarGrilla
'            Me.ucFacturacionProductos.CargaProductosPorIdOrden
'
'
'        End If
'    End If
'
'End Sub

Sub ConfigurarFechaIngreso()
    
    mo_cmbFechaIngreso.ListField = "DescripcionLarga"
    mo_cmbFechaIngreso.BoundColumn = "IdCuentaAtencion"

End Sub

Sub ConfigurarPuntosDeCarga()
    
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    
    Set mo_cmbIdPuntoCarga.RowSource = mo_ReglasComunes.SeleccionarPuntosDeCarga()

End Sub

Sub ConfigurarTiposHistoriaClinica()
        
        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

End Sub







Sub CargaPorCuentaAtencion()
    Dim rs As New Recordset
    Dim rsRespuesta As New Recordset
    Dim oPaciente As New doPaciente
    Dim oCuentaAtencion As New DOCuentaAtencion
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    Me.MousePointer = 11
    Set oCuentaAtencion = mo_ReglasFacturacion.CuentasAtencionSeleccionarPorId(Val(txtNroCuenta.Text))
    If oCuentaAtencion.idPaciente = 0 Then
       MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
    Else
       Dim oBuscaCabecera As New sighfacturacion.dllFactFacBienServicio
       Set rs = oBuscaCabecera.RetornaDatosDeCabecera(oCuentaAtencion.idCuentaAtencion)
'       oConexion.CommandTimeout = 300
'       oConexion.Open SIGHComun.CadenaConexion
'       lcSql = "Select ca.IdCuentaAtencion, ca.FechaApertura,ca.HoraApertura, ca.IdEstado," & _
'                 " ca.FechaApertura + ' ' + ca.HoraApertura as DescripcionLarga " & _
'                 " from FacturacionCuentasAtencion ca, Pacientes pa " & _
'                 " where ca.IdPaciente = pa.IdPaciente " & _
'                 " and ca.IdEstado = 1 " & _
'                 " and  ca.IdCuentaAtencion=" & oCuentaAtencion.idCuentaAtencion
'       rs.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If rs.RecordCount > 0 Then
            Set oPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oCuentaAtencion.idPaciente)
            txtPaciente = oPaciente.ApellidoPaterno + " " + oPaciente.ApellidoMaterno + " " + oPaciente.PrimerNombre
             ml_IdPaciente = oPaciente.idPaciente
             txtNroHistoriaBusqueda.Text = oPaciente.NroHistoriaClinica
            Set mo_DOAtencion = mo_ReglasFacturacion.SeleccionarUltimaAtencion(ml_IdPaciente, rs!idCuentaAtencion)
           
            Set Me.ucFacturacionProductos.Atencion = mo_DOAtencion
            Set mo_cmbFechaIngreso.RowSource = rs
            cmbFechaIngreso.ListIndex = 0
            Me.ucFacturacionProductos.LimpiarGrilla
            Me.ucFacturacionProductos.CargaProductosPorIdOrden
            If mo_DOAtencion.IdTipoServicio <> 1 Then  'Hospitalizacion y Emergencia
               cmbFormaPago.BoundText = "5"
               cmbFormaPago.Enabled = False
            End If
       Else
            MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
       End If
    End If
    Me.MousePointer = 1
End Sub


Sub Impresion()
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim iFila As Integer: Dim lnTotal As Double
    Dim rsReporte As New Recordset
    Dim mo_ReporteUtil As New ReporteUtil
    
        MousePointer = 11
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\GalenHos.xls")
        oWorkBookPlantilla.Worksheets("facturacion_bs").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
        '*******************************************Inicio de Impresion
        oWorkSheet.Cells(1, 2).Value = "Nº de Orden: " & mo_DOFactOrdenServicio.IdOrden
        oWorkSheet.Cells(3, 3).Value = txtPaciente.Text
        oWorkSheet.Cells(3, 6).Value = txtNroCuenta.Text
        oWorkSheet.Cells(4, 3).Value = cmbIdPuntoDeCarga.Text
        oWorkSheet.Cells(4, 6).Value = txtNroHistoriaBusqueda.Text
        oWorkSheet.Cells(5, 3).Value = cmbFormaPago.Text
        If txtHentrega.Text <> "__:__" Then
           oWorkSheet.Cells(5, 6).Value = CDate(txtFentrega.Text & " " & txtHentrega.Text)
        End If
        oWorkSheet.Cells(6, 2).Value = ""
        oWorkSheet.Cells(6, 3).Value = ""
        
        Set rsReporte = Me.ucFacturacionProductos.FacturacionProductos
        iFila = 9: lnTotal = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
           oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!codigo
           oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields!NombreProducto
           oWorkSheet.Cells(iFila, 5).Value = Format(rsReporte.Fields!cantidad, "####,###")
           oWorkSheet.Cells(iFila, 6).Value = Format(rsReporte.Fields!PrecioUnitario, "####,##0.000")
           oWorkSheet.Cells(iFila, 7).Value = Format(rsReporte.Fields!totalPorPagar, "####,##0.00")
           lnTotal = lnTotal + rsReporte.Fields!totalPorPagar
           iFila = iFila + 1
           rsReporte.MoveNext
        Loop
        iFila = iFila + 1
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
        oWorkSheet.Cells(iFila, 2).Value = "Total: "
        oWorkSheet.Cells(iFila, 7).Value = Format(lnTotal, "####,##0.00")
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
        'oWorkBook.Close SaveChanges:=False
        MousePointer = 1
End Sub

Sub CargaDataCombos()
      On Error GoTo error_pcargadrs
      Dim oConexion As New ADODB.Connection
      Dim lcSql As String
      
      oConexion.Open sighcomun.CadenaConexion
      oConexion.CursorLocation = adUseClient
      
      oRsFormaPago.Open "select * from TiposFinanciamiento where esFuenteFinanciamiento=1 order by descripcion", oConexion, adOpenKeyset, adLockOptimistic
      Set cmbFormaPago.RowSource = oRsFormaPago
      cmbFormaPago.ListField = "Descripcion"
      cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
      cmbFormaPago.BoundText = 1
      
      'oConexion.Close
      'Set oConexion = Nothing
      Exit Sub
error_pcargadrs:
      If Err.Number = 3705 Then
         oRsFormaPago.Close
         Resume
      Else
         MsgBox Err.Number & " Descripcion: " & Err.Description, vbCritical, "ERROR"
      End If

End Sub

Sub CargaOrdenYaRegistradaEnSisSoat()
    If mi_Opcion = sghAgregar Then
        Dim LnIdFormaPago As Long
        mi_Opcion = sghModificar
        Me.IdOrden = txtNroOrdenSisSoat.Text
        CargarDatosAlFormulario
        LnIdFormaPago = Me.ucFacturacionProductos.OrdenRegistradaYaprobadaPorSisSoat
        If esfecha(Left(mo_DOFactOrdenServicio.FechaOrden, 10), "DD/MM/AAAA") = False Then
        
           If LnIdFormaPago > 1 Then
                cmbFormaPago.BoundText = LnIdFormaPago  '3=sis, 4=soat
                cmbFormaPago.Enabled = False
                Frame3.Enabled = True
                txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL 'Format(Now, "HH:MM")
                txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL  'Format(Date, "DD/MM/YYYY")
                txtHentrega.Enabled = True
                txtFentrega.Enabled = True
                lbDocumentoYaRegistradoEnSeguros = True
                'Cargar datos de los servicios
                Me.ucFacturacionProductos.LimpiarGrilla
                Me.ucFacturacionProductos.DocumentoYaRegistradoEnSeguros = lbDocumentoYaRegistradoEnSeguros
                Me.ucFacturacionProductos.IdOrden = Me.IdOrden
                Me.ucFacturacionProductos.IdEstadoOrden = mo_DOFactOrdenServicio.IdEstadoOrden
                Me.ucFacturacionProductos.CargaProductosPorIdOrden
            Else
                MsgBox "Ese 'Nro de Orden' NO ha sido Aprobado en SIS o SOAT", vbCritical, "Mensaje"
                btnCancelar_Click
            End If
        Else
            MsgBox "Ese 'Nro de Orden' ha sido Registrado y Aprobado en SIS o SOAT, pero ya fue despachado", vbCritical, "Mensaje"
            btnCancelar_Click
        End If
   End If
End Sub






