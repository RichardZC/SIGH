VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form FacOrdenBienInsumoDetalle 
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   Icon            =   "FacOrdenBienInsumoDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin UltraGrid.SSUltraGrid grdPacientesEncontrados 
      Height          =   2265
      Left            =   4140
      TabIndex        =   13
      Top             =   1050
      Visible         =   0   'False
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   3995
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
      Caption         =   "SSUltraGrid1"
   End
   Begin VB.Frame fraDatosHistoria 
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
      Height          =   1905
      Left            =   60
      TabIndex        =   14
      Top             =   1170
      Width           =   13725
      Begin VB.Frame frmFarmacia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   90
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   10545
         Begin MSDataListLib.DataCombo cmbFormaPago 
            Height          =   360
            Left            =   7230
            TabIndex        =   23
            Top             =   300
            Width           =   3045
            _ExtentX        =   5371
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
         Begin MSDataListLib.DataCombo cmbFarmacia 
            Height          =   360
            Left            =   1020
            TabIndex        =   24
            Top             =   300
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Farmacia"
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
            Left            =   150
            TabIndex        =   26
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago"
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
            Left            =   5880
            TabIndex        =   25
            Top             =   360
            Width           =   1305
         End
      End
      Begin VB.TextBox txtIdOrden 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4230
         MaxLength       =   30
         TabIndex        =   18
         Top             =   510
         Width           =   1365
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
         Left            =   5700
         TabIndex        =   17
         Top             =   540
         Width           =   2610
      End
      Begin VB.TextBox txtPaciente 
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
         Height          =   360
         Left            =   90
         MaxLength       =   30
         TabIndex        =   16
         Top             =   510
         Width           =   4125
      End
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
         Left            =   8430
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   540
         Width           =   2805
      End
      Begin MSMask.MaskEdBox txtHentrega 
         Height          =   315
         Left            =   12750
         TabIndex        =   27
         Top             =   540
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
         Left            =   11370
         TabIndex        =   28
         Top             =   540
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   $"FacOrdenBienInsumoDetalle.frx":000C
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
         Left            =   120
         TabIndex        =   19
         Top             =   210
         Width           =   13290
      End
   End
   Begin Galenhos.ucFacturacionItems ucFacturacionProductos 
      Height          =   4335
      Left            =   90
      TabIndex        =   5
      Top             =   3120
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   7805
   End
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   13725
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
         Height          =   315
         Left            =   9180
         TabIndex        =   30
         Top             =   630
         Width           =   1545
      End
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
         Left            =   10740
         TabIndex        =   12
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtApellidoMaternoBusqueda 
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
         Left            =   6645
         TabIndex        =   11
         Top             =   630
         Width           =   1260
      End
      Begin VB.TextBox txtApellidoPaternoBusqueda 
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
         Left            =   5460
         TabIndex        =   10
         Top             =   630
         Width           =   1185
      End
      Begin VB.TextBox txtPrimerNombreBusqueda 
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
         Left            =   7905
         TabIndex        =   9
         Top             =   630
         Width           =   1230
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   12360
         Picture         =   "FacOrdenBienInsumoDetalle.frx":00C2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   630
         Width           =   1305
      End
      Begin VB.TextBox txtNroHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         MaxLength       =   9
         TabIndex        =   4
         Top             =   630
         Width           =   1365
      End
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1845
      End
      Begin Threed.SSOption optConHistoriaClinica 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Con Historia Clínica"
         Value           =   -1
      End
      Begin Threed.SSOption optSinHistoriaClinica 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sin Historia Clínica"
      End
      Begin VB.Label Label50 
         Caption         =   $"FacOrdenBienInsumoDetalle.frx":2D0B
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
         Left            =   2370
         TabIndex        =   8
         Top             =   300
         Width           =   9855
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
         Left            =   7380
         TabIndex        =   6
         Top             =   1290
         Width           =   2205
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   90
      TabIndex        =   0
      Top             =   7590
      Width           =   13680
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacOrdenBienInsumoDetalle.frx":2D97
         DownPicture     =   "FacOrdenBienInsumoDetalle.frx":31F7
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
         Picture         =   "FacOrdenBienInsumoDetalle.frx":366C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacOrdenBienInsumoDetalle.frx":3AE1
         DownPicture     =   "FacOrdenBienInsumoDetalle.frx":3FA5
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
         Picture         =   "FacOrdenBienInsumoDetalle.frx":4491
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FacOrdenBienInsumoDetalle"
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
Dim mo_Formulario As New SIGHComun.Formulario
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico

Dim mo_cmbIdPuntoCarga As New SIGHComun.ListaDespleglable
Dim mo_cmbIdEstado As New SIGHComun.ListaDespleglable
Dim mo_cmbFechaIngreso As New SIGHComun.ListaDespleglable
Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHComun.ListaDespleglable
Dim mo_DOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
Dim mo_DOAtencion As New DOAtencion
Dim ml_IdPaciente As Long
Dim ml_IdTipoFinanciamiento As Long
Dim oRsFormaPago As New ADODB.Recordset
Dim oRsFarmacias As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnCodigoFarmacia  As Long
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
                    If MsgBox("Los datos se agregaron correctamente" & Chr(13) & "  " & Chr(13) & "      ¿desea IMPRIMIR?", vbQuestion + vbYesNo, "Estado de Cuenta") = vbYes Then
                         Impresion
                    End If
                    
                    Me.txtIdOrden = mo_DOFactOrdenBienInsumo.IdOrden
                    MsgBox "Los datos se agregaron correctamente" + Chr(13) + "Se ha generado la orden número " & mo_DOFactOrdenBienInsumo.IdOrden, vbInformation, Me.Caption
                    'LimpiarFormulario
                    'txtNroCuenta.SetFocus
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
    txtPaciente.Text = ""
    txtIdOrden.Text = ""
    cmbFechaIngreso.Text = ""
    txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL   'Format(Now, "HH:MM")
    txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL    'Format(Date, "DD/MM/YYYY")
    ml_IdPaciente = 0
    Me.ucFacturacionProductos.LimpiarGrilla

End Sub

Function ValidarDatosObligatorios() As Boolean
    ValidarDatosObligatorios = False
    If Round(Me.ucFacturacionProductos.DevuelveTotalPagar, 2) <= 0 Then
       MsgBox "El Importe Total es 0.....verifique", vbInformation, Me.Caption
       Exit Function
    End If
    ValidarDatosObligatorios = True
End Function

Sub CargaDatosAlObjetosDeDatos()

   With mo_DOFactOrdenBienInsumo
   Select Case mi_Opcion
   Case sghAgregar
        .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL     'Now
        .FechaModificacion = 0
        If txtHentrega.Text <> "__:__" Then
           .FechaOrden = CDate(txtFentrega.Text & " " & txtHentrega.Text) 'Now
        End If
        .idAtencion = IIf(optConHistoriaClinica.Value, mo_DOAtencion.idAtencion, 0)
        .IdPuntoCarga = mo_cmbIdPuntoCarga.BoundText
        .IdUsuarioAuditoria = ml_IdUsuario
        .IdUsuarioCrea = ml_IdUsuario
        .IdUsuarioModifica = 0
        .IdComprobantePago = 0
        .IdEstadoOrden = 1 '1 Registrado 4 Pagado 9 Anulado
        .idPaciente = ml_IdPaciente
        If frmFarmacia.Visible = True Then
            .IdFormaPago = cmbFormaPago.BoundText
            .idFarmacia = cmbFarmacia.BoundText
        End If
    Case sghModificar
        If frmFarmacia.Visible = True Then
            .IdFormaPago = cmbFormaPago.BoundText
            .idFarmacia = cmbFarmacia.BoundText
            If txtHentrega.Text <> "__:__" Then
               .FechaOrden = CDate(txtFentrega.Text & " " & txtHentrega.Text) 'Now
            End If
        End If
    End Select
   End With
   
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
      
   ValidarReglas = True
End Function
Function AgregarDatos() As Boolean

    AgregarDatos = mo_ReglasFacturacion.FactOrdenBienInsumoAgregar(mo_DOFactOrdenBienInsumo, Me.ucFacturacionProductos.FacturacionProductos, Me.ucFacturacionProductos.ProductosEliminados, ml_IdUsuario)
   
End Function

Function ModificarDatos() As Boolean
    
    ModificarDatos = mo_ReglasFacturacion.FactOrdenBienInsumoModificar(mo_DOFactOrdenBienInsumo, Me.ucFacturacionProductos.FacturacionProductos, Me.ucFacturacionProductos.ProductosEliminados, ml_IdUsuario)

End Function

Function EliminarDatos() As Boolean
    
    EliminarDatos = mo_ReglasFacturacion.FactOrdenBienInsumoEliminar(mo_DOFactOrdenBienInsumo)

End Function

Private Sub btnBuscarPaciente_Click()
    Dim rsHistorias As New Recordset
    Dim oDOPaciente As New doPaciente
    If txtNroOrdenSisSoat.Text <> "" Then
        CargaOrdenYaRegistradaEnSisSoat
    ElseIf txtNroHistoria.Text <> "" Then
        txtNroHistoria_LostFocus
    ElseIf txtNroCuenta.Text <> "" Then
        CargaPorCuentaAtencion
    Else
        oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda
        oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
        oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
        oDOPaciente.IdDocIdentidad = 1
        oDOPaciente.nroDocumento = ""
        
        If Me.txtApellidoPaternoBusqueda.Text = "" And Me.txtApellidoMaternoBusqueda.Text = "" And _
        Me.txtPrimerNombreBusqueda.Text = "" And _
        Me.txtNroCuenta.Text = "" Then
            MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        Set rsHistorias = mo_AdminAdmision.PacientesFiltrar(oDOPaciente)
        Screen.MousePointer = vbDefault
        
        Set grdPacientesEncontrados.DataSource = rsHistorias
            
        'Si hay una sola coincidencia
        If rsHistorias.RecordCount = 1 Then
            txtNroHistoria.Text = rsHistorias.Fields!NroHistoriaClinica
            mo_cmbIdTipoGenHistoriaClinica.BoundText = rsHistorias.Fields!IdTipoNumeracion
            txtNroHistoria_LostFocus
        ElseIf rsHistorias.RecordCount > 1 Then
            Me.grdPacientesEncontrados.Visible = True
        ElseIf rsHistorias.RecordCount = 0 Then
            
            MsgBox "No se encontró datos", vbInformation, Me.Caption
            Me.grdPacientesEncontrados.Visible = False
            txtApellidoMaternoBusqueda = ""
            txtPrimerNombreBusqueda = ""
            txtApellidoPaternoBusqueda = ""
            txtNroCuenta.Text = ""
            
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores Me.grdPacientesEncontrados, SIGHComun.GrillaConFilasBicolor
        Screen.MousePointer = vbDefault
    End If




      'txtNroHistoria_LostFocus
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub Form_Activate()
    If txtNroCuenta.Text <> "" And mi_Opcion = sghAgregar Then
        'Es una cuenta que se AGREGA desde SEGUROS (sis/soat/convenio/exo)
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

    Me.optConHistoriaClinica.Value = True

    Me.ucFacturacionProductos.IdUsuario = ml_IdUsuario
    Me.ucFacturacionProductos.Inicializar
    Me.ucFacturacionProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    Me.ucFacturacionProductos.TipoProducto = sghbien
    Me.ucFacturacionProductos.IdPuntoCarga = ml_PuntoCarga

    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Ordenes de Bien e Insumos"
    Case sghModificar
        Me.Caption = "Modificar Ordenes de Bien e Insumos"
    Case sghConsultar
        Me.Caption = "Consultar Ordenes de Bien e Insumos"
    Case sghEliminar
        Me.Caption = "Eliminar Ordenes de Bien e Insumos"
    End Select
    
    CargarDatosAlFormulario
    
End Sub

Sub CargaDataCombos()
      On Error GoTo error_pcargadrs
      Dim oConexion As New ADODB.Connection
      Dim lcSql As String
      
      oConexion.Open SIGHComun.CadenaConexion
      oConexion.CursorLocation = adUseClient
      
      lcSql = "select distinct RolesPermisos.IdPermiso,permisos.modulo from (((Empleados" & _
            " left join UsuariosRoles on Empleados.IdEmpleado = UsuariosRoles.IdEmpleado)" & _
            " left join Roles on UsuariosRoles.IdRol = Roles.IdRol)" & _
            " left join RolesPermisos on Roles.IdRol = RolesPermisos.IdRol)" & _
            " left join Permisos on RolesPermisos.IdPermiso = Permisos.IdPermiso" & _
            " where Empleados.Idempleado = " & ml_IdUsuario & " and permisos.modulo='Farmacia'"
      oRsFarmacias.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
      If oRsFarmacias.RecordCount > 0 Then
         lnCodigoFarmacia = oRsFarmacias.Fields!idPermiso
         oRsFarmacias.Close
         frmFarmacia.Visible = True
         oRsFarmacias.Open "select * from permisos where modulo='Farmacia' order by idPermiso", oConexion, adOpenKeyset, adLockOptimistic
         Set cmbFarmacia.RowSource = oRsFarmacias
         cmbFarmacia.ListField = "descripcion"
         cmbFarmacia.BoundColumn = "idPermiso"
         cmbFarmacia.BoundText = lnCodigoFarmacia
      Else
         frmFarmacia.Visible = False
         oRsFarmacias.Close
         optSinHistoriaClinica.Visible = False
      End If
      
      oRsFormaPago.Open "select * from TiposFinanciamiento where esFarmacia=1 order by descripcion", oConexion, adOpenKeyset, adLockOptimistic
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




Sub CargarDatosAlFormulario()
 Select Case mi_Opcion
     Case sghAgregar
        txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL     'Format(Now, "HH:MM")
        txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL    'Format(Date, "DD/MM/YYYY")
     Case sghModificar
        Me.fraDatosHistoria.Enabled = False
        fraDatosAtencion.Enabled = False
        CargarDatosALosControles
     Case sghConsultar
        Me.fraDatosHistoria.Enabled = False
        fraDatosAtencion.Enabled = False
        CargarDatosALosControles
     Case sghEliminar
        Me.fraDatosHistoria.Enabled = False
        fraDatosAtencion.Enabled = False
        CargarDatosALosControles
 End Select
End Sub

Sub CargarDatosALosControles()

        'Carga datos de la orden
        Set mo_DOFactOrdenBienInsumo = mo_ReglasFacturacion.FactOrdenBienInsumoSeleccionarPorId(Me.IdOrden)
        
        If Not mo_DOFactOrdenBienInsumo Is Nothing Then
             With mo_DOFactOrdenBienInsumo
                mo_cmbIdPuntoCarga.BoundText = mo_DOFactOrdenBienInsumo.IdPuntoCarga
                 Me.txtIdOrden = Me.IdOrden
                 mb_ExistenDatos = True
             End With
             
            Select Case mo_DOFactOrdenBienInsumo.IdEstadoOrden
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
            If Not IsNull(mo_DOFactOrdenBienInsumo.IdFormaPago) Then
                cmbFormaPago.BoundText = mo_DOFactOrdenBienInsumo.IdFormaPago
                cmbFarmacia.BoundText = mo_DOFactOrdenBienInsumo.idFarmacia
            End If
            optSinHistoriaClinica.Value = IIf(mo_DOFactOrdenBienInsumo.idPaciente = 0, True, False)
            optConHistoriaClinica.Value = IIf(mo_DOFactOrdenBienInsumo.idPaciente = 0, False, True)
            txtHentrega.Text = Format(mo_DOFactOrdenBienInsumo.FechaOrden, "HH:MM")
            txtFentrega.Text = Format(mo_DOFactOrdenBienInsumo.FechaOrden, "DD/MM/YYYY")
         Else
            mb_ExistenDatos = False
            Exit Sub
         End If
         
        'Cargar datos del paciente y de la atencion
        Set mo_DOAtencion = mo_AdminAdmision.AtencionesSeleccionarPorId(mo_DOFactOrdenBienInsumo.idAtencion)
        Set Me.ucFacturacionProductos.Atencion = mo_DOAtencion
        cmbFechaIngreso.Text = mo_DOAtencion.FechaIngreso
        txtNroCuenta.Text = mo_DOAtencion.idCuentaAtencion
        
        Dim oDOPaciente As New doPaciente
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DOAtencion.idPaciente)
        If Not oDOPaciente Is Nothing Then
            txtPaciente.Text = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre
            mo_cmbIdTipoGenHistoriaClinica.BoundText = oDOPaciente.IdTipoNumeracion
            Me.txtNroHistoria.Text = oDOPaciente.NroHistoriaClinica
        End If
         
        'Cargar datos de los servicios
        Me.ucFacturacionProductos.LimpiarGrilla
        'Me.ucFacturacionProductos.DocumentoYaRegistradoEnSeguros = lbDocumentoYaRegistradoEnSeguros
        Me.ucFacturacionProductos.IdOrden = Me.IdOrden
        Me.ucFacturacionProductos.IdEstadoOrden = mo_DOFactOrdenBienInsumo.IdEstadoOrden
        Me.ucFacturacionProductos.CargaProductosPorIdOrden
   
   
        Select Case mi_Opcion
        Case sghModificar
        Case sghEliminar
        Case sghConsultar
        End Select
   
   
End Sub




Private Sub optConHistoriaClinica_Click(Value As Integer)
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, True
    mo_Formulario.HabilitarDeshabilitar txtPaciente, False
    mo_Formulario.HabilitarDeshabilitar txtIdOrden, False
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    txtNroHistoria.Text = ""
    txtPaciente.Text = ""

End Sub



Private Sub optSinHistoriaClinica_Click(Value As Integer)
    Me.ucFacturacionProductos.LimpiarGrilla
    Me.ucFacturacionProductos.IdOrden = -1
    Me.ucFacturacionProductos.CargaProductosPorIdOrden
            
    mo_cmbIdTipoGenHistoriaClinica.BoundText = ""
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar txtPaciente, False
    mo_Formulario.HabilitarDeshabilitar txtIdOrden, False
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    txtNroHistoria.Text = ""
    txtPaciente.Text = ""

End Sub




Private Sub txtNroHistoria_LostFocus()
    Dim oPaciente As New doPaciente
    Dim rsRespuesta As New Recordset
    If txtNroHistoria = "" Then
        Exit Sub
    End If
    Me.MousePointer = 11
    
    oPaciente.NroHistoriaClinica = Val(txtNroHistoria)
    oPaciente.IdTipoNumeracion = Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
    Set rsRespuesta = mo_AdminAdmision.PacientesFiltrar(oPaciente)
    
    If rsRespuesta.RecordCount = 0 Then
        MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
    ElseIf rsRespuesta.RecordCount = 1 Then
        ml_IdPaciente = rsRespuesta!idPaciente
        txtPaciente = rsRespuesta!ApellidoPaterno + " " + rsRespuesta!ApellidoMaterno + " " + rsRespuesta!PrimerNombre
       
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
            MsgBox "No tiene cuenta", vbInformation, "Búsqueda"
        End If
    Else
        MsgBox "No tiene cuenta", vbInformation, "Búsqueda"
    End If
    Me.MousePointer = 1
End Sub

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


Private Sub grdPacientesEncontrados_DblClick()
    Dim rsPaciente As Recordset
    On Error Resume Next
    Set rsPaciente = Me.grdPacientesEncontrados.DataSource
    txtNroHistoria.Text = rsPaciente.Fields!NroHistoriaClinica
    mo_cmbIdTipoGenHistoriaClinica.BoundText = rsPaciente.Fields!IdTipoNumeracion
    txtNroHistoria_LostFocus
    Me.grdPacientesEncontrados.Visible = False: DoEvents
End Sub

Private Sub grdPacientesEncontrados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPacientesEncontrados.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Width = 1300
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Width = 2000

    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult. Fec Ing."
    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult. Fec Egr."
    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Width = 1500
End Sub

Private Sub grdPacientesEncontrados_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        grdPacientesEncontrados.Visible = False
    End If
    
End Sub

Private Sub grdPacientesEncontrados_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = vbKeyReturn Then
        grdPacientesEncontrados_DblClick
    End If
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
'        oConexion.CommandTimeout = 300
'        oConexion.Open SIGHComun.CadenaConexion
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
             txtNroHistoria.Text = oPaciente.NroHistoriaClinica
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
        oWorkSheet.Cells(1, 2).Value = "Nº de Orden: " & mo_DOFactOrdenBienInsumo.IdOrden
        oWorkSheet.Cells(3, 3).Value = txtPaciente.Text
        oWorkSheet.Cells(3, 6).Value = txtNroCuenta.Text
        oWorkSheet.Cells(4, 3).Value = cmbIdPuntoDeCarga.Text
        oWorkSheet.Cells(4, 6).Value = txtNroHistoria.Text
        oWorkSheet.Cells(5, 3).Value = cmbFormaPago.Text
        If txtHentrega.Text <> "__:__" Then
           oWorkSheet.Cells(5, 6).Value = CDate(txtFentrega.Text & " " & txtHentrega.Text)
        End If
        oWorkSheet.Cells(6, 3).Value = ""
        oWorkSheet.Cells(6, 2).Value = ""
        
        Set rsReporte = Me.ucFacturacionProductos.FacturacionProductos
        iFila = 9: lnTotal = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
           oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!codigo
           oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields!NombreProducto
           oWorkSheet.Cells(iFila, 5).Value = Format(rsReporte.Fields!cantidad, "####,###")
           oWorkSheet.Cells(iFila, 6).Value = Format(rsReporte.Fields!precioUnitario, "####,##0.000")
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


Sub CargaOrdenYaRegistradaEnSisSoat()
    If mi_Opcion = sghAgregar Then
        Dim LnIdFormaPago As Long
        mi_Opcion = sghModificar
        Me.IdOrden = txtNroOrdenSisSoat.Text
        CargarDatosAlFormulario
        LnIdFormaPago = Me.ucFacturacionProductos.OrdenRegistradaYaprobadaPorSisSoat
        If esfecha(Left(mo_DOFactOrdenBienInsumo.FechaOrden, 10), "DD/MM/AAAA") = False Then
           If LnIdFormaPago > 1 Then
                cmbFormaPago.BoundText = LnIdFormaPago  '2=sis, 3=soat, 4=convenio, 9=exo
                cmbFarmacia.BoundText = lnCodigoFarmacia
                cmbFormaPago.Enabled = False
                fraDatosHistoria.Enabled = True
                txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL  'Format(Now, "HH:MM")
                txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL 'Format(Date, "DD/MM/YYYY")
                lbDocumentoYaRegistradoEnSeguros = True
                'Cargar datos de los servicios
                Me.ucFacturacionProductos.LimpiarGrilla
                Me.ucFacturacionProductos.DocumentoYaRegistradoEnSeguros = lbDocumentoYaRegistradoEnSeguros
                Me.ucFacturacionProductos.IdOrden = Me.IdOrden
                Me.ucFacturacionProductos.IdEstadoOrden = mo_DOFactOrdenBienInsumo.IdEstadoOrden
                Me.ucFacturacionProductos.CargaProductosPorIdOrden
                
            Else
                MsgBox "Ese 'Nro de Orden' NO ha sido Aprobado en SIS/SOAT/EXO/Convenio", vbCritical, "Mensaje"
                btnCancelar_Click
            End If
        Else
            MsgBox "Ese 'Nro de Orden' ha sido Registrado y Aprobado en SIS/SOAT/EXO/Convenio, pero ya fué despachado", vbCritical, "Mensaje"
            btnCancelar_Click
        End If
   End If
End Sub

