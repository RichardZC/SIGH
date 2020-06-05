VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FarmAlmacenDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FarmAlmacenDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1065
      Left            =   0
      TabIndex        =   7
      Top             =   4500
      Width           =   5790
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarmAlmacenDetalle.frx":0CCA
         DownPicture     =   "FarmAlmacenDetalle.frx":118E
         Height          =   700
         Left            =   2940
         Picture         =   "FarmAlmacenDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarmAlmacenDetalle.frx":1B66
         DownPicture     =   "FarmAlmacenDetalle.frx":1FC6
         Height          =   700
         Left            =   1395
         Picture         =   "FarmAlmacenDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
      Height          =   4515
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5790
      Begin VB.CheckBox chkEsUnidosis 
         Alignment       =   1  'Right Justify
         Caption         =   "Es FARMACIA UNIDOSIS"
         Height          =   255
         Left            =   3165
         TabIndex        =   26
         Top             =   1380
         Width           =   2520
      End
      Begin VB.TextBox txtCodigoDigemid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4290
         MaxLength       =   20
         TabIndex        =   24
         Top             =   960
         Width           =   1410
      End
      Begin VB.Frame Frame1 
         Caption         =   "REGENERAR SALDOS automáticamente"
         Height          =   2655
         Left            =   135
         TabIndex        =   10
         Top             =   1770
         Width           =   5550
         Begin VB.CheckBox chkDomingo 
            Caption         =   "Domingo"
            Height          =   240
            Left            =   780
            TabIndex        =   21
            Top             =   2190
            Width           =   1140
         End
         Begin VB.CheckBox chkSabado 
            Caption         =   "Sábado"
            Height          =   240
            Left            =   780
            TabIndex        =   20
            Top             =   1875
            Width           =   1080
         End
         Begin VB.CheckBox chkViernes 
            Caption         =   "Viernes"
            Height          =   240
            Left            =   780
            TabIndex        =   19
            Top             =   1560
            Width           =   1140
         End
         Begin VB.CheckBox chkJueves 
            Caption         =   "Jueves"
            Height          =   240
            Left            =   780
            TabIndex        =   18
            Top             =   1245
            Width           =   1230
         End
         Begin VB.CheckBox chkMiercoles 
            Caption         =   "Miércoles"
            Height          =   240
            Left            =   780
            TabIndex        =   17
            Top             =   945
            Width           =   1110
         End
         Begin VB.CheckBox chkMartes 
            Caption         =   "Martes"
            Height          =   240
            Left            =   780
            TabIndex        =   16
            Top             =   630
            Width           =   990
         End
         Begin VB.CheckBox chkLunes 
            Caption         =   "Lunes"
            Height          =   240
            Left            =   780
            TabIndex        =   15
            Top             =   315
            Width           =   825
         End
         Begin VB.TextBox txtEstadoRegenerar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3675
            MaxLength       =   20
            TabIndex        =   14
            Top             =   630
            Width           =   1770
         End
         Begin MSMask.MaskEdBox txtHrInicioRegenerar 
            Height          =   315
            Left            =   4695
            TabIndex        =   22
            Top             =   270
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
         Begin VB.Label Label5 
            Caption         =   $"FarmAlmacenDetalle.frx":28B0
            ForeColor       =   &H000000FF&
            Height          =   1350
            Left            =   2535
            TabIndex        =   23
            Top             =   1095
            Width           =   2925
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado actual "
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2535
            TabIndex        =   13
            Top             =   675
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hora de inicio del proceso"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2535
            TabIndex        =   12
            Top             =   300
            Width           =   2100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Días"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   11
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.ComboBox cmbTipoAlmacen 
         Height          =   330
         Left            =   960
         TabIndex        =   9
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   4755
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código DIGEMID"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2970
         TabIndex        =   25
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   990
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Width           =   645
      End
   End
End
Attribute VB_Name = "FarmAlmacenDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Almacenes
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_DoFarmAlmacen As New DoFarmAlmacen
Dim ms_MensajeError As String
Dim mb_ExistenDatos As Boolean
Dim mo_ReglasFarmacia As New ReglasFarmacia
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdAlmacen As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_cmbTipoAlmacen As New SIGHEntidades.ListaDespleglable
Dim mo_AdminComun As New ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let IdDependenciaExt(lValue As Long)
   ml_IdAlmacen = lValue
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
     mo_Formulario.HabilitarDeshabilitar Me.txtCodigo, False
     mo_Formulario.HabilitarDeshabilitar Me.txtCodigoDigemid, False
     'RZC 13/02/2020 Cambio5 Inicio
     If mi_Opcion = sghEliminar Then
        mo_Formulario.HabilitarDeshabilitar Me.txtDescripcion, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoAlmacen, False
     End If
     If mi_Opcion = sghModificar Then
        mo_Formulario.HabilitarDeshabilitar Me.txtDescripcion, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoAlmacen, False
     End If
    '     If mi_Opcion <> sghAgregar Then
    '        mo_Formulario.HabilitarDeshabilitar Me.txtDescripcion, False
    '        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoAlmacen, False
    '     End If
    'RZC 13/02/2020 Cambio5 Fin
     Select Case mi_Opcion
     Case sghAgregar
         CargaUltimoCorrelativoIdAlmacen
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

Sub CargaUltimoCorrelativoIdAlmacen()
    txtCodigo.Text = Trim(Str(mo_ReglasFarmacia.CargaUltimoCorrelativoIdAlmacen))
End Sub



Private Sub Form_Initialize()
    Set mo_cmbTipoAlmacen.MiComboBox = Me.cmbTipoAlmacen
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
   
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Farmacia"
       Case sghModificar
           Me.Caption = "Modificar Farmacia"
       Case sghConsultar
           Me.Caption = "Consultar Farmacia"
       Case sghEliminar
           Me.Caption = "Anular Farmacia"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
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
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   'LimpiarFormulario
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
       If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox "Los datos se Anularon correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
        End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "No hay el Id" + Chr(13)
   End If
   If Trim(Me.txtDescripcion) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre del almacén" + Chr(13)
       txtDescripcion.SetFocus
   End If
   If Trim(Me.cmbTipoAlmacen) = "" Then
        sMensaje = sMensaje + "Ingrese tipo de almacén" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If

   ValidarDatosObligatorios = True

End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   Dim orsTemp As New Recordset
   Dim lnCorrelativo As Integer
   With mo_DoFarmAlmacen
        .IdAlmacen = Val(Me.txtCodigo.Text)
        .Descripcion = UCase(Me.txtDescripcion.Text)
        .idEstado = sghEstadoTabla.sghRegistrado
        .idTipoLocales = "F"
        .idTipoSuministro = "0" & mo_cmbTipoAlmacen.BoundText
        .IdUsuarioAuditoria = ml_idUsuario
        Set orsTemp = mo_ReglasFarmacia.FarmAlmacenObtieneCodigoSismed("0" & mo_cmbTipoAlmacen.BoundText)
        If orsTemp.RecordCount > 0 Then
            lnCorrelativo = CInt(IIf(IsNull(orsTemp.Fields!correlativo), 0, orsTemp.Fields!correlativo)) 'Actualizado 06102014
            lnCorrelativo = lnCorrelativo + 1
        Else
            lnCorrelativo = 1
        End If
        .CodigoSismed = CStr(lcBuscaParametro.SeleccionaFilaParametro(208)) + "F" + "0" + CStr(lnCorrelativo)
        .esUnidosis = IIf(chkEsUnidosis.Value = 1, 1, 0)
   End With
   CargaRestoDatos
End Sub

Sub CargaRestoDatos()
   With mo_DoFarmAlmacen
        .regenerarEstado = Me.txtEstadoRegenerar.Text
        .regenerarHora = Me.txtHrInicioRegenerar.Text
        .regenerarDias = IIf(chkLunes.Value = 1, "2", "") & _
                        IIf(chkMartes.Value = 1, "3", "") & _
                        IIf(chkMiercoles.Value = 1, "4", "") & _
                        IIf(chkJueves.Value = 1, "5", "") & _
                        IIf(chkViernes.Value = 1, "6", "") & _
                        IIf(chkSabado.Value = 1, "7", "") & _
                        IIf(chkDomingo.Value = 1, "1", "")
   End With

End Sub

Sub ActualizaItemsUnidosisEnTablaFactCatalalogoBI()
    If mo_DoFarmAlmacen.esUnidosis = 1 Then
        Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
        Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos
        Dim oConexion As New ADODB.Connection
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim lcMensaje As String, lcCodigo As String, lcCodigoConPunto As String, lcListaCodigos As String
        Dim lnItem As Integer
        lcMensaje = ""
        lcListaCodigos = "/"
        lnItem = 0
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        Set oRsTmp1 = mo_ReglasFarmacia.farmUnidosisSeleccionarTodos(oConexion)
        If oRsTmp1.RecordCount > 0 Then
           Set oCatalogoBienesInsumos.Conexion = oConexion
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lnItem = lnItem + 1
              lcCodigoConPunto = Trim(oRsTmp1!codigo) & SIGHEntidades.Pto
              lcCodigo = Trim(oRsTmp1!codigo)
              lcListaCodigos = lcListaCodigos & lcCodigoConPunto & "/"
              If lnItem > 10 Then
                 lnItem = 0
                 lcListaCodigos = lcListaCodigos & Chr(13)
              End If
              Set oRsTmp2 = oCatalogoBienesInsumos.SeleccionarPorCodigo(lcCodigoConPunto, oConexion)
              If oRsTmp2.RecordCount = 0 Then
                 oRsTmp2.Close
                 Set oRsTmp2 = oCatalogoBienesInsumos.SeleccionarPorCodigo(lcCodigo, oConexion)
                 If oRsTmp2.RecordCount = 0 Then
                    lcMensaje = lcMensaje & "No existe el CODIGO: " & oRsTmp1!codigo & Chr(13)
                 Else
                    oDOCatalogoBienesInsumos.idProducto = oRsTmp2!idProducto
                    oDOCatalogoBienesInsumos.IdUsuarioAuditoria = SIGHEntidades.Usuario
                    If oCatalogoBienesInsumos.SeleccionarPorId(oDOCatalogoBienesInsumos) Then
                       oDOCatalogoBienesInsumos.codigo = lcCodigoConPunto
                       oDOCatalogoBienesInsumos.Nombre = Left(oRsTmp1!Descripcion, 300)
                       oDOCatalogoBienesInsumos.FormaFarmaceutica = Left(oRsTmp1!umConsumo, 10)
                       If oCatalogoBienesInsumos.Insertar(oDOCatalogoBienesInsumos) Then
                       End If
                    End If
                 End If
              End If
              oRsTmp2.Close
              oRsTmp1.MoveNext
           Loop
           If lcMensaje = "" Then
              MsgBox "Se agregó ITEMS para las FARMACIAS DE UNIDOSIS" & Chr(13) & _
                     "       no olvidar asignar PRECIOS a cada uno:  " & Chr(13) & _
                     lcListaCodigos, vbInformation, ""
           Else
              MsgBox lcMensaje, vbInformation, ""
           End If
        End If
        Set oDOCatalogoBienesInsumos = Nothing
        Set oCatalogoBienesInsumos = Nothing
        Set oConexion = Nothing
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
    End If
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   mo_DoFarmAlmacen.Descripcion = mo_DoFarmAlmacen.Descripcion & "-" & Me.cmbTipoAlmacen.Text
   AgregarDatos = mo_ReglasFarmacia.FarmAlmacenAgregar(mo_DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
   ActualizaItemsUnidosisEnTablaFactCatalalogoBI
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
   With mo_DoFarmAlmacen
       .Descripcion = UCase(txtDescripcion.Text)
       .IdUsuarioAuditoria = ml_idUsuario
       .idTipoSuministro = "0" & mo_cmbTipoAlmacen.BoundText
       .esUnidosis = IIf(chkEsUnidosis.Value = 1, 1, 0)
   End With
   CargaRestoDatos
   ModificarDatos = mo_ReglasFarmacia.farmalmacenmodificar(mo_DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
   ActualizaItemsUnidosisEnTablaFactCatalalogoBI
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
   With mo_DoFarmAlmacen
       .idEstado = sghEstadoTabla.sghAnulado
       .IdUsuarioAuditoria = ml_idUsuario
   End With
   EliminarDatos = mo_ReglasFarmacia.farmalmacenmodificar(mo_DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_DoFarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(ml_IdAlmacen)
    If Not mo_DoFarmAlmacen Is Nothing Then
        With mo_DoFarmAlmacen
            Me.txtDescripcion = .Descripcion
            Me.txtCodigo = .IdAlmacen
            mo_cmbTipoAlmacen.BoundText = .idTipoSuministro
            txtCodigoDigemid.Text = .CodigoSismed
            Me.txtEstadoRegenerar.Text = .regenerarEstado
            Me.txtHrInicioRegenerar.Text = IIf(.regenerarHora = "", SIGHEntidades.HORA_VACIA_HM, .regenerarHora)
            If InStr(.regenerarDias, "1") > 0 Then
               Me.chkDomingo.Value = 1
            End If
            If InStr(.regenerarDias, "2") > 0 Then
               Me.chkLunes.Value = 1
            End If
            If InStr(.regenerarDias, "3") > 0 Then
               Me.chkMartes.Value = 1
            End If
            If InStr(.regenerarDias, "4") > 0 Then
               Me.chkMiercoles.Value = 1
            End If
            If InStr(.regenerarDias, "5") > 0 Then
               Me.chkJueves.Value = 1
            End If
            If InStr(.regenerarDias, "6") > 0 Then
               Me.chkViernes.Value = 1
            End If
            If InStr(.regenerarDias, "7") > 0 Then
               Me.chkSabado.Value = 1
            End If
            chkEsUnidosis.Value = .esUnidosis
            mb_ExistenDatos = True
            If .idEstado = 0 Then
               btnAceptar.Enabled = False
            End If
        End With
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    ml_IdAlmacen = 0
    
    Me.txtDescripcion = ""
    Me.txtCodigo = ""
    chkEsUnidosis.Value = 0
End Sub

Sub CargarComboBoxes()
    mo_cmbTipoAlmacen.BoundColumn = "idTipoSuministro"
    mo_cmbTipoAlmacen.ListField = "descripcion"
    Set mo_cmbTipoAlmacen.RowSource = mo_AdminComun.FarmAlmacenesTipoSuministroSeleccionarTodos
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
   End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_DoFarmAlmacen = Nothing
    Set mo_ReglasFarmacia = Nothing

End Sub
