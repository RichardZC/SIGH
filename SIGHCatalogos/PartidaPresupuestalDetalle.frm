VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form PartidaPresupuestalDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "PartidaPresupuestalDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   30
      TabIndex        =   8
      Top             =   1140
      Width           =   12195
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
         Height          =   945
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   11910
         Begin VB.TextBox txtCodBusqueda 
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCptBusqueda 
            Height          =   315
            Left            =   1080
            TabIndex        =   12
            Top             =   480
            Width           =   7605
         End
         Begin VB.CommandButton btnLimpiar 
            Height          =   315
            Left            =   10395
            Picture         =   "PartidaPresupuestalDetalle.frx":0CCA
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   9030
            Picture         =   "PartidaPresupuestalDetalle.frx":38A6
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Código       Nombre"
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
            TabIndex        =   15
            Top             =   270
            Width           =   6825
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
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
            Left            =   150
            TabIndex        =   14
            Top             =   810
            Width           =   7635
         End
      End
      Begin UltraGrid.SSUltraGrid grdCpt 
         Height          =   5145
         Left            =   120
         TabIndex        =   16
         Top             =   1230
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   9075
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
         Caption         =   "Procedimientos asignados"
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   45
      TabIndex        =   5
      Top             =   90
      Width           =   12210
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   960
         MaxLength       =   250
         TabIndex        =   1
         Top             =   600
         Width           =   11085
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   30
      TabIndex        =   2
      Top             =   7710
      Width           =   12180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "PartidaPresupuestalDetalle.frx":64EF
         DownPicture     =   "PartidaPresupuestalDetalle.frx":694F
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
         Left            =   4687
         Picture         =   "PartidaPresupuestalDetalle.frx":6DC4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "PartidaPresupuestalDetalle.frx":7239
         DownPicture     =   "PartidaPresupuestalDetalle.frx":76FD
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
         Left            =   6232
         Picture         =   "PartidaPresupuestalDetalle.frx":7BE9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "PartidaPresupuestalDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Partidas Presupuestales
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_PartidaPresupuestal As New DOPartidasPresupuestales
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_idPartidaPresupuestal As Long
Dim mo_AdminComun As New ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion

Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mrs_Cpt As New Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idPartidaPresupuestal(lValue As Long)
   ml_idPartidaPresupuestal = lValue
End Property
Property Get idPartidaPresupuestal() As Long
   idPartidaPresupuestal = ml_idPartidaPresupuestal
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
         
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

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Partida Presupuestal"
       Case sghModificar
           Me.Caption = "Modificar Partida Presupuestal"
       Case sghConsultar
           Me.Caption = "Consultar Partida Presupuestal"
       Case sghEliminar
           Me.Caption = "Eliminar Partida Presupuestal"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       CreaTemporal
       SkinConfigura
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
'        btnAceptar.Picture = LoadPicture("")
'        btnCancelar.Picture = LoadPicture("")
        btnAceptar.Caption = ""
        btnCancelar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores Me.grdCpt, "99"
  End If
ErrSkin:
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
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Me.MousePointer = 11
    mrs_Cpt.Filter = ""
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                   'LimpiarFormulario
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
   Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Trim(Me.txtDescripcion) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre de la Partida" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   If mi_Opcion = sghAgregar Then
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oRsTmp1 = mo_ReglasComunes.FactPartidasPresupuestalesSeleccionarTodas(oConexion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           oRsTmp1.Find "codigo='" & Me.txtCodigo.Text & "'"
           If Not oRsTmp1.EOF Then
              MsgBox "Ya existe ese CODIGO", vbInformation, Me.Caption
              Set oConexion = Nothing
              Set oRsTmp1 = Nothing
              Set mo_ReglasComunes = Nothing
              'Exit Function
           End If
        End If
        Set oConexion = Nothing
        Set oRsTmp1 = Nothing
        Set mo_ReglasComunes = Nothing
   End If
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   
   With mo_PartidaPresupuestal
        .codigo = Me.txtCodigo.Text
        .Descripcion = Me.txtDescripcion.Text
        
        .IdUsuarioAuditoria = Me.idUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminComun.PartidasPresupuestalesAgregar(mo_PartidaPresupuestal, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)
   If AgregarDatos = True Then
        'Graba id de Centro costo en Procedimientos CPT
        Dim oRsTmp As New Recordset
        Dim lcSql As String
        If mrs_Cpt.RecordCount > 0 Then
           mrs_Cpt.MoveFirst
           Do While Not mrs_Cpt.EOF
             If mrs_Cpt.Fields!seleccionar = True Then
                mo_AdminComun.CatalogoServiciosActualizaPartidaXidProducto mo_PartidaPresupuestal.idPartidaPresupuestal, mrs_Cpt.Fields!IdProducto, False
             End If
             mrs_Cpt.MoveNext
           Loop
        End If
        Set oRsTmp = Nothing
    End If
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_AdminComun.PartidasPresupuestalesModificar(mo_PartidaPresupuestal, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)
    'Graba id de Centro costo en Procedimientos CPT
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       Do While Not mrs_Cpt.EOF
         If mrs_Cpt.Fields!seleccionar = True Then
            mo_AdminComun.CatalogoServiciosActualizaPartidaXidProducto mo_PartidaPresupuestal.idPartidaPresupuestal, mrs_Cpt.Fields!IdProducto, False
         Else
            mo_AdminComun.CatalogoServiciosActualizaPartidaXidProducto mo_PartidaPresupuestal.idPartidaPresupuestal, mrs_Cpt.Fields!IdProducto, True
         End If
         mrs_Cpt.MoveNext
       Loop
    End If
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   'Graba id de Centro costo en Procedimientos CPT
   mo_AdminComun.CatalogoServicioActualizaNULLenIdPartida mo_PartidaPresupuestal.idPartidaPresupuestal
   mo_AdminComun.CatalogoBienesInsumosActualizaNULLenIdPartida mo_PartidaPresupuestal.idPartidaPresupuestal
   '
   EliminarDatos = mo_AdminComun.PartidasPresupuestalesEliminar(mo_PartidaPresupuestal, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    mo_Formulario.HabilitarDeshabilitar Me.txtCodigo, False
    Set mo_PartidaPresupuestal = mo_AdminComun.PartidasPresupuestalesSeleccionarPorId(Me.idPartidaPresupuestal)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminComun.MensajeError, vbInformation, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_PartidaPresupuestal Is Nothing Then
        With mo_PartidaPresupuestal
            Me.txtDescripcion = .Descripcion
            Me.txtCodigo = .codigo
            
            mb_ExistenDatos = True
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

    Me.idPartidaPresupuestal = 0
    
    Me.txtDescripcion = ""
    Me.txtCodigo = ""
    
End Sub

Sub CargarComboBoxes()
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mrs_Cpt = Nothing
End Sub



Private Sub grdCpt_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdCpt.Bands(0).Columns("IdProducto").Hidden = True
    '
    grdCpt.Bands(0).Columns("Cpt").Header.Caption = "Cpt"
    grdCpt.Bands(0).Columns("Cpt").Width = 1000
    '
    grdCpt.Bands(0).Columns("Producto").Header.Caption = "Procedimiento"
    grdCpt.Bands(0).Columns("Producto").Width = 9500
    '
    grdCpt.Bands(0).Columns("Seleccionar").Header.Caption = "Seleccionar"
    grdCpt.Bands(0).Columns("Seleccionar").Width = 700

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

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub


Sub CreaTemporal()
    If mrs_Cpt.State = adStateOpen Then Set mrs_Cpt = Nothing
    With mrs_Cpt
          .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
          .Fields.Append "Cpt", adVarChar, 20, adFldIsNullable
          .Fields.Append "Producto", adVarChar, 250, adFldIsNullable
          .Fields.Append "Seleccionar", adBoolean
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    mo_Apariencia.ConfigurarFilasBiColores Me.grdCpt, sighentidades.GrillaConFilasBicolor
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    Set oRsTmp = mo_ReglasFacturacion.CatalogoServiciosSeleccionarTodos
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          mrs_Cpt.AddNew
          mrs_Cpt.Fields!IdProducto = oRsTmp.Fields!IdProducto
          mrs_Cpt.Fields!cpt = oRsTmp.Fields!codigo
          mrs_Cpt.Fields!Producto = Left(oRsTmp.Fields!Nombre, 250)
          If oRsTmp.Fields!IdPartida = ml_idPartidaPresupuestal Then
             mrs_Cpt.Fields!seleccionar = True
          Else
             mrs_Cpt.Fields!seleccionar = False
          End If
          mrs_Cpt.Update
          oRsTmp.MoveNext
       Loop
       mrs_Cpt.MoveFirst
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    mrs_Cpt.Sort = "seleccionar desc"
    Set Me.grdCpt.DataSource = mrs_Cpt
    
End Sub

Private Sub btnBuscar_Click()
   If txtCodBusqueda.Text = "" And Me.txtCptBusqueda.Text = "" Then
      MsgBox "Para la búsqueda debe ingresar el Código CPT o Parte del Procedimiento", vbInformation, Me.Caption
      Exit Sub
   End If
   
   If txtCodBusqueda.Text <> "" Then
      mrs_Cpt.MoveFirst
      mrs_Cpt.Find "cpt='" & Trim(txtCodBusqueda.Text) & "'"
      If mrs_Cpt.EOF Then
         MsgBox "No se encontró datos para ese código"
      End If
   Else
       'Por Nombre
        mrs_Cpt.Filter = "Producto like '%" & Trim(Me.txtCptBusqueda.Text) & "%'"
   End If
End Sub

Private Sub btnLimpiar_Click()
    txtCodBusqueda.Text = ""
    Me.txtCptBusqueda.Text = ""
    mrs_Cpt.Filter = ""
End Sub

