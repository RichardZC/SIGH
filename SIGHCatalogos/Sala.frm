VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form Sala 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   11610
   ClientTop       =   4545
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7770
   Begin VB.Frame frSalaPtosCarga 
      Height          =   2895
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   7695
      Begin VB.CommandButton btnQuitar 
         DisabledPicture =   "Sala.frx":0000
         DownPicture     =   "Sala.frx":038B
         Height          =   315
         Left            =   3930
         Picture         =   "Sala.frx":071E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   645
         Width           =   795
      End
      Begin VB.CommandButton btnAgregar 
         DisabledPicture =   "Sala.frx":0AAF
         DownPicture     =   "Sala.frx":0E98
         Height          =   315
         Left            =   2850
         Picture         =   "Sala.frx":12A4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   645
         Width           =   795
      End
      Begin VB.ComboBox cmbIdPtoCarga 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Sala.frx":16B0
         Left            =   1290
         List            =   "Sala.frx":16B2
         TabIndex        =   12
         Text            =   "cmbIdPtoCarga"
         Top             =   240
         Width           =   6135
      End
      Begin UltraGrid.SSUltraGrid grdPuntosDeCarga 
         Height          =   1560
         Left            =   120
         TabIndex        =   11
         Top             =   1170
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   2752
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
         Caption         =   "Puntos de Carga"
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pto de Carga"
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
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Sala"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7680
      Begin VB.ComboBox cmbIdTipoModalidadSala 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   240
         Width           =   5445
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   2
         Top             =   990
         Width           =   5505
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   1
         Top             =   630
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "Modalidad Sala"
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCodigoCIE2004 
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
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1050
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   0
      TabIndex        =   5
      Top             =   4470
      Width           =   7680
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "Sala.frx":16B4
         DownPicture     =   "Sala.frx":1B14
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
         Left            =   2565
         Picture         =   "Sala.frx":1F89
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "Sala.frx":23FE
         DownPicture     =   "Sala.frx":28C2
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
         Left            =   4080
         Picture         =   "Sala.frx":2DAE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "Sala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de tipo de modalidad de sala
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_ImagSala As New DOImagSala
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdSala As Long
Dim mo_ReglasImagenes As New ReglasImagenes
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_cmbIdTipoModalidadSala As New sighentidades.ListaDespleglable
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mb_RegistroInactivo As Boolean
Dim mrs_PuntosCarga As New Recordset
Dim mb_ResultadoOperacion As Boolean

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
Property Let IdSala(lValue As Long)
   ml_IdSala = lValue
End Property
Property Get IdSala() As Long
   IdSala = ml_IdSala
End Property

Property Let ResultadoOperacion(bValue As Boolean)
   mb_ResultadoOperacion = bValue
End Property

Property Get ResultadoOperacion() As Boolean
   ResultadoOperacion = mb_ResultadoOperacion
End Property

Private Sub btnAgregar_Click()
    Dim lbNuevo As Boolean
    Dim lcIdServicio As String
    If mo_cmbIdPuntoCarga.BoundText = "" Then
        MsgBox "Seleccione Punto de carga a agregar", vbInformation, "Punto de Carga"
        Exit Sub
    End If
    If cmbIdPtoCarga.Text <> "" Then
        If ExistePuntoCarga() = True Then
            MsgBox "Ya se asigno punto de carga a esta sala", vbInformation, Me.Caption
            Exit Sub
        End If
        agregarPuntoCarga Val(mo_cmbIdPuntoCarga.BoundText), Trim(cmbIdPtoCarga.Text)
        mo_cmbIdPuntoCarga.BoundText = ""
    End If
End Sub

Private Sub btnQuitar_Click()
    On Error Resume Next
    EliminarRegistroSeleccionadoDeRs mrs_PuntosCarga, "", "Puntos de Carga"
End Sub

Private Sub cmbIdTipoModalidadSala_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoModalidadSala
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoModalidadSala.MiComboBox = cmbIdTipoModalidadSala
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
    
'    mo_Apariencia.ConfigurarFilasBiColores grdPuntosDeCarga, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub grdPuntosDeCarga_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPuntosDeCarga.Bands(0).Columns("idPuntoCarga").Header.Caption = "Id"
    grdPuntosDeCarga.Bands(0).Columns("Descripcion").Header.Caption = "Descripcion"
    
    grdPuntosDeCarga.Bands(0).Columns("idPuntoCarga").Width = 600
    grdPuntosDeCarga.Bands(0).Columns("Descripcion").Width = grdPuntosDeCarga.Width - _
                                grdPuntosDeCarga.Bands(0).Columns("idPuntoCarga").Width - 600
    
    
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
   AdministrarKeyPreview KeyCode
End Sub
Private Sub txtCodigo_LostFocus()
    txtCodigo = UCase(txtCodigo)
   mo_Formulario.MarcarComoVacio txtCodigo
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
         Frame1.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         Frame1.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
    CreaTemporal
    CargaComboBoxes
       
    Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agregar Sala"
        Case sghModificar
            Me.Caption = "Modificar Sala"
        Case sghConsultar
            Me.Caption = "Consultar Sala"
        Case sghEliminar
            Me.Caption = "Eliminar Sala"
    End Select
    CargarDatosAlFormulario
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
    
    mo_Apariencia.ConfigurarFilasBiColores grdPuntosDeCarga, sighentidades.GrillaConFilasBicolor
End Sub

Sub CargaComboBoxes()
    mo_cmbIdTipoModalidadSala.BoundColumn = "IdTipoModalidadSala"
    mo_cmbIdTipoModalidadSala.ListField = "TipoModalidadSala"
    Set mo_cmbIdTipoModalidadSala.RowSource = mo_ReglasImagenes.ImagTipoModalidadSalaSeleccionarTodos()
    
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
    mo_cmbIdPuntoCarga.ListField = ":Descripcion| (|:TipoPunto| )"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSourceTextCompuesto = mo_AdminComun.SeleccionarPuntosDeCarga()
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
   mb_ResultadoOperacion = False
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            mb_RegistroInactivo = False
           If ValidarReglas() Then
                'Dim bSuccess As Boolean
                If mb_RegistroInactivo = False Then
                    mb_ResultadoOperacion = AgregarDatos()
                Else
                    mb_ResultadoOperacion = ModificarDatos()
                End If
               If mb_ResultadoOperacion Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   Me.txtCodigo.SetFocus
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasImagenes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   mb_ResultadoOperacion = True
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasImagenes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   mb_ResultadoOperacion = True
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasImagenes.MensajeError, vbExclamation, Me.Caption
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
  
   If Val(mo_cmbIdTipoModalidadSala.BoundText) = 0 Then
       sMensaje = sMensaje + "Seleccione Tipo de Modalidad de Sala" + Chr(13)
   End If
   If Me.txtCodigo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código " + Chr(13)
   End If
   If Me.txtDescripcion = "" Then
       sMensaje = sMensaje + "Ingrese la descripción" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   'si existe y esta inactivo dejar que siga su curso
   CargaDatosAlObjetosDeDatos
   Dim oRs As ADODB.Recordset
   
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
        If mi_Opcion = sghAgregar Then
            mo_ImagSala.IdSala = 0
        Else
            mo_ImagSala.IdSala = ml_IdSala
        End If
        Set oRs = mo_ReglasImagenes.ImagSalaVerificarPorNombre(mo_ImagSala)
         If oRs Is Nothing Then
             MsgBox mo_ReglasImagenes.MensajeError, vbInformation, "Sala"
             Exit Function
         Else
             Dim lNumeroError As Long
             lNumeroError = ValidarDuplicadoRegistro(mi_Opcion, oRs)
             If lNumeroError > 0 Then
                MsgBox mapearMensageErrorDupplicado(lNumeroError), vbInformation, "Sala"
                Exit Function
             End If
         End If
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   With mo_ImagSala
        .codigo = Me.txtCodigo.Text
        .Sala = UCase(Me.txtDescripcion.Text)
        .IdTipoModalidadSala = mo_cmbIdTipoModalidadSala.BoundText
        .IdUsuarioAuditoria = Me.idUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_ReglasImagenes.ImagSalaAgregar(mo_ImagSala, ObtenerPuntosCarga(), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_ReglasImagenes.ImagSalaModificar(mo_ImagSala, ObtenerPuntosCarga(), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_ReglasImagenes.ImagSalaEliminar(mo_ImagSala, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    Set mo_ImagSala = mo_ReglasImagenes.ImagSalaSeleccionarPorId(Me.IdSala)
    If mo_ReglasImagenes.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_ReglasImagenes.MensajeError, vbInformation, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_ImagSala Is Nothing Then
        With mo_ImagSala
            Me.txtCodigo = .codigo
            Me.txtDescripcion = .Sala
            mo_cmbIdTipoModalidadSala.BoundText = .IdTipoModalidadSala
            mb_ExistenDatos = True
        End With
        CreaTemporal
        cargarPuntosCargaPorSala mo_ImagSala.IdSala
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
    CreaTemporal
    Me.IdSala = 0
    Me.txtCodigo = ""
    Me.txtDescripcion = ""
    mo_cmbIdTipoModalidadSala.BoundText = ""
    Err = 0
End Sub


Private Function ValidarDuplicadoRegistro(iOpcion As sghOpciones, oRs As ADODB.Recordset) As Integer
    Dim lTotalRegistros As Long
    ValidarDuplicadoRegistro = 0
    
    If oRs.RecordCount > 0 Then
        lTotalRegistros = oRs.RecordCount
        
        oRs.MoveFirst
        While oRs.EOF = False
            If oRs.Fields!EsActivo = True Then
                'duplicado activo
                ValidarDuplicadoRegistro = 1
                Exit Function
            Else
                If iOpcion = sghModificar Then
                    'registro duplicado e inactivo si se permite la edicion ocasionaria dos registros uno activo y otro inactivo
                    ValidarDuplicadoRegistro = 2
                    Exit Function
                ElseIf iOpcion = sghAgregar Then
                    'si hay mas de un registro inactivo duplicado no se puede permitir activar uno al azar
                    If lTotalRegistros > 1 Then
                        ValidarDuplicadoRegistro = 3
                        Exit Function
                    Else
'                        ml_IdSala = oRs.Fields!IdSala
                        mo_ImagSala.IdSala = oRs.Fields!IdSala
                        mb_RegistroInactivo = True
                        Exit Function
                    End If
                End If
            End If
            oRs.MoveNext
        Wend
    End If
End Function

Private Function mapearMensageErrorDupplicado(lNumeroError As Long) As String
    Dim sMessage As String
    Select Case lNumeroError
        Case 1:
            sMessage = "Existe un registro con la misma descripción o código"
        Case 2:
            sMessage = "Existe un registro con la misma descripción o código que ha sido eliminado"
        Case 3:
            sMessage = "Existe mas de un registro con la misma descripción o código que han sido eliminados"
    End Select
    
    mapearMensageErrorDupplicado = sMessage
End Function

Sub CreaTemporal()
    If mrs_PuntosCarga.State = adStateOpen Then mrs_PuntosCarga.Close
    With mrs_PuntosCarga
          .Fields.Append "IdPuntoCarga", adInteger, 4, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 200, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdPuntosDeCarga.DataSource = mrs_PuntosCarga
End Sub


Private Function cargarPuntosCargaPorSala(lIdSala As Long)
    Dim oRsPuntosCarga As ADODB.Recordset
    Dim oImagSalaPuntoCarga As New DOImagSalaPuntoCarga
    
    oImagSalaPuntoCarga.IdSala = lIdSala
    CreaTemporal
    Set oRsPuntosCarga = mo_ReglasImagenes.ImagSalaPuntoCargaFiltrarPorIdSala(oImagSalaPuntoCarga)
    If Not (oRsPuntosCarga Is Nothing) Then
        If oRsPuntosCarga.RecordCount > 0 Then
            oRsPuntosCarga.MoveFirst
            While oRsPuntosCarga.EOF = False
                agregarPuntoCarga oRsPuntosCarga.Fields!idPuntoCarga, oRsPuntosCarga.Fields!Descripcion
                oRsPuntosCarga.MoveNext
            Wend
            oRsPuntosCarga.MoveFirst
        End If
    End If
End Function

Private Function ExistePuntoCarga() As Boolean
    Dim oRs As ADODB.Recordset
    Dim bReturnValue As Boolean
    
    bReturnValue = False
    
    Set oRs = mrs_PuntosCarga.Clone()
    
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        oRs.Find "idPuntoCarga=" & mo_cmbIdPuntoCarga.BoundText
        If Not oRs.EOF Then
           bReturnValue = True
        End If
    End If
    ExistePuntoCarga = bReturnValue
End Function


Public Function EliminarRegistroSeleccionadoDeRs(oRs As ADODB.Recordset, Optional sMessageWarning = "", _
                        Optional sTitleMessage As String = "", _
                        Optional bShowWarning As Boolean = True, _
                        Optional bShowMessage As Boolean = True) As Boolean
On Error GoTo miError
    EliminarRegistroSeleccionadoDeRs = False
    Dim iResponseWarnig As Integer
    
    If sMessageWarning = "" Then
        sMessageWarning = "¿Desea eliminar registro seleccionado?"
    End If
    If sTitleMessage = "" Then
        sTitleMessage = "Eliminar Registro"
    End If
    
    If oRs.RecordCount = 0 Then
        MsgBox "No hay registros para eliminar", vbInformation, sTitleMessage
        Exit Function
    End If
    
    If bShowWarning = True Then
        If MsgBox(sMessageWarning, vbYesNo, sTitleMessage) = vbNo Then
            EliminarRegistroSeleccionadoDeRs = False
            Exit Function
        End If
    End If
    
    
    If oRs.RecordCount > 0 Then
        With oRs
            If Not .EOF And Not .BOF Then
               .Delete
               .Update
               If Not (.BOF = True And .EOF = True) Then
                    .MovePrevious
                    If .BOF = True Then
                        .MoveNext
                    End If
               End If
               EliminarRegistroSeleccionadoDeRs = True
            Else
                MsgBox "Seleccione registro a Eliminar", vbInformation, sTitleMessage
            End If
        End With
    End If
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, sTitleMessage
    End If
End Function

Private Function agregarPuntoCarga(lidPuntoCarga As Long, sDescripcion As String) As Boolean
    mrs_PuntosCarga.AddNew
    mrs_PuntosCarga.Fields!idPuntoCarga = Val(lidPuntoCarga)
    mrs_PuntosCarga.Fields!Descripcion = Trim(sDescripcion)
    mrs_PuntosCarga.Update
End Function

Private Function ObtenerPuntosCarga() As Collection
    Dim oPuntosCarga As New Collection
    Dim oRs As ADODB.Recordset
    Dim oDOImagSalaPuntoCarga As DOImagSalaPuntoCarga
    
    Set oRs = mrs_PuntosCarga.Clone()
    
    If oRs.RecordCount > 0 Then
        While oRs.EOF = False
            Set oDOImagSalaPuntoCarga = New DOImagSalaPuntoCarga
            
            oDOImagSalaPuntoCarga.idPuntoCarga = oRs.Fields!idPuntoCarga
            oPuntosCarga.Add oDOImagSalaPuntoCarga
            
            oRs.MoveNext
        Wend
    End If
    
    Set ObtenerPuntosCarga = oPuntosCarga
End Function
