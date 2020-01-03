VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Configación de Formularios
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------------
'   Función:        ConfigurarTitulos
'   Descripción:    Configura los titulos de los formularios basado en la opcion
'                   que se va a realizar.
'   Parámetros:     sNombre:    Es el nombre del formulario
'                   sOpcion:    Es la opcion Agregar, Modificar, Consultar, eliminar
'                   oForm:      Es el formulario que se desea configurar
'------------------------------------------------------------------------------------
Sub ConfigurarTitulos(sNombre As String, sOpcion As sghOpciones, oForm As Object)
    
    Select Case sOpcion
    Case sghAgregar
        oForm.Caption = "Agregar " + sNombre
    Case sghModificar
        oForm.Caption = "Modificar " + sNombre
    Case sghConsultar
        oForm.Caption = "Consultar " + sNombre
    Case sghEliminar
        oForm.Caption = "Eliminar " + sNombre
    End Select

End Sub

'------------------------------------------------------------------------------------
'   Función:        ConfigurarTipoLetra
'   Descripción:    Configura el tipo de letra para cada uno de los elementos del formulario
'   Parámetros:     sNombre:    Es el nombre del formulario
'                   sOpcion:    Es la opcion Agregar, Modificar, Consultar, eliminar
'                   oForm:      Es el formulario que se desea configurar
'------------------------------------------------------------------------------------
Sub ConfigurarTipoLetra(sTipoLetra As String, sTamanio As String, oForm As Object)
Dim oControl As Object

    Exit Sub

    oForm.Font.Name = sTipoLetra
    oForm.Font.Size = sTamanio
    For Each oControl In oForm.Controls
        On Error Resume Next
        oControl.Font.Name = sTipoLetra
        oControl.Font.Size = sTamanio
    Next
    
End Sub
Sub ConfigurarTipoLetraDeControles(oControls As Object)
Dim oControl As Object

    For Each oControl In oControls
        On Error Resume Next
        oControl.Font.Name = "Tahoma"
        oControl.Font.Size = "8"
    Next
    
End Sub


'------------------------------------------------------------------------------------
'   Función:        InicializarDatosDelFormulario
'   Descripción:    Inicializa (limpia) los datos de un formulario
'   Parámetros:     oForm:      Es el formulario que se desea configurar
'------------------------------------------------------------------------------------
Sub InicializarDatosDelFormulario(oForm As Object)
Dim oControl As Object

    For Each oControl In oForm.Controls
        
        If TypeOf oControl Is TextBox Then
            oControl.Text = ""
        End If
        
        If TypeOf oControl Is ComboBox Then
            oControl.ListIndex = -1
        End If
        
        If TypeOf oControl Is ListBox Then
            oControl.ListIndex = -1
        End If
        
        If TypeOf oControl Is MSFlexGrid Then
            oControl.FixedRows = 0
            oControl.Rows = 1
            oControl.Rows = 2
            oControl.FixedRows = 1
        End If
    
    Next
    
End Sub

Sub MarcarComoVacio(oControl As Object)
    
    On Error Resume Next
    If oControl.Locked = True Then
        Exit Sub
    End If
    
    If Trim(oControl.Text) = "" Then
        oControl.BackColor = &HD0FDFC   '&HCACBFD
    Else
        oControl.BackColor = &H80000005
    End If
    
End Sub
Sub HabilitarAlerta(oControl As Object, EsHabilitado As Boolean)
    If EsHabilitado = False Then
       HabilitarDeshabilitar oControl, False
    Else
       If TypeOf oControl Is TextBox Or TypeOf oControl Is ComboBox Or TypeOf oControl Is DataCombo Or TypeName(oControl) = "PVComboBox" Then
            oControl.Enabled = True
            oControl.Locked = True
            oControl.BackColor = vbRed
            oControl.ForeColor = vbWhite
       End If
    End If
End Sub
Sub HabilitarDeshabilitar(oControl As Object, EsHabilitado As Boolean)

    On Error Resume Next
    
    If Not EsHabilitado Then
        
        If TypeOf oControl Is Frame Then
            oControl.BackColor = &HF9EADF
            oControl.Enabled = False
        End If
        If TypeOf oControl Is Label Then
            oControl.ForeColor = &H808080
        End If
        If TypeOf oControl Is TextBox Or TypeOf oControl Is ComboBox Or TypeOf oControl Is DataCombo Or TypeName(oControl) = "PVComboBox" Then
            oControl.Locked = True
            oControl.BackColor = &HF9EADF   '&HE0E0E0
            oControl.ForeColor = &H808080
            
        End If
        If TypeOf oControl Is CommandButton Then
            oControl.ForeColor = &H808080
            oControl.Enabled = False
        End If
        If TypeOf oControl Is MaskEdBox Then
            oControl.BackColor = &HF9EADF
            oControl.ForeColor = &H808080
            oControl.Enabled = False
        End If
        'ZCC_2006
        If TypeOf oControl Is CheckBox Then
            oControl.Enabled = False
        End If
        
    Else
        If TypeOf oControl Is Frame Then
            oControl.BackColor = &H8000000F
            oControl.Enabled = True
        End If
        If TypeOf oControl Is Label Then
            oControl.ForeColor = &H0&
        End If
        If TypeOf oControl Is TextBox Or TypeOf oControl Is ComboBox Or TypeOf oControl Is DataCombo Or TypeName(oControl) = "PVComboBox" Then
            oControl.Locked = False
            oControl.BackColor = &HFFFFFF
            oControl.ForeColor = &H0&
        End If
        If TypeOf oControl Is CommandButton Then
            oControl.ForeColor = &H0&
            oControl.Enabled = True
        End If
        If TypeOf oControl Is MaskEdBox Then
            oControl.BackColor = &HFFFFFF
            oControl.ForeColor = &H0&
            oControl.Enabled = True
        End If
        'ZCC_2006
        If TypeOf oControl Is CheckBox Then
            oControl.Enabled = True
        End If
    End If
    
End Sub
Sub VisibleNoVisible(oControl As Object, EsVisible As Boolean)
    
    If Not EsVisible Then
        oControl.Visible = False
    Else
        oControl.Visible = True
    End If
    
End Sub

Sub LlenaComboConAnios(cmbAnios As Object)
    Dim lnAnio As Integer
    For lnAnio = Year(Date) To 1980 Step -1
        cmbAnios.AddItem Trim(str(lnAnio))
    Next
    cmbAnios.ListIndex = 0
End Sub

Sub LlenaComboConMeses(cmbMes As Object)
        cmbMes.AddItem "Enero"
        cmbMes.AddItem "Febrero"
        cmbMes.AddItem "Marzo"
        cmbMes.AddItem "Abril"
        cmbMes.AddItem "Mayo"
        cmbMes.AddItem "Junio"
        cmbMes.AddItem "Julio"
        cmbMes.AddItem "Agosto"
        cmbMes.AddItem "Setiembre"
        cmbMes.AddItem "Octubre"
        cmbMes.AddItem "Noviembre"
        cmbMes.AddItem "Diciembre"
        cmbMes.ListIndex = Month(Date) - 1
End Sub
'mgaray
Public Function ControlEstaHabilitado(oControl As Object) As Boolean
    On Error Resume Next
    Dim EsHabilitado As Boolean
    
    EsHabilitado = False
    
    If TypeOf oControl Is Frame Or TypeOf oControl Is CommandButton _
                    Or TypeOf oControl Is MaskEdBox Or TypeOf oControl Is CheckBox Then
        If oControl.Enabled = True Then
            EsHabilitado = True
        End If
    End If
    If TypeOf oControl Is TextBox Or TypeOf oControl Is ComboBox Or TypeOf oControl Is DataCombo Then
        If oControl.Locked = False And oControl.Enabled = True Then
            EsHabilitado = True
        End If
    End If
   
    ControlEstaHabilitado = EsHabilitado
    Err = 0
End Function

Public Sub controlSelectText(oControl As Object)
    If ControlEstaHabilitado(oControl) Then
        On Error Resume Next
        oControl.SelStart = 0
        oControl.SelLength = Len(oControl.Text)
        Err = 0
    End If
End Sub

