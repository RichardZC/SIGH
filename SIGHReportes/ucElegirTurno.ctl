VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucElegirTurno 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   LockControls    =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   2025
   Begin VB.Frame fraTurnos 
      Caption         =   "Turnos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   -15
      TabIndex        =   0
      Top             =   15
      Width           =   2055
      Begin Threed.SSOption optAmbos 
         Height          =   240
         Left            =   195
         TabIndex        =   1
         Top             =   255
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   423
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
         Caption         =   "Ambos"
         Value           =   -1
      End
      Begin Threed.SSOption optManana 
         Height          =   240
         Left            =   195
         TabIndex        =   2
         Top             =   570
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   423
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
         Caption         =   "Mañana"
      End
      Begin Threed.SSOption optTarde 
         Height          =   240
         Left            =   195
         TabIndex        =   3
         Top             =   1140
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   423
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
         Caption         =   "Tarde"
      End
      Begin MSMask.MaskEdBox txtMinicio 
         Height          =   315
         Left            =   495
         TabIndex        =   4
         Top             =   825
         Width           =   735
         _ExtentX        =   1296
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
      Begin MSMask.MaskEdBox txtTinicio 
         Height          =   315
         Left            =   495
         TabIndex        =   5
         Top             =   1425
         Width           =   735
         _ExtentX        =   1296
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
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtmFinal 
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   825
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
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
      Begin MSMask.MaskEdBox txtTfinal 
         Height          =   315
         Left            =   1260
         TabIndex        =   7
         Top             =   1425
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
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
   End
End
Attribute VB_Name = "ucElegirTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Historia Clinica
'        Programado por: Barrantes D
'        Fecha: Octubre 2018
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event SeModificoTurnos(lnTurno As sghTurnos, lcHrMinicio As String, lcHrMfinal As String, _
                              lcHrTinicio As String, lcHrTfinal As String)


Sub Inicializar()
    TurnosCargar
End Sub

Sub TurnosCargar()
    txtMinicio.Text = "07:00"
    txtmFinal.Text = "13:00"
    txtTinicio.Text = "14:00"
    txtTfinal.Text = "18:00"
    Dim lcTurnosTotales As String
    lcTurnosTotales = SIGHEntidades.TurnosMananaTarde
    If lcTurnosTotales = "" Then
       TurnosGrabar
       lcTurnosTotales = SIGHEntidades.TurnosMananaTarde
    End If
    '/hh:mm/hh:mm/hh:mm/hh:mm/
    '1234567890123456789012345
    txtMinicio.Text = Mid(lcTurnosTotales, 2, 5)
    txtmFinal.Text = Mid(lcTurnosTotales, 8, 5)
    txtTinicio.Text = Mid(lcTurnosTotales, 14, 5)
    txtTfinal.Text = Mid(lcTurnosTotales, 20, 5)
    
End Sub

Sub TurnosGrabar()
    SIGHEntidades.TurnosMananaTarde = "/" & txtMinicio.Text & "/" & txtmFinal.Text & "/" & txtTinicio.Text & "/" & txtTfinal.Text & "/"
End Sub


Private Sub optAmbos_Click(Value As Integer)
  If optAmbos.Value = True Then
    RaiseEvent SeModificoTurnos(IIf(optAmbos.Value = True, sghTurnos.sghTurnoAmbos, _
                               IIf(optManana.Value = True, sghTurnos.sghTurnoManana, sghTurnos.sghTurnoTarde)), _
                               txtMinicio.Text, txtmFinal.Text, txtTinicio.Text, txtTfinal.Text)
  End If
End Sub

Private Sub optManana_Click(Value As Integer)
   If optManana.Value = True Then
    RaiseEvent SeModificoTurnos(IIf(optAmbos.Value = True, sghTurnos.sghTurnoAmbos, _
                               IIf(optManana.Value = True, sghTurnos.sghTurnoManana, sghTurnos.sghTurnoTarde)), _
                               txtMinicio.Text, txtmFinal.Text, txtTinicio.Text, txtTfinal.Text)
   
   End If
End Sub

Private Sub optTarde_Click(Value As Integer)
   If optTarde.Value = True Then
    RaiseEvent SeModificoTurnos(IIf(optAmbos.Value = True, sghTurnos.sghTurnoAmbos, _
                               IIf(optManana.Value = True, sghTurnos.sghTurnoManana, sghTurnos.sghTurnoTarde)), _
                               txtMinicio.Text, txtmFinal.Text, txtTinicio.Text, txtTfinal.Text)
   End If
End Sub



Private Sub txtmFinal_LostFocus()
  If SIGHEntidades.EsHora(txtmFinal.Text) Then
    If CDate("01/01/2000 " & txtMinicio.Text) < CDate("01/01/2000 " & txtmFinal.Text) Then
      TurnosGrabar
      RaiseEvent SeModificoTurnos(IIf(optAmbos.Value = True, sghTurnos.sghTurnoAmbos, _
                               IIf(optManana.Value = True, sghTurnos.sghTurnoManana, sghTurnos.sghTurnoTarde)), _
                               txtMinicio.Text, txtmFinal.Text, txtTinicio.Text, txtTfinal.Text)
    Else
       txtmFinal.Text = "13:00"
    End If
  Else
    txtmFinal.Text = "13:00"
  End If
End Sub

Private Sub txtMinicio_LostFocus()
    If SIGHEntidades.EsHora(txtMinicio.Text) Then
       If CDate("01/01/2000 " & txtMinicio.Text) < CDate("01/01/2000 " & txtmFinal.Text) Then
         TurnosGrabar
         RaiseEvent SeModificoTurnos(IIf(optAmbos.Value = True, sghTurnos.sghTurnoAmbos, _
                               IIf(optManana.Value = True, sghTurnos.sghTurnoManana, sghTurnos.sghTurnoTarde)), _
                               txtMinicio.Text, txtmFinal.Text, txtTinicio.Text, txtTfinal.Text)
       Else
          txtMinicio.Text = "07:00"
       End If
    Else
        txtMinicio.Text = "07:00"
    End If
End Sub





Private Sub txtTfinal_LostFocus()
  If SIGHEntidades.EsHora(txtTfinal.Text) Then
    If CDate("01/01/2000 " & txtTinicio.Text) < CDate("01/01/2000 " & txtTfinal.Text) Then
      TurnosGrabar
      RaiseEvent SeModificoTurnos(IIf(optAmbos.Value = True, sghTurnos.sghTurnoAmbos, _
                               IIf(optManana.Value = True, sghTurnos.sghTurnoManana, sghTurnos.sghTurnoTarde)), _
                               txtMinicio.Text, txtmFinal.Text, txtTinicio.Text, txtTfinal.Text)
    Else
      txtTfinal.Text = "18:00"
    End If
  Else
    txtTfinal.Text = "18:00"
  End If
End Sub

Private Sub txtTinicio_LostFocus()
  If SIGHEntidades.EsHora(txtTinicio.Text) Then
    If CDate("01/01/2000 " & txtTinicio.Text) < CDate("01/01/2000 " & txtTfinal.Text) Then
      TurnosGrabar
      RaiseEvent SeModificoTurnos(IIf(optAmbos.Value = True, sghTurnos.sghTurnoAmbos, _
                               IIf(optManana.Value = True, sghTurnos.sghTurnoManana, sghTurnos.sghTurnoTarde)), _
                               txtMinicio.Text, txtmFinal.Text, txtTinicio.Text, txtTfinal.Text)
    Else
      txtTinicio.Text = "14:00"
    End If
  Else
    txtTinicio.Text = "14:00"
  End If
End Sub
