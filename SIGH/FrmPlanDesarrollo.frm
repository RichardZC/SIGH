VERSION 5.00
Begin VB.Form frmPlanDesarrollo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desarrollo Psicomotor"
   ClientHeight    =   7635
   ClientLeft      =   4650
   ClientTop       =   2055
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnImprime 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Picture         =   "FrmPlanDesarrollo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   7080
      Width           =   1005
   End
   Begin VB.VScrollBar vsForm 
      Height          =   7095
      Left            =   10320
      TabIndex        =   146
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7095
      ScaleWidth      =   10215
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.PictureBox picDetail 
         BorderStyle     =   0  'None
         Height          =   9975
         Left            =   0
         Picture         =   "FrmPlanDesarrollo.frx":04D9
         ScaleHeight     =   9975
         ScaleWidth      =   10335
         TabIndex        =   1
         Top             =   15
         Width           =   10335
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   4200
            TabIndex        =   145
            Text            =   "x"
            Top             =   250
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   4530
            TabIndex        =   144
            Text            =   "x"
            Top             =   250
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   4200
            TabIndex        =   143
            Text            =   "x"
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   4530
            TabIndex        =   142
            Text            =   "x"
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   4200
            TabIndex        =   141
            Text            =   "x"
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   4530
            TabIndex        =   140
            Text            =   "x"
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   4200
            TabIndex        =   139
            Text            =   "x"
            Top             =   920
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   4530
            TabIndex        =   138
            Text            =   "x"
            Top             =   920
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   7
            Left            =   4200
            TabIndex        =   137
            Text            =   "x"
            Top             =   1560
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   7
            Left            =   4485
            TabIndex        =   136
            Text            =   "x"
            Top             =   1575
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   8
            Left            =   4200
            TabIndex        =   135
            Text            =   "x"
            Top             =   1800
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   8
            Left            =   4530
            TabIndex        =   134
            Text            =   "x"
            Top             =   1800
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   5
            Left            =   4200
            TabIndex        =   133
            Text            =   "x"
            Top             =   1160
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   5
            Left            =   4530
            TabIndex        =   132
            Text            =   "x"
            Top             =   1160
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   6
            Left            =   4200
            TabIndex        =   131
            Text            =   "x"
            Top             =   1360
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   6
            Left            =   4530
            TabIndex        =   130
            Text            =   "x"
            Top             =   1360
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   9
            Left            =   4200
            TabIndex        =   129
            Text            =   "x"
            Top             =   2020
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   9
            Left            =   4530
            TabIndex        =   128
            Text            =   "x"
            Top             =   2020
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   10
            Left            =   4200
            TabIndex        =   127
            Text            =   "x"
            Top             =   2240
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   10
            Left            =   4530
            TabIndex        =   126
            Text            =   "x"
            Top             =   2240
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   11
            Left            =   4200
            TabIndex        =   125
            Text            =   "x"
            Top             =   2460
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   11
            Left            =   4530
            TabIndex        =   124
            Text            =   "x"
            Top             =   2460
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   12
            Left            =   4200
            TabIndex        =   123
            Text            =   "x"
            Top             =   2660
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   12
            Left            =   4530
            TabIndex        =   122
            Text            =   "x"
            Top             =   2660
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   13
            Left            =   4200
            TabIndex        =   121
            Text            =   "x"
            Top             =   2880
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   13
            Left            =   4530
            TabIndex        =   120
            Text            =   "x"
            Top             =   2880
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   14
            Left            =   4200
            TabIndex        =   119
            Text            =   "x"
            Top             =   3120
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   14
            Left            =   4530
            TabIndex        =   118
            Text            =   "x"
            Top             =   3120
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   37
            Left            =   8440
            TabIndex        =   117
            Text            =   "x"
            Top             =   260
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   37
            Left            =   8770
            TabIndex        =   116
            Text            =   "x"
            Top             =   260
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   38
            Left            =   8440
            TabIndex        =   115
            Text            =   "x"
            Top             =   465
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   38
            Left            =   8770
            TabIndex        =   114
            Text            =   "x"
            Top             =   465
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   39
            Left            =   8440
            TabIndex        =   113
            Text            =   "x"
            Top             =   705
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   39
            Left            =   8770
            TabIndex        =   112
            Text            =   "x"
            Top             =   705
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   15
            Left            =   4200
            TabIndex        =   111
            Text            =   "x"
            Top             =   3760
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   15
            Left            =   4530
            TabIndex        =   110
            Text            =   "x"
            Top             =   3760
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   16
            Left            =   4200
            TabIndex        =   109
            Text            =   "x"
            Top             =   3975
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   16
            Left            =   4530
            TabIndex        =   108
            Text            =   "x"
            Top             =   3975
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   17
            Left            =   4200
            TabIndex        =   107
            Text            =   "x"
            Top             =   4185
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   17
            Left            =   4530
            TabIndex        =   106
            Text            =   "x"
            Top             =   4185
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   18
            Left            =   4200
            TabIndex        =   105
            Text            =   "x"
            Top             =   4405
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   18
            Left            =   4530
            TabIndex        =   104
            Text            =   "x"
            Top             =   4405
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   21
            Left            =   4200
            TabIndex        =   103
            Text            =   "x"
            Top             =   5065
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   21
            Left            =   4530
            TabIndex        =   102
            Text            =   "x"
            Top             =   5065
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   19
            Left            =   4200
            TabIndex        =   101
            Text            =   "x"
            Top             =   4635
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   19
            Left            =   4530
            TabIndex        =   100
            Text            =   "x"
            Top             =   4635
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   20
            Left            =   4200
            TabIndex        =   99
            Text            =   "x"
            Top             =   4860
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   20
            Left            =   4530
            TabIndex        =   98
            Text            =   "x"
            Top             =   4860
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   42
            Left            =   8440
            TabIndex        =   97
            Text            =   "x"
            Top             =   2020
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   42
            Left            =   8770
            TabIndex        =   96
            Text            =   "x"
            Top             =   2020
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   43
            Left            =   8440
            TabIndex        =   95
            Text            =   "x"
            Top             =   2250
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   43
            Left            =   8770
            TabIndex        =   94
            Text            =   "x"
            Top             =   2250
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   44
            Left            =   8440
            TabIndex        =   93
            Text            =   "x"
            Top             =   2440
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   44
            Left            =   8770
            TabIndex        =   92
            Text            =   "x"
            Top             =   2440
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   45
            Left            =   8440
            TabIndex        =   91
            Text            =   "x"
            Top             =   2660
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   45
            Left            =   8770
            TabIndex        =   90
            Text            =   "x"
            Top             =   2660
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   46
            Left            =   8440
            TabIndex        =   89
            Text            =   "x"
            Top             =   2885
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   46
            Left            =   8770
            TabIndex        =   88
            Text            =   "x"
            Top             =   2885
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   47
            Left            =   8440
            TabIndex        =   87
            Text            =   "x"
            Top             =   3105
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   47
            Left            =   8770
            TabIndex        =   86
            Text            =   "x"
            Top             =   3105
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   48
            Left            =   8440
            TabIndex        =   85
            Text            =   "x"
            Top             =   3315
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   48
            Left            =   8770
            TabIndex        =   84
            Text            =   "x"
            Top             =   3315
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   51
            Left            =   8440
            TabIndex        =   83
            Text            =   "x"
            Top             =   3985
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   51
            Left            =   8770
            TabIndex        =   82
            Text            =   "x"
            Top             =   3985
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   49
            Left            =   8440
            TabIndex        =   81
            Text            =   "x"
            Top             =   3555
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   49
            Left            =   8770
            TabIndex        =   80
            Text            =   "x"
            Top             =   3555
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   50
            Left            =   8440
            TabIndex        =   79
            Text            =   "x"
            Top             =   3750
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   50
            Left            =   8770
            TabIndex        =   78
            Text            =   "x"
            Top             =   3750
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   40
            Left            =   8440
            TabIndex        =   77
            Text            =   "x"
            Top             =   1120
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   40
            Left            =   8770
            TabIndex        =   76
            Text            =   "x"
            Top             =   1120
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   41
            Left            =   8440
            TabIndex        =   75
            Text            =   "x"
            Top             =   1360
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   41
            Left            =   8770
            TabIndex        =   74
            Text            =   "x"
            Top             =   1360
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   52
            Left            =   8440
            TabIndex        =   73
            Text            =   "x"
            Top             =   4200
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   52
            Left            =   8770
            TabIndex        =   72
            Text            =   "x"
            Top             =   4200
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   24
            Left            =   4200
            TabIndex        =   71
            Text            =   "x"
            Top             =   5735
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   24
            Left            =   4530
            TabIndex        =   70
            Text            =   "x"
            Top             =   5735
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   22
            Left            =   4200
            TabIndex        =   69
            Text            =   "x"
            Top             =   5280
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   22
            Left            =   4530
            TabIndex        =   68
            Text            =   "x"
            Top             =   5280
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   23
            Left            =   4200
            TabIndex        =   67
            Text            =   "x"
            Top             =   5505
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   23
            Left            =   4530
            TabIndex        =   66
            Text            =   "x"
            Top             =   5505
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   55
            Left            =   8440
            TabIndex        =   65
            Text            =   "x"
            Top             =   5080
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   55
            Left            =   8770
            TabIndex        =   64
            Text            =   "x"
            Top             =   5080
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   53
            Left            =   8440
            TabIndex        =   63
            Text            =   "x"
            Top             =   4640
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   53
            Left            =   8770
            TabIndex        =   62
            Text            =   "x"
            Top             =   4640
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   54
            Left            =   8440
            TabIndex        =   61
            Text            =   "x"
            Top             =   4860
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   54
            Left            =   8770
            TabIndex        =   60
            Text            =   "x"
            Top             =   4860
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   27
            Left            =   4200
            TabIndex        =   59
            Text            =   "x"
            Top             =   6820
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   27
            Left            =   4530
            TabIndex        =   58
            Text            =   "x"
            Top             =   6820
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   25
            Left            =   4200
            TabIndex        =   57
            Text            =   "x"
            Top             =   6390
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   25
            Left            =   4530
            TabIndex        =   56
            Text            =   "x"
            Top             =   6390
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   26
            Left            =   4200
            TabIndex        =   55
            Text            =   "x"
            Top             =   6600
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   26
            Left            =   4530
            TabIndex        =   54
            Text            =   "x"
            Top             =   6600
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   32
            Left            =   4200
            TabIndex        =   53
            Text            =   "x"
            Top             =   8595
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   32
            Left            =   4530
            TabIndex        =   52
            Text            =   "x"
            Top             =   8595
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   30
            Left            =   4200
            TabIndex        =   51
            Text            =   "x"
            Top             =   8160
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   30
            Left            =   4530
            TabIndex        =   50
            Text            =   "x"
            Top             =   8160
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   31
            Left            =   4200
            TabIndex        =   49
            Text            =   "x"
            Top             =   8365
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   31
            Left            =   4530
            TabIndex        =   48
            Text            =   "x"
            Top             =   8365
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   35
            Left            =   4200
            TabIndex        =   47
            Text            =   "x"
            Top             =   9455
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   35
            Left            =   4530
            TabIndex        =   46
            Text            =   "x"
            Top             =   9455
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   33
            Left            =   4200
            TabIndex        =   45
            Text            =   "x"
            Top             =   9000
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   33
            Left            =   4530
            TabIndex        =   44
            Text            =   "x"
            Top             =   9000
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   34
            Left            =   4200
            TabIndex        =   43
            Text            =   "x"
            Top             =   9225
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   34
            Left            =   4530
            TabIndex        =   42
            Text            =   "x"
            Top             =   9225
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   36
            Left            =   4200
            TabIndex        =   41
            Text            =   "x"
            Top             =   9675
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   36
            Left            =   4530
            TabIndex        =   40
            Text            =   "x"
            Top             =   9675
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   28
            Left            =   4200
            TabIndex        =   39
            Text            =   "x"
            Top             =   7260
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   28
            Left            =   4530
            TabIndex        =   38
            Text            =   "x"
            Top             =   7260
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   29
            Left            =   4200
            TabIndex        =   37
            Text            =   "x"
            Top             =   7485
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   29
            Left            =   4530
            TabIndex        =   36
            Text            =   "x"
            Top             =   7485
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   56
            Left            =   8440
            TabIndex        =   35
            Text            =   "x"
            Top             =   5520
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   56
            Left            =   8770
            TabIndex        =   34
            Text            =   "x"
            Top             =   5520
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   57
            Left            =   8440
            TabIndex        =   33
            Text            =   "x"
            Top             =   5730
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   57
            Left            =   8770
            TabIndex        =   32
            Text            =   "x"
            Top             =   5730
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   58
            Left            =   8440
            TabIndex        =   31
            Text            =   "x"
            Top             =   5940
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   58
            Left            =   8770
            TabIndex        =   30
            Text            =   "x"
            Top             =   5940
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   59
            Left            =   8440
            TabIndex        =   29
            Text            =   "x"
            Top             =   6165
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   59
            Left            =   8770
            TabIndex        =   28
            Text            =   "x"
            Top             =   6165
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   62
            Left            =   8440
            TabIndex        =   27
            Text            =   "x"
            Top             =   6825
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   62
            Left            =   8770
            TabIndex        =   26
            Text            =   "x"
            Top             =   6825
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   60
            Left            =   8440
            TabIndex        =   25
            Text            =   "x"
            Top             =   6390
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   60
            Left            =   8770
            TabIndex        =   24
            Text            =   "x"
            Top             =   6390
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   61
            Left            =   8440
            TabIndex        =   23
            Text            =   "x"
            Top             =   6615
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   61
            Left            =   8770
            TabIndex        =   22
            Text            =   "x"
            Top             =   6615
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   65
            Left            =   8440
            TabIndex        =   21
            Text            =   "x"
            Top             =   7470
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   65
            Left            =   8770
            TabIndex        =   20
            Text            =   "x"
            Top             =   7470
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   63
            Left            =   8440
            TabIndex        =   19
            Text            =   "x"
            Top             =   7035
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   63
            Left            =   8770
            TabIndex        =   18
            Text            =   "x"
            Top             =   7035
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   64
            Left            =   8440
            TabIndex        =   17
            Text            =   "x"
            Top             =   7260
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   64
            Left            =   8770
            TabIndex        =   16
            Text            =   "x"
            Top             =   7260
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   66
            Left            =   8440
            TabIndex        =   15
            Text            =   "x"
            Top             =   7700
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   66
            Left            =   8770
            TabIndex        =   14
            Text            =   "x"
            Top             =   7700
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   67
            Left            =   8440
            TabIndex        =   13
            Text            =   "x"
            Top             =   7910
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   67
            Left            =   8770
            TabIndex        =   12
            Text            =   "x"
            Top             =   7910
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   68
            Left            =   8440
            TabIndex        =   11
            Text            =   "x"
            Top             =   8130
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   68
            Left            =   8770
            TabIndex        =   10
            Text            =   "x"
            Top             =   8130
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   69
            Left            =   8440
            TabIndex        =   9
            Text            =   "x"
            Top             =   8355
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   69
            Left            =   8770
            TabIndex        =   8
            Text            =   "x"
            Top             =   8355
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   72
            Left            =   8440
            TabIndex        =   7
            Text            =   "x"
            Top             =   9020
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   72
            Left            =   8770
            TabIndex        =   6
            Text            =   "x"
            Top             =   9020
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   70
            Left            =   8440
            TabIndex        =   5
            Text            =   "x"
            Top             =   8560
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   70
            Left            =   8770
            TabIndex        =   4
            Text            =   "x"
            Top             =   8560
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloSi 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   71
            Left            =   8440
            TabIndex        =   3
            Text            =   "x"
            Top             =   8795
            Width           =   255
         End
         Begin VB.TextBox itemDesarrolloNo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   71
            Left            =   8770
            TabIndex        =   2
            Text            =   "x"
            Top             =   8795
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "frmPlanDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Muestra el Plan de Desarrollo de un niño
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Dim mo_rsPlanDesarrollo As New ADODB.Recordset

Property Set rsDesarrollo(lValue As ADODB.Recordset)
   Set mo_rsPlanDesarrollo = lValue
End Property

Private Sub btnImprime_Click()
'mgaray201411b
On Error GoTo miError
    picDetail.Picture = picDetail.Image
    'mgaray20141013
    Printer.Orientation = 2
    Printer.PaintPicture picDetail.Picture, 250, 250, Printer.Width - 500, Printer.Height - 500
    Printer.EndDoc
miError:
    If Err Then
        If Err.Number <> 482 Then
            MsgBox Err.Number & " : " & Err.Description, vbCritical, "Error de Impresión"
        End If
    End If
End Sub

Private Sub ocultarControles()
    Dim i As Integer
    For i = 1 To itemDesarrolloSi.Count
        itemDesarrolloSi(i).Visible = False
        itemDesarrolloNo(i).Visible = False
    Next i
End Sub

Private Sub Form_Activate()
    Inicializar
End Sub

Private Sub Form_Initialize()
    vsForm.Top = 0
    picDetail.Top = 0
    picDetail.Left = 0
    vsForm.Top = picContainer.Top
    vsForm.Height = picContainer.Height
    
    vsForm.Max = picContainer.Height - picContainer.Top
    vsForm.SmallChange = Abs(vsForm.Max \ 16) + 1
    vsForm.LargeChange = Abs(vsForm.Max \ 4) + 1
End Sub

Private Sub Form_Load()
    ocultarControles
End Sub

Private Sub Inicializar()
On Error GoTo miError
    If Not (mo_rsPlanDesarrollo Is Nothing) Then
        If Not (mo_rsPlanDesarrollo.BOF = True And mo_rsPlanDesarrollo.EOF = True) Then
            mo_rsPlanDesarrollo.MoveFirst
            picDetail.AutoRedraw = True
            While mo_rsPlanDesarrollo.EOF = False
                If Not IsNull(mo_rsPlanDesarrollo!EjecutaAccion) Then
                    If mo_rsPlanDesarrollo!EjecutaAccion = True Then
'                        itemDesarrolloSi(mo_rsPlanDesarrollo!ItemDesarrollo).Visible = True
                        picDetail.CurrentX = itemDesarrolloSi(mo_rsPlanDesarrollo!IdItemDesarrollo).Left
                        picDetail.CurrentY = itemDesarrolloSi(mo_rsPlanDesarrollo!IdItemDesarrollo).Top
                        picDetail.Print "x"
                    Else
'                        itemDesarrolloNo(mo_rsPlanDesarrollo!ItemDesarrollo).Visible = True
                        picDetail.CurrentX = itemDesarrolloNo(mo_rsPlanDesarrollo!IdItemDesarrollo).Left
                        picDetail.CurrentY = itemDesarrolloNo(mo_rsPlanDesarrollo!IdItemDesarrollo).Top
                        picDetail.Print "x"
                    End If
                End If
                mo_rsPlanDesarrollo.MoveNext
                
                picDetail.CurrentX = 0
                picDetail.CurrentY = 0
                
            Wend
            picDetail.AutoRedraw = False
        End If
    End If
miError:
    If Err Then
        MsgBox Err.Description, vbInformation, "Mensaje"
    End If
End Sub

Private Sub vsForm_Change()
    picDetail.Top = -vsForm.Value
End Sub
