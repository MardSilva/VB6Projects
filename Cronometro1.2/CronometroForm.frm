VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cron�metro 1.2 - Corporate"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTempoDecorrido 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   7
      Top             =   5040
      Width           =   4095
   End
   Begin VB.TextBox txtTempFinal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtTempInicial 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton btnParar 
      Caption         =   "Parar"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton btnIniciar 
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Tempo Decorrido (Total)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Tempo Final"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tempo Inicial"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'criando uma vari�vel global
Dim tempoInicial As Date

Private Sub btnIniciar_Click()

'instanciando a vari�vel tempoInicial na sub-rotina e definindo ela como now.
tempoInicial = Now

'atribuindo a vari�vel tempoInicial j� formatada no TextBox
txtTempInicial.Text = Format(tempoInicial, "hh:mm:ss")


End Sub

Private Sub btnParar_Click()

'criando a vari�vel para receber o tempo final, e definindo ela como now.
Dim tempoFinal As Date
tempoFinal = Now

'atribuindo a vari�vel tempoFinal j� formatada no TextBox
txtTempFinal.Text = Format(tempoFinal, "hh:mm:ss")

'agora fazendo o c�lculo da vari�vel tempoDecorrido para a diferen�a entre elas.
Dim tempoDecorrido As Date
tempoDecorrido = tempoFinal - tempoInicial

txtTempoDecorrido.Text = Format(tempoDecorrido, "hh:mm:ss")

'fazendo uma condi��o para exibir como segundos/minutos.
If (tempoDecorrido <= 59) Then
    txtTempoDecorrido.Text = Format(tempoDecorrido, "hh:mm:ss") & " segundos"

    ElseIf (tempoDecorrido >= 60) Then
    txtTempoDecorrido.Text = Format(tempoDecorrido, "hh:mm:ss") & " minutos"

End If

End Sub
