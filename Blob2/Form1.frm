VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Bouton"
      Height          =   525
      Left            =   3330
      TabIndex        =   2
      Top             =   1890
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bouton"
      Height          =   525
      Left            =   3330
      TabIndex        =   1
      Top             =   690
      Width           =   1245
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   3165
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5583
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
UserControl11.GetFileFromDB "c:\customer.mdb", "Select * from customers where Name Like 'Adam'", "Picture", "c:\test.jpg"

End Sub

Private Sub Command2_Click()
UserControl11.PutFileToDB "c:\customer.mdb", "Select * from customers where Name Like 'Adam'", "Picture", "c:\adam.jpg"

End Sub


