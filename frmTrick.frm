VERSION 5.00
Begin VB.Form frmTrick 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Trick !"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command Buttons Forecolor"
      ForeColor       =   &H00400000&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "Command 1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Command 1"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   4
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Command 1"
         ForeColor       =   &H00008080&
         Height          =   495
         Index           =   3
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Command 1"
         ForeColor       =   &H00004000&
         Height          =   495
         Index           =   2
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Command 1"
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Command 1"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTrick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is just a little trick to mimic a commandbutton
'and have its forecolor changeable.

Private Sub Check1_Click(Index As Integer)
    'if you click on checkbox, it's value will change
    'to vbChecked and you must changed its value
    'back to vbUnchecked to make it look like
    'a Command Button.
    If Check1(Index).Value = vbChecked Then
        Check1(Index).Value = vbUnchecked
            'Put your code in here !
    End If
End Sub

'Note:
'     You can also used Option Button to
'     achieved the same effect !
