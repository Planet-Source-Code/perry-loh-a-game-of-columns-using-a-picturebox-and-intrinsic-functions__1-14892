VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Height          =   1635
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3840
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Height          =   1275
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   3570
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblMain = "VeeBee Columns" & vbCrLf & " ~ Perry Loh ~ " & _
                vbCrLf & "skeevs@hotmail.com" & _
                vbCrLf & vbCrLf & "Arrow Keys - Move Columns" & _
                vbCrLf & "Space Bar - Rotate Columns"
End Sub
