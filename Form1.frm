VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TextBox w/TipText and Effects"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Example 2"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Example 1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSearch 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label lblAdd 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   840
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Effect: Default, MouseMove and Clicked. Also 3D/Flat.  Click the form for the Default appearance to return."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   120
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'By Jim K on June, 2005
'Makes a nice effect to TextBox's
'I wrote it for use with single line TextBox's
'where i thought a TipText could look nice.
'A simple but nice effect for e.g. Search Field
'or other "Input" Fields.


'To hold values <False/True>
Dim ValTB1, ValTB2 As Boolean

Private Sub Command1_Click()

    lblSearch.Caption = "(" & Text1.Text & ") was not found"
    
End Sub

Private Sub Command2_Click()
    
    lblAdd.Caption = Text2.Text
    
End Sub

Private Sub Form_Load()
    
    'Setup controls
    SetTxtEffect Text1, &HEAF3F4, vbWhite, 0, " Insert a word or text to search for", Shape1, vbBlue
    SetTxtEffect Text2, &HEAF3F4, vbWhite, 1, " Type an e-mail address here...", Shape2, vbRed
    Command1.Top = Shape1.Top
    Command1.Height = Shape1.Height
    Command2.Top = Shape2.Top
    Command2.Height = Shape2.Height
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Rewrite propertys and set new values after a
    'TextBox is clicked
    Val = False
    ValTB1 = False
    ValTB2 = False
    SetTxtEffect Text1, &HEAF3F4, vbWhite, 0, " Insert a word or text to search for", Shape1, vbBlue
    SetTxtEffect Text2, &HEAF3F4, vbWhite, 1, " Type an e-mail address here...", Shape2, vbRed

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'MouseOver
    Shape1.Visible = ValTB1
    Shape2.Visible = ValTB2
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Form_MouseDown Button, Shift, X, Y
    
End Sub

Private Sub lblAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Form_MouseDown Button, Shift, X, Y
    
End Sub

Private Sub lblSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Form_MouseDown Button, Shift, X, Y
    
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Set new values and hold the clicked box's border
    'to show, and Box2 not to.
    Val = True
    ValTB1 = True
    ValTB2 = False
    SetTxtEffect Text1, &HEAF3F4, vbWhite, 0, "", Shape1, vbBlue
    
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Shape1.Visible = True
    
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Set new values
    Val = True
    ValTB1 = False
    ValTB2 = True
    SetTxtEffect Text2, &HEAF3F4, vbWhite, 0, "", Shape2, vbRed
    
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Shape2.Visible = True
    
End Sub
