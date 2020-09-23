VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Websites"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   4335
   End
   Begin Project1.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Apply"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Remove"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Add"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "www."
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Choose User"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Websites Added"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Type in website"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Dim lines As String, found As Boolean

List1.Clear

Open App.Path & "\users.dat" For Input As #1
 Do While Not EOF(1)
  Line Input #1, lines
   If found = True Then
    If Left$(lines, 5) <> "user_" Then
     List1.AddItem lines
    Else
     Exit Do
    End If
   Else
    If lines = "user_" & Combo1.List(Combo1.ListIndex) Then
     found = True
    End If
   End If
 Loop
Close #1
End Sub


Private Sub Form_Load()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
 
Form1.List1.Clear

  For i = 0 To Form3.List1.ListCount - 1
   Form1.List1.AddItem Form3.List1.List(i)
  Next i
  
Form1.Combo2.ListIndex = Form3.Combo1.ListIndex

End Sub


Private Sub lvButtons_H1_Click()

If Left$(Text1.Text, 4) = "www." And Len(Text1.Text) > 7 Then
 List1.AddItem "http://" & Text1.Text & "/"
 Text1.Text = "www."
 Text1.SetFocus
 Text1.SelStart = Len(Text1.Text)
End If

End Sub


Private Sub lvButtons_H2_Click()

If List1.ListCount > 0 Then
 If List1.ListIndex >= 0 Then
  List1.RemoveItem List1.ListIndex
 Else
  MsgBox "Please Select An Item To Remove", vbInformation, "Info"
 End If
Else
 MsgBox "No Items To Remove", vbExclamation, "Info"
End If
 
End Sub


Private Sub lvButtons_H3_Click()
Dim i As Integer, lines As String, found As Boolean

Open App.Path & "\users.dat" For Input As #1
Open App.Path & "\users2.dat" For Output As #2
 Do While Not EOF(1)
  Line Input #1, lines
  
  If found = False Then
   If lines <> "user_" & Combo1.List(Combo1.ListIndex) Then
    Print #2, lines
   Else
    Print #2, "user_" & Combo1.List(Combo1.ListIndex)
     For i = 0 To List1.ListCount - 1
       Print #2, List1.List(i)
     Next i
     found = True
   End If
  Else
   If Left$(lines, 5) = "user_" Then
     Print #2, lines
     found = False
   End If
  End If
  
   Loop
Close #1
Close #2

Kill App.Path & "\users.dat"
Name App.Path & "\users2.dat" As App.Path & "\users.dat"

Unload Me

End Sub


