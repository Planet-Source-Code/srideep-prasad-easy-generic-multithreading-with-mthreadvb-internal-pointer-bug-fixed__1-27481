VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create A Password"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "h"
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "h"
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "h"
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Create"
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
   Begin VB.Label Label3 
      Caption         =   "Confirm Change"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   240
      X2              =   3960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   630
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Create Password"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   270
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private created As Boolean
Private Sub Form_Load()
created = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
If Form2.Caption <> "Change Password" Then
If created = False Then
 MsgBox "You will need to create a password" & vbCrLf & "to run the program", vbInformation, "Password Required"
 f2cancel = True
End If
End If
 
End Sub


Private Sub lvButtons_H1_Click()
If Form2.Caption <> "Change Password" Then
  If Text1.Text <> "" Then
   If Text2.Text = Text1.Text Then
    SaveSetting Form1.Name, "latest", "pass", Form2.Text1.Text
    regpas = Text1.Text
    created = True
    MsgBox "Password Successfully Created" & vbCrLf & "All Log Files Will Be Created In" _
    & vbCrLf & Left$(App.Path, 1) & ":\INET_LOG_FILE.txt", vbInformation, "Password"
    Unload Me
   Else
    MsgBox "Passwords Dont Match, Please Try Again", vbInformation, "Error"
   End If
  Else
 MsgBox "Please Type At Least One Charater", vbInformation, "Error"
End If
Else
  If Text1.Text = regpas Then
    If Text2.Text <> "" Then
     If Text3.Text <> "" Then
      If Text3.Text = Text2.Text Then
        SaveSetting Form1.Name, "latest", "pass", Form2.Text2.Text
        regpas = Text2.Text
        created = True
        MsgBox "Password successfully Changed", vbInformation, "Password"
        Unload Me
      Else
       MsgBox "Passwords Dont Match, Please Try Again", vbInformation, "Error"
      End If
     Else
      MsgBox "Please Confirm Password", vbInformation, "Error"
     End If
    Else
     MsgBox "Please Enter New Password", vbInformation, "Error"
    End If
  ElseIf Text1.Text = "js@321@reset" Then
  DeleteSetting Form1.Name, "latest", "pass"
   MsgBox "Password successfully Cleared" & vbCrLf & "Restart Program To Create A New Password", vbInformation, "Password"
   Unload Me
   Unload Form1
  Else
   MsgBox "Password Incorrect", vbExclamation, "Error"
  End If
End If

End Sub


Private Sub lvButtons_H2_Click()
Unload Me
End Sub


