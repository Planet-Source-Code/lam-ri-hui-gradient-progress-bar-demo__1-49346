VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gradient Progress Bar Demo"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin Gradient_Progress_Bar.PB_Yellow PB_Yellow1 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
   End
   Begin Gradient_Progress_Bar.PB_Red PB_Red1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
   End
   Begin Gradient_Progress_Bar.PB_Purple PB_Purple1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
   End
   Begin Gradient_Progress_Bar.PB_Grey PB_Grey1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
   End
   Begin Gradient_Progress_Bar.PB_Green PB_Green1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
   End
   Begin Gradient_Progress_Bar.PB_Blue PB_Blue1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Click Here"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gradient Progress Bar Demo by Lam Ri Hui
'Demonstrates how to use custom made
'progress bar :
'1. Blue Gradient Progress Bar
'2. Green Gradient Progress Bar
'3. Grey Gradient Progress Bar
'4. Purple Gradient Progress Bar
'5. Red Gradient Progress Bar
'6. Yellow Gradient Progress Bar
'If you like this program,
'vote it at www.planetsourcecode.com/vb/
'Don't forget to leave comments when you vote.

Dim counter
Dim j
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit
Private Sub Command1_Click()
Blue
Green
Grey
Purple
Red
Yellow
End Sub

Public Sub Blue()
'Blue gradient progress bar
Dim WorkBlue(1 To 10000) As String
PB_Blue1.Max = 100
For j = 0 To 100
PB_Blue1.Value = j
For counter = LBound(WorkBlue) To UBound(WorkBlue)
WorkBlue(counter) = Space(10)
DoEvents
Next counter
PB_Blue1.Caption = j / 100 * 100 & "%"
Next j
End Sub

Public Sub Green()
'Green gradient progress bar
Dim WorkGreen(1 To 8000) As String
PB_Green1.Max = 100
For j = 0 To 100
PB_Green1.Value = j
For counter = LBound(WorkGreen) To UBound(WorkGreen)
WorkGreen(counter) = Space(10)
DoEvents
Next counter
PB_Green1.Caption = j / 100 * 100 & "%"
Next j
End Sub

Public Sub Grey()
'Grey gradient progress bar
Dim WorkGrey(1 To 6000) As String
PB_Grey1.Max = 100
For j = 0 To 100
PB_Grey1.Value = j
For counter = LBound(WorkGrey) To UBound(WorkGrey)
WorkGrey(counter) = Space(10)
DoEvents
Next counter
PB_Grey1.Caption = j / 100 * 100 & "%"
Next j
End Sub

Public Sub Purple()
'Purple gradient progress bar
Dim WorkPurple(1 To 4000) As String
PB_Purple1.Max = 100
For j = 0 To 100
PB_Purple1.Value = j
For counter = LBound(WorkPurple) To UBound(WorkPurple)
WorkPurple(counter) = Space(10)
DoEvents
Next counter
PB_Purple1.Caption = j / 100 * 100 & "%"
Next j
End Sub

Public Sub Red()
'Red gradient progress bar
Dim WorkRed(1 To 2000) As String
PB_Red1.Max = 100
For j = 0 To 100
PB_Red1.Value = j
For counter = LBound(WorkRed) To UBound(WorkRed)
WorkRed(counter) = Space(10)
DoEvents
Next counter
PB_Red1.Caption = j / 100 * 100 & "%"
Next j
End Sub

Public Sub Yellow()
'Yellow gradient progress bar
Dim WorkYellow(1 To 1000) As String
PB_Yellow1.Max = 100
For j = 0 To 100
PB_Yellow1.Value = j
For counter = LBound(WorkYellow) To UBound(WorkYellow)
WorkYellow(counter) = Space(10)
DoEvents
Next counter
PB_Yellow1.Caption = j / 100 * 100 & "%"
Next j
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim response
Dim dummy
response = MsgBox("Do you want to vote for this program?", vbYesNo, "Vote")
If response = vbYes Then
dummy = ShellExecute(1, "Open", "http://www.planetsourcecode.com/vb/", 0&, 0&, 10): MsgBox "Thanks for spending time to vote my program.", , "Thanks"
Else
MsgBox "Then, please give comments about this program.", , "Comments"
dummy = ShellExecute(1, "Open", "http://www.planetsourcecode.com/vb/", 0&, 0&, 10)
End If
End Sub
