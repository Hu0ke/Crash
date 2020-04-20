VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "密码吧辅助破译工具v1.0"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   14520
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   4
      Left            =   12240
      TabIndex        =   16
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   3
      Left            =   9960
      TabIndex        =   15
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   2
      Left            =   7680
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   1
      Left            =   5400
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      ItemData        =   "Form1.frx":0006
      Left            =   3120
      List            =   "Form1.frx":0008
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   4
      Left            =   12240
      TabIndex        =   11
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Index           =   4
      Left            =   12240
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   3
      Left            =   9960
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Index           =   3
      Left            =   9960
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   2
      Left            =   7680
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Index           =   2
      Left            =   7680
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Index           =   1
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Txtin 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Form1.frx":000A
      Top             =   1560
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gftxt As Integer
Dim txtcheck As Boolean
Private Sub Combo1_Click(index As Integer)
Select Case index
Case 0
    If Txtin.Text <> "" And Combo1(0).Text <> "可用的方法" Then
        Command1(0).Enabled = True
    Else
        Command1(0).Enabled = False
    End If
Case Else
    If Combo1(index).Text <> "可用的方法" And Combo1(index).ListCount > 0 Then
        Command1(index).Enabled = True
    Else
        Command1(index).Enabled = False
    End If
End Select
End Sub

Private Sub Command1_Click(index As Integer)
Select Case index
Case 0 To 3
    If Combo1(index).Text = "B.摩斯码" Then Call 摩斯码(Txtin.Text, index)
    If Combo1(index).Text = "C.英文字母表数字互译" Then Call 英文字母表互译(Txtin.Text, index)
    Combo1(index + 1).Enabled = True
    Combo1(index + 1).Text = "可用的方法"
    'Call 智能检测(List1(index).Selected, index)
Case 4
    If Combo1(index).Text = "B.摩斯码" Then Call 摩斯码(Txtin.Text, index)
    If Combo1(index).Text = "C.英文字母表数字互译" Then Call 英文字母表互译(Txtin.Text, index)
    'Call 智能检测(List1(index).Selected, index)
End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
txtcheck = False
gftxt = 0
Text1.TabIndex = 0
Txtin.Text = "*单击输入*" & vbCrLf & _
"请输入密文（规格符合范围：数字，英文字母，摩斯码以及分割标点）"
Text1.Text = "支持方式：" & vbCrLf & "字母类：A.凯撒移位，B.摩斯码，C.字母表互译" & vbCrLf _
& "数字类：C.字母表互译，D.asc码，E.九宫格"
Text1.BackColor = vbGreen
For i = 0 To 4
    Command1(i).Enabled = False
    Command1(i).Caption = "第" & i & "层解密"
    If i = 0 Then Combo1(i).Text = "请输入密文" Else Combo1(i).Text = "请解密上一层"
    Combo1(i).Enabled = False
Next i
End Sub
Private Sub Txtin_Change()
Dim i As Integer
For i = 1 To Len(Txtin.Text)
    If (Mid(Txtin.Text, i, 1) > "a" And Mid(Txtin.Text, i, 1) < "z") _
    Or (Mid(Txtin.Text, i, 1) > "A" And Mid(Txtin.Text, i, 1) < "Z") _
    Or (Mid(Txtin.Text, i, 1) >= "0" And Mid(Txtin.Text, i, 1) <= "9") _
    Or (Mid(Txtin.Text, i, 1) = "." Or Mid(Txtin.Text, i, 1) = "*") _
    Or (Mid(Txtin.Text, i, 1) = "-" Or Mid(Txtin.Text, i, 1) = "/") Then
        txtcheck = True
    Else
        txtcheck = False
    End If
Next i
If Txtin.Text <> "" And txtcheck = True Then
    Combo1(0).Text = "可用的方法"
    Call 智能检测(Txtin.Text, -1)
    Combo1(0).Enabled = True
Else
    Combo1(0).Clear
    Combo1(0).Text = "请输入密文"
    Combo1(0).Enabled = False
End If
End Sub

Private Sub Txtin_GotFocus()
gftxt = gftxt + 1
If gftxt = 1 Then Txtin.Text = "" Else gftxt = 1
End Sub

Public Function 摩斯码(txt As String, index As Integer)
Dim a As String, txtout As String
Dim tou As Integer, wei As Integer
Dim i As Integer, j As Integer
tou = 1: wei = 0: txtout = "": a = ""
If Right(txt, 1) <> "/" Then txt = txt & "/"
For i = wei + 1 To Len(txt)
    If Mid(txt, i, 1) = "/" Then
    wei = i
    For j = tou To wei - 1
        a = a & Mid(txt, j, 1)
    Next j
    tou = wei + 1
        Select Case a
        Case ".-", "*-"
            txtout = txtout & "a"
        Case "-...", "-***"
            txtout = txtout & "b"
        Case "-.-.", "-*-*"
            txtout = txtout & "c"
        Case "-..", "-**"
            txtout = txtout & "d"
        Case ".", "*"
            txtout = txtout & "e"
        Case "..-.", "**-*"
            txtout = txtout & "f"
        Case "--.", "--*"
            txtout = txtout & "g"
        Case "....", "****"
            txtout = txtout & "h"
        Case "..", "**"
            txtout = txtout & "i"
        Case ".---", "*---"
            txtout = txtout & "j"
        Case "-.-", "-*-"
            txtout = txtout & "k"
        Case ".-..", "*-**"
            txtout = txtout & "l"
        Case "--"
            txtout = txtout & "m"
        Case "-.", "-*"
            txtout = txtout & "n"
        Case "---"
            txtout = txtout & "o"
        Case ".--.", "*--*"
            txtout = txtout & "p"
        Case "--.-", "--*-"
            txtout = txtout & "q"
        Case ".-.", "*-*"
            txtout = txtout & "r"
        Case "...", "***"
            txtout = txtout & "s"
        Case "-"
            txtout = txtout & "t"
        Case "..-", "**-"
            txtout = txtout & "u"
        Case "...-", "***-"
            txtout = txtout & "v"
        Case ".--", "*--"
            txtout = txtout & "w"
        Case "-..-", "-**-"
            txtout = txtout & "x"
        Case "-.--", "-*--"
            txtout = txtout & "y"
        Case "--..", "--**"
            txtout = txtout & "z"
        Case "-----", "-----": txtout = txtout & "0"
        Case ".----", "*----": txtout = txtout & "1"
        Case "..---", "**---": txtout = txtout & "2"
        Case "...--", "***--": txtout = txtout & "3"
        Case "....-", "****-": txtout = txtout & "4"
        Case ".....", "*****": txtout = txtout & "5"
        Case "-....", "-****": txtout = txtout & "6"
        Case "--...", "--***": txtout = txtout & "7"
        Case "---..", "---**": txtout = txtout & "8"
        Case "----.", "----*": txtout = txtout & "9"
        Case ".-.-.-", "*-*-*-"
            txtout = txtout & "."
        Case "---...", "---***"
            txtout = txtout & ":"
        Case "--..--", "--**--"
            txtout = txtout & ","
        Case "-.-.-.", "-*-*-*"
            txtout = txtout & ";"
        Case "..--..", "**--**"
            txtout = txtout & "?"
        Case "-...-", "-***-"
            txtout = txtout & "="
        Case ".----.", "*----*"
            txtout = txtout & """"
        Case "-..-.", "-**-*"
            txtout = txtout & "/"
        Case "-.-.--", "-*-*--"
            txtout = txtout & "!"
        Case "-....-", "-****-"
            txtout = txtout & "-"
        Case "..--.-", "**--*-"
            txtout = txtout & "_"
        Case ".-..", "*-**"
            txtout = txtout & "”"
        Case "-.--.", "-*--*"
            txtout = txtout & "("
        Case "-.--.-", "-*--*-"
            txtout = txtout & ")"
        Case "...-..-", "***-**-"
            txtout = txtout & "$"
        Case "....", "****"
            txtout = txtout & "&"
        Case ".--.-.", "*--*-*"
            txtout = txtout & "@"
        Case ".-.-.", "*-*-*"
            txtout = txtout & "+"
        Case Else
            txtout = txtout & "无此码"
        End Select
        a = ""
    End If
    If i = Len(txt) Then Exit For
Next i
摩斯码 = txtout
List1(index).AddItem 摩斯码
End Function

Public Function 英文字母表互译(txt As String, index As Integer)
Dim zmb(1 To 26) As String, sz(1 To 26) As String, txtout As String
Dim a As Variant, k As Integer, i As Integer
a = "": txtout = ""
For i = 1 To 26
    zmb(i) = Chr(96 + i)
    If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 6 Or i = 7 Or i = 8 Or i = 9 Then
        sz(i) = "0" & i
    Else
        sz(i) = i
    End If
Next i

For i = 1 To Len(txt)
    If IsNumeric(Mid(txt, i, 1)) = True Then
        a = Mid(txt, i, 2): i = i + 1
        For k = 1 To 26
            If a = sz(k) Then a = zmb(k)
        Next k
    ElseIf Asc(Mid(txt, i, 1)) >= 65 And Asc(Mid(txt, i, 1)) <= 91 Then
        a = Chr(Asc(Mid(txt, i, 1)) + 32)
        For k = 1 To 26
            If a = zmb(k) Then a = sz(k)
        Next k
    ElseIf Asc(Mid(txt, i, 1)) >= 97 And Asc(Mid(txt, i, 1)) <= 123 Then
        a = Mid(txt, i, 1)
        For k = 1 To 26
            If a = zmb(k) Then a = sz(k)
        Next k
    Else
        a = Mid(txt, i, 1)
    End If
    txtout = txtout & a
Next i
英文字母表互译 = txtout
List1(index).AddItem 英文字母表互译
End Function

Public Sub 智能检测(txt As String, index As Integer)
Dim abcde(1 To 5) As Boolean, i As Integer, p As Integer
For i = 1 To 5
    abcde(i) = False
Next i

If IsNumeric(txt) = True Then
    abcde(3) = True: abcde(4) = True: abcde(5) = True
ElseIf Mid(txt, 1, 1) = "." Or Mid(txt, 1, 1) = "*" Or Mid(txt, 1, 1) = "-" Or Mid(txt, 1, 1) = "/" Then
    p = 1
    For i = 1 To Len(txt)
        If Mid(txt, i, 1) <> "." And Mid(txt, i, 1) <> "*" _
        And Mid(txt, i, 1) <> "-" And Mid(txt, i, 1) <> "/" Then p = 0
    Next i
    If p = 1 Then abcde(2) = True
Else
    abcde(1) = True: abcde(3) = True
End If

Combo1(index + 1).Clear
Combo1(index + 1).Text = "可用的方法"
For i = 1 To 5
    If abcde(i) = True Then
        Select Case i
        Case 1: Combo1(index + 1).AddItem "A.凯撒移位"
        Case 2: Combo1(index + 1).AddItem "B.摩斯码"
        Case 3: Combo1(index + 1).AddItem "C.英文字母表数字互译"
        Case 4: Combo1(index + 1).AddItem "D.asc码"
        Case 5: Combo1(index + 1).AddItem "E.九宫格"
        End Select
    End If
Next i
End Sub
