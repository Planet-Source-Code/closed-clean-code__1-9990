VERSION 5.00
Begin VB.Form CleanCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Variable Declarations"
   ClientHeight    =   4320
   ClientLeft      =   2175
   ClientTop       =   1890
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3870
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3870
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "The following variables have been declared, but are not used anywhere in the code."
      Height          =   465
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   4395
   End
End
Attribute VB_Name = "CleanCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub Command1_Click()
    Dim Varname As String
    Dim Linenum As Long
    If List1.text = "" Then Exit Sub
    Varname = Trim(Mid(List1.text, 1, 25))
    Linenum = CInt(Trim(Mid(List1.text, 32)))
    If InStr(1, Connect.VBInstance.ActiveCodePane.CodeModule.Lines(Linenum, 1), ", " & Varname) > 0 Then
        Connect.VBInstance.ActiveCodePane.CodeModule.ReplaceLine Linenum, Replace(Connect.VBInstance.ActiveCodePane.CodeModule.Lines(Linenum, 1), ", " & Varname, "")
    ElseIf InStr(1, Connect.VBInstance.ActiveCodePane.CodeModule.Lines(Linenum, 1), Varname & ", ") > 0 Then
        Connect.VBInstance.ActiveCodePane.CodeModule.ReplaceLine Linenum, Replace(Connect.VBInstance.ActiveCodePane.CodeModule.Lines(Linenum, 1), Varname & ", ", "")
    ElseIf InStr(1, Connect.VBInstance.ActiveCodePane.CodeModule.Lines(Linenum, 1), Varname) > 0 Then
        Connect.VBInstance.ActiveCodePane.CodeModule.ReplaceLine Linenum, Replace(Connect.VBInstance.ActiveCodePane.CodeModule.Lines(Linenum, 1), Varname, "")
    End If
    If LCase(Connect.VBInstance.ActiveCodePane.CodeModule.Lines(Linenum, 1)) = "dim" Then Connect.VBInstance.ActiveCodePane.CodeModule.DeleteLines Linenum
    List1.Clear
    Call Form_Load
End Sub

Private Sub Form_Load()
    On Error GoTo bad
    Dim count As Long
    Dim tempstr As String
    Dim x As Long
    Dim CurrLine As String
    Dim DeclareEnd As Integer
    Dim Declares As New Collection
    Dim LineCounts As New Collection
    For count = 1 To Connect.VBInstance.ActiveCodePane.CodeModule.CountOfLines
        CurrLine = Connect.VBInstance.ActiveCodePane.CodeModule.Lines(count, 1)
        If Mid(LCase(Trim(CurrLine)), 1, 4) = "dim " Then
            tempstr = Mid(Trim(CurrLine), 5)
            For x = 1 To counter(tempstr, ", ")
                Declares.Add Sepinfo(tempstr, ", ", x - 1)
                LineCounts.Add count
            Next
        End If
    Next
    For count = 1 To Connect.VBInstance.ActiveCodePane.CodeModule.CountOfLines
        CurrLine = Connect.VBInstance.ActiveCodePane.CodeModule.Lines(count, 1)
        If Mid(LCase(Trim(CurrLine)), 1, 4) <> "dim " And CurrLine <> "" Then
            x = 1
            Do While x <= Declares.count
                If IsVariableIn(CurrLine, RemoveArray(Mid(Declares.Item(x), 1, InStr(1, Declares.Item(x), " ") - 1))) Then
                    Declares.Remove x
                    LineCounts.Remove x
                    x = x - 1
                End If
                x = x + 1
            Loop
        End If
    Next
bad:
    For count = 1 To Declares.count
        List1.AddItem Printable(Declares.Item(count), 25) & "Line: " & LineCounts.Item(count)
    Next
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Connect.Hide
    Unload Me
End Sub

Private Sub OKButton_Click()
    Connect.Hide
    Unload Me
End Sub

Private Function IsVariableIn(CodeLine As String, Varname As String) As Boolean
    Dim TimesFound As Integer
    Dim count As Integer
    Dim CurrLoc As Integer
    TimesFound = counter(CodeLine, Varname) - 1
    IsVariableIn = False
    CurrLoc = 0
    For count = 1 To TimesFound
        IsVariableIn = True
        CurrLoc = InStr(CurrLoc + 1, CodeLine, Varname)
        If CurrLoc < 1 Then
            IsVariableIn = False
            Exit Function
        End If
        If CurrLoc > 1 Then
            If IsAcceptable(Mid(CodeLine, CurrLoc - 1, 1)) = True Then IsVariableIn = False
        End If
        If Len(Varname) + CurrLoc < Len(CodeLine) Then
            If IsAcceptable(Mid(CodeLine, CurrLoc + Len(Varname), 1)) = True Then IsVariableIn = False
        End If
        If IsVariableIn = True Then Exit Function
    Next

End Function

Private Function IsAcceptable(AChar As String) As Boolean
    Dim AscCode As Integer
    AscCode = Asc(LCase(AChar))
    If (AscCode < Asc("a") Or AscCode > Asc("z")) And (AscCode < Asc("0") Or AscCode > Asc("9")) And AChar <> "_" Then
    Else
        IsAcceptable = True
    End If
End Function

Private Function Sepinfo(text As String, Key As String, Index As Long) As String
    Dim TempArr() As String
    ReDim TempArr(counter(text, Key))
    TempArr = Split(text, Key)
    Sepinfo = TempArr(Index)
End Function


Private Function counter(text As String, Key As String) As Integer
    Dim count As Long
    Dim total As Integer
    On Error GoTo bad
    For count = 1 To Len(text)
        If Mid(text, count, Len(Key)) = Key Then total = total + 1
    Next
bad:
    counter = total + 1
End Function

Private Function Printable(Varname As String, PrintSpace As Integer) As String
    Dim count As Integer
    If Len(Varname) > PrintSpace Then Varname = Mid(Varname, 1, PrintSpace)
    For count = 1 To PrintSpace - Len(Varname)
        Varname = Varname & " "
    Next
    Printable = Varname
End Function

Private Function RemoveArray(Varname As String) As String
    If InStr(1, Varname, "(") > 0 Then
        Varname = Mid(Varname, 1, InStr(1, Varname, "(") - 1)
    End If
    RemoveArray = Varname
End Function
