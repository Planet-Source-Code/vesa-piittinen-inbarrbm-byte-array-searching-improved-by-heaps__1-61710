VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InStr vs. InBArrBM (by Vesa Piittinen)"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBenchmark 
      Caption         =   "Benchmark results:"
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   9015
      Begin VB.ListBox lstBenchmark 
         Height          =   1455
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.CommandButton cmdCodeCompare 
      Caption         =   "B&enchmark InStr"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdCodeBenchmark 
      Caption         =   "&Benchmark InBArrBM"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Frame fraCode 
      Caption         =   "Function call details:"
      Height          =   855
      Left            =   2760
      TabIndex        =   2
      Top             =   2640
      Width           =   6375
      Begin VB.TextBox txtCodeIterations 
         Height          =   285
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   10
         Text            =   "1000"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkUnicode 
         Caption         =   "Unicode (InBArrBM)"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkTextCompare 
         Caption         =   "TextCompare"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblCode 
         Caption         =   "Iterations:"
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.ListBox lstCodes 
      Height          =   3375
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type CODEDATA
    Name As String
    Text As String
End Type

Dim CodeList() As CODEDATA
Private Function CompareTest(ByVal SearchString As String, ByVal SearchKeyword As String, ByVal Iterations As Long) As Long
    Dim A As Long, CompareMode As Long
    Dim tmpArr() As Byte, tmpResult As Long
    ' error checking
    If LenB(SearchString) = 0 Then Exit Function
    If Iterations < 1 Then Exit Function
    ' get settings
    CompareMode = CLng(chkTextCompare.Value)
    ' the real benchmark
    For A = 1 To Iterations
        tmpResult = InStr(1, SearchString, SearchKeyword, CompareMode)
    Next A
    ' output the result (convert to comparable character position)
    CompareTest = (tmpResult - 1)
End Function
Private Sub GetCode(ByVal TestName As String, ByRef SearchText As String, ByRef SearchKey As String)
    Select Case TestName
        Case "Long keyword"
            SearchText = Space$(999) & String$(99, "x") & Space$(999)
            SearchKey = String$(99, "x")
        Case "Longer keyword"
            SearchText = Space$(99999) & String$(9999, "x") & Space$(99999)
            SearchKey = String$(9999, "x")
        Case "TextCompare"
            SearchText = Space$(99999) & String$(500, "Xx") & Space$(99999)
            SearchKey = String$(500, "xX")
        Case "VBspeed: InStr Call 1"
            SearchText = "http://www.xbeat.net/vbspeed/index.htm"
            SearchKey = ":"
        Case "VBspeed: InStr Call 2"
            SearchText = "http://www.xbeat.net/vbspeed/index.htm"
            SearchKey = "."
        Case "VBspeed: InStr Call 3"
            SearchText = "http://www.xbeat.net/vbspeed/index.htm"
            SearchKey = "m"
        Case "VBspeed: InStr Call 4"
            SearchText = "http://www.xbeat.net/vbspeed/index.htm"
            SearchKey = "M"
        Case "VBspeed: InStr Call 5"
            SearchText = "http://www.xbeat.net/vbspeed/index.htm"
            SearchKey = "www"
        Case "VBspeed: InStr Call 6"
            SearchText = "http://www.xbeat.net/vbspeed/index.htm"
            SearchKey = "WWW"
        Case "VBspeed: InStr Call 7"
            SearchText = Space$(999) & String$(99, "x") & Space$(99)
            SearchKey = "x"
        Case "VBspeed: InStr Call 8"
            SearchText = Space$(99)
            SearchKey = "x"
        Case "VBspeed: InStr Call 8 (mod)"
            SearchText = Space$(99)
            SearchKey = "xxxxx"
    End Select
End Sub
Private Sub InitCodes()
    Dim Codes() As String, Buffer As String
    Dim A As Long
    lstCodes.Clear
    Open App.Path & "\benchmark.txt" For Input As #1
        Buffer = Input(LOF(1), #1)
    Close #1
    Codes = Split(Buffer, "~" & vbCrLf)
    ReDim CodeList(UBound(Codes) - 1)
    For A = 0 To UBound(CodeList)
        With CodeList(A)
            .Name = Left$(Codes(A), InStr(Codes(A), vbCrLf) - 1)
            .Text = Mid$(Codes(A), InStr(Codes(A), vbCrLf) + 2)
            lstCodes.AddItem .Name
            lstCodes.ItemData(lstCodes.NewIndex) = A
        End With
    Next A
    If lstCodes.ListCount Then lstCodes.ListIndex = 0
End Sub
Public Function IsGoodInBArrBM(Optional fLigaturesToo As Boolean) As Boolean
    ' verify correct InBArrBM returns, 2005-07-08
    ' based on InStr test code available at VBspeed
    ' returns True if all tests are passed
    Dim fFailed As Boolean
    Dim Temp() As Byte
    Dim Test As Long
    
    ' replace "InBArrBM" with the name of your function
    Temp = "abc"
    If InBArrBM(Temp, "b") <> 2 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "ab") <> 0 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "aB") <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "aB", , vbTextCompare) <> 0 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "ab", 2) <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "ab", 4) <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "ab", 6) <> -1 Then Stop: fFailed = True
    Temp = "aaabcab"
    If InBArrBM(Temp, "abc", 6) <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "", 6) <> 6 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "", 8) <> 8 Then Stop: fFailed = True
    Erase Temp
    If InBArrBM(Temp, "", 6) <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "c") <> -1 Then Stop: fFailed = True
    
    Temp = "abcdabcd"
    If InBArrBM(Temp, "abcd") <> 0 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "Ab") <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrBM(Temp, "Ab", , vbTextCompare) <> 0 Then Stop: fFailed = True
    
    Temp = "a" & String$(50000, "b")
    If InBArrBM(Temp, "a") <> 0 Then Stop: fFailed = True
    
    ' unicode
    Temp = "a€€c"
    If InBArrBM(Temp, "€") <> 2 Then Stop: fFailed = True
    
    ' the 4 stooges: š/Š, œ/Œ, ž/Ž, ÿ/Ÿ (154/138, 156/140, 158/142, 255/159)
    Temp = "Hašiš"
    If InBArrBM(Temp, "Š", , vbTextCompare) <> 4 Then Stop: fFailed = True
    ' ligatures  textcompare (VBspeed entries do NOT have to pass this test)
    If fLigaturesToo Then
        ' ligatures, a digraphemic fun house: ss/ß, ae/æ, oe/œ, th/þ
        Temp = "Straße"
        If InBArrBM(Temp, "ss", , vbTextCompare) <> 8 Then Stop: fFailed = True
    End If
    
    ' well done
    IsGoodInBArrBM = Not fFailed
End Function
Public Function IsGoodInBArrBMANSI() As Boolean
    ' verify correct InBArrBM returns (ANSI mode), 2005-07-08
    ' based on InStr test code available at VBspeed
    ' returns True if all tests are passed
    Dim fFailed As Boolean
    Dim Temp() As Byte
    Dim Test As Long
    
    ' replace ".InStr01" with the name of your function
    Temp = StrConv("abc", vbFromUnicode)
    If InBArrBM(Temp, "b", , , False) <> 1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "ab", , , False) <> 0 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "aB", , , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "aB", , vbTextCompare, False) <> 0 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "ab", 1, , False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "ab", 2, , False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "ab", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("aaabcab", vbFromUnicode)
    If InBArrBM(Temp, "abc", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "", 3, , False) <> 3 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "", 4, , False) <> 4 Then Stop: fFailed = True
    Erase Temp
    If InBArrBM(Temp, "", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "c", , , False) <> -1 Then Stop: fFailed = True
    
    Temp = StrConv("abcdabcd", vbFromUnicode)
    If InBArrBM(Temp, "abcd", , , False) <> 0 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "Ab", , , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrBM(Temp, "Ab", , vbTextCompare, False) <> 0 Then Stop: fFailed = True
    
    Temp = StrConv("a" & String$(50000, "b"), vbFromUnicode)
    If InBArrBM(Temp, "a", , , False) <> 0 Then Stop: fFailed = True
    
    ' well done
    IsGoodInBArrBMANSI = Not fFailed
End Function
Private Function PerformTest(ByVal SearchString As String, ByVal SearchKeyword As String, ByVal Iterations As Long) As Long
    Dim A As Long, CompareMode As Long, UnicodeMode As Boolean
    Dim tmpArr() As Byte, tmpResult As Long
    ' error checking
    If LenB(SearchString) = 0 Then Exit Function
    If Iterations < 1 Then Exit Function
    ' get settings
    CompareMode = CLng(chkTextCompare.Value)
    UnicodeMode = (chkUnicode.Value = vbChecked)
    ' convert string to byte array
    If UnicodeMode Then
        tmpArr = SearchString
    Else
        tmpArr = StrConv(SearchString, vbFromUnicode)
    End If
    ' the real benchmark
    For A = 1 To Iterations
        tmpResult = InBArrBM(tmpArr, SearchKeyword, 0, CompareMode, UnicodeMode)
    Next A
    ' output the result (and convert to comparable character position)
    If UnicodeMode Then
        If tmpResult >= 0 Then
            PerformTest = tmpResult \ 2
        Else
            PerformTest = -1
        End If
    Else
        PerformTest = tmpResult
    End If
End Function
Private Sub cmdCodeBenchmark_Click()
    Dim tmpResult As Long, TestName As String, Iterations As Long
    Dim tmpSearch As String, tmpKey As String
    If lstCodes.ListIndex < 0 Then Exit Sub
    Me.MousePointer = vbHourglass
    TestName = CodeList(lstCodes.ItemData(lstCodes.ListIndex)).Name
    Iterations = CLng(Val(txtCodeIterations.Text))
    GetCode TestName, tmpSearch, tmpKey
    hpc_Start
    tmpResult = PerformTest(tmpSearch, tmpKey, Iterations)
    lstBenchmark.AddItem "[BARR] " & TestName & " results: value is " & tmpResult & ", time " & Format$(hpc_Finish / 1000, "0.000") & " s [" & Iterations & " iterations]", 0
    lstBenchmark.ListIndex = 0
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdCodeCompare_Click()
    Dim tmpResult As Long, TestName As String, Iterations As Long
    Dim tmpSearch As String, tmpKey As String
    If lstCodes.ListIndex < 0 Then Exit Sub
    Me.MousePointer = vbHourglass
    TestName = CodeList(lstCodes.ItemData(lstCodes.ListIndex)).Name
    Iterations = CLng(Val(txtCodeIterations.Text))
    GetCode TestName, tmpSearch, tmpKey
    hpc_Start
    tmpResult = CompareTest(tmpSearch, tmpKey, Iterations)
    lstBenchmark.AddItem "[INSTR] " & TestName & " results: value is " & tmpResult & ", time " & Format$(hpc_Finish / 1000, "0.000") & " s [" & Iterations & " iterations]", 0
    lstBenchmark.ListIndex = 0
    Me.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    InitCodes
    On Error Resume Next
    Debug.Print 1 \ 0
    If Err.Number Then
        Err.Clear
        MsgBox "Compile please!" & vbCrLf & vbCrLf & _
            "File > Make StrVsBArr.exe > Options > Compile > make sure Advanced Optimizations are all ticked." & vbCrLf & _
            "This makes sure you get the most speed possible with native VB6 code.", _
            vbInformation, "Running benchmark under the IDE is a CRIME"
    End If
    On Error GoTo 0
End Sub
Private Sub lstCodes_Click()
    If lstCodes.ListIndex < 0 Then Exit Sub
    txtCode.Text = CodeList(lstCodes.ItemData(lstCodes.ListIndex)).Text
    chkTextCompare.Value = Abs(InStr(txtCode.Text, "Compare = vbTextCompare") > 0)
End Sub
