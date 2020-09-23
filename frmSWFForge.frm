VERSION 5.00
Begin VB.Form frmSWFForge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SWF Forge"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5910
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdSelectPlayerFile 
      Caption         =   "..."
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtPlayerPath 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   1440
      Width           =   4260
   End
   Begin VB.OptionButton optForge 
      Appearance      =   0  'Flat
      Caption         =   "&Compile"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optForge 
      Appearance      =   0  'Flat
      Caption         =   "&Parse"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdForge 
      Caption         =   "&Forge!"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   3720
      Width           =   1035
   End
   Begin VB.TextBox txtSWFPath 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   4260
   End
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1695
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   1920
      Width           =   4775
   End
   Begin VB.TextBox txtEXEPath 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   4260
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   3720
      Width           =   1035
   End
   Begin VB.CommandButton cmdSelectSWFFile 
      Caption         =   "..."
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdSelectEXEFile 
      Caption         =   "..."
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Mode:"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Player:"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "EXE File:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   810
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "SWF File:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   810
   End
End
Attribute VB_Name = "frmSWFForge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdExit_Click()
        Unload Me
End Sub


Private Sub cmdForge_Click()
    On Error GoTo HandleError
        
        If Trim(txtEXEPath.Text) <> "" And Trim(txtSWFPath.Text) <> "" Then
            Dim o_ucmStorage As New cStorage
            
            With o_ucmStorage
                Dim o_lngTotal As Long
                Dim o_lngData As Long
                Dim o_strTime As String
                Dim o_lngPlayer As Long
                Dim o_strRet As String
                
                If optForge(0).Value Then
                    o_strRet = "Parsing: " & IIf(.Parse(txtEXEPath.Text, txtSWFPath.Text, o_lngTotal, o_lngData, o_strTime, txtPlayerPath.Text), "ok", "failed") & vbCrLf
                    o_strRet = o_strRet & "Original size: " & o_lngTotal & " bytes," & vbCrLf & "extracted size: " & o_lngData & " bytes"
                Else
                    o_strRet = "Compiling: " & IIf(.Compile(txtPlayerPath.Text, txtSWFPath.Text, txtEXEPath.Text, o_lngPlayer, o_lngData, o_lngTotal, o_strTime), "ok", "failed") & vbCrLf
                    o_strRet = o_strRet & "Player size: " & o_lngPlayer & " bytes," & vbCrLf & "swf size: " & o_lngData & " bytes," & vbCrLf & "output size: " & o_lngTotal & " bytes"
                End If
            End With
            
            Set o_ucmStorage = Nothing
            
            txtContent.Text = o_strRet
        Else
            MsgBox "Please select exe file and swf file!"
        End If
        
    Exit Sub
    
HandleError:
    MsgBox Err.Description
End Sub


Private Sub cmdSelectEXEFile_Click()
        Dim o_blnRet As Boolean
        
        With New cFileDlg
            With .Filters
                .Clear
                .AddFilter "Compiled SWF EXE file", "*.exe"
            End With
            .Flags = ofnPathMustExist Or ofnFileMustExist
            .InitDir = GetFilePath(txtEXEPath.Text)
            If optForge(0).Value Then
                o_blnRet = .ShowOpen(Me.hwnd)
            Else
                o_blnRet = .ShowSave(Me.hwnd)
            End If
            If o_blnRet Then
                txtEXEPath.Text = .FileName
            End If
        End With
End Sub


Private Sub cmdSelectPlayerFile_Click()
        Dim o_blnRet As Boolean
        
        With New cFileDlg
            With .Filters
                .Clear
                .AddFilter "Flash Player", "*.exe"
            End With
            .Flags = ofnPathMustExist
            If optForge(1).Value Then
                .Flags = .Flags Or ofnFileMustExist
            End If
            .InitDir = GetFilePath(txtPlayerPath.Text)
            If optForge(0).Value Then
                o_blnRet = .ShowSave(Me.hwnd)
            Else
                o_blnRet = .ShowOpen(Me.hwnd)
            End If
            If o_blnRet Then
                txtPlayerPath.Text = .FileName
            End If
        End With
End Sub


Private Sub cmdSelectSWFFile_Click()
        Dim o_blnRet As Boolean
        
        With New cFileDlg
            .Filters.AddFilter "SWF file", "*.swf"
            .Flags = ofnPathMustExist Or ofnCreatePrompt Or ofnOverwritePrompt
            .InitDir = GetFilePath(txtSWFPath.Text)
            If optForge(1).Value Then
                o_blnRet = .ShowOpen(Me.hwnd)
            Else
                o_blnRet = .ShowSave(Me.hwnd)
                If .FileName <> "" Then
                    If InStr(1, .FileName, ".swf") = 0 Then
                        .FileName = .FileName & ".swf"
                    End If
                End If
            End If
            If o_blnRet Then
                txtSWFPath.Text = .FileName
            End If
        End With
End Sub


Public Function GetFilePath(ByVal strFile As String) As String
        If strFile <> "" Then
            If InStr(1, strFile, "\") <> 0 Then
                GetFilePath = Left$(strFile, InStrRev(strFile, "\") - 1)
            End If
        End If
End Function


Private Sub Form_Load()
        Dim o_intFileNum  As Integer
        Dim o_bytData() As Byte
        
        o_intFileNum = FreeFile
        
        'extract a compiled swf so that you guys will not find it too hard
        'since the player could be extracted from it:)
        Open App.Path & "\Marshi Maro.exe" For Binary As #o_intFileNum
        
        o_bytData = LoadResData("Data", "CUSTOM")
        
        Put #o_intFileNum, , o_bytData
        
        Close #o_intFileNum

End Sub


Private Sub Form_Unload(Cancel As Integer)
        Set frmSWFForge = Nothing
        
        End
End Sub
