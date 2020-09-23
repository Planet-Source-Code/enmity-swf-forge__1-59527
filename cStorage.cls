VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************************************************
'* Free to use, without any warranty, Autor: Unruledboy@hotmail.com
'************************************************************************
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'default chunk size, bigger chunk for better performance but more memory usage
Private Const mc_lngChunkSize As Long = 8192
'since the oldest player we can find is at least bigger than 300,000 bytes(300k)
Private Const mc_lngStartPosition As Long = 300000

'ok, it is up to you to define the chunk size:)
Private m_lngChunckSize As Long
'this IS the swf signature
Private m_bytDataSectionHeader(0 To 2) As Byte



Public Property Get ChunkSize() As Long
    ChunkSize = m_lngChunckSize
End Property

Public Property Let ChunkSize(ByVal lngChunkSize As Long)
    m_lngChunckSize = lngChunkSize
End Property


'quick & dirty way to find a byte array in another byte array, hey, it just WORKS!
Private Function FindSection(ByRef bytChunk() As Byte, ByRef bytHeader() As Byte, ByVal lngPos As Long) As Long
    FindSection = InStrB(lngPos, bytChunk, bytHeader, vbBinaryCompare)
End Function


'ok, this is what we need, the parser :)
Public Function Parse(ByVal strExeFilePath As String, _
                      ByVal strSWFFilePath As String, _
                      ByRef lngTotalLength As Long, _
                      ByRef lngDataLength As Long, _
                      ByRef strTime As String, _
                      Optional ByVal strPlayer As String = "") As Boolean
    Dim o_ucmTimer As New cPrecTimer
    
    o_ucmTimer.ResetTimer
    
    Dim o_intFileNum As Integer
    Dim o_bytChunk() As Byte
    Dim o_bytPending() As Byte
    Dim o_blnRet As Boolean
    Dim o_lngDataPosition As Long
    Dim o_lngFoundPosition As Long
    
    o_intFileNum = FreeFile
    
    Open strExeFilePath For Binary As #o_intFileNum
    
    lngTotalLength = LOF(o_intFileNum)
    
    If lngTotalLength > mc_lngStartPosition Then
        o_lngDataPosition = mc_lngStartPosition
        
        Seek #o_intFileNum, o_lngDataPosition
        
        'til we found the signature
        Do Until o_lngFoundPosition > 0 Or o_blnRet
            If o_lngDataPosition + m_lngChunckSize > lngTotalLength Then
                m_lngChunckSize = lngTotalLength - o_lngDataPosition
                o_blnRet = True
            End If
            
            ReDim o_bytChunk(0 To m_lngChunckSize - 1) As Byte
            Get #o_intFileNum, , o_bytChunk
            
            If Not IsObjectArrayInit(o_bytPending) Then
                ReDim o_bytPending(m_lngChunckSize - 1) As Byte
            Else
                ReDim Preserve o_bytPending(0 To UBound(o_bytPending) + m_lngChunckSize) As Byte
            End If
            
            CopyMemory o_bytPending(UBound(o_bytPending) - m_lngChunckSize + 1), o_bytChunk(0), UBound(o_bytChunk) + 1
            
            o_lngFoundPosition = FindSection(o_bytPending, m_bytDataSectionHeader, 1)
            
            o_lngDataPosition = o_lngDataPosition + m_lngChunckSize
            
            If o_lngDataPosition > lngTotalLength Then Exit Do
        Loop
        
        'if we found it
        If o_lngFoundPosition > 0 Then
            
            Dim o_intFileNum2 As Integer
            
            lngDataLength = lngTotalLength - mc_lngStartPosition - 1
            
            o_lngFoundPosition = mc_lngStartPosition + o_lngFoundPosition - 1
                      
            'extract the player as well
            If strPlayer <> "" Then
                o_intFileNum2 = FreeFile
                
                Open strPlayer For Binary As #o_intFileNum2
                
                Seek #o_intFileNum, 1
                
                ReDim o_bytChunk(0 To o_lngFoundPosition - 1) As Byte
                Get #o_intFileNum, , o_bytChunk
                Put #o_intFileNum2, , o_bytChunk

                Close #o_intFileNum2
            End If
            
            Seek #o_intFileNum, o_lngFoundPosition
            
            o_intFileNum2 = FreeFile
            
            Open strSWFFilePath For Binary As #o_intFileNum2
            
            o_lngDataPosition = o_lngFoundPosition
            
            'copy chunk by chunk to destination swf file
            Do Until EOF(o_intFileNum) Or o_blnRet
                If o_lngDataPosition + m_lngChunckSize > lngTotalLength Then
                    m_lngChunckSize = lngTotalLength - o_lngDataPosition
                    o_blnRet = True
                End If

                ReDim o_bytChunk(0 To m_lngChunckSize - 1) As Byte
                Get #o_intFileNum, , o_bytChunk
                Put #o_intFileNum2, , o_bytChunk

                o_lngDataPosition = o_lngDataPosition + m_lngChunckSize
            Loop
            
            Close #o_intFileNum2
        
            o_blnRet = True
        Else
            o_blnRet = False
        End If
    Else
        o_blnRet = False
    End If
    
    Close #o_intFileNum
    
    'hey, remember to release memory!
    Erase o_bytChunk
    Erase o_bytPending
    
    o_ucmTimer.StopTimer
    strTime = o_ucmTimer.ToString
    
    Set o_ucmTimer = Nothing
    
    Parse = o_blnRet
End Function


'ok, this is what we need, the compiler :)
Public Function Compile(ByVal strPlayerFilePath As String, _
                        ByVal strSWFFilePath As String, _
                        ByVal strExeFilePath As String, _
                        ByRef lngPlayerLength As Long, _
                        ByRef lngDataLength As Long, _
                        ByRef lngTotalLength As Long, _
                        ByRef strTime As String) As Boolean
    Dim o_ucmTimer As New cPrecTimer
    
    o_ucmTimer.ResetTimer
    
    Dim o_intFileNum As Integer
    Dim o_bytChunk() As Byte
    Dim o_blnRet As Boolean
    Dim o_lngDataPosition As Long
    Dim o_lngFoundPosition As Long
    
    o_intFileNum = FreeFile
    Open strSWFFilePath For Binary As #o_intFileNum
    
    lngDataLength = LOF(o_intFileNum)
    
    'first let verify whether it is a swf file or not
    If lngDataLength > UBound(m_bytDataSectionHeader) + 1 Then
        ReDim o_bytChunk(0 To UBound(m_bytDataSectionHeader)) As Byte
        Get #o_intFileNum, , o_bytChunk
    
        o_blnRet = FindSection(o_bytChunk, m_bytDataSectionHeader, 1) > 0
    Else
        o_blnRet = False
    End If
    
    Close #o_intFileNum
    
    'it IS a SWF file
    If o_blnRet Then
        Dim o_intFileNum2 As Integer
        
        o_intFileNum2 = FreeFile
        Open strExeFilePath For Binary As #o_intFileNum2
        
        o_intFileNum = FreeFile
        Open strPlayerFilePath For Binary As #o_intFileNum
        
        lngPlayerLength = LOF(o_intFileNum)
        
        o_lngDataPosition = 0
        
        o_blnRet = False

        'copy the player file chunk by chunk to destination exe file
        Do Until EOF(o_intFileNum) Or o_blnRet
            If o_lngDataPosition + m_lngChunckSize > lngTotalLength Then
                m_lngChunckSize = lngPlayerLength - o_lngDataPosition
                o_blnRet = True
            End If

            ReDim o_bytChunk(0 To m_lngChunckSize - 1) As Byte
            Get #o_intFileNum, , o_bytChunk
            Put #o_intFileNum2, , o_bytChunk

            o_lngDataPosition = o_lngDataPosition + m_lngChunckSize
        Loop
        
        Close #o_intFileNum

        
        o_intFileNum = FreeFile
        Open strSWFFilePath For Binary As #o_intFileNum
        
        lngDataLength = LOF(o_intFileNum)
        
        o_lngDataPosition = 0
        
        o_blnRet = False

        'copy the swf file chunk by chunk to destination exe file
        Do Until EOF(o_intFileNum) Or o_blnRet
            If o_lngDataPosition + m_lngChunckSize > lngTotalLength Then
                m_lngChunckSize = lngDataLength - o_lngDataPosition
                o_blnRet = True
            End If

            ReDim o_bytChunk(0 To m_lngChunckSize - 1) As Byte
            Get #o_intFileNum, , o_bytChunk
            Put #o_intFileNum2, , o_bytChunk

            o_lngDataPosition = o_lngDataPosition + m_lngChunckSize
        Loop
        
        lngTotalLength = LOF(o_intFileNum2)
        
        Close #o_intFileNum

        Close #o_intFileNum2
        
        o_blnRet = True
    Else
        o_blnRet = False
    End If
        
    'hey, remember to release memory!
    Erase o_bytChunk
    
    o_ucmTimer.StopTimer
    strTime = o_ucmTimer.ToString
    
    Set o_ucmTimer = Nothing
    
    Compile = o_blnRet
End Function


Private Function IsObjectArrayInit(ByRef vntObject As Variant) As Boolean

    On Error GoTo HandleError
        
        IsObjectArrayInit = CBool(UBound(vntObject) + 1)

    Exit Function
        
HandleError:

    IsObjectArrayInit = False

End Function


Private Sub Class_Initialize()
    m_lngChunckSize = mc_lngChunkSize
    
    m_bytDataSectionHeader(0) = 70
    m_bytDataSectionHeader(1) = 87
    m_bytDataSectionHeader(2) = 83
End Sub
