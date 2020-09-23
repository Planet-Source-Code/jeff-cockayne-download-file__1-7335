VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmDownload.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5520
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1485
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   313
      FullHeight      =   41
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   2197
      TabIndex        =   0
      Top             =   3050
      Width           =   1215
   End
   Begin VB.Label RateLabel 
      Caption         =   "RateLabel"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2400
      Width           =   3765
   End
   Begin VB.Label TransferRate 
      Caption         =   "Transfer rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label ToLabel 
      Caption         =   "ToLabel"
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   2100
      Width           =   3765
   End
   Begin VB.Label DownloadTo 
      Caption         =   "Download to:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2100
      Width           =   1575
   End
   Begin VB.Label TimeLabel 
      Caption         =   "TimeLabel"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   3765
   End
   Begin VB.Label SourceLabel 
      Caption         =   "SourceLabel"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   5250
   End
   Begin VB.Label EstimatedTimeLeft 
      Caption         =   "Estimated time left:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label StatusLabel 
      Caption         =   "StatusLabel"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   915
      Width           =   5235
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CancelSearch As Boolean

Public Function DownloadFile(strURL As String, _
                             strDestination As String, _
                             Optional UserName As String = Empty, _
                             Optional Password As String = Empty) _
                             As Boolean

' Funtion DownloadFile: Download a file via HTTP
'
' Author:   Jeff Cockayne
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

Const CHUNK_SIZE As Long = 1024 ' Download chunk size
Const ROLLBACK As Long = 4096   ' Bytes to roll back on resume
                                ' You can be less conservative,
                                ' and roll back less, but I
                                ' don't recommend it.
Dim bData() As Byte             ' Data var
Dim blnResume As Boolean        ' True if resuming download
Dim intFile As Integer          ' FreeFile var
Dim lngBytesReceived As Long    ' Bytes received so far
Dim lngFileLength As Long       ' Total length of file in bytes
Dim lngX                        ' Temp long var
Dim sglLastTime As Single          ' Time last chunk received
Dim sglRate As Single           ' Var to hold transfer rate
Dim sglTime As Single           ' Var to hold time remaining
Dim strFile As String           ' Temp filename var
Dim strHeader As String         ' HTTP header store
Dim strHost As String           ' HTTP Host

On Local Error GoTo InternetErrorHandler

' Start with Cancel flag = False
CancelSearch = False

' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
              
SourceLabel = Empty
TimeLabel = Empty
ToLabel = Empty
RateLabel = Empty

' Pre-open the AVI
With Animation1
    .AutoPlay = True
    .Open App.Path & "\DOWNLD2.AVI"
End With

' Show the download status form
Show
' Move form into view
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

StartDownload:

If blnResume Then
    StatusLabel = "Resuming download..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    StatusLabel = "Getting file information..."
End If
' Give the system time to update the form gracefully
DoEvents

' Download file
With Inet1
    .URL = strURL
    .UserName = UserName
    .Password = Password
    ' GET file, sending the magic resume input header...
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
    ' While initiating connection, yield CPU to Windows
    While .StillExecuting
        DoEvents
        ' If user pressed Cancel button on StatusForm
        ' then fail, cancel, and exit this download
        If CancelSearch Then GoTo ExitDownload
    Wend

    StatusLabel = "Saving:"
    SourceLabel = FitText(SourceLabel, strHost & " from " & .RemoteHost)
    ToLabel = FitText(ToLabel, strDestination)

    ' Get first header ("HTTP/X.X XXX ...")
    strHeader = .GetHeader
End With

' Trap common HTTP response codes
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK
        ' If resuming, however, this is a failure
        If blnResume Then
            ' Delete partially downloaded file
            Kill strDestination
            ' Prompt
            If MsgBox("The server is unable to resume this download." & _
                      vbCr & vbCr & _
                      "Do you want to continue anyway?", _
                      vbExclamation + vbYesNo, _
                      "Unable to Resume Download") = vbYes Then
                    ' Yes - continue anyway:
                    ' Set resume flag to False
                    blnResume = False
                Else
                    ' No - cancel
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
            
    Case "206"  ' 206=Partial Content, which is GREAT when resuming!
    
    Case "204"  ' No content
        MsgBox "Nothing to download!", _
               vbInformation, _
               "No Content"
        CancelSearch = True
        GoTo ExitDownload
        
    Case "401"  ' Not authorized
        MsgBox "Authorization failed!", _
               vbCritical, _
               "Unauthorized"
        CancelSearch = True
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        MsgBox "The file, " & _
               """" & Inet1.URL & """" & _
               " was not found!", _
               vbCritical, _
               "File Not Found"
        CancelSearch = True
        GoTo ExitDownload
        
    Case vbCrLf ' Empty header
        MsgBox "Cannot establish connection." & vbCr & vbCr & _
               "Check your Internet connection and try again.", _
               vbExclamation, _
               "Cannot Establish Connection"
        CancelSearch = True
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MsgBox "The server returned the following response:" & _
               vbCr & vbCr & _
               strHeader, _
               vbCritical, _
               "Error Downloading File"
        CancelSearch = True
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
If blnResume = False Then
    ' Set timer for gauging download speed
    sglLastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If

' Check for available disk space first...
' If on a physical or mapped drive. Can't with a UNC path.
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, _
                          InStr(strDestination, "\"))) < lngFileLength Then
        ' Not enough free space to download file
        MsgBox "There is not enough free space on disk for this file." _
               & vbCr & vbCr & "Please free up some disk space and try again.", _
               vbCritical, _
               "Insufficient Disk Space"
        GoTo ExitDownload
    End If
End If

' Prepare display
'
' Progress Bar
With ProgressBar
    .Value = 0
    .Max = lngFileLength
End With

' Give system a chance to show AVI
DoEvents

' Reset bytes received counter if not resuming
If blnResume = False Then lngBytesReceived = 0


On Local Error GoTo FileErrorHandler

' Create destination directory, if necessary
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If

' If no errors occurred, then spank the file to disk
intFile = FreeFile()        ' Set intFile to an unused file.
' Open a file to write to.
Open strDestination For Binary Access Write As #intFile
' If resuming, then seek byte position in downloaded file
' where we last left off...
If blnResume Then Seek #intFile, lngBytesReceived + 1
Do
    ' Get chunks...
    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    Put #intFile, , bData   ' Put it into our destination file
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    TimeLabel = FormatTime(sglTime) & _
                   " (" & _
                   FormatFileSize(lngBytesReceived) & _
                   " of " & _
                   FormatFileSize(lngFileLength) & _
                   " copied)"
    RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
    ProgressBar.Value = lngBytesReceived
    Me.Caption = Format((lngBytesReceived / lngFileLength), "##0%") & _
                 " of " & strFile & " Completed"
Loop While UBound(bData, 1) > 0       ' Loop while there's still data...
Close #intFile

ExitDownload:
' Success if the # of bytes transferred = content length
If lngBytesReceived = lngFileLength Then
    StatusLabel = "Download completed!"
    DownloadFile = True
Else
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        ' Resume? (If not cancelled)
        If CancelSearch = False Then
            If MsgBox("The connection with the server was reset." & _
                      vbCr & vbCr & _
                      "Click ""Retry"" to resume downloading the file." & _
                      vbCr & "(Approximate time remaining: " & FormatTime(sglTime) & ")" & _
                      vbCr & vbCr & _
                      "Click ""Cancel"" to cancel downloading the file.", _
                      vbExclamation + vbRetryCancel, _
                      "Download Incomplete") = vbRetry Then
                    ' Yes
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If
    ' No or unresumable failure:
    ' Delete partially downloaded file
    If Not Dir(strDestination) = Empty Then Kill strDestination
    DownloadFile = False
End If

CleanUp:
' Close AVI
Animation1.Close

' Make sure that the Internet connection is closed...
Inet1.Cancel
' ...and exit this function
Unload Me

Exit Function

InternetErrorHandler:
    ' Err# 9 occurs when UBound(bData,1) < 0
    If Err.Number = 9 Then Resume Next
    ' Other errors...
    MsgBox "Error: " & Err.Description & " occurred.", _
           vbCritical, _
           "Error Downloading File"
    Err.Clear
    GoTo ExitDownload
    
FileErrorHandler:
    MsgBox "Cannot write file to disk." & _
           vbCr & vbCr & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Error Downloading File"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
    
End Function


Private Sub CancelButton_Click()
    StatusLabel = "Cancelling..."
    Animation1.Close
    CancelSearch = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' For some reason, the Cancel=True and Default=True
' Properties of the "Cancel" button are not working...
' This is the first time that's ever happened to me.
' Any ideas?
If KeyCode = vbKeyEscape Then CancelButton_Click

End Sub

Private Sub Form_Load()
    
' Move form off-screen until it is ready (completely drawn)
Move -Width, -Height

End Sub


Private Sub Form_Unload(Cancel As Integer)

' Move form off-screen so that it disappears "instantly"
Move -Width, -Height
DoEvents

End Sub


