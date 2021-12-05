Attribute VB_Name = "screenSamplerPainter"
Option Explicit
Option Base 0

Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
    ByVal hDC As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long) As Long


Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "GDI32" ( _
    ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "GDI32" ( _
    ByVal hDCDest As Long, ByVal XDest As Long, _
    ByVal YDest As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hDCSrc As Long, _
    ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
    As Long

Private Declare Function GetDC Lib "USER32" ( _
    ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "USER32" ( _
    ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Declare Function DeleteDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function CreateThread Lib "kernel32.dll" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private hThreadScreenSampler As Long '//handle to the screen sampler thread

Private hThreadScreenPainter As Long '//handle to the screen painter thread

Private screenWidth, screenHeight As Long

Private painterThreadProcRunning As Boolean '//information regarding the status of the thread, maintained by the thread. True if running

Private exitPainterThread As Boolean '//switch that controls the thread. set to True if we want the thread to stop

Private enableSampling As Boolean '//setting this to False will disable capturing screen

Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Private Type CapturedScreen
            memoryDC As Long '//the memory DC to hold captured screen's bitmap
            memoryBitmap As Long '//the bitmap selected into the memoryDC
End Type

Const CAPTURE_WINDOW As Long = 3000 '//stores this many milliseconds worth of screen capture
Const SAMPLING_INTERVAL As Long = 100 '//sampling interval in milliseconds
Private samplesPerWindow As Integer '//samples per capture window
Private capturedScreenArray() As CapturedScreen '//array for holding the screen captures
Private callRecord As Integer
Private arrayRolling As Boolean '//True if the DCArray has started rolling

Public Sub RunScreenSampler()
    
    samplesPerWindow = CAPTURE_WINDOW \ SAMPLING_INTERVAL
    ReDim capturedScreenArray(samplesPerWindow)
    
    screenWidth = Screen.Width \ Screen.TwipsPerPixelX '//update the global var
    screenHeight = Screen.Height \ Screen.TwipsPerPixelY '//update the global var
    
    enableSampling = True '//enable sampling
    hThreadScreenSampler = CreateThread(0, 0, AddressOf SamplerThreadProc, 0, 0, 0)
End Sub

Private Sub ShiftCapturedArrayLeft()
'//this sub shifts the capturedScreen array over to the left by one to make room for one more recent DC at the last position by ditching the oldest DC at the first position
    
    DeleteDC capturedScreenArray(0).memoryDC '//free up the resource used by the oldest DC
    DeleteObject capturedScreenArray(0).memoryBitmap '//delete the oldest bitmap. ignoring this step causes serious memory leaks as I learnt it first hand(RAM just kept getting stuffed to 100%).
    
    Dim i As Integer
    For i = 0 To UBound(capturedScreenArray) - 1 '//iterate to the penultimate element
        capturedScreenArray(i) = capturedScreenArray(i + 1)
    Next i
End Sub

Private Sub SamplerThreadProc()
    Do
        If enableSampling Then '//only record the screen if toggle enabled
            
            callRecord = callRecord + 1 '//records how many times samples have been taken within the sampling window
            If callRecord > samplesPerWindow Then
                arrayRolling = True
                callRecord = callRecord Mod samplesPerWindow '//roll over the counter. so callRecord can only range from 1 to samplesPerWindow
            End If
            If arrayRolling Then
                Call ShiftCapturedArrayLeft '//slide the elements of the array to the left by one
            End If
            
            Call RecordCurrentScreensBitmapDC(callRecord - 1, arrayRolling)
            
        End If
        Sleep SAMPLING_INTERVAL
    Loop

End Sub


Public Sub RecordCurrentScreensBitmapDC(ByVal useIndex As Integer, ByVal rolledOver As Boolean)
    Dim hDCSrc As Long
    Dim hBmp, hBmpPrev As Long
    Dim r As Boolean
    
    Dim hDCMemory As Long '//this holds the captured screen
    hDCSrc = GetDC(0) 'Get DC for entire screen
    
    hDCMemory = CreateCompatibleDC(hDCSrc) 'Get a memory DC similar to hDCSrc
    
    hBmp = CreateCompatibleBitmap(hDCSrc, screenWidth, screenHeight)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    r = BitBlt(hDCMemory, 0, 0, screenWidth, screenHeight, hDCSrc, 0, 0, vbSrcCopy)
    
    'hBmp = SelectObject(hDCMemory, hBmpPrev)
    ReleaseDC 0, hDCSrc
    
    If Not rolledOver Then '//if the array hasn't started sliding yet
        capturedScreenArray(useIndex).memoryDC = hDCMemory '//add the current captured DC to the array
        capturedScreenArray(useIndex).memoryBitmap = hBmp '//add the current bitmap to the array
    Else '//if the array is sliding
        capturedScreenArray(UBound(capturedScreenArray)).memoryDC = hDCMemory  '//add the current captured DC to the end of the array
        capturedScreenArray(UBound(capturedScreenArray)).memoryBitmap = hBmp '//add the current captured DC's bitmap to the end of the array
    End If
        
End Sub



Public Sub StopScreenPainter()
    exitPainterThread = True
End Sub


Public Sub RunScreenPainter()
    If Not painterThreadProcRunning Then '//if the thread is not running
        exitPainterThread = False '//reset the exit switch
        hThreadScreenPainter = CreateThread(0, 0, AddressOf PainterThreadProc, 0, 0, 0)
    End If
End Sub


Public Sub PainterThreadProc()
    enableSampling = False '//disable capturing screen (we don't want to capture the thing we're painting)
    painterThreadProcRunning = True '//update the global variable, say we're painting
    
    
    Dim hDCSrc As Long
    Dim r As Boolean
    hDCSrc = GetDC(0) '//get DC of the entire screen
    
    Do
        r = BitBlt(hDCSrc, 0, 0, screenWidth, screenHeight, capturedScreenArray(0).memoryDC, 0, 0, vbSrcCopy) '//replace current screen's content with the oldest element of the history array (this will be the screen captured at (t - CAPTURE_WINDOW))
        Sleep 10
    Loop While exitPainterThread = False
    
    ReleaseDC 0, hDCSrc '//release screen's DC
    
    painterThreadProcRunning = False '//update the global variable, say we've stopped painting
    enableSampling = True '//enable capture again
    'MessageBox 0, "Painter thread exited", "", 0
End Sub


