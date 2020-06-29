Attribute VB_Name = "Console"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants, functions, and structures used by this program:
Private Type CONSOLE_CURSOR_INFO
   dwSize As Long
   bVisible As Long
End Type

Private Type COORD
   x As Integer
   y As Integer
End Type

Private Type SMALL_RECT
   Left As Integer
   Top As Integer
   Right As Integer
   Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
   dwSize As COORD
   dwCursorPosition As COORD
   wAttributes As Integer
   srWindow As SMALL_RECT
   dwMaximumWindowSize As COORD
End Type

Private Declare Function AllocConsole Lib "Kernel32.dll" () As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function FreeConsole Lib "Kernel32.dll" () As Long
Private Declare Function GetConsoleCP Lib "Kernel32.dll" () As Long
Private Declare Function GetConsoleCursorInfo Lib "Kernel32.dll" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function GetConsoleMode Lib "Kernel32.dll" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
Private Declare Function GetConsoleOutputCP Lib "Kernel32.dll" () As Long
Private Declare Function GetConsoleProcessList Lib "Kernel32.dll" (lpdwProcessList As Long, ByVal dwProcessCount As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "Kernel32.dll" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function GetConsoleTitleA Lib "Kernel32.dll" (ByVal lpConsoleTitle As String, ByVal nSize As Integer) As Integer
Private Declare Function GetConsoleWindow Lib "Kernel32.dll" () As Long
Private Declare Function GetLargestConsoleWindowSize Lib "Kernel32.dll" (ByVal hConsoleOutput As Long) As COORD
Private Declare Function GetNumberOfConsoleInputEvents Lib "Kernel32.dll" (ByVal hConsoleInput As Long, lpNumberOfEvents As Long) As Long
Private Declare Function GetNumberOfConsoleMouseButtons Lib "Kernel32.dll" (lpNumberOfMouseButtons As Long) As Long
Private Declare Function GetProcessImageFileNameW Lib "Psapi.dll" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetStdHandle Lib "Kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadConsoleA Lib "Kernel32.dll" (ByVal m_hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Long) As Long
Private Declare Function SetConsoleActiveScreenBuffer Lib "Kernel32.dll" (ByVal hConsoleOutput As Long) As Long
Private Declare Function SetConsoleCtrlHandler Lib "Kernel32.dll" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
Private Declare Function SetConsoleMode Lib "Kernel32.dll" (ByVal m_hConsoleOutput As Long, ByVal dwMode As Long) As Long
Private Declare Function SetConsoleTitleA Lib "Kernel32.dll" (ByVal lpConsoleTitle As String) As Long
Private Declare Function WriteConsoleA Lib "Kernel32.dll" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Long) As Long

Private Const ENABLE_ECHO_INPUT As Long = 4
Private Const ENABLE_LINE_INPUT As Long = 2
Private Const ENABLE_MOUSE_INPUT As Long = 16
Private Const ENABLE_PROCESSED_INPUT As Long = 1
Private Const ENABLE_PROCCESED_OUTPUT As Long = 1
Private Const ENABLE_WINDOW_INPUT As Long = 8
Private Const ENABLE_WRAP_AT_EOL_OUTPUT  As Long = 2
Private Const ERROR_SUCCESS As Long = 0
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const STD_ERROR_HANDLE As Long = -12
Private Const STD_INPUT_HANDLE As Long = -10
Private Const STD_OUTPUT_HANDLE As Long = -11

'The constants and variables used by this program:
Private Const MAX_PATH As Long = 260           'Defines the maximum length allowed for a directory/file path.
Private Const MAX_SHORT_STRING As Long = 256   'Defines the maximum length allowed for a short string buffer.
Private Const MAX_STRING As Long = 65535       'Defines the maximum length allowed for a string buffer.

Private ErrorH As Long    'Contains the console's error handle.
Private InputH As Long    'Contains the console's input handle.
Private OutputH As Long   'Contains the console's output handle.

'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Text As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   If Not ErrorCode = ERROR_SUCCESS Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Text = "API error code: " & CStr(ErrorCode) & " - " & Description & vbCrLf
      WriteConsoleA ErrorH, Text, Len(Text), CLng(0), CLng(0)
      Text = "Return value: " & CStr(ReturnValue) & vbCrLf & vbCrLf
      WriteConsoleA ErrorH, Text, Len(Text), CLng(0), CLng(0)
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles the control signals from the console.
Private Function ControlHandler(CtrlType As Long) As Long
On Error GoTo ErrorTrap
EndRoutine:
   ControlHandler = CLng(False)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure displays console related information.
Private Sub DisplayConsoleInformation()
On Error GoTo ErrorTrap
Dim BufferInformation As CONSOLE_SCREEN_BUFFER_INFO
Dim Cursor As CONSOLE_CURSOR_INFO
Dim InputEventCount As Long
Dim InputMode As Long
Dim LargestSize As COORD
Dim MouseButtonCount As Long
Dim OutputMode As Long

   LargestSize = GetLargestConsoleWindowSize(OutputH)
   
   CheckForError GetConsoleCursorInfo(OutputH, Cursor)
   CheckForError GetConsoleMode(InputH, InputMode)
   CheckForError GetConsoleMode(OutputH, OutputMode)
   CheckForError GetConsoleScreenBufferInfo(OutputH, BufferInformation)
   CheckForError GetNumberOfConsoleInputEvents(OutputH, InputEventCount)
   CheckForError GetNumberOfConsoleMouseButtons(MouseButtonCount)
   
   WriteToConsole "[Console Information]" & vbCrLf
   WriteToConsole "Code page: " & CStr(CheckForError(GetConsoleCP())) & vbCrLf
   WriteToConsole "Cursor size: " & CStr(Cursor.dwSize) & vbCrLf
   WriteToConsole "Cursor visible: " & CStr(Cursor.bVisible) & vbCrLf
   WriteToConsole "Error handle: " & CStr(ErrorH) & vbCrLf
   WriteToConsole "Input handle: " & CStr(InputH) & vbCrLf
   WriteToConsole "Input mode: 0x" & Hex$(InputMode) & vbCrLf
   WriteToConsole "Largest size: " & CStr(LargestSize.x) & ", " & CStr(LargestSize.y) & vbCrLf
   WriteToConsole "Number of mouse buttons used: " & CStr(MouseButtonCount) & vbCrLf
   WriteToConsole "Number of unhandled input events: " & CStr(InputEventCount) & vbCrLf
   WriteToConsole "Output code page: " & CStr(CheckForError(GetConsoleOutputCP())) & vbCrLf
   WriteToConsole "Output handle: " & CStr(OutputH) & vbCrLf
   WriteToConsole "Output mode: 0x" & Hex$(OutputMode) & vbCrLf
   WriteToConsole "Title: " & GetConsoleTitle() & vbCrLf
   WriteToConsole "Window handle: " & CStr(CheckForError(GetConsoleWindow())) & vbCrLf
   
   WriteToConsole vbCrLf
   WriteToConsole "[Screen buffer information]" & vbCrLf
   With BufferInformation
      WriteToConsole "Character attributes: " & CStr(.wAttributes) & vbCrLf
      WriteToConsole "Cursor position: " & CStr(.dwCursorPosition.x) & ", " & CStr(.dwCursorPosition.y) & vbCrLf
      WriteToConsole "Lower right corner: " & CStr(.srWindow.Right) & ", " & CStr(.srWindow.Bottom) & vbCrLf
      WriteToConsole "Maximum window size: " & CStr(.dwMaximumWindowSize.x) & ", " & CStr(.dwMaximumWindowSize.y) & vbCrLf
      WriteToConsole "Size: " & CStr(.dwSize.x) & ", " & CStr(.dwSize.y) & vbCrLf
      WriteToConsole "Upper left corner: " & CStr(.srWindow.Left) & ", " & CStr(.srWindow.Top) & vbCrLf
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns a list of processes attached to the console.
Public Sub GetConsoleProcesses(ProcessList() As Long)
On Error GoTo ErrorTrap
Dim ProcessCount As Long

   ReDim ProcessList(1 To 1) As Long
   ProcessCount = CheckForError(GetConsoleProcessList(ProcessList(1), UBound(ProcessList())))
   
   ReDim ProcessList(LBound(ProcessList()) To ProcessCount) As Long
   CheckForError GetConsoleProcessList(ProcessList(1), UBound(ProcessList()))
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the console's title.
Private Function GetConsoleTitle() As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim Title As String

   Title = String$(MAX_SHORT_STRING, vbNullChar)
   Length = CheckForError(GetConsoleTitleA(Title, Len(Title)))
   If Length > 0 Then Title = Left$(Title, Length)
   
EndRoutine:
   GetConsoleTitle = Title
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure handles any errors that occur.
Private Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   
   On Error Resume Next
   WriteToConsole "Error: " & CStr(ErrorCode) & " - " & Description & vbCrLf, IsError:=True
End Sub

'This procedure initializes this program.
Private Sub Initialize()
On Error GoTo ErrorTrap

   CheckForError AllocConsole()
   
   With App
      CheckForError SetConsoleTitleA("Console v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName)
   End With
   
   CheckForError SetConsoleCtrlHandler(AddressOf ControlHandler, CLng(True))
   
   ErrorH = CheckForError(GetStdHandle(STD_ERROR_HANDLE))
   InputH = CheckForError(GetStdHandle(STD_INPUT_HANDLE))
   OutputH = CheckForError(GetStdHandle(STD_OUTPUT_HANDLE))
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays a list of the processes that are using the console.
Private Sub ListConsoleProcesses()
On Error GoTo ErrorTrap
Dim Index As Long
Dim ProcessList() As Long

   GetConsoleProcesses ProcessList()
   For Index = LBound(ProcessList()) To UBound(ProcessList())
      WriteToConsole CStr(ProcessList(Index)) & vbTab & GetProcessImageName(ProcessList(Index)) & vbCrLf
   Next Index
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the process image name for the specified process id.
Private Function GetProcessImageName(ProcessId As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim ProcessHandle As Long
Dim ProcessImageName As String
 
   ProcessHandle = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(False), ProcessId))
   ProcessImageName = String$(MAX_PATH, vbNullChar)
   Length = CheckForError(GetProcessImageFileNameW(ProcessHandle, StrPtr(ProcessImageName), Len(ProcessImageName)))
   CheckForError CloseHandle(ProcessHandle)
    
EndRoutine:
   GetProcessImageName = Left$(ProcessImageName, Length)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure lists the current directory's contents.
Private Sub ListDirectory()
On Error GoTo ErrorTrap
Dim FileName As String

   FileName = Dir$("*.*", vbDirectory Or vbHidden Or vbNormal Or vbSystem Or vbVolume)
   Do Until FileName = Empty
      WriteToConsole FileName & vbCrLf
      FileName = Dir$()
   Loop
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure manages the console.
Private Sub Main()
On Error GoTo ErrorTrap
Dim CommandV As String
Dim Parameters As String
Dim Position As Long

   Initialize
   
   WriteToConsole "Type ""help"" for a list of commands." & vbCrLf
   Do
      DoEvents
      WriteToConsole ">"
      CommandV = LCase$(Trim$(ReadFromConsole()))
      
      If Not CommandV = Empty Then
         Position = InStr(CommandV, " ")
         If Position > 0 Then
            Parameters = Mid$(CommandV, Position + 1)
            CommandV = Left$(CommandV, Position - 1)
         End If
         
         Select Case CommandV
            Case "chdir"
               ChDir Parameters
            Case "chdrive"
               ChDrive Parameters
            Case "consoleinfo"
               DisplayConsoleInformation
            Case "curdir"
               WriteToConsole CurDir$ & vbCrLf
            Case "exit", "quit"
               Quit
            Case "help", "?"
               WriteToConsole "chdrive" & vbCrLf
               WriteToConsole "chdir" & vbCrLf
               WriteToConsole "consoleinfo" & vbCrLf
               WriteToConsole "curdir" & vbCrLf
               WriteToConsole "exit" & vbCrLf
               WriteToConsole "help" & vbCrLf
               WriteToConsole "listdir" & vbCrLf
               WriteToConsole "listprocesses" & vbCrLf
               WriteToConsole "quit" & vbCrLf
               WriteToConsole "runprg" & vbCrLf
            Case "listdir"
               ListDirectory
            Case "listprocesses"
               ListConsoleProcesses
            Case "runprg"
               Shell Parameters
            Case Else
               WriteToConsole "<ERROR>" & vbCrLf
         End Select
      End If
   Loop
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume Next
End Sub

'This procedure closes the console and ends the program.
Private Sub Quit()
On Error GoTo ErrorTrap
   CheckForError FreeConsole()
   End
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure reads from the console.
Private Function ReadFromConsole() As String
On Error GoTo ErrorTrap
Dim CharactersRead As Long
Dim Text As String

   Text = String$(MAX_SHORT_STRING, vbNullChar)
   CheckForError ReadConsoleA(InputH, Text, Len(Text), CharactersRead, CLng(0))
   If CharactersRead < Len(vbCrLf) Then CharactersRead = Len(vbCrLf)
   
EndRoutine:
   ReadFromConsole = Left$(Text, CharactersRead - 2)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure writes the specified text to the console.
Private Sub WriteToConsole(Text As String, Optional IsError As Boolean = False)
On Error GoTo ErrorTrap
   If IsError Then
      CheckForError WriteConsoleA(ErrorH, Text, Len(Text), CLng(0), CLng(0))
   Else
      CheckForError WriteConsoleA(OutputH, Text, Len(Text), CLng(0), CLng(0))
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


