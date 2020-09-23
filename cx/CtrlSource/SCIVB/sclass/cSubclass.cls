VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* cSubclass - Subclassing class. Use either the implemented interface class, iSubclass, or the
'*             supplied type library.
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'*************************************************************************************************

Option Explicit

'-cSublass declarations---------------------------------------------------------------------------
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const CODE_LEN      As Long = 276                                   'Subclass thunk machine-code length in bytes
Private Const WNDPROC_OFF   As Long = &H34                                  'WndProc thunk offset into the allocated memory
Private Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))    'Number of bytes to allocate per hWnd thunk
Private Const PAGE_RWX      As Long = &H40                                  'Make allocated memory executable
Private Const MEM_COMMIT    As Long = &H1000                                'Commit allocated memory
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 0                                     'Shutdown flag index into the machine-code thunk
Private Const IDX_NWNDPROC  As Long = 1                                     'Used internally in the thunk, tracks original WndProc recursion
Private Const IDX_NCALLBACK As Long = 2                                     'Used internally in the thunk, tracks original callback recursion
Private Const IDX_HWND      As Long = 3                                     'hWnd index into the machine-code thunk
Private Const IDX_EBMODE    As Long = 4                                     'EbMode function address index into the machine-code thunk
Private Const IDX_CWP       As Long = 5                                     'CallWindowProc function address index into the machine-code thunk
Private Const IDX_SWL       As Long = 6                                     'SetWindowsLong function address index into the machine-code thunk
Private Const IDX_FREE      As Long = 7                                     'VirtualFree function address index into the machine-code thunk
Private Const IDX_ME        As Long = 8                                     'ObjPtr(Me) address index into the machine-code thunk
Private Const IDX_WNDPROC   As Long = 9                                     'Original WndProc address index into the machine-code thunk
Private Const IDX_CALLBACK  As Long = 10                                    'zWndProc address index into the machine-code thunk
Private Const IDX_BTABLE    As Long = 11                                    'Before table address index into the machine-code thunk
Private Const IDX_ATABLE    As Long = 12                                    'After table address index into the machine-code thunk
Private Const IDX_EBX       As Long = 15                                    'Data address index into the machine-code thunk

Private z_Code(34)          As Currency                                     'The subclass thunk's machine-code is initialised here
Private z_Data(32)          As Long                                         'Array whose data pointer is re-mapped to arbitary memory addresses
Private z_DataDataPtr       As Long                                         'Address of z_Data()'s SafeArray data pointer
Private z_DataOrigData      As Long                                         'Address of z_Data()'s original data
Private z_hWnds             As Collection                                   'Window handle/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)

'-cSubclass code----------------------------------------------------------------------------------
Private Sub Class_Initialize()                                              'Class initialization
  RtlMoveMemory VarPtr(z_DataDataPtr), VarPtrArray(z_Data), 4               'Get the address of z_Data()'s SafeArray header
  z_DataDataPtr = z_DataDataPtr + 12                                        'Bump the address to point to the pvData data pointer
  RtlMoveMemory VarPtr(z_DataOrigData), z_DataDataPtr, 4                    'Get the value of z_Data()'s SafeArray pvData data pointer

  'Initialise the machine-code thunk
  z_Code(6) = -330064573519993.2416@: z_Code(7) = 131176846802010.352@: z_Code(8) = 117184300020928.1512@: z_Code(9) = 14200749441411.2628@: z_Code(10) = 4011464988223.474@: z_Code(11) = -151320947323777.8432@: z_Code(12) = -171686151974904.1888@: z_Code(13) = 576750112840.0921@: z_Code(14) = 353082210901452.7861@: z_Code(15) = 539529232.6866@: z_Code(16) = -208529527005181.5168@: z_Code(17) = 34438797751651.6657@: z_Code(18) = -6207657131060.2869@: z_Code(19) = 291551690021556.6453@: z_Code(20) = -826407423118013.5425@: z_Code(21) = -6292285434835.2443@: z_Code(22) = 59785192422972.0131@: z_Code(23) = 850256790485906.9675@: z_Code(24) = -6179475612041.2372@: z_Code(25) = 146479469761128.9715@: z_Code(26) = 428830235516392.5503@: z_Code(27) = 843074343097751.6544@: z_Code(28) = 640058813842636.8985@: z_Code(29) = 89734221035447.6031@: z_Code(30) = 502803189117959.0655@: z_Code(31) = 347840509002750.366@: z_Code(32) = -856577002795587.9936@: z_Code(33) = -428647568023007.1227@: z_Code(34) = 47624.986@
  
  zMap VarPtr(z_Code(0))                                                    'Map the address of z_Code()'s first element to the z_Data() array
  z_Data(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                            'Store the EbMode function address in the subclass thunk machine-code
  z_Data(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store the CallWindowProc function address in the subclass thunk machine-code
  z_Data(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the subclass thunk machine-code
  z_Data(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the SetWindowLong function address in the subclass thunk machine-code
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
  
  Set z_hWnds = New Collection                                              'Create the window-handle/thunk-memory-address collection
End Sub

'Class termination
Private Sub Class_Terminate()
  UnSubAll                                                                  'Unsubclass all existing subclassing
  Set z_hWnds = Nothing                                                     'Destroy the window-handle/thunk-address collection
End Sub

'Subclass the specified window handle
Public Function Subclass(ByVal lng_hWnd As Long, ByVal oOwner As iSubclass) As Boolean
  Dim nAddr     As Long
  Dim nCallback As Long

  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError "Subclass", "Invalid window handle"
  End If

  RtlMoveMemory VarPtr(nCallback), ObjPtr(oOwner), 4                        'Get the address of my vTable
  zMap nCallback + &H1C                                                     'Map the first implemented interface
  nCallback = z_Data(0)                                                     'Calculate the address of the first implemented interface

  nAddr = VirtualAlloc(ByVal 0&, MEM_LEN, MEM_COMMIT, PAGE_RWX)             'Allocate memory
  RtlMoveMemory nAddr, VarPtr(z_Code(0)), CODE_LEN                          'Copy the machine-code to the allocated memory

  On Error GoTo Catch                                                       'Catch double subclassing
    z_hWnds.Add nAddr, "h" & lng_hWnd                                       'Add the hWnd/thunk-address to the collection
  On Error GoTo 0

  zMap nAddr                                                                'Map z_Data() to the subclass thunk machine-code
  z_Data(IDX_EBX) = nAddr                                                   'Patch the data address
  z_Data(IDX_HWND) = lng_hWnd                                               'Store the window handle
  z_Data(IDX_BTABLE) = nAddr + CODE_LEN                                     'Store the address of the before table
  z_Data(IDX_ATABLE) = z_Data(IDX_BTABLE) + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table
  z_Data(IDX_ME) = ObjPtr(oOwner)                                           'Store the Me object address in the subclass thunk machine-code
  z_Data(IDX_CALLBACK) = nCallback                                          'Store the zWndProc address in the subclass thunk machine-code
  z_Data(IDX_WNDPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, nAddr + _
                                                              WNDPROC_OFF)  'Set the WndProc
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
  Subclass = True                                                           'Indicate success
  Exit Function                                                             'Exit

Catch:
  zError "Subclass", "Window handle is already subclassed"
End Function

'Terminate all subclassing
Public Sub UnSubAll()
  Dim i     As Long
  Dim nAddr As Long

  With z_hWnds
    For i = .Count To 1 Step -1                                             'Loop through the collection of window handles in reverse order
      nAddr = .Item(i)                                                      'Map z_Data() to the window handle's thunk address
      If IsBadCodePtr(nAddr) = 0 Then                                       'Ensure that the thunk hasn't freed itself
        zMap nAddr                                                          'Map the thunk memory to the z_Data() array
        Unsubclass z_Data(IDX_HWND)                                         'Unsubclass
      End If
    Next i                                                                  'Next member of the collection
        
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
  End With
End Sub

'Unsubclass the specified window handle
Public Sub Unsubclass(ByVal lng_hWnd As Long)
  Dim nAddr As Long
  Dim sKey  As String
  
  With z_hWnds
    On Error GoTo Catch
    
    sKey = "h" & lng_hWnd                                                   'Construct the window handle collection key
    nAddr = .Item(sKey)                                                     'Lookup the thunk memory address associated with the window handle
    DelMsg lng_hWnd, ALL_MESSAGES, MSG_BEFORE_AFTER                         'Delete all messages
    zMap nAddr                                                              'Map the thunk memory to the z_Data() array
    z_Data(IDX_SHUTDOWN) = -1                                               'Set the thunk shutdown flag
    CallWindowProcA nAddr + WNDPROC_OFF, 0, 0, 0, 0                         'Ensures that Unsubclassing from a Form/UserControl's _Terminate event will cause the thunk to release itself
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
    .Remove sKey                                                            'Remove the specified window handle from the collection
  End With
  
  Exit Sub
  
Catch:
  zError "Unsubclass", "Window handle isn't subclassed", False
End Sub

'Add the message value to the window handle's specified callback table
Public Sub AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen_iSC = eMsgWhen_iSC.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be added to the before original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Add the message to the before table
  End If

  If When And MSG_AFTER Then                                                'If message is to be added to the after original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Add the message to the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Delete the message value from the window handle's specified callback table
Public Sub DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen_iSC = eMsgWhen_iSC.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be deleted from the before original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Delete the message from the before table
  End If

  If When And MSG_AFTER Then                                                'If the message is to be deleted from the after original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Delete the message from the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Call the original WndProc
Public Function CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  CallOrigWndProc = CallWindowProcA(z_Data(IDX_WNDPROC), lng_hWnd, uMsg, _
                                                         wParam, lParam)    'Call the original WndProc of the passed window handle parameter
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Function

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = z_Data(0)                                                      'Get the current table entry count

    If nCount >= MSG_ENTRIES Then                                           'Check that we aren't about to overflow the message table
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values", False
      Exit Sub
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = 0 Then                                                 'If the element is free...
        z_Data(i) = uMsg                                                    'Use this element
        Exit Sub                                                            'Bail
      ElseIf z_Data(i) = uMsg Then                                          'If the message is already in the table...
        Exit Sub                                                            'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    z_Data(nCount) = uMsg                                                   'Store the message in the appended table entry
  End If

  z_Data(0) = nCount                                                        'Store the new table entry count
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    z_Data(0) = 0                                                           'Zero the table entry count
  Else
    nCount = z_Data(0)                                                      'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = uMsg Then                                              'If the message is found...
        z_Data(i) = 0                                                       'Null the msg value -- also frees the element for re-use
        Exit Sub                                                            'Exit
      End If
    Next i                                                                  'Next message table entry
    
    zError "Message " & "&H" & Hex$(uMsg) & " not found in table", False
  End If
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String, Optional bEnd As Boolean = True)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  
  MsgBox sMsg & ".", IIf(bEnd, vbCritical, vbExclamation) + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
  
  If bEnd Then
    ' Nothing?
  End If
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map z_Data() to the specified address
Private Sub zMap(ByVal nAddr As Long)
  RtlMoveMemory z_DataDataPtr, VarPtr(nAddr), 4                             'Set z_Data()'s SafeArray data pointer to the specified address
End Sub

'Map z_Data() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  On Error GoTo Catch                                                       'Catch unsubclassed window handles
  
  zMap_hWnd = z_hWnds("h" & lng_hWnd)                                       'Get the thunk address
  zMap zMap_hWnd                                                            'Map z_Data() to the thunk address
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function
