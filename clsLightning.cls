VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLightning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'***************************************************************************************
'*  Lightning!   1.6.5 Standalone Registry Control Class                               *
'*                                                                                     *
'*  Created:     November 13, 2005                                                     *
'*  Updated:     April 12, 2006                                                        *
'*  Purpose:     Comprehensive Registry Control Class                                  *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.6.5                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  Throughout Project                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************

'/~  List of exposed functions ~/

'/~  Value Types ~/
'/~  Access_Check - Test user access rights
'/~  Read_BEndian - read a big endian value
'/~  Write_BEndian - write a big_endian value
'/~  Read_Binary - read a binary value
'/~  Write_Binary - write a binary value
'/~  Read_Dword - read a dword value
'/~  Write_Dword - write a dword value
'/~  Read_Link - read a binary link value
'/~  Write_Link - write a binary link value
'/~  List_Data - list all data blocks in a subkeys values
'/~  List_Values - puts all of a keys values into a collection
'/~  Read_LEndian - read a little endian value
'/~  Write_LEndian - write a little_endian value
'/~  Read_Multi - read a multi_sz value
'/~  Write_Multi - write a multi_sz value
'/~  Read_MultiCN - reads a multi value and splits into a collection
'/~  Write_MultiCN - converts a collection into a multi_sz value
'/~  Write_Qword - write a 64bit dword number
'/~  Read_Qword - read a 64bit dword number
'/~  Read_ResDescriptor - read hardware resource description (hex)binary
'/~  Write_ResDescriptor = write hardware resource description (hex)binary
'/~  Read_ResourceList - read a hardware resource list (hex)binary
'/~  Write_ResourceList - write to a hardware resource list (hex)binary
'/~  Read_ResRequired - read a hardware resource requirements list (hex)binary
'/~  Write_ResRequired - write to a hardware resource requirements list (hex)binary
'/~  Read_String - read a string(sz) or expanded string(expand_sz)
'/~  Write_String - write a string value
'/~  Write_Expanded - write an expanded string value

'/~  Key Related Functions ~/
'/~  List_Keys - puts all subkeys under specified branch into a collection
'/~  Key_Exists - test if key exists
'/~  Create_Key - create a new key
'/~  Delete_Key - delete a key

'/~  Value Related Functions ~/
'/~  Write_Value - write value types: 1)sz 2)expand_sz 3)multi_sz 4)binary 5)dword 6)little_endian 7)big_endian
'/~  Delete_Value - delete a value
'/~  Search_Value - search for a value under the key

'/~  Conversion Routines ~/
'/~  Make_LEndian16 - convert integer to 16bit little_endian
'/~  Make_LEndian32 - convert long to 32bit little_endian
'/~  Make_BEndian32 - convert long to big endian format

'/~  Error Logging ~/
'/~  Get_Error - interprets errors passed from dll
'/~  Log_Error - sends errors to a log file

Public Event ErrorCond(ByVal sRoutine As String, ByVal sKey As String, ByVal sError As String)

'/* time structure
Private Type FILETIME
    dwLowDateTime                              As Long
    dwHighDateTime                             As Long
End Type

'/* security structure
Private Type SECURITY_ATTRIBUTES
    nLength                                    As Long
    lpSecurityDescriptor                       As Long
    bInheritHandle                             As Boolean
End Type

'key constants
Public Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'/* value types
Public Enum Reg_Type
    REG_NONE = 0                            '/* No value type
    REG_SZ = 1                              '/* Unicode NULL terminated string
    REG_EXPAND_SZ = 2                       '/* Unicode NULL terminated string
    REG_BINARY = 3                          '/* Binary data
    REG_DWORD = 4                           '/* 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4             '/* 32-bit number
    REG_DWORD_BIG_ENDIAN = 5                '/* 32-bit number, high byte first
    REG_LINK = 6                            '/* Unicode symbolic link
    REG_MULTI_SZ = 7                        '/* Array of Unicode strings
    REG_RESOURCE_LIST = 8                   '/* Hardware resource description
    REG_FULL_RESOURCE_DESCRIPTOR = 9        '/* Hardware resource description
    REG_RESOURCE_REQUIREMENTS_LIST = 10     '/* Resource requirements
    REG_QWORD_LITTLE_ENDIAN = 11            '/* 64bit dword value
End Enum

'/* access paramaters
Private Const KEY_ALL_ACCESS               As Long = &HF003F
Private Const KEY_CREATE_LINK              As Long = &H20
Private Const KEY_CREATE_SUB_KEY           As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Private Const KEY_EXECUTE                  As Long = &H20019
Private Const KEY_NOTIFY                   As Long = &H10
Private Const KEY_QUERY_VALUE              As Long = &H1
Private Const KEY_READ                     As Long = &H20019
Private Const KEY_SET_VALUE                As Long = &H2
Private Const KEY_WRITE                    As Long = &H20006
Private Const REG_OPTION_NON_VOLATILE      As Long = &H0
Private Const REG_ERR_OK                   As Long = &H0
Private Const REG_ERR_NOT_EXIST            As Long = &H1
Private Const REG_ERR_NOT_STRING           As Long = &H2
Private Const REG_ERR_NOT_DWORD            As Long = &H4

'/* error handling
Private Const ERROR_NONE                   As Long = &H0
Private Const ERROR_BADDB                  As Long = &H1
Private Const ERROR_BADKEY                 As Long = &H2
Private Const ERROR_CANTOPEN               As Long = &H3
Private Const ERROR_CANTREAD               As Long = &H4
Private Const ERROR_CANTWRITE              As Long = &H5
Private Const ERROR_OUTOFMEMORY            As Long = &H6
Private Const ERROR_ARENA_TRASHED          As Long = &H7
Private Const ERROR_ACCESS_DENIED          As Long = &H8
Private Const ERROR_INVALID_PARAMETERS     As Long = &H57
Private Const ERROR_MORE_DATA              As Long = &HEA
Private Const ERROR_NO_MORE_ITEMS          As Long = &H103

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal length As Long)
                                                                     
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
                                                                                
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long
                                                                                      
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, _
                                                                                ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, _
                                                                                    ByVal lpValueName As String) As Long

Private Declare Function RegCreatekey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, _
                                                                                    ByVal lpSubKey As String, _
                                                                                    ByVal Reserved As Long, _
                                                                                    ByVal lpClass As String, _
                                                                                    ByVal dwOptions As Long, _
                                                                                    ByVal samDesired As Long, _
                                                                                    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                    phkResult As Long, _
                                                                                    lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  lpData As Any, _
                                                                                  ByVal cbData As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, _
                                                                                ByVal dwIndex As Long, _
                                                                                ByVal lpName As String, _
                                                                                lpcbName As Long, _
                                                                                lpReserved As Long, _
                                                                                ByVal lpClass As String, _
                                                                                lpcbClass As Long, _
                                                                                lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, _
                                                                                ByVal dwIndex As Long, _
                                                                                ByVal lpValueName As String, _
                                                                                lpcbValueName As Long, _
                                                                                ByVal lpReserved As Long, _
                                                                                lpType As Long, _
                                                                                lpData As Byte, _
                                                                                lpcbData As Long) As Long

Private m_bIntercept    As Boolean
Private m_bLogging      As Boolean
Private m_bNotify       As Boolean


Public Property Let p_Intercept(PropVal As Boolean)
    '/* enable error interception
    m_bIntercept = PropVal
End Property

Public Property Let p_Logging(PropVal As Boolean)
    '/* enable error logging
    m_bLogging = PropVal
End Property

Public Property Let p_Notify(PropVal As Boolean)
    '/* enable error notification
    m_bNotify = PropVal
End Property


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                              PROCESSING CORE
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Public Function Access_Test(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Boolean

    '/* test user access

    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security attributes
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If lRetVal = ERROR_NONE Then
        Access_Test = True
    End If

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Access_Test", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Function Read_BEndian(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As String

    '/* read an big_endian value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim sBuffer         As String
    Dim slength         As Long

    On Error GoTo Handler

    '/* open root key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler
    
    '/* create a buffer
    sBuffer = Space$(255)
    slength = 255
    
    '/* query key for string value
    lRetVal = RegQueryValueEx(lHKey, Value, 0, REG_DWORD_BIG_ENDIAN, ByVal sBuffer, slength)
    
    '/* remove null terminator and add value
    If lRetVal = ERROR_NONE Then
        sBuffer = Left$(sBuffer, slength - 1)
        Read_BEndian = sBuffer
    End If

Handler:
    '/* close key and set result
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_BEndian", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_BEndian(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, _
       ByVal vData As Variant)

    '/* write a big endian value (32b number)
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* required security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key and test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write value, set error, and close key
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_DWORD_BIG_ENDIAN, vData, 4)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_BEndian", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_Binary(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Variant

    '/* read a binary value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim byBuffer()      As Byte
    Dim lBuffersize     As Long
    Dim sTemp           As String
    Dim i               As Long

    On Error GoTo Handler

    '/* open root key and test for value type
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler
    
    '/* get buffer size
    lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_BINARY, ByVal 0&, lBuffersize)
    
    '/* read into buffer
    If lRetVal = ERROR_NONE Then
        ReDim byBuffer(lBuffersize - 1) As Byte
        lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_BINARY, byBuffer(0), lBuffersize)
        '/* format value
        'For i = 0 To UBound(byBuffer)
        '    sTemp = sTemp & Format$(Trim$(Hex$(byBuffer(i))), "0#")
        'Next i
        '/* set value
        Read_Binary = byBuffer 'LTrim$(sTemp)
    End If

Handler:
    '/* set error and close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_Binary", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_Binary(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant, _
       ByRef bData() As Byte)

    '/* write a binary value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler
    
    '/* write binary, set error, and close
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_BINARY, bData(0), UBound(bData) + 1)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_Binary", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_DWord(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Long

    '/* read a dword value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lBuffer         As Long

    On Error GoTo Handler

    '/* open root key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler
    
    '/* query value
    lRetVal = RegQueryValueEx(lHKey, Value, 0, REG_DWORD, lBuffer, 4)
    
    '/* set value
    If lRetVal = ERROR_NONE Then
        Read_DWord = lBuffer
    Else
        Read_DWord = 0
    End If

Handler:
    '/* close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_Binary", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_DWord(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, _
       ByVal lData As Long)

    '/* write a DWORD value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* required security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key and test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler
    
    '/* write value, set error, and close key
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_DWORD, lData, 4)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_DWord", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_Link(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Variant

    '/* read a link value
    Dim lHKey               As Long
    Dim lRetVal             As Long
    Dim byBuffer()          As Byte
    Dim lBuffersize         As Long

    On Error GoTo Handler

    '/* open root key and test for value type
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* get buffer size
    lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_LINK, ByVal 0&, lBuffersize)
    
    '/* read into buffer
    If lRetVal = ERROR_NONE Then
        ReDim byBuffer(lBuffersize - 1) As Byte
        lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_LINK, byBuffer(0), lBuffersize)
        '/* set value
        Read_Link = byBuffer
    End If

Handler:
    '/* set error and close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_Link", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_Link(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant, _
       ByRef bData() As Byte)

    '/* write a link value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write link, set error, and close
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_LINK, bData(0), UBound(bData) + 1)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_Link", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_LEndian(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Long

    '/* read a dword value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lBuffer         As Long

    On Error GoTo Handler

    '/* open root key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* query value
    lRetVal = RegQueryValueEx(lHKey, Value, 0, REG_DWORD_LITTLE_ENDIAN, lBuffer, 4)
    
    '/* set value
    If lRetVal = ERROR_NONE Then
        Read_LEndian = lBuffer
    Else
        Read_LEndian = 0
    End If

Handler:
    '/* close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_LEndian", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_LEndian(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, _
       ByVal vData As Variant)

    '/* write a little endian value (same as dword)

    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* required security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key and test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write value, set error, and close key
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_DWORD_LITTLE_ENDIAN, vData, 4)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_LEndian", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_Multi(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String) As String

    '/* read a multi_sz value

    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim sBuffer         As String
    Dim length          As Long
    Dim resString       As String
    Dim resBinary()     As Byte

    On Error GoTo Handler

    '/* open root key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, &HF003F, lHKey)
    If lRetVal = ERROR_NONE Then
        '/* get buffer size
        length = 1024
        ReDim resBinary(0 To length - 1) As Byte
        lRetVal = RegQueryValueEx(lHKey, Value, 0, &H7, resBinary(0), length)
        '/* if more data, resize our byte array
        If lRetVal = &HEA Then
            ReDim resBinary(0 To length - 1) As Byte
            lRetVal = RegQueryValueEx(lHKey, Value, 0, &H7, resBinary(0), length)
        End If
        '/* read into buffer
        If lRetVal = ERROR_NONE Then
            resString = Space$(length - 2)
            CopyMemory ByVal resString, resBinary(0), length - 2
            sBuffer = resString
            If Len(Trim_Null(sBuffer)) > 0 Then
                Read_Multi = resString
            End If
        End If
    End If

Handler:
    '/* set error and close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_Multi", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_Multi(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, _
       ByVal sData As String)

    '/* write a multi-sz value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, &H20006, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write multi, set error, and close key
    lRetVal = RegSetValueEx(lHKey, Value, 0, &H7, ByVal sData, Len(sData))

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_Multi", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_MultiCN(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal KeyVal As String) As Collection

    '//process multi_sz value
    '//and return collection
    Dim aTemp()         As String
    Dim cTemp           As New Collection
    Dim sTemp           As String
    Dim l               As Long
    Dim lReturn         As Long

    On Error GoTo Handler

    Set cTemp = New Collection
    '//get multi string
    sTemp = Read_Multi(RootKey, SubKey, KeyVal)
    If Len(sTemp) = 0 Then GoTo Handler
    
    '//get values and add to collection
    aTemp = Split(sTemp, vbNullChar)
    For l = 0 To UBound(aTemp)
        cTemp.Add aTemp(l)
    Next l
    
Handler:
    '//test for error
    If Not lReturn = 0 Then
        If m_bIntercept Then
            Error_State lReturn, "Read_MultiCN", SubKey
        End If
    Else
        '//set collection and release library
        Set Read_MultiCN = cTemp
        Set cTemp = Nothing
    End If

End Function

Public Sub Write_MultiCN(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal KeyVal As String, _
       ByVal DataSet As Collection)

    '//process a collection and
    '//write to registry as a
    '//multi_sz entry
    Dim sTemp           As String
    Dim sValue          As String
    Dim cTemp           As New Collection
    Dim l               As Long
    Dim lRetVal         As Long
    Dim Item            As Variant

    On Error GoTo Handler

    For Each Item In DataSet
        sTemp = sTemp & Item & vbNullChar
    Next
    '//add terminating null char
    sTemp = sTemp & vbNullChar
    '//write to registry
    Write_Multi RootKey, SubKey, KeyVal, sTemp
    
Handler:
    On Error GoTo 0

End Sub

Public Function Read_QWord(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Currency

    '/* read a qword value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lBuffer         As Long

    On Error GoTo Handler

    '/* open root key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* query value
    lRetVal = RegQueryValueEx(lHKey, Value, 0, REG_QWORD_LITTLE_ENDIAN, lBuffer, 4)
    '/* set value
    If lRetVal = ERROR_NONE Then
        Read_QWord = lBuffer
    Else
        Read_QWord = 0
    End If

Handler:
    '/* close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_QWord", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_QWord(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, _
       ByVal cData As Currency)

    '/* write a DWORD value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* required security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key and test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write value, set error, and close key
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_QWORD_LITTLE_ENDIAN, cData, 4)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_QWord", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_ResourceList(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Variant

    '/* read a resource list value(binary)
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim byBuffer()      As Byte
    Dim lBuffersize     As Long

    On Error GoTo Handler

    '/* open root key and test for value type
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* get buffer size
    lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_RESOURCE_LIST, ByVal 0&, lBuffersize)
    
    '/* read into buffer
    If lRetVal = ERROR_NONE Then
        ReDim byBuffer(lBuffersize - 1) As Byte
        lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_RESOURCE_LIST, byBuffer(0), lBuffersize)
        '/* set value
        Read_ResourceList = byBuffer
    End If

Handler:
    '/* set error and close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_ResourceList", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_ResourceList(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant, _
       ByRef bData() As Byte)

    '/* write a resource list value(binary)
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write resource, set error, and close
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_RESOURCE_LIST, bData(0), UBound(bData) + 1)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_ResourceList", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_ResDescriptor(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Variant

    '/* read a resource list value(binary)
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim byBuffer()      As Byte
    Dim lBuffersize     As Long

    On Error GoTo Handler

    '/* open root key and test for value type
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* get buffer size
    lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_FULL_RESOURCE_DESCRIPTOR, ByVal 0&, lBuffersize)
    
    '/* read into buffer
    If lRetVal = ERROR_NONE Then
        ReDim byBuffer(lBuffersize - 1) As Byte
        lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_FULL_RESOURCE_DESCRIPTOR, byBuffer(0), lBuffersize)
        '/* set value
        Read_ResDescriptor = byBuffer
    End If

Handler:
    '/* set error and close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_ResDescriptor", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_ResDescriptor(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant, _
       ByRef bData() As Byte)

    '/* write a resource list value(binary)
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write resource, set error, and close
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_FULL_RESOURCE_DESCRIPTOR, bData(0), UBound(bData) + 1)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_ResDescriptor", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_ResRequired(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Variant

    '/* read a resource requirements value(binary)
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim byBuffer()      As Byte
    Dim lBuffersize     As Long

    On Error GoTo Handler

    '/* open root key and test for value type
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* get buffer size
    lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_RESOURCE_REQUIREMENTS_LIST, ByVal 0&, lBuffersize)
    
    '/* read into buffer
    If lRetVal = ERROR_NONE Then
        ReDim byBuffer(lBuffersize - 1) As Byte
        lRetVal = RegQueryValueEx(lHKey, Value, 0&, REG_RESOURCE_REQUIREMENTS_LIST, byBuffer(0), lBuffersize)
        '/* set value
        Read_ResRequired = byBuffer
    End If

Handler:
    '/* set error and close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_ResRequired", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_ResRequired(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant, _
       ByRef bData() As Byte)

    '/* write a resource requirements value(binary)
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security structure
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* test access
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write resource, set error, and close
    lRetVal = RegSetValueEx(lHKey, Value, 0&, REG_RESOURCE_REQUIREMENTS_LIST, bData(0), UBound(bData) + 1)

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_ResRequired", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Read_String(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, Optional BufferSize As Long = 255) As String

    '/* read an SZ value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim sBuffer         As String
    Dim slength         As Long

    On Error GoTo Handler

    '/* open root key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler
    
retry:
    '/* create a buffer
    sBuffer = Space$(BufferSize)
    '/* query key for string value
    lRetVal = RegQueryValueEx(lHKey, Value, 0, REG_SZ, ByVal sBuffer, BufferSize)
    '/* read string into buffer
    If lRetVal = ERROR_NONE Then
        sBuffer = Left$(sBuffer, BufferSize - 1)
        Read_String = sBuffer
    ElseIf lRetVal = 234 Then
        'BufferSize = BufferSize + 255
        GoTo retry
    End If

Handler:
    '/* close key and set result
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Read_String", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Sub Write_String(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, _
       ByVal Data As String)

    '/* write an SZ value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security attributes
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write string, set error, and close
    lRetVal = RegSetValueEx(lHKey, Value, 0, REG_SZ, ByVal Data, Len(Data))

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_String", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Sub Write_Expanded(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String, _
       ByVal sData As String)

    '/* write an SZ value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim lDeposit        As Long
    Dim tSecAttrib      As SECURITY_ATTRIBUTES

    On Error GoTo Handler

    '/* security attributes
    With tSecAttrib
        .nLength = Len(tSecAttrib)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key
    lRetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, tSecAttrib, lHKey, lDeposit)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* write string, set error, and close
    lRetVal = RegSetValueEx(lHKey, Value, 0, REG_EXPAND_SZ, ByVal sData, Len(sData))

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Write_Expanded", SubKey
        End If
    End If
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Sub

Public Function Key_Exist(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Boolean

    '/* check for key
    Dim lHKey           As Long
    Dim lRetVal         As Long

    On Error GoTo Handler

    '/* if 0 returned, key is valid
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_QUERY_VALUE, lHKey)
    If lRetVal = ERROR_NONE Then
        Key_Exist = True
    End If

Handler:
    RegCloseKey lHKey
    On Error GoTo 0

End Function

Public Function Value_Exist(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As String) As Boolean

    '/* read an SZ value
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim sBuffer         As String
    Dim slength         As Long
    Dim DataType        As Long

    On Error GoTo Handler

    '/* open root key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, lHKey)
    If lRetVal = ERROR_NONE Then
        '/* create a buffer
        sBuffer = Space$(255)
        slength = 255
        '/* query key for string value
        lRetVal = RegQueryValueEx(lHKey, Value, 0, DataType, ByVal sBuffer, slength)
        '/* read string into buffer
        If lRetVal = ERROR_NONE Or lRetVal = 234 Then
            Value_Exist = True
        End If
    End If

Handler:
    '/* close key and set result
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Function Create_Key(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Boolean

    '/* create a new key
    Dim lHKey           As Long
    Dim lRetVal         As Long

    'On Error GoTo Handler

    '/* create the key
    lRetVal = RegCreatekey(RootKey, SubKey, lHKey)
    Debug.Print "lrv "; lRetVal
Handler:
    '/* set error and close
    Create_Key = (lRetVal = 0)
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Function Delete_Key(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Boolean

    '/* delete a key
    Dim lRetVal         As Long

    On Error GoTo Handler

    '/* delete key
    lRetVal = RegDeleteKey(RootKey, SubKey)
    '/* set error
    Delete_Key = (lRetVal = 0)

Handler:
    On Error GoTo 0

End Function

Public Function Delete_Value(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal Value As Variant) As Boolean

    '/* delete a value
    Dim lRetVal         As Long
    Dim handle          As Long

    On Error GoTo Handler

    '/* open key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, handle)
    
    '/* exit on error
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    '/* delete value
    lRetVal = RegDeleteValue(handle, Value)
    Delete_Value = (lRetVal = 0)
    
Handler:
    '/* set error and close key
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Delete_Value", SubKey
        End If
    End If
    lRetVal = RegCloseKey(handle)
    On Error GoTo 0

End Function

Public Function List_Keys(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Collection

    '/* list all keys and add to collection
    Dim KeyName         As String
    Dim keylen          As Long
    Dim classname       As String
    Dim classlen        As Long
    Dim lastwrite       As FILETIME
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim Index           As Long
    Dim cTemp           As New Collection

    On Error GoTo Handler

    Set cTemp = New Collection
    '/* open key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, lHKey)
    If Not lRetVal = ERROR_NONE Then
        Set List_Keys = Nothing
        GoTo Handler
    End If
    Index = 0
    
    '/* loop through keys and add to collection
    Do
        KeyName = Space$(255)
        keylen = 255
        classname = Space$(255)
        classlen = 255
        lRetVal = RegEnumKeyEx(lHKey, Index, KeyName, keylen, ByVal 0, classname, classlen, lastwrite)
        If lRetVal = ERROR_NONE Then
            KeyName = Left$(KeyName, keylen)
            cTemp.Add KeyName
        End If
        Index = Index + 1
    Loop Until Not lRetVal = 0

Handler:
    '/* set collection and close
    If cTemp.Count > 0 Then
        Set List_Keys = cTemp
    Else
        If m_bIntercept Then
            Error_State lRetVal, "List_Keys", SubKey
        End If
        Set List_Keys = Nothing
    End If
    Set cTemp = Nothing
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Function ColList_Keys(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Collection

    '/* list all keys and add to collection
    Dim KeyName         As String
    Dim keylen          As Long
    Dim classname       As String
    Dim classlen        As Long
    Dim lastwrite       As FILETIME
    Dim lHKey           As Long
    Dim lRetVal         As Long
    Dim Index           As Long
    Dim cTemp           As New Collection

    On Error GoTo Handler

    Set ColList_Keys = New Collection
    '/* open key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    Index = 0
    
    '/* loop through keys and add to collection
    Do
        KeyName = Space$(255)
        keylen = 255
        classname = Space$(255)
        classlen = 255
        lRetVal = RegEnumKeyEx(lHKey, Index, KeyName, keylen, ByVal 0, classname, classlen, lastwrite)
        If lRetVal = ERROR_NONE Then
            KeyName = Left$(KeyName, keylen)
            ColList_Keys.Add KeyName
        End If
        Index = Index + 1
    Loop Until lRetVal <> 0

Handler:
    '/* set collection and close
    If cTemp.Count > 0 Then
        Set ColList_Keys = cTemp
    Else
        If m_bIntercept Then
            Error_State lRetVal, "List_Keys", SubKey
        End If
        Set ColList_Keys = Nothing
    End If
    Set cTemp = Nothing
    lRetVal = RegCloseKey(lHKey)
    On Error GoTo 0

End Function

Public Function List_Values(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Collection

    '/* list all values and add to a collection
    Dim Value           As String
    Dim ValueLen        As Long
    Dim DataType        As Long
    Dim Data(0 To 254)  As Byte
    Dim DataLen         As Long
    Dim lHKey           As Long
    Dim Index           As Long
    Dim lRetVal         As Long
    Dim cTemp           As New Collection

    On Error GoTo Handler

    Set cTemp = New Collection
    '/* open key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_QUERY_VALUE, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    Index = 0
    '/* loop through values and add to collection
    Do
        DataLen = 255
        Value = Space$(255)
        ValueLen = 255
        lRetVal = RegEnumValue(lHKey, Index, Value, ValueLen, 0, DataType, Data(0), DataLen)
        If lRetVal = ERROR_NONE Then
            Value = Left$(Value, ValueLen)
            Select Case DataType
                Case REG_SZ, REG_EXPAND_SZ
                    cTemp.Add Value
                Case REG_MULTI_SZ
                    cTemp.Add Value
                Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
                    cTemp.Add Value
                Case REG_BINARY
                    cTemp.Add Value
            End Select
        End If
        Index = Index + 1
    Loop Until Not lRetVal = 0

    '/* set error, collection, and close key
    Set List_Values = cTemp
    Set cTemp = Nothing
    lRetVal = RegCloseKey(lHKey)

    Exit Function

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Delete_Value", SubKey
        End If
        Set List_Values = Nothing
    End If
    On Error GoTo 0

End Function

Public Function List_Data(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String) As Collection

    '/* list all data for values in a subkey
    Dim Value           As String
    Dim ValueLen        As Long
    Dim DataType        As Long
    Dim Data(0 To 254)  As Byte
    Dim DataLen         As Long
    Dim lHKey           As Long
    Dim Index           As Long
    Dim lRetVal         As Long
    Dim cTemp           As New Collection

    On Error GoTo Handler

    Set cTemp = New Collection
    '/* open key
    lRetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_QUERY_VALUE, lHKey)
    If Not lRetVal = ERROR_NONE Then GoTo Handler

    Index = 0
    '/* loop through values and add to collection
    Do
        DataLen = 255
        Value = Space$(255)
        ValueLen = 255
        lRetVal = RegEnumValue(lHKey, Index, Value, ValueLen, 0, DataType, Data(0), DataLen)
        If lRetVal = ERROR_NONE Then
            Value = Left$(Value, ValueLen)
            Select Case DataType
                Case REG_SZ, REG_EXPAND_SZ
                    cTemp.Add Read_String(RootKey, SubKey, Value), Value
                Case REG_MULTI_SZ
                    cTemp.Add Read_Multi(RootKey, SubKey, Value), Value
                Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
                    cTemp.Add Read_DWord(RootKey, SubKey, Value), Value
                Case REG_BINARY
                    cTemp.Add Read_Binary(RootKey, SubKey, Value), Value
            End Select
        End If
        Index = Index + 1
    Loop Until Not lRetVal = 0

    '/* set error, collection, and close key
    Set List_Data = cTemp
    Set cTemp = Nothing
    lRetVal = RegCloseKey(lHKey)

    Exit Function

Handler:
    If lRetVal <> 0 Then
        If m_bIntercept Then
            Error_State lRetVal, "Delete_Value", SubKey
        End If
        Set List_Data = Nothing
    End If
    On Error GoTo 0

End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           PERIPHERAL ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Search_Value(ByVal RootKey As HKEY_Type, _
       ByVal KeyVal As String, _
       ByVal SearchVal As String) As Boolean

    '/* search for a value
    Dim Item            As Variant
    Dim lReturn         As Long

    On Error GoTo Handler
    
    '/* enumerate values and compare
    '/* to search item
    For Each Item In List_Values(RootKey, KeyVal)
        If LCase$(Item) = SearchVal Then
            Search_Value = True
            Exit For
        End If
    Next Item
    
    '/* log errors
    If Not lReturn = 0 Then Get_Error lReturn
        
Handler:
    On Error GoTo 0
    
End Function

Public Function Value_Exists(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal KeyVal As String) As Boolean

    '/* test for key
    On Error GoTo Handler

    '/* if no error then key exists
    If Value_Exist(RootKey, SubKey, KeyVal) Then
        Value_Exists = True
    End If

Handler:
    On Error GoTo 0

End Function

Public Sub Write_Value(ByVal RootKey As HKEY_Type, _
       ByVal SubKey As String, _
       ByVal KeyVal As String, _
       ByVal ValData As String, _
       ByVal DataType As Integer)

    '/* create a new value with data
    '/* supports sz, multi_sz, expand_sz
    '/* binary, dword, little_endian
    '/* big_endian and link
    Dim j               As Integer
    Dim lData           As Long
    Dim MByte()         As Byte
    Dim Btemp()         As String
    Dim lReturn         As Long

    On Error GoTo Handler

    Select Case DataType
        Case 1
            '/* sz
            Write_String RootKey, SubKey, KeyVal, ValData
        
        Case 2
            '/* expand_sz
            Write_Expanded RootKey, SubKey, KeyVal, ValData
            
        Case 3
            '/* multi_sz
            Write_Multi RootKey, SubKey, KeyVal, ValData
            
        Case 4
            '/* binary
            Btemp() = Split(ValData, " ")
            For j = 0 To UBound(Btemp) - 1
                ReDim Preserve MByte(j)
                MByte(j) = CByte(Btemp(j))
            Next j
            Write_Binary RootKey, SubKey, KeyVal, MByte
        
        Case 5
            '/* dword
            lData = CLng(ValData)
            Write_DWord RootKey, SubKey, KeyVal, lData
    
        Case 6
            '/* little_endian
            Write_LEndian RootKey, SubKey, KeyVal, lData
            
        Case 7
            '/* big_endian
            Write_BEndian RootKey, SubKey, KeyVal, lData
    End Select
    
    Erase Btemp
    Erase MByte

Handler:
    On Error GoTo 0

End Sub

Public Function Trim_Null(Item As String) As String

    '/* trim nulls

    Dim pos             As Integer

    On Error GoTo Handler
    '/* trim nulls for sz_multi
    pos = InStr(Item, vbNullChar)
    If pos Then
        Item = Left$(Item, pos - 1)
    End If
    Trim_Null = Item

Handler:
    On Error GoTo 0

End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           CONVERSION ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'/*  Thanks go out to actorics.de PSC submission

Private Function Low_Byte(ByVal iNum As Integer) As Byte

    On Error GoTo Handler

    Low_Byte = iNum And &HFF

Handler:
    On Error GoTo 0

End Function

Private Function High_Byte(ByVal iNum As Integer) As Byte

    On Error GoTo Handler

    High_Byte = (iNum And &HFF00&) \ 256

Handler:
    On Error GoTo 0

End Function

Private Function Make_Word(ByVal bLow As Byte, ByVal bHigh As Byte) As Integer

    On Error GoTo Handler

    If bHigh And &H80 Then
        Make_Word = ((bHigh * 256&) + bLow) Or &HFFFF0000
    Else
        Make_Word = (bHigh * 256) + bLow
    End If

Handler:
    On Error GoTo 0

End Function

Public Function Make_LEndian16(ByVal iVal As Integer) As Integer
    '/* create a 16bit little_endian
    On Error GoTo Handler

    Dim B(1)            As Byte

    B(0) = Low_Byte(iVal)
    B(1) = High_Byte(iVal)
    Make_LEndian16 = Make_Word(B(1), B(0))

Handler:
    On Error GoTo 0

End Function

Public Function Make_LEndian32(lVal As Long) As String
    '/* create 32bit little_endian

    Dim b0              As Byte
    Dim B1              As Byte
    Dim B2              As Byte
    Dim b3              As Byte

    On Error GoTo Handler

    b0 = (lVal And &HFF&)
    B1 = (lVal And &HFF00&) \ &H100&
    B2 = (lVal And &HFF0000) \ &H10000
    b3 = (lVal And &H7F000000) \ &H1000000 - 128 * (lVal < 0)
    Make_LEndian32 = Chr(b0) & Chr(B1) & Chr(B2) & Chr(b3)

Handler:
    On Error GoTo 0

End Function

Public Function Make_BEndian32(lVal As Long) As String
    '/* create 32bit big_endian

    Dim b0              As Byte
    Dim B1              As Byte
    Dim B2              As Byte
    Dim b3              As Byte

    On Error GoTo Handler

    b0 = (lVal And &HFF&)
    B1 = (lVal And &HFF00&) \ &H100&
    B2 = (lVal And &HFF0000) \ &H10000
    b3 = (lVal And &H7F000000) \ &H1000000 - 128 * (lVal < 0)
    Make_BEndian32 = Chr(b3) & Chr(B2) & Chr(B1) & Chr(b0)

Handler:
    On Error GoTo 0

End Function

Public Function Convert_Unicode(ByVal sWord As String) As Variant
    '/* convert ascii to unicode

    On Error GoTo Handler

    Convert_Unicode = StrConv(sWord, vbUnicode)
    
Handler:
    On Error GoTo 0

End Function

Public Function Convert_Byte(ByVal sVal As String) As Byte
    '/* convert string to byte array

    Dim bVal()          As Byte
    Dim i               As Long

    On Error GoTo Handler

    ReDim bVal(0 To Len(sVal))
    For i = 0 To UBound(bVal())
        bVal(i) = Mid$(sVal, i, 1)
    Next i
    
Handler:
    On Error GoTo 0

End Function

Public Function Convert_Curr(ByVal sValue As String) As Currency

    Dim l               As Long
    Dim Negative        As Boolean

    sValue = Trim$(sValue)
    If Left$(sValue, 1) = "-" Then
        Negative = True
        sValue = Mid$(sValue, 2)
    End If
    
    l = Len(sValue)
    If l < 4 Then
        Convert_Curr = CCur(IIf(Negative, "-0.", "0.") & Right$("0000" & sValue, 4))
    Else
        Convert_Curr = CCur(IIf(Negative, "-", "") & Left$(sValue, l - 4) & "." & Right$(sValue, 4))
    End If
    
End Function

Public Function Convert_Text(ByVal cValue As Currency) As String

    Dim Temp            As String
    Dim l               As Long

    Temp = Format$(cValue, "#.0000")
    l = Len(Temp)
    Temp = Left$(Temp, l - 5) & Right$(Temp, 4)
    
    Do While Len(Temp) > 1 And Left$(Temp, 1) = "0"
        Temp = Mid$(Temp, 2)
    Loop
    
    Do While Len(Temp) > 2 And Left$(Temp, 2) = "-0"
        Temp = "-" & Mid$(Temp, 3)
    Loop
    
    Temp = Temp / 10000
    Convert_Text = Temp
        
End Function

Public Function Get_KeyName(RootKey As HKEY_Type) As String
    Select Case RootKey
        Case HKEY_Type.HKEY_CLASSES_ROOT
            Get_KeyName = "HKEY_CLASSES_ROOT"
        Case HKEY_Type.HKEY_CURRENT_CONFIG
            Get_KeyName = "HKEY_CURRENT_CONFIG"
        Case HKEY_Type.HKEY_CURRENT_USER
            Get_KeyName = "HKEY_CURRENT_USER"
        Case HKEY_Type.HKEY_DYN_DATA
            Get_KeyName = "HKEY_DYN_DATA"
        Case HKEY_Type.HKEY_LOCAL_MACHINE
            Get_KeyName = "HKEY_LOCAL_MACHINE"
        Case HKEY_Type.HKEY_PERFORMANCE_DATA
            Get_KeyName = "HKEY_PERFORMANCE_DATA"
        Case HKEY_Type.HKEY_USERS
            Get_KeyName = "HKEY_USERS"
        Case Else
            Get_KeyName = ""
    End Select
End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           ERROR LOGGING
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Private Sub Error_State(ByVal lErrNum As Long, _
       ByVal sRoutine As String, _
       ByVal sKey As String)

    '/* send the error to log,
    '/* and raise notification

    Dim sErrDesc        As String

    sErrDesc = Get_Error(lErrNum)
    If m_bIntercept Then
        Log_Error "Description: " + sErrDesc + " Routine: " + " Location: " + sKey
    End If
    If m_bNotify Then
        RaiseEvent ErrorCond(sRoutine, sKey, sErrDesc)
    End If

End Sub

Private Function Get_Error(ByVal lErrNum As Long) As String
    '/* interpret registry errors

    On Error GoTo Handler

    Select Case lErrNum
        Case 1
            Get_Error = "Error: " & lErrNum & " Bad Database. The database is invalid"
        Case 2
            Get_Error = "Error: " & lErrNum & " Bad Key Name. Key is corrupt or does not exist"
        Case 3
            Get_Error = "Error: " & lErrNum & " Key Locked. The key requested does not exist"
        Case 4
            Get_Error = "Error: " & lErrNum & " Failed Read. The key requested can not be read accessed"
        Case 5
            Get_Error = "Error: " & lErrNum & " Failed Write. The key requested can not be write accessed"
        Case 6
            Get_Error = "Error: " & lErrNum & " Out Of Memory. The process has exceeded its memory allocation"
        Case 7
            Get_Error = "Error: " & lErrNum & " Work Area Invalid. The process work area appears to be corrupt"
        Case 8
            Get_Error = "Error: " & lErrNum & " Access Is Denied. Access to this key has been denied"
        Case 87
            Get_Error = "Error: " & lErrNum & " Invalid Parameters. The call parameters passed are invalid"
        Case 234
            Get_Error = "Error: " & lErrNum & " More Data. The container has more data"
        Case 259
            Get_Error = "Error: " & lErrNum & " No More Items. No more items for this query"
        Case Else
            Get_Error = "Error: An Unknown Error has occured"
    End Select

Handler:
    On Error GoTo 0

End Function

Private Sub Log_Error(ByVal ErrDesc As String)
    '/* log errors

    On Error Resume Next

    '/* write to log
    Open App.Path & "\err.log" For Append As #1
    Print #1, CStr(Now) & vbTab & ErrDesc
    Close #1

    On Error GoTo 0

End Sub
