VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Time Converter"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Set TargetZone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6150
      TabIndex        =   11
      Top             =   2700
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set SourceZone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6180
      TabIndex        =   10
      Top             =   990
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1140
      TabIndex        =   7
      Top             =   4110
      Width           =   1635
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   1170
      TabIndex        =   3
      Top             =   2490
      Width           =   4875
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   4875
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2745
   End
   Begin VB.Label Label4 
      Caption         =   "Set selected Timezone as default target."
      Height          =   615
      Left            =   6150
      TabIndex        =   9
      Top             =   3030
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Set selected Timezone as default source."
      Height          =   615
      Left            =   6180
      TabIndex        =   8
      Top             =   1350
      Width           =   1935
   End
   Begin VB.Label sResult 
      BackColor       =   &H80000018&
      Caption         =   "Result will be displayed here. "
      Height          =   465
      Left            =   3030
      TabIndex        =   6
      Top             =   4110
      Width           =   2985
   End
   Begin VB.Label Label2 
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   5
      Top             =   2790
      Width           =   675
   End
   Begin VB.Label lblFrom 
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Time to convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Operating System version information declares
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const VER_PLATFORM_WIN32_WINDOWS = 1

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'Time Zone information declares

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type REGTIMEZONEINFORMATION
        Bias As Long
        StandardBias As Long
        DaylightBias As Long
        StandardDate As SYSTEMTIME
        DaylightDate As SYSTEMTIME
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function SetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'Registry information declares
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_ARENA_TRASHED = 7&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_CREATED_NEW_KEY = &H1

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long
    
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'Registry editing constants, types and APIs
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
    lpdwDisposition As Long) As Long
    
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long
    ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'Registry Constants
Const SKEY_NT = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\TIME ZONES"
Const SKEY_9X = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\TIME ZONES"

Const SCurrTZKEY = "SYSTEM\CurrentControlSet\Control\TimeZoneInformation"

'The following declaration is different from the one in the API viewer.
'To disable implicit ANSI<->UNICODE conversion, it changes the variable
'   types of lpMultiByteStr and plWideCharStr to Any.
'Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
'Private Const CP_ACP = 0
'Private Const MB_PRECOMPOSED = &H1

Dim SubKey As String

Dim tzNames() As String, tzDisplayNames() As String
Dim tzSource As TIME_ZONE_INFORMATION, tzTarget As TIME_ZONE_INFORMATION
Dim sSourceZone As String, sTargetZone As String

Private Sub Command2_Click()
    Dim lRetVal As Long, hKeyResult As Long
    Dim sSecurity As SECURITY_ATTRIBUTES
    
    lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\smallApps\TimeConverter", _
        0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sSecurity, hKeyResult, _
        lRetVal1)

    If lRetVal = ERROR_SUCCESS Then
        lRetVal = RegSetValueEx(hKeyResult, "DefaultSourceZone", 0, REG_SZ, ByVal tzNames(List1.ListIndex), Len(tzNames(List1.ListIndex)))
        If lRetVal <> ERROR_SUCCESS Then
            MsgBox "Couldn't set value as Default." & vbCrLf & "Problem writing to registry.", , "Err :: Time Calculator"
        Else
            RegFlushKey (hKeyResult)
            RegCloseKey (hKeyResult)
            MsgBox "SourceZone set to """ & List1.Text & """." & vbCrLf & "Change will apply next time you open TimeConverter.", , "Alert :: TimeConverter - smallApps"
        End If
    End If
    
End Sub

Private Sub Command3_Click()
    Dim lRetVal As Long, hKeyResult As Long
    Dim sSecurity As SECURITY_ATTRIBUTES
    
    lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\smallApps\TimeConverter", _
        0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sSecurity, hKeyResult, _
        lRetVal1)

    If lRetVal = ERROR_SUCCESS Then
        lRetVal = RegSetValueEx(hKeyResult, "DefaultTargetZone", 0, REG_SZ, ByVal tzNames(List2.ListIndex), Len(tzNames(List2.ListIndex)))
        If lRetVal <> ERROR_SUCCESS Then
            MsgBox "Couldn't set value as Default." & vbCrLf & "Problem writing to registry.", , "Err :: Time Calculator"
        Else
            RegFlushKey (hKeyResult)
            RegCloseKey (hKeyResult)
            MsgBox "TargetZone set to """ & List2.Text & """." & vbCrLf & "Change will apply next time you open TimeConverter.", , "Alert :: TimeConverter - smallApps"
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim lRetVal As Long, hKeyResult As Long
    Dim iMinsBiasSource As Long, iMinsBiasTarget As Long
    Dim sysTimeUTC As SYSTEMTIME

    'GetSystemTime sysTimeUTC
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey & "\" & tzNames(List1.ListIndex), 0&, KEY_ALL_ACCESS, hKeyResult)
    If lRetVal = ERROR_SUCCESS Then
        lRetVal = RegQueryValueEx(hKeyResult, "TZI", 0, ByVal 0&, tzSource, Len(tzSource))
        If lRetVal = ERROR_SUCCESS Then
            iMinsBiasSource = tzSource.Bias
        End If
    End If
    RegCloseKey (hKeyResult)
                                                  
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey & "\" & tzNames(List2.ListIndex), 0&, KEY_ALL_ACCESS, hKeyResult)
    If lRetVal = ERROR_SUCCESS Then
        lRetVal = RegQueryValueEx(hKeyResult, "TZI", 0, ByVal 0&, tzSource, Len(tzSource))
        If lRetVal = ERROR_SUCCESS Then
            iMinsBiasTarget = tzSource.Bias
        End If
    End If
    RegCloseKey (hKeyResult)
    
    Dim tActualUTC As Date
    
    tActualUTC = DateAdd("n", iMinsBiasSource, CDate(Text1.Text))
    sResult = CStr(DateAdd("n", -(iMinsBiasTarget), tActualUTC))
    sResult.FontName = "Verdana"
    sResult.FontBold = True
    sResult.FontSize = 10
    
End Sub

Private Sub Form_Load()
    Dim lRetVal As Long, lResult As Long, lCurIdx As Long
    Dim lDataLen As Long, lValueLen As Long, hKeyResult As Long
    Dim strValue As String, strValue1 As String, hKeyResult1 As Long
    Dim lValueLen1 As Long
    
    Dim osV As OSVERSIONINFO
    
    'Win9X and WinNT have a slightly different registry structure.
    'Determine the operating system and set a module variable to the
    'correct subkey.
    
    Text1.Text = Now
    
    osV.dwOSVersionInfoSize = Len(osV)
    Call GetVersionEx(osV)
    If osV.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        SubKey = SKEY_NT
    Else
        SubKey = SKEY_9X
    End If

    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey, 0, _
                KEY_ALL_ACCESS, hKeyResult)
    
    If lRetVal = ERROR_SUCCESS Then
        lCurIdx = 0
        lValueLen = 32
        
        Do
            strValue = String(lValueLen, 0)
            
            lResult = RegEnumKey(hKeyResult, lCurIdx, strValue, lValueLen)
            If lResult = ERROR_SUCCESS Then
                lRetVal = RegOpenKeyEx(hKeyResult, trimNull(strValue), 0, _
                    KEY_ALL_ACCESS, hKeyResult1)

                If lRetVal = ERROR_SUCCESS Then
                
                    lRetVal = RegQueryValueEx(hKeyResult1, "Display", 0, _
                         0, ByVal 0, lValueLen1)
                
                    If lRetVal = ERROR_SUCCESS Then
                        strValue1 = String(lValueLen1, 0)
                        lRetVal = RegQueryValueEx(hKeyResult1, "Display", 0, _
                             0, ByVal strValue1, Len(strValue1))
                        If lRetVal = ERROR_SUCCESS Then
                            ReDim Preserve tzNames(lCurIdx)
                            ReDim Preserve tzDisplayNames(lCurIdx)
                            tzNames(lCurIdx) = trimNull(strValue)
                            tzDisplayNames(lCurIdx) = trimNull(strValue1)
                        End If
                    End If
                    RegCloseKey (hKeyResult1)
                End If
            End If

            lCurIdx = lCurIdx + 1
        Loop While lResult = ERROR_SUCCESS
        RegCloseKey hKeyResult

        'Sort array values.
        Dim sTempValue, iLoop1, iLoop2
        For iLoop1 = 0 To UBound(tzDisplayNames) - 1
            For iLoop2 = 1 To UBound(tzDisplayNames)
                If tzDisplayNames(iLoop2 - 1) < tzDisplayNames(iLoop2) Then
                    sTempValue = tzDisplayNames(iLoop2 - 1)
                    tzDisplayNames(iLoop2 - 1) = tzDisplayNames(iLoop2)
                    tzDisplayNames(iLoop2) = sTempValue
                
                    sTempValue = tzNames(iLoop2 - 1)
                    tzNames(iLoop2 - 1) = tzNames(iLoop2)
                    tzNames(iLoop2) = sTempValue
                End If
            Next
        Next

        'populate list with array values
        For lCurIdx = 0 To UBound(tzNames)
            List1.AddItem tzDisplayNames(lCurIdx)
            List2.AddItem tzDisplayNames(lCurIdx)
        Next
        
        ReadRegistrySettings
        
    Else
        List1.AddItem "Could not open registry for timezones"
        List2.AddItem "Could not open registry for timezones"
    End If
End Sub

Private Sub ReadRegistrySettings()

    'create key and subkey in registry, if not there.
    Dim lRetVal As Long, hKeyResult As Long, lRetVal1 As Long
    Dim hKeyResult1 As Long, sSecurity As SECURITY_ATTRIBUTES
    Dim lValueLen As Long, lValueType As Long

    'CREATE OR OPEN MAIN APPLICATION KEY
    lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\smallApps", _
        0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sSecurity, hKeyResult, _
        lRetVal1)
    
    If lRetVal = ERROR_SUCCESS Then
        
        'IF NEW KEY CREATED PUT DEFAULT VALUE
        If lRetVal1 = REG_CREATED_NEW_KEY Then
            lRetVal1 = RegSetValueEx(hKeyResult, "", 0, REG_SZ, ByVal "Harish's Applications", Len("Harish's Applications"))
            If lRetVal1 = ERROR_SUCCESS Then
                RegFlushKey (hKeyResult)
            End If
        End If
        
        'CREATE OR OPEN SUBKEY TIMECONVERTER
        lRetVal = RegCreateKeyEx(hKeyResult, "TimeConverter", _
            0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sSecurity, _
            hKeyResult1, lRetVal1)

        If lRetVal = ERROR_SUCCESS Then
            
            'DETERMINING DEFAULT SOURCE FROM REGISTRY.
            'if source not found, default to local timezone.
            lRetVal = RegQueryValueEx(hKeyResult1, "DefaultSourceZone", 0, lValueType, ByVal 0, lValueLen)
            
            If lRetVal = ERROR_SUCCESS Then
                sSourceZone = String(lValueLen, Chr$(0))
                lRetVal = RegQueryValueEx(hKeyResult1, "DefaultSourceZone", 0, _
                        0, ByVal sSourceZone, Len(sSourceZone))
            End If
            
            If lRetVal <> ERROR_SUCCESS Then
                'value not found
                'set default source to current timezone
                Dim tzCurr As TIME_ZONE_INFORMATION
                Dim sStrName As String, iIndex1 As Integer
                
                lRetVal = GetTimeZoneInformation(tzCurr)
                If lRetVal <> TIME_ZONE_ID_INVALID Then
                    sSourceZone = trimNull(CStr(tzCurr.StandardName))
                End If
            Else
                sSourceZone = trimNull(sSourceZone)
            End If
            
            'DETERMINING DEFAULT TARGET FROM REGISTRY.
            'if the default target value is not found, no value is pre-selected.
            'no default target
            lRetVal = RegQueryValueEx(hKeyResult1, "DefaultTargetZone", 0, lValueType, ByVal 0, lValueLen)
            
            If lRetVal = ERROR_SUCCESS Then
                sTargetZone = String(lValueLen, Chr$(0))
                lRetVal = RegQueryValueEx(hKeyResult1, "DefaultTargetZone", 0, _
                        0, ByVal sTargetZone, Len(sTargetZone))
                        
                If lRetVal = ERROR_SUCCESS Then
                    sTargetZone = trimNull(sTargetZone)
                End If
                
            End If
            
            RegCloseKey (hKeyResult)
            RegCloseKey (hKeyResult1)
        End If
    End If

    'set preselected values for listboxes
        For iIndex1 = 0 To List1.ListCount - 1
            If tzNames(iIndex1) = sSourceZone Then
                List1.Selected(iIndex1) = True
                Exit For
            End If
        Next
        For iIndex1 = 0 To List2.ListCount - 1
            If tzNames(iIndex1) = sTargetZone Then
                List2.Selected(iIndex1) = True
                Exit For
            End If
        Next
End Sub

Private Function trimNull(strIn As String) As String
    Dim nNull
    nNull = InStr(strIn, vbNullChar)
    Select Case nNull
        Case Is > 1
            trimNull = Left(strIn, nNull - 1)
        Case 1
            trimNull = ""
        Case 0
            trimNull = Trim(strIn)
    End Select
End Function
