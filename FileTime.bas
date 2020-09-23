Attribute VB_Name = "modFileTime"
Option Explicit
'                                                                 ©Rd
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                          FILE FUNCTIONS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
   (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpSpec As String) As Long

Private Declare Function SetAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpSpec As String, ByVal dwAttributes As Long) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OF_READ = &H0&
Private Const FILE_SHARE_READ = &H1&
Private Const FILE_SHARE_WRITE = &H2&
Private Const OPEN_EXISTING = &H3&
Private Const FILE_ATTRIBUTE_NORMAL = &H80&
Private Const INVALID_HANDLE_VALUE = &HFFFFFFFF

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                  FILETIME AND SYSTEMTIME FUNCTIONS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The LocalFileTimeToFileTime function converts a local file time
' to a file time based on the Universal Time Convention (UTC).

Private Declare Function LocalFileTimeToFileTime Lib "kernel32" _
    (lpLocalFileTime As Currency, lpFileTime As Currency) As Long

' The FileTimeToLocalFileTime function converts a file time based
' on a UTC-based time to a local file time.

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" _
    (lpFileTime As Currency, lpLocalFileTime As Currency) As Long

' Functions involving local time use the current settings for the time
' zone and daylight saving time. Therefore, if it is daylight saving time,
' the function will take daylight saving time into account, even if the
' time you are converting is in standard time.

'Public Type FILETIME
'    dwLowDateTime As Long  ' The low 32 bits of the Win32 date/time value.
'    dwHighDateTime As Long ' The upper 32 bits of the Win32 date/time value.
'End Type

Public Type FILETIME
    ftDateTime As Currency
End Type

' The FILETIME structure holds an unsigned 64-bit date and time value for
' a file. This value represents the number of 100-nanosecond units since
' the beginning of January 1, 1601. The FILETIME data structure is used
' in the time conversion functions between DOS and Win32.

' It is not recommended that you add and subtract values from the FILETIME
' structure to obtain relative times. Instead, you should copy the FILETIME
' structure to a LARGE_INTEGER structure and use normal 64-bit arithmetic
' on the LARGE_INTEGER value.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' It is not recommended that you add and subtract values from the SYSTEMTIME
' structure to obtain relative times. Instead, you should convert the SYSTEMTIME
' structure to a FILETIME structure, then copy the resulting FILETIME structure
' to a LARGE_INTEGER structure and use normal 64-bit arithmetic on the
' LARGE_INTEGER value.

' The SystemTimeToFileTime function converts a system time to a file time.
' The wDayOfWeek member of the SYSTEMTIME structure may be ignored?

Private Declare Function SystemTimeToFileTime Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME, lpFileTime As Currency) As Long

' The SYSTEMTIME structure represents a date and time using individual members
' for the month, day, year, weekday, hour, minute, second, and millisecond.

Public Type SYSTEMTIME
    wYear As Integer         ' Specifies the current year.
    wMonth As Integer        ' Specifies the current month; January = 1, February = 2, and so on.
    wDayOfWeek As Integer    ' Specifies the current day of the week; Sunday = 0, Monday = 1, and so on.
    wDay As Integer          ' Specifies the current day of the month.
    wHour As Integer         ' Specifies the current hour.
    wMinute As Integer       ' Specifies the current minute.
    wSecond As Integer       ' Specifies the current second.
    wMilliseconds As Integer ' Specifies the current millisecond.
End Type

' The FileTimeToSystemTime function converts a 64-bit file time to
' system time format. This function fails for FILETIME values that
' are greater than 0x7FFFFFFFFFFFFFF.

Private Declare Function FileTimeToSystemTime Lib "kernel32" _
    (lpFileTime As Currency, lpSystemTime As SYSTEMTIME) As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The VariantTimeToSystemTime function converts the variant Date
' representation of time to system time values. See the comments below.

Private Declare Function VariantTimeToSystemTime Lib "oleaut32" _
    (ByVal vtime As Date, lpSystemTime As SYSTEMTIME) As Long

' The SystemTimeToVariantTime function converts system time
' values to the variant Date format.

Private Declare Function SystemTimeToVariantTime Lib "oleaut32" _
    (lpSystemTime As SYSTEMTIME, vtime As Date) As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' System Time

' All translations between UTC time and local time are based on the
' following formula:

' UTC = local time + bias
 
' The bias is the difference, in minutes, between UTC time and local
' time. A variant time is stored as an 8-byte real value (double),
' representing a date between January 1, 1753 and December 31, 2078,
' inclusive.

' The value 2.0 represents January 1, 1900; 3.0 represents January 2,
' 1900, and so on. Adding 1 to the value increments the date by a day.
' The fractional part of the value represents the time of day.

' Therefore, 2.5 represents noon on January 1, 1900; 3.25 represents
' 6:00 a.m. on January 2, 1900, and so on. Negative numbers represent
' the dates prior to December 30, 1899.

' Using the SYSTEMTIME structure is useful because:

' It spans all time/date periods. MS-DOS date/time is limited to
' representing only those dates between 1/1/1980 and 12/31/2107.

' The date/time elements are all easily accessible without needing to
' do any bit decoding.

' The National Language Support data and time formatting functions
' GetDateFormat and GetTimeFormat take a SYSTEMTIME value as input.

' SYSTEMTIME is the default Win32 time and date data format supported
' by Windows NT and Windows 95.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The GetSystemTimeAsFileTime function obtains the current system date
' and time. The information is in Coordinated Universal Time (UTC) format.

' This function is equivalent to using GetSystemTime and passing the
' result to SystemTimeToFileTime. This sub does not return a value.

Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32" _
    (lpFileTime As Currency)

' The GetSystemTime function retrieves the current system date and time.
' The system time is expressed in Coordinated Universal Time (UTC).

Private Declare Sub GetSystemUTCTime Lib "kernel32" Alias "GetSystemTime" _
    (lpSystemTime As SYSTEMTIME)

' The SetSystemTime function sets the current system time and date.
' The system time is expressed in Coordinated Universal Time (UTC).

Private Declare Function SetSystemUTCTime Lib "kernel32" Alias "SetSystemTime" _
    (lpSystemTime As SYSTEMTIME) As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Local Time

' While the system uses UTC-based time internally, your applications
' will generally display the local time — the date and time of day for
' your time zone. Therefore, to ensure correct results, you must be
' aware of whether a function expects to receive a UTC-based time or
' a local time, and whether the function returns a UTC-based time or
' a local time.

' You can retrieve the local time by using the GetLocalTime function.
' The GetLocalTime function retrieves the current local date and time,
' assigning to a SYSTEMTIME structure.

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

' GetLocalTime converts the system time to a local time based on the
' current time-zone settings and copies the result to a SYSTEMTIME
' structure. You can set the system time by using the SetLocalTime
' function. SetLocalTime assumes you have specified a local time and
' converts to UTC before setting the system time.

' The SetLocalTime function sets the current local time and date.

Private Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The GetFileTime function retrieves the date and time that a file was
' created, last accessed, and last modified. The file handle must have
' been created/opened with GENERIC_READ access to the file.

Private Declare Function GetFileTime Lib "kernel32" _
    (ByVal hFile As Long, lpCreationTime As Currency, _
     lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long

' Each lp FileTime parameter can be NULL (0&) if the application does
' not need to get this information.

' The SetFileTime function sets the date and time that a file was
' created, last accessed, or last modified. The file handle must have
' been created/opened with GENERIC_WRITE access to the file.

Private Declare Function SetFileTime Lib "kernel32" _
    (ByVal hFile As Long, lpCreationTime As Currency, _
     lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long

' Each lp FileTime parameter can be NULL (0&) if the application does
' not need to set this information.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The FileTimeToDosDateTime function converts a 64-bit file time
' to MS-DOS date and time values.

Private Declare Function FileTimeToDosDateTime Lib "kernel32" _
    (lpFileTime As Currency, ByVal lpFatDate As Long, _
     ByVal lpFatTime As Long) As Long

' The DosDateTimeToFileTime function converts MS-DOS date and time
' values to a 64-bit file time.

Private Declare Function DosDateTimeToFileTime Lib "kernel32" _
    (ByVal wFatDate As Long, ByVal wFatTime As Long, _
     lpFileTime As Currency) As Long

' MS-DOS records file dates and times as packed 16-bit values.

' MS-DOS date and time
Public Type DOSDATETIME
    iFileDate As Integer  ' 16-bit MS-DOS date
    iFileTime As Integer  ' 16-bit MS-DOS time
End Type

' An MS-DOS date has the following format:
'   Bits    Contents
'   0–4     Day of the month (1–31).
'   5–8     Month (1 = January, 2 = February, and so on).
'   9–15    Year offset from 1980 (add 1980 to get the actual year).

' An MS-DOS time has the following format:
'   Bits    Contents
'   0–4     Second divided by 2.
'   5–10    Minute (0–59).
'   11–15   Hour (0– 23 on a 24-hour clock).

' IMPORTANT: The MS-DOS date format can represent only dates
' between 1/1/1980 and 12/31/2107; this conversion fails if the
' input file time is outside this range.

' Also, the MS-DOS date format is accurate to two-second intervals
' which has far less precision than provided by the Date format.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The VariantTimeToDosDateTime function converts the variant Date
' representation of a date and time to MS-DOS date and time values.

Private Declare Function VariantTimeToDosDateTime Lib "oleaut32" _
    (ByVal vtime As Date, lpFatDate As Long, lpFatTime As Long) As Long

' The DosDateTimeToVariantTime function converts the MS-DOS representation
' of time to the date and time representation stored in a variant Date.

Private Declare Function DosDateTimeToVariantTime Lib "oleaut32" _
    (ByVal lpFatDate As Long, ByVal lpFatTime As Long, vtime As Date) As Long

' A variant time is stored as an 8-byte real value representing a
' date between January 1, 1753 and December 31, 2078, inclusive.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Return Values

' If the above functions succeed, the return value is nonzero.
' If the above functions fail, the return value is zero. To get
' extended error information, call Err.LastDllError.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                     COMPARE FILETIME FUNCTION
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The CompareFileTime function compares two 64-bit file times.

Public Declare Function CompareFileTime Lib "kernel32" _
    (lpFileTime1 As Currency, lpFileTime2 As Currency) As eCompareFileTime

' The return value is one of the following values:

Public Enum eCompareFileTime
    Older = -1&  ' First file time is less than second file time
    Equal = 0&   ' First file time is equal to second file time
    Newer = 1&   ' First file time is greater than second file time
End Enum

#If False Then
    Dim Older, Equal, Newer
#End If

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                  FILETIME CONVERSION FUNCTIONS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Note - Do not use this function for dates before 1/1/1980
Public Function FileTimeToDosTime(ByVal cFileTime As Currency) As DOSDATETIME
    On Error GoTo HandleIt
    Dim lDate As Long, lTime As Long
    ' Convert the file time format to dos format
    FileTimeToDosDateTime cFileTime, lDate, lTime
    FileTimeToDosTime.iFileDate = lDate
    FileTimeToDosTime.iFileTime = lTime
HandleIt:
End Function

' Note - Do not use this function for dates before 1/1/1980
Public Function DosTimeToFileTime(tDos As DOSDATETIME) As Currency
    On Error GoTo HandleIt
    Dim lDate As Long, lTime As Long
    lDate = tDos.iFileDate
    lTime = tDos.iFileTime
    ' Convert the dos format to file time format
    DosDateTimeToFileTime lDate, lTime, DosTimeToFileTime
HandleIt:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Note - Do not use this function for dates before 1/1/1980
Public Function SysTimeToDosTime(tSystemTime As SYSTEMTIME, Optional ByVal bIsLocalTime As Boolean = True) As DOSDATETIME
    On Error GoTo HandleIt
    Dim cFileTime As Currency
    Dim lDate As Long, lTime As Long
    ' Convert the system time to file time format
    SystemTimeToFileTime tSystemTime, cFileTime
    If bIsLocalTime Then
       ' Convert the local time to file time
       LocalFileTimeToFileTime cFileTime, cFileTime
    End If
    ' Convert the file time to dos format
    FileTimeToDosDateTime cFileTime, lDate, lTime
    SysTimeToDosTime.iFileDate = lDate
    SysTimeToDosTime.iFileTime = lTime
HandleIt:
End Function

' Note - Do not use this function for dates before 1/1/1980
Public Function DosTimeToSysTime(tDos As DOSDATETIME, Optional ByVal bAsLocalTime As Boolean = True) As SYSTEMTIME
    On Error GoTo HandleIt
    Dim cFileTime As Currency
    Dim lDate As Long, lTime As Long
    lDate = tDos.iFileDate: lTime = tDos.iFileTime
    ' Convert the dos time to file time format
    DosDateTimeToFileTime lDate, lTime, cFileTime
    If bAsLocalTime Then
       ' Convert the file time to local time zone
       FileTimeToLocalFileTime cFileTime, cFileTime
    End If
    ' Convert the file time to system time
    FileTimeToSystemTime cFileTime, DosTimeToSysTime
HandleIt:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function SysTimeToFileTime(tSystemTime As SYSTEMTIME, Optional ByVal bIsLocalTime As Boolean = True) As Currency
    On Error GoTo HandleIt
    ' Convert the system time to file time format
    SystemTimeToFileTime tSystemTime, SysTimeToFileTime
    If bIsLocalTime Then
       ' Convert the local time zone to file time
       LocalFileTimeToFileTime SysTimeToFileTime, SysTimeToFileTime
    End If
HandleIt:
End Function

Public Function FileTimeToSysTime(ByVal cFileTime As Currency, Optional ByVal bAsLocalTime As Boolean = True) As SYSTEMTIME
    On Error GoTo HandleIt
    If bAsLocalTime Then
       ' Convert the file time to the local time zone
       FileTimeToLocalFileTime cFileTime, cFileTime
    End If
    ' Convert the file time to system time format
    FileTimeToSystemTime cFileTime, FileTimeToSysTime
HandleIt:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function CurrentTimeAsFileTime() As Currency
    On Error GoTo HandleIt
    ' Get the system time in the file time format
    GetSystemTimeAsFileTime CurrentTimeAsFileTime
    ' The GetSystemTimeAsFileTime function is equivalent to using
    ' GetSystemTime and passing the result to SystemTimeToFileTime
HandleIt:
End Function

Public Function CurrentTimeAsSysTime(Optional ByVal bAsLocalTime As Boolean = True) As SYSTEMTIME
    On Error GoTo HandleIt
    ' Get the current system time
    If bAsLocalTime Then
       ' Return the current time as local time zone
       GetLocalTime CurrentTimeAsSysTime
    Else
       ' Return the current time as UTC time format
       GetSystemUTCTime CurrentTimeAsSysTime
    End If
HandleIt:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub SetSystemTime(tSystemTime As SYSTEMTIME, Optional ByVal bIsLocalTime As Boolean = True)
    On Error GoTo HandleIt
    ' Set the system time to the specified time
    If bIsLocalTime Then
       ' Pass the time as local time zone
       SetLocalTime tSystemTime
    Else
       ' Pass the time as UTC time format
       SetSystemUTCTime tSystemTime
    End If
HandleIt:
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function DateToFileTime(ByVal dtDate As Date, Optional ByVal bIsLocalDate As Boolean = True) As Currency
    On Error GoTo HandleIt
    Dim tSysTime As SYSTEMTIME
    ' Convert the variant Date to system time format
    VariantTimeToSystemTime dtDate, tSysTime
    ' Convert the system time to file time format
    SystemTimeToFileTime tSysTime, DateToFileTime
    If bIsLocalDate Then
       ' Convert the local time to file time
       LocalFileTimeToFileTime DateToFileTime, DateToFileTime
    End If
HandleIt:
End Function

Public Function FileTimeToDate(ByVal cFileTime As Currency, Optional ByVal bAsLocalDate As Boolean = True) As Date
    On Error GoTo HandleIt
    Dim tSysTime As SYSTEMTIME
    If bAsLocalDate Then
       ' Convert file time to local time zone
       FileTimeToLocalFileTime cFileTime, cFileTime
    End If
    ' Convert the file time to system time format
    FileTimeToSystemTime cFileTime, tSysTime
    ' Convert the system time to variant Date format
    SystemTimeToVariantTime tSysTime, FileTimeToDate
HandleIt:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function DateToSystemTime(ByVal dtDate As Date) As SYSTEMTIME
    On Error GoTo HandleIt
    ' Convert the variant Date to system time format
    VariantTimeToSystemTime dtDate, DateToSystemTime
HandleIt:
End Function

Public Function SystemTimeToDate(tSystemTime As SYSTEMTIME) As Date
    On Error GoTo HandleIt
    ' Convert the system time format to a variant Date
    SystemTimeToVariantTime tSystemTime, SystemTimeToDate
HandleIt:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                      MODIFYING FILE TIMES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Property Get FileCreated(sFileSpec As String) As Currency
    On Error GoTo HandleIt
    If PathExists(sFileSpec) Then
        ' Gets the creation file time for the specified file
        Dim hFile As Long, cJunk1 As Currency, cJunk2 As Currency

        hFile = CreateFile(sFileSpec, GENERIC_READ, FILE_SHARE_READ, 0&, _
                           OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)

        If (hFile = INVALID_HANDLE_VALUE) Then Exit Property

        GetFileTime hFile, FileCreated, cJunk1, cJunk2
    End If
HandleIt:
    If (hFile <> 0) Then CloseHandle hFile
End Property

Public Property Let FileCreated(sFileSpec As String, ByVal cCreationTime As Currency)
    On Error GoTo HandleIt
    If PathExists(sFileSpec) Then
        ' Changess the creation file time for the specified file
        Dim hFile As Long, cJunk As Currency
        Dim cAccessTime As Currency, cLastWrite As Currency

        hFile = CreateFile(sFileSpec, GENERIC_READ Or GENERIC_WRITE, 0&, _
                           0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)

        If (hFile = INVALID_HANDLE_VALUE) Then Exit Property

        GetFileTime hFile, cJunk, cAccessTime, cLastWrite
        SetFileTime hFile, cCreationTime, cAccessTime, cLastWrite
    End If
HandleIt:
    If (hFile <> 0) Then CloseHandle hFile
End Property

Public Property Get FileLastModified(sFileSpec As String) As Currency
    On Error GoTo HandleIt
    If PathExists(sFileSpec) Then
        ' Gets the last write file time for the specified file
        Dim hFile As Long, cJunk1 As Currency, cJunk2 As Currency

        hFile = CreateFile(sFileSpec, GENERIC_READ, FILE_SHARE_READ, 0&, _
                           OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)

        If (hFile = INVALID_HANDLE_VALUE) Then Exit Property

        GetFileTime hFile, cJunk1, cJunk2, FileLastModified
    End If
HandleIt:
    If (hFile <> 0) Then CloseHandle hFile
End Property

Public Property Let FileLastModified(sFileSpec As String, ByVal cLastWriteTime As Currency)
    On Error GoTo HandleIt
    If PathExists(sFileSpec) Then
        ' Sets the last write file time for the specified file
        Dim hFile As Long, cJunk As Currency
        Dim cAccessTime As Currency, cCreateTime As Currency

        hFile = CreateFile(sFileSpec, GENERIC_READ Or GENERIC_WRITE, 0&, _
                           0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)

        If (hFile = INVALID_HANDLE_VALUE) Then Exit Property

        GetFileTime hFile, cCreateTime, cAccessTime, cJunk
        SetFileTime hFile, cCreateTime, cAccessTime, cLastWriteTime
    End If
HandleIt:
    If (hFile <> 0) Then CloseHandle hFile
End Property

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                         SUPPORT FUNCTION
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function PathExists(sSpec As String) As Boolean
    If (LenB(sSpec) <> 0) Then
        PathExists = (GetAttributes(sSpec) <> INVALID_HANDLE_VALUE)
    End If
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ :) ++
