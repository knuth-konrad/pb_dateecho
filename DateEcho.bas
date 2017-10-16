'------------------------------------------------------------------------------
'Purpose  : Echo the current date/time to STDOut
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 04.05.2017
'           - #Break on to prevent console context menu changes
'           15.05.2017
'           - Application manifest added
'------------------------------------------------------------------------------
#Compile Exe ".\DateEcho.exe"
#Option Version5
#Break On
#Dim All

#Debug Error Off
#Tools Off

DefLng A-Z

%VERSION_MAJOR = 2
%VERSION_MINOR = 0
%VERSION_REVISION = 0

' Version Resource information
#Include ".\DateEchoRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
%DATE_TIME_ERROR = -1&
%DATE_SHORT = 0&
%DATE_LONG = 1&
%TIME = 2&
%DATE_TIME_SHORT = 3&
%DATE_TIME_LONG = 4&
%DATE_TIME_CUSTOM = 5&
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include "win32api.inc"
#Include "sautilcc.inc"
'----------------------------------------------------------------------------

Function PBMain()
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 08.02.2017
'           - Code reformatting
'           16.10.2017
'           - New parameter: /m=<custom mask>
'------------------------------------------------------------------------------
   Local sMask, sResult As String
   Local lLocale As Long

   lLocale = GetUserDefaultLCID()

   ' Application intro
   Print ""
   ConHeadline "DateEcho", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2003-2017", $COMPANY_NAME
   Print ""

   Select Case As Long ParseCmd(Command$, sMask)
   Case %DATE_TIME_ERROR
      Call ShowHelp
   Case %DATE_SHORT
      Print ShortDate(lLocale)
      StdOut ShortDate(lLocale)
   Case %DATE_LONG
      Print LongDate(lLocale)
      StdOut LongDate(lLocale)
   Case %TIME
      Print Time$
      StdOut Time$
   Case %DATE_TIME_SHORT
      Print ShortDate(lLocale) & ", " & Time$
      StdOut ShortDate(lLocale) & ", " & Time$
   Case %DATE_TIME_LONG
      Print LongDate(lLocale) & ", " & Time$
      StdOut LongDate(lLocale) & ", " & Time$
   Case %DATE_TIME_CUSTOM
      sResult = CustomDate(sMask)
      Print sResult
      StdOut sResult
   End Select

End Function
'---------------------------------------------------------------------------

Function ParseCmd(ByVal sCmd As String, ByRef sMask As String) As Long
'------------------------------------------------------------------------------
'Purpose  : Parses the command line parameters
'
'Prereq.  : -
'Parameter: sCmd  - Parameters passed to the application
'           sMask - (ByRef!) - returns the mask passed with the
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 08.02.2017
'           - Code reformatting
'------------------------------------------------------------------------------
   Local i As Long
   ReDim asArg(1 To 2) As String

   If InStr(sCmd, Any "<>") Then
      sCmd = Left$(sCmd, InStr(sCmd, Any "<>") -1)
   End If

   Parse sCmd, asArg(), Any "=:"

   ReDim Preserve asArg(1 To ParseCount(sCmd, Any "=:"))

   If UBound(asArg) < 1 Then
      ParseCmd = %DATE_TIME_ERROR
      Exit Function
   End If

   Replace "/" With "" In asArg(1)

   Select Case As Const$ LCase$(asArg(1))
   Case "d"
      Select Case Trim$(LCase$(asArg(2)))
      Case "l"
         ParseCmd = %DATE_LONG
      Case "s"
         ParseCmd = %DATE_SHORT
      Case Else
         ParseCmd = %DATE_TIME_ERROR
      End Select
      Exit Function

   Case "dt"
      Select Case Trim$(LCase$(asArg(2)))
      Case "l"
         ParseCmd = %DATE_TIME_LONG
      Case "s"
         ParseCmd = %DATE_TIME_SHORT
      Case Else
         ParseCmd = %DATE_TIME_ERROR
      End Select
      Exit Function

   Case "m"
         ParseCmd = %DATE_TIME_CUSTOM
         sMask = Trim$(LCase$(asArg(2)))

   Case "t"
      ParseCmd = %TIME
      Exit Function
   Case Else
      ParseCmd = %DATE_TIME_ERROR
   End Select

End Function
'---------------------------------------------------------------------------

Sub ShowHelp()

Print "DateEcho prints the current date/time to STDOUT. This might be usefull"
Print "to log dates/times in batch processing jobs."
Print ""
Print "Usage: DateEcho /<Date|Time>=<Format>"
Print "  i.e. DateEcho /dt=l"
Print "- or -"
Print "       DateEcho /m=<custom format>"
Print "  i.e. DateEcho /m=yyyy-dd-mm-wd_hh"
Print ""
Print "Parameter"
Print "========="
Print ""
Print "<Date|Time>"
Print "-----------"
Print "/dt       - prints date and time"
Print "/d        - prints date only"
Print "/t        - prints time only"
Print "/m:<mask> - print date and/or time formatted as the supplied custom mask"
Print ""
Print "<Format>"
Print "--------"
Print "l - use long date format as defined in system settings (Control Panel)"
Print "s - use short date format as defined in system settings (Control Panel)"
Print ""
Print "Custom format"
Print "-------------"
Print "The following variables may be used with the mask parameter:"
Print "yyyy - 4-digit year"
Print "yy   - 2-digit year (with leading zero)"
Print "mm   - 2-digit month (with leading zero)"
Print "dd   - 2-digit day (with leading zero)"
Print "wd   - 1-digit day of week (0=Sunday ... 6=Saturday)"
Print "hh   - 2 digit hour, 24 h format (with leading zero)"
Print "nn   - 2 digit minute (with leading zero)"
Print "ss   - 2 digit second (with leading zero)"
Print "ms   - 3 digit millisecond (with leading zero)"
Print ""
Print "Any other character present will be returned 'as is'."

End Sub
'---------------------------------------------------------------------------

Function ShortDate(iLocale As Long) As String
  ' Retrieve Windows short date for the country(language ID) involved

  Local szDate As AsciiZ * 11, st As SYSTEMTIME
  GetLocalTime st
  GetDateFormat iLocale, %DATE_SHORTDATE, st, ByVal %Null, szDate, SizeOf(szDate)
  Function = szDate
End Function
'---------------------------------------------------------------------------

Function LongDate(iLocale As Long) As String
  ' Retrieve Windows long date for the country(language ID) involved

  Local szDate As AsciiZ * 64, st As SYSTEMTIME
  GetLocalTime st
  GetDateFormat iLocale, %DATE_LONGDATE, st, ByVal %Null, szDate, SizeOf(szDate)
  Function = szDate
End Function
'---------------------------------------------------------------------------

Function CustomDate(ByVal sMask As String) As String

   Local sResult As String, st As SYSTEMTIME
   GetLocalTime st

   sResult = sMask

   Replace "yyyy" With Format$(st.wYear, "0000") In sResult
   Replace "yy" With Right$(Format$(st.wYear, "0000"), 2) In sResult
   Replace "mm" With Format$(st.wMonth, "00") In sResult
   Replace "dd" With Format$(st.wDay, "00") In sResult
   Replace "wd" With Format$(st.wDayOfWeek, "00") In sResult
   Replace "hh" With Format$(st.wHour, "00") In sResult
   Replace "nn" With Format$(st.wMinute, "00") In sResult
   Replace "ss" With Format$(st.wSecond, "00") In sResult
   Replace "ms" With Format$(st.wMilliseconds, "000") In sResult

   CustomDate = sResult

End Function
'---------------------------------------------------------------------------
