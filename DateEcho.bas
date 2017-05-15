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

%VERSION_MAJOR = 1
%VERSION_MINOR = 0
%VERSION_REVISION = 3

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
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include "win32api.inc"
#Include "sautilcc.inc"

Declare Function ParseCmd(ByVal sCmd As String) As Long
Declare Sub ShowHelp()
Declare Function ShortDate(iLocale As Long) As String
Declare Function LongDate(iLocale As Long) As String
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
'------------------------------------------------------------------------------
   Local sTime As String
   Local lLocale As Long

   lLocale = GetUserDefaultLCID()

   ' Application intro
   Print ""
   ConHeadline "DateEcho", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2003-2017", $COMPANY_NAME
   Print ""

   Select Case As Long ParseCmd(Command$)
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
   End Select

End Function
'---------------------------------------------------------------------------

Function ParseCmd(ByVal sCmd As String) As Long
'------------------------------------------------------------------------------
'Purpose  : Parses the command line parameters
'
'Prereq.  : -
'Parameter: sCmd  - Parameters passed to the application
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
Print "Usage: DateEcho /<Date|Time>:<Format>"
Print "  i.e. DateEcho /dt=l"
Print ""
Print "Parameter"
Print "========="
Print ""
Print "<Date|Time>"
Print "-----------"
Print "/dt - prints date and time"
Print "/d  - prints date only"
Print "/t  - prints time only"
Print ""
Print "<Format>"
Print "--------"
Print "l - use long date format as defined in system settings (Control Panel)"
Print "s - use short date format as defined in system settings (Control Panel)"

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
