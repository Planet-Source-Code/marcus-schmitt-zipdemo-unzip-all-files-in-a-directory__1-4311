Attribute VB_Name = "ZipUnzip"
Public Type ZIPnames
    s(0 To 99) As String
End Type

Private Type ZPOPT
  fSuffix As Long
  fEncrypt As Long
  fSystem As Long
  fVolume As Long
  fExtra As Long
  fNoDirEntries As Long
  fExcludeDate As Long
  fIncludeDate As Long
  fVerbose As Long
  fQuiet As Long
  fCRLF_LF As Long
  fLF_CRLF As Long
  fJunkDir As Long
  fRecurse As Long
  fGrow As Long
  fForce As Long
  fMove As Long
  fDeleteEntries As Long
  fUpdate As Long
  fFreshen As Long
  fJunkSFX As Long
  fLatestTime As Long
  fComment As Long
  fOffsets As Long
  fPrivilege As Long
  fEncryption As Long
  fRepair As Long
  flevel As Byte
  date As String
  szRootDir As String
End Type

Private Type ZIPUSERFUNCTIONS
  DllPrnt As Long
  DLLPASSWORD As Long
  DLLCOMMENT As Long
  DLLSERVICE As Long
End Type


Dim MYOPT As ZPOPT
Dim MYUSER As ZIPUSERFUNCTIONS

Private Declare Function ZpInit Lib "zip32.dll" _
(ByRef Zipfun As ZIPUSERFUNCTIONS) As Long

Private Declare Function ZpSetOptions Lib "zip32.dll" _
(ByRef Opts As ZPOPT) As Long


Private Declare Function ZpGetOptions Lib "zip32.dll" () As ZPOPT

Private Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action


Private Type U_CBChar
  ch(32800) As Byte
End Type

Private Type Z_CBChar
  ch(4096) As Byte
End Type


Private Type CBCh
  ch(256) As Byte
End Type


Private Type DCLIST
  ExtractOnlyNewer As Long
  SpaceToUnderscore As Long
  PromptToOverwrite As Long
  fQuiet As Long
  ncflag As Long
  ntflag As Long
  nvflag As Long
  nUflag As Long
  nzflag As Long
  ndflag As Long
  noflag As Long
  naflag As Long
  nZIflag As Long
  C_flag As Long
  fPrivilege As Long
  Zip As String
  ExtractDir As String
End Type


Private Type USERFUNCTION
  DllPrnt As Long
  DLLSND As Long
  DLLREPLACE As Long
  DLLPASSWORD As Long
  DLLMESSAGE As Long
  DLLSERVICE As Long
  TotalSizeComp As Long
  TotalSize As Long
  CompFactor As Long
  NumMembers As Long
  cchComment As Integer
End Type


Private Type UZPVER
  structlen As Long
  flag As Long
  beta As String * 10
  date As String * 20
  zlib As String * 10
  Unzip(1 To 4) As Byte
  zipinfo(1 To 4) As Byte
  os2dll As Long
  windll(1 To 4) As Byte
End Type

Private Declare Function windll_unzip Lib "unzip32.dll" _
    (ByVal ifnc As Long, ByRef ifnv As ZIPnames, _
     ByVal xfnc As Long, ByRef xfnv As ZIPnames, _
     dcll As DCLIST, Userf As USERFUNCTION) As Long

Private Declare Sub UzpVersion2 Lib "unzip32.dll" _
    (uzpv As UZPVER)


Dim MYDCL As DCLIST
Dim U_MYUSER As USERFUNCTION
Dim MYVER As UZPVER

Public vbzipnum As Long, vbzipmes As String
Public vbzipnam As ZIPnames
Public vbxnames As ZIPnames



Function FnPtr(ByVal lp As Long) As Long
  FnPtr = lp
End Function


Function DllPrnt(ByRef fname As Z_CBChar, ByVal AnzChars As Long) As Long
  Dim t&, a$

  On Error Resume Next
  
  For t = 0 To AnzChars
    If fname.ch(t) <> 10 Then
      a = a + Chr(fname.ch(t))
    End If
  Next t
  DllPrnt = 0
End Function


Function DllPass(ByRef s1 As Byte, x As Long, _
  ByRef s2 As Byte, ByRef s3 As Byte) As Long

  On Error Resume Next
  
  
  DllPass = 1
End Function

Function DllRep(ByRef fname As U_CBChar) As Long
  Dim s0$, xx As Long

  On Error Resume Next
  
  DllRep = 100
  s0 = ""
  For xx = 0 To 255
    If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 + Chr(fname.ch(xx))
  Next xx
  xx = MsgBox("Datei '" + s0 + "' überschreiben?", vbYesNoCancel Or vbQuestion, "Frage")
  If xx = vbNo Then Exit Function
  If xx = vbCancel Then
    DllRep = 104
    Exit Function
  End If
  
  DllRep = 102
End Function


Sub ReceiveDllMessage(ByVal ucsize As Long, _
  ByVal csiz As Long, _
  ByVal cfactor As Integer, _
  ByVal mo As Integer, _
  ByVal dy As Integer, _
  ByVal yr As Integer, _
  ByVal hh As Integer, _
  ByVal mm As Integer, _
  ByVal c As Byte, ByRef fname As CBCh, _
  ByRef meth As CBCh, ByVal crc As Long, _
  ByVal fCrypt As Byte)
  Dim s0$, xx As Long
  Dim strout As String * 80

  
  On Error Resume Next
  strout = Space(80)
  If vbzipnum = 0 Then
      Mid$(strout, 1, 50) = "Filename:"
      Mid$(strout, 53, 4) = "Size"
      Mid$(strout, 62, 4) = "Date"
      Mid$(strout, 71, 4) = "Time"
      vbzipmes = strout + vbCrLf
      strout = Space(80)
  End If
  s0 = ""
  For xx = 0 To 255
      If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 & Chr$(fname.ch(xx))
  Next xx
  Mid$(strout, 1, 50) = Mid$(s0, 1, 50)
  Mid$(strout, 51, 7) = Right$("        " + Str$(ucsize), 7)
  Mid$(strout, 60, 3) = Right$(Str$(dy), 2) + "/"
  Mid$(strout, 63, 3) = Right$("0" + Trim$(Str$(mo)), 2) + "/"
  Mid$(strout, 66, 2) = Right$("0" + Trim$(Str$(yr)), 2)
  Mid$(strout, 70, 3) = Right$(Str$(hh), 2) + ":"
  Mid$(strout, 73, 2) = Right$("0" + Trim$(Str$(mm)), 2)
  
  vbzipmes = vbzipmes + strout + vbCrLf
  vbzipnum = vbzipnum + 1

End Sub



Function szTrim(szString As String) As String
  Dim pos As Integer, ln As Integer
  
  pos = InStr(szString, Chr$(0))
  ln = Len(szString)
  Select Case pos
    Case Is > 1
      szTrim = Trim(Left(szString, pos - 1))
    Case 1
      szTrim = ""
    Case Else
      szTrim = Trim(szString)
  End Select
  
End Function


Sub VBUnzip(fname As String, extdir As String, _
  prom As Integer, ovr As Integer, _
  mess As Integer, dirs As Integer, numfiles As Long, numxfiles As Long)
  Dim xx As Long
  
  MYDCL.ExtractOnlyNewer = 0      ' 1=extract only newer
  MYDCL.SpaceToUnderscore = 0     ' 1=convert space to underscore
  MYDCL.PromptToOverwrite = prom  ' 1=prompt to overwrite required
  MYDCL.fQuiet = 0                ' 2=no messages 1=less 0=all
  MYDCL.ncflag = 0                ' 1=write to stdout
  MYDCL.ntflag = 0                ' 1=test zip
  MYDCL.nvflag = mess             ' 0=extract 1=list contents
  MYDCL.nUflag = 0                ' 1=extract only newer
  MYDCL.nzflag = 0                ' 1=display zip file comment
  MYDCL.ndflag = dirs             ' 1=honour directories
  MYDCL.noflag = ovr              ' 1=overwrite files
  MYDCL.naflag = 0                ' 1=convert CR to CRLF
  MYDCL.nZIflag = 0               ' 1=Zip Info Verbose
  MYDCL.C_flag = 0                ' 1=Case insensitivity, 0=Case Sensitivity
  MYDCL.fPrivilege = 0            ' 1=ACL 2=priv
  MYDCL.Zip = fname               ' ZIP name
  MYDCL.ExtractDir = extdir       ' Extraction directory, NULL if extracting
                                  ' to current directory

  U_MYUSER.DllPrnt = FnPtr(AddressOf DllPrnt)
  U_MYUSER.DLLSND = 0& ' not supported
  U_MYUSER.DLLREPLACE = FnPtr(AddressOf DllRep)
  U_MYUSER.DLLPASSWORD = FnPtr(AddressOf DllPass)
  U_MYUSER.DLLMESSAGE = FnPtr(AddressOf ReceiveDllMessage)
  U_MYUSER.DLLSERVICE = 0& ' not coded yet :)
  
  
  With MYVER
    .structlen = Len(MYVER)
    .beta = Space$(9) & vbNullChar
    .date = Space$(19) & vbNullChar
    .zlib = Space$(9) & vbNullChar
  End With
  
 
  Call UzpVersion2(MYVER)

 
  xx = windll_unzip(numfiles, vbzipnam, numxfiles, vbxnames, MYDCL, U_MYUSER)
  If xx <> 0 Then
    MsgBox "Error: " & fname, vbCritical, "Fehler"
  End If

End Sub


Function DllComm(ByRef s1 As Z_CBChar) As Z_CBChar
  On Error Resume Next
  s1.ch(0) = vbNullString
  DllComm = s1
End Function

Function VBZip(argc As Integer, zipname As String, _
  mynames As ZIPnames, junk As Integer, _
  recurse As Integer, updat As Integer, _
  freshen As Integer, basename As String) As Long
  Dim hmem As Long, xx As Integer
  Dim retcode As Long
  
  On Error Resume Next
  
  
  MYUSER.DllPrnt = FnPtr(AddressOf DllPrnt)
  MYUSER.DLLPASSWORD = FnPtr(AddressOf DllPass)
  MYUSER.DLLCOMMENT = FnPtr(AddressOf DllComm)
  MYUSER.DLLSERVICE = 0&
  retcode = ZpInit(MYUSER)
  
 
  MYOPT.fSuffix = 0        ' include suffixes (not yet implemented)
  MYOPT.fEncrypt = 0       ' 1 if encryption wanted
  MYOPT.fSystem = 0        ' 1 to include system/hidden files
  MYOPT.fVolume = 0        ' 1 if storing volume label
  MYOPT.fExtra = 0         ' 1 if including extra attributes
  MYOPT.fNoDirEntries = 0  ' 1 if ignoring directory entries
  MYOPT.fExcludeDate = 0   ' 1 if excluding files earlier than a specified date
  MYOPT.fIncludeDate = 0   ' 1 if including files earlier than a specified date
  MYOPT.fVerbose = 0       ' 1 if full messages wanted
  MYOPT.fQuiet = 0         ' 1 if minimum messages wanted
  MYOPT.fCRLF_LF = 0       ' 1 if translate CR/LF to LF
  MYOPT.fLF_CRLF = 0       ' 1 if translate LF to CR/LF
  MYOPT.fJunkDir = junk    ' 1 if junking directory names
  MYOPT.fRecurse = recurse ' 1 if recursing into subdirectories
  MYOPT.fGrow = 0          ' 1 if allow appending to zip file
  MYOPT.fForce = 0         ' 1 if making entries using DOS names
  MYOPT.fMove = 0          ' 1 if deleting files added or updated
  MYOPT.fDeleteEntries = 0 ' 1 if files passed have to be deleted
  MYOPT.fUpdate = updat    ' 1 if updating zip file--overwrite only if newer
  MYOPT.fFreshen = freshen ' 1 if freshening zip file--overwrite only
  MYOPT.fJunkSFX = 0       ' 1 if junking sfx prefix
  MYOPT.fLatestTime = 0    ' 1 if setting zip file time to time of latest file in archive
  MYOPT.fComment = 0       ' 1 if putting comment in zip file
  MYOPT.fOffsets = 0       ' 1 if updating archive offsets for sfx Files
  MYOPT.fPrivilege = 0     ' 1 if not saving privelages
  MYOPT.fEncryption = 0    'Read only property!
  MYOPT.fRepair = 0        ' 1=> fix archive, 2=> try harder to fix
  MYOPT.flevel = 0         ' compression level - should be 0!!!
  MYOPT.date = vbNullString ' "12/31/79"? US Date?
  MYOPT.szRootDir = basename
  

  retcode = ZpSetOptions(MYOPT)
  
 
  retcode = ZpArchive(argc, zipname, mynames)
  
  VBZip = retcode
End Function



