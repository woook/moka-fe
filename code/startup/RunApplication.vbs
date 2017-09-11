'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2007
'
' NAME: 
'
' AUTHOR: neil.gaffney@gstt.nhs.uk
' DATE  : 09/10/2009
'
' COMMENT: 
'
'==========================================================================

Const OverwriteExisting = True
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oNet = CreateObject("WScript.NetWork")
Set oExplorer1 = CreateObject("InternetExplorer.Application")

sComputerName = oNet.ComputerName

sDestFolder = "C:\Program Files\GSTT\Moka\"
sAcessLoc = """C:\Program Files\Microsoft Office\OFFICE11\MSACCESS.EXE"""

sAppName = "Moka"
sDir = "\\gstt.local\data\Apps\Moka\FrontEnd\"
sTarget = "Moka.htm"

sAdminAccess = "C:\Program Files\GSTT\Moka\admin.mdb"
sSourceAdminAccess = "\\gstt.local\apps\Moka\FrontEnd\admin.mdb"

sGeneticsClinicAccess = "C:\Program Files\GSTT\Moka\clinicial.mdb"
sSourceGeneticsClinicAccess = "\\gstt.local\apps\Moka\FrontEnd\clinicial.mdb"

sCytogeneticsAccess = "C:\Program Files\GSTT\Moka\master.mdb"
sSourceCytogeneticsAccess = "\\gstt.local\apps\Moka\FrontEnd\master.mdb"

sAcessLoc = """C:\Program Files\Microsoft Office\OFFICE11\MSACCESS.EXE"""

If oFSO.FileExists(sAdminAccess) Then
Set sDestAdminAccessFile = oFSO.GetFile(sAdminAccess)
End If

If oFSO.FileExists(sSourceAdminAccess) Then
Set sSourceAdminAccessFile = oFSO.GetFile(sSourceAdminAccess)
End If 

If oFSO.FileExists(sGeneticsClinicAccess) Then
Set sDestClinicAccess = oFSO.GetFile(sGeneticsClinicAccess)
End If

If oFSO.FileExists(SSourceGeneticsClinicAccess) Then
Set sSourceGeneticsClinicAccessFile = oFSO.GetFile(sSourceGeneticsClinicAccess)
End If 

If oFSO.FileExists(sCytogeneticsAccess) Then
Set sDestCytogeneticsAccess = oFSO.GetFile(sCytogeneticsAccess)
End If 

If oFSO.FileExists(SSourceCytogeneticsAccess) Then
 
Set sSourceCytogeneticsAccessFile = oFSO.GetFile(sSourceCytogeneticsAccess)
End If 

oExplorer1.Navigate "\\gstt.local\data\Apps\Moka\FrontEnd\Moka.htm"
oExplorer1.ToolBar = 0
oExplorer1.StatusBar = 0
oExplorer1.Width = 450
oExplorer1.Height = 350
oExplorer1.Visible = 1
oExplorer1.Document.bgcolor = "#3bb9ff"
Do While (oExplorer1.Document.All.ButtonClicked.Value = "")
    Wscript.Sleep 250
Loop

sSelection = oExplorer1.Document.All.ButtonClicked.Value
oExplorer1.Quit
Wscript.Sleep 250

Select Case sSelection
Case "Genetics Clinic Access"
   CopyGeneticsAccess
   oShell.Run sAcessLoc & " " & """" & sGeneticsClinicAccess & """"
'wscript.echo sAcessLoc & " " & sClinicAccess
Case "Cytogenetics Access"
   CopyCytogeneticsAccess
 
   oShell.Run sAcessLoc & " " & """" & sCytogeneticsAccess & """"
'wscript.echo sAcessLoc & " " & sScientistAccess
Case "Admin Access"
   CopyAdminAccess 
   oShell.Run sAcessLoc & " " & """" & sAdminAccess & """"
'wscript.echo sAcessLoc & " " & sAdminAccess
Case "Cancelled"
     Wscript.Quit
End Select
Sub CopyGeneticsAccess
Set oClinicalFile = oFSO.OpenTextFile("C:\Program Files\GSTT\Moka\clinicial.txt", 1, false)
sClinicalVer = oClinicalFile.ReadLine
sClinicalDate = CStr(Day(sSourceGeneticsClinicAccessFile.DateLastModified) & "/" & (Month(sSourceGeneticsClinicAccessFile.DateLastModified)) & "/" & Year(sSourceGeneticsClinicAccessFile.DateLastModified)& " " & Hour(sSourceGeneticsClinicAccessFile.DateLastModified) & ":" & minute(sSourceGeneticsClinicAccessFile.DateLastModified)& ":" & second(sSourceGeneticsClinicAccessFile.DateLastModified))
oClinicalFile.Close
If sClinicalVer <> sClinicalDate Then
Set oExplorer1 = WScript.CreateObject("InternetExplorer.Application")
oExplorer1.Navigate "about:blank"   
oExplorer1.ToolBar = 0
oExplorer1.StatusBar = 0
oExplorer1.Width=400
oExplorer1.Height = 200 
oExplorer1.Top = 0
oExplorer1.Document.bgcolor = "#C8B560"
Do While (oExplorer1.Busy)
   Wscript.Sleep 200
Loop    
oExplorer1.Visible = 1             
oExplorer1.Document.Body.InnerHTML = "Retrieving updated Genetics Clinic file. " _
   & "This may take several minutes to complete."
   
oFSO.CopyFile sSourceGeneticsClinicAccessFile, sDestFolder, True
strComputer = "."
Set colServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2"). _
   ExecQuery("Select * from Win32_Service")
For Each objService In colServices
   Wscript.Sleep 200
Next
oExplorer1.Document.Body.InnerHTML = "Genetics Clinic file is now up to date."
oExplorer1.Quit
Set oClinicalFileW = oFSO.OpenTextFile("C:\Program Files\GSTT\Moka\clinicial.txt", 2)
oClinicalFileW.WriteLine sClinicalDate 
oClinicalFileW.Close
End If
End Sub

Sub CopyCytogeneticsAccess
 
    
Set oMasterFile = oFSO.OpenTextFile("C:\Program Files\GSTT\Moka\master.txt", 1, false)
sMasterVer = oMasterFile.ReadLine
sMasterDate = CStr(Day(sSourceCytogeneticsAccessFile.DateLastModified) & "/" & (Month(sSourceCytogeneticsAccessFile.DateLastModified)) & "/" & Year(sSourceCytogeneticsAccessFile.DateLastModified)& " " & Hour(sSourceCytogeneticsAccessFile.DateLastModified) & ":" & minute(sSourceCytogeneticsAccessFile.DateLastModified)& ":" & second(sSourceCytogeneticsAccessFile.DateLastModified)) 
oMasterFile.close
If sMasterVer <> sMasterDate Then
Set oExplorer1 = WScript.CreateObject("InternetExplorer.Application")
oExplorer1.Navigate "about:blank"   
oExplorer1.ToolBar = 0
oExplorer1.StatusBar = 0
oExplorer1.Width=400
oExplorer1.Height = 200 
oExplorer1.Top = 0
oExplorer1.Document.bgcolor = "#C8B560"
Do While (oExplorer1.Busy)
   Wscript.Sleep 200
Loop    
oExplorer1.Visible = 1             
oExplorer1.Document.Body.InnerHTML = "Retrieving updated Cytogenetics file. " _
   & "This may take several minutes to complete."
   
oFSO.CopyFile sSourceCytogeneticsAccessFile, sDestFolder, True
strComputer = "."
Set colServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2"). _
   ExecQuery("Select * from Win32_Service")
For Each objService In colServices
   Wscript.Sleep 200
Next
oExplorer1.Document.Body.InnerHTML = "Cytogenetics file is now up to date."
oExplorer1.Quit
Set oMasterFileW = oFSO.OpenTextFile("C:\Program Files\GSTT\Moka\master.txt", 2)
oMasterFileW.WriteLine sMasterDate
oMasterFileW.close
End If
 
End Sub

Sub CopyAdminAccess  
Set oAdminFile = oFSO.OpenTextFile("C:\Program Files\GSTT\Moka\Admin.txt", 1, false)
sAdminVer = oAdminFile.ReadLine
sAdminDate = CStr(Day(sSourceAdminAccessFile.DateLastModified) & "/" & (Month(sSourceAdminAccessFile.DateLastModified)) & "/" & Year(sSourceAdminAccessFile.DateLastModified)& " " & Hour(sSourceAdminAccessFile.DateLastModified) & ":" & minute(sSourceAdminAccessFile.DateLastModified)& ":" & second(sSourceAdminAccessFile.DateLastModified)) 
oAdminFile.Close
If sAdminVer <> sAdminDate Then
Set oExplorer1 = WScript.CreateObject("InternetExplorer.Application")
oExplorer1.Navigate "about:blank"   
oExplorer1.ToolBar = 0
oExplorer1.StatusBar = 0
oExplorer1.Width=400
oExplorer1.Height = 200 
oExplorer1.Top = 0
oExplorer1.Document.bgcolor = "#C8B560"
Do While (oExplorer1.Busy)
   Wscript.Sleep 200
Loop    
oExplorer1.Visible = 1             
oExplorer1.Document.Body.InnerHTML = "Retrieving updated Admin file. " _
   & "This may take several minutes to complete."
   
oFSO.CopyFile sSourceAdminAccessFile, sDestFolder, True
strComputer = "."
Set colServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2"). _
   ExecQuery("Select * from Win32_Service")
For Each objService In colServices
   Wscript.Sleep 200
Next
oExplorer1.Document.Body.InnerHTML = "Admin file is now up to date."
oExplorer1.Quit
Set oAdminFile = oFSO.OpenTextFile("C:\Program Files\GSTT\Moka\Admin.txt", 2)
    oAdminFile.WriteLine sAdminDate
    oAdminFile.Close
End If
End Sub 
