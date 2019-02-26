Option Explicit

Include "FileUtilityClass.vbs"

Dim fileUtility
Set fileUtility = New FileUtility

Dim demoFileStr
demoFileStr = fileUtility.exportFileStr("demo.json")

WScript.Echo demoFileStr

' 疑似Include用関数
Sub Include(ByVal strFile)
  Dim objFSO , objStream , strDir

  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
  strDir = objFSO.GetFile(WScript.ScriptFullName).ParentFolder 

  Set objStream = objFSO.OpenTextFile(strDir & "\" & strFile, 1)

  ExecuteGlobal objStream.ReadAll() 
  objStream.Close 

  Set objStream = Nothing 
  Set objFSO = Nothing
End Sub