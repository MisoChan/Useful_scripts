Option Explicit

'--------------------------------------------------------------------
'
'   ディレクトリ内Excelファイルマクロ順次実行WSH
'
'
'   Author N.Kishino 2020/11/19
'
'--------------------------------------------------------------------




Dim filepath 
Dim macrofilepath 
Dim macrofilepath_array()
Dim macrofilename
Dim procname 

' WshShell オブジェクト
Dim objWshShell     
Set objWshShell = WScript.CreateObject("WScript.Shell")

'日時をyyyymmdd形式に変換
Function getDate(date)
	getDate = Year(date)
	getDate = getDate & Right( "0" & Month(date) , 2)
	getDate = getDate & Right( "0" & Day(date) , 2)
End Function


'フォルダ作成
Private function CreateFolder()
    
    Dim objFSO      ' FileSystemObject
    Dim strFolder   ' フォルダ名
    Dim strMessage  ' 表示用メッセージ
    
    strFolder = objWshShell.CurrentDirectory & "\" & "SQL_" & getDate(Now())
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then

        If objFSO.FolderExists(strFolder) = True Then
            objFSO.DeleteFolder(strFolder)
            WScript.Echo "フォルダ " & strFolder & " を削除します。"
        end if

        objFSO.CreateFolder(strFolder)
        If Err.Number = 0 Then
            strMessage = "フォルダ " & strFolder & " を作成しました。"
        Else
            strMessage = "エラー: " & Err.Description
        End If

        WScript.Echo strMessage
    Else
        WScript.Echo "エラー: " & Err.Description
    End If
	CreateFolder = strFolder 
end function



'ファイルパス指定 第一引数 エクセル格納ディレクトリを入力する。
filepath = WScript.Arguments(0)

'マクロが格納されたExcelを格納
macrofilepath = WScript.Arguments(1)

'実行するマクロ名
procname = WScript.Arguments(2)

'第一引数を元にファイル一覧を取得
Dim filereader
set filereader = CreateObject("Scripting.FileSystemObject")

set filereader = filereader.getFolder(filepath)

Dim excel 
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
'Excelを見えるようにして最前列表示
excel.Visible = true
CreateObject("WScript.Shell").AppActivate excel.Caption
Dim exportdir
exportdir=CreateFolder()

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
macrofilename  = fso.GetFileName(macrofilepath) 

Dim ext
Set ext = CreateObject("Scripting.FileSystemObject")

Dim excelpath
Dim file
for each file in filereader.files
	excelpath = filePath & "\" & file.Name
	If ext.GetExtensionName(excelpath) = "xlsx" Then	
		
		
		excel.Workbooks.Open macrofilepath
		excel.Workbooks.Open excelpath
		
		excel.Application.Run  macrofilename & "!" & procname
		
		excel.Workbooks.Close 
	End If
next 



