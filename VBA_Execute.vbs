Option Explicit

'--------------------------------------------------------------------
'
'   ディレクトリ内Excelファイルマクロ順次実行WSH
'   第一引数： マクロ実行対象Excelディレクトリパス
'   第二引数： マクロ格納Excelパス
'   第三引数： マクロ実行プロシージャ/関数名
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



