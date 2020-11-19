Option Explicit

'--------------------------------------------------------------------
'
'   �f�B���N�g����Excel�t�@�C���}�N���������sWSH
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

' WshShell �I�u�W�F�N�g
Dim objWshShell     
Set objWshShell = WScript.CreateObject("WScript.Shell")

'������yyyymmdd�`���ɕϊ�
Function getDate(date)
	getDate = Year(date)
	getDate = getDate & Right( "0" & Month(date) , 2)
	getDate = getDate & Right( "0" & Day(date) , 2)
End Function


'�t�H���_�쐬
Private function CreateFolder()
    
    Dim objFSO      ' FileSystemObject
    Dim strFolder   ' �t�H���_��
    Dim strMessage  ' �\���p���b�Z�[�W
    
    strFolder = objWshShell.CurrentDirectory & "\" & "SQL_" & getDate(Now())
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then

        If objFSO.FolderExists(strFolder) = True Then
            objFSO.DeleteFolder(strFolder)
            WScript.Echo "�t�H���_ " & strFolder & " ���폜���܂��B"
        end if

        objFSO.CreateFolder(strFolder)
        If Err.Number = 0 Then
            strMessage = "�t�H���_ " & strFolder & " ���쐬���܂����B"
        Else
            strMessage = "�G���[: " & Err.Description
        End If

        WScript.Echo strMessage
    Else
        WScript.Echo "�G���[: " & Err.Description
    End If
	CreateFolder = strFolder 
end function



'�t�@�C���p�X�w�� ������ �G�N�Z���i�[�f�B���N�g������͂���B
filepath = WScript.Arguments(0)

'�}�N�����i�[���ꂽExcel���i�[
macrofilepath = WScript.Arguments(1)

'���s����}�N����
procname = WScript.Arguments(2)

'�����������Ƀt�@�C���ꗗ���擾
Dim filereader
set filereader = CreateObject("Scripting.FileSystemObject")

set filereader = filereader.getFolder(filepath)

Dim excel 
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
'Excel��������悤�ɂ��čőO��\��
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



