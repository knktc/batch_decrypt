;�������ܹ���
;knktc 2011-7-3

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListviewConstants.au3>
#include <EditConstants.au3>
#Include <GuiListView.au3>

Global $DropFilesArr[1]

GUICreate("�������ܹ���", 473, 342, -1, -1, -1, $WS_EX_ACCEPTFILES)
GUIRegisterMsg(0x233, "WM_DROPFILES_FUNC")

GUICtrlCreateLabel("�������ܵ��ļ�", 20, 16)
$Listview_encrypt_files = GUICtrlCreateListView("�����ļ�", 11, 33, 446, 150, -1, $WS_EX_CLIENTEDGE)
GUICtrlSendMsg(-1, $LVM_SETEXTENDEDLISTVIEWSTYLE, $LVS_EX_HEADERDRAGDROP, $LVS_EX_HEADERDRAGDROP)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)

GUICtrlCreateLabel("ѡ������ļ���", 20, 189)
$Editbox_output_folder = GUICtrlCreateEdit("", 11, 208, 408, 23, 0x1000 + $ES_AUTOHSCROLL)
GUICtrlCreateLabel("�粻ָ������ļ��У�����ܺ��ļ��洢�ڼ����ļ���ͬĿ¼��", 20, 241)

$Button_add_file = GUICtrlCreateButton("���", 23, 267, 75, 23)
$Button_remove_file = GUICtrlCreateButton("ɾ��", 141, 267, 75, 23)
$Button_remove_all = GUICtrlCreateButton("���", 249, 267, 75, 23)
$Button_decrypt = GUICtrlCreateButton("����", 369, 267, 75, 23)


GUISetState()

While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop

			
		Case $msg = $Button_add_file
		
		;���ɾ����ť��ɾ��ѡ�е��ļ�
		Case $msg = $Button_remove_file
			_GUICtrlListView_DeleteItemsSelected($Listview_encrypt_files)	
		
		;�����հ�ť��ɾ��listview�����е��ļ�
		Case $msg = $Button_remove_all
			_GUICtrlListView_DeleteAllItems($Listview_encrypt_files)
		
		Case $msg = $Button_decrypt
		$password = Call("_GetPassword")		
		$file_count = _GUICtrlListView_GetItemCount($Listview_encrypt_files)
		For $i = 0 To $file_count-1 
			$input_filepath = _GUICtrlListView_GetItemText($Listview_encrypt_files, $i)
			$output_filepath = _GetOutputFilepath($input_filepath)
			_DecryptSingleFile($input_filepath, $output_filepath, $password)
		Next
		
		Case $msg = $GUI_EVENT_DROPPED
            For $i = 1 To UBound($DropFilesArr)-1
				GUICtrlCreateListViewItem($DropFilesArr[$i], $Listview_encrypt_files)
			Next
			GUICtrlSendMsg($Listview_encrypt_files, $LVM_SETCOLUMNWIDTH, 0, -1)
	EndSelect
WEnd


Func _GetOutputFilepath($func_input_filepath)
	$func_output_filepath = StringTrimRight($func_input_filepath, 4)
	Return $func_output_filepath
EndFunc

Func _GetSpecOutputFilepath($func_input_filepath, $func_specify_folder)
	$split_filepath_array = StringSplit($func_input_filepath, "\", 1)
	$split_array_count = $split_filepath_array[0]
	$func_filename = $split_filepath_array[$split_array_count]
	$func_output_filepath = $func_specify_folder & "\" & $func_filename
	Return $func_output_filepath
EndFunc
	

Func _GetPassword()
	$func_decrypt_password = InputBox("��������", "�������������", "" ,"*")
	Return $func_decrypt_password
EndFunc
	
Func _DecryptSingleFile($func_input_filepath, $func_output_filepath, $func_password)
	Run(@ComSpec & " /c " & 'gpg --passphrase=' & $func_password & ' -o ' & $func_output_filepath & ' -d ' &$func_input_filepath, "", @SW_HIDE)
EndFunc

;�����ϳ���һ����ק�ļ���listview�еĺ���������Ҫ�о��¾����ʵ�ַ���
Func WM_DROPFILES_FUNC($hWnd, $msgID, $wParam, $lParam)
    Local $nSize, $pFileName
    Local $nAmt = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", 0xFFFFFFFF, "ptr", 0, "int", 255)
    For $i = 0 To $nAmt[0] - 1
        $nSize = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", $i, "ptr", 0, "int", 0)
        $nSize = $nSize[0] + 1
        $pFileName = DllStructCreate("char[" & $nSize & "]")
        DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", $i, "ptr", _
            DllStructGetPtr($pFileName), "int", $nSize)
        ReDim $DropFilesArr[$i + 2]
        $DropFilesArr[$i+1] = DllStructGetData($pFileName, 1)
        $pFileName = 0
    Next
    $DropFilesArr[0] = UBound($DropFilesArr)-1
EndFunc
		