;批量解密工具
;knktc 2011-7-3

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListviewConstants.au3>
#include <EditConstants.au3>
#Include <GuiListView.au3>

Global $DropFilesArr[1]

GUICreate("批量解密工具", 473, 342, -1, -1, -1, $WS_EX_ACCEPTFILES)
GUIRegisterMsg(0x233, "WM_DROPFILES_FUNC")

GUICtrlCreateLabel("添加需解密的文件", 20, 16)
$Listview_encrypt_files = GUICtrlCreateListView("解密文件", 11, 33, 446, 150, -1, $WS_EX_CLIENTEDGE)
GUICtrlSendMsg(-1, $LVM_SETEXTENDEDLISTVIEWSTYLE, $LVS_EX_HEADERDRAGDROP, $LVS_EX_HEADERDRAGDROP)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)

GUICtrlCreateLabel("选择输出文件夹", 20, 189)
$Editbox_output_folder = GUICtrlCreateEdit("", 11, 208, 408, 23, 0x1000 + $ES_AUTOHSCROLL)
GUICtrlCreateLabel("如不指定输出文件夹，则解密后文件存储于加密文件相同目录下", 20, 241)

$Button_add_file = GUICtrlCreateButton("添加", 23, 267, 75, 23)
$Button_remove_file = GUICtrlCreateButton("删除", 141, 267, 75, 23)
$Button_remove_all = GUICtrlCreateButton("清空", 249, 267, 75, 23)
$Button_decrypt = GUICtrlCreateButton("解密", 369, 267, 75, 23)


GUISetState()

While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop

			
		Case $msg = $Button_add_file
		
		;点击删除按钮后删除选中的文件
		Case $msg = $Button_remove_file
			_GUICtrlListView_DeleteItemsSelected($Listview_encrypt_files)	
		
		;点击清空按钮后删除listview中所有的文件
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
	$func_decrypt_password = InputBox("解密密码", "请输入解密密码", "" ,"*")
	Return $func_decrypt_password
EndFunc
	
Func _DecryptSingleFile($func_input_filepath, $func_output_filepath, $func_password)
	Run(@ComSpec & " /c " & 'gpg --passphrase=' & $func_password & ' -o ' & $func_output_filepath & ' -d ' &$func_input_filepath, "", @SW_HIDE)
EndFunc

;从网上抄了一个拖拽文件到listview中的函数，还需要研究下具体的实现方法
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
		