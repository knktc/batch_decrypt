;批量解密工具
;knktc 2011-7-3

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListviewConstants.au3>
#include <EditConstants.au3>
#Include <GuiListView.au3>
#include <File.au3>

Global $DropFilesArr[1]

GUICreate("批量解密工具", 473, 342, -1, -1, -1, $WS_EX_ACCEPTFILES)
GUIRegisterMsg(0x233, "WM_DROPFILES_FUNC")

;“设置”菜单
$Menu_configure = GUICtrlCreateMenu("设置(&C)")
$Menu_configure_gpgpath = GUICtrlCreateMenuItem("选择gpg.exe路径", $Menu_configure)
$Menu_configure_import_keys = GUICtrlCreateMenuItem("导入密钥", $Menu_configure)
$Menu_configure_separator1 = GUICtrlCreateMenuItem("", $Menu_configure)
$Menu_configure_exit = GUICtrlCreateMenuItem("退出", $Menu_configure)

;“帮助”菜单
$Menu_help = GUICtrlCreateMenu("帮助(&H)")
$Menu_help_about = GUICtrlCreateMenuItem("关于", $Menu_help)

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

;配置文件路径
$config_file_path = @WorkingDir & "\configure.ini"

;启动时检查是否有配置文件，是否已指定gpg.exe的位置
while 1	
	If FileExists($config_file_path) = 0 Then
		_FileCreate($config_file_path)
		_ChooseGpgPath($config_file_path)
	ElseIf IniRead($config_file_path, "batch_decrypt", "gpg_path" , "") = "" Then
		_ChooseGpgPath($config_file_path)
	Else
		ExitLoop
	EndIf
WEnd


GUISetState()

;等待接受按键信息
While 1
	$msg = GUIGetMsg()
	Select
		;接收关闭信息
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
		Case $msg = $Menu_configure_exit
			ExitLoop
		
		;点击“设置” -- "选择gpg.exe路径"按钮后配置gpg文件指向
		Case $msg = $Menu_configure_gpgpath
			_ChooseGpgPath($config_file_path)
		
		;点击“帮助”--“关于”按钮后弹出关于信息
		Case $msg = $Menu_help_about
			MsgBox(0, "关于", "利用GnuPG来实现批量文件解密的工具" & @CRLF & "希望能帮助您解决一些小麻烦" & @CRLF & "2011 www.knktc.com")
		
		
		;点击“添加”按钮后选择需要解密的文件	
		Case $msg = $Button_add_file
			$add_file_path = FileOpenDialog("选择加密文件", @DesktopDir & "\", "加密文件 (*.asc; *.pgp) |所有文件 (*.*)", 1 + 2 + 4)
			If $add_file_path <> "" Then
				GUICtrlCreateListViewItem($add_file_path, $Listview_encrypt_files)
				GUICtrlSendMsg($Listview_encrypt_files, $LVM_SETCOLUMNWIDTH, 0, -1)
			EndIf
		
		;点击删除按钮后删除选中的文件
		Case $msg = $Button_remove_file
			_GUICtrlListView_DeleteItemsSelected($Listview_encrypt_files)	
		
		;点击清空按钮后删除listview中所有的文件
		Case $msg = $Button_remove_all
			_GUICtrlListView_DeleteAllItems($Listview_encrypt_files)
		
		;点击“解密”按钮后开始解密列表中所有文件
		Case $msg = $Button_decrypt
			$password = Call("_GetPassword")		
			$gpg_path = IniRead($config_file_path, "batch_decrypt", "gpg_path" , "")
			$file_count = _GUICtrlListView_GetItemCount($Listview_encrypt_files)
			For $i = 0 To $file_count-1 
				$input_filepath = _GUICtrlListView_GetItemText($Listview_encrypt_files, $i)
				$output_filepath = _GetOutputFilepath($input_filepath)
				_DecryptSingleFile($gpg_path, $input_filepath, $output_filepath, $password)
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
	
;获取解密密码函数
Func _GetPassword()
	$func_decrypt_password = InputBox("解密密码", "请输入解密密码", "" ,"*")
	Return $func_decrypt_password
EndFunc

;使用gpg解密的函数
Func _DecryptSingleFile($func_gpg_path, $func_input_filepath, $func_output_filepath, $func_password)
	Run(@ComSpec & " /c " & '""' & $func_gpg_path & '""' & ' --passphrase=' & $func_password & ' -o ' & $func_output_filepath & ' -d ' &$func_input_filepath, "", @SW_HIDE)
EndFunc

Func _ChooseGpgPath($func_config_file_path)
	$selected_gpg_path = FileOpenDialog("请选择gpg.exe文件位置", "@ProgramFilesDir", "Gnupg主程序(gpg.exe)|可执行程序(*.exe)", 1 + 2)
	If $selected_gpg_path <> "" Then
		IniWrite($func_config_file_path, "batch_decrypt", "gpg_path", $selected_gpg_path)
	EndIf	
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
		