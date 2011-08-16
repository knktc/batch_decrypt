;批量解密工具
;knktc 2011-7-3

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListviewConstants.au3>
#include <EditConstants.au3>
#Include <GuiListView.au3>
#include <File.au3>

Global $DropFilesArr[1]

;批量解密主窗口的建立
$GUI_batch_decrypt = GUICreate("批量解密工具", 473, 342, -1, -1, -1, $WS_EX_ACCEPTFILES)
GUIRegisterMsg(0x233, "WM_DROPFILES_FUNC")

;“设置”菜单
$Menu_configure = GUICtrlCreateMenu("设置(&C)")
$Menu_configure_config = GUICtrlCreateMenuItem("运行设置", $Menu_configure)
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

;配置窗口的建立
$GUI_config_window = GUICreate("设置", 386, 224, -1, -1, $WS_DLGFRAME, $WS_EX_ACCEPTFILES, $GUI_batch_decrypt)

;Select gpg.exe path
$Label_select_gpg_path = GUICtrlCreateLabel("选择gpg.exe路径", 12, 26)
;GUICtrlSetColor($Label_select_gpg_path, 0xff0000)
$Input_gpgpath = GUICtrlCreateInput("", 12, 40, 276, 21)
GUICtrlSetState($Input_gpgpath, $GUI_DROPACCEPTED)
$Button_browse_gpgpath = GUICtrlCreateButton("浏览...", 296, 38, 75, 23)

;Select secret ring path
$Label_select_secring_path = GUICtrlCreateLabel("选择私钥路径", 12, 72)
$Input_secringpath = GUICtrlCreateInput("", 12, 86, 276, 21)
GUICtrlSetState($Input_secringpath, $GUI_DROPACCEPTED)
$Button_browse_secringpath = GUICtrlCreateButton("浏览...", 296, 84, 75, 23)

;Select public ring path
$Label_select_pubring_path = GUICtrlCreateLabel("选择公钥路径", 12, 118)
$Input_pubringpath = GUICtrlCreateInput("", 12, 132, 276, 21)
GUICtrlSetState($Input_pubringpath, $GUI_DROPACCEPTED)
$Button_browse_pubringpath = GUICtrlCreateButton("浏览...", 296, 130, 75, 23)

;ok and cancel
$Button_config_ok = GUICtrlCreateButton("确定", 213, 177, 75, 23)
$Button_config_cancel = GUICtrlCreateButton("取消", 296, 177, 75, 23)


;配置文件路径
$config_file_path = @WorkingDir & "\configure.ini"

;启动时检查是否有配置文件
;如果没有配置文件存在则弹出设置窗口要求用户进行第一次设置
while 1
	If FileExists($config_file_path) = 0 Then
		_FileCreate($config_file_path)				
		While 1
			$startup_msg = GUIGetMsg()
		GUISetState(@SW_SHOW, $GUI_config_window)
		Select
			Case $startup_msg = $Button_config_ok
				_GetAndWriteConfig($Input_gpgpath, $Input_secringpath, $Input_pubringpath, $config_file_path)
				GUISetState(@SW_HIDE, $GUI_config_window)
				ExitLoop
			
			Case $startup_msg = $Button_config_cancel
				GUISetState(@SW_HIDE, $GUI_config_window)
				ExitLoop
				
			;在设置中选择gpg.exe的地址
			Case $startup_msg = $Button_browse_gpgpath
				_BrowseGpgPath($Input_gpgpath)

			;在设置中选择私钥地址
			Case $startup_msg = $Button_browse_secringpath
				_BrowseSkrPath($Input_secringpath)
		
			;在设置中选择公钥地址
			Case $startup_msg = $Button_browse_pubringpath
				_BrowsePkrPath($Input_pubringpath)			
		EndSelect
		WEnd
	Else
		ExitLoop		
	EndIf
WEnd

GUISetState(@SW_SHOW, $GUI_batch_decrypt)

;等待接受按键信息
While 1
	$msg = GUIGetMsg()
	Select
		;接收关闭信息
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
		
		;点击“设置”--“运行设置”后弹出设置窗口
		Case $msg = $Menu_configure_config
			GUISetState(@SW_DISABLE, $GUI_batch_decrypt)
			GUISetState(@SW_SHOW, $GUI_config_window)
			_GetAndSetInput($config_file_path, $Input_gpgpath, $Input_secringpath, $Input_pubringpath)
		
		;点击设置窗口中的确定按钮后进行的操作
		Case $msg = $Button_config_ok
			_GetAndWriteConfig($Input_gpgpath, $Input_secringpath, $Input_pubringpath, $config_file_path)
			GUISetState(@SW_ENABLE, $GUI_batch_decrypt)
			GUISetState(@SW_HIDE, $GUI_config_window)
		
		;点击设置窗口中的取消按钮后进行的操作
		Case $msg = $Button_config_cancel
			GUISetState(@SW_ENABLE, $GUI_batch_decrypt)
			GUISetState(@SW_HIDE, $GUI_config_window)
		
		;按菜单中的退出按钮退出程序运行	
		Case $msg = $Menu_configure_exit
			ExitLoop
		
		;在设置中选择gpg.exe的地址
		Case $msg = $Button_browse_gpgpath
			_BrowseGpgPath($Input_gpgpath)

		;在设置中选择私钥地址
		Case $msg = $Button_browse_secringpath
			_BrowseSkrPath($Input_secringpath)
		
		;在设置中选择公钥地址
		Case $msg = $Button_browse_pubringpath
			_BrowsePkrPath($Input_pubringpath)
		
		;点击“帮助”--“关于”按钮后弹出关于信息
		Case $msg = $Menu_help_about
			MsgBox(0, "关于", "利用GnuPG来实现批量文件解密的工具" & @CRLF & "希望能帮助您解决一些小麻烦" & @CRLF & "2011 www.knktc.com")
		
		;点击“添加”按钮后选择需要解密的文件	
		Case $msg = $Button_add_file
			$add_file_path = FileOpenDialog("选择加密文件", @DesktopDir & "\", "加密文件 (*.asc; *.pgp) |所有文件 (*.*)", 1 + 2)
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
			$gpg_path = IniRead($config_file_path, "batch_decrypt", "gpg_path", "")
			$skr_path = IniRead($config_file_path, "batch_decrypt", "skr_path", "")
			$pkr_path = IniRead($config_file_path, "batch_decrypt", "pkr_path", "")
			$file_count = _GUICtrlListView_GetItemCount($Listview_encrypt_files)
			For $i = 0 To $file_count-1 
				$input_filepath = _GUICtrlListView_GetItemText($Listview_encrypt_files, $i)
				$output_filepath = _GetOutputFilepath($input_filepath)
				_DecryptSingleFile($gpg_path, $skr_path, $pkr_path, $input_filepath, $output_filepath, $password)
			Next
		
		Case $msg = $GUI_EVENT_DROPPED
            For $i = 1 To UBound($DropFilesArr)-1
				GUICtrlCreateListViewItem($DropFilesArr[$i], $Listview_encrypt_files)
			Next
			GUICtrlSendMsg($Listview_encrypt_files, $LVM_SETCOLUMNWIDTH, 0, -1)
	EndSelect
WEnd

;获取输出文件路径的函数
;使用点号来分隔文件路径，输出除最后一个点号之前的路径
Func _GetOutputFilepath($func_input_filepath)
	$func_path_array = StringSplit($func_input_filepath, ".")
	Local $length = 0
	$length = UBound($func_path_array)
	$length = $length - 2
	$func_output_filepath = _ArrayToString($func_path_array, ".", 1, $length)
	Return $func_output_filepath
EndFunc

Func _GetSpecOutputFilepath($func_input_filepath, $func_specify_folder)
	$split_filepath_array = StringSplit($func_input_filepath, "\", 1)
	$split_array_count = $split_filepath_array[0]
	$func_filename = $split_filepath_array[$split_array_count]
	$func_output_filepath = $func_specify_folder & "\" & $func_filename
	Return $func_output_filepath
EndFunc
	
;获取解密密码函数、
Func _GetPassword()
	$func_decrypt_password = InputBox("解密密码", "请输入解密密码", "" ,"*")
	Return $func_decrypt_password
EndFunc

;使用gpg解密单个文件的函数
Func _DecryptSingleFile($func_gpg_path, $func_skr_path, $func_pkr_path, $func_input_filepath, $func_output_filepath, $func_password)
	$command = '"' & $func_gpg_path & '"' _
		     & ' --secret-keyring ' & '"' & $func_skr_path & '"' _ 
			 &' --keyring ' & '"' & $func_pkr_path & '"' _ 
			 &' --passphrase=' & $func_password _ 
			 & ' -o ' & '"' & $func_output_filepath & '"' _ 
			 & ' -d ' & '"' & $func_input_filepath & '"'
	Run($command, "", @SW_HIDE)		
EndFunc

;浏览gpg.exe文件，并将选择的路径写入到路径输入框中
Func _BrowseGpgPath($func_input_gpgpath)
	$selected_gpg_path = FileOpenDialog("请选择gpg.exe文件位置", "@ProgramFilesDir", "Gnupg主程序(gpg.exe)|可执行程序(*.exe)", 1 + 2)
	GUICtrlSetData($func_input_gpgpath, $selected_gpg_path)
EndFunc

;浏览私钥文件，并将选择的路径写入到路径输入框中
Func _BrowseSkrPath($func_input_secringpath)
	$selected_skr_path = FileOpenDialog("选择私钥文件", @DesktopDir & "\", "所有文件 (*.*)", 1 + 2)
	GUICtrlSetData($func_input_secringpath, $selected_skr_path)
EndFunc

;浏览公钥文件，并将选择的路径写入到路径输入框中
Func _BrowsePkrPath($func_input_pubringpath)
	$selected_pkr_path = FileOpenDialog("选择公钥文件", @DesktopDir & "\", "所有文件 (*.*)", 1 + 2)
	GUICtrlSetData($func_input_pubringpath, $selected_pkr_path)
EndFunc

;读取输入框中的数据并将数据写入到配置文件中的函数
Func _GetAndWriteConfig($func_input_gpgpath, $func_input_secringpath, $func_input_pubringpath, $func_config_file_path)
	Local $func_gpgpath = GUICtrlRead($func_input_gpgpath)
	Local $func_secringpath = GUICtrlRead($func_input_secringpath)
	Local $func_pubringpath = GUICtrlRead($func_input_pubringpath)
	IniWrite($func_config_file_path, "batch_decrypt", "gpg_path", $func_gpgpath)
	IniWrite($func_config_file_path, "batch_decrypt", "skr_path", $func_secringpath)
	IniWrite($func_config_file_path, "batch_decrypt", "pkr_path", $func_pubringpath)
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

;用于读取配置文件中的配置信息并将信息显示到配置窗口相关输入框中的函数
Func _GetAndSetInput($func_config_file_path, $func_input_gpgpath, $func_input_secringpath, $func_input_pubringpath)
	$func_gpg_path = IniRead($func_config_file_path, "batch_decrypt", "gpg_path", "")
	GUICtrlSetData($func_input_gpgpath, $func_gpg_path)
	$func_skr_path = IniRead($func_config_file_path, "batch_decrypt", "skr_path", "")
	GUICtrlSetData($func_input_secringpath, $func_skr_path)
	$func_pkr_path = IniRead($func_config_file_path, "batch_decrypt", "pkr_path", "")
	GUICtrlSetData($func_input_pubringpath, $func_pkr_path)
EndFunc