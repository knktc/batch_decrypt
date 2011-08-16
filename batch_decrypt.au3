;�������ܹ���
;knktc 2011-7-3

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListviewConstants.au3>
#include <EditConstants.au3>
#Include <GuiListView.au3>
#include <File.au3>

Global $DropFilesArr[1]

;�������������ڵĽ���
$GUI_batch_decrypt = GUICreate("�������ܹ���", 473, 342, -1, -1, -1, $WS_EX_ACCEPTFILES)
GUIRegisterMsg(0x233, "WM_DROPFILES_FUNC")

;�����á��˵�
$Menu_configure = GUICtrlCreateMenu("����(&C)")
$Menu_configure_config = GUICtrlCreateMenuItem("��������", $Menu_configure)
$Menu_configure_separator1 = GUICtrlCreateMenuItem("", $Menu_configure)
$Menu_configure_exit = GUICtrlCreateMenuItem("�˳�", $Menu_configure)

;���������˵�
$Menu_help = GUICtrlCreateMenu("����(&H)")
$Menu_help_about = GUICtrlCreateMenuItem("����", $Menu_help)

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

;���ô��ڵĽ���
$GUI_config_window = GUICreate("����", 386, 224, -1, -1, $WS_DLGFRAME, $WS_EX_ACCEPTFILES, $GUI_batch_decrypt)

;Select gpg.exe path
$Label_select_gpg_path = GUICtrlCreateLabel("ѡ��gpg.exe·��", 12, 26)
;GUICtrlSetColor($Label_select_gpg_path, 0xff0000)
$Input_gpgpath = GUICtrlCreateInput("", 12, 40, 276, 21)
GUICtrlSetState($Input_gpgpath, $GUI_DROPACCEPTED)
$Button_browse_gpgpath = GUICtrlCreateButton("���...", 296, 38, 75, 23)

;Select secret ring path
$Label_select_secring_path = GUICtrlCreateLabel("ѡ��˽Կ·��", 12, 72)
$Input_secringpath = GUICtrlCreateInput("", 12, 86, 276, 21)
GUICtrlSetState($Input_secringpath, $GUI_DROPACCEPTED)
$Button_browse_secringpath = GUICtrlCreateButton("���...", 296, 84, 75, 23)

;Select public ring path
$Label_select_pubring_path = GUICtrlCreateLabel("ѡ��Կ·��", 12, 118)
$Input_pubringpath = GUICtrlCreateInput("", 12, 132, 276, 21)
GUICtrlSetState($Input_pubringpath, $GUI_DROPACCEPTED)
$Button_browse_pubringpath = GUICtrlCreateButton("���...", 296, 130, 75, 23)

;ok and cancel
$Button_config_ok = GUICtrlCreateButton("ȷ��", 213, 177, 75, 23)
$Button_config_cancel = GUICtrlCreateButton("ȡ��", 296, 177, 75, 23)


;�����ļ�·��
$config_file_path = @WorkingDir & "\configure.ini"

;����ʱ����Ƿ��������ļ�
;���û�������ļ������򵯳����ô���Ҫ���û����е�һ������
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
				
			;��������ѡ��gpg.exe�ĵ�ַ
			Case $startup_msg = $Button_browse_gpgpath
				_BrowseGpgPath($Input_gpgpath)

			;��������ѡ��˽Կ��ַ
			Case $startup_msg = $Button_browse_secringpath
				_BrowseSkrPath($Input_secringpath)
		
			;��������ѡ��Կ��ַ
			Case $startup_msg = $Button_browse_pubringpath
				_BrowsePkrPath($Input_pubringpath)			
		EndSelect
		WEnd
	Else
		ExitLoop		
	EndIf
WEnd

GUISetState(@SW_SHOW, $GUI_batch_decrypt)

;�ȴ����ܰ�����Ϣ
While 1
	$msg = GUIGetMsg()
	Select
		;���չر���Ϣ
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
		
		;��������á�--���������á��󵯳����ô���
		Case $msg = $Menu_configure_config
			GUISetState(@SW_DISABLE, $GUI_batch_decrypt)
			GUISetState(@SW_SHOW, $GUI_config_window)
			_GetAndSetInput($config_file_path, $Input_gpgpath, $Input_secringpath, $Input_pubringpath)
		
		;������ô����е�ȷ����ť����еĲ���
		Case $msg = $Button_config_ok
			_GetAndWriteConfig($Input_gpgpath, $Input_secringpath, $Input_pubringpath, $config_file_path)
			GUISetState(@SW_ENABLE, $GUI_batch_decrypt)
			GUISetState(@SW_HIDE, $GUI_config_window)
		
		;������ô����е�ȡ����ť����еĲ���
		Case $msg = $Button_config_cancel
			GUISetState(@SW_ENABLE, $GUI_batch_decrypt)
			GUISetState(@SW_HIDE, $GUI_config_window)
		
		;���˵��е��˳���ť�˳���������	
		Case $msg = $Menu_configure_exit
			ExitLoop
		
		;��������ѡ��gpg.exe�ĵ�ַ
		Case $msg = $Button_browse_gpgpath
			_BrowseGpgPath($Input_gpgpath)

		;��������ѡ��˽Կ��ַ
		Case $msg = $Button_browse_secringpath
			_BrowseSkrPath($Input_secringpath)
		
		;��������ѡ��Կ��ַ
		Case $msg = $Button_browse_pubringpath
			_BrowsePkrPath($Input_pubringpath)
		
		;�����������--�����ڡ���ť�󵯳�������Ϣ
		Case $msg = $Menu_help_about
			MsgBox(0, "����", "����GnuPG��ʵ�������ļ����ܵĹ���" & @CRLF & "ϣ���ܰ��������һЩС�鷳" & @CRLF & "2011 www.knktc.com")
		
		;�������ӡ���ť��ѡ����Ҫ���ܵ��ļ�	
		Case $msg = $Button_add_file
			$add_file_path = FileOpenDialog("ѡ������ļ�", @DesktopDir & "\", "�����ļ� (*.asc; *.pgp) |�����ļ� (*.*)", 1 + 2)
			If $add_file_path <> "" Then
				GUICtrlCreateListViewItem($add_file_path, $Listview_encrypt_files)
				GUICtrlSendMsg($Listview_encrypt_files, $LVM_SETCOLUMNWIDTH, 0, -1)
			EndIf		
		
		;���ɾ����ť��ɾ��ѡ�е��ļ�
		Case $msg = $Button_remove_file
			_GUICtrlListView_DeleteItemsSelected($Listview_encrypt_files)
		
		;�����հ�ť��ɾ��listview�����е��ļ�
		Case $msg = $Button_remove_all
			_GUICtrlListView_DeleteAllItems($Listview_encrypt_files)

		;��������ܡ���ť��ʼ�����б��������ļ�
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

;��ȡ����ļ�·���ĺ���
;ʹ�õ�����ָ��ļ�·������������һ�����֮ǰ��·��
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
	
;��ȡ�������뺯����
Func _GetPassword()
	$func_decrypt_password = InputBox("��������", "�������������", "" ,"*")
	Return $func_decrypt_password
EndFunc

;ʹ��gpg���ܵ����ļ��ĺ���
Func _DecryptSingleFile($func_gpg_path, $func_skr_path, $func_pkr_path, $func_input_filepath, $func_output_filepath, $func_password)
	$command = '"' & $func_gpg_path & '"' _
		     & ' --secret-keyring ' & '"' & $func_skr_path & '"' _ 
			 &' --keyring ' & '"' & $func_pkr_path & '"' _ 
			 &' --passphrase=' & $func_password _ 
			 & ' -o ' & '"' & $func_output_filepath & '"' _ 
			 & ' -d ' & '"' & $func_input_filepath & '"'
	Run($command, "", @SW_HIDE)		
EndFunc

;���gpg.exe�ļ�������ѡ���·��д�뵽·���������
Func _BrowseGpgPath($func_input_gpgpath)
	$selected_gpg_path = FileOpenDialog("��ѡ��gpg.exe�ļ�λ��", "@ProgramFilesDir", "Gnupg������(gpg.exe)|��ִ�г���(*.exe)", 1 + 2)
	GUICtrlSetData($func_input_gpgpath, $selected_gpg_path)
EndFunc

;���˽Կ�ļ�������ѡ���·��д�뵽·���������
Func _BrowseSkrPath($func_input_secringpath)
	$selected_skr_path = FileOpenDialog("ѡ��˽Կ�ļ�", @DesktopDir & "\", "�����ļ� (*.*)", 1 + 2)
	GUICtrlSetData($func_input_secringpath, $selected_skr_path)
EndFunc

;�����Կ�ļ�������ѡ���·��д�뵽·���������
Func _BrowsePkrPath($func_input_pubringpath)
	$selected_pkr_path = FileOpenDialog("ѡ��Կ�ļ�", @DesktopDir & "\", "�����ļ� (*.*)", 1 + 2)
	GUICtrlSetData($func_input_pubringpath, $selected_pkr_path)
EndFunc

;��ȡ������е����ݲ�������д�뵽�����ļ��еĺ���
Func _GetAndWriteConfig($func_input_gpgpath, $func_input_secringpath, $func_input_pubringpath, $func_config_file_path)
	Local $func_gpgpath = GUICtrlRead($func_input_gpgpath)
	Local $func_secringpath = GUICtrlRead($func_input_secringpath)
	Local $func_pubringpath = GUICtrlRead($func_input_pubringpath)
	IniWrite($func_config_file_path, "batch_decrypt", "gpg_path", $func_gpgpath)
	IniWrite($func_config_file_path, "batch_decrypt", "skr_path", $func_secringpath)
	IniWrite($func_config_file_path, "batch_decrypt", "pkr_path", $func_pubringpath)
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

;���ڶ�ȡ�����ļ��е�������Ϣ������Ϣ��ʾ�����ô������������еĺ���
Func _GetAndSetInput($func_config_file_path, $func_input_gpgpath, $func_input_secringpath, $func_input_pubringpath)
	$func_gpg_path = IniRead($func_config_file_path, "batch_decrypt", "gpg_path", "")
	GUICtrlSetData($func_input_gpgpath, $func_gpg_path)
	$func_skr_path = IniRead($func_config_file_path, "batch_decrypt", "skr_path", "")
	GUICtrlSetData($func_input_secringpath, $func_skr_path)
	$func_pkr_path = IniRead($func_config_file_path, "batch_decrypt", "pkr_path", "")
	GUICtrlSetData($func_input_pubringpath, $func_pkr_path)
EndFunc