'�錾(�s����ꕔ�ȗ��j
Dim objFSO
Dim objRec
Dim objDel
Dim objDir
Dim objLog
Dim time_full
Dim CRLF
Dim message
Dim report
Dim log
Dim mode
Dim earth
Dim earth_e
Dim Setting(1)
Dim kion_conf_13
Dim kousuiryou_conf_13
Dim kion_t
Dim kousuiryou_t
Dim quit_conf
Dim kion_conf
Dim kousuiryou_conf
Dim Error_check
Dim Error_Number
Dim summer(3)
Dim winter(3)
Dim kannsou_genkai
Dim kansou
Dim kansou_jpn
Dim sum_kousuiryou
Dim max_kousuiryou
Dim max_kousuiryou_month
Dim min_kousuiryou
Dim min_kousuiryou_month
Dim sum_kion
Dim ave_kion
Dim max_kion
Dim max_kion_month
Dim min_kion
Dim min_kion_month
Dim warm_months
Dim dbl
Dim kikoutai
Dim kikouku
Dim kikouku_d
Dim kikouku_log
Dim kion_double(11)
Dim kousuiryou_double(11)
Dim input_kion
Dim def_kion
Dim input_kousuiryou
Dim def_kousuiryou

'1.3 2.5 1.4 7.3 4.7 2.5 1.0 2.4 1.5 3.37 13.5 224.5682

'�z��̏�����
for i = 1 to 12
	kion_double(i-1) = CDbl(0)
	kousuiryou_double(i-1) = CDbl(0)
next

'FileSystemObject
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

CRLF = Chr(13) & Chr(10)
def_kion = "�C��(��)"
def_kousuiryou = "�~����(mm)"
mode = "Normal"
Setting(0) = "1"
log = "First Err.Number = " & Err.Number & Chr(13)
On Error Resume next
Error_check = CDbl(0.0)


for i = 1 to 1000000000

	Error_Number = "0"
	Err.Clear
	
	log = log & "-----Kion For:" & CStr(i) & Chr(13) & Chr(13)
	
	input_kion = inputbox("1������12���܂ł̋C�����X�y�[�X�ŋ�؂��ē��͂��Ă��������B",mode,def_kion)

	log = log & "Kion Input:" & input_kion & Chr(13)
	'�f�o�b�O���[�h
	If input_kion = "mode debug" then
		mode = "Debug"
		log = log & "Mode changed to Debug" & Chr(13)
		Error_Number = "debug"
	End If

	'�m�[�}�����[�h
	If input_kion = "mode normal" then
		mode = "Normal"
		log = log & "Mode changed to Normal" & Chr(13)
		Error_Number = "Normal"
	End If

	'�����G�ߌ^�C��Off
	If input_kion = "Setting 0 Off" then
		Setting(0) = "0"
		log = log & "Setting 0 changed to Off" & Chr(13)
		Error_Number = "Setting 0 Off"
	End If

	'�����G�ߌ^�C��On
	If input_kion = "Setting 0 On" then
		Setting(0) = "1"
		log = log & "Setting 0 changed to On" & Chr(13)
		Error_Number = "Setting 0 On"
	End If



	'�L�����Z������
	If IsEmpty(input_kion) = true then
		Error_Number = "cancell"
		log = log & "Cancell Pressed" & Chr(13)
		quit_conf = msgbox ("�{���ɏI�����܂���? ���͂����f�[�^�͎����܂��B",vbYesNo,mode)
		If quit_conf = 6 then
			log = log & "Confirmed by user" & Chr(13) & "Quitting..." & Chr(13)
			log = Replace(CStr(log),Chr(13),CRLF)
			Set objFile = objFSO.OpenTextFile("Keppen-v9.log",8,True)
			objFile.WriteLine("-------------------------------" & log & Chr(13) & Chr(13) & Chr(13))
			objFile.Close
			WScript.quit
			log = log & "Quit Failed" & Chr(13)
		Else
			log = log & "Continued by user" & Chr(13)
		End If
	End If

	'���[�̋󔒂��폜
	input_kion = RTrim(input_kion)
	input_kion = LTrim(input_kion)
	'�S�p�X�y�[�X�𔼊p�ɒu��
	input_kion = Replace (input_kion,"�@"," ")
	'                                 ������͑S�p�X�y�[�X
	'2�ȏ�̘A���������p�X�y�[�X���ЂƂɂ���
	for k = 1 to 100
		input_kion = Replace(input_kion,"  "," ")
	'                                        ������͔��p�X�y�[�X�~2
	next


	'��OK����
	If input_kion = "" And Error_Number = "0" then
		log = log & "Error:Kion Input is Empty" & Chr(13)
		Error_Number = "kion_empty"
		msgbox ("�C������͂��Ă�������")
	End If

	kion = split(input_kion)
	log = log & "Kion Split Err.Number = " & CStr(Err.Number) & Chr(13)
	Err.Clear

	for n = 1 to 1000000
		kion_t = kion_t & n & "��:" & kion(n-1) & Chr(13)
		If Err.Number = 9 then
			log = log & "End making kion_t:" & n & Chr(13)
			exit for
		ElseIf Err.Number <> 0 then
			msgbox("�\�����ʃG���[�������������ߏI�����܂��B�G���[�ԍ������`����������" & Chr(13) & "�G���[�ԍ�:" & Err.Number)
			WScript.Quit
		End If
	next
	Err.Clear

	'��������
	for j = 1 to 12
		Error_Check = CDbl(kion(j-1))
		If Err.Number = 13 And Error_Number = "0" then
			log = log & "Error:Kion Not a number   " & CStr(j) & " = " & kion(j-1) & Chr(13)
			msgbox ("�����ȊO�̕�������͂��Ȃ��ł�������" & Chr(13) & Chr(13) & kion_t)
			Error_Number = "kion_Nan"
			exit for
		End If
		Err.Clear
	next

	'�X�y�[�X��؂蔻��
	If InStr(input_kion," ") = 0 And Error_Number = "0" then
		log = log & "Error:Kion String space not found" & Chr(13)
		Error_Number = "kion_no_space"
		msgbox ("�C���ƋC���̊Ԃ̓X�y�[�X�ŋ�؂��Ă��������B" & Chr(13) & Chr(13) & kion_t)
	End If
	
	If IsEmpty(kion(11)) then
		If Err.Number = 9  And Error_Number = "0" then
			Error_Number = "kion<12"
			log = log & "Error:Kion 12th Number is Empty" & Chr(13)
			msgbox ("�C����12���͂���Ă��܂���B���肷��ɂ�1������12���܂ł̋C�����K�v�ł��B" & Chr(13) & Chr(13) & kion_t)
			Err.Clear
		End If
	End If

	If IsEmpty(kion(12)) then
	End If

	If Err.Number = 0  And Error_Number = "0" then
		log = log & "Warning:There is 13th kion number" & Chr(13)
		kion_conf_13 = msgbox("�C����13�ȏ���͂���Ă��܂��B13�ڈȍ~�͌v�Z���疳������܂�����낵���ł��傤��?" & Chr(13) & Chr(13) & kion_t,vbYesNo)
		If (kion_conf_13) = 7 then
			Error_Number = "kion_13"
			log = log & "Returned to inputbox" & Chr(13)
		Else
			log = log & "Continued" & Chr(13)
		End If
	End If
	Err.Clear

	'�ُ�l����
	Err.Clear
	for m = 1 to 12
		If CDbl(kion(m-1)) < CDbl(-100) And Error_Number = "0" then
			If Err.Number = 0 then
				log = log & "Warning:Kion less than -100   " & CStr(m) & " = " & kion(m-1) & "��" & Chr(13)
				kion_conf = msgbox("�C���ɋɒ[�ɏ������l������܂�:" & m & "��" & kion(m-1) & "�x" & Chr(13) & "���s���܂����H" & Chr(13) & Chr(13) & kion_t,vbYesNo)
				If kion_conf = 7 then
					Error_Number = "Kion_Low"
					log = log & "Returned to inputbox" & Chr(13)
				End If
				If kion_conf = 6 then
					log = log & "Continued" & Chr(13)
					Exit for
				End If
			End If
		End If
		
		def_kion = input_kion
		
		If Err.Number = 0 then
			If CDbl(kion(m-1)) > CDbl(100) And Error_Number = "0" then
				log = log & "Warning:Kion more than 100   " & CStr(m) & " = " & kion(m-1) & "��" & Chr(13)
				kion_conf = msgbox("�C���ɋɒ[�ɑ傫���l������܂�:" & m & "��" & kion(m-1) & "�x" & Chr(13) & "���s���܂����H" & Chr(13) & Chr(13) & kion_t,vbYesNo)
				If kion_conf = 7 then
					Error_Number = "Kion_High"
					log = log & "Returned to inputbox" & Chr(13)
				End If
				If kion_conf = 6 then
					log = log & "Continued" & Chr(13)
					Exit for
				End If
			End If
		End If
	next


	If Error_Number = "0" And Err.Number = 0 then
		log = log & "Kion Final Error check passed" & Chr(13)
		Exit For
	ElseIf Error_Number = "0" then
		log = log & "Error:Kion Final Error Check Failed" & Chr(13) & "Err.Number = " & Err.Number & Chr(13)
	Else
		log = log & "Error:Kion Final Error Check Failed" & Chr(13) & "Error_Number = " & Error_Number & Chr(13)
	End If
next

log = log & "Kion Input Finished" & Chr(13)

for i = 1 to 1000000000
	Err.Clear
	Error_Number = "0"
	kousuiryou_t = ""
	log = log & "-----Rain For:" & CStr(i) & Chr(13) & Chr(13)
	input_kousuiryou = inputbox("1������12���܂ł̍~���ʂ��X�y�[�X�ŋ�؂��ē��͂��Ă��������B",mode,def_kousuiryou)
	log = log & "Rain Input:" & input_kousuiryou & Chr(13)

	'�f�o�b�O���[�h
	If input_kousuiryou = "mode debug" then
		mode = "Debug"
		log = log & "Mode changed to Debug" & Chr(13)
		Error_Number = "debug"
	End If

	'�m�[�}�����[�h
	If input_kousuiryou = "mode normal" then
		mode = "Normal"
		log = log & "Mode changed to Normal" & Chr(13)
		Error_Number = "Normal"
	End If

	'�����G�ߌ^�C��Off
	If input_kion = "Setting 0 Off" then
		Setting(0) = "0"
		log = log & "Setting 0 changed to Off" & Chr(13)
		Error_Number = "Setting 0 Off"
	End If

	'�����G�ߌ^�C��On
	If input_kion = "Setting 0 On" then
		Setting(0) = "1"
		log = log & "Setting 0 changed to On" & Chr(13)
		Error_Number = "Setting 0 On"
	End If


	'�L�����Z������
	If IsEmpty(input_kousuiryou) = true then
		Error_Number = "cancell"
		log = log & "Cancell Pressed"
		quit_conf = msgbox ("�{���ɏI�����܂���? ���͂����f�[�^�͎����܂��B",vbYesNo,mode)
		If quit_conf = 6 then
			log = log & "Confirmed by user" & Chr(13) & "Quitting" & Chr(13)
			log = Replace(CStr(log),Chr(13),CRLF)
			Set objFile = objFSO.OpenTextFile("Keppen-v9.log",8,True)
			objFile.WriteLine("-------------------------------" & log & Chr(13) & Chr(13) & Chr(13))
			objFile.Close
			WScript.quit
			log = log & "Quit Failed" & Chr(13)
		Else
			log = log & "Continued by user" & Chr(13)
		End If
	End If

	input_kousuiryou = RTrim(input_kousuiryou)
	input_kousuiryou = LTrim(input_kousuiryou)
	input_kousuiryou = Replace (input_kousuiryou,"�@"," ")
	for k = 1 to 100
		input_kousuiryou = Replace(input_kousuiryou,"  "," ")
	next

	'�󂩂ǂ�������
	If input_kousuiryou = "" And Error_Number = "0" then
		Error_Number = "kousuiryou_empty"
		log = log & "Error:Rain input is Empty" & Chr(13)
		msgbox ("�~���ʂ���͂��Ă�������")
	End If

	kousuiryou = split(input_kousuiryou)
	log = log & "Rain Split Err.Number = " & CStr(Err.Number) & Chr(13)
	Err.Clear

	for n = 1 to 1000000
		kousuiryou_t = kousuiryou_t & n & "��:" & kousuiryou(n-1) & Chr(13)
		If Err.Number = 9 then
			log = log & "End making kousuiryou_t:" & n & Chr(13)
			exit for
		ElseIf Err.Number <> 0 then
			msgbox("�\�����ʃG���[�������������ߏI�����܂��B�G���[�ԍ������`����������" & Chr(13) & "�G���[�ԍ�:" & Err.Number)
			WScript.Quit
		End If
	next
	Err.Clear

	'��������
	for j = 1 to 12
		Error_Check = CDbl(kousuiryou(j-1))
		If Err.Number = 13 And Error_Number = "0" then
			log = log & "Error:Rain Not a Number   " & CStr(j) & " = " & kousuiryou(j) & Chr(13)
			msgbox("�����ȊO�̕�������͂��Ȃ��ł�������" & Chr(13) & Chr(13) & kousuiryou_t)
			Error_Number = "kousuiryou_Nan"
			Exit for
		End If
		Err.Clear
	next

	'�X�y�[�X��؂蔻��
	If InStr(input_kousuiryou," ") = 0 And Error_Number = "0" then
		log = log & "Error:Rain String space not found" & Chr(13)
		Error_Number = "kousuiryou_no_space"
		msgbox("�~���ʂƍ~���ʂ̊Ԃ̓X�y�[�X�ŋ�؂��Ă��������B" & Chr(13) & Chr(13) & kousuiryou_t)
	End If
	
	If IsEmpty(kousuiryou(11)) then
		If Err.Number = 9  And Error_Number = "0" then
			log = log & "Error:Rain 12th Number is Empty"
			Error_Number = "kousuiryou<12"
			msgbox("�~���ʂ�12���͂���Ă��܂���B���肷��ɂ�1������12���܂ł̍~���ʂ��K�v�ł��B" & Chr(13) & Chr(13) & kousuiryou_t)
		End If
	End If
	
	If IsEmpty(kousuiryou(12)) then
	End If

	If Err.Number = 0  And Error_Number = "0" then
		log = log & "Warning:There is 13th kousuiryou number" & Chr(13)
		kousuiryou_conf_13 = msgbox("�~���ʂ�13�ȏ���͂���Ă��܂��B13�ڈȍ~�͌v�Z���疳������܂�����낵���ł��傤��?" & Chr(13) & Chr(13) & kousuiryou_t,vbYesNo)
		If (kousuiryou_conf_13) = 7 then
			Error_Number = "kousuiryou_13"
			log = log & "Returned to inputbox" & Chr(13)
		Else
			log = log & "Continued" & Chr(13)
		End If
	End If
	Err.Clear
	

	'�ُ�l����
	Err.Clear
	for m = 1 to 12
		If CDbl(kousuiryou(m-1)) < CDbl(-100) And Error_Number = "0" then
			If Err.Number = 0 then
				log = log & "Warning:Rain less than -100   " & CStr(m) & " = " & kousuiryou(m-1) & "��" & Chr(13)
				kousuiryou_conf = msgbox ("�~���ʂɋɒ[�ɏ������l������܂�:" & m & "��" & kousuiryou(m-1) & "�x" & Chr(13) & "���s���܂����H" & Chr(13) & Chr(13) & kousuiryou_t,vbYesNo)
				If kousuiryou_conf = 7 then
					Error_Number = "Kousuiryou_Low"
				End If
				If kousuiryou_conf = 6 then
					exit for
				End If
			End If
		End If

		If CDbl(kousuiryou(m-1)) > CDbl(100) And Error_Number = "0" then
			If Err.Number = 0 then
				log = log & "Warning:Rain more than 100   " & CStr(m) & " = " & kousuiryou(m-1) & "��" & Chr(13)
				kousuiryou_conf = msgbox ("�~���ʂɋɒ[�ɑ傫���l������܂�:" & m & "��" & kousuiryou(m-1) & "�x" & Chr(13) & "���s���܂����H" & Chr(13) & Chr(13) & kousuiryou_t,vbYesNo)
				If kousuiryou_conf = 7 then
					log = log & "Returned to inputbox" & Chr(13)
					Error_Number = "Kousuiryou_High"
				End If
				If kousuiryou_conf = 6 then
					log = log & "Continued" & Chr(13)
					Exit for
				End If
			End If
		End If
	next

	def_kousuiryou = input_kousuiryou

	If Error_Number = "0" And Err.Number = 0 then
		log = log & "Rain Error check passed" & Chr(13)
		Exit For
	ElseIf Error_Number = "0" then
		log = log & "Error:Rain Final Error Check Failed" & Chr(13) & "Err.Number = " & Err.Number & Chr(13)
	Else
		log = log & "Error:Rain Final Error Check Failed" & Chr(13) & "Error_Number = " & Error_Number & Chr(13)
	End If
next

log = log & "Input Finished" & Chr(13)
On Error goto 0

'10�x�ȏ�̌��̐������߂�
warm_months = CDbl(0.0)
for i = 1 to 12
	If CDbl(kion(i-1)) >= CDbl(10.0) then
		warm_months = warm_months + CDbl(1.0)
	End If
next
If warm_months <> CDbl(1) then
	dbl = "s"
Else
	dbl = ""
End If
'�N�ԍ~���ʂ����߂�
sum_kousuiryou = CDbl(0.0)
for i = 1 to 12
	sum_kousuiryou = sum_kousuiryou + CDbl(kion(i-1))
next

'�N���ϋC�������߂�
sum_kion = CDbl(0.0)
for i = 1 to 12
	sum_kion = sum_kion+CDbl(kion(i-1))
next
ave_kion = CDbl(sum_kion/12)

'�Ŋ����ƍŊ����C�������߂�
min_kion = CDbl(kion(0))
min_kion_month = Cstr(1)
for i = 1 to 11
	If min_kion > CDbl(kion(i)) then
		min_kion_month = Cstr((i+1))
		min_kion = CDbl(kion(i))
	End If
next

'�Œg���ƍŒg���C�������߂�
max_kion=CDbl(kion(0))
max_kion_month = Cstr(1)
for i = 1 to 11
	If max_kion < CDbl(kion(i)) then
		max_kion_month = Cstr((i+1))
		max_kion = CDbl(kion(i))
	End If
next

'�ő��J���ƍő��J���~���ʂ����߂�
max_kousuiryou=CDbl(kousuiryou(0))
max_kousuiryou_month = Cstr(1)
for i = 1 to 11
	If max_kousuiryou < CDbl(kousuiryou(i)) then
		max_kousuiryou_month = Cstr((i+1))
		max_kousuiryou = CDbl(kousuiryou(i))
	End If
next

'�ŏ��J���ƍŏ��J���~���ʂ����߂�
min_kousuiryou=CDbl(kousuiryou(0))
min_kousuiryou_month = Cstr(1)
for i = 1 to 11
	If min_kousuiryou > CDbl(kousuiryou(i)) then
		min_kousuiryou_month = Cstr((i+1))
		min_kousuiryou = CDbl(kousuiryou(i))
	End If
next
	

'���т��ǂ����̔���
If max_kion < CDbl(10.0) Then

	'����
	report = "�ő��J��:" & max_kousuiryou_month & "��    " & max_kousuiryou & "mm" & Chr(13) & "�ŏ��J��:" & min_kousuiryou_month & "��    "  & min_kousuiryou & "mm" & Chr(13) & "�Œg��:" & max_kion_month & "��    " & max_kion & "��" & Chr(13) & "�Ŋ���:" & min_kion_month & "��    " & min_kion & "��" & Chr(13) & "�N�ԍ~����:" & sum_kousuiryou & "mm" & Chr(13) & "�N���ϋC��:" & ave_kion & "��" & Chr(13) & "10���ȏ�̌�:" & warm_months & "����"
	
	log = log & Chr(13) & Chr(13) & "Max_Rain = " & max_kousuiryou_month & "    " & max_kousuiryou & "mm" & Chr(13) & "Min_Rain = " & min_kousuiryou_month & "    "  & min_kousuiryou & "mm" & Chr(13) & "Max_Kion = " & max_kion_month & "    " & max_kion & "��" & Chr(13) & "Min_Kion = " & min_kion_month & "    " & min_kion & "��" & Chr(13) & "Sum_Rain = " & sum_kousuiryou & "mm" & Chr(13) & "Ave_Kion = " & ave_kion & "��" & Chr(13) & "Warm_Months = " & warm_months & "month" & dbl
	
	kikoutai = "����"
	If max_kion < CDbl(0.0) Then
		kikouku = "�X��C��(EF)"
		kikouku_d = "�X��C��(EF)"
		kikouku_log = "EF"
	Else
		kikouku = "�c���h���C��(ET)"
		kikouku_d = "�c���h���C��(ET)"
		kikouku_log = "ET"
	End If
Else

	'���тł͂Ȃ�
	'�k�������씼�����̔��f
	If CDbl(kion(6))+CDbl(kion(7)) < CDbl(kion(0))+CDbl(kion(1)) Then


		'�씼��
		earth = "�씼��"
		earth_e = "South"
		summer(0) = "12"
		summer(1) = "1"
		summer(2) = "2"
		summer(3) = "3"
		winter(0) = "6"
		winter(1) = "7"
		winter(2) = "8"
		winter(3) = "9"
	Else

		'�k����
		earth = "�k����"
		earth_e = "North"
		summer(0) = "6"
		summer(1) = "7"
		summer(2) = "8"
		summer(3) = "9"
		winter(0) = "12"
		winter(1) = "1"
		winter(2) = "2"
		winter(3) = "3"
	End If


	'�����G�ߌ^�����߂�
	kansou = "full"
	If Setting(0) = 1 then
		'�����G�ߌ^�C������
		for i = 1 to 4
			If summer(i-1) = max_kousuiryou_month and (min_kousuiryou*10) < max_kousuiryou then
				kansou = "winter"
			ElseIf winter(i-1) = min_kousuiryou_month and (min_kousuiryou*10) < max_kousuiryou then
				kansou = "winter"
			End If

			If winter(i-1) = max_kousuiryou_month and min_kousuiryou*3 < max_kousuiryou then
				kansou = "summer"
			ElseIf summer(i-1) = min_kousuiryou_month and min_kousuiryou*3 < max_kousuiryou then
				kansou = "summer"
			End If
		next
	Else
		'�����G�ߌ^�C���Ȃ�
		for i = 1 to 4
			If summer(i-1) = max_kousuiryou_month and (min_kousuiryou*10) < max_kousuiryou then
				kansou = "winter"
			End If

			If winter(i-1) = max_kousuiryou_month and min_kousuiryou*3 < max_kousuiryou then
				kansou = "summer"
			End If
		next
	End If

	'�������E�����߂�
	Select Case kansou
		Case "full"
			'�N�������^
			kansou_genkai = CDbl(20*(ave_kion + CDbl(7.0)))
		Case "summer"
			'�Ċ������^
			kansou_genkai = CDbl(20*ave_kion)
		Case "winter"
			'�~�G�����^
			kansou_genkai = CDbl(20*(ave_kion + CDbl(14.0)))
		Case Else
			'�G���[
			msgbox("Debug Message" & Chr(13) & "Incorrect String $kansou")
	End Select

	Select Case kansou
		Case "full"
			kansou_jpn = "�N�������^"
		Case "summer"
			kansou_jpn = "��"
		Case "winter"
			kansou_jpn = "�~"
	End Select

	report = "�ő��J��:" & max_kousuiryou_month & "��    " & max_kousuiryou & "mm" & Chr(13) & "�ŏ��J��:" & min_kousuiryou_month & "��    "  & min_kousuiryou & "mm" & Chr(13) & "�Œg��:" & max_kion_month & "��    " & max_kion & "��" & Chr(13) & "�Ŋ���:" & min_kion_month & "��    " & min_kion & "��" & Chr(13) & "�N�ԍ~����:" & sum_kousuiryou & "mm" & Chr(13) & "�N���ϋC��:" & ave_kion & "��" & Chr(13) & "10���ȏ�̌�:" & warm_months & "����" & Chr(13) & "�n��:" & earth & Chr(13) & "���G:" & kansou_jpn & Chr(13) & "�������E:" & kansou_genkai & "mm"
	
	log = log & Chr(13) & "Max_Rain = " & max_kousuiryou_month & "    " & max_kousuiryou & "mm" & Chr(13) & "Min_Rain = " & min_kousuiryou_month & "    "  & min_kousuiryou & "mm" & Chr(13) & "Max_Kion = " & max_kion_month & "    " & max_kion & "��" & Chr(13) & "Min_Kion = " & min_kion_month & "    " & min_kion & "��" & Chr(13) & "Sum_Rain = " & sum_kousuiryou & "mm" & Chr(13) & "Ave_Kion = " & ave_kion & "��" & Chr(13) & "Warm_Month" & dbl & " = "& warm_months & "month" & dbl & Chr(13) & "Region = " & earth_e & Chr(13) & "Dry_Season = " & kansou & Chr(13) & "Drying_Limit = " & kansou_genkai & "mm"
	
'�����т̔���
	If sum_kousuiryou < kansougenkai then

		'������
		kikoutai = "������"
		If sum_kousuiryou < CDbl(kansougenkai/2) then
			kikouku = "�����C��(BW)"
			If ave_kion < CDbl(18.0) then
				kikouku_d = "�����C��(BWk)"
				kikouku_log = "BWk"
			Else
				kikouku_d = "�����C��(BWh)"
				kikouku_log = "BWh"
			End If
		Else
			kikouku = "�X�e�b�v�C��(BS)"
			If ave_kion < CDbl(18.0) then
				kikouku_d = "�X�e�b�v�C��(BSk)"
				kikouku_log = "BSk"
			Else
				kikouku_d = "�X�e�b�v�C��(BSh)"
				kikouku_log = "BSh"
			End If
		End If

	Else

		'�����тłȂ�
		If min_kion > CDbl(18.0) then
			kikoutai = "�M��"

			If min_kousuiryou > 60 then
						kikouku = "�M�щJ�ыC��(Af)"
						kikouku_d = "�M�щJ�ыC��(Af)"
			Else
				If min_kousuiryou < CDbl(100 - sum_kousuiryou*0.04) then
					If kansou = "summer" And min_kousuiryou < CDbl(30.0) then
						kikouku = "�T�o�i�C��(As)"
						kikouku_d = "�M�щċG���J�C��(As)"
						kikouku_log = "As"
					Else
						kikouku = "�T�o�i�C��(Aw)"
						kikouku_d = "�T�o�i�C��(Aw)"
						kikouki_log = "Aw"
					End If
				Else
					kikouku = "�T�o�i�C��(Am)"
					kikouku_d = "�M�у����X�[���C��(Am)"
					kikouku_log = "Am"
				End If
			End If
		ElseIf min_kion < CDbl(-3.0) then
			kikoutai = "������(���)"
			Select Case kansou
				Case "winter"
					kikouku = "�����ѓ~�G���J�C��(Dw)"
					If max_kion > CDbl(22.0) then
						kikouku_d = "�����ѓ~�G���J�C��(Dwa)"
						kikouku_log = "Dwa"
					ElseIf warm_months > CDbl(3.0) then
						kikouku_d = "�����ѓ~�G���J�C��(Dwb)" 
						kikouku_log = "Dwb"
					ElseIf min_kion > CDbl(-38.0) then
						kikouku_d = "�����ѓ~�G���J�C��(Dwc)"
						kikouku_log = "Dwc"
					Else
						kikouku_d = "�����ѓ~�G���J�C��(Dwd)"
						kikouku_log =  "Dwd"
					End If
				Case "summer"
					If min_kousuiryou < CDbl(30.0) then
						kikouku = "�Y���Ȃ�(���n�n���C���C��)"
						If max_kion > CDbl(22.0) then
							kikouku_d = "���n�n���C���C��(Dsa)"
							kikouku_log = "Dsa"
						ElseIf warm_months > CDbl(3.0) then
							kikouku_d = "���n�n���C���C��(Dsb)" 
							kikouku_log = "Dsb"
						ElseIf min_kion > CDbl(-38.0) then
							kikouku_d = "���n�n���C���C��(Dsc)"
							kikouku_log = "Dsc"
						Else
							kikouku_d = "���n�n���C���C��(Dsd)"
							kikouku_log = "Dsd"
						End If
					Else
						kikouku = "�����ю����C��(Df)"
						If max_kion > CDbl(22.0) then
							kikouku_d = "�����ю����C��(Dfa)"
							kikouku_log = "Dfa(From Ds)"
						ElseIf warm_months >= CDbl(4.0) then
							kikouku_d = "�����ю����C��(Dfb)" 
							kikouku_log = "Dfb(From Ds)"
						ElseIf min_kion > CDbl(-38.0) then
							kikouku_d = "�����ю����C��(Dfc)"
							kikouku_log = "Dfc(From Ds)"
						Else
							kikouku_d = "�����ю����C��(Dfd)"
							kikouku_log = "Dfd(From Ds)"
						End If
					End If
				Case "full"
					kikouku = "�����ю����C��(Df)"
					If max_kion > CDbl(22.0) then
						kikouku_d = "�����ю����C��(Dfa)"
						kikouku_log = "Dfa"
					ElseIf warm_months >= CDbl(4.0) then
						kikouku_d = "�����ю����C��(Dfb)" 
						kikouku_log = "Dfb"
					ElseIf min_kion > CDbl(-38.0) then
						kikouku_d = "�����ю����C��(Dfc)"
						kikouku_log = "Dfc"
					Else
						kikouku_d = "�����ю����C��(Dfd)"
						kikouku_log = "Dfd"
					End If
			End Select

		Else
			kikoutai = "����"
			Select Case kansou
				Case "winter"
					kikouku = "���g�~�G���J�C��(Cw)"
					If max_kion > CDbl(22.0) then
						kikouku_d = "���g�~�G���J�C��(Cwa)"
						kikouku_log = "Cwa"
					ElseIf warm_months >= CDbl(4.0) then
						kikouku_d = "���g�~�G���J�C��(Cwb)" 
						kikouku_log = "Cwb"
					Else
						kikouku_d = "���g�~�G���J�C��(Cwc)"
						kikouku_log = "Cwc"
					End If
				Case "summer"
					If min_kousuiryou < CDbl(30.0) then
						kikouku = "�n���C���C��(Cs)"
						If max_kion > CDbl(22.0) then
							kikouku_d = "�n���C���C��(Csa)"
							kikouku_log = "Csa"
						ElseIf warm_months >= CDbl(4.0) then
							kikouku_d = "�n���C���C��(Csb)"
							kikouku_log = "Csb"
						Else
							kikouku_d = "�n���C���C��(Csc)"
							kikouku_log = "Csc"
						End If
					Else
						If max_kion > CDbl(22.0) then
							kikouku = "���g�����C��(Cfa)"
							kikouku_d = "���g�����C��(Cfa)"
							kikouku_log = "Cfa(From Cs)"
						ElseIf warm_months >= CDbl(4.0) then
							kikouku = "���݊C�m���C��(Cfb)" 
							kikouku_d = "���݊C�m���C��(Cfb)" 
							kikouku_log = "Cfb(From Cs)"
						Else
							kikouku = "���݊C�m���C��(Cfc)"
							kikouku_d = "���݊C�m���C��(Cfc)"
							kikouku_log = "Cfc(From Cs)"
						End If
					End If
				Case "full"
					If max_kion > CDbl(22.0) then
						kikouku = "���g�����C��(Cfa)"
						kikouku_d = "���g�����C��(Cfa)"
						kikouku_log = "Cfa"
					ElseIf warm_months >= CDbl(4.0) then
						kikouku = "���݊C�m���C��(Cfb)" 
						kikouku_d = "���݊C�m���C��(Cfb)" 
						kikouku_log = "Cfb"
					Else
						kikouku = "���݊C�m���C��(Cfc)"
						kikouku_d = "���݊C�m���C��(Cfc)"
						kikouku_log = "Cfc"
					End If
			End Select

		End If
	End If
End If

message = "�C���:" & kikoutai & Chr(13) & "12�C���:" & kikouku & Chr(13) & "�ڍ׋C���:" & kikouku_d & Chr(13) & Chr(13) & "�ڍ�" & Chr(13) & report

log = log & Chr(13) & "Climatic_zone = " & kikouku_log
msgbox (message)


On Error Resume Next
Err.Clear
message = Replace(message,Chr(13),CRLF)
log = Replace(CStr(log),Chr(13),CRLF)
message = time_full & Chr(13) & Chr(10) & message
time_full = Date & "-" &  hour(now) & "-" & minute(now) & "-" & second(now)
time_full = Replace(time_full,"/","-")

'�L�^�̏�������
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objRec = objFSO.OpenTextFile("log.txt",8,True)
objRec.WriteLine("-------------------------------------" & CRLF & CRLF & time_full & CRLF & message & CRLF & CRLF & CRLF)
If Err.Number <> 0 then
	msgbox ("�L�^�t�@�C���̏������݂Ɏ��s���܂����B")
	log = log & "Cannot Write Record:" & Err.Number & " " & Err.Description & Chr(13)
	Err.Clear
End If
objRec.Close
Set objRec = Nothing

Err.Clear

Set objLog = objFSO.OpenTextFile("Keppen-v9.log",8,true)
objLog.WriteLine("---------------------------------------" & CRLF & time_full & CRLF & log & CRLF & CRLF & CRLF)
objLog.close
Set objLog = Nothing
If Err.Number <> 0 then
	msgbox ("���O�t�@�C���̏������݂Ɏ��s���܂����B")
	Err.Clear
End If






