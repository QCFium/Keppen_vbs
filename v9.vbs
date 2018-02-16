'宣言(都合上一部省略）
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

'配列の初期化
for i = 1 to 12
	kion_double(i-1) = CDbl(0)
	kousuiryou_double(i-1) = CDbl(0)
next

'FileSystemObject
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

CRLF = Chr(13) & Chr(10)
def_kion = "気温(℃)"
def_kousuiryou = "降水量(mm)"
mode = "Normal"
Setting(0) = "1"
log = "First Err.Number = " & Err.Number & Chr(13)
On Error Resume next
Error_check = CDbl(0.0)


for i = 1 to 1000000000

	Error_Number = "0"
	Err.Clear
	
	log = log & "-----Kion For:" & CStr(i) & Chr(13) & Chr(13)
	
	input_kion = inputbox("1月から12月までの気温をスペースで区切って入力してください。",mode,def_kion)

	log = log & "Kion Input:" & input_kion & Chr(13)
	'デバッグモード
	If input_kion = "mode debug" then
		mode = "Debug"
		log = log & "Mode changed to Debug" & Chr(13)
		Error_Number = "debug"
	End If

	'ノーマルモード
	If input_kion = "mode normal" then
		mode = "Normal"
		log = log & "Mode changed to Normal" & Chr(13)
		Error_Number = "Normal"
	End If

	'乾燥季節型修正Off
	If input_kion = "Setting 0 Off" then
		Setting(0) = "0"
		log = log & "Setting 0 changed to Off" & Chr(13)
		Error_Number = "Setting 0 Off"
	End If

	'乾燥季節型修正On
	If input_kion = "Setting 0 On" then
		Setting(0) = "1"
		log = log & "Setting 0 changed to On" & Chr(13)
		Error_Number = "Setting 0 On"
	End If



	'キャンセル判定
	If IsEmpty(input_kion) = true then
		Error_Number = "cancell"
		log = log & "Cancell Pressed" & Chr(13)
		quit_conf = msgbox ("本当に終了しますか? 入力したデータは失われます。",vbYesNo,mode)
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

	'両端の空白を削除
	input_kion = RTrim(input_kion)
	input_kion = LTrim(input_kion)
	'全角スペースを半角に置換
	input_kion = Replace (input_kion,"　"," ")
	'                                 ↑これは全角スペース
	'2個以上の連続した半角スペースをひとつにする
	for k = 1 to 100
		input_kion = Replace(input_kion,"  "," ")
	'                                        ↑これは半角スペース×2
	next


	'空白OK判定
	If input_kion = "" And Error_Number = "0" then
		log = log & "Error:Kion Input is Empty" & Chr(13)
		Error_Number = "kion_empty"
		msgbox ("気温を入力してください")
	End If

	kion = split(input_kion)
	log = log & "Kion Split Err.Number = " & CStr(Err.Number) & Chr(13)
	Err.Clear

	for n = 1 to 1000000
		kion_t = kion_t & n & "月:" & kion(n-1) & Chr(13)
		If Err.Number = 9 then
			log = log & "End making kion_t:" & n & Chr(13)
			exit for
		ElseIf Err.Number <> 0 then
			msgbox("予期せぬエラーが発生したため終了します。エラー番号をお伝えください" & Chr(13) & "エラー番号:" & Err.Number)
			WScript.Quit
		End If
	next
	Err.Clear

	'数字判定
	for j = 1 to 12
		Error_Check = CDbl(kion(j-1))
		If Err.Number = 13 And Error_Number = "0" then
			log = log & "Error:Kion Not a number   " & CStr(j) & " = " & kion(j-1) & Chr(13)
			msgbox ("数字以外の文字を入力しないでください" & Chr(13) & Chr(13) & kion_t)
			Error_Number = "kion_Nan"
			exit for
		End If
		Err.Clear
	next

	'スペース区切り判定
	If InStr(input_kion," ") = 0 And Error_Number = "0" then
		log = log & "Error:Kion String space not found" & Chr(13)
		Error_Number = "kion_no_space"
		msgbox ("気温と気温の間はスペースで区切ってください。" & Chr(13) & Chr(13) & kion_t)
	End If
	
	If IsEmpty(kion(11)) then
		If Err.Number = 9  And Error_Number = "0" then
			Error_Number = "kion<12"
			log = log & "Error:Kion 12th Number is Empty" & Chr(13)
			msgbox ("気温が12個入力されていません。判定するには1月から12月までの気温が必要です。" & Chr(13) & Chr(13) & kion_t)
			Err.Clear
		End If
	End If

	If IsEmpty(kion(12)) then
	End If

	If Err.Number = 0  And Error_Number = "0" then
		log = log & "Warning:There is 13th kion number" & Chr(13)
		kion_conf_13 = msgbox("気温が13個以上入力されています。13個目以降は計算から無視されますがよろしいでしょうか?" & Chr(13) & Chr(13) & kion_t,vbYesNo)
		If (kion_conf_13) = 7 then
			Error_Number = "kion_13"
			log = log & "Returned to inputbox" & Chr(13)
		Else
			log = log & "Continued" & Chr(13)
		End If
	End If
	Err.Clear

	'異常値判定
	Err.Clear
	for m = 1 to 12
		If CDbl(kion(m-1)) < CDbl(-100) And Error_Number = "0" then
			If Err.Number = 0 then
				log = log & "Warning:Kion less than -100   " & CStr(m) & " = " & kion(m-1) & "℃" & Chr(13)
				kion_conf = msgbox("気温に極端に小さい値があります:" & m & "月" & kion(m-1) & "度" & Chr(13) & "続行しますか？" & Chr(13) & Chr(13) & kion_t,vbYesNo)
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
				log = log & "Warning:Kion more than 100   " & CStr(m) & " = " & kion(m-1) & "℃" & Chr(13)
				kion_conf = msgbox("気温に極端に大きい値があります:" & m & "月" & kion(m-1) & "度" & Chr(13) & "続行しますか？" & Chr(13) & Chr(13) & kion_t,vbYesNo)
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
	input_kousuiryou = inputbox("1月から12月までの降水量をスペースで区切って入力してください。",mode,def_kousuiryou)
	log = log & "Rain Input:" & input_kousuiryou & Chr(13)

	'デバッグモード
	If input_kousuiryou = "mode debug" then
		mode = "Debug"
		log = log & "Mode changed to Debug" & Chr(13)
		Error_Number = "debug"
	End If

	'ノーマルモード
	If input_kousuiryou = "mode normal" then
		mode = "Normal"
		log = log & "Mode changed to Normal" & Chr(13)
		Error_Number = "Normal"
	End If

	'乾燥季節型修正Off
	If input_kion = "Setting 0 Off" then
		Setting(0) = "0"
		log = log & "Setting 0 changed to Off" & Chr(13)
		Error_Number = "Setting 0 Off"
	End If

	'乾燥季節型修正On
	If input_kion = "Setting 0 On" then
		Setting(0) = "1"
		log = log & "Setting 0 changed to On" & Chr(13)
		Error_Number = "Setting 0 On"
	End If


	'キャンセル判定
	If IsEmpty(input_kousuiryou) = true then
		Error_Number = "cancell"
		log = log & "Cancell Pressed"
		quit_conf = msgbox ("本当に終了しますか? 入力したデータは失われます。",vbYesNo,mode)
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
	input_kousuiryou = Replace (input_kousuiryou,"　"," ")
	for k = 1 to 100
		input_kousuiryou = Replace(input_kousuiryou,"  "," ")
	next

	'空かどうか判定
	If input_kousuiryou = "" And Error_Number = "0" then
		Error_Number = "kousuiryou_empty"
		log = log & "Error:Rain input is Empty" & Chr(13)
		msgbox ("降水量を入力してください")
	End If

	kousuiryou = split(input_kousuiryou)
	log = log & "Rain Split Err.Number = " & CStr(Err.Number) & Chr(13)
	Err.Clear

	for n = 1 to 1000000
		kousuiryou_t = kousuiryou_t & n & "月:" & kousuiryou(n-1) & Chr(13)
		If Err.Number = 9 then
			log = log & "End making kousuiryou_t:" & n & Chr(13)
			exit for
		ElseIf Err.Number <> 0 then
			msgbox("予期せぬエラーが発生したため終了します。エラー番号をお伝えください" & Chr(13) & "エラー番号:" & Err.Number)
			WScript.Quit
		End If
	next
	Err.Clear

	'数字判定
	for j = 1 to 12
		Error_Check = CDbl(kousuiryou(j-1))
		If Err.Number = 13 And Error_Number = "0" then
			log = log & "Error:Rain Not a Number   " & CStr(j) & " = " & kousuiryou(j) & Chr(13)
			msgbox("数字以外の文字を入力しないでください" & Chr(13) & Chr(13) & kousuiryou_t)
			Error_Number = "kousuiryou_Nan"
			Exit for
		End If
		Err.Clear
	next

	'スペース区切り判定
	If InStr(input_kousuiryou," ") = 0 And Error_Number = "0" then
		log = log & "Error:Rain String space not found" & Chr(13)
		Error_Number = "kousuiryou_no_space"
		msgbox("降水量と降水量の間はスペースで区切ってください。" & Chr(13) & Chr(13) & kousuiryou_t)
	End If
	
	If IsEmpty(kousuiryou(11)) then
		If Err.Number = 9  And Error_Number = "0" then
			log = log & "Error:Rain 12th Number is Empty"
			Error_Number = "kousuiryou<12"
			msgbox("降水量が12個入力されていません。判定するには1月から12月までの降水量が必要です。" & Chr(13) & Chr(13) & kousuiryou_t)
		End If
	End If
	
	If IsEmpty(kousuiryou(12)) then
	End If

	If Err.Number = 0  And Error_Number = "0" then
		log = log & "Warning:There is 13th kousuiryou number" & Chr(13)
		kousuiryou_conf_13 = msgbox("降水量が13個以上入力されています。13個目以降は計算から無視されますがよろしいでしょうか?" & Chr(13) & Chr(13) & kousuiryou_t,vbYesNo)
		If (kousuiryou_conf_13) = 7 then
			Error_Number = "kousuiryou_13"
			log = log & "Returned to inputbox" & Chr(13)
		Else
			log = log & "Continued" & Chr(13)
		End If
	End If
	Err.Clear
	

	'異常値判定
	Err.Clear
	for m = 1 to 12
		If CDbl(kousuiryou(m-1)) < CDbl(-100) And Error_Number = "0" then
			If Err.Number = 0 then
				log = log & "Warning:Rain less than -100   " & CStr(m) & " = " & kousuiryou(m-1) & "℃" & Chr(13)
				kousuiryou_conf = msgbox ("降水量に極端に小さい値があります:" & m & "月" & kousuiryou(m-1) & "度" & Chr(13) & "続行しますか？" & Chr(13) & Chr(13) & kousuiryou_t,vbYesNo)
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
				log = log & "Warning:Rain more than 100   " & CStr(m) & " = " & kousuiryou(m-1) & "℃" & Chr(13)
				kousuiryou_conf = msgbox ("降水量に極端に大きい値があります:" & m & "月" & kousuiryou(m-1) & "度" & Chr(13) & "続行しますか？" & Chr(13) & Chr(13) & kousuiryou_t,vbYesNo)
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

'10度以上の月の数を求める
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
'年間降水量を求める
sum_kousuiryou = CDbl(0.0)
for i = 1 to 12
	sum_kousuiryou = sum_kousuiryou + CDbl(kion(i-1))
next

'年平均気温を求める
sum_kion = CDbl(0.0)
for i = 1 to 12
	sum_kion = sum_kion+CDbl(kion(i-1))
next
ave_kion = CDbl(sum_kion/12)

'最寒月と最寒月気温を求める
min_kion = CDbl(kion(0))
min_kion_month = Cstr(1)
for i = 1 to 11
	If min_kion > CDbl(kion(i)) then
		min_kion_month = Cstr((i+1))
		min_kion = CDbl(kion(i))
	End If
next

'最暖月と最暖月気温を求める
max_kion=CDbl(kion(0))
max_kion_month = Cstr(1)
for i = 1 to 11
	If max_kion < CDbl(kion(i)) then
		max_kion_month = Cstr((i+1))
		max_kion = CDbl(kion(i))
	End If
next

'最多雨月と最多雨月降水量を求める
max_kousuiryou=CDbl(kousuiryou(0))
max_kousuiryou_month = Cstr(1)
for i = 1 to 11
	If max_kousuiryou < CDbl(kousuiryou(i)) then
		max_kousuiryou_month = Cstr((i+1))
		max_kousuiryou = CDbl(kousuiryou(i))
	End If
next

'最少雨月と最少雨月降水量を求める
min_kousuiryou=CDbl(kousuiryou(0))
min_kousuiryou_month = Cstr(1)
for i = 1 to 11
	If min_kousuiryou > CDbl(kousuiryou(i)) then
		min_kousuiryou_month = Cstr((i+1))
		min_kousuiryou = CDbl(kousuiryou(i))
	End If
next
	

'寒帯かどうかの判定
If max_kion < CDbl(10.0) Then

	'寒帯
	report = "最多雨月:" & max_kousuiryou_month & "月    " & max_kousuiryou & "mm" & Chr(13) & "最少雨月:" & min_kousuiryou_month & "月    "  & min_kousuiryou & "mm" & Chr(13) & "最暖月:" & max_kion_month & "月    " & max_kion & "℃" & Chr(13) & "最寒月:" & min_kion_month & "月    " & min_kion & "℃" & Chr(13) & "年間降水量:" & sum_kousuiryou & "mm" & Chr(13) & "年平均気温:" & ave_kion & "℃" & Chr(13) & "10℃以上の月:" & warm_months & "ヶ月"
	
	log = log & Chr(13) & Chr(13) & "Max_Rain = " & max_kousuiryou_month & "    " & max_kousuiryou & "mm" & Chr(13) & "Min_Rain = " & min_kousuiryou_month & "    "  & min_kousuiryou & "mm" & Chr(13) & "Max_Kion = " & max_kion_month & "    " & max_kion & "℃" & Chr(13) & "Min_Kion = " & min_kion_month & "    " & min_kion & "℃" & Chr(13) & "Sum_Rain = " & sum_kousuiryou & "mm" & Chr(13) & "Ave_Kion = " & ave_kion & "℃" & Chr(13) & "Warm_Months = " & warm_months & "month" & dbl
	
	kikoutai = "寒帯"
	If max_kion < CDbl(0.0) Then
		kikouku = "氷雪気候(EF)"
		kikouku_d = "氷雪気候(EF)"
		kikouku_log = "EF"
	Else
		kikouku = "ツンドラ気候(ET)"
		kikouku_d = "ツンドラ気候(ET)"
		kikouku_log = "ET"
	End If
Else

	'寒帯ではない
	'北半球か南半球かの判断
	If CDbl(kion(6))+CDbl(kion(7)) < CDbl(kion(0))+CDbl(kion(1)) Then


		'南半球
		earth = "南半球"
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

		'北半球
		earth = "北半球"
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


	'乾燥季節型を求める
	kansou = "full"
	If Setting(0) = 1 then
		'乾燥季節型修正あり
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
		'乾燥季節型修正なし
		for i = 1 to 4
			If summer(i-1) = max_kousuiryou_month and (min_kousuiryou*10) < max_kousuiryou then
				kansou = "winter"
			End If

			If winter(i-1) = max_kousuiryou_month and min_kousuiryou*3 < max_kousuiryou then
				kansou = "summer"
			End If
		next
	End If

	'乾燥限界を求める
	Select Case kansou
		Case "full"
			'年中湿潤型
			kansou_genkai = CDbl(20*(ave_kion + CDbl(7.0)))
		Case "summer"
			'夏期乾燥型
			kansou_genkai = CDbl(20*ave_kion)
		Case "winter"
			'冬季乾燥型
			kansou_genkai = CDbl(20*(ave_kion + CDbl(14.0)))
		Case Else
			'エラー
			msgbox("Debug Message" & Chr(13) & "Incorrect String $kansou")
	End Select

	Select Case kansou
		Case "full"
			kansou_jpn = "年中湿潤型"
		Case "summer"
			kansou_jpn = "夏"
		Case "winter"
			kansou_jpn = "冬"
	End Select

	report = "最多雨月:" & max_kousuiryou_month & "月    " & max_kousuiryou & "mm" & Chr(13) & "最小雨月:" & min_kousuiryou_month & "月    "  & min_kousuiryou & "mm" & Chr(13) & "最暖月:" & max_kion_month & "月    " & max_kion & "℃" & Chr(13) & "最寒月:" & min_kion_month & "月    " & min_kion & "℃" & Chr(13) & "年間降水量:" & sum_kousuiryou & "mm" & Chr(13) & "年平均気温:" & ave_kion & "℃" & Chr(13) & "10℃以上の月:" & warm_months & "ヶ月" & Chr(13) & "地域:" & earth & Chr(13) & "乾季:" & kansou_jpn & Chr(13) & "乾燥限界:" & kansou_genkai & "mm"
	
	log = log & Chr(13) & "Max_Rain = " & max_kousuiryou_month & "    " & max_kousuiryou & "mm" & Chr(13) & "Min_Rain = " & min_kousuiryou_month & "    "  & min_kousuiryou & "mm" & Chr(13) & "Max_Kion = " & max_kion_month & "    " & max_kion & "℃" & Chr(13) & "Min_Kion = " & min_kion_month & "    " & min_kion & "℃" & Chr(13) & "Sum_Rain = " & sum_kousuiryou & "mm" & Chr(13) & "Ave_Kion = " & ave_kion & "℃" & Chr(13) & "Warm_Month" & dbl & " = "& warm_months & "month" & dbl & Chr(13) & "Region = " & earth_e & Chr(13) & "Dry_Season = " & kansou & Chr(13) & "Drying_Limit = " & kansou_genkai & "mm"
	
'乾燥帯の判定
	If sum_kousuiryou < kansougenkai then

		'乾燥帯
		kikoutai = "乾燥帯"
		If sum_kousuiryou < CDbl(kansougenkai/2) then
			kikouku = "砂漠気候(BW)"
			If ave_kion < CDbl(18.0) then
				kikouku_d = "砂漠気候(BWk)"
				kikouku_log = "BWk"
			Else
				kikouku_d = "砂漠気候(BWh)"
				kikouku_log = "BWh"
			End If
		Else
			kikouku = "ステップ気候(BS)"
			If ave_kion < CDbl(18.0) then
				kikouku_d = "ステップ気候(BSk)"
				kikouku_log = "BSk"
			Else
				kikouku_d = "ステップ気候(BSh)"
				kikouku_log = "BSh"
			End If
		End If

	Else

		'乾燥帯でない
		If min_kion > CDbl(18.0) then
			kikoutai = "熱帯"

			If min_kousuiryou > 60 then
						kikouku = "熱帯雨林気候(Af)"
						kikouku_d = "熱帯雨林気候(Af)"
			Else
				If min_kousuiryou < CDbl(100 - sum_kousuiryou*0.04) then
					If kansou = "summer" And min_kousuiryou < CDbl(30.0) then
						kikouku = "サバナ気候(As)"
						kikouku_d = "熱帯夏季少雨気候(As)"
						kikouku_log = "As"
					Else
						kikouku = "サバナ気候(Aw)"
						kikouku_d = "サバナ気候(Aw)"
						kikouki_log = "Aw"
					End If
				Else
					kikouku = "サバナ気候(Am)"
					kikouku_d = "熱帯モンスーン気候(Am)"
					kikouku_log = "Am"
				End If
			End If
		ElseIf min_kion < CDbl(-3.0) then
			kikoutai = "亜寒帯(冷帯)"
			Select Case kansou
				Case "winter"
					kikouku = "亜寒帯冬季少雨気候(Dw)"
					If max_kion > CDbl(22.0) then
						kikouku_d = "亜寒帯冬季少雨気候(Dwa)"
						kikouku_log = "Dwa"
					ElseIf warm_months > CDbl(3.0) then
						kikouku_d = "亜寒帯冬季少雨気候(Dwb)" 
						kikouku_log = "Dwb"
					ElseIf min_kion > CDbl(-38.0) then
						kikouku_d = "亜寒帯冬季少雨気候(Dwc)"
						kikouku_log = "Dwc"
					Else
						kikouku_d = "亜寒帯冬季少雨気候(Dwd)"
						kikouku_log =  "Dwd"
					End If
				Case "summer"
					If min_kousuiryou < CDbl(30.0) then
						kikouku = "該当なし(高地地中海性気候)"
						If max_kion > CDbl(22.0) then
							kikouku_d = "高地地中海性気候(Dsa)"
							kikouku_log = "Dsa"
						ElseIf warm_months > CDbl(3.0) then
							kikouku_d = "高地地中海性気候(Dsb)" 
							kikouku_log = "Dsb"
						ElseIf min_kion > CDbl(-38.0) then
							kikouku_d = "高地地中海性気候(Dsc)"
							kikouku_log = "Dsc"
						Else
							kikouku_d = "高地地中海性気候(Dsd)"
							kikouku_log = "Dsd"
						End If
					Else
						kikouku = "亜寒帯湿潤気候(Df)"
						If max_kion > CDbl(22.0) then
							kikouku_d = "亜寒帯湿潤気候(Dfa)"
							kikouku_log = "Dfa(From Ds)"
						ElseIf warm_months >= CDbl(4.0) then
							kikouku_d = "亜寒帯湿潤気候(Dfb)" 
							kikouku_log = "Dfb(From Ds)"
						ElseIf min_kion > CDbl(-38.0) then
							kikouku_d = "亜寒帯湿潤気候(Dfc)"
							kikouku_log = "Dfc(From Ds)"
						Else
							kikouku_d = "亜寒帯湿潤気候(Dfd)"
							kikouku_log = "Dfd(From Ds)"
						End If
					End If
				Case "full"
					kikouku = "亜寒帯湿潤気候(Df)"
					If max_kion > CDbl(22.0) then
						kikouku_d = "亜寒帯湿潤気候(Dfa)"
						kikouku_log = "Dfa"
					ElseIf warm_months >= CDbl(4.0) then
						kikouku_d = "亜寒帯湿潤気候(Dfb)" 
						kikouku_log = "Dfb"
					ElseIf min_kion > CDbl(-38.0) then
						kikouku_d = "亜寒帯湿潤気候(Dfc)"
						kikouku_log = "Dfc"
					Else
						kikouku_d = "亜寒帯湿潤気候(Dfd)"
						kikouku_log = "Dfd"
					End If
			End Select

		Else
			kikoutai = "温帯"
			Select Case kansou
				Case "winter"
					kikouku = "温暖冬季少雨気候(Cw)"
					If max_kion > CDbl(22.0) then
						kikouku_d = "温暖冬季少雨気候(Cwa)"
						kikouku_log = "Cwa"
					ElseIf warm_months >= CDbl(4.0) then
						kikouku_d = "温暖冬季少雨気候(Cwb)" 
						kikouku_log = "Cwb"
					Else
						kikouku_d = "温暖冬季少雨気候(Cwc)"
						kikouku_log = "Cwc"
					End If
				Case "summer"
					If min_kousuiryou < CDbl(30.0) then
						kikouku = "地中海性気候(Cs)"
						If max_kion > CDbl(22.0) then
							kikouku_d = "地中海性気候(Csa)"
							kikouku_log = "Csa"
						ElseIf warm_months >= CDbl(4.0) then
							kikouku_d = "地中海性気候(Csb)"
							kikouku_log = "Csb"
						Else
							kikouku_d = "地中海性気候(Csc)"
							kikouku_log = "Csc"
						End If
					Else
						If max_kion > CDbl(22.0) then
							kikouku = "温暖湿潤気候(Cfa)"
							kikouku_d = "温暖湿潤気候(Cfa)"
							kikouku_log = "Cfa(From Cs)"
						ElseIf warm_months >= CDbl(4.0) then
							kikouku = "西岸海洋性気候(Cfb)" 
							kikouku_d = "西岸海洋性気候(Cfb)" 
							kikouku_log = "Cfb(From Cs)"
						Else
							kikouku = "西岸海洋性気候(Cfc)"
							kikouku_d = "西岸海洋性気候(Cfc)"
							kikouku_log = "Cfc(From Cs)"
						End If
					End If
				Case "full"
					If max_kion > CDbl(22.0) then
						kikouku = "温暖湿潤気候(Cfa)"
						kikouku_d = "温暖湿潤気候(Cfa)"
						kikouku_log = "Cfa"
					ElseIf warm_months >= CDbl(4.0) then
						kikouku = "西岸海洋性気候(Cfb)" 
						kikouku_d = "西岸海洋性気候(Cfb)" 
						kikouku_log = "Cfb"
					Else
						kikouku = "西岸海洋性気候(Cfc)"
						kikouku_d = "西岸海洋性気候(Cfc)"
						kikouku_log = "Cfc"
					End If
			End Select

		End If
	End If
End If

message = "気候帯:" & kikoutai & Chr(13) & "12気候区:" & kikouku & Chr(13) & "詳細気候区:" & kikouku_d & Chr(13) & Chr(13) & "詳細" & Chr(13) & report

log = log & Chr(13) & "Climatic_zone = " & kikouku_log
msgbox (message)


On Error Resume Next
Err.Clear
message = Replace(message,Chr(13),CRLF)
log = Replace(CStr(log),Chr(13),CRLF)
message = time_full & Chr(13) & Chr(10) & message
time_full = Date & "-" &  hour(now) & "-" & minute(now) & "-" & second(now)
time_full = Replace(time_full,"/","-")

'記録の書き込み
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objRec = objFSO.OpenTextFile("log.txt",8,True)
objRec.WriteLine("-------------------------------------" & CRLF & CRLF & time_full & CRLF & message & CRLF & CRLF & CRLF)
If Err.Number <> 0 then
	msgbox ("記録ファイルの書き込みに失敗しました。")
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
	msgbox ("ログファイルの書き込みに失敗しました。")
	Err.Clear
End If






