#include <MsgBoxConstants.au3>
#Include <Misc.au3>
#include <Array.au3>
#include <Excel.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include <ColorConstants.au3>


AutoItSetOption('WinTitleMatchMode',2)
HotKeySet("{ESC}", "stop")
$small = 100
$medium = 700
$large = 3000
Global $x = 0
Global $y = 0
Global $n = 0
Global $m = 0
Global $s1 = 115
$titel = "DCAP TAST ver. 1.2 (beta)"
$excel = "\Dataark_v1.2.xlsm"
$text = ""
Global $h = 80
Global $hplus = 25

Screen()
GUI1()

While 1
   Switch GUIGetMsg()
		 Case $GUI_EVENT_CLOSE
			Exit
		 Case $BUT1
			SplashOn()
			Opstart()
			If $opstart = 1 Then
			   Excel()
			   FirstSetup()
			   BE()
			   UA()
			   YD()
			   KV()
			   TA()
			   TT()
			   AG()
			   AL()
			   UT()
			   DOR()
			   VI()
			   BT()
			   IT()
			   IA()
			   TN()
			   EL()
			   GL()
			   BV()
			   VA()
			   VE()
			   MsgBox($MB_SYSTEMMODAL, "Done", "Programmet er færdi med at taste", 5)
			EndIf
			SplashOfff()
   EndSwitch
WEnd


;----------------------------------------------------------------------------------------------------------------------------------------------------------
;           OPSTART KONTROL           OPSTART KONTROL           OPSTART KONTROL           OPSTART KONTROL           OPSTART KONTROL
;----------------------------------------------------------------------------------------------------------------------------------------------------------
Func Opstart()
   Global $text = "Kontrolere om DCAP er opsat korrekt"
   SplashU()

   Global $opstart = 0
   WinActivate("Chrome", "")
   Sleep(300)
   Send("{PGUP}")
   Sleep(1000)
   Global $seach = PixelSearch ( 1, 0, 2, 600, 0xF1F1F1)
   If Not @error Then
	  Global $opstart = 1
   Else
	  SplashOfff()
	  MsgBox($MB_SYSTEMMODAL, "", "Programmet kunne ikke finde DCAP.")
	  Global $opstart = 0
   EndIf
EndFunc

;----------------------------------------------------------------------------------------------------------------------------------------------------------
;           FirstSetup           FirstSetup           FirstSetup           FirstSetup           FirstSetup           FirstSetup           FirstSetup
;----------------------------------------------------------------------------------------------------------------------------------------------------------
Func FirstSetup()
   Global $text = "Programmet taster, du kan bruge ''ESC'' til at stoppe programmet."
   SplashU()

   WinActivate("Chrome", "")
   Sleep(300)
   MouseMove($seach[0], $seach[1]+50, 2)
   Sleep(100)
   MouseClick("LEFT")
   Sleep(100)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
EndFunc

;----------------------------------------------------------------------------------------------------------------------------------------------------------
;           NESTE           NESTE           NESTE           NESTE           NESTE           NESTE           NESTE           NESTE           NESTE
;----------------------------------------------------------------------------------------------------------------------------------------------------------
Func Neste()
   Sleep($small)
   Send("{DOWN}")
   Sleep($small)
   Send("{DOWN}")
   Sleep($small)
   Send("{DOWN}")
   Sleep($small)
   Send("{DOWN}")
   Sleep($small)
   Send("{DOWN}")
   Sleep($small)
   Send("{DOWN}")
   Sleep($small)
   Send("{DOWN}")
   Sleep($medium)
   Global $seach = PixelSearch ( 900, 200, 901, 1080, 0xBACFD5)
   If Not @error Then
	  MouseMove($seach[0], $seach[1]+30, 2)
	  Sleep(100)
	  MouseClick("LEFT")
	  Sleep($medium)
	  DCAP_Arbejder_Check()

   Else
	  MsgBox($MB_SYSTEMMODAL, "", "Programmet kan ikke finde den næste bygningsdel")
	  Exit
   EndIf

EndFunc

;----------------------------------------------------------------------------------------------------------------------------------------------------------
;           DCAP_Arbejder_Check           DCAP_Arbejder_Check           DCAP_Arbejder_Check           DCAP_Arbejder_Check           DCAP_Arbejder_Check
;----------------------------------------------------------------------------------------------------------------------------------------------------------
Func DCAP_Arbejder_Check()
   $timer_mini = 0
   $timer_sec = 0

   Do
	  $x = 1
	  $seach = PixelSearch ( 900, 300, 901, 700, 0xB0B0B0)
	  If Not @error Then
		 $x = 0
	  EndIf

	  $seach = PixelSearch ( 900, 300, 901, 700, 0x808F93)
	  If Not @error Then
		 $x = 0
	  EndIf

	  $seach = PixelSearch ( 900, 300, 901, 700, 0xE1E1E1)
	  If Not @error Then
		 $x = 0
	  EndIf

	  Sleep(200)
	  $timer_mini = $timer_mini + 0.2001

	  If $timer_mini > 1 Then
		 $timer_mini = 0
		 $timer_sec = $timer_sec + 1
	  EndIf

	  If $timer_sec > 20 Then
		 Global $text = "DCAP har ''arbejdet'' i mere end 20 sek. Hvis det ikke løser sig skal hjemmeside og DCAP_Hacks genstartes."
		 SplashU()
	  ElseIf $timer_sec > 10 Then
		 Global $text = "DCAP har ''arbejdet'' i mere end 10 sek. Vent lidt endnu."
		 SplashU()
	  ElseIf $timer_sec > 5 Then
		 Global $text = "DCAP har ''arbejdet'' i 5 sek. Vent lidt endnu."
		 SplashU()
	  ElseIf $timer_sec > 2 Then
		 Global $text = "DCAP har ''arbejdet'' i 2 sek."
		 SplashU()
	  EndIf
   Until $x = 1
EndFunc
;----------------------------------------------------------------------------------------------------------------------------------------------------------
;           UDFYLDT_CHECK           UDFYLDT_CHECK           UDFYLDT_CHECK           UDFYLDT_CHECK           UDFYLDT_CHECK           UDFYLDT_CHECK
;----------------------------------------------------------------------------------------------------------------------------------------------------------

Func Udfyldt_check()
   $y = 0
   $z = 0

   Do
	  $x = 0
	  $seach = PixelSearch ( 768, 200, 769, 1080, 0xFF4040)
	  If Not @error Then
		 $y = 1
		 $z = 410
		 MsgBox($MB_SYSTEMMODAL, "", "Der er allerede tastet arbejder ind i denne bygningsdel. Sagens skal være helt tom for at DCAP_Hacks virker. Du kan enten stoppe indtastningen ved at trykke ''OK'' og så ''ESC'', eller du kan trykke ''OK'' og slette de indtastede arbejder + vedligehold i bygningsdelen (programmet fortsætter automatisk 10 sek. efter du har trykket ''OK''.")
		 Sleep(10000)
	  Else
		 $x = 1
	  EndIf
   Until $x = 1

   Do
	  $x = 0
	  $seach = PixelSearch ( 652, 200, 653, 1080, 0xFF4040)
	  If Not @error Then
		 $y = 1
		 $z = 230
		 MsgBox($MB_SYSTEMMODAL, "", "Der er allerede tastet arbejder ind i denne bygningsdel. Sagens skal være helt tom for at DCAP_Hacks virker. Du kan enten stoppe indtastningen ved at trykke ''OK'' og så ''ESC'', eller du kan trykke ''OK'' og slette de indtastede arbejder + vedligehold i bygningsdelen (programmet fortsætter automatisk 10 sek. efter du har trykket ''OK''.")
		 Sleep(10000)
	  Else
		 $x = 1
	  EndIf
   Until $x = 1

   If $y = 1 Then
	  $seach = PixelSearch ( $z, 200, $z+1, 1080, 0x2D4550)
	  If Not @error Then
		 MouseMove($seach[0], $seach[1]+30, 2)
		 Sleep(100)
		 MouseClick("LEFT")
		 Sleep($medium)
		 MouseClick("LEFT")
		 Sleep($medium)
		 Global $text = "Programmet taster, du kan bruge ''ESC'' til at stoppe programmet."
		 SplashU()
	  Else
		 MsgBox($MB_SYSTEMMODAL, "", "Programmet kunne ikke "
	  EndIf
   EndIf
EndFunc

;----------------------------------------------------------------------------------------------------------------------------------------------------------
;          EXCEL         EXCEL         EXCEL         EXCEL         EXCEL         EXCEL         EXCEL         EXCEL         EXCEL         EXCEL
;----------------------------------------------------------------------------------------------------------------------------------------------------------
Func Excel()

Global $text = "Indlæser data fra Excel"
SplashU()


Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel", "Error - EXCEL ER IKKE ÅBENT " & @CRLF & "@error = " & @error & ", @extended = " & @extended)
$oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & $excel)
If @error Then
	MsgBox($MB_SYSTEMMODAL,"Excel", "Error - DATAARK KAN IKKE FINDES, KONTROLLER AT ''DATAARK.XLSM'' LIGGER I SAMME MAPPE SOM PROGRAMMET")
	_Excel_Close($oExcel)
	Exit
EndIf

;STAMDATA (SD)
Global $SD = _Excel_RangeRead($oWorkbook, Default, "C5:C15")

;BELÆGNING (BE)
Global $BE1 = _Excel_RangeRead($oWorkbook, Default, "B53:O56")
Global $BE_N = 4

;UDVENDIGT AFLØB (UA)
Global $UA1 = _Excel_RangeRead($oWorkbook, Default, "B67:O70")
Global $UA_N = 4

;YDERVÆG (YD)
Global $YD1 = _Excel_RangeRead($oWorkbook, Default, "B85:O91")
Global $YD_N = 7

;KVISTE (KV)
Global $KV1 = _Excel_RangeRead($oWorkbook, Default, "B100:O101")
Global $KV_N = 2

;TAGDÆKNING (TA)
Global $TA1 = _Excel_RangeRead($oWorkbook, Default, "B112:O118")
Global $TA_N = 7

;TAGTERRASSE (TT)
Global $TT1 = _Excel_RangeRead($oWorkbook, Default, "B125:O127")
Global $TT_N = 3

;ALTANGANG (AG)
Global $AG1 = _Excel_RangeRead($oWorkbook, Default, "B134:O136")
Global $AG_N = 3

;ALTANER (AL)
Global $AL1 = _Excel_RangeRead($oWorkbook, Default, "B145:O147")
Global $AL_N = 3

;UDVENDIGE TRAPPER (UT)
Global $UT1 = _Excel_RangeRead($oWorkbook, Default, "B154:O157")
Global $UT_N = 4

;DØRE (DO)
Global $DO1 = _Excel_RangeRead($oWorkbook, Default, "B167:O172")
Global $DO_N = 6

;VINDUER (VI)
Global $VI1 = _Excel_RangeRead($oWorkbook, Default, "B179:O183")
Global $VI_N = 5

;BADEVÆRELSER OG TOILETTER (BT)
Global $BT1 = _Excel_RangeRead($oWorkbook, Default, "B192:O194")
Global $BT_N = 3

;INDVENDIGE TRAPPER (IT)
Global $IT1 = _Excel_RangeRead($oWorkbook, Default, "B203:O205")
Global $IT_N = 3

;INDVENDIGT AFLØBSSYSTEM (IA)
Global $IA1 = _Excel_RangeRead($oWorkbook, Default, "B214:O217")
Global $IA_N = 4

;TAGRENDER OG NEDLØB (TN)
Global $TN1 = _Excel_RangeRead($oWorkbook, Default, "B227:O231")
Global $TN_N = 5

;EL (EL)
Global $EL1 = _Excel_RangeRead($oWorkbook, Default, "B241:O243")
Global $EL_N = 3

;GAS OG LUFT (GL)
Global $GL1 = _Excel_RangeRead($oWorkbook, Default, "B254:O256")
Global $GL_N = 3

;BRUGSVANDSSYSTEM (BV)
Global $BV1 = _Excel_RangeRead($oWorkbook, Default, "B266:O270")
Global $BV_N = 5

;VARMEANLÆG (VA)
Global $VA1 = _Excel_RangeRead($oWorkbook, Default, "B283:O286")
Global $VA_N = 4

;VENTILATION (VE)
Global $VE1 = _Excel_RangeRead($oWorkbook, Default, "B298:O302")
Global $VE_N = 5

EndFunc

;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;						ARBEJDER						ARBEJDER						ARBEJDER						ARBEJDER						ARBEJDER
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func Arbejder()
   HotKeySet("{ESC}", "stop")
   $n = 0
   $p = 0

   ;INDSÆT ARBEJDER
   Do
	  If $Data[$n][0] = "" Then
	  Else
		 If $p = 1 Then
			$m = 0
			Do
			   Send("+{TAB}")
			   Sleep($small)
			   $m = $m + 1
			Until $m = 4
		 EndIf

		 ClipPut($Data[$n][0])
		 Send("^v")
		 Sleep($small)
		 Send("{TAB}")
		 Sleep($small)
		 ClipPut($Data[$n][2])
		 Send("^v")
		 Sleep($small)
		 Send("{TAB}")
		 Sleep($small)
		 ClipPut($Data[$n][3])
		 Send("^v")
		 Sleep($small)
		 Send("{TAB}")
		 Sleep($small)
		 ClipPut($Data[$n][4])
		 Send("^v")
		 Send("{TAB}")
		 Sleep($small)
		 Send("{SPACE}")
		 Sleep($medium)
		 $p = 1
	  EndIf
	  $n = $n + 1
   Until $n = $Data_N

   ;INDSÆT VEDLIGEHOLD
   $n = 0
   $o = 0
   $p = 0
   Do
	  If $Data[$n][7] = "" Then
	  Else
		 $o = 1
	  EndIf
	  $n = $n + 1
   Until $n = $Data_N

   If $o = 1 Then
	  Sleep($small)
	  Send("{TAB}")
	  Sleep($small)
	  Send("{SPACE}")
	  Sleep($small)
	  DCAP_Arbejder_Check()
	  Send("{TAB}")
	  Sleep($small)
	  $n = 0
	  Do
		 If $Data[$n][7] = "" Then
		 Else
			If $p = 1 Then
			   $m = 0
			   Do
				  Send("+{TAB}")
				  Sleep($small)
				  $m = $m + 1
			   Until $m = 6
			EndIf
			Sleep($small)
			ClipPut($Data[$n][7])
			Send("^v")
			Sleep($small)
			Send("{TAB}")
			Sleep($small)
			ClipPut($Data[$n][9])
			Send("^v")
			Sleep($small)
			Send("{TAB}")
			Sleep($small)
			ClipPut($Data[$n][10])
			Send("^v")
			Sleep($small)
			Send("{TAB}")
			Sleep($small)
			ClipPut($Data[$n][11])
			Send("^v")
			Sleep($small)
			Send("{TAB}")
			Sleep($small)
			ClipPut($Data[$n][12])
			Send("^v")
			Sleep($small)
			Send("{TAB}")
			Sleep($small)
			ClipPut($Data[$n][13])
			Send("^v")
			Send("{TAB}")
			Sleep($small)
			Send("{SPACE}")
			Sleep($medium)
			$p = 1
		 EndIf
		 $n = $n + 1
	  Until $n = $Data_N
   EndIf
EndFunc



;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 BELÆGNING
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func BE()

HotKeySet("{ESC}", "stop")
Global $Data = $BE1
Global $Data_N = $BE_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX1) = 1 Then
   Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 15
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Send($SD[2]); Boligareal
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 9
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()
EndIf
Neste()

EndFunc

;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 UDVENDIGT AFLØB
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func UA()

HotKeySet("{ESC}", "stop")
Global $Data = $UA1
Global $Data_N = $UA_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX2) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 22
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Send($SD[2]); Boligareal
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 16
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 YDERVÆGGE
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func YD()

HotKeySet("{ESC}", "stop")
Global $Data = $YD1
Global $Data_N = $YD_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX3) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 19
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 15
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 KVIST
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func KV()

HotKeySet("{ESC}", "stop")
Global $Data = $KV1
Global $Data_N = $KV_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX4) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 10
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][1]); Antal
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 5
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 TAG
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func TA()

HotKeySet("{ESC}", "stop")
Global $Data = $TA1
Global $Data_N = $TA_N
Global $s1 = 280

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX5) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 20
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[2]); Boligareal
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 11
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()
EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 TAGTERRASSE
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func TT()

HotKeySet("{ESC}", "stop")
Global $Data = $TT1
Global $Data_N = $TT_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX6) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 12
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 9
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 ALTANGANG
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func AG()

HotKeySet("{ESC}", "stop")
Global $Data = $AG1
Global $Data_N = $AG_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX7) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 12
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 9
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 ALTANER
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func AL()

HotKeySet("{ESC}", "stop")
Global $Data = $AL1
Global $Data_N = $AL_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX8) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 12
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 9
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 UDVENDIGE TRAPPER
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func UT()

HotKeySet("{ESC}", "stop")
Global $Data = $UT1
Global $Data_N = $UT_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX9) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 11
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 6
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 DØRE
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func DOR()

HotKeySet("{ESC}", "stop")
Global $Data = $DO1
Global $Data_N = $DO_N
Global $s1 = 480

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX10) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 20
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 17
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 VINDUER
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func VI()

HotKeySet("{ESC}", "stop")
Global $Data = $VI1
Global $Data_N = $VI_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX11) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 22
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 19
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 BAD OG TOILET
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func BT()

HotKeySet("{ESC}", "stop")
Global $Data = $BT1
Global $Data_N = $BT_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX12) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 15
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[2]); Boligareal
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 9
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 INDVENDIGE TRAPPER
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func IT()

HotKeySet("{ESC}", "stop")
Global $Data = $IT1
Global $Data_N = $IT_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX13) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 11
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 6
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 INDVENDIGT AFLØBSSYSTEM
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func IA()

HotKeySet("{ESC}", "stop")
Global $Data = $IA1
Global $Data_N = $IA_N
Global $s1 = 650

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX14) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 11
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[1][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 4
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIF
Neste()

EndFunc

;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 TAGRENDER OG NEDLØB
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func TN()

HotKeySet("{ESC}", "stop")
Global $Data = $TN1
Global $Data_N = $TN_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX15) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 20
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[2]); Boligareal
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 14
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 EL
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func EL()

HotKeySet("{ESC}", "stop")
Global $Data = $EL1
Global $Data_N = $EL_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX16) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 11
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 5
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 GAS OG LUFT
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func GL()

HotKeySet("{ESC}", "stop")
Global $Data = $GL1
Global $Data_N = $GL_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX17) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 9
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 6
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 BRUGSVANDSSYSTEM
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func BV()

HotKeySet("{ESC}", "stop")
Global $Data = $BV1
Global $Data_N = $BV_N
Global $s1 = 820

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX18) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 13
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 7
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc

;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 VARMAANLÆG
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func VA()

HotKeySet("{ESC}", "stop")
Global $Data = $VA1
Global $Data_N = $VA_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX19) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 15
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 9
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()

EndIf
Neste()

EndFunc

;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                 VENTILATIONSANLÆG
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func VE()

HotKeySet("{ESC}", "stop")
Global $Data = $VE1
Global $Data_N = $VE_N

$n = 0
$o = 0
Do
   If $Data[$n][0] = "" Then
   Else
	  $o = 1
   EndIf
   $n = $n + 1
Until $n = $Data_N

If GUICtrlRead($BOX20) = 1 Then
Else
   $o = 0
EndIf

If $o = 0 Then
Else
   ;INDSÆT STAMDATA
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($Data[0][5]); Levetid
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send("{TAB}")
   Sleep($small)
   Send($SD[8]); Etager
   Sleep($small)
   $m = 0
   Do
	  Send("{TAB}")
	  Sleep($small)
	  $m = $m + 1
   Until $m = 6
   Udfyldt_check()
   Send("{SPACE}")
   Sleep($small)
   DCAP_Arbejder_Check()
   Send("{TAB}")
   Sleep($small)

   ;INDSÆT ARBEJDER OG VEDLIGEHOLD
   Arbejder()
EndIf

EndFunc



;----------------------------------------------------------------------------------------------------------------------------------------------------------
;           GUI           GUI           GUI           GUI           GUI           GUI           GUI           GUI           GUI           GUI
;----------------------------------------------------------------------------------------------------------------------------------------------------------

Func GUI1()
   $bredde = 230
   $hojde = 700

   Global $GUI1 = GUICreate($titel, 230, 700, @DeskTopWidth-230, 80, -1 , $WS_EX_TOPMOST)

   $hPic_background = GUICtrlCreatePic(@ScriptDir & "\Pic\BAGGRUND.jpg", 0, 0, 348, 715)
   GUICtrlSetState($hPic_background, $GUI_DISABLE)

   GUISetFont(14, 800)
   GUICtrlCreateLabel("DOMUTECH", 20, 20)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

   GUISetFont(8, 300)
   GUICtrlCreateLabel($titel, 20, 40)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

   GUISetFont(11, 600)
   GUICtrlCreateLabel("Bygningsdele der indtastes:", 20, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   $h = $h + $hplus

   ;BYGNINGSDELE-------------------------------

   GUISetFont(11, 400)
   GUICtrlCreateLabel("Belægning", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX1 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus


   GUICtrlCreateLabel("Afløbsystem (udv.)", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX2 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel("Ydervægge", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX3 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel("Kviste", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX4 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel("Tagdækning", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX5 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel("Tagterrasser", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX6 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel("Altangange", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX7 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel("Altaner", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX8 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel("Trapper (udv.)", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX9 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Døre", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX10 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Vinduer", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX11 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Badeværelser", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX12 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Trapper (indv.)", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX13 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Afløbssystem (indv.)", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX14 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Tagrender", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX15 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("El", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX16 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Gas og Luft", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX17 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Brugsvand", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX18 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Varmeanlæg", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX19 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus

   GUICtrlCreateLabel ("Ventilation", 40, $h)
   GUICtrlSetColor(-1, $COLOR_WHITE)
   GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
   Global $BOX20 = GUICtrlCreateCheckbox("", 20, $h, 14, 14)
   GUICtrlSetState(-1, $GUI_CHECKED)
   $h = $h + $hplus


   Global $BUT1 = GUICtrlCreateButton("START INDTASTNING", 40, $h+20, 170, 30)

   GUISetState(@SW_SHOW, $GUI1)
EndFunc


;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;             SPLASH             SPLASH             SPLASH             SPLASH             SPLASH             SPLASH             SPLASH             SPLASH
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func SplashOn()
   Global $SPLASH = GuiCreate("", 1000, 100, -1, 0, BitOr($WS_POPUP,$WS_DLGFRAME), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
   GUISetBkColor(0x66FF66)
   GUISetFont(12, 700)
   Global $Label1 = GUICtrlCreateLabel("Programmet arbejder, brug ikke mus eller keyboard.", 0, 30, 1000, 30, $SS_CENTER)
   Global $Label2 = GUICtrlCreateLabel($text, 0, 60, 1000, 30, $SS_CENTER)
   GuiSetState()
EndFunc

Func SplashU()
   GUICtrlSetData ($Label2, $text)
EndFunc

Func SplashOfff()
   GUIDelete($SPLASH)
EndFunc

;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;             DIVERSE             DIVERSE             DIVERSE             DIVERSE             DIVERSE             DIVERSE             DIVERSE
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Func Screen()
   If @DeskTopHeight = 1080 Then
   Else
	  MsgBox($MB_SYSTEMMODAL, "Fejl", "Din skærmopløsning er: " & @DeskTopWidth  & " x " & @DeskTopHeight & ". Programmet virker kun på skærme med en opløsning på 1920 x 1080. Bruger du en tablet, skal du indstille din monitor ''hovedskærm''. Har du en 1920 x 1080 monitor skal du sætte skalering til 100%. Har du en 4K skærm skal du sænke opløsningen til 1920 x 1080. Og elles kontrakt Caroline/Danny, så finder vi en løsning.")
	  Exit
   EndIf
EndFunc

Func stop()
   Exit
EndFunc

