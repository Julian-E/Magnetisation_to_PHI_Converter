#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
Opt("WinTitleMatchMode", 1)

#Region # Gui zum Auswählen, am Anfang aktiv, nach Convert-Click versteckt
$Form1 = GUICreate("AK Rentschler MvsH Converter", 434, 361, 192, 124)
GUISetBkColor(0xC0C0C0)
$Label1 = GUICtrlCreateLabel("Exel File:", 33, 24, 46, 17)
$Input1 = GUICtrlCreateInput("", 32, 48, 369, 21)
$Label2 = GUICtrlCreateLabel("Output File:", 33, 120, 58, 17)
$Input2 = GUICtrlCreateInput("", 32, 144, 369, 21)
$Button1 = GUICtrlCreateButton("Save...", 32, 168, 75, 25)
$Button2 = GUICtrlCreateButton("Search...", 32, 72, 75, 25)
$Button3 = GUICtrlCreateButton("Convert", 288, 240, 115, 89)
$Label3 = GUICtrlCreateLabel("Field / Oe:", 56, 243, 139, 17, $SS_RIGHT)
$Input3 = GUICtrlCreateInput("A11", 200, 240, 41, 21)
$Label4 = GUICtrlCreateLabel("Temperature / K:", 56, 275, 139, 17, $SS_RIGHT)
$Input4 = GUICtrlCreateInput("B11", 200, 272, 41, 21)
$Label5 = GUICtrlCreateLabel("Molar Magnetization / N*My:", 56, 306, 139, 17, $SS_RIGHT)
$Input5 = GUICtrlCreateInput("M11", 200, 304, 41, 21)
$Group1 = GUICtrlCreateGroup("Exel Start-Cells", 32, 216, 217, 121)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW, $Form1)
#EndRegion

#Region #  Lade-Gui NACH Auswählen, am Anfang versteckt, nach Convert-Click aktiv
$Form2 = GUICreate("WAIT - AK Rentschler MvsH Converter", 434, 361, 192, 124)
GUISetBkColor(0xC0C0C0)
$Label6 = GUICtrlCreateLabel("Please wait, ", 128, 120, 171, 43)
GUICtrlSetFont(-1, 24, 800, 0, "Calibri")
$Label7 = GUICtrlCreateLabel("process takes a few moments ;)", 8, 160, 421, 43)
GUICtrlSetFont(-1, 24, 800, 0, "Calibri")
GUISetState(@SW_HIDE, $Form2)
#EndRegion ### END Koda GUI section ###

Func _ArrayAssign($sArray, $vValue, $Flag = 0)
	Local $iIsArray, $sStruct, $sElem
	If Not StringRegExp($sArray, "^\h*\w+(?:\h*\[\h*\d+\h*\])*\h*$") Then Return SetError(1, 0, 0)

	Local $sVarname = StringRegExpReplace($sArray, "^\h*(\w+)[\h\d\[\]]*$", "$1")
	If Not @extended Then Return SetError(1, 0, 0)

	Local $aDims = StringRegExp($sArray, "\[\h*(\d+)\h*\]", 3)
	If @error Then
		$iIsArray = True
		If Not IsArray($vValue) Then Return SetError(1, 0, 0)
		For $i = 1 To UBound($vValue, 0)
			$sStruct &= "[" & UBound($vValue, $i)& "]"
		Next

	Else
		$iIsArray = False
		For $i = 0 To UBound($aDims) - 1
			$sStruct &= "[" & $aDims[$i] + 1  & "]"
			$sElem &= "[" & $aDims[$i] & "]"
		Next
	EndIf

	If IsDeclared($sVarname) Then
		If $Flag Then Return SetError(1, 0, 0)
	Else
		If Not Assign($sVarname, "", 2) Then Return SetError(1, 0, 0)
		Local $aTmp = _ArrayDeclare($sStruct)
		If @error Then Return SetError(1, 0, 0)
		If Not Execute("__ArrayAssignValue($" & $sVarname & ", $aTmp, 1)") Then Return SetError(1, 0, 0)
	EndIf

	If $iIsArray Then
		Execute("__ArrayAssignValue($" & $sVarname & ", $vValue, 1)")
	Else
		Execute("__ArrayAssignValue($" & $sVarname & $sElem & ", $vValue )")
	EndIf
	If @error Then Return SetError(1, 0, 0)

	Return 1
 EndFunc

Func __ArrayAssignValue(ByRef $aArray, $aValues, $iFlag = 0)
	If $iFlag And Not IsArray($aValues) Then Return SetError(1, 0, 0)
	$aArray = $aValues
	Return 1
 EndFunc ; ===> __ArrayAssignValue

Func Convert()

$handleGUI1 = WinGetHandle("[TITLE:AK Rentschler MvsH Converter; CLASS:AutoIt v3 GUI]")
$handleGUI2 = WinGetHandle("[TITLE:WAIT - AK Rentschler MvsH Converter; CLASS:AutoIt v3 GUI]")
$positionGUI1 = WinGetPos($handleGUI1)
Winmove($handleGUI2,"",$positionGUI1[0],$positionGUI1[1])
GUISetState(@SW_HIDE, $Form1)
GUISetState(@SW_SHOW, $Form2)

Local $loadpath = GUICtrlRead($Input1)
Local $savepath = GUICtrlRead($Input2)
Local $Fieldstart = GUICtrlRead($Input3)
Local $Tempstart = GUICtrlRead($Input4)
Local $Magstart = GUICtrlRead($Input5)

#Region# Split Exel Start-Cells to get the letter and the number: Field
   $split = StringSplit($Fieldstart,"",2)
   $numberdigits = UBound($split)
   $number = String("")
   For $a = 1 to $numberdigits-1
	  $number = String($number & $split[$a])
   Next
   $Fieldstart = INT($number)
   Local $FieldstartLetter = $split[0]
   ;MsgBox(0,"Buchstabe",$FieldstartLetter)
   ;MsgBox(0,"Nummer",$Fieldstart)
#Endregion#

#Region# Split Exel Start-Cells to get the letter and the number: Temp
   $split = StringSplit($Tempstart,"",2)
   $numberdigits = UBound($split)
   $number = String("")
   For $a = 1 to $numberdigits-1
	  $number = String($number & $split[$a])
   Next
   $Tempstart = INT($number)
   Local $TempstartLetter = $split[0]
  ; MsgBox(0,"Buchstabe",$TempstartLetter)
  ; MsgBox(0,"Nummer",$Tempstart)
#Endregion#

#Region# Split Exel Start-Cells to get the letter and the number: Mag
   $split = StringSplit($Magstart,"",2)
   $numberdigits = UBound($split)
   $number = String("")
   For $a = 1 to $numberdigits-1
	  $number = String($number & $split[$a])
   Next
   $Magstart = INT($number)
   Local $MagstartLetter = $split[0]
   ;MsgBox(0,"Buchstabe",$MagstartLetter)
   ;MsgBox(0,"Nummer",$Magstart)
#Endregion#

#Region# Open Exelsheet, if already opened: dont hide, if not opened before: hide window!
   ;Split Filepath to get Name---------
   $split = StringSplit($loadpath,"\",2)
   $numberfragments = UBound($split)
   $split = StringSplit($split[$numberfragments-1],".",2)
   ;($split[0] = Dateiname)
   $exists = WinExists("[TITLE:"&$split[0]&"; CLASS:XLMAIN]")
   ;-------------------------------------------$MB_SYSTEMMODAL
   Local $oExcel = _Excel_Open()
   If @error Then
	  MsgBox($MB_SYSTEMMODAL, "Error", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Exit
   EndIf
   Local $oWorkbook = _Excel_BookOpen($oExcel, $loadpath)
   If $exists = 0 Then
	  Local $handle = WinGetHandle("[TITLE:"&$split[0]&"; CLASS:XLMAIN]")
	  WinSetState($handle,"",@SW_HIDE)
   EndIf
   If @error Then
	  MsgBox($MB_SYSTEMMODAL, "Error", "Couldn't load Exel sheet. Path correct?")
		 _Excel_BookClose($oWorkbook)
	  Exit
   EndIf
#EndRegion

Local $f = 0
Local $t = 0
Local $v = 0
Local $aField[$f]
Local $aTemp[$t]
Local $aValue[$v]

#Region# Lese Feld und speicher in $aField[$j]
While 1
Local $Read = _Excel_RangeRead($oWorkbook, Default, $FieldstartLetter & $Fieldstart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Error", "Couldn't load Exel sheet. Path correct?")
If $Read = "" Then
   ExitLoop
Else
   _ArrayAdd($aField, $Read)
   $f = $f+1
   $Fieldstart = $Fieldstart+1
EndIf
WEnd
#EndRegion

#Region# Lese Temp und speicher in $aTemp[$j]
While 2
Local $Read = _Excel_RangeRead($oWorkbook, Default, $TempstartLetter & $Tempstart)
$round = Round($Read,0)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Error", "Couldn't load Exel sheet. Path correct?")
If $Read = "" Then
   ExitLoop
Else
   _ArrayAdd($aTemp, $round)
   $t = $t+1
   $Tempstart = $Tempstart+1
EndIf
WEnd
#EndRegion

#Region# Lese Wert und speicher in $aValue[$j]
While 3
Local $Read = _Excel_RangeRead($oWorkbook, Default, $MagstartLetter & $Magstart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Error", "Couldn't load Exel sheet. Path correct?")
If $Read = "" Then
   ExitLoop
Else
   _ArrayAdd($aValue, $Read)
   $v = $v+1
   $Magstart = $Magstart+1
EndIf
WEnd
#EndRegion

;_ArrayDisplay($aField)
;_ArrayDisplay($aTemp)
;_ArrayDisplay($aValue)

#Region# Bestimme Anzahl an Feldern
Local $amountfields = 0
Local $check = $aTemp[0]
For $t = 0 to UBound($aTemp) - 1
   If $check = $aTemp[$t] Then
	  $amountfields = $amountfields +1
   EndIf
Next
;MsgBox(0,"Felder",$amountfields)
#EndRegion#

#Region# Bestimme Anzahl an Temperaturen
Local $amounttemps = ($f) / $amountfields
;MsgBox(0,"temps",$amounttemps)
#EndRegion

#Region# Trage Felder in eine neue 2D Array ein, geordnet nach Temperatur
Local $i = 0
Local $j = 0
Global $groupField[$amountfields][$amounttemps]
;_ArrayDisplay($groupField)
Local $helperrepeat = 0
Local $helperfieldswitch = 0
Local $arrayrow = 0
$f = 0
;MsgBox(0,"Field1",$aField[$f])
While $helperrepeat < $amounttemps

  while  $helperfieldswitch < $amountfields
	  _ArrayAssign("groupField["&$i&"]["&$j&"]", $aField[$f])
	  $f = $f + 1
	  $helperfieldswitch = $helperfieldswitch + 1
	  $i = $i + 1
  WEnd
   $j = $j + 1
   $i = 0
   $helperfieldswitch = 0
   $helperrepeat = $helperrepeat + 1
WEnd
;_ArrayDisplay($groupField)
#EndRegion

#Region# Mittel die Felder und trage sie in eine 1D Array ein
Global $aFieldave[$amountfields]
Local $aveFieldvalue = 0
Local $k = 0
$i = 0
$j = 0
$helperrepeat = 0
$helperfieldswitch = 0
While $helperfieldswitch < $amountfields

   While $helperrepeat < $amounttemps
	  $aveFieldvalue = $aveFieldvalue + $groupField[$i][$j]
	  $j = $j +1
	  $helperrepeat = $helperrepeat + 1
   WEnd
   $aveFieldvalue = $aveFieldvalue / $amounttemps
   _ArrayAssign("aFieldave["&$k&"]", $aveFieldvalue)
   $aveFieldvalue = 0
   $k = $k + 1
   $i = $i + 1
   $j = 0
   $helperrepeat = 0
   $helperfieldswitch = $helperfieldswitch + 1
WEnd
;_ArrayDisplay($aFieldave)
#EndRegion

#Region# Schreibe gemittelte Felder und die Messwerte in Datei, geordnet wie es PHI X_mag.exp verlangt
$i = 0
$k = 0
$helperrepeat = 0
$helperfieldswitch = 0
Local $writeField
Local $writeValue
While $helperfieldswitch < $amountfields
   $writeField = StringFormat("%f", $aFieldave[$k]/10000)
   FileWrite($savepath, ($writeField) & @TAB)
   $k = $k +1

   While $helperrepeat < $amounttemps
	  $writeValue = StringFormat("%e", $aValue[$i])
	  FileWrite($savepath, $writeValue & @TAB)
	  $i = $i + $amountfields
	  $helperrepeat = $helperrepeat +1
   WEnd
   FileWrite($savepath, @CRLF)
   $helperrepeat = 0
   $helperfieldswitch = $helperfieldswitch + 1
   $i = $helperfieldswitch
WEnd
#EndRegion

If $exists = 0 Then ;Falls das Exelsheet for dem Programm nciht geöffnet war wird es hier wieder beendet!
   _Excel_BookClose($oWorkbook)
EndIf

Exit

EndFunc

While 1
   $nMsg = GUIGetMsg()

   if $nMsg = $GUI_EVENT_CLOSE Then
	  Exit
   ElseIf $nMsg = $Button2 Then ; search button
	  $loadDialog = FileOpenDialog("Search...",@ScriptDir,"Exel file (*.xlsx)|Exel 97-2003 (*.xls)|CSV (MS-DOS) (*.csv)|All files (*.*)" )
	  if @error Then
		 $loadpath = GUICtrlRead($Input1)
		 GUICtrlSetData($Input1,$loadpath)
	  Else
		 GUICtrlSetData($Input1,$loadDialog)
	  EndIf
   ElseIf $nMsg = $Button1 Then ; save Button
	  $saveDialog = FileSaveDialog("Save...","","Txt file (*.txt) |Dat file (*.dat)|Exp file (*.exp)|All files (*.*)" )
	  if @error Then
		 $savepath = GUICtrlRead($Input2)
		 GUICtrlSetData($Input2,$savepath)
	  Else
		 GUICtrlSetData($Input2,$saveDialog)
	  EndIf
   ElseIf $nMsg = $Button3 Then
	  Convert()
	  Exit
   EndIf
WEnd
