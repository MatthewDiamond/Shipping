; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


Macro1:
InputBox, test, Beep
InputBox, test, boop
WinHttpDownloadToFile("https://raw.githubusercontent.com/MatthewDiamond/Shipping/master/MyScript.ahk", "./")
Run, .\MyScript.ahk
Return


WinHttpDownloadToFile(UrlList, DestFolder)
{
	UrlList := StrReplace(UrlList, "`n", ";")
	UrlList := StrReplace(UrlList, ",", ";")
	DestFolder := RTrim(DestFolder, "\") . "\"

	Loop, Parse, UrlList, `;, %A_Space%%A_Tab%
	{
		Url := A_LoopField, FileName := DestFolder . RegExReplace(A_LoopField, ".*/")
		whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
		whr.Open("GET", Url, True)
		whr.Send()
		If (whr.WaitForResponse())
		{
			ado := ComObjCreate("ADODB.Stream")
			ado.Type := 1 ; adTypeBinary
			ado.Open
			ado.Write(whr.ResponseBody)
			ado.SaveToFile(FileName, 2)
			ado.Close
		}
	}
}
