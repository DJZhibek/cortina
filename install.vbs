Option Explicit
Const cRegKey = "kc_cortina"

Dim Reg : Set Reg = SDB.Registry

If Reg.OpenKey(cRegKey, True) Then

	If Not Reg.ValueExists("SearchTitle") Then
		Reg.BoolValue("SearchTitle") = False
	End If
	If Not Reg.ValueExists("SearchGenre") Then
		Reg.BoolValue("SearchGenre") = True
	End If
	If Not Reg.ValueExists("SearchCustomTags") Then
		Reg.BoolValue("SearchCustomTags") = False
	End If
	If Not Reg.ValueExists("SearchPath") Then
		Reg.BoolValue("SearchPath") = True
	End If	
	If Not Reg.ValueExists("CortinaLen") Then
		Reg.IntValue("CortinaLen") = 45 ' seconds
	End If
	If Not Reg.ValueExists("FadeIn") Then
		Reg.IntValue("FadeIn") = 1
	End If
	If Not Reg.ValueExists("FadeOut") Then
		Reg.IntValue("FadeOut") = 10
	End If		
	If Not Reg.ValueExists("GapTime") Then
		Reg.IntValue("GapTime") = 3
	End If
	If Not Reg.ValueExists("CortinaVolume") Then
		Reg.IntValue("CortinaVolume") = 70
	End If		


	Reg.CloseKey
End If
