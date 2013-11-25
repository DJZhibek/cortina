'
'	MediaMonkey Script
'	Name: cortina
'	[kc_cortina]
'	Filename=cortina.vbs
'	Description=Turn songs into cortinas without editing
'	Language=VBScript
'	ScriptType=0 Auto Script

'	cortina.vbs is copied to scripts\auto
'	Use  Play/Cortinas  at any time to set options.
'	--------------------------------------------------------------------	
'	This program is free software: you can redistribute it and/or modify
'	it under the terms of the GNU General Public License as published by
'	the Free Software Foundation, either version 3 of the License, or
'	(at your option) any later version.
'
'	This program is distributed in the hope that it will be useful,
'	but WITHOUT ANY WARRANTY; without even the implied warranty of
'	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'	GNU General Public License for more details.
'
'	You should have received a copy of the GNU General Public License
'	along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
Option Explicit


' registry key name
Const cRegKey = "kc_cortina"

' set the default values
Dim bSearchTitle : bSearchTitle = False			' do we search song title for 'cortina'?
Dim bSearchGenre : bSearchGenre = False			' do we search song genre for 'cortina'?
Dim bSearchAllCustomTags : bSearchAllCustomTags = False	' do we search song Custom1 tag for 'cortina'?
Dim bSearchPath : bSearchPath = False			' do we search song path for 'cortina'?
Dim iCortinaLen : iCortinaLen = 45				' default cortina length in seconds (includes fade-in and fade-out time)
Dim dCortinaVolume : dCortinaVolume = 0.6		' default cortina volume multiplier
Dim iFadeIn : iFadeIn = 0						' default fade-in time in seconds
Dim iFadeOut : iFadeOut = 5						' default fade-out time in seconds
Dim iGapTime : iGapTime = 1						' default gap time in seconds (additional silence added after cortina)
Dim dSongVolume : dSongVolume = 1.0				' storage for current playback volume (copied before cortina volume modifies it)


' cortina status constants
Const cNone = 0
Const cFadeIn = 1
Const cFullVolume = 2
Const cFadeOut = 3
Const cGap = 4
Dim CortinaState : CortinaState = cNone			' keeps track of where we are in the cortina playback

' globals used during cortina playback
Dim iStateCounter			' timer counter used by all cortina states
Dim dVolumeInc				' volume increment for fade-in and fade-out (calculated from cortina volume and fade times)

' global objects used to track cortina progress
Dim ProgressDisplay : Set ProgressDisplay = Nothing		' used to display cortina progress text
Dim ProgressTimer : Set ProgressTimer = Nothing			' used to keep track of how long the cortina has been playing
Dim StateTimer : Set StateTimer = Nothing				' used as interrupt timer for altering cortina volume

' Called when MediaMonkey is starting up
Sub OnStartup
	' Add a menu item for easy access to cortina settings from the Play menu
	Dim MenuItem : Set MenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_Play, -2, -1) ' last item in the second to last part of Play menu
	MenuItem.Caption = "Cortinas"
	Script.RegisterEvent MenuItem, "OnClick", "ShowForm"
	MenuItem.Visible = True
	
	ReadSettings	' Read previously saved settings
	CreateTimers	' Create timers used by cortinas
	
	' Register MediaMonkey events we need to act on
	Script.RegisterEvent SDB, "OnPlay", "Event_OnPlay"
	Script.RegisterEvent SDB, "OnPlaybackEnd", "Cleanup"
	Script.RegisterEvent SDB, "OnShutdown", "Cleanup"
	
End Sub

' create timers for cortina playback
Sub CreateTimers()
	' keeps track of cortina playback time
	Set ProgressTimer = SDB.CreateTimer(10000) ' correct times are set later
	ProgressTimer.Enabled = False
	Script.RegisterEvent ProgressTimer, "OnTimer", "OnProgressTimer"

	' used as an interrupt timer for changing corinta volume
	Set StateTimer = SDB.CreateTimer(10000)  ' correct times are set later
	StateTimer.Enabled = False
	Script.RegisterEvent StateTimer, "OnTimer", "OnStateTimer"
End Sub

' Disable timers, set cortina state to none, destroy reference to the player progress display
Sub DisableTimers()
	On Error Resume Next
	If IsObject(StateTimer) Then StateTimer.Enabled = False
	If IsObject(ProgressTimer) Then ProgressTimer.Enabled = False
	If IsObject(ProgressDisplay) Then Set ProgressDisplay = Nothing
	On Error GoTo 0
End Sub

' cleanup function for playback ending and MediaMonkey shutdown
Sub Cleanup()
	CortinaState = cNone
	DisableTimers
	SDB.Player.Volume = dSongVolume
End Sub

' tests current song to see if it is a cortina
Function Is_Cortina()
	Dim objSongData : Set objSongData = SDB.Player.CurrentSong
	Is_Cortina = False
	
	' test selected locations to see if the word "cortina" exists (not case sensitive)
	If bSearchGenre Then
		If Instr(1,objSongData.Genre,"cortina",1) > 0 Then Is_Cortina = True
		Exit Function
	End If
	
	If bSearchTitle Then
		If Instr(1,objSongData.Title,"cortina",1) > 0 Then Is_Cortina = True
		Exit Function
	End If

	If bSearchPath Then
		If Instr(1,objSongData.Path,"cortina",1) > 0 Then Is_Cortina = True
		Exit Function
	End If

	If bSearchAllCustomTags Then
		' try to be a little faster by only searching until "cortina" is found
		If Instr(1,objSongData.Custom1,"cortina",1) > 0 Then 
			Is_Cortina = True
		Else 
			If Instr(1,objSongData.Custom2,"cortina",1) > 0 Then 
				Is_Cortina = True
			Else
				If Instr(1,objSongData.Custom3,"cortina",1) > 0 Then
					Is_Cortina = True
				Else
					If Instr(1,objSongData.Custom4,"cortina",1) > 0 Then
						Is_Cortina = True
					Else
						If Instr(1,objSongData.Custom5,"cortina",1) > 0 Then Is_Cortina = True
					End If
				End If
			End If
		End If
	End If

End Function

' Calculate the length of the full volume portion of the cortina in seconds
Function FullVolumeLength()
	Dim iSongLength : iSongLength = CInt(SDB.Player.CurrentSong.SongLength / 1000.0) ' convert current song length to seconds
	
	' subtract fade-in and fade-out times from cortina length
	If iSongLength > iCortinaLen Then
		FullVolumeLength = iCortinaLen - (iFadeIn + iFadeOut)
	Else
		FullVolumeLength = iSongLength - (iFadeIn + iFadeOut)
	End If
	
	' if this happens, we are really in trouble, cortina might end up longer than the song
	If FullVolumeLength < 1 Then FullVolumeLength = 0

End Function

' Calculate fade out volume
Function FadeOutVolume(dFadeLength, dStartVolume, dTimePoint)
	FadeOutVolume = dStartVolume * ((cos((dTimePoint/dFadeLength) * 3.1415) + 1.0) / 2.0)
	'FadeOutVolume = dStartVolume * ((1.0 / (dTimePoint / dFadeLength + 0.05) - 1.0) / 19.0)
	If FadeOutVolume < 0.0 Then FadeOutVolume = 0.0
End Function

' called when a song starts to play
Sub Event_OnPlay()
	
	' check if this song is a cortina
	If Is_Cortina() = False Then Exit Sub

	ReadSettings ' get current settings

	' do nothing if cortina length is not greater than zero
	If iCortinaLen <= 0 Then Exit Sub
	
	' retrieve current song length
	Dim iSongLength
	iSongLength = CInt(SDB.Player.CurrentSong.SongLength / 1000.0) ' convert current song length to seconds

	Dim iTotalTime		' Total time for cortina and silence gap
	If iSongLength < iCortinaLen Then ' if song is shorter than song length, shorten cortina length
		iTotalTime = iSongLength + iGapTime
	Else
		iTotalTime = iCortinaLen + iGapTime
	End If

	' set up a progress display for the cortina
	Set ProgressDisplay = SDB.Progress
	ProgressDisplay.MaxValue = iTotalTime
	ProgressTimer.Interval = 1000 ' 1000 ms = 1 second
	ProgressDisplay.Text = "Cortina: " & iTotalTime & " seconds left."
	
	' save current playback volume
	dSongVolume = SDB.Player.Volume
	
	' calculate cortina volume
	dCortinaVolume = dSongVolume * dCortinaVolume
	
	' If fade in timer is used, start it
	If iFadeIn > 0 Then 
		SDB.Player.Volume = 0.0		' turn volume all the way down
		CortinaState = cFadeIn		' set cortina state tracker to fade-in
		iStateCounter = iFadeIn * 4	' convert seconds to quarter seconds
		' calculate volume decrement value from current settings
		dVolumeInc = dCortinaVolume / CDbl(iStateCounter)	
		StateTimer.Interval = 250	' set timer interval to 1/4 second (250ms)
	Else
		' No fade in, set cortina volume and start full cortina volume timer
		SDB.Player.Volume = dCortinaVolume
		CortinaState = cFullVolume
		iStateCounter = FullVolumeLength()
		StateTimer.Interval = 1000	' set timer interval to 1 second for full volume part of cortina (1000ms)		
	End If		
		
	ProgressTimer.Enabled = True ' start the cortina progress display timer
	StateTimer.Enabled = True	' start the interrupt timer

End Sub

' Update the progress display (shows a progress bar and info text while cortina is playing)
Sub OnProgressTimer(thisTimer)
	' check if progress was terminated
	If Not isObject(ProgressDisplay) Then 
		Set ProgressDisplay = Nothing
		Exit Sub
	End If
	If ProgressDisplay.Terminate = True Then
		thisTimer.Enabled = False
		Set ProgressDisplay = Nothing
		Exit Sub
	End If
	
	' thisTimer is a reference to the ProgressTimer passed this interrupt handler
	If thisTimer.Enabled Then
		'thisTimer.Enabled = False
		ProgressDisplay.Increase
		ProgressDisplay.Text = "Cortina: " & ProgressDisplay.MaxValue - ProgressDisplay.Value & " seconds left."
		'thisTimer.Enabled = True
	End If
	' check if we need to turn off the progress display
	If Not isObject(ProgressDisplay) Or ProgressDisplay.Terminate = True Or ProgressDisplay.Value >= ProgressDisplay.MaxValue Then
		thisTimer.Enabled = False
		Set ProgressDisplay = Nothing
	End If
End Sub

' Stop cortina playback, reset volume, start playing next song.
Sub GoToNextSong()
	Dim Player : Set Player = SDB.Player
	Player.Stop

	While Player.isPlaying
		SDB.ProcessMessages
	WEnd	

	If Player.CurrentSongIndex < Player.PlaylistCount Then
		Player.Next
		Player.Play
	End If
	Player.Volume = dSongVolume
	
End Sub

' Handle Volume Timer event
Sub OnStateTimer(thisTimer)
	If thisTimer.Enabled = False Then Exit Sub
	
	thisTimer.Enabled = False ' Stop timer to avoid triggering while we are here
	
	' Act on current state of cortina
	Select Case CortinaState
		Case cNone ' Should never happen
			Exit Sub

		Case cFadeIn ' handle fade in timer interrupt
			' increment the fade-in timer
			If iStateCounter > 0 Then
				If SDB.Player.Volume + dVolumeInc < dCortinaVolume Then
					SDB.Player.Volume = SDB.Player.Volume + dVolumeInc
				Else
					SDB.Player.Volume = dCortinaVolume
				End If
				iStateCounter = iStateCounter - 1
			Else
				' Switch to FullVolume State
				SDB.Player.Volume = dCortinaVolume
				iStateCounter = FullVolumeLength()
				thisTimer.Interval = 1000
				CortinaState = cFullVolume
			End If
			thisTimer.Enabled = True	' re-start the volume interrupt timer

		Case cFullVolume	' handle full volume portion of cortina
			If iStateCounter > 0 Then
				iStateCounter = iStateCounter - 1	' decrement length counter
				thisTimer.Enabled = True			' re-start the volume interrupt timer
			Else
				' we are done with the full volume part of the cortina, start the fade-out if any
				If iFadeOut > 0 Then
					CortinaState = cFadeOut			' set state to fade-out
					iStateCounter = iFadeOut * 4	' convert seconds to quarter seconds
					' calculate volume decrement value from current settings
					dVolumeInc = dCortinaVolume / CDbl(iStateCounter+1)
					thisTimer.Interval = 250		' set the timer to 1/4 second (250ms)
					thisTimer.Enabled = True		' re-start the volume interrupt timer
				Else
					' if no fade-out, check if silence gap is set
					If iGapTime > 0 Then
						SDB.Player.Volume = 0.0		' turn volume all the way down
						CortinaState = cGap			' set state to gap time
						iStateCounter = iGapTime	' set counter to gap time
						thisTimer.Interval = 1000	' set interrupt timer to 1 second (1000ms)
						thisTimer.Enabled = True	' re-start the interrupt timer
					Else
						' no gap either, we are done
						CortinaState = cNone
						DisableTimers
						GoToNextSong
					End If			
				End If
			End If
			
			
		Case cFadeOut	' handle fade-out timer interrupt	
			If iStateCounter > 0 Then
				' still fading out, lower volume
				'If SDB.Player.Volume > dVolumeInc Then
					'SDB.Player.Volume = SDB.Player.Volume - dVolumeInc
				If SDB.Player.Volume > 0.0 Then
					SDB.Player.Volume = FadeOutVolume(CDbl(iFadeOut * 4), dCortinaVolume, CDbl((iFadeOut * 4)-iStateCounter))
				Else
					SDB.Player.Volume = 0.0
				End If
				iStateCounter = iStateCounter - 1
				thisTimer.Enabled = True		' re-start the interrupt timer
			Else
				SDB.Player.Volume = 0.0		' fade-out is finished, make sure volume is zero
				If iGapTime > 0 Then		' check if a silence gap is set
					CortinaState = cGap				' set state to gap
					iStateCounter = iGapTime		' set counter to gap time
					thisTimer.Interval = 1000		' gap timer is 1 second (1000ms)
					thisTimer.Enabled = True		' re-start the interrupt timer
				Else
					' no gap, we are done
					CortinaState = cNone
					DisableTimers
					GoToNextSong
				End If
			End If
			
		Case cGap	' handle silence gap timer interrupt
			If iStateCounter > 0 Then
				iStateCounter = iStateCounter - 1	' decrement silence gap counter
				thisTimer.Enabled = True			' re-start the interrupt timer
			Else
				CortinaState = cNone		' silence gap is finish, set state to none
				DisableTimers
				GoToNextSong
			End If
		
	End Select
		
End Sub

' Save settings from options form to the registry
Sub SaveSettings(Form1)
	Dim Reg : Set Reg = SDB.Registry
	Set Form1 = Form1.Common

	If Reg.OpenKey(cRegKey, True) Then
		Reg.BoolValue("SearchTitle") = CBool(Form1.ChildControl("cb_SearchTitles").Checked)
		Reg.BoolValue("SearchGenre") = CBool(Form1.ChildControl("cb_SearchGenre").Checked)
		Reg.BoolValue("SearchCustomTags") = CBool(Form1.ChildControl("cb_SearchCustomTags").Checked)
		Reg.BoolValue("SearchPath") = CBool(Form1.ChildControl("cb_SearchPath").Checked)
		Reg.IntValue("CortinaLen") = CInt(Form1.ChildControl("tb_CortinaLen").Value)
		Reg.IntValue("FadeIn") = CInt(Form1.ChildControl("tb_FadeIn").Value)
		Reg.IntValue("FadeOut") = CInt(Form1.ChildControl("tb_FadeOut").Value)
		Reg.IntValue("GapTime") = CInt(Form1.ChildControl("tb_GapTime").Value)
		Reg.IntValue("CortinaVolume") = CInt(Form1.ChildControl("tb_CortinaVolume").Value)
		Reg.CloseKey

		ReadSettings ' Read the current settings back into global variables
	End If
End Sub

' read settings from registry and assign them to gloabl variables
Sub ReadSettings()
	Dim Reg : Set Reg = SDB.Registry
	
	' if a value does not already exist, then the default value set at startup will used
	If Reg.OpenKey(cRegKey, True) Then
		If Reg.ValueExists("SearchTitle") Then
			bSearchTitle = Reg.BoolValue("SearchTitle")
		End If
		If Reg.ValueExists("SearchGenre") Then
			bSearchGenre = Reg.BoolValue("SearchGenre")
		End If
		If Reg.ValueExists("SearchCustomTags") Then
			bSearchAllCustomTags = Reg.BoolValue("SearchCustomTags")
		End If
		If Reg.ValueExists("SearchPath") Then
			bSearchPath = Reg.BoolValue("SearchPath")
		End If
		If Reg.ValueExists("CortinaLen") Then
			iCortinaLen = Reg.IntValue("CortinaLen")
		End If
		If Reg.ValueExists("FadeIn") Then
			iFadeIn = Reg.IntValue("FadeIn")
		End If
		If Reg.ValueExists("FadeOut") Then
			iFadeOut = Reg.IntValue("FadeOut")
		End If		
		If Reg.ValueExists("GapTime") Then
			iGapTime = Reg.IntValue("GapTime")
		End If
		If Reg.ValueExists("CortinaVolume") Then
			dCortinaVolume = CDbl(Reg.IntValue("CortinaVolume")) / 100.0
		End If		

		Reg.CloseKey	
	End If
End Sub

' update cortina length display when trackbar is changed
Sub length_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_tbCurLen")
	Label1.Caption = obj.Value & " sec"
End Sub

' update fade-in display when trackbar is changed
Sub fadein_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_tbCurFadeIn")
	Label1.Caption = obj.Value & " sec"
End Sub

' update fade-out display when trackbar is changed
Sub fadeout_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_tbCurFadeOut")
	Label1.Caption = obj.Value & " sec"
End Sub

' update silence gap length display when trackbar is changed
Sub gap_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_tbCurGap")
	Label1.Caption = obj.Value & " sec"
End Sub

' update cortina volume display when trackbar is changed
Sub volume_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_CortinaVolume")
	Label1.Caption = "Cortina Volume: " & obj.Value & "%"
End Sub


' Displays the options form
Sub ShowForm(arg)
	'*******************************************************************'
	'* Form produced by MMVBS Form Creator (http://trixmoto.net/mmvbs) *'
	'*******************************************************************'
	
	Dim Form1 : Set Form1 = SDB.UI.NewForm
	Form1.BorderStyle = 1
	Form1.Caption = "Cortina Options"
	Form1.FormPosition = 4
	Form1.SavePositionName = "FrmCortinaPos"
	Form1.StayOnTop = True
	Form1.Common.SetRect 100,100,480,560
	Form1.Common.ControlName = "frm_CortinaOptions"
		
	Dim Label1 : Set Label1 = SDB.UI.NewLabel(Form1)
	Label1.Common.SetRect 15,6,120,17
	Label1.Caption = "Select where to search for the word ""cortina"""
	
	Dim CheckBox1 : Set CheckBox1 = SDB.UI.NewCheckBox(Form1)
	CheckBox1.Caption = "Title"
	CheckBox1.Checked = bSearchTitle
	CheckBox1.Common.SetRect 22,25,98,20
	CheckBox1.Common.ControlName = "cb_SearchTitles"
	CheckBox1.Common.Hint = "Search song title for the word ""cortina"""
	
	Dim CheckBox2 : Set CheckBox2 = SDB.UI.NewCheckBox(Form1)
	CheckBox2.Caption = "Genre"
	CheckBox2.Checked = bSearchGenre
	CheckBox2.Common.SetRect 132,25,98,20
	CheckBox2.Common.ControlName = "cb_SearchGenre"
	CheckBox2.Common.Hint = "Search song genre for the word ""cortina"""
	
	Dim CheckBox3 : Set CheckBox3 = SDB.UI.NewCheckBox(Form1)
	CheckBox3.Caption = "Custom Tags"
	CheckBox3.Checked = bSearchAllCustomTags
	CheckBox3.Common.SetRect 238,25,110,20
	CheckBox3.Common.ControlName = "cb_SearchCustomTags"
	CheckBox3.Common.Hint = "Search song's Custom tags for the word ""cortina"""

	Dim CheckBox4 : Set CheckBox4 = SDB.UI.NewCheckBox(Form1)
	CheckBox4.Caption = "Path"
	CheckBox4.Checked = bSearchPath
	CheckBox4.Common.SetRect 354,25,110,20
	CheckBox4.Common.ControlName = "cb_SearchPath"
	CheckBox4.Common.Hint = "Search song's directory path on disk for the word ""cortina"""
		
	Dim TrackBar1 : Set TrackBar1 = SDB.UI.NewTrackBar(Form1)
	TrackBar1.MaxValue = 240
	TrackBar1.MinValue = 15
	TrackBar1.Value = iCortinaLen
	TrackBar1.Common.SetRect 16,60,450,45
	TrackBar1.Common.ControlName = "tb_CortinaLen"
	TrackBar1.Common.Hint = "Cortina length"
	Script.RegisterEvent TrackBar1, "OnChange", "length_change"

	
	Dim Label2 : Set Label2 = SDB.UI.NewLabel(Form1)
	Label2.Autosize = False
	Label2.Common.SetRect 15,105,65,17
	Label2.Caption = "15 sec"
	
	Dim Label3 : Set Label3 = SDB.UI.NewLabel(Form1)
	Label3.Common.SetRect 432,105,65,17
	Label3.Caption = "240 sec"
	
	Dim Label4 : Set Label4 = SDB.UI.NewLabel(Form1)
	Label4.Common.SetRect 182,105,70,17
	Label4.Caption = "Cortina Length:"
	
	Dim Label5 : Set Label5 = SDB.UI.NewLabel(Form1)
	Label5.Common.SetRect 264,105,80,17
	Label5.Common.ControlName = "lbl_tbCurLen"
	Label5.Caption = iCortinaLen & " sec"

	
	Dim TrackBar2 : Set TrackBar2 = SDB.UI.NewTrackBar(Form1)
	TrackBar2.MaxValue = 10
	TrackBar2.Value = iFadeIn
	TrackBar2.Common.SetRect 16,147,450,45
	TrackBar2.Common.ControlName = "tb_FadeIn"
	TrackBar2.Common.Hint = "Cortina Fade In Time"
	Script.RegisterEvent TrackBar2, "OnChange", "fadein_change"
	
	Dim Label6 : Set Label6 = SDB.UI.NewLabel(Form1)
	Label6.Common.SetRect 16,190,65,17
	Label6.Caption = "0 sec"
	
	Dim Label7 : Set Label7 = SDB.UI.NewLabel(Form1)
	Label7.Common.SetRect 191,190,65,17
	Label7.Caption = "Fade In Time:"
	
	Dim Label8 : Set Label8 = SDB.UI.NewLabel(Form1)
	Label8.Common.SetRect 264,190,65,17
	Label8.Common.ControlName = "lbl_tbCurFadeIn"
	Label8.Caption = iFadeIn & " sec"
	
	Dim Label9 : Set Label9 = SDB.UI.NewLabel(Form1)
	Label9.Common.SetRect 431,190,65,17
	Label9.Caption = "10 sec"
	
	Dim TrackBar3 : Set TrackBar3 = SDB.UI.NewTrackBar(Form1)
	TrackBar3.MaxValue = 15
	TrackBar3.MinValue = 1
	TrackBar3.Value = iFadeOut
	TrackBar3.Common.SetRect 16,230,450,45
	TrackBar3.Common.ControlName = "tb_FadeOut"
	TrackBar3.Common.Hint = "Cortina Fade Out Time"
	Script.RegisterEvent TrackBar3, "OnChange", "fadeout_change"
	
	Dim Label10 : Set Label10 = SDB.UI.NewLabel(Form1)
	Label10.Common.SetRect 186,275,65,17
	Label10.Caption = "Fade Out Time:"
	
	Dim Label11 : Set Label11 = SDB.UI.NewLabel(Form1)
	Label11.Common.SetRect 16,275,65,17
	Label11.Caption = "1 sec"
	
	Dim Label12 : Set Label12 = SDB.UI.NewLabel(Form1)
	Label12.Common.SetRect 430,275,85,17
	Label12.Caption = "15 sec"
	
	Dim Label13 : Set Label13 = SDB.UI.NewLabel(Form1)
	Label13.Common.SetRect 264,275,65,17
	Label13.Common.ControlName = "lbl_tbCurFadeOut"
	Label13.Caption = iFadeOut & " sec"
	
	Dim TrackBar4 : Set TrackBar4 = SDB.UI.NewTrackBar(Form1)
	TrackBar4.MaxValue = 5
	TrackBar4.MinValue = 0
	TrackBar4.Value = iGapTime
	TrackBar4.Common.SetRect 16,317,450,45
	TrackBar4.Common.ControlName = "tb_GapTime"
	TrackBar4.Common.Hint = "Additional gap of silence after cortina"
	Script.RegisterEvent TrackBar4, "OnChange", "gap_change"
	
	Dim Label14 : Set Label14 = SDB.UI.NewLabel(Form1)
	Label14.Common.SetRect 17,359,65,17
	Label14.Caption = "0 sec"
	
	Dim Label15 : Set Label15 = SDB.UI.NewLabel(Form1)
	Label15.Common.SetRect 208,359,65,17
	Label15.Caption = "Gap Time:"
	
	Dim Label16 : Set Label16 = SDB.UI.NewLabel(Form1)
	Label16.Common.SetRect 264,359,80,17
	Label16.Common.ControlName = "lbl_tbCurGap"
	Label16.Caption = iGapTime & " sec"
	
	Dim Label17 : Set Label17 = SDB.UI.NewLabel(Form1)
	Label17.Common.SetRect 437,359,65,17
	Label17.Caption = "5 sec"
	
	Dim iCortinaVolume: iCortinaVolume = CInt(dCortinaVolume * 100.0)
	Dim TrackBar5 : Set TrackBar5 = SDB.UI.NewTrackBar(Form1)
	TrackBar5.MaxValue = 100
	TrackBar5.MinValue = 0
	TrackBar5.Value = iCortinaVolume
	TrackBar5.Common.SetRect 16,402,450,45
	TrackBar5.Common.ControlName = "tb_CortinaVolume"
	TrackBar5.Common.Hint = "Percentage to lower Cortina volume relative to tandas"
	Script.RegisterEvent TrackBar5, "OnChange", "volume_change"
	
	Dim Label18 : Set Label18 = SDB.UI.NewLabel(Form1)
	Label18.Common.SetRect 200,445,65,17
	Label18.Common.ControlName = "lbl_CortinaVolume"
	Label18.Common.Hint = "Lowers cortina volume relative to tandas"
	Label18.Caption = "Cortina Volume: " & iCortinaVolume & "%"
	
	Dim Button1 : Set Button1 = SDB.UI.NewButton(Form1)
	Button1.Default = True
	Button1.Caption = "Save"
	Button1.ModalResult = 1
	Button1.Common.SetRect 393,500,75,25
	Button1.Common.ControlName = "btn_SaveOptions"
	
	Dim Button2 : Set Button2 = SDB.UI.NewButton(Form1)
	Button2.Cancel = True
	Button2.Caption = "Cancel"
	Button2.ModalResult = 2
	Button2.Common.SetRect 307,500,75,25
	Button2.Common.ControlName = "btn_CancelOptions"	
	
	'*******************************************************************'
	'* End of form                              Richard Lewis (c) 2007 *'
	'*******************************************************************'
	
	' Show form as modal, save settings if Save button is clicked on exit
	If Form1.ShowModal = 1 Then SaveSettings(Form1)

	'Set Form1 = Nothing	
End Sub
