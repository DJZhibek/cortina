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
Dim bSearchTitle : bSearchTitle = False			' search song title for 'cortina'?
Dim bSearchGenre : bSearchGenre = False			' search song genre for 'cortina'?
Dim bSearchPath : bSearchPath = False			' search song path for 'cortina'?
Dim bSearchAllCustomTags : bSearchAllCustomTags = False	' search song Custom1 tag for 'cortina'?
Dim iCortinaLen : iCortinaLen = 45			' default cortina length in seconds (includes fade-in and fade-out time)
Dim dCortinaVolume : dCortinaVolume = 0.7		' default cortina volume multiplier
Dim iFadeIn : iFadeIn = 1				' default fade-in time in seconds
Dim iFadeOut : iFadeOut = 5				' default fade-out time in seconds
Dim iGapTime : iGapTime = 3				' default gap time in seconds (additional silence added after cortinas and songs)
Dim dSongVolume : dSongVolume = 1.0			' storage for current playback volume (copied before cortina volume modifies it)


' cortina setting constants
Const cSecLabel = " sec"
Const cCortinaMin = 15
Const cCortinaMax = 240
Const cFadeInMin = 0
Const cFadeInMax = 10
Const cFadeOutMin = 0
Const cFadeOutMax = 15
Const cGapMin = 0
Const cGapMax = 10


' cortina state constants
Const cNone = 0
Const cFadeIn = 1
Const cFullVolume = 2
Const cFadeOut = 3
Const cGap = 4
Dim CortinaState : CortinaState = cNone	' keeps track of where we are in the cortina playback

' globals used during cortina playback
Dim iStateCounter : iStateCounter = 0	' timer counter used by all cortina states
Dim dVolumeInc : dVolumeInc = 1.0		' volume increment for fade-in and fade-out (calculated from cortina volume and fade times)
Dim bDoingCortina : bDoingCortina = False	' True is a cortina is currently being processed

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
	
	ToggleCrossfade(False)		' Turn off crossfade, backup current setting
	ReadSettings				' Read previously saved settings
	CreateTimers				' Create timers used by cortinas
	
	' Register MediaMonkey events we need to act on
	Script.RegisterEvent SDB, "OnPlay", "Event_OnPlay"
	Script.RegisterEvent SDB, "OnPause", "Event_OnPause"
	Script.RegisterEvent SDB, "OnStop", "Event_OnStop"
	Script.RegisterEvent SDB, "OnTrackEnd", "Event_TrackEnd"
	Script.RegisterEvent SDB, "OnShutdown", "Shutdown"
	
	dSongVolume = SDB.Player.Volume ' Save start up volume setting
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

' Disable timers
Sub DisableTimers()
	On Error Resume Next
		StateTimer.Enabled = False
		ProgressTimer.Enabled = False
	On Error GoTo 0
End Sub

' cleanup function for playback ending
Sub Cleanup()
	DisableTimers
	Set ProgressDisplay = Nothing
	CortinaState = cNone
End Sub

' cleanup function for MediaMonkey shutdown
Sub Shutdown()
	Cleanup
	On Error Resume Next
		Script.UnregisterAllEvents
		ToggleCrossfade(True)	' Restore old crossfade settings on shutdown
	On Error GoTo 0
End Sub

' tests current song to see if it is a cortina
Function Is_Cortina()
	Dim objSongData : Set objSongData = SDB.Player.CurrentSong
	Is_Cortina = False
	
	' test selected locations to see if the word "cortina" exists (not case sensitive)
	If bSearchGenre Then
		If Instr(1,objSongData.Genre,"cortina",1) > 0 Then 
			Is_Cortina = True
			Exit Function
		End If
	End If
	
	If bSearchTitle Then
		If Instr(1,objSongData.Title,"cortina",1) > 0 Then 
			Is_Cortina = True
			Exit Function
		End If
	End If

	If bSearchPath Then
		If Instr(1,objSongData.Path,"cortina",1) > 0 Then
			Is_Cortina = True
			Exit Function
		End If
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
	Dim iSongLength
	iSongLength = SDB.Player.CurrentSong.StopTime - SDB.Player.CurrentSong.StartTime ' Calculate real playtime
	iSongLength = CInt(iSongLength / 1000.0) ' convert current song length to seconds
	
	' subtract fade-in and fade-out times from cortina length
	If iSongLength > iCortinaLen Then
		FullVolumeLength = iCortinaLen - (iFadeIn + iFadeOut)
	Else
		FullVolumeLength = iSongLength - (iFadeIn + iFadeOut)
	End If
	
	' if this happens, we are really in trouble, cortina might end up longer than the song
	If FullVolumeLength < 1 Then FullVolumeLength = 0

End Function

' Calculate fade out volume at any point in time
Function FadeOutVolume(dFadeLength, dStartVolume, dTimePoint)
	FadeOutVolume = dStartVolume * ((cos((dTimePoint/dFadeLength) * 3.1415) + 1.0) / 2.0)
	If FadeOutVolume < 0.001 Then FadeOutVolume = 0.0
End Function


' play next song, if applicable
Sub GoToNextSong()
	If SDB.Player.CurrentSongIndex + 1 < SDB.Player.PlaylistCount Then
		If bDoingCortina Then SDB.Player.Next
		SDB.ProcessMessages
	End If
End Sub

' Get what the next state should be without changing the current state
Function GetNextState(curState)
	GetNextState = curState

	If GetNextState = cNone Then
		If iFadeIn > 0 Then
			GetNextState = cFadeIn
			Exit Function
		End If
		GetNextState = cFullVolume
	End If
	
	If GetNextState = cFadeIn Then
		If FullVolumeLength() > 0 Then
			GetNextState = cFullVolume
			Exit Function
		End If
		GetNextState = cFadeOut
	End If
		
	If GetNextState = cFullVolume Then	
		If iFadeOut > 0 Then
			GetNextState = cFadeOut
			Exit Function
		End If
		GetNextState = cGap
	End If

	If GetNextState = cFadeOut Then	
		GetNextState = cGap
		Exit Function
	End If

	If GetNextState = cGap Then	
		GetNextState = cNone
	End If
End Function

Sub SetupState(newState)
	DisableTimers
	
	Dim Player : Set Player = SDB.Player

	Select Case newState
		Case cNone
			CortinaState = cNone
			Exit Sub
			
		Case cFadeIn					' fade in timer is used, start it
			CortinaState = cFadeIn		' set cortina state to fade-in
			Player.Volume = 0.0
			iStateCounter = iFadeIn * 4	' convert seconds to quarter seconds		
			dVolumeInc = dCortinaVolume / CDbl(iStateCounter) ' calculate volume decrement value from current settings
			StateTimer.Interval = 250	' set timer interval to 1/4 second (250ms)
			On Error Resume Next
				If ProgressDisplay Is Nothing Then
					Set ProgressDisplay = SDB.Progress
					ProgressTimer.Interval = 1000 ' 1000 ms = 1 second
					ProgressDisplay.MaxValue = iCortinaLen
				End If
				ProgressDisplay.Text = "Cortina: " & iCortinaLen & " seconds left."
			On Error GoTo 0

		Case cFullVolume	' No fade in, set cortina volume and start full cortina volume timer
			CortinaState = cFullVolume
			Player.Volume = dCortinaVolume
			iStateCounter = FullVolumeLength()
			StateTimer.Interval = 1000	' set timer interval to 1 second for full volume part of cortina (1000ms)
			On Error Resume Next
				If ProgressDisplay Is Nothing Then
					Set ProgressDisplay = SDB.Progress
					ProgressTimer.Interval = 1000 ' 1000 ms = 1 second			
					ProgressDisplay.MaxValue = iCortinaLen
					ProgressDisplay.Value = 0
				End If
				ProgressDisplay.Text = "Cortina: " & ProgressDisplay.MaxValue - ProgressDisplay.Value & " seconds left."
			On Error GoTo 0
			
		Case cFadeOut
			CortinaState = cFadeOut		' set cortina state to fade-out
			Player.Volume = dCortinaVolume
			iStateCounter = iFadeOut * 4	' convert seconds to quarter seconds
			StateTimer.Interval = 250	' set timer interval to 1/4 second (250ms)
			On Error Resume Next
				If ProgressDisplay Is Nothing Then
					Set ProgressDisplay = SDB.Progress
					ProgressTimer.Interval = 1000 ' 1000 ms = 1 second			
					ProgressDisplay.MaxValue = iCortinaLen
					ProgressDisplay.Value = 0				
				End If
				ProgressDisplay.Text = "Cortina: " & ProgressDisplay.MaxValue - ProgressDisplay.Value & " seconds left."
			On Error GoTo 0
			
		Case cGap
			CortinaState = cGap
			iStateCounter = iGapTime * 4	' convert seconds to quarter seconds
			StateTimer.Interval = 250	' set timer interval to 1 second for silence gap (250ms)
			On Error Resume Next
				If ProgressDisplay Is Nothing Then
					Set ProgressDisplay = SDB.Progress
					ProgressTimer.Interval = 1000 ' 1000 ms = 1 second
				End If
				ProgressDisplay.MaxValue = iGapTime
				ProgressDisplay.Value = 0		
				ProgressDisplay.Text = "Silence Gap: " & iGapTime & " seconds left."
			On Error GoTo 0
			If bDoingCortina Then Player.Stop
	End Select		

	SDB.ProcessMessages
	ProgressTimer.Enabled = True	' start the cortina progress display timer
	StateTimer.Enabled = True	' start the interrupt timer

End Sub

' called when a song starts to play
Sub Event_OnPlay()
	' Make sure stop after current is enabled
	SDB.Player.StopAfterCurrent = True
		
	' Make sure progress display is not shown
	Set ProgressDisplay = Nothing
	
	' check if this song is a cortina
	bDoingCortina = Is_Cortina()
	If bDoingCortina = False Then 
		If CortinaState = cNone Then dSongVolume = SDB.Player.Volume  ' save current playback volume
		CortinaState = cNone
		Exit Sub
	End If
	CortinaState = cNone
	
	ReadSettings 					' get current settings

	' retrieve current song length
	Dim iSongLength
	iSongLength = CInt((SDB.Player.CurrentSong.StopTime - SDB.Player.CurrentSong.StartTime) / 1000.0) ' convert to seconds

	If iSongLength < iCortinaLen Then ' if song is shorter than cortina length use song length instead
		iCortinaLen = iSongLength
	End If	
	
	' calculate cortina volume
	dCortinaVolume = dSongVolume * dCortinaVolume
	
	SetupState(GetNextState(CortinaState))

End Sub

' Handle Pause button toggle
Sub Event_OnPause()
	' Re-enable timers if pause was release, otherwise disabled timers
	If SDB.Player.isPaused = False And SDB.Player.isPlaying = True And CortinaState <> cNone Then
		If IsObject(StateTimer) Then StateTimer.Enabled = True
		If IsObject(ProgressTimer) Then ProgressTimer.Enabled = True
	Else
		If IsObject(StateTimer) Then StateTimer.Enabled = False
		If IsObject(ProgressTimer) Then ProgressTimer.Enabled = False
	End If
	
End Sub

' Handle playback ending because end of track was reached during playback
Sub Event_TrackEnd()
	Dim Player : Set Player = SDB.Player
	
	' If last song already played then ignore this
	If Player.CurrentSongIndex + 1 < Player.PlaylistCount Then
		SetupState(GetNextState(cFadeOut))	' check if silence gap is set
	End If
End Sub

' Handle stop button being pressed or Player.Stop being called
Sub Event_OnStop()
	If CortinaState=cGap Then 
		GoToNextSong()
	Else
		Cleanup
	End If
End Sub

' Update the progress display (shows progress bar and text info while cortina or gap is playing)
Sub OnProgressTimer(thisTimer)
	
	On Error Resume Next ' Avoid errors caused by state timer event destroying ProgressDisplay before we can look at it	
		thisTimer.Enabled = False
		
		' check if progress display exists
		If ProgressDisplay Is Nothing Or isObject(ProgressDisplay)<>True Then 
			Set ProgressDisplay = Nothing
			On Error GoTo 0
			Exit Sub
		End If
		If ProgressDisplay.Terminate = True Or ProgressDisplay.Value >= ProgressDisplay.MaxValue Then 
			Set ProgressDisplay = Nothing
			On Error GoTo 0
			Exit Sub
		End If
		
			
		Dim iTimeLeft : iTimeLeft = ProgressDisplay.MaxValue - ProgressDisplay.Value - 1
		
		If 0 > iTimeLeft Then
			Set ProgressDisplay = Nothing
			On Error GoTo 0
			Exit Sub
		End If

		If CortinaState = cGap Then
			ProgressDisplay.Text = "Silence Gap: " & iTimeLeft & " seconds left."
		Else
			ProgressDisplay.Text = "Cortina: " & iTimeLeft & " seconds left."
		End If
		
		SDB.ProcessMessages
		ProgressDisplay.Increase
		thisTimer.Enabled = True		
	On Error GoTo 0
End Sub

' Handle Volume Timer event
Sub OnStateTimer(thisTimer)
	If thisTimer.Enabled = False Then Exit Sub
	thisTimer.Enabled = False ' Stop timer to avoid triggering while we are here
	
	Dim Player : Set Player = SDB.Player
	
	' Act on current state of cortina
	Select Case CortinaState
		Case cFadeIn		' handle fade in timer interrupt
			If iStateCounter > 0 Then
				' increment the fade-in volume
				If Player.Volume + dVolumeInc < dCortinaVolume Then
					Player.Volume = Player.Volume + dVolumeInc
				Else
					Player.Volume = dCortinaVolume
				End If
				iStateCounter = iStateCounter - 1
				thisTimer.Enabled = True	' re-start the volume interrupt timer
				Exit Sub					' and exit
			End If
			SetupState(GetNextState(CortinaState)) ' Fade In is done, setup next state and exit
			Exit Sub
	

		Case cFullVolume	' handle full volume portion of cortina
			If iStateCounter > 0 Then
				iStateCounter = iStateCounter - 1	' decrement length counter
				thisTimer.Enabled = True			' re-start the volume interrupt timer
				Exit Sub
			End If
			
			SetupState(GetNextState(CortinaState)) ' Full Volume is done, setup next state and exit
			Exit Sub
			
		Case cFadeOut		' handle fade-out timer interrupt	
			If iStateCounter > 0 Then
				' still fading out, lower volume
				If Player.Volume > 0.0 Then
					Player.Volume = FadeOutVolume(CDbl(iFadeOut * 4), dCortinaVolume, CDbl((iFadeOut * 4)-iStateCounter))
				Else
					Player.Volume = 0.0
				End If
				iStateCounter = iStateCounter - 1
				thisTimer.Enabled = True		' re-start the interrupt timer
				Exit Sub
			End If
			SetupState(GetNextState(CortinaState)) ' Fade Out is done, setup next state and exit
			Exit Sub
			
		Case cGap		' handle silence gap timer interrupt
			If iStateCounter > 0 Then
				iStateCounter = iStateCounter - 1	' decrement silence gap counter
				thisTimer.Enabled = True		' re-start the interrupt timer
				Exit Sub
			End If
			
	End Select

	CortinaState = cNone		' silence gap is finished, set state to none
	DisableTimers
	If bDoingCortina <> True Then GoToNextSong

	Player.Volume = dSongVolume
	Player.Play
End Sub

Sub ToggleCrossfade(iOnOff)
	
	Dim Player : Set Player = SDB.Player
	Dim Reg : Set Reg = SDB.Registry
	If Reg.OpenKey(cRegKey, True) Then
		If iOnOff = False Then
			Reg.BoolValue("CrossfadeState") = Player.IsCrossfade
			Player.IsCrossfade = False
		Else
			Player.IsCrossfade = Reg.BoolValue("CrossfadeState")
		End If
		Reg.CloseKey
	End If
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
	Label1.Caption = obj.Value & cSecLabel
End Sub

' update fade-in display when trackbar is changed
Sub fadein_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_tbCurFadeIn")
	Label1.Caption = obj.Value & cSecLabel
End Sub

' update fade-out display when trackbar is changed
Sub fadeout_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_tbCurFadeOut")
	Label1.Caption = obj.Value & cSecLabel
End Sub

' update silence gap length display when trackbar is changed
Sub gap_change(obj)
	Dim Form1 : Set Form1 = obj.Common.TopParent.Common
	Dim Label1 : Set Label1 = Form1.ChildControl("lbl_tbCurGap")
	Label1.Caption = obj.Value & cSecLabel
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
	'* Form produced by MMVBS Form Creator                             *'
	'* (https://code.google.com/archive/p/mmvbs/)                      *'
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
	TrackBar1.MaxValue = cCortinaMax
	TrackBar1.MinValue = cCortinaMin
	TrackBar1.Value = iCortinaLen
	TrackBar1.Common.SetRect 16,60,450,45
	TrackBar1.Common.ControlName = "tb_CortinaLen"
	TrackBar1.Common.Hint = "Cortina length"
	Script.RegisterEvent TrackBar1, "OnChange", "length_change"

	
	Dim Label2 : Set Label2 = SDB.UI.NewLabel(Form1)
	Label2.Autosize = False
	Label2.Common.SetRect 15,105,65,17
	Label2.Caption = cCortinaMin & cSecLabel
	
	Dim Label3 : Set Label3 = SDB.UI.NewLabel(Form1)
	Label3.Common.SetRect 432,105,65,17
	Label3.Caption = cCortinaMax & cSecLabel
	
	Dim Label4 : Set Label4 = SDB.UI.NewLabel(Form1)
	Label4.Common.SetRect 182,105,70,17
	Label4.Caption = "Cortina Length:"
	
	Dim Label5 : Set Label5 = SDB.UI.NewLabel(Form1)
	Label5.Common.SetRect 264,105,80,17
	Label5.Common.ControlName = "lbl_tbCurLen"
	Label5.Caption = iCortinaLen & cSecLabel

	
	Dim TrackBar2 : Set TrackBar2 = SDB.UI.NewTrackBar(Form1)
	TrackBar2.MaxValue = 10
	TrackBar2.Value = iFadeIn
	TrackBar2.Common.SetRect 16,147,450,45
	TrackBar2.Common.ControlName = "tb_FadeIn"
	TrackBar2.Common.Hint = "Cortina Fade In Time"
	Script.RegisterEvent TrackBar2, "OnChange", "fadein_change"
	
	Dim Label6 : Set Label6 = SDB.UI.NewLabel(Form1)
	Label6.Common.SetRect 16,190,65,17
	Label6.Caption = cFadeInMin & cSecLabel
	
	Dim Label7 : Set Label7 = SDB.UI.NewLabel(Form1)
	Label7.Common.SetRect 191,190,65,17
	Label7.Caption = "Fade In Time:"
	
	Dim Label8 : Set Label8 = SDB.UI.NewLabel(Form1)
	Label8.Common.SetRect 264,190,65,17
	Label8.Common.ControlName = "lbl_tbCurFadeIn"
	Label8.Caption = iFadeIn & cSecLabel
	
	Dim Label9 : Set Label9 = SDB.UI.NewLabel(Form1)
	Label9.Common.SetRect 431,190,65,17
	Label9.Caption = cFadeInMax & cSecLabel
	
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
	Label11.Caption = cFadeOutMin & cSecLabel
	
	Dim Label12 : Set Label12 = SDB.UI.NewLabel(Form1)
	Label12.Common.SetRect 430,275,85,17
	Label12.Caption = cFadeOutMax & cSecLabel
	
	Dim Label13 : Set Label13 = SDB.UI.NewLabel(Form1)
	Label13.Common.SetRect 264,275,65,17
	Label13.Common.ControlName = "lbl_tbCurFadeOut"
	Label13.Caption = iFadeOut & cSecLabel
	
	Dim TrackBar4 : Set TrackBar4 = SDB.UI.NewTrackBar(Form1)
	TrackBar4.MaxValue = 10
	TrackBar4.MinValue = 0
	TrackBar4.Value = iGapTime
	TrackBar4.Common.SetRect 16,317,450,45
	TrackBar4.Common.ControlName = "tb_GapTime"
	TrackBar4.Common.Hint = "Additional gap of silence after cortina"
	Script.RegisterEvent TrackBar4, "OnChange", "gap_change"
	
	Dim Label14 : Set Label14 = SDB.UI.NewLabel(Form1)
	Label14.Common.SetRect 17,359,65,17
	Label14.Caption = cGapMin & cSecLabel
	
	Dim Label15 : Set Label15 = SDB.UI.NewLabel(Form1)
	Label15.Common.SetRect 208,359,65,17
	Label15.Caption = "Gap Time:"
	
	Dim Label16 : Set Label16 = SDB.UI.NewLabel(Form1)
	Label16.Common.SetRect 264,359,80,17
	Label16.Common.ControlName = "lbl_tbCurGap"
	Label16.Caption = iGapTime & cSecLabel
	
	Dim Label17 : Set Label17 = SDB.UI.NewLabel(Form1)
	Label17.Common.SetRect 437,359,65,17
	Label17.Caption = cGapMax & cSecLabel
	
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
	
	' Show form as modal, save settings if Save button is clicked on exit
	If Form1.ShowModal = 1 Then SaveSettings(Form1)

	'Set Form1 = Nothing	
End Sub
