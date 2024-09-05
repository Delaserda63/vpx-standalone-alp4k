' ****************************************************************
'             VP_Cooks THE MATRIX VISUAL PINBALL X 10.7.Final
'      Including JP's Arcade Physics 3.0.1 - Tnx to JP Salas for all the kind help!
'		
'		PuPPack pupevents added by Enthusiast
'		nFozzy/Rothbauerw physics conversion by MerlinRTP
'		Thanks to VP Cooks for permissions to work with the table script!
'		VR Room by DaRdog81 for the room and Mike DA Spike for the hybrid scripting and menu Options_KeyDown
'          NOTE : USE LEFT AND RIGHT MAGNA FOR MENU OPTIONS
                 
            
' ****************************************************************

'On the DOF website put THE MATRIX
'DOF commands:
'101 Left Flipper
'102 Right Flipper
'103 left slingshot
'104 right slingshot
'105
'106
'107
'108 RIGHT Bumper
'109
'110
'111 Scoop
'118
'119
'120 AutoFire
'122 knocker
'123 ballrelease
'204 Lamp posts (teenagers)
'205 Blue targets
'206 Ramps, lanes and spinners
'207 In and outlanes
'208 Plunger

Option Explicit 

'*************************** PuP Settings for this table ********************************

'****** PuP Variables ******

Dim usePUP: Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!

usePUP = True
cPuPPack = "TheMatrix"    ' name of the PuP-Pack / PuPVideos folder for this table

'****** VR Options ****** DO NOT MODIFY NEXT LINE., PRESS LEFT AND RIGHT MAGNA DURING ATTRACT MODE TO CHANGE OPTIONS
Dim VRRoomChoice : VRRoomChoice = 1				
Dim VRTest : Vrtest = False


'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

' This is used for the startup and control of Pinup Player

Sub PuPStart(cPuPPack)
    If PUPStatus=true then Exit Sub
    If usePUP=true then
        Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
        If PuPlayer is Nothing Then
            usePUP=false
            PUPStatus=false
        Else
            PuPlayer.B2SInit "",cPuPPack 'start the Pup-Pack
            PUPStatus=true
        End If
    End If
End Sub

Sub pupevent(EventNum)
    if (usePUP=false or PUPStatus=false) then Exit Sub
    PuPlayer.B2SData "E"&EventNum,1  'send event to Pup-Pack
End Sub

' ******* How to use PUPEvent to trigger / control a PuP-Pack *******

' Usage: pupevent(EventNum)

' EventNum = PuP Exxx trigger from the PuP-Pack

' Example: pupevent 102

' This will trigger E102 from the table's PuP-Pack

' DO NOT use any Exxx triggers already used for DOF (if used) to avoid any possible confusion

'**********************************************APRON LEFT***************************

Randomize

Const BallSize = 50        ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1       ' standard ball mass in JP's VPX Physics 3.0.1
Const SongVolume = 0.3     ' 1 is full volume, but I set it quite low to listen better the other sounds since I use headphones, adjust to your setup :)
Const FlippersBlood = False 'set it to false if you don't like that

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
End Sub

'//////////////////////////////////////////////////////////////////////
dim ScorbitActive
ScorbitActive					= 0 	' Is Scorbit Active	
Const     ScorbitShowClaimQR	= 1 	' If Scorbit is active this will show a QR Code  on ball 1 that allows player to claim the active player from the app
Const     ScorbitQRLeft		= 0 	' Make Claim QR Code show in lower left (0)  or lower right (1)
Const     ScorbitUploadLog		= 0 	' Store local log and upload after the game is over 
Const     ScorbitAlternateUUID  = 0 	' Force Alternate UUID from Windows Machine and saves it in VPX Users directory (C:\Visual Pinball\User\ScorbitUUID.dat)
'/////////////////////////////////////////////////////////////////////


' Define any Constants
Const cGameName = "TheMatrix"
Const myVersion = "1.01"
Const BallSaverTime = 10     ' in seconds of the first ball
Const MaxMultiplier = 5      ' limit playfield multiplier
Const MaxBonusMultiplier = 5 'limit Bonus multiplier
Const BallsPerGame = 3     ' usually 3 or 5
Const MaxMultiballs = 13     ' max number of balls during multiballs
Dim MaxPlayers : MaxPlayers = 4 ' from 1 to 4

'----- VR Room Auto-Detect -----
Dim VR_Obj, VR_Room, VRMode

' Use FlexDMD if in FS mode
Dim UseFlexDMD
If Table1.ShowDT = True then
    UseFlexDMD = False
Else
    UseFlexDMD = True
End If
'FlexDMD in high or normal quality
'change it to True if you have an LCD screen, 256x64
'or keep it False if you have a real DMD at 128x32 in size
DIM FlexDMDHighQuality : FlexDMDHighQuality = True

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim PlayfieldMultiplier(4)
Dim PFxSeconds
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot(4)
Dim SuperJackpot(4)
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim SuperSkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode
Dim x
Dim bFirstBall(4)
Dim bOnTheFirstBallScorbit

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock(4)
Dim BallsInHole

' Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs
Dim mMagnet
Dim cbLeft    'captive ball at the magnet

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 0.5        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFXDOF("fx_kicker", 141, DOFPulse, DOFContactors), SoundFXDOF("fx_solenoid", 141, DOFPulse, DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Magnet
    Set mMagnet = New cvpmMagnet
    With mMagnet
        .InitMagnet Magnet, 35
        .GrabCenter = True
        .CreateEvents "mMagnet"
    End With

    Set cbLeft = New cvpmCaptiveBall
    With cbLeft
        .InitCaptive CapTrigger, MagnetPost, CapKicker, 0
        .ForceTrans = .7
        .MinForce = 3.5
        .CreateEvents "cbLeft"
        .Start
    End With

    'Load Menu and settings
	Options_Load
	'************ PuP-Pack Startup **************

	PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack


    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Credits = 0
    Loadhs

    ' Initalise the DMD display
    DMD_Init


    if bFreePlay or Credits > 1 Then DOF 125, DOFOn

    ' Init main variables and any other flags
    bAttractMode = False
    bOnTheFirstBall = False
	bOnTheFirstBallScorbit = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    PFxSeconds = 0
    bGameInPlay = False
    bAutoPlunger = False
    bMusicOn = True
    BallsOnPlayfield = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJackpot = False
    bInstantInfo = False
    ' set any lights for the attract mode
    GiOff
	'pupevent 798
    StartAttractMode

    ' Start the RealTime timer
    RealTime.Enabled = 1


	If RenderingMode = 2 or Table1.ShowFSS or VRTest Then 'VR or FSS hide all that is not needed :D 
				
		'disable table objects that should not be visible
		ramp103.visible = False
		Wall8.visible = False
		Wall8.SideVisible = False
		Wall9.visible = False
		Wall9.SideVisible = False
		rpeg009.visible = False
		rpeg009.SideVisible = False
		rrail.visible = False
		lrail.visible = False

		'External DMD not used as we use the internal DMD 
		UseFlexDMD = false		'already set during launch of table
		
		'Move DMD to correct location
		digitgrid.rotx = - 86
		digitgrid.x = digitgrid.x + 1204
		digitgrid.y = digitgrid.y - 310 
		digitgrid.height = digitgrid.height + 284 
		digit041.rotx = -86
		digit041.x = digit041.x + 1204
		digit041.y = digit041.y - 310 
		digit041.height = digit041.height + 284 
		Dim VrObj
		For Each VrObj in DMDUpper
			VrObj.rotx = -86
			VrObj.x = VrObj.x +1204 
			VrObj.y = VrObj.y - 295
			VrObj.height = VrObj.height + 301 
		Next
		For Each VrObj in DMDLower
			VrObj.rotx = -86
			VrObj.x = VrObj.x + 1204
			VrObj.y = VrObj.y - 334 
			VrObj.height = VrObj.height + 276 
		Next

		If usepup Then
			VR_Backglass.visible = False
			VR_PupBackglass.visible = True
			VR_TopperScreen.visible = True
			VR_TopperTV.visible = True
			VRPupTopper.VideoCapWidth=400	
			VRPupTopper.VideoCapHeight=100
			VRPupTopper.visible=true	
			VRPupTopper.TimerEnabled=true
			VRPupTopper.TimerInterval=60
		Else
			VR_Backglass.visible = True
			VR_PupBackglass.visible = False
			VRPupTopper.visible=False
			VR_TopperScreen.visible = False
			VR_TopperTV.visible = False
		End If

		'Weird blue light. Move it to sentinal
		flasher008.rotx = 0
		flasher008.roty = 0
		flasher008.x = 758
		flasher008.y = 930
		flasher008.height = 128
	end if
	if ScorbitActive = 1 and usePUP Then PUPInit  'this should be called in table1_init at bottom after all else b2s/controller running
End Sub 

    ' KICKER RAMPA MATRIX ''

Sub KickerRampaMatrix_Hit
	pupevent 823
    PlaySound "sfx_rampamatrix"

	vpmtimer.addtimer 2500,"KickerRampaMatrix.kick 10, 325 '"
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)


    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound SoundFX("fx_nudge",0), 0, 1, 1, 0.25

    If keycode = MechanicalTilt Then PlaySound SoundFX("fx_nudge",0), 0, 1, 1, 0.25

	If RenderingMode = 2 or Table1.ShowFSS or VRTest Then 
		If keyCode=LeftFlipperKey Then :FlipperButtonLeft.X = FlipperButtonLeft.X + 8: 	
		If keyCode=RightFlipperKey Then  FlipperButtonRight.X = FlipperButtonRight.X - 8
		If keycode = StartGameKey then Primary_StartButton.y = Primary_StartButton.y - 6 
		'If keycode = LeftMagnaSave Then bLutActive = True
	End if
    'If keycode = RightMagnaSave Then
    '    If bLutActive Then
    '        NextLUT
    '    End If
    'End If



    If Keycode = AddCreditKey Then
        Credits = Credits + 1
        if bFreePlay = False Then DOF 125, DOFOn
        If(Tilted = False) Then
            DMDFlush
            DMD "", CL("CREDITS " & Credits), "", eNone, eNone, eNone, 500, True, "fx_coin"
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
		TimerVRPlunger.enabled = true 
		TimerVRPlunger2.enabled = False
    End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

		If keycode = LeftTiltKey Then CheckTilt 'only check the tilt during game
		If keycode = RightTiltKey Then CheckTilt
		If keycode = CenterTiltKey Then CheckTilt

		If keycode = MechanicalTilt Then CheckTilt

		If keycode = LeftFlipperKey Then SolLFlipper 1:InstantInfoTimer.Enabled = True:RotateLaneLights 1
		If keycode = RightFlipperKey Then SolRFlipper 1:InstantInfoTimer.Enabled = True:RotateLaneLights 0

		If keycode = StartGameKey Then
			If((PlayersPlayingGame < MaxPlayers) AND(bOnTheFirstBall = True) ) Then

				If(bFreePlay = True) Then
					PlayersPlayingGame = PlayersPlayingGame + 1
					TotalGamesPlayed = TotalGamesPlayed + 1
					DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
				Else
					If(Credits > 0) then
						PlayersPlayingGame = PlayersPlayingGame + 1
						TotalGamesPlayed = TotalGamesPlayed + 1
						Credits = Credits - 1
						DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
						If Credits < 1 And bFreePlay = False Then DOF 125, DOFOff
					Else
							' Not Enough Credits to start a game.
							DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, "vo_givemeyourmoney"
					End If
				End If
			End If
		End If


   Else ' If (GameInPlay)

		If keycode = StartGameKey Then
			If(bFreePlay = True) Then
				If(BallsOnPlayfield = 0) Then
					ResetForNewGame()
				End If
			Else
				If(Credits > 0) Then
					If(BallsOnPlayfield = 0) Then
						Credits = Credits - 1
						If Credits < 1 And bFreePlay = False Then DOF 125, DOFOff
						ResetForNewGame()
					End If
				Else
					' Not Enough Credits to start a game.
					DMDFlush
					DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, "vo_givemeyourmoney"
					ShowTableInfo
				End If
			End If
		End If
		If bInOptions Then
			Options_KeyDown keycode
			Exit Sub
		End If
		If keycode = LeftMagnaSave  Then
			If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
		ElseIf keycode = RightMagnaSave  Then
			If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
		End If

    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

    'If keycode = LeftMagnaSave Then bLutActive = False

	If RenderingMode = 2 or Table1.ShowFSS or VRTest Then 
		If keyCode=LeftFlipperKey Then :FlipperButtonLeft.X = FlipperButtonLeft.X - 8: 	
		If keyCode=RightFlipperKey Then  FlipperButtonRight.X = FlipperButtonRight.X + 8 
		If keycode = StartGameKey then Primary_StartButton.y = Primary_StartButton.y + 6
	End If
    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
		TimerVRPlunger.enabled = False 
		TimerVRPlunger2.enabled = True
    End If

	If keycode = LeftMagnaSave And Not bInOptions Then bOptionsMagna = False
    If keycode = RightMagnaSave And Not bInOptions Then bOptionsMagna = False

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
        If keycode = RightFlipperKey Then
            SolRFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
    End If



End Sub




Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    If NOT hsbModeActive Then
        bInstantInfo = True
        DMDFlush
        InstantInfo
    End If
End Sub

'*************
' Pause Table
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub Table1_Exit
    Savehs
	if UsePuP and (RenderingMode = 2 or Table1.ShowFSS or vrtest) then VRPupTopper.TimerEnabled=false ': VRPupTopper.VideoCapUpdate = nothing
	if UsePuP and (RenderingMode = 2 or Table1.ShowFSS or vrtest) then VRPupTopper.VideoCapUpdate = "" : SET PuPlayer = Nothing
    If UseFlexDMD Then FlexDMD.Run = False
    If B2SOn = true Then Controller.Stop
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled AND bFlippersEnabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.EOSTorque = 0.75:LeftFlipper.RotateToEnd
        LeftFlipper001.EOSTorque = 0.75:LeftFlipper001.RotateToEnd
        If FlippersBlood Then LeftSplat
        Else
            PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
            LeftFlipper.EOSTorque = 0.2:LeftFlipper.RotateToStart
            LeftFlipper001.EOSTorque = 0.2:LeftFlipper001.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled AND bFlippersEnabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.EOSTorque = 0.75:RightFlipper.RotateToEnd
        If FlippersBlood Then RightSplat
        Else
            PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
            RightFlipper.EOSTorque = 0.2:RightFlipper.RotateToStart
    End If
End Sub

' flippers hit Sound
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable

Sub LeftFlipper_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub LeftFlipper001_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch Activeball, RightFlipper, RFCount, parm
	RightFlipperCollide parm
End Sub

Sub LeftFlipperCollide(parm)
        FlipperLeftHitParm = parm/10
        If FlipperLeftHitParm > 1 Then
                FlipperLeftHitParm = 1
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
        RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
        FlipperRightHitParm = parm/10
        If FlipperRightHitParm > 1 Then
                FlipperRightHitParm = 1
        End If
        FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
        RandomSoundRubberFlipper(parm)
End Sub


'--------

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub


Sub RandomSoundRubberFlipper(parm)
        PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub
Dim RSplat, LSplat

Sub RightSplat
    RSplat = 0
    Rightblood_Timer
End Sub

Sub Rightblood_Timer
    Select Case RSplat
        Case 0:Rightblood.ImageA = "blood1":Rightblood.Visible = 1:Rightblood.TimerEnabled = 1
        Case 1:Rightblood.ImageA = "blood2"
        Case 2:Rightblood.ImageA = "blood3"
        Case 3:Rightblood.ImageA = "blood4"
        Case 4:Rightblood.ImageA = "blood5"
        Case 5:Rightblood.ImageA = "blood6"
        Case 6:Rightblood.Visible = 0:Rightblood.TimerEnabled = 0
    End Select
    RSplat = RSplat + 1
End Sub

Sub LeftSplat
    LSplat = 0
    Leftblood_Timer
End Sub

Sub Leftblood_Timer
    Select Case LSplat
        Case 0:Leftblood.ImageA = "blood1a":Leftblood.Visible = 1:Leftblood.TimerEnabled = 1
        Case 1:Leftblood.ImageA = "blood2a"
        Case 2:Leftblood.ImageA = "blood3a"
        Case 3:Leftblood.ImageA = "blood4a"
        Case 4:Leftblood.ImageA = "blood5a"
        Case 5:Leftblood.ImageA = "blood6a"
        Case 6:Leftblood.Visible = 0:Leftblood.TimerEnabled = 0
    End Select
    LSplat = LSplat + 1
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                       'Called when table is nudged
    Dim BOT
    BOT = GetBalls
    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub
    Tilt = Tilt + TiltSensitivity                   'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity) AND(Tilt <= 15) Then 'show a warning
        DMD "_", CL("CAREFUL NEO"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
    End if
    If(NOT Tilted) AND Tilt > 15 Then 'If more that 15 then TILT the table
        'display Tilt
        InstantInfoTimer.Enabled = False
        DMDFlush
		pupevent 928 'topper
		pupevent 927 'lady in red
        DMD CL("YOU"), CL("TILTED"), "", eNone, eNone, eNone, 200, False, ""
        'PlaySound "vo_yousuck" &RndNbr(5)
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
        StopMBmodes
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        Tilted = True
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        LeftFlipper001.RotateToStart
        RightFlipper.RotateToStart
        Bumper1.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        Tilted = False
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper1.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0) Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        vpmtimer.Addtimer 2000, "EndOfBall() '"
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'*****************************************
'         Music as wav sounds
' in VPX 10.7 you may use also mp3 or ogg
'*****************************************

Dim Song
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            PlaySound Song, -1, SongVolume
        End If
    End If
End Sub

Sub ChangeSong
    If bJasonMBStarted Then
        PlaySong "BBmu_multiball1"
    ElseIf bFreddyMBStarted Then
        PlaySong "BBmu_multiball3"
    ElseIf bMichaelMBStarted Then
        PlaySong "BBmu_multiball2"
    ElseIf Mode(CurrentPLayer, 0) Then
        PlaySong "BBmu_pursuit"
    Else
        PlaySong "BBmu_main" &Balls
    End If
End Sub

Sub StopSong(name)
    StopSound name
End Sub

'********************
' Play random quotes
'********************

Sub PlayQuote 'Jason's mom
    PlaySound "vo_mother" &RndNbr(45)
End Sub

Sub PlayHighScoreQuote
    Select Case RndNbr(3)
        Case 1:PlaySound "vo_awesomescore"
        Case 2:PlaySound "vo_excellentscore"
        Case 3:PlaySound "vo_greatscore"
    End Select
End Sub

Sub PlayNotGoodScore
    Select Case RndNbr(4)
        Case 1:PlaySound "vo_heywhathappened"
        Case 2:PlaySound "vo_didyouscoreanypoints"
        Case 3:PlaySound "vo_youneedflipperskills"
        Case 4:PlaySound "vo_youmissedeverything"
        Case 5:PlaySound "vo_thatwasprettybad"
        Case 6:PlaySound "vo_thatwasprettybad2"
    End Select
End Sub

Sub PlayEndQuote
    Select Case RndNbr(9)
        Case 1:PlaySound "vo_hahaha1"
        Case 2:PlaySound "vo_hahaha2"
        Case 3:PlaySound "vo_hahaha3"
        Case 4:PlaySound "vo_hahaha4"
        Case 5:PlaySound "vo_hastalaviatababy"
        Case 6:PlaySound "vo_hastalaviatababy2"
        Case 7:PlaySound "vo_Illbeback"
        Case 8:PlaySound "vo_seeyoulater"
        Case 9:PlaySound "vo_youredonebyebye"
    End Select
End Sub

Sub PlayThunder
    PlaySound "sfx_thunder" &RndNbr(7)
End Sub

Sub PlaySword
    PlaySound "sfx_sword" &RndNbr(5)
End Sub

Sub PlayKill
    PlaySound "sfx_kill" &RndNbr(10)
End Sub

Sub PlayElectro
    PlaySound "sfx_electro" &RndNbr(9)
End Sub

'**********************
'     GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************

Dim OldGiState
OldGiState = -1   'start witht the Gi off

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub ChangeGIIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGILights
        bulb.IntensityScale = factor
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 0 Then '-1 means no balls, 0 is the first captive ball, 1 is the second captive ball...)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    PlaySoundAt "fx_GiOn", li036 'about the center of the table
    DOF 118, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    PlaySoundAt "fx_GiOff", li036 'about the center of the table
    DOF 118, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

' GI, light & flashers sequence effects

Sub GiEffect(n)
    Dim ii
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 40
            LightSeqGi.Play SeqBlinking, , 15, 25
        Case 2 'random
            LightSeqGi.UpdateInterval = 25
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqBlinking, , 10, 20
        Case 4 'seq up
            LightSeqGi.UpdateInterval = 3
            LightSeqGi.Play SeqUpOn, 25, 3
        Case 5 'seq down
            LightSeqGi.UpdateInterval = 3
            LightSeqGi.Play SeqDownOn, 25, 3
    End Select
End Sub

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 40
            LightSeqInserts.Play SeqBlinking, , 15, 25
        Case 2 'random
            LightSeqInserts.UpdateInterval = 25
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 20
            LightSeqInserts.Play SeqBlinking, , 10, 10
        Case 4 'center - used in the bonus count
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqCircleOutOn, 15, 2
        Case 5 'top down
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 15, 2
        Case 6 'down to top
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 15, 1
        Case 7 'center from the magnet
            LightSeqMG.UpdateInterval = 4
            LightSeqMG.Play SeqCircleOutOn, 15, 1
    End Select
End Sub

Sub FlashEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqFlashers.Play SeqAlloff
        Case 1 'all blink
            LightSeqFlashers.UpdateInterval = 40
            LightSeqFlashers.Play SeqBlinking, , 15, 25
        Case 2 'random
            LightSeqFlashers.UpdateInterval = 25
            LightSeqFlashers.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqFlashers.UpdateInterval = 20
            LightSeqFlashers.Play SeqBlinking, , 10, 20
        Case 4 'center
            LightSeqFlashers.UpdateInterval = 4
            LightSeqFlashers.Play SeqCircleOutOn, 15, 2
        Case 5 'top down
            LightSeqFlashers.UpdateInterval = 4
            LightSeqFlashers.Play SeqDownOn, 15, 1
        Case 6 'down to top
            LightSeqFlashers.UpdateInterval = 4
            LightSeqFlashers.Play SeqUpOn, 15, 1
        Case 7 'top flashers left right
            LightSeqTopFlashers.UpdateInterval = 10
            LightSeqTopFlashers.Play SeqRightOn, 50, 10
    End Select
End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v3.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.4, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v3.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls, 20 balls, from 0 to 19
Const lob = 1     'number of locked balls
Const maxvel = 40 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
        aBallShadow(b).Y = 1500
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z -24

        If BallVel(BOT(b) ) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 10
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        End If

        ' jps ball speed control
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero

Const VolumeDial = 0.8

Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim snd
	Select Case Int(Rnd * 7) + 1
		Case 1
			snd = "Ball_Collide_1"
		Case 2
			snd = "Ball_Collide_2"
		Case 3
			snd = "Ball_Collide_3"
		Case 4
			snd = "Ball_Collide_4"
		Case 5
			snd = "Ball_Collide_5"
		Case 6
			snd = "Ball_Collide_6"
		Case 7
			snd = "Ball_Collide_7"
	End Select
	
	PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)

	FlipperCradleCollision ball1, ball2, velocity

End Sub

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / table1.width-1

        if tmp > 7000 Then
                tmp = 7000
        elseif tmp < -7000 Then
                tmp = -7000
        end if

    If tmp> 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10))
    End If
End Function'************************************
' Diverse Collection Hit Sounds v3.0
'************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

'extra collections in this table
Sub aBlueRubbers_Hit(idx)
    Select Case RndNbr(15)
        Case 1:Playsound "vo_hahaha1"
        Case 2:Playsound "vo_hahaha2"
        Case 3:Playsound "vo_hahaha3"
        Case 4:Playsound "vo_toobusytoaim"
        Case 5:Playsound "vo_youmissedeverything"
        Case 6:Playsound "vo_yousuck1"
        Case 7:Playsound "vo_yousuck2"
    End Select
End Sub

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True

	if ScorbitActive = 1 And (Scorbit.bNeedsPairing) = False Then 
		Scorbit.StartSession()
	End If

    'resets the score display, and turn off attract mode
    StopAttractMode
    GiOn
	
    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
	bOnTheFirstBallScorbit = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        PlayfieldMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

    ' initialise any other flags
    Tilt = 0

    ' initialise specific Game variables
    Game_Init()

    ' you may wish to start some music, play a sound, do whatever at this point

    vpmtimer.addtimer 1500, "FirstBall '"
End Sub

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with

Sub FirstBall
    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' show neo letter
    TimerParaPc.enabled = 1
    DarkNight.Visible = 1
    ' a ball will be created at the end of the text
    ' create a new ball in the shooters lane
    ' CreateNewBall()
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
    ' make sure the correct display is upto date
    DMDScoreNow

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1

    ' reduce the playfield multiplier
    DecreasePlayfieldMultiplier

    ' reset any drop targets, lights, game Mode etc..

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False

    'Reset any table specific
    ResetNewBallVariables

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    'and the skillshot
    bSkillShotReady = True

'MerlinRTP added code
	if bFirstBall(CurrentPlayer) = False Then
		pupevent 801
		pupevent 833
	End IF

	bFirstBall(CurrentPlayer) = False
'Change the music ?
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass 
	
    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease 
    BallRelease.Kick 90, 4
	'pupevent 797  ' shuts off attractmode video if needed
	

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield > 1 Then
        DOF 143, DOFPulse
        bMultiBallMode = True
        bAutoPlunger = True
    End If
End Sub

' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
    'and eject the first ball
    CreateMultiballTimer_Timer
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If BallsOnPlayfield < MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                CreateMultiballTimer.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            CreateMultiballTimer.Enabled = False
			
        End If
    End If
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded

Sub EndOfBall()
    Dim AwardPoints, TotalBonus, ii
'	pupevent 800
'	pupevent 832
    AwardPoints = 0
    TotalBonus = 10 'yes 10 points :)
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False 
	

    ' only process any of this if the table is not tilted.
    '(the tilt recovery mechanism will handle any extra balls or end of game)
	
    If NOT Tilted Then
        StopSong Song
        PlaySound "BBsfx_suspense" 
		
        'Count the bonus. This table uses several bonus
        DMD CL("BONUS"), "", "", eNone, eNone, eNone, 1000, True, ""

        'Weapons collected / Programs Hacked X 1.500.000
        AwardPoints = Weapons(CurrentPlayer) * 1500000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL("HACKED PROGRAMS " & Weapons(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, "mu_bonus"

        'Counselors killed / Completed Missions X 750.000
        AwardPoints = CounselorsKilled(CurrentPlayer) * 750000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL("COMPLETED MISSIONS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, "mu_bonus"

        'Teenagers killed / Clones Killed x 300.000
        AwardPoints = TeensKilled(CurrentPlayer) * 300000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL("SMITH CLONES KILLED"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, "mu_bonus"

        'Loops X 150.000
        AwardPoints = LoopHits(CurrentPlayer) * 150000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL("LOOP COMBOS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, "mu_bonus"

        'Combos X 150.000
        AwardPoints = ComboHits(CurrentPlayer) * 150000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL("RAMP COMBOS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, "mu_bonus"

        'Bumpers X 50.000
        AwardPoints = BumperHits(CurrentPlayer) * 50000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL("BUMPER HITS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, "mu_bonus"

        If TotalBonus > 5000000 Then
            DMD CL("TOTAL BONUS X MULT"), CL(FormatScore(TotalBonus * BonusMultiplier(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, True, "vo_heynicebonus"
        Else
            DMD CL("TOTAL BONUS X MULT"), CL(FormatScore(TotalBonus * BonusMultiplier(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, True, "vo_notbad"
        End If
        AddScore2 TotalBonus * BonusMultiplier(CurrentPlayer)

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 9000, "EndOfBall2 '"
    Else 'if tilted then only add a short delay and move to the 2nd part of the end of the ball
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the CurrentPlayer)
'
Sub EndOfBall2()
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If ExtraBallsAwards(CurrentPlayer) > 0 Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' if no more EB's then turn off any Extra Ball light if there was any
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            LightShootAgain.State = 0
            
        End If

        ' You may wish to do a bit of a song AND dance at this point
		pupevent 816
		pupevent 839
        DMD CL("EXTRA BALL"), CL("SHOOT AGAIN"), "", eNone, eBlink, eNone, 1500, True, "vo_replay"

        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            ' debug.print "No More Balls, High Score Entry"
            ' Submit the CurrentPlayers score to the High Score system
            CheckHighScore()
        ' you may wish to play some music at this point
			
        Else

            ' not the last ball (for that player)
            ' if multiple players are playing then move onto the next one
            EndOfBallComplete()
        End If
    End If
End Sub

' This function is called when the end of bonus display
' (or high score entry finished) AND it either end the game or
' move onto the next player (or the next ball of the same player)
'
Sub EndOfBallComplete()
    Dim NextPlayer
'	pupevent 800
'	pupevent 832
    'debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1) Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame) Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        EndOfGame()

    ' you may wish to put a Game Over message on the desktop/backglass

    Else
        ' set the next player
        CurrentPlayer = NextPlayer

        ' make sure the correct display is up to date
        DMDScoreNow

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame > 1 Then
            Select Case CurrentPlayer
                Case 1:DMD "", CL("PLAYER 1"), "", eNone, eNone, eNone, 1000, True, "vo_player1"
                Case 2:DMD "", CL("PLAYER 2"), "", eNone, eNone, eNone, 1000, True, "vo_player2"
                Case 3:DMD "", CL("PLAYER 3"), "", eNone, eNone, eNone, 1000, True, "vo_player3"
                Case 4:DMD "", CL("PLAYER 4"), "", eNone, eNone, eNone, 1000, True, "vo_player4"
            End Select
        Else
            DMD "", CL("PLAYER 1"), "", eNone, eNone, eNone, 1000, True, "vo_youareup"
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    'debug.print "End Of Game"
    bGameInPLay = False

	StopScorbit

    ' just ended your game then play the end of game tune
    PlaySound "BBmu_death"
	
    vpmtimer.AddTimer 2500, "PlayEndQuote '"
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball

    ' set any lights for the attract mode
    GiOff
	pupevent 867 'game over topper	
	pupevent 804 'game over video
	
    vpmtimer.AddTimer 10000, "StartAttractMode '" 'allows end of game pup video to play, then restarts Attractmode
' you may wish to light any Game Over Light you may have
	
	pupevent 799 'restart attract mode video
End Sub

'this calculates the ball number in play
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Sub Drain_Hit()
    ' Destroy the ball
    Drain.DestroyBall
    If bGameInPLay = False Then Exit Sub 'don't do anything, just delete the ball
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' pretend to knock the ball into the ball storage mech
    PlaySoundAt "fx_drain", Drain
    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True) AND(Tilted = False) Then

        ' is the ball saver active,
        If(bBallSaverActive = True) Then

            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            ' we kick the ball with the autoplunger
            bAutoPlunger = True
            ' you may wish to put something on a display or play a sound at this point
			
            ' stop the ballsaver timer during the launch ball saver time, but not during multiballs
            If NOT bMultiBallMode Then
				Debug "Ball SAVE"
				pupevent 831
				pupevent 803
                DMD "_", CL("BALL SAVED"), "_", eNone, eBlinkfast, eNone, 2500, True, "vo_giveballback"
				'pupevent 803
                BallSaverTimerExpired_Timer
            End If
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and
                    changesong
                    ' turn off any multiball specific lights
                    ChangeGi white
                    ChangeGIIntensity 1
                    'stop any multiball modes of this game
                    StopMBmodes
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                ' End Mode and timers
                StopSong Song
                ChangeGi white
                ChangeGIIntensity 1
				pupevent 800
				pupevent 832
                ' Show the end of ball animation
                ' and continue with the end of ball
                ' DMD something?
                StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, if there is no animation then move to the end of ball
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

'****************
'DIM PARA ESCRITURA PC
'****************

Dim FlashershPC
Dim FadeParaPCOff

'**********************
' TIMER PARA ESCRITURA PC
'**********************

Sub TimerParaPc_Timer
	'pupevent 869  'blank video
	pupevent 821  ' video wake up neo
	pupevent 837  ' topper wake up neo
Select Case FlashershPC
    Case 0:  whiteRabbit001.visible = 0 
    Case 1:  FlasherWakeUpNeo.visible = 1
    Case 2:  TheMatrix.visible = 1
    Case 3:  HasYou.visible = 1
    Case 4:  FollowThe.visible = 1
    Case 5:  WhiteRabbit.visible = 1 
    Case 6:  whiteRabbit001.visible = 0
    Case 7,8 
    Case 9:  FadeParaPCOff = 0: FadeParaPC.Enabled = 1
	Case Else:  TimerParaPc.Enabled = 0
End Select
FlashershPC = FlashershPC + 1
End Sub

Sub FadeParaPC_Timer
Select Case FadeParaPCOff
    Case 0:  FlasherWakeUpNeo.visible = 0
    Case 1:  TheMatrix.visible = 0
    Case 2:  HasYou.visible = 0
    Case 3:  FollowThe.visible = 0
    Case 4:  WhiteRabbit.visible = 0
    Case 5:  whiteRabbit001.visible = 1
    case 6:  whiteRabbit001.visible = 0
    Case 7:  whiteRabbit001.visible = 1
    Case 8:  whiteRabbit001.visible = 0
    Case 9:  whiteRabbit001.visible = 1
    Case 10: whiteRabbit001.visible = 0
    Case 11: whiteRabbit001.visible = 1
    Case 12: whiteRabbit001.visible = 0
    Case 13:  FadeParaPC.Enabled = 0: DarkNight.Visible = 0: CreateNewBall '0n firstball only
    pupevent 871 'set Bbackground scrolling matrix 
End Select
FadeParaPCOff = FadeParaPCOff + 1
End Sub


Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySoundAt "fx_sensor", swPlungerRest
    DOF 208, DOFOn
    bBallInPlungerLane = True

	if bOnTheFirstBallScorbit And ScorbitActive = 1 And (Scorbit.bNeedsPairing) = false then ScorbitClaimQR(True)
    ' turn on Launch light is there is one
    'LaunchLight.State = 2
    ' be sure to update the Scoreboard after the animations, if any
    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        vpmtimer.addtimer 1500, "PlungerIM.AutoFire:DOF 120, DOFPulse:DOF 124, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), swPlungerRest:bAutoPlunger = False '"
    End If
    'Start the skillshot lights & variables if any
    If bSkillShotReady Then
        PlaySong "BBmu_wait"
        UpdateSkillshot()
        ' show the message to shoot the ball in case the player has fallen sleep
        swPlungerRest.TimerEnabled = 1
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    lighteffect 6
    bBallInPlungerLane = False
    DOF 208, DOFOff
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ChangeSong
        ResetSkillShotTimer.Enabled = 1
		ScorbitClaimQR(False)
		hideScorbit 'backup call to make sure all scorbit QR codes are gone
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
	bOnTheFirstBallScorbit = False
' turn off LaunchLight
' LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds
Sub swPlungerRest_Timer
    IF bOnTheFirstBall Then
        Select Case RndNbr(5)
            Case 1:DMD CL("NEO"), CL("WELCOME BACK"), "_", eNone, eNone, eNone, 2000, True, "vo_backforsomemoretorture"
            Case 2:DMD CL("KEEP"), CL("FIGHTING"), "_", eNone, eNone, eNone, 2000, True, "vo_Iknewyoullbeback"
            Case 3:DMD CL("DISCONNCT"), CL("YOU"), "_", eNone, eNone, eNone, 2000, True, "vo_areyouplayingthisgame"
            Case 4:DMD CL("YOU HAVE"), CL("WORK TO DO"), "_", eNone, eNone, eNone, 2000, True, "vo_shoothereandhere"
            Case 5:DMD CL("ARE YOU READY"), CL("FOR IT"), "_", eNone, eNone, eNone, 2000, True, "vo_welcomeback"
        End Select
    Else
        Select Case RndNbr(4)
            Case 1:DMD CL("FOLLOW"), CL("THE WHITE RABBIT"), "_", eNone, eNone, eNone, 2000, True, "vo_timetowakeup"
            Case 2:DMD CL("ENTER THE MATRIX"), CL(" NOW"), "_", eNone, eNone, eNone, 2000, True, "vo_whatareyouwaitingfor"
            Case 3:DMD CL("COME ON "), CL("MISTER ANDERSON"), "_", eNone, eNone, eNone, 2000, True, "vo_heypulltheplunger"
            Case 4:DMD CL("WAKE UP"), CL("NEO"), "_", eNone, eNone, eNone, 2000, True, "vo_areyouplayingthisgame"
        End Select
    End If
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    BallSaverTimerExpired.Interval = 2000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 2000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
  
   
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    BallSaverSpeedUpTimer.Enabled = False 'ensure this timer is also stopped
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
    
    ' if the table uses the same lights for the extra ball or replay then turn them on if needed
    If ExtraBallsAwards(CurrentPlayer) > 0 Then
        LightShootAgain.State = 1
      
    End If
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
    
  
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board

Sub AddScore(points) 'normal score routine
    If Tilted Then Exit Sub
    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
' you may wish to check to see if the player has gotten a replay
End Sub

Sub AddScore2(points) 'used in jackpots, skillshots, combos, and bonus as it does not use the PlayfieldMultiplier
    If Tilted Then Exit Sub
    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points
End Sub

' Add bonus to the bonuspoints AND update the score board

Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If Tilted Then Exit Sub
    ' add the bonus to the current players bonus variable
    BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points)
    ' Jackpots only generally increment in multiball mode AND not tilted
    ' but this doesn't have to be the case
    If Tilted Then Exit Sub

    ' If(bMultiBallMode = True) Then
    Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
    DMD "_", CL("INCREASED JACKPOT"), "_", eNone, eNone, eNone, 1000, True, ""
' you may wish to limit the jackpot to a upper limit, ie..
'	If (Jackpot >= 6000000) Then
'		Jackpot = 6000000
' 	End if
'End if
End Sub

Sub AddSuperJackpot(points) 'not used in this table
    If Tilted Then Exit Sub
End Sub

Sub AddBonusMultiplier(n)
    Dim NewBonusLevel
    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <= MaxBonusMultiplier) then
        ' then add and set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier(NewBonusLevel)
        DMD "_", CL("BONUS X " &NewBonusLevel), "_", eNone, eBlink, eNone, 2000, True, ""
    Else
        AddScore2 500000
        DMD "_", CL("500000"), "_", eNone, eNone, eNone, 1000, True, ""
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    UpdateBonusXLights(Level)
End Sub

Sub UpdateBonusXLights(Level) '4 lights in this table, from 2x to 5x
    ' Update the lights
    Select Case Level
        Case 1:li021.State = 0:li022.State = 0:li023.State = 0:li024.State = 0
        Case 2:li021.State = 1:li022.State = 0:li023.State = 0:li024.State = 0
        Case 3:li021.State = 1:li022.State = 1:li023.State = 0:li024.State = 0
        Case 4:li021.State = 1:li022.State = 1:li023.State = 1:li024.State = 0
        Case 5:li021.State = 1:li022.State = 1:li023.State = 1:li024.State = 1
    End Select
End Sub

Sub AddPlayfieldMultiplier(n)
    Dim snd
    Dim NewPFLevel
    ' if not at the maximum level x
    if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
        SetPlayfieldMultiplier(NewPFLevel)
        PlayThunder
        DMD "_", CL("PLAYFIELD X " &NewPFLevel), "_", eNone, eBlink, eNone, 2000, True, snd
        LightEffect 4
        ' Play a voice sound
        Select Case NewPFLevel
            Case 2:PlaySound "vo_2xplayfield"
            Case 3:PlaySound "vo_3xplayfield"
            Case 4:PlaySound "vo_4xplayfield"
            Case 5:PlaySound "vo_5xplayfield"
        End Select
    Else 'if the max is already lit
        AddScore2 500000
        DMD "_", CL("500000"), "_", eNone, eNone, eNone, 2000, True, ""
    End if
    ' restart the PlayfieldMultiplier timer to reduce the multiplier
    PFXTimer.Enabled = 0
    PFXTimer.Enabled = 1
End Sub

Sub PFXTimer_Timer
    DecreasePlayfieldMultiplier
End Sub

Sub DecreasePlayfieldMultiplier 'reduces by 1 the playfield multiplier
    Dim NewPFLevel
    ' if not at 1 already
    if(PlayfieldMultiplier(CurrentPlayer) > 1) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) - 1
        SetPlayfieldMultiplier(NewPFLevel)
    Else
        PFXTimer.Enabled = 0
    End if
End Sub

' Set the Playfield Multiplier to the specified level AND set any lights accordingly

Sub SetPlayfieldMultiplier(Level)
    ' Set the multiplier to the specified level
    PlayfieldMultiplier(CurrentPlayer) = Level
    UpdatePFXLights(Level)
End Sub

Sub UpdatePFXLights(Level) '4 lights in this table, from 2x to 5x
    ' Update the playfield multiplier lights
    Select Case Level
        Case 1:li025.State = 0:li026.State = 0:li027.State = 0:li027.State = 0
        Case 2:li025.State = 1:li026.State = 0:li027.State = 0:li027.State = 0
        Case 3:li025.State = 0:li026.State = 1:li027.State = 0:li027.State = 0
        Case 4:li025.State = 0:li026.State = 0:li027.State = 1:li027.State = 0
        Case 5:li025.State = 0:li026.State = 0:li027.State = 0:li027.State = 1
    End Select
' perhaps show also the multiplier in the DMD?
End Sub

Sub AwardExtraBall()
    '   If NOT bExtraBallWonThisBall Then 'in this table you can win several extra balls
	pupevent 816
    DMD "_", CL("EXTRA BALL WON"), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    DOF 124, DOFPulse
    PLaySound "vo_extraball"
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    'bExtraBallWonThisBall = True
    LightShootAgain.State = 1  'light the shoot again lamp
  
    GiEffect 2
    LightEffect 2
'    END If
End Sub

Sub AwardSpecial()
    DMD "_", CL("EXTRA GAME WON"), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    DOF 124, DOFPulse
    Credits = Credits + 1
    If bFreePlay = False Then DOF 125, DOFOn
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardJackpot() 'only used for the final mode
    DMD CL("JACKPOT"), CL(FormatScore(Jackpot(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 1500, True, "vo_Jackpot" &RndNbr(6)
    DOF 126, DOFPulse
    AddScore2 Jackpot(CurrentPlayer)
    Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + 100000
    LightEffect 2
    GiEffect 2
    FlashEffect 2
End Sub

Sub AwardTargetJackpot() 'award a target jackpot after hitting all targets
    DMD CL("TARGET JACKPOT"), CL(FormatScore(TargetJackpot(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 1500, True, "vo_Jackpot" &RndNbr(6)
    DOF 126, DOFPulse
    AddScore2 TargetJackpot(CurrentPlayer)
    TargetJackpot(CurrentPlayer) = TargetJackpot(CurrentPlayer) + 150000
    li069.State = 0
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardSuperJackpot() 'not used in this table as there are several superjackpots but I keep it as a reference
    DMD CL("SUPER JACKPOT"), CL(FormatScore(SuperJackpot(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_superjackpot"
    DOF 126, DOFPulse
    AddScore2 SuperJackpot(CurrentPlayer)
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardWeaponsSuperJackpot()
    DMD CL("DEJA VU JACKPOT"), CL(FormatScore(WeaponSJValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_superjackpot"
    DOF 126, DOFPulse
    AddScore2 WeaponSJValue(CurrentPlayer)
    WeaponSJValue(CurrentPlayer) = WeaponSJValue(CurrentPlayer) + ((Score(CurrentPlayer) * 0.2) \ 10) * 10 'increase the weapons score with 20%
    aWeaponSJactive = False
    li060.State = 0
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
	pupevent 865
	pupevent 866
    DMD CL("SKILLSHOT"), CL(FormatScore(SkillshotValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 1000, False, "vo_greatshot"
    DMD CL("SKILLSHOT"), CL(FormatScore(SkillshotValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 1000, True, "vo_greatshot"
    DOF 127, DOFPulse
    Addscore2 SkillShotValue(CurrentPlayer)
    ' increment the skillshot value with 100.000
    SkillShotValue(CurrentPlayer) = SkillShotValue(CurrentPlayer) + 100000
    'do some light show
    GiEffect 2
    LightEffect 2
End Sub

Sub AwardSuperSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
	pupevent 865
	pupevent 866 
    DMD CL("SUPER SKILLSHOT"), CL(FormatScore(SuperSkillshotValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 1000, False, "vo_greatshot"
    DMD CL("SUPER SKILLSHOT"), CL(FormatScore(SuperSkillshotValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 1000, True, "vo_excellentshot"
    DOF 127, DOFPulse
    Addscore2 SuperSkillshotValue(CurrentPlayer)
    ' increment the superskillshot value with 1.000.000
    SuperSkillshotValue(CurrentPlayer) = SuperSkillshotValue(CurrentPlayer) + 1000000
    'do some light show
    GiEffect 2
    LightEffect 2
End Sub

Sub aSkillshotTargets_Hit(idx) 'stop the skillshot if any other target is hit
    If bSkillshotReady then ResetSkillShotTimer_Timer
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(cGameName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 100000 End If
    x = LoadValue(cGameName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(cGameName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 100000 End If
    x = LoadValue(cGameName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(cGameName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 100000 End If
    x = LoadValue(cGameName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(cGameName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 100000 End If
    x = LoadValue(cGameName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(cGameName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff:End If
    x = LoadValue(cGameName, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
End Sub

Sub Savehs
    SaveValue cGameName, "HighScore1", HighScore(0)
    SaveValue cGameName, "HighScore1Name", HighScoreName(0)
    SaveValue cGameName, "HighScore2", HighScore(1)
    SaveValue cGameName, "HighScore2Name", HighScoreName(1)
    SaveValue cGameName, "HighScore3", HighScore(2)
    SaveValue cGameName, "HighScore3Name", HighScoreName(2)
    SaveValue cGameName, "HighScore4", HighScore(3)
    SaveValue cGameName, "HighScore4Name", HighScoreName(3)
    SaveValue cGameName, "Credits", Credits
    SaveValue cGameName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

Sub Reseths
    HighScoreName(0) = "AAA"
    HighScoreName(1) = "BBB"
    HighScoreName(2) = "CCC"
    HighScoreName(3) = "DDD"
    HighScore(0) = 1500000
    HighScore(1) = 1400000
    HighScore(2) = 1300000
    HighScore(3) = 1200000
    Savehs
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(3)
Dim hsCurrentDigit
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash

Sub CheckHighscore()
    Dim tmp
    tmp = Score(CurrentPlayer)

    If tmp > HighScore(0) Then 'add 1 credit for beating the highscore
        Credits = Credits + 1
        DOF 125, DOFOn
    End If

    If tmp > HighScore(3) Then
        PlaySound SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
        DOF 121, DOFPulse
        HighScore(3) = tmp
        PlayHighScoreQuote
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
        PlayNotGoodScore
    End If
End Sub

Sub HighScoreEntryInit()
	pupevent 868
	pupevent 870
    hsbModeActive = True
    PlaySound "vo_enterinitials"
    hsLetterFlash = 0

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' < is back arrow
    hsCurrentLetter = 1
    DMDFlush()
    HighScoreDisplayNameNow()

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter > len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit > 0) then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayNameNow()
        end if
    end if
End Sub

Sub HighScoreDisplayNameNow()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreDisplayName()
    Dim i
    Dim TempTopStr
    Dim TempBotStr

    TempTopStr = "YOUR NAME:"
    dLine(0) = ExpandLine(TempTopStr)
    DMDUpdate 0

    TempBotStr = "    > "
    if(hsCurrentDigit > 0) then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit > 1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit > 2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3) then
        if(hsLetterFlash <> 0) then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit < 1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit < 2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & " <    "
    dLine(1) = ExpandLine(TempBotStr)
    DMDUpdate 1
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2) then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False

    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ") then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    EndOfBallComplete()
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 3
        For j = 0 to 2
            If HighScore(j) < HighScore(j + 1) Then
                tmp = HighScore(j + 1)
                tmp2 = HighScoreName(j + 1)
                HighScore(j + 1) = HighScore(j)
                HighScoreName(j + 1) = HighScoreName(j)
                HighScore(j) = tmp
                HighScoreName(j) = tmp2
            End If
        Next
    Next
End Sub

'*********
'   LUT
'*********

Dim bLutActive, LUTImage
Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1) MOD 10:UpdateLUT:SaveLUT:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
    End Select
End Sub

' *************************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' only 5 effects: none, scroll left, scroll right, blink and blinkfast
' 3 Lines, treats all 3 lines as text.
' 1st and 2nd lines are 20 characters long
' 3rd line is just 1 character
' Example format:
' DMD "text1","text2","backpicture", eNone, eNone, eNone, 250, True, "sound"
' Short names:
' dq = display queue
' de = display effect
' *************************************************************************

Const eNone = 0        ' Instantly displayed
Const eScrollLeft = 1  ' scroll on from the right
Const eScrollRight = 2 ' scroll on from the left
Const eBlink = 3       ' Blink (blinks for 'TimeOn')
Const eBlinkFast = 4   ' Blink (blinks for 'TimeOn') at user specified intervals (fast speed)

Const dqSize = 64

Dim dqHead
Dim dqTail
Dim deSpeed
Dim deBlinkSlowRate
Dim deBlinkFastRate

Dim dLine(2)
Dim deCount(2)
Dim deCountEnd(2)
Dim deBlinkCycle(2)

Dim dqText(2, 64)
Dim dqEffect(2, 64)
Dim dqTimeOn(64)
Dim dqbFlush(64)
Dim dqSound(64)

Dim FlexDMD
Dim DMDScene

Sub DMD_Init() 'default/startup values
    If UseFlexDMD Then
        Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
        If Not FlexDMD is Nothing Then
            If FlexDMDHighQuality Then
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 256
                FlexDMD.Height = 64
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + i * 12, 6, 12, 22
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + (i - 20) * 12, 34, 12, 22
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            Else
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 128
                FlexDMD.Height = 32
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + i * 6, 3, 6, 11
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + (i - 20) * 6, 17, 6, 11
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            End If
        End If
    End If

    Dim i, j
    DMDFlush()
    deSpeed = 20
    deBlinkSlowRate = 10
    deBlinkFastRate = 5
    For i = 0 to 2
        dLine(i) = Space(20)
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
        dqTimeOn(i) = 0
        dqbFlush(i) = True
        dqSound(i) = ""
    Next
    dLine(2) = " "
    For i = 0 to 2
        For j = 0 to 64
            dqText(i, j) = ""
            dqEffect(i, j) = eNone
        Next
    Next
    DMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDFlush()
    Dim i
    DMDTimer.Enabled = False
    DMDEffectTimer.Enabled = False
    dqHead = 0
    dqTail = 0
    For i = 0 to 2
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
    Next
End Sub

Sub DMDScore()
    Dim tmp, tmp1, tmp1a, tmp1b, tmp2
    if(dqHead = dqTail) Then
        ' default when no modes are active
        tmp = RL(FormatScore(Score(Currentplayer) ) )
        tmp1 = FL("PLAYER " &CurrentPlayer, "BALL " & Balls)
        'back image
        If bJasonMBStarted Then
            tmp2 = "d_jason"
        ElseIf bFreddyMBStarted Then
            tmp2 = "d_freddy"
        ElseIf bMichaelMBStarted Then
            tmp2 = "d_michael"
        Else
            tmp2 = "d_border"
        End If
        'info on the second line
        Select Case Mode(CurrentPlayer, 0)
            Case 0: 'no Mode active
                If bTommyStarted Then
                    tmp1 = "SHOOT THE RIGHT RAMP"
                ElseIf bPoliceStarted Then
                    tmp1 = CL("HIT SENTINEL TARGET")
                End If
            Case 1: 'spinners
                If Not ReadyToKill Then
                    tmp1 = FL("SPINNERS LEFT", SpinNeeded-SpinCount)
                Else
                    tmp1 = CL("SHOOT THE SCOOP")
                End If
            Case 2:
                If Not ReadyToKill Then
                    tmp1 = FL("HITS LEFT", 4-TargetModeHits)
                Else
                    tmp1 = CL("SHOOT THE MAGNET")
                End If
            Case 3:
                If Not ReadyToKill Then
                    tmp1 = FL("HITS LEFT", 5-TargetModeHits)
                Else
                    tmp1 = CL("SHOOT THE SCOOP")
                End If
            Case 4:tmp1 = FL("HITS LEFT", 5-TargetModeHits)
            Case 5:
                If Not ReadyToKill Then
                    tmp1 = FL("HITS LEFT", 4-TargetModeHits)
                Else
                    tmp1 = CL("SHOOT THE TANK")
                End If
            Case 6:tmp1 = FL("HITS LEFT", 5-TargetModeHits)
            Case 7:tmp1 = FL("HITS LEFT", 4-TargetModeHits)
            Case 8:tmp1 = FL("HITS LEFT", 5-TargetModeHits)
            Case 9:tmp1 = FL("HITS LEFT", 5-TargetModeHits)
            Case 10:tmp1 = FL("HITS LEFT", 6-TargetModeHits)
            Case 11:tmp1 = FL("HITS LEFT", 6-TargetModeHits)
            Case 12:tmp1 = FL("SPINNERS LEFT", SpinNeeded-SpinCount)
            Case 13:tmp1 = FL("HITS LEFT", 6-TargetModeHits)
            Case 14:tmp1 = FL("HITS LEFT", 6-TargetModeHits)
            Case 15:tmp1 = CL("SHOOT JACKPOTS")
        End Select
    End If
    DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    if(dqTail < dqSize) Then
        if(Text0 = "_") Then
            dqEffect(0, dqTail) = eNone
            dqText(0, dqTail) = "_"
        Else
            dqEffect(0, dqTail) = Effect0
            dqText(0, dqTail) = ExpandLine(Text0)
        End If

        if(Text1 = "_") Then
            dqEffect(1, dqTail) = eNone
            dqText(1, dqTail) = "_"
        Else
            dqEffect(1, dqTail) = Effect1
            dqText(1, dqTail) = ExpandLine(Text1)
        End If

        if(Text2 = "_") Then
            dqEffect(2, dqTail) = eNone
            dqText(2, dqTail) = "_"
        Else
            dqEffect(2, dqTail) = Effect2
            dqText(2, dqTail) = Text2 'it is always 1 letter in this table
        End If

        dqTimeOn(dqTail) = TimeOn
        dqbFlush(dqTail) = bFlush
        dqSound(dqTail) = Sound
        dqTail = dqTail + 1
        if(dqTail = 1) Then
            DMDHead()
        End If
    End If
End Sub

Sub DMDHead()
    Dim i
    deCount(0) = 0
    deCount(1) = 0
    deCount(2) = 0
    DMDEffectTimer.Interval = deSpeed

    For i = 0 to 2
        Select Case dqEffect(i, dqHead)
            Case eNone:deCountEnd(i) = 1
            Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead) )
            Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead) )
            Case eBlink:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
            Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
        End Select
    Next
    if(dqSound(dqHead) <> "") Then
        PlaySound(dqSound(dqHead) )
    End If
    DMDEffectTimer.Enabled = True
End Sub

Sub DMDEffectTimer_Timer()
    DMDEffectTimer.Enabled = False
    DMDProcessEffectOn()
End Sub

Sub DMDTimer_Timer()
    Dim Head
    DMDTimer.Enabled = False
    Head = dqHead
    dqHead = dqHead + 1
    if(dqHead = dqTail) Then
        if(dqbFlush(Head) = True) Then
            DMDScoreNow()
        Else
            dqHead = 0
            DMDHead()
        End If
    Else
        DMDHead()
    End If
End Sub

Sub DMDProcessEffectOn()
    Dim i
    Dim BlinkEffect
    Dim Temp

    BlinkEffect = False

    For i = 0 to 2
        if(deCount(i) <> deCountEnd(i) ) Then
            deCount(i) = deCount(i) + 1

            select case(dqEffect(i, dqHead) )
                case eNone:
                    Temp = dqText(i, dqHead)
                case eScrollLeft:
                    Temp = Right(dLine(i), 19)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead), 21 - deCount(i), 1)
                    Temp = Temp & Left(dLine(i), 19)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkSlowRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(20)
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkFastRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(20)
                    End If
            End Select

            if(dqText(i, dqHead) <> "_") Then
                dLine(i) = Temp
                DMDUpdate i
            End If
        End If
    Next

    if(deCount(0) = deCountEnd(0) ) and(deCount(1) = deCountEnd(1) ) and(deCount(2) = deCountEnd(2) ) Then

        if(dqTimeOn(dqHead) = 0) Then
            DMDFlush()
        Else
            if(BlinkEffect = True) Then
                DMDTimer.Interval = 10
            Else
                DMDTimer.Interval = dqTimeOn(dqHead)
            End If

            DMDTimer.Enabled = True
        End If
    Else
        DMDEffectTimer.Enabled = True
    End If
End Sub

Sub Dmdoptionstimer_timer 'The DMDoption timer interval should be -1, so executes at the display frame rate
	Options_UpdateDMD

End Sub

Function ExpandLine(TempStr) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(20)
    Else
        if Len(TempStr) > Space(20) Then
            TempStr = Left(TempStr, Space(20) )
        Else
            if(Len(TempStr) < 20) Then
                TempStr = TempStr & Space(20 - Len(TempStr) )
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString

    NumString = CStr(abs(Num) )

    For i = Len(NumString) -3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1) ) then
            NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1) ) + 48) & right(NumString, Len(NumString) - i)
        end if
    Next
    FormatScore = NumString
End function

Function FL(NumString1, NumString2) 'Fill line
    Dim Temp, TempStr
    Temp = 20 - Len(NumString1) - Len(NumString2)
    TempStr = NumString1 & Space(Temp) & NumString2
    FL = TempStr
End Function

Function CL(NumString) 'center line
    Dim Temp, TempStr
    Temp = (20 - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(NumString) 'right line
    Dim Temp, TempStr
    Temp = 20 - Len(NumString)
    TempStr = Space(Temp) & NumString
    RL = TempStr
End Function

'**************
' Update DMD
'**************

Sub DMDUpdate(id)
    Dim digit, value
    If UseFlexDMD Then FlexDMD.LockRenderThread
    Select Case id
        Case 0 'top text line
            For digit = 0 to 19
                DMDDisplayChar mid(dLine(0), digit + 1, 1), digit
            Next
        Case 1 'bottom text line
            For digit = 20 to 39
                DMDDisplayChar mid(dLine(1), digit -19, 1), digit
            Next
        Case 2 ' back image - back animations
            If dLine(2) = "" OR dLine(2) = " " Then dLine(2) = "d_border"
            Digits(40).ImageA = dLine(2)
            If UseFlexDMD Then DMDScene.GetImage("Back").Bitmap = FlexDMD.NewImage("", "VPX." & dLine(2) & "&dmd=2").Bitmap
    End Select
    If UseFlexDMD Then FlexDMD.UnlockRenderThread
End Sub

Sub DMDDisplayChar(achar, adigit)
    If achar = "" Then achar = " "
    achar = ASC(achar)
    Digits(adigit).ImageA = Chars(achar)
    If UseFlexDMD Then DMDScene.GetImage("Dig" & adigit).Bitmap = FlexDMD.NewImage("", "VPX." & Chars(achar) & "&dmd=2&add").Bitmap
End Sub

'****************************
' JP's new DMD using flashers
'****************************

Dim Digits, Chars(255), Images(255)

DMDInit

Sub DMDInit
    Dim i
    Digits = Array(digit001, digit002, digit003, digit004, digit005, digit006, digit007, digit008, digit009, digit010, _
        digit011, digit012, digit013, digit014, digit015, digit016, digit017, digit018, digit019, digit020,            _
        digit021, digit022, digit023, digit024, digit025, digit026, digit027, digit028, digit029, digit030,            _
        digit031, digit032, digit033, digit034, digit035, digit036, digit037, digit038, digit039, digit040,            _
        digit041)
    For i = 0 to 255:Chars(i) = "d_empty":Next

    Chars(32) = "d_empty"
    Chars(33) = ""       '!
    Chars(34) = ""       '"
    Chars(35) = ""       '#
    Chars(36) = ""       '$
    Chars(37) = ""       '%
    Chars(38) = ""       '&
    Chars(39) = ""       ''
    Chars(40) = ""       '(
    Chars(41) = ""       ')
    Chars(42) = ""       '*
    Chars(43) = ""       '+
    Chars(44) = ""       '
    Chars(45) = ""       '-
    Chars(46) = "d_dot"  '.
    Chars(47) = ""       '/
    Chars(48) = "d_0"    '0
    Chars(49) = "d_1"    '1
    Chars(50) = "d_2"    '2
    Chars(51) = "d_3"    '3
    Chars(52) = "d_4"    '4
    Chars(53) = "d_5"    '5
    Chars(54) = "d_6"    '6
    Chars(55) = "d_7"    '7
    Chars(56) = "d_8"    '8
    Chars(57) = "d_9"    '9
    Chars(60) = "d_less" '<
    Chars(61) = ""       '=
    Chars(62) = "d_more" '>
    Chars(64) = ""       '@
    Chars(65) = "d_a"    'A
    Chars(66) = "d_b"    'B
    Chars(67) = "d_c"    'C
    Chars(68) = "d_d"    'D
    Chars(69) = "d_e"    'E
    Chars(70) = "d_f"    'F
    Chars(71) = "d_g"    'G
    Chars(72) = "d_h"    'H
    Chars(73) = "d_i"    'I
    Chars(74) = "d_j"    'J
    Chars(75) = "d_k"    'K
    Chars(76) = "d_l"    'L
    Chars(77) = "d_m"    'M
    Chars(78) = "d_n"    'N
    Chars(79) = "d_o"    'O
    Chars(80) = "d_p"    'P
    Chars(81) = "d_q"    'Q
    Chars(82) = "d_r"    'R
    Chars(83) = "d_s"    'S
    Chars(84) = "d_t"    'T
    Chars(85) = "d_u"    'U
    Chars(86) = "d_v"    'V
    Chars(87) = "d_w"    'W
    Chars(88) = "d_x"    'X
    Chars(89) = "d_y"    'Y
    Chars(90) = "d_z"    'Z
    Chars(94) = "d_up"   '^
    '    Chars(95) = '_
    Chars(96) = "d_0a"  '0.
    Chars(97) = "d_1a"  '1. 'a
    Chars(98) = "d_2a"  '2. 'b
    Chars(99) = "d_3a"  '3. 'c
    Chars(100) = "d_4a" '4. 'd
    Chars(101) = "d_5a" '5. 'e
    Chars(102) = "d_6a" '6. 'f
    Chars(103) = "d_7a" '7. 'g
    Chars(104) = "d_8a" '8. 'h
    Chars(105) = "d_9a" '9. 'i
    Chars(106) = ""     'j
    Chars(107) = ""     'k
    Chars(108) = ""     'l
    Chars(109) = ""     'm
    Chars(110) = ""     'n
    Chars(111) = ""     'o
    Chars(112) = ""     'p
    Chars(113) = ""     'q
    Chars(114) = ""     'r
    Chars(115) = ""     's
    Chars(116) = ""     't
    Chars(117) = ""     'u
    Chars(118) = ""     'v
    Chars(119) = ""     'w
    Chars(120) = ""     'x
    Chars(121) = ""     'y
    Chars(122) = ""     'z
    Chars(123) = ""     '{
    Chars(124) = ""     '|
    Chars(125) = ""     '}
    Chars(126) = ""     '~
End Sub

'********************
' Real Time updates
'********************
'used for all the real time updates

Sub Realtime_Timer
    RollingUpdate
    
    LeftFlipperTop002NEW.RotZ = LeftFlipper.CurrentAngle 
    LeftFlipperTop001.RotZ = LeftFlipper001.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
' add any other real time update subs, like gates or diverters, flippers
End Sub

'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version

    If TypeName(MyLight) = "Light" Then

        If FinalState = 2 Then
            FinalState = MyLight.State 'Keep the current light state
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

        Dim steps

        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
        If FinalState = 2 Then                      'Keep the current flasher state
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

        ' Start blink timer and create timer subroutine
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' 11 colors: red, orange, amber, yellow...
'******************************************

'colors
Const red = 5
Const orange = 4
Const amber = 6
Const yellow = 3
Const darkgreen = 7
Const green = 2
Const blue = 1
Const darkblue = 8
Const purple = 9
Const white = 11
Const teal = 10

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Select Case col
        Case red
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case orange
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case amber
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 8, 0)
            n.colorfull = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 16, 0)
            n.colorfull = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(64, 0, 96)
            n.colorfull = RGB(128, 0, 192)
        Case white
            n.color = RGB(193, 91, 0)
            n.colorfull = RGB(255, 197, 143)
        Case teal
            n.color = RGB(1, 64, 62)
            n.colorfull = RGB(2, 128, 126)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub SetFlashColor(n, col, stat) 'stat 0 = off, 1 = on, -1= no change - no blink for the flashers, use FlashForMs
    Select Case col
        Case red
            n.color = RGB(255, 0, 0)
        Case orange
            n.color = RGB(255, 64, 0)
        Case amber
            n.color = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 192)
        Case white
            n.color = RGB(255, 197, 143)
        Case teal
            n.color = RGB(2, 128, 126)
    End Select
    If stat <> -1 Then
        n.Visible = stat
    End If
End Sub

'*************************
' Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

Sub StartRainbow(n) 'n is a collection
    set RainbowLights = n
    RGBStep = 0
    RGBFactor = 5
    rRed = 255
    rGreen = 0
    rBlue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow()
    RainbowTimer.Enabled = 0
    TurnOffArrows
End Sub

Sub TurnOffArrows() 'during Modes when changing modes
    For each x in aArrows
        SetLightColor x, white, 0
    Next
End Sub

Sub TurnOnArrows(incolor) 'blink during Modes
    For each x in aArrows
        SetLightColor x, incolor, 2
    Next
End Sub

Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
        Case 0 'Green
            rGreen = rGreen + RGBFactor
            If rGreen > 255 then
                rGreen = 255
                RGBStep = 1
            End If
        Case 1 'Red
            rRed = rRed - RGBFactor
            If rRed < 0 then
                rRed = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            rBlue = rBlue + RGBFactor
            If rBlue > 255 then
                rBlue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            rGreen = rGreen - RGBFactor
            If rGreen < 0 then
                rGreen = 0
                RGBStep = 4
            End If
        Case 4 'Red
            rRed = rRed + RGBFactor
            If rRed > 255 then
                rRed = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            rBlue = rBlue - RGBFactor
            If rBlue < 0 then
                rBlue = 0
                RGBStep = 0
            End If
    End Select
    For each obj in RainbowLights
        obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
    Next
End Sub

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    Dim ii
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1) Then
        DMD CL("LAST SCORE"), CL("PLAYER 1 " &FormatScore(Score(1) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(2) Then
        DMD CL("LAST SCORE"), CL("PLAYER 2 " &FormatScore(Score(2) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(3) Then
        DMD CL("LAST SCORE"), CL("PLAYER 3 " &FormatScore(Score(3) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(4) Then
        DMD CL("LAST SCORE"), CL("PLAYER 4 " &FormatScore(Score(4) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    DMD "", CL("GAME OVER"), "", eNone, eBlink, eNone, 2000, False, ""
    If bFreePlay Then
        DMD "", CL("FREE PLAY"), "", eNone, eBlink, eNone, 2000, False, ""
    Else
        If Credits > 0 Then
            DMD CL("CREDITS " & Credits), CL("PRESS START"), "", eNone, eBlink, eNone, 2000, False, ""
        Else
            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 2000, False, ""
        End If
    End If
    DMD "        VP COOKS", "          PRESENTS", "d_jppresents", eNone, eNone, eNone, 3000, False, ""
    DMD "", "", "d_title", eNone, eNone, eNone, 4000, False, ""
    DMD "", CL("ROM VERSION " &myversion), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("HIGHSCORES"), Space(20), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CL("HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CL("HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(20), Space(20), "", eScrollLeft, eScrollLeft, eNone, 500, False, ""
End Sub

Sub StartAttractMode
	'pupevent 799
    StartLightSeq
    StartRainbow aArrows
    DMDFlush
    ShowTableInfo
    'PlaySong "BBmu_game_over"
End Sub

Sub StopAttractMode
	'pupevent 797
	'pupevent 798
'	pupevent 831
    StopRainbow
    DMDScoreNow
    LightSeqAttract.StopPlay
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 10
    'LightSeqAttract.Play SeqAllOff
    LightSeqAttract.Play SeqDiagUpRightOn, 25, 2
    LightSeqAttract.Play SeqStripe1VertOn, 25
    LightSeqAttract.Play SeqClockRightOn, 180, 2
    LightSeqAttract.Play SeqFanLeftUpOn, 50, 2
    LightSeqAttract.Play SeqFanRightUpOn, 50, 2
    LightSeqAttract.Play SeqScrewRightOn, 50, 2

    LightSeqAttract.Play SeqDiagDownLeftOn, 25, 2
    LightSeqAttract.Play SeqStripe2VertOn, 25, 2
    LightSeqAttract.Play SeqFanLeftDownOn, 50, 2
    LightSeqAttract.Play SeqFanRightDownOn, 50, 2
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
End Sub

Sub LightSeqTopFlashers_PlayDone()
    FlashEffect 7
End Sub

'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' droptargets, animations, timers, etc
Sub VPObjects_Init
End Sub

' tables variables and Mode init
Dim bRotateLights
Dim Mode(4, 15) 'the first 4 is the current player, contains status of the Mode, 0 not started, 1 won, 2 started
Dim Weapons(4)  ' collected weapons
Dim LoopCount
Dim LoopHits(4)
Dim LoopValue(4)
Dim SlingCount 'used for the db2s animation
Dim ComboCount
Dim ComboHits(4)
Dim ComboValue(4)
Dim Mystery(4, 4) 'inlane lights for each player
Dim BumperHits(4)
Dim BumperNeededHits(4)
Dim TargetHits(4, 7) '6 targets + the bumper -the blue lights
Dim WeaponHits(4, 6) '6 lights, lanes and the magnet post
Dim aWeaponSJactive
Dim WeaponSJValue(4)
Dim bFlippersEnabled
Dim bTommyStarted
Dim TommyCount 'counts the seconds left
Dim TommyValue
Dim bPoliceStarted
Dim PoliceTargetHits
Dim PoliceCount 'counts the seconds, used for the police jackpot
Dim TeensKilled(4)
Dim TeensKilledValue(4)
Dim CounselorsKilled(4)
Dim CenterSpinnerHits(4)
Dim LeftSpinnerHits(4)
Dim RightSpinnerHits(4)
Dim TargetJackpot(4)
Dim bJasonMBStarted
Dim bFreddyMBStarted
Dim bMichaelMBStarted
Dim ArrowMultiPlier(8) 'used for the Jackpot multiplier and the color for the Jason MB arrow lights
Dim FreddySJValue
Dim MichaelSJValue
'variables used only in the modes
Dim NewMode
Dim ReadyToKill ' final shot in a mode
Dim SpinCount
Dim SpinNeeded
Dim TargetModeHits 'mode 2,3 hits
Dim EndModeCountdown
Dim BlueTargetsCount
Dim ArrowsCount

Sub Game_Init() 'called at the start of a new game
	Debug "Game INIT"
	Debug4 "Game INIT"
    Dim i, j
    bExtraBallWonThisBall = False
    TurnOffPlayfieldLights()
    FlashershPC = 0

    'Init Variables
    bRotateLights = True
    aWeaponSJactive = False
    bFlippersEnabled = True 'only disabled if the police or Tommy catches you
    bTommyStarted = False
    TommyCount = 1          'we set it to 1 because it also acts as a multiplier in the hurry up
    TommyValue = 500000
    bPoliceStarted = False
    PoliceTargetHits = 0
    bJasonMBStarted = False
    bFreddyMBStarted = False
    bMichaelMBStarted = False
    FreddySJValue = 1000000
    MichaelSJValue = 1000000
    NewMode = 0
    SpinCount = 0
    ReadyToKill = False
    SpinNeeded = 0
    EndModeCountdown = 0
    BlueTargetsCount = 0
    ArrowsCount = 0
    For i = 0 to 4
        SkillShotValue(i) = 500000
        SuperSkillShotValue(i) = 5000000
        LoopValue(i) = 500000
        ComboValue(i) = 500000
        BumperHits(i) = 0
        BumperNeededHits(i) = 10
        Weapons(i) = 0
        WeaponSJValue(i) = 3500000
        TeensKilled(i) = 0
        TeensKilledValue(i) = 250000
        CounselorsKilled(i) = 0
        LoopHits(i) = 0
        ComboHits(i) = 0
        CenterSpinnerHits(i) = 0
        LeftSpinnerHits(i) = 0
        RightSpinnerHits(i) = 0
        TargetJackpot(i) = 500000
        Jackpot(i) = 500000 'only used in the last mode
        BallsInLock(i) = 0
        ArrowMultiPlier(i) = 1
		bFirstBall(i) = True
    Next
    For i = 0 to 4
        For j = 0 to 4
            Mystery(i, j) = 0
        Next
    Next
    For i = 0 to 4
        For j = 0 to 15
            Mode(i, j) = 0
        Next
    Next
    For i = 0 to 4
        For j = 0 to 7
            TargetHits(i, j) = 1
        Next
    Next
    For i = 0 to 4
        For j = 0 to 6
            WeaponHits(i, j) = 1
        Next
    Next
    LoopCount = 0
    ComboCount = 0
    TeenTimer.Enabled = 1
End Sub

Sub InstantInfo
    Dim tmp
    DMD CL("INSTANT INFO"), "", "", eNone, eNone, eNone, 1000, False, ""
    Select Case NewMode
        Case 1 'A.J Mason = Super Spinners
            DMD CL("CURRENT MODE"), CL("WHITE RABBIT"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT THE SPINNERS"), CL("AND SCOOP TO FINISH"), "", eNone, eNone, eNone, 2000, False, ""
        Case 2 'Adam = 5 Targets at semi random
            DMD CL("CURRENT MODE"), CL("AGENT SMITH"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("HIT THE LIT TARGETS"), CL("AND MAGNET TO FINISH"), "", eNone, eNone, eNone, 2000, False, ""
        Case 3 'Brandon = 5 Flashing Shots 90 seconds to complete
            DMD CL("CURRENT MODE"), CL("ORACLE"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT THE LIGHTS"), CL("YOU HAVE 90 SECONDS"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT THE SCOOP"), CL("BEFORE TIME IS UP"), "", eNone, eNone, eNone, 2000, False, ""
        Case 4 'Chad = 5 Orbits       
			DMD CL("CURRENT MODE"), CL("SMITH CLONES"), "", eNone, eNone, eNone, 4000, False, ""    
            DMD CL("SHOOT 5"), CL("LIT ORBITS"), "", eNone, eNone, eNone, 2000, False, ""
        Case 5 'Deborah = Shoot 4 lights 60 seconds      
			DMD CL("CURRENT MODE"), CL("SERAPH"), "", eNone, eNone, eNone, 4000, False, ""   
            DMD CL("SHOOT 4 LIGHTS"), CL("YOU HAVE 60 SECONDS"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("AND HOLE TO FINISH"), CL("AND HOLE TO FINISH"), "", eNone, eNone, eNone, 2000, False, ""
        Case 6 'Eric = Shoot the ramps		
            DMD CL("CURRENT MODE"), CL("KEYMAKER"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("ERIC"), CL("SHOOT 5 RAMPS"), "", eNone, eNone, eNone, 2000, False, ""
        Case 7 'Jenny=  Target Frenzy		
            DMD CL("CURRENT MODE"), CL("MEROVINGIAN"), "", eNone, eNone, eNone, 4000, False, ""  
            DMD CL("SHOOT 4"), CL("LIT TARGETS"), "", eNone, eNone, eNone, 2000, False, ""
        Case 8 'Mitch = 5 Targets in rotation           
			DMD CL("CURRENT MODE"), CL("THE TWINS"), "", eNone, eNone, eNone, 4000, False, ""  
            DMD CL(""), CL("SHOOT 5 LIT TARGETS"), "", eNone, eNone, eNone, 2000, False, ""
        Case 9 'Fox = Magnet post
            DMD CL("CURRENT MODE"), CL("THE ARCHITECT"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT THE MAGNET"), CL("5 TIMES"), "", eNone, eNone, eNone, 2000, False, ""
        Case 10 'Victoria = Ramps and Orbits
            DMD CL("CURRENT MODE"), CL("TRAINMAN"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT 6 RAMPS"), CL("OR ORBITS"), "", eNone, eNone, eNone, 2000, False, ""
        Case 11 'Kenny = 5 Blue Targets at random
            DMD CL("CURRENT MODE"), CL("DEUS EX MACHINA"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT 6"), CL("BLUE TARGETS"), "", eNone, eNone, eNone, 2000, False, ""
        Case 12 'Sheldon = Super Spinners at random			
            DMD CL("CURRENT MODE"), CL("CYPHER"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT THE SPINNERS"), CL("AT RANDOM"), "", eNone, eNone, eNone, 2000, False, ""
        Case 13 'Tiffany = Follow the Lights
            DMD CL("CURRENT MODE"), CL("ZION SPEECH"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT 6 LIT"), CL("LIGHTS"), "", eNone, eNone, eNone, 2000, False, ""
        Case 14 'Vanessa = Follow the Lights random
            DMD CL("CURRENT MODE"), CL("PERSEPHONE"), "", eNone, eNone, eNone, 4000, False, ""
            DMD CL("SHOOT 6 LIT"), CL("LIGHTS"), "", eNone, eNone, eNone, 2000, False, ""
    End Select
    DMD CL("YOUR SCORE"), CL(FormatScore(Score(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("EXTRA BALLS"), CL(ExtraBallsAwards(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("PLAYFIELD MULTIPLIER"), CL("X " &PlayfieldMultiplier(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("BONUS MULTIPLIER"), CL("X " &BonusMultiplier(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("SKILLSHOT VALUE"), CL(FormatScore(SkillshotValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("SUPR SKILLSHOT VALUE"), CL(FormatScore(SuperSkillshotValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("RAMP COMBO VALUE"), CL(FormatScore(ComboValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("LOOP COMBO VALUE"), CL(FormatScore(LoopValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("DEJA VU JACKPOT"), CL(FormatScore(WeaponSJValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("TARGET JACKPOT"), CL(FormatScore(TargetJackpot(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    If Score(1) Then
        DMD CL("PLAYER 1 SCORE"), CL(FormatScore(Score(1) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
    If Score(2) Then
        DMD CL("PLAYER 2 SCORE"), CL(FormatScore(Score(2) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
    If Score(3) Then
        DMD CL("PLAYER 3 SCORE"), CL(FormatScore(Score(3) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
    If Score(4) Then
        DMD CL("PLAYER 4 SCORE"), CL(FormatScore(Score(4) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
End Sub

Sub StopMBmodes 'stop multiball modes after loosing the last multibal
    If bJasonMBStarted Then StopJasonMultiball
    If bFreddyMBStarted Then StopFreddyMultiball
    If bMichaelMBStarted Then StopMichaelMultiball
	pupevent 829
    If NewMode = 15 Then StopMode
End Sub

Sub StopEndOfBallMode()                         'this sub is called after the last ball in play is drained, reset skillshot, modes, timers
    If li048.State then SuperJackpotTimer_Timer 'to turn off the timer
    If bPoliceStarted Then StopPolice
    if bTommyStarted Then StopTommyJarvis
    StopMode
End Sub

Sub ResetNewBallVariables() 'reset variables and lights for a new ball or player
    'turn on or off the needed lights before a new ball is released
    TurnOffPlayfieldLights
    libumper.State = 0
    Flasher002.Visible = 0
    'set up the lights according to the player achievments
    BonusMultiplier(CurrentPlayer) = 1 'no need to update light as the 1x light do not exists
    UpdateTargetLights
    UpdateWeaponLights                 ' the W lights
    UpdateWeaponLights2                ' the collected weapons
    aWeaponSJactive = False
    UpdateLockLights                   ' turn on the lock lights for the current player
    UpdateModeLights                   ' show the killed counselors
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateTargetLights 'CurrentPlayer
    Li031.State = TargetHits(CurrentPlayer, 1)
    Li049.State = TargetHits(CurrentPlayer, 2)
    Li050.State = TargetHits(CurrentPlayer, 3)
    Li051.State = TargetHits(CurrentPlayer, 4)
    Li058.State = TargetHits(CurrentPlayer, 5)
    Li057.State = TargetHits(CurrentPlayer, 6)
    Li079.State = TargetHits(CurrentPlayer, 7)
End Sub

Sub TurnOffBlueTargets 'Turns off all blue targets at the start of a mode that uses the blue targets
    Li031.State = 0
    Li049.State = 0
    Li050.State = 0
    Li051.State = 0
    Li058.State = 0
    Li057.State = 0
    Li079.State = 0
End Sub

Sub ResetTargetLights 'CurrentPlayer
    Dim j
    For j = 0 to 7
        TargetHits(CurrentPlayer, j) = 1
    Next
    UpdateTargetLights
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    LightSeqSkillshot.Play SeqAllOff
    DMD CL("HIT LIT LIGHT"), CL("FOR SKILLSHOT"), "", eNone, eNone, eNone, 3000, True, ""
    li034.State = 2
   
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    bRotateLights = True
    LightSeqSkillshot.StopPlay
    Li034.State = 0
    
    DMDScoreNow
End Sub

' *********************************************************************
'                        Table Object Hit Events
'
' Any target hit Sub will follow this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/Mode this trigger is a member of
' - set the "LastSwitchHit" variable in case it is needed later
' *********************************************************************

'*********************************************************
' Slingshots has been hit
' In this table the slingshots change the outlanes lights

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFcontactors), Lemk
    DOF 105, DOFPulse
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 530
    ' check modes
    ' add some effect to the table?
    If B2sOn then
        SlingCount = 0
        SlingTimer.Enabled = 1
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub



Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_slingshot", 104, DOFPulse, DOFcontactors), Remk
    DOF 106, DOFPulse
    PlaySound "sfx_lasergunAPU" 
    FlashForms FlasherArmaDerecha,250, 15, 0
    FlashForms FlasherArmaIzquierda,300, 25, 0
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 530
    ' check modes
    ' add some effect to the table?
    If B2sOn then
        SlingCount = 0
        SlingTimer.Enabled = 1
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub SlingTimer_Timer
    Select case SlingCount
        Case 0, 2, 4, 6, 8:Controller.B2SSetData 10, 1
        Case 1, 3, 5, 7, 9:Controller.B2SSetData 10, 0
        Case 10:SlingTimer.Enabled = 0
    End Select
    SlingCount = SlingCount + 1
End Sub

'***********************
'        Bumper
'***********************
' Bumper Jackpot is scored when the bumper light is on
' the value is always 200.000 + 20% of the score

Sub Bumper1_Hit ' W6
    If Tilted Then Exit Sub
    Dim tmp
	pupevent 802
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), Bumper1
    If B2sOn then
        SlingCount = 0
        SlingTimer.Enabled = 1
    End If
    DOF 138, DOFPulse
    ' add some points
    If libumper.State Then 'the light is on so give the bumper Jackpot
        FlashForms libumper, 1500, 75, 0
        FlashForms Flasher002, 1500, 75, 0
        DOF 127, DOFPulse
        tmp = 100000 + INT(Score(CurrentPlayer) * 0.01) * 10 'the bumper jackpot is 100.000 + 10% of the score
        DMD CL("BUMPER JACKPOT"), CL(FormatScore(tmp) ), "", eNone, eNone, eNone, 1500, True, "vo_jackpot" &RndNbr(6)
        AddScore2 tmp
    Else 'score normal points
        AddScore 1000
    End If
    ' check for modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2, 7, 8, 11
				If li079.State Then
					TargetModeHits = TargetModeHits + 1
					li079.State = 0
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 7) = 0
				CheckTargets
		End Select
	End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper1"
    ' increase the bumper hit count and increase the bumper value after each 30 hits
    BumperHits(CurrentPlayer) = BumperHits(CurrentPlayer) + 1
    ' Check the bumper hits to lit the bumper to collect the bumper jackpot
    If BumperHits(CurrentPlayer) = BumperNeededHits(CurrentPlayer) Then
        libumper.State = 1
        Flasher002.Visible = 1
        BumperNeededHits(CurrentPlayer) = BumperNeededHits(CurrentPlayer) + 10 + RndNbr(10)
    End If
End Sub

'*********
' Lanes
'*********
' in and outlanes - mystery ?
Sub Trigger001_Hit
    PLaySoundAt "fx_sensor", Trigger001
    DOF 207, DOFPulse
    If Tilted Then Exit Sub
    Addscore 5000
    Mystery(CurrentPlayer, 1) = 1
    CheckMystery
End Sub

Sub Trigger002_Hit
    PLaySoundAt "fx_sensor", Trigger002
    DOF 207, DOFPulse
    If Tilted Then Exit Sub
    Addscore 1000
    Mystery(CurrentPlayer, 2) = 1
    CheckMystery
End Sub

Sub Trigger003_Hit
    PLaySoundAt "fx_sensor", Trigger003
    DOF 207, DOFPulse
    If Tilted Then Exit Sub
    Addscore 1000
    Mystery(CurrentPlayer, 3) = 1
    CheckMystery
End Sub

Sub Trigger004_Hit
    PLaySoundAt "fx_sensor", Trigger004
    DOF 207, DOFPulse
    If Tilted Then Exit Sub
    Addscore 5000
    Mystery(CurrentPlayer, 4) = 1
    CheckMystery
End Sub

Sub UpdateMysteryLights
    'update lane lights
    li017.State = Mystery(CurrentPlayer, 1)
    li018.State = Mystery(CurrentPlayer, 2)
    li019.State = Mystery(CurrentPlayer, 3)
    li020.State = Mystery(CurrentPlayer, 4)
    If Mystery(CurrentPlayer, 1) + Mystery(CurrentPlayer, 2) + Mystery(CurrentPlayer, 3) + Mystery(CurrentPlayer, 4) = 4 Then
        li078.State = 1
    End If
End Sub

Sub RotateLaneLights(n) 'n is the direction, 1 or 0, left or right. They are rotated by the flippers
    Dim tmp
    If bRotateLights Then
        If n = 1 Then
            tmp = Mystery(CurrentPlayer, 1)
            Mystery(CurrentPlayer, 1) = Mystery(CurrentPlayer, 2)
            Mystery(CurrentPlayer, 2) = Mystery(CurrentPlayer, 3)
            Mystery(CurrentPlayer, 3) = Mystery(CurrentPlayer, 4)
            Mystery(CurrentPlayer, 4) = tmp
        Else
            tmp = Mystery(CurrentPlayer, 4)
            Mystery(CurrentPlayer, 4) = Mystery(CurrentPlayer, 3)
            Mystery(CurrentPlayer, 3) = Mystery(CurrentPlayer, 2)
            Mystery(CurrentPlayer, 2) = Mystery(CurrentPlayer, 1)
            Mystery(CurrentPlayer, 1) = tmp
        End If
    End If
    UpdateMysteryLights
End Sub

'table lanes
Sub Trigger005_Hit
    PLaySoundAt "fx_sensor", Trigger005
    If Tilted Then Exit Sub
    Addscore 1000
    If bSkillShotReady Then li034.State = 0
    If bMichaelMBStarted AND li032.State Then 'award the michael super jackpot
        DOF 126, DOFPulse
        DMD CL("SUPER JACKPOT"), CL(FormatScore(MichaelSJValue) ), "_", eBlink, eNone, eNone, 1000, True, "vo_superjackpot"
        Addscore2 MichaelSJValue
        MichaelSJValue = 1000000
        li032.State = 0
        LightEffect 2
        GiEffect 2
    End If
End Sub

Sub Trigger006_Hit 'skillshot 1
    PLaySoundAt "fx_sensor", Trigger006
    If Tilted Then Exit Sub
    Addscore 1000
    If bSkillShotReady AND li034.State Then AwardSkillshot
End Sub

Sub Trigger008_Hit 'end top loop
    PLaySoundAt "fx_sensor", Trigger008
    DOF 206, DOFPulse
    If Tilted Then Exit Sub
    Addscore 5000
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger008"
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 4:TargetModeHits = TargetModeHits + 1:CheckWinMode
		End Select
	End If
End Sub

Sub Trigger009_Hit 'right loop
	pupevent 826
    PLaySoundAt "fx_sensor", Trigger009
    DOF 206, DOFPulse
    If Tilted Then Exit Sub
    PlayThunder
    Addscore 10000
    Flashforms f2A, 800, 50, 0
    Flashforms F2B, 800, 50, 0
    Flashforms Flasher004, 800, 50, 0
    Flashforms Flasher005, 800, 50, 0
    If LastSwitchHit = "Trigger008" Then
        AwardLoop
    Else
        li061.State = 2 'super loops light
        LoopCount = 1
    End If
    If F002.State Then TeenKilled:F002.State = 0
    ' Weapons Super Jackpot
    If aWeaponSJactive AND li060.State Then AwardWeaponsSuperJackpot
    'Jason multiball
    If bJasonMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(7) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        LightEffect 2
        GiEffect 2
        Addscore2 ArrowMultiPlier(7) * 1000000
        If ArrowMultiPlier(7) < 5 Then
            ArrowMultiPlier(7) = ArrowMultiPlier(7) + 1
            UpdateArrowLights
        End If
    End If
    'Freddy multiball
    If bFreddyMBStarted and li060.State Then
        DOF 126, DOFPulse
        DMD CL("SUPER JACKPOT"), CL(FormatScore(FreddySJValue) ), "_", eBlink, eNone, eNone, 1000, True, "vo_superjackpot"
        Addscore2 FreddySJValue
        FreddySJValue = 1000000
        li060.State = 0
        LightEffect 2
        GiEffect 2
    End If
    'Modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 5, 10, 13, 14
				TargetModeHits = TargetModeHits + 1
				CheckWinMode
			Case 15
				AwardJackpot
		End Select
	End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger009"
End Sub

'****************************
' extra triggers - no sound
'****************************

Sub Gate001_Hit 'superskillshot
    If Tilted Then Exit Sub
    Addscore 1000
    If bSkillShotReady Then AwardSuperSkillshot
End Sub

Sub Trigger011_Hit 'cabin playfield , only active when the ball move upwards
    If Tilted OR ActiveBall.VelY > 0 Then Exit Sub
    Addscore 5000
    If aWeaponSJactive Then     'the lit Super Jackpot light is lit, so lit the Super Jackpot Light at the right loop
        SuperJackpotTimer_Timer 'call the timer to stop the 30s timer and blinking light at the cabin
        li060.State = 2
    End If
    'Jason multiball
    If bJasonMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(4) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        LightEffect 2
        GiEffect 2
        Addscore2 ArrowMultiPlier(4) * 1000000
        If ArrowMultiPlier(4) < 5 Then
            ArrowMultiPlier(4) = ArrowMultiPlier(4) + 1
            UpdateArrowLights
        End If
    End If
    ' lock balls
    Select Case BallsInLock(CurrentPlayer)
        Case 0: 'enabled the first lock
			pupevent 830
			pupevent 836
            DMD "_", CL("LOCK IS LIT"), "_", eNone, eNone, eNone, 1000, True, "vo_lockislit"
            BallsInLock(CurrentPlayer) = 1
            UpdateLockLights
        Case 1: 'lock 1
			pupevent 828
			pupevent 835
            DMD "_", CL("BALL 1 LOCKED"), "_", eNone, eNone, eNone, 1000, True, "vo_ball1locked"
            BallsInLock(CurrentPlayer) = 2
            UpdateLockLights
        Case 2: 'lock 2
			pupevent 828
			pupevent 835
            DMD "_", CL("BALL 2 LOCKED"), "_", eNone, eNone, eNone, 1000, True, "vo_ball2locked"
            BallsInLock(CurrentPlayer) = 3
            UpdateLockLights
        Case 3 'lock 3 - start multiball if there is not a multiball already
            If NOT bMultiBallMode Then
                DMD "_", CL("BALL 3 LOCKED"), "_", eNone, eNone, eNone, 1000, True, "vo_ball3locked"
                BallsInLock(CurrentPlayer) = 4
                UpdateLockLights
                lighteffect 2
                Flasheffect 5
                StartJasonMultiball
            End If
    End Select
    'Modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 3
				If li066.State Then
					TargetModeHits = TargetModeHits + 1
					li066.State = 0
					CheckWinMode
				End If
			Case 5, 13, 14
				TargetModeHits = TargetModeHits + 1
				CheckWinMode
				If ReadyToKill Then 'kill her :)
					WinMode
				End If
			Case 15
				AwardJackpot
		End Select
	End If
End Sub

Sub UpdateLockLights
    Select Case BallsInLock(CurrentPlayer)
        Case 0:li071.State = 0:li072.State = 0:li073.State = 0
        Case 1:li073.State = 2                                 'enabled the first lock
        Case 2:li073.State = 1:li072.State = 2                 'lock 1
        Case 3:li072.State = 1:li071.State = 2                 'lock 2
        Case 4:li071.State = 0:li072.State = 0:li073.State = 0 'lock 3
    End Select
End Sub

Sub Trigger012_Hit 'left spinner - W1
    DOF 206, DOFPulse
    If Tilted Then Exit Sub
    FlashForMs F5, 1000, 75, 0
    If B2sOn then
        SlingCount = 0
        SlingTimer.Enabled = 1
    End If
    'weapon hit
    If WeaponHits(CurrentPlayer, 1) Then 'if the light is lit then turn it off
        WeaponHits(CurrentPlayer, 1) = 0
        CheckWeapons
    End If
    'teen killed
    If F001.State Then TeenKilled:F001.State = 0
    'Jason multiball
    If bJasonMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(1) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        LightEffect 2
        GiEffect 2
        Addscore2 ArrowMultiPlier(1) * 1000000
        If ArrowMultiPlier(1) < 5 Then
            ArrowMultiPlier(1) = ArrowMultiPlier(1) + 1
            UpdateArrowLights
        End If
    End If
    'Michael multiball
    If bMichaelMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        Addscore2 1000000
        MichaelSJValue = MichaelSJValue + 2000000
        If MichaelSJValue >= 5000000 Then
            li032.State = 2
            Select Case RndNbr(10)
                Case 1:DMD "_", "SPINNER JACKPOTS LIT", "_", eBlink, eNone, eNone, 1000, True, "vo_getthestupidjackpot"
                Case Else:DMD "_", "SPINNER JACKPOTS LIT", "_", eBlink, eNone, eNone, 1000, True, "vo_getthesuperjackpot"
            End Select
            LightEffect 2
            GiEffect 2
        End If
    End If
    'Modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 3
				If li062.State Then
					TargetModeHits = TargetModeHits + 1
					li062.State = 0
					CheckWinMode
				End If
			Case 5, 13, 14
				TargetModeHits = TargetModeHits + 1
				CheckWinMode
			Case 15
				AwardJackpot
		End Select
	End If
End Sub

Sub Trigger013_Hit 'behind right spinner
    DOF 206, DOFPulse
    If Tilted Then Exit Sub
    FlashForMs F4, 1000, 75, 0
    If B2sOn then
        SlingCount = 0
        SlingTimer.Enabled = 1
    End If
    'weapon Hit
    If WeaponHits(CurrentPlayer, 6) Then 'if the light is lit then turn it off
        WeaponHits(CurrentPlayer, 6) = 0
        CheckWeapons
    End If
    If li077.State Then 'give the special, which in this table is an add-a-ball
        PlaySound "vo_special"
        AddMultiball 1
        li077.State = 0
    End If
    'Jason multiball
    If bJasonMBStarted Then
        DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(8) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        LightEffect 2
        GiEffect 2
        Addscore2 ArrowMultiPlier(8) * 1000000
        If ArrowMultiPlier(8) < 5 Then
            ArrowMultiPlier(8) = ArrowMultiPlier(8) + 1
            UpdateArrowLights
        End If
    End If
    'Michael multiball
    If bMichaelMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        Addscore2 1000000
        MichaelSJValue = MichaelSJValue + 2000000
        If MichaelSJValue >= 5000000 Then
            li032.State = 2
            Select Case RndNbr(10)
                Case 1:DMD "_", "SPINNER JACKPOTS LIT", "_", eBlink, eNone, eNone, 1000, True, "vo_getthestupidjackpot"
                Case Else:DMD "_", "SPINNER JACKPOTS LIT", "_", eBlink, eNone, eNone, 1000, True, "vo_getthesuperjackpot"
            End Select
            LightEffect 2
            GiEffect 2
        End If
    End If
    'Modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select case NewMode
			Case 5, 13, 14
				TargetModeHits = TargetModeHits + 1
				CheckWinMode
			Case 15
				AwardJackpot
		End Select
	End If
End Sub

Sub Trigger014_Hit 'center spinner for loop awards - W4
    DOF 206, DOFPulse
    If Tilted Then Exit Sub
    FlashForMs F3, 1000, 75, 0
    If B2sOn then
        SlingCount = 0
        SlingTimer.Enabled = 1
    End If
    If LastSwitchHit = "Trigger008" Then
        AwardLoop
    Else
        li061.State = 2 'super loops light
        LoopCount = 1
    End If
    If F005.State Then TeenKilled:F005.State = 0
    'weapon hit
    If WeaponHits(CurrentPlayer, 4) Then 'if the light is lit then turn it off
        WeaponHits(CurrentPlayer, 4) = 0
        CheckWeapons
    End If
    'Jason multiball
    If bJasonMBStarted Then
        DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(5) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        LightEffect 2
        GiEffect 2
        Addscore2 ArrowMultiPlier(5) * 1000000
        If ArrowMultiPlier(5) < 5 Then
            ArrowMultiPlier(5) = ArrowMultiPlier(5) + 1
            UpdateArrowLights
        End If
    End If
    'Michael multiball
    If bMichaelMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        Addscore2 1000000
        MichaelSJValue = MichaelSJValue + 2000000
        If MichaelSJValue >= 5000000 Then
            li032.State = 2
            Select Case RndNbr(10)
                Case 1:DMD "_", "SPINNER JACKPOTS LIT", "_", eBlink, eNone, eNone, 1000, True, "vo_getthestupidjackpot"
                Case Else:DMD "_", "SPINNER JACKPOTS LIT", "_", eBlink, eNone, eNone, 1000, True, "vo_getthesuperjackpot"
            End Select
            LightEffect 2
            GiEffect 2
        End If
    End If
    'Modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 3
				If li067.State Then
					TargetModeHits = TargetModeHits + 1
					li067.State = 0
					CheckWinMode
				End If
			Case 5, 6, 10, 13, 14
				TargetModeHits = TargetModeHits + 1
				CheckWinMode
			Case 15
				AwardJackpot
		End Select
	End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger014"
End Sub

Sub Trigger015_Hit 'right ramp done - W5
    If Tilted Then Exit Sub
    Addscore 5000
	pupevent 826
    If LastSwitchHit = "Trigger015" Then
        AwardCombo
    Else
        ComboCount = 1
    End If
    If F003.State Then TeenKilled:F003.State = 0
    'weapon hit
    If WeaponHits(CurrentPlayer, 5) Then 'if the light is lit then turn it off
        WeaponHits(CurrentPlayer, 5) = 0
        CheckWeapons
    End If
    If bTommyStarted Then 'the Tommy Jarvis hurry up is started so award the jackpot
        AwardTommyJackpot
        StopTommyJarvis
    End If
    'Jason multiball
    If bJasonMBStarted Then
        DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(6) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        LightEffect 2
        GiEffect 2
        Addscore2 ArrowMultiPlier(6) * 1000000
        If ArrowMultiPlier(6) < 5 Then
            ArrowMultiPlier(6) = ArrowMultiPlier(6) + 1
            UpdateArrowLights
        End If
    End If
    'Freddy multiball
    If bFreddyMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        Addscore2 1000000
        FreddySJValue = FreddySJValue + 2000000
        If FreddySJValue >= 5000000 Then
            li060.State = 2
            Select Case RndNbr(10)
                Case 1:DMD "_", CL("SUPERJACKPOT IS LIT"), "_", eBlink, eNone, eNone, 1000, True, "vo_getthestupidjackpot"
                Case Else:DMD "_", CL("SUPERJACKPOT IS LIT"), "_", eBlink, eNone, eNone, 1000, True, "vo_getthesuperjackpot"
            End Select
            LightEffect 2
            GiEffect 2
        End If
    End If
    'Modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 3
				If li068.State Then
					TargetModeHits = TargetModeHits + 1
					li068.State = 0
					CheckWinMode
				End If
			Case 5, 10, 13, 14
				TargetModeHits = TargetModeHits + 1
				CheckWinMode
			Case 15
				AwardJackpot
		End Select
	End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger015"
End Sub

Sub Trigger016_Hit 'left ramp done - W2
    DOF 206, DOFPulse
    If Tilted Then Exit Sub
    Addscore 5000
    If LastSwitchHit = "Trigger016" Then
        AwardCombo
    Else
        ComboCount = 1
    End If
    If F004.State Then TeenKilled:F004.State = 0
    If li069.State Then AwardTargetJackpot
    'weapon hit
    If WeaponHits(CurrentPlayer, 2) Then 'if the light is lit then turn it off
        WeaponHits(CurrentPlayer, 2) = 0
        CheckWeapons
    End If
    'Jason multiball
    If bJasonMBStarted Then
        DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(2) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        LightEffect 2
        GiEffect 2
        Addscore2 ArrowMultiPlier(2) * 1000000
        If ArrowMultiPlier(2) < 5 Then
            ArrowMultiPlier(2) = ArrowMultiPlier(2) + 1
            UpdateArrowLights
        End If
    End If
    'Freddy multiball
    If bFreddyMBStarted Then
        DOF 127, DOFPulse
        DMD CL("JACKPOT"), CL(FormatScore(1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
        Addscore2 1000000
        FreddySJValue = FreddySJValue + 2000000
        If FreddySJValue >= 5000000 Then
            li060.State = 2
            Select Case RndNbr(10)
                Case 1:DMD "_", CL("SUPERJACKPOT IS LIT"), "_", eBlink, eNone, eNone, 1000, True, "vo_getthestupidjackpot"
                Case Else:DMD "_", CL("SUPERJACKPOT IS LIT"), "_", eBlink, eNone, eNone, 1000, True, "vo_getthesuperjackpot"
            End Select
            LightEffect 2
            GiEffect 2
        End If
    End If
    'Modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 3
				If li064.State Then
					TargetModeHits = TargetModeHits + 1
					li064.State = 0
					CheckWinMode
				End If
			Case 5, 6, 10, 13, 14
				TargetModeHits = TargetModeHits + 1
				CheckWinMode
		End Select
	End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger016"
End Sub

Sub Trigger017_Hit 'cabin top
    DOF 204, DOFPulse
    If Tilted Then Exit Sub
    ChangeGi red
    ChangeGIIntensity 2
    GiEffect 1
    FlashEffect 1
    PlaySound "sfx_dialphone"
    FlashForms FlasheOndasCabina, 325, 50, 0 
    vpmTimer.AddTimer 2500, "ChangeGi white:ChangeGIIntensity 1 '"
    ' Select Mode
    Select Case Mode(CurrentPlayer, 0)
        Case 0:SelectMode 'no mode is active then activate another mode
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger017"
End Sub

Sub Trigger018_Hit 'left loop - only light effect
    '	DOF 204, DOFPulse
    If Tilted Then Exit Sub
    Flashforms f1A, 800, 50, 0
    Flashforms F1B, 800, 50, 0
    Flashforms Flasher006, 800, 50, 0
    Flashforms Flasher007, 800, 50, 0
End Sub

Sub Triggerholograma_Hit
    If LastSwitchHit = "Trigger017" Then
    Flashforms Flasher009, 800, 50, 0
    Flashforms FlasherManos004, 600, 50, 1
    Flashforms FlasherManos003, 600, 50, 1
    Flashforms FlasherAsciiGigante, 300, 50, 0
    vpmtimer.addtimer 300, "Flashforms FlashThematrix, 800, 50, 1 '"
    vpmtimer.addtimer 500, "Flashforms FlasherBala, 800, 125, 0 '"
    vpmtimer.addtimer 600, "Flashforms FlasherBala2, 900, 125, 0 '"
    vpmtimer.addtimer 700, "Flashforms FlasherBala3, 1000, 125, 0 '"
    vpmtimer.addtimer 500, "PlaySound""sfx_balacomplex"" '"
    End If

End Sub


'*****************
'SENTINEL FLASHER'
'*****************



Dim SentinelFlashEstado

    SentinelFlashEstado = 0
    

Sub FlasherGrandeSentinelMode_timer
    SentinelFlashEstado = ABS(SentinelFlashEstado - 1)
    FlasherGrandeSentinelMode.Visible = SentinelFlashEstado

End Sub


   ' FLASHER OJOS BLONK

    Dim FlasherOjosEstado

    FlasherOjosEstado = 0

    FlasherManos003.TimerEnabled = 1

Sub TimerOjos1_timer

    FlasherOjosEstado = ABS(FlasherOjosEstado - 1)

    FlasherManos003.Visible = FlasherOjosEstado

End Sub

' FLASHER OJOS BLONK 2

    Dim FlasherOjosEstado2

    FlasherOjosEstado2 = 0

    FlasherManos004.TimerEnabled = 1

 

Sub TimerOjos2_timer

    FlasherOjosEstado2 = ABS(FlasherOjosEstado2 - 1)

    FlasherManos004.Visible = FlasherOjosEstado2

End Sub


   ' FLASHER OJOS BLONK 3

    Dim FlasherOjosEstado3

    FlasherOjosEstado3 = 0

    FlasherManos016.TimerEnabled = 1

 

Sub TimerOjosBlink016_timer

    FlasherOjosEstado3 = ABS(FlasherOjosEstado3 - 1)

    FlasherManos016.Visible = FlasherOjosEstado3

End Sub

'*************************
   'FLASHER PULSERA BLINK MANO
'*************************



    FlasherPulseraMano.TimerEnabled = 1

 

Sub TimerPulsera_timer

    Flashforms FlasherPulseraMano, 300, 20, 1

End Sub

'*********************
' gira pulsera timer
'*********************
Dim RotAngle
RotAngle = 0

Sub TimerGiraPuls_Timer
    RotAngle = (RotAngle+ 1)MOD 360
    FlasherPulseraMano.RotZ = RotAngle
    End Sub

'***************
'   Giro Bebe 
'***************

Dim RotAngle3
RotAngle3 = 0

Sub QueridoBebe_Timer
    RotAngle3 = (RotAngle3+ 1)MOD 360
    Bebe1.Roty = RotAngle3
    End Sub
    
'********************
' GIRA FONDO LOGO
'********************

Dim RotAngle5
RotAngle5 = 0

Sub TimerParaFondoLogo_Timer
    RotAngle5 = (RotAngle5+ 1)MOD 360
    FlasherVPCOOKS1.Rotz = RotAngle5
    End Sub




'********************
' GIRA FONDO vp logo
'********************

Dim RotAngle4
RotAngle4 = 0

Sub TimerVueltaLogo_Timer
    RotAngle4 = (RotAngle4+ 1)MOD 360
    FlasherlogoVP.RotX = RotAngle4
    End Sub


'************
'  Targets
'************

Sub Target001_Hit 'police
    PLaySoundAtBall SoundFXDOF("fx_Target",205,DOFPulse,DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5000
    Flashforms f5, 800, 50, 0
    Flashforms Flasher001, 800, 50, 0
    PlayElectro
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2, 7, 8, 11
				If li031.State Then
					TargetModeHits = TargetModeHits + 1
					li031.State = 0
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 1) = 0
				CheckTargets
		End Select
	End If
    If NOT bPoliceStarted Then
        PoliceTargetHits = PoliceTargetHits + 1
        If PoliceTargetHits = 3 Then
            StartPolice
        End If
    End If
End Sub

Sub Target002_Hit 'right target 1
    PLaySoundAtBall SoundFXDOF("fx_Target",205,DOFPulse,DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5000
    PlayElectro
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2, 7, 8, 11
				If li057.State Then
					TargetModeHits = TargetModeHits + 1
					li057.State = 0
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 6) = 0
				CheckTargets
		End Select
	End If
End Sub

Sub Target003_Hit 'right target 2
    PLaySoundAtBall SoundFXDOF("fx_Target",205,DOFPulse,DOFTargets)
    Addscore 5000
    If Tilted Then Exit Sub
    PlayElectro
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2, 7, 8, 11
				If li058.State Then
					TargetModeHits = TargetModeHits + 1
					li058.State = 0
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 5) = 0
				CheckTargets
		End Select
	End If
End Sub

Sub Target004_Hit 'cabin target 1
    PLaySoundAtBall SoundFXDOF("fx_Target",109,DOFPulse,DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5000
    PlayElectro
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2, 7, 8, 11
				If li049.State Then
					TargetModeHits = TargetModeHits + 1
					li049.State = 0
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 2) = 0
				CheckTargets
		End Select
	End If
End Sub

Sub Target005_Hit 'cabin target 2
    PLaySoundAtBall SoundFXDOF("fx_Target",109,DOFPulse,DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5000
    PlayElectro
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2, 7, 8, 11
				If li050.State Then
					TargetModeHits = TargetModeHits + 1
					li050.State = 0
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 3) = 0
				CheckTargets
		End Select
	End If
End Sub

Sub Target006_Hit 'loop target
    PLaySoundAtBall SoundFXDOF("fx_Target",109,DOFPulse,DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5000
    PlayElectro
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2, 7, 8, 11
				If li051.State Then
					TargetModeHits = TargetModeHits + 1
					li051.State = 0
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 4) = 0
				CheckTargets
		End Select
	End If
End Sub

Sub Target007_Hit 'magnet target W3
    PLaySoundAtBall SoundFXDOF("fx_Target",107,DOFPulse,DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5000
    'weapon hit
    If WeaponHits(CurrentPlayer, 3) Then 'if the light is lit then turn it off
        WeaponHits(CurrentPlayer, 3) = 0
        CheckWeapons
    End If
    'extra ball
    If li070.State Then AwardExtraBall:li070.State = 0
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 2
				If li065.State Then
					WinMode
				End If
			Case 9
				If li065.State Then
					TargetModeHits = TargetModeHits + 1
					CheckWinMode
				End If
			Case Else
				TargetHits(CurrentPlayer, 1) = 0
				CheckTargets
		End Select
	End If
End Sub

Sub CheckTargets
    Dim tmp, i
    tmp = 0
    UpdateTargetLights
    For i = 1 to 7
        tmp = tmp + TargetHits(CurrentPlayer, i)
    Next
    If tmp = 0 then 'all targets are hit so turn on the target Jackpot
        DMD "_", CL("TARGET JACKPOT S LIT"), "", eNone, eNone, eNone, 1000, True, "vo_shoottheleftramp"
        li069.State = 1
        LightSeqBLueTargets.Play SeqBlinking, , 15, 25
        For i = 1 to 7 'and reset them
            TargetHits(CurrentPlayer, i) = 1
        Next
        UpdateTargetLights
    End If
End Sub

'*************
'  Spinners
'*************

Sub Spinner001_Spin 'right
    PlaySoundAt "fx_spinner", Spinner001
    DOF 200, DOFPulse
    If Tilted Then Exit Sub
    Addscore 1000
    RightSpinnerHits(CurrentPlayer) = RightSpinnerHits(CurrentPlayer) + 1
    ' check for add-a-aball during multiballs or during normal play
    If RightSpinnerHits(CurrentPlayer) >= 100 Then
        LitSpecial
        RightSpinnerHits(CurrentPlayer) = 0
    End If
    CheckSpinners
    'chek modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 1
				If SpinCount < SpinNeeded Then
					SpinCount = SpinCount + 1
					CheckWinMode
				End If
			Case 12
				If li059.State AND SpinCount < SpinNeeded Then
					SpinCount = SpinCount + 1
					CheckWinMode
				End If
		End Select
	End If
End Sub

Sub LitSpecial
    DMD "_", CL("SPECIAL IS LIT"), "", eNone, eNone, eNone, 1000, True, "vo_specialislit"
    li077.State = 1
End Sub

Sub Spinner002_Spin 'center
    PlaySoundAt "fx_spinner", Spinner002
    DOF 201, DOFPulse
    If Tilted Then Exit Sub
    Addscore 1000
    CenterSpinnerHits(CurrentPlayer) = CenterSpinnerHits(CurrentPlayer) + 1
    ' check for Bonus multiplier
    If CenterSpinnerHits(CurrentPlayer) >= 30 Then
        li074.State = 1
    End If
    If CenterSpinnerHits(CurrentPlayer) >= 40 Then
        AddBonusMultiplier 1
        CenterSpinnerHits(CurrentPlayer) = 0
        li074.State = 0
    End If
    CheckSpinners
    'chek modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 1
				If SpinCount < SpinNeeded Then
					SpinCount = SpinCount + 1
					CheckWinMode
				End If
			Case 12
				If li067.State AND SpinCount < SpinNeeded Then
					SpinCount = SpinCount + 1
					CheckWinMode
				End If
		End Select
	End If
End Sub

Sub Spinner003_Spin 'left
    PlaySoundAt "fx_spinner", Spinner003
    DOF 202, DOFPulse
    If Tilted Then Exit Sub
    Addscore 1000
    LeftSpinnerHits(CurrentPlayer) = LeftSpinnerHits(CurrentPlayer) + 1
    CheckSpinners
    'chek modes
	if Mode(CurrentPlayer, 0) = NewMode Then
		Select Case NewMode
			Case 1
				If SpinCount < SpinNeeded Then
					SpinCount = SpinCount + 1
					CheckWinMode
				End If
			Case 12
				If li062.State AND SpinCount < SpinNeeded Then
					SpinCount = SpinCount + 1
					CheckWinMode
				End If
		End Select
	End If
End Sub

Sub CheckSpinners
End Sub


'*********
' scoop
'*********

Dim aBall

Sub scoop_Hit
    PlaySoundAt "fx_hole_enter", scoop
    BallsinHole = BallsInHole + 1
    Set aBall = ActiveBall
    scoop.TimerEnabled = 1
    If Tilted Then Exit Sub
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    ' kick out the ball during hurry-ups
    If bTommyStarted OR bPoliceStarted Then vpmtimer.addtimer 500, "kickBallOut '"
    ' check for modes
    Addscore 5000
    Flashforms f4, 800, 50, 0
    Flashforms Flasher003, 800, 50, 0
    ' check modes
    If(Mode(CurrentPlayer, NewMode) = 2) AND(Mode(CurrentPlayer, 0) = 0) Then 'the mode is ready, so start it
		pupevent 827
		vpmtimer.addtimer 4000, "StartMode '"
    'PlaySound "sfx_gritomatrix"   
	
    Flashforms FlasherAsciiGigante, 200, 40, 0
        Exit Sub
    End If
    If li078.State Then
        AwardMystery 'after the award the ball will be kicked out
        li078.State = 0
    Else
        Select Case NewMode
            Case 1, 3
                If ReadyToKill Then
                    WinMode
                Else
                    vpmtimer.addtimer 2500, "kickBallOut '"
                End If  
            Case Else
                ' Nothing left to do, so kick out the ball
                vpmtimer.addtimer 2500, "kickBallOut '"
        End Select
    End If
End Sub

Sub scoop_Timer
    If aBall.Z > 0 Then
        aBall.Z = aBall.Z -5
    Else
        scoop.Destroyball
        Me.TimerEnabled = 0
        If Tilted Then kickBallOut
    End If
End Sub

Sub kickBallOut
    If BallsinHole > 0 Then
        BallsinHole = BallsInHole - 1
        PlaySoundAt SoundFXDOF("fx_popper", 111, DOFPulse, DOFcontactors), scoop
        DOF 124, DOFPulse
        scoop.CreateSizedBallWithMass BallSize / 2, BallMass
        scoop.kick 235, 22, 1
        Flashforms F4, 500, 50, 0
        vpmtimer.addtimer 400, "kickBallOut '" 'kick out the rest of the balls, if any
    End If
End Sub

'*************
' Magnet
'*************

Sub Trigger007_Hit
    DOF 206, DOFPulse
    If Tilted Then Exit Sub
    If ActiveBall.VelY > 10 Then
        ActiveBall.VelY = 10
    End If
    mMagnet.MagnetOn = True
    DOF 112, DOFOn
    Me.TimerEnabled = 1 'to turn off the Magnet
    'Jason multiball
    If bJasonMBStarted Then
        If ActiveBall.VelY < 0 Then 'this means the ball going up
            DMD CL("JACKPOT"), CL(FormatScore(ArrowMultiPlier(3) * 1000000) ), "_", eBlink, eNone, eNone, 1000, True, "vo_jackpot" &RndNbr(6)
            LightEffect 2
            GiEffect 2
            Addscore2 ArrowMultiPlier(3) * 1000000
            If ArrowMultiPlier(3) < 5 Then
                ArrowMultiPlier(3) = ArrowMultiPlier(3) + 1
                UpdateArrowLights
            End If
        End If
    End If
    'Modes
    If ActiveBall.VelY < 0 Then 'this means the ball going up
		if Mode(CurrentPlayer, 0) = NewMode Then
			Select Case NewMode
				Case 5, 13, 14
					TargetModeHits = TargetModeHits + 1
					CheckWinMode
				Case 15
					AwardJackpot
			End Select
		End If
	End If
End Sub

Sub Trigger007_Timer
    Me.TimerEnabled = 0
    ReleaseMagnetBalls
End Sub

Sub ReleaseMagnetBalls 'mMagnet off and release the ball if any
    Dim ball
    mMagnet.MagnetOn = False
    DOF 112, DOFOff
    For Each ball In mMagnet.Balls
        With ball
            .VelX = 0
            .VelY = 1
        End With
    Next
End Sub

'*******************
'    RAMP COMBOS
'*******************
' don't time out
' starts at 500K for a 2 way combo and it is doubled on each combo
' shots that count as ramp combos:
' Left Ramp and Right Ramp
' shots to the same ramp also count as loops

Sub AwardCombo
    ComboCount = ComboCount + 1
    Select Case ComboCount
        Case 1: 'just starting
        Case 2:DMD CL("COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_combo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 3:DMD CL("2X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 2) ), "", eNone, eNone, eNone, 1500, True, "vo_2xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 4:DMD CL("3X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 3) ), "", eNone, eNone, eNone, 1500, True, "vo_3xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 5:DMD CL("4X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 4) ), "", eNone, eNone, eNone, 1500, True, "vo_4xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 6:DMD CL("5X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 5) ), "", eNone, eNone, eNone, 1500, True, "vo_5xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 7:DMD CL("SUPER COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 7) ), "", eNone, eNone, eNone, 1500, True, "vo_supercombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1:DOF 126, DOFPulse
        Case Else:DMD CL("RED PILL COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 10) ), "", eNone, eNone, eNone, 1500, True, "vo_superdupercombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1:DOF 126, DOFPulse
    End Select
    AddScore2 ComboValue(CurrentPlayer) * ComboCount
    ComboValue(CurrentPlayer) = ComboValue(CurrentPlayer) + 100000
End Sub

Sub aComboTargets_Hit(idx) 'reset the combo count if the ball hits another target/trigger
    ComboCount = 0
End Sub

'*******************
'    LOOP COMBOS
'*******************
' starts at 500K for a 2 way combo and it is doubled on each combo
' shots that count as loop combos:
' Center loop and Right Loop
' shots to the same loop also count as loops

Sub AwardLoop
    LoopCount = LoopCount + 1
    Select Case LoopCount
        Case 1: 'just starting
        Case 2:DMD CL("COMBO"), CL(FormatScore(LoopValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_combo":LoopHits(CurrentPlayer) = LoopHits(CurrentPlayer) + 1
        Case 3:DMD CL("2X COMBO"), CL(FormatScore(LoopValue(CurrentPlayer) * 2) ), "", eNone, eNone, eNone, 1500, True, "vo_2xcombo":LoopHits(CurrentPlayer) = LoopHits(CurrentPlayer) + 1
        Case 4:DMD CL("3X COMBO"), CL(FormatScore(LoopValue(CurrentPlayer) * 3) ), "", eNone, eNone, eNone, 1500, True, "vo_3xcombo":LoopHits(CurrentPlayer) = LoopHits(CurrentPlayer) + 1
        Case 5:DMD CL("4X COMBO"), CL(FormatScore(LoopValue(CurrentPlayer) * 4) ), "", eNone, eNone, eNone, 1500, True, "vo_4xcombo":LoopHits(CurrentPlayer) = LoopHits(CurrentPlayer) + 1
        Case 6:DMD CL("5X COMBO"), CL(FormatScore(LoopValue(CurrentPlayer) * 5) ), "", eNone, eNone, eNone, 1500, True, "vo_5xcombo":LoopHits(CurrentPlayer) = LoopHits(CurrentPlayer) + 1
        Case 7:DMD CL("SUPER COMBO"), CL(FormatScore(LoopValue(CurrentPlayer) * 7) ), "", eNone, eNone, eNone, 1500, True, "vo_supercombo":LoopHits(CurrentPlayer) = LoopHits(CurrentPlayer) + 1
        Case Else:DMD CL("RED PILL COMBO"), CL(FormatScore(LoopValue(CurrentPlayer) * 10) ), "", eNone, eNone, eNone, 1500, True, "vo_superdupercombo":LoopHits(CurrentPlayer) = LoopHits(CurrentPlayer) + 1
    End Select
    AddScore2 LoopValue(CurrentPlayer) * LoopCount
    LoopValue(CurrentPlayer) = LoopValue(CurrentPlayer) + 100000
End Sub

Sub aLoopTargets_Hit(idx) 'reset the loop count if the ball hits another target/trigger
    li061.State = 0       'turn off also the super loops light
    LoopCount = 0
End Sub

'*******************
'  Teenager kill / Clone Kill
'*******************

' the timer will change the current teenager by lighting the light on top of her

Sub TeenTimer_Timer
    Select Case RndNbr(15)
        Case 1:F001.State = 2:F002.State = 0:F003.State = 0:F004.State = 0:F005.State = 0:PlayQuote:DOF 204, DOFPulse
        Case 2:F001.State = 0:F002.State = 2:F003.State = 0:F004.State = 0:F005.State = 0:PlayQuote:DOF 204, DOFPulse
        Case 3:F001.State = 0:F002.State = 0:F003.State = 2:F004.State = 0:F005.State = 0:PlayQuote:DOF 204, DOFPulse
        Case 4:F001.State = 0:F002.State = 0:F003.State = 0:F004.State = 2:F005.State = 0:PlayQuote:DOF 204, DOFPulse
        Case 5:F001.State = 0:F002.State = 0:F003.State = 0:F004.State = 0:F005.State = 2:PlayQuote:DOF 204, DOFPulse
        Case Else:F001.State = 0:F002.State = 0:F003.State = 0:F004.State = 0:F005.State = 0
    End Select
End Sub

Sub TeenKilled 'a teen has been killed
    'DMD animation
    DMD "", "", "d_goa", eNone, eNone, eNone, 150, False, "sfx_kill" &RndNbr(10)
    DMD "", "", "d_gob", eNone, eNone, eNone, 150, False, ""
    DMD "", "", "d_goc", eNone, eNone, eNone, 150, False, ""
    DMD "", "", "d_god", eNone, eNone, eNone, 150, False, ""
    DMD "", "", "d_gof", eNone, eNone, eNone, 150, False, ""
    DMD CL("Y"), "", "d_go1", eNone, eNone, eNone, 100, False, ""
    DMD CL("YO"), "", "d_go2", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU"), "", "d_go3", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU "), "", "d_go4", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU K"), "", "d_go5", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KI"), "", "d_go6", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KIL"), "", "d_go7", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLE"), "", "d_go8", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), "", "d_go9", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A"), "d_go10", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A "), "d_go11", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A C"), "d_go12", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A CL"), "d_go13", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A CLO"), "d_go14", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A CLON"), "d_go15", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A CLONE"), "d_go16", eNone, eNone, eNone, 100, False, ""
    DMD CL("YOU KILLED"), CL("A CLONE"), "d_go17", eNone, eNone, eNone, 100, False, ""
    
    Select Case RndNbr(3)
        Case 1:DMD CL("CLONE DESTROYED"), CL(FormatScore(TeensKilledValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_excellent"
        Case 2:DMD CL("CLONE DESTROYED"), CL(FormatScore(TeensKilledValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_greatshot"
        Case 3:DMD CL("CLONE DESTROYED"), CL(FormatScore(TeensKilledValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_notbad"
    End Select
    TeensKilled(CurrentPlayer) = TeensKilled(CurrentPlayer) + 1
    AddScore2 TeensKilledValue(CurrentPlayer)
    TeensKilledValue(CurrentPlayer) = TeensKilledValue(CurrentPlayer) + 50000
    LightEffect 5
    GiEffect 5
    'check
    If TeensKilled(CurrentPlayer) MOD 6 = 0 Then StartTommyJarvis 'start Tommy Jarvis hurry up after each 6th killed teenager
       'start police after each 4th killed teenager
    If TeensKilled(CurrentPlayer) MOD 10 = 0 Then LitExtraBall    'lit the extra ball
End Sub

Sub LitExtraBall
	
    DMD "_", CL("EXTRA BALL IS LIT"), "", eNone, eNone, eNone, 1500, True, "vo_extraballislit"
    li070.State = 1
End Sub

'********************************
' Weapons - Playfield multiplier
'********************************

Sub CheckWeapons
    Dim a, j
    a = 0
    For j = 1 to 6
        a = a + WeaponHits(CurrentPlayer, j)
    Next
    'debug.print a
    If a = 0 Then
        UpgradeWeapons
    Else
        LightSeqWeaponLights.UpdateInterval = 25
        LightSeqWeaponLights.Play SeqBlinking, , 15, 25
        UpdateWeaponLights
        PlaySword
    End If
End Sub

Sub UpdateWeaponLights 'CurrentPlayer
    Li035.State = WeaponHits(CurrentPlayer, 1)
    li037.State = WeaponHits(CurrentPlayer, 2)
    li038.State = WeaponHits(CurrentPlayer, 3)
    li039.State = WeaponHits(CurrentPlayer, 4)
    li040.State = WeaponHits(CurrentPlayer, 5)
    Li036.State = WeaponHits(CurrentPlayer, 6)
End Sub

Sub ResetWeaponLights 'CurrentPlayer
    Dim j
    For j = 1 to 6
        WeaponHits(CurrentPlayer, j) = 1
    Next
    UpdateWeaponLights
End Sub

Sub UpgradeWeapons 'increases the playfield multiplier
    Weapons(CurrentPlayer) = Weapons(CurrentPlayer) + 1
    UpdateWeaponLights2
    AddPlayfieldMultiplier 1
    ResetWeaponLights
    aWeaponSJactive = True
    StartSuperJackpot   'turn on the lits SJ at the cabin
End Sub

Sub UpdateWeaponLights2 'collected weapons
    Select Case Weapons(CurrentPlayer)
        Case 1:li041.State = 1:li042.State = 0:li043.State = 0:li044.State = 0:li045.State = 0:li046.State = 0:li047.State = 0
        Case 2:li041.State = 1:li042.State = 1:li043.State = 0:li044.State = 0:li045.State = 0:li046.State = 0:li047.State = 0
        Case 3:li041.State = 1:li042.State = 1:li043.State = 1:li044.State = 0:li045.State = 0:li046.State = 0:li047.State = 0
        Case 4:li041.State = 1:li042.State = 1:li043.State = 1:li044.State = 1:li045.State = 0:li046.State = 0:li047.State = 0
        Case 5:li041.State = 1:li042.State = 1:li043.State = 1:li044.State = 1:li045.State = 1:li046.State = 0:li047.State = 0
        Case 6:li041.State = 1:li042.State = 1:li043.State = 1:li044.State = 1:li045.State = 1:li046.State = 1:li047.State = 0
        Case 7:li041.State = 1:li042.State = 1:li043.State = 1:li044.State = 1:li045.State = 1:li046.State = 1:li047.State = 1
    End Select
End Sub

'*******************************************
' Super Jackpot at the cabin and right Loop
'*******************************************

' 30 seconds timer to lit the Super Jackpot at the right loop
' once the Super Jackpot light is lit it will be lit until the end of the ball

Sub StartSuperJackpot 'lits the cabin's red SJ light
    li048.BlinkInterval = 160
    li048.State = 2
    SuperJackpotTimer.Enabled = 1
    SuperJackpotSpeedTimer.Enabled = 1
End Sub

Sub SuperJackpotTimer_Timer            '30 seconds hurry up to turn off the red light at the cabin
    SuperJackpotSpeedTimer.Enabled = 0 'to be sure it is stopped
    SuperJackpotTimer.Enabled = 0
    li048.State = 0
End Sub

Sub SuperJackpotSpeedTimer_Timer 'after 25 seconds speed opp the blinking of the lit SJ red cabin light
    DMD "_", CL("HURRY UP NEO"), "_", eNone, eBlink, eNone, 1500, True, "vo_hurryup"
    SuperJackpotSpeedTimer.Enabled = 0
    li048.BlinkInterval = 80
    li048.State = 2
End Sub

'*************************
' Tommy Jarvis - Hurry up
'*************************
' it starts after each 6 teenagers killed
' a 30 seconds hurry up starts at the right ramp
' hit the right ramp to throw a deadly blow to your archie enemy
' fail and your flippers will die for 3 seconds

Sub StartTommyJarvis 'Tommy's hurry up
    DMD CL("DEJA VU"), CL("HURRY UP"), "d_border", eNone, eNone, eNone, 2500, True, "vo_shoottherightramp" &RndNbr(2)
    bTommyStarted = True
    FlashEffect 7
    Li030.BlinkInterval = 160
    Li030.State = 2
    TommyCount = 1
    TommyValue = 500000
    TommyHurryUpTimer.Enabled = 1
End Sub

Sub StopTommyJarvis
    TommyHurryUpTimer.Enabled = 0
    TommyRecoverTimer.Enabled = 0
    bTommyStarted = False
    LightSeqTopFlashers.StopPlay
    Li030.State = 0
    TommyCount = 1
End Sub

Sub TommyHurryUpTimer_Timer '30 seconds, runs once a second
    TommyCount = TommyCount + 1
    If TommyCount = 20 Then 'speed up the Light
        DMD "_", CL("HURRY UP"), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_timerunningout"
        Li030.BlinkInterval = 80
        Li030.State = 2
    End If
    If TommyCount = 31 Then 'the 30 seconds are over
        TommyHurryUpTimer.Enabled = 0
        DMD CL("DEJA VU"), CL("KILLED YOU"), "d_border", eNone, eNone, eNone, 2500, True, "vo_holyshityoushootme"
        DisableFlippers True          'disable the flippers...
        ChangeGi purple
        ChangeGIIntensity 2
        TommyRecoverTimer.Enabled = 1 '... start the 3 seconds timer to enable the table again
    End If
End Sub

Sub TommyRecoverTimer_Timer 'after disabling the table for 3 seconds then enable it
    TommyRecoverTimer.Enabled = 0
    StopTommyJarvis
    DisableFlippers False
    ChangeGi white 
    ChangeGIIntensity 1
End Sub

Sub AwardTommyJackpot() 'scores the seconds multiplied by 500.000, the longer you wait the higher the score
    Dim a
    a = TommyCount * 500000
    Select Case RndNbr(3)
        Case 1:DMD CL("DEJA VU HURRY JACKPT"), CL(FormatScore(a) ), "d_border", eNone, eBlinkFast, eNone, 2500, True, "vo_waytogo"
        Case 2:DMD CL("DEJA VU HURRY JACKPT"), CL(FormatScore(a) ), "d_border", eNone, eBlinkFast, eNone, 2500, True, "vo_welldone"
        Case 3:DMD CL("DEJA VU HURRY JACKPT"), CL(FormatScore(a) ), "d_border", eNone, eBlinkFast, eNone, 2500, True, "vo_younailedit"
    End Select
    DOF 126, DOFPulse
    AddScore2 a
    LightEffect 2
    GiEffect 2
End Sub

'***********************
' The Police - Hurry up
'***********************
' similar to Tommy hurry up, but different score
' it starts after 5 police hits, 2 killed counselors or 4 teenagers
' 30 seconds to hit 2 ramps or 3 police targets
' scores 5.000.000 for each left second
' fail and your flippers will die for 3 seconds

Sub StartPolice 'police hurry up
    EnableBallSaver 30
    bPoliceStarted = True
    FlasherGrandeSentinelMode.timerenabled = 1
    SentinelKicker.Enabled = 1
    DarkNight.visible = 1
    FantasmaPosicionZ = -120 'unsure it is down, and that it will go up  
    FantasmaDireccionZ = 5
    FantasmaTimerZ.Enabled = 1
    StopSong Song
    PoliceL1.BlinkInterval = 160
    PoliceL2.BlinkInterval = 160
    Li033.BlinkInterval = 160
    PoliceL1.State = 2
    PoliceL2.State = 2
    Li033.State = 2
    PlaySound "sfx_siren1", -1
    PoliceCount = 30
    PoliceHurryUpTimer.Enabled = 1
    PoliceTargetHits = 0 'reset the count
    UpdateClock 10
End Sub

Sub StopPolice
	Pupevent 925
	Pupevent 926
    SentinelKicker.Enabled = 0
    FantasmaTimer.Enabled = 0 'stop side to side movement
    FantasmaPosicionZ = 65 'unsure it is up, and that it will go down  
    FantasmaDireccionZ = -10
    FlasherGrandeSentinelMode.timerenabled = 0
    FlasherGrandeSentinelMode.Visible = 0
    vpmtimer.addtimer 300, "FantasmaTimerZ.Enabled = 1:DarkNight.visible = 0 '"
    ChangeSong
    PoliceHurryUpTimer.Enabled = 0
    PoliceRecoverTimer.Enabled = 0
    bPoliceStarted = False
    StopSound "sfx_siren1"
    PoliceL1.State = 0
    PoliceL2.State = 0
    Li033.State = 0
    PoliceTargetHits = 0     'reset the count
    PoliceCount = 0
    DisableFlippers False    'ensure they are not disabled
    ChangeGi white 
    ChangeGIIntensity 1     
    TurnOffClock     '
    DeactivarSentinel
    PlaySong "mu_Morpheus"

End Sub

Sub PoliceHurryUpTimer_Timer '30 seconds, runs once a second
    PoliceCount = PoliceCount - 1
    If PoliceCount = 8 Then  'speed up the Lights
        DMD "_", CL("HURRY UP"), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_timerunningout"
        PoliceL1.BlinkInterval = 80
        PoliceL2.BlinkInterval = 80
        Li033.BlinkInterval = 80
        PoliceL1.State = 2
        PoliceL2.State = 2
        Li033.State = 2
    End If
    If PoliceCount = 0 Then
        PoliceHurryUpTimer.Enabled = 0
        DMD CL("THE SENTINEL"), CL("WON"), "d_border", eNone, eNone, eNone, 2500, True, "vo_holyshityoushootme"
        
        ChangeGi Blue
        ChangeGIIntensity 2
        PoliceRecoverTimer.Enabled = 1 '... start the 3 seconds timer to enable the table again
    End If
End Sub

Sub PoliceRecoverTimer_Timer 'after disabling the table for 3 seconds then enable it again
    PoliceRecoverTimer.Enabled = 0
    StopPolice
End Sub

Sub AwardPoliceJackpot() 'scores the seconds left multiplied by 5.000.000
    Dim a
    DOF 126, DOFPulse
    a = PoliceCount * 5000000
    Select Case RndNbr(3)
        Case 1:DMD CL("SENTINEL JACKPOT"), CL(FormatScore(a) ), "d_border", eNone, eBlinkFast, eNone, 2500, True, "vo_waytogo"
        Case 2:DMD CL("SENTINEL JACKPOT"), CL(FormatScore(a) ), "d_border", eNone, eBlinkFast, eNone, 2500, True, "vo_welldone"
        Case 3:DMD CL("SENTINEL JACKPOT"), CL(FormatScore(a) ), "d_border", eNone, eBlinkFast, eNone, 2500, True, "vo_younailedit"
    End Select
    AddScore2 a
    LightEffect 2
    GiEffect 2
End Sub

Sub SentinelKicker_Hit
	
    PlaySound "sfx_rampamatrix"

    
    
    vpmtimer.addtimer 2500,"SentinelKicker.kick 90, 60 '"
    ' vpmTimer.AddTimer 30000, "FantasmaTimerZ.Enabled = 1 '" 'move to the end of the police hurry up
    SentinelKicker.Enabled = 0
    
    
End Sub

Sub DisableFlippers(enabled)
    If enabled Then
        SolLFlipper 0
        SolRFlipper 0
        bFlippersEnabled = 0
    Else
        bFlippersEnabled = 1
    End If
End Sub

'********************************
'        Digital clock
'********************************

Dim ClockDigits(4), ClockChars(10)

ClockDigits(0) = Array(a00, a02, a05, a06, a04, a01, a03) 'clock left digit
ClockDigits(1) = Array(a10, a12, a15, a16, a14, a11, a13)
ClockChars(0) = Array(1, 1, 1, 1, 1, 1, 0)                '0
ClockChars(1) = Array(0, 1, 1, 0, 0, 0, 0)                '1
ClockChars(2) = Array(1, 1, 0, 1, 1, 0, 1)                '2
ClockChars(3) = Array(1, 1, 1, 1, 0, 0, 1)                '3
ClockChars(4) = Array(0, 1, 1, 0, 0, 1, 1)                '4
ClockChars(5) = Array(1, 0, 1, 1, 0, 1, 1)                '5
ClockChars(6) = Array(1, 0, 1, 1, 1, 1, 1)                '6
ClockChars(7) = Array(1, 1, 1, 0, 0, 0, 0)                '7
ClockChars(8) = Array(1, 1, 1, 1, 1, 1, 1)                '8
ClockChars(9) = Array(1, 1, 1, 1, 0, 1, 1)                '9

Sub UpdateClock(myTime)
    Dim a, b, i
    a = myTime \ 10
    b = myTime MOD 10
    For i = 0 to 6
        ClockDigits(0)(i).State = ClockChars(a)(i)
        ClockDigits(1)(i).State = ClockChars(b)(i)
    Next
End Sub

Sub TurnOffClock
    Dim i
    For i = 0 to 6
        ClockDigits(0)(i).State = 0
        ClockDigits(1)(i).State = 0
    Next
End Sub

'*********************
 'FANTASMA MOVIMIENTO  /  GHOST MOVEMENT
'*********************
Dim FantasmaDireccion, FantasmaPosicion

FantasmaDireccion = -2
FantasmaPosicion = 351

Sub FantasmaTimer_Timer
	pupevent 825
	pupevent 840
    FantasmaPosicion = FantasmaPosicion + FantasmaDireccion 
    Fantasma.x = FantasmaPosicion
    FlasherFantasma001.X = FantasmaPosicion -15
    FlasherFantasma002.X = FantasmaPosicion +15
    FlasherGrandeSentinelMode.X = FantasmaPosicion
    If FantasmaPosicion <= 265 Then FantasmaDireccion = 2
    If FantasmaPosicion >= 680 Then FantasmaDireccion = -2
    'check for sentinel targets
    DeactivarSentinel 'deactiva todos y activa uno de los targets /  Deactivate all and activate one of the targets
    If FantasmaPosicion > 264 AND FantasmaPosicion < 297 Then Sentarget001.Collidable = 1
    If FantasmaPosicion > 297 AND FantasmaPosicion < 359 Then Sentarget002.Collidable = 1
    If FantasmaPosicion > 359 AND FantasmaPosicion < 421 Then Sentarget003.Collidable = 1
    If FantasmaPosicion > 421 AND FantasmaPosicion < 483 Then Sentarget004.Collidable = 1
    If FantasmaPosicion > 483 AND FantasmaPosicion < 546 Then Sentarget005.Collidable = 1
    If FantasmaPosicion > 546 AND FantasmaPosicion < 608 Then Sentarget006.Collidable = 1
    If FantasmaPosicion > 608 AND FantasmaPosicion < 670 Then Sentarget007.Collidable = 1
    If FantasmaPosicion > 670 AND FantasmaPosicion < 730 Then Sentarget008.Collidable = 1
End Sub

Sub DeactivarSentinel 'targets debajo del sentinel  / Targets below the Sentinel
    For each x in SenTargets
        x.collidable = 0
    Next

	'pupevent 925
End Sub

Sub SenTargets_Hit(idx)
    If bPoliceStarted Then
        PoliceTargetHits = PoliceTargetHits + 1
        PlaySound "sfx_lasergunSEN" 
        Flashforms FlasherFantasma001, 300, 25, 0
        Flashforms FlasherFantasma002, 330, 20, 0
        UpdateClock 10 - PoliceTargetHits
        If PoliceTargetHits = 10 Then
            AwardPoliceJackpot
            StopPolice
        End If
     End If
End Sub


    'FANTASMA MOVIMIENTO eje z
Dim FantasmaDireccionZ, FantasmaPosicionZ
 
'Posicion inicial, debajo del playfield / Starting position, below the playfield
FantasmaDireccionZ = 5
FantasmaPosicionZ = -120

Sub FantasmaTimerZ_Timer
    FantasmaTimer.Enabled = 0 'asegúrate que el movimiento está apagado mientras sube y baja / Make sure the movement is off as it goes up and down
    FantasmaPosicionZ = FantasmaPosicionZ + FantasmaDireccionZ 
    Fantasma.z = FantasmaPosicionZ
    If FantasmaPosicionZ <= -120 Then FantasmaTimerz.Enabled = 0
    If FantasmaPosicionZ >= 65 Then FantasmaTimerz.Enabled = 0:FantasmaTimer.Enabled = 1 'solamente activar el moviemitno horizontal cuando está arriba / Only activate the horizontal motion when it is up
End Sub

'***********************************
' Jason Multiball - Main multiball
'***********************************
' shoot the cabin to start locking
' lock 3 balls under the cabin to start
' all main shots are Lit
' 15 seconds ball saver
' each succesive shot will increase the color and score
' Blue shots: 1 million points
' Green Shots: 2 million points
' Yellow shots: 3 million points
' Orange shots: 4 million points
' Red shots: 5 million points

Sub StartJasonMultiball
    Dim i
    If bMultiBallMode Then Exit Sub 'do not start if already in a multiball mode
    bJasonMBStarted = True
    EnableBallSaver 15
	SelectMultiballEventPairs
	pupevent MultiballStart
	pupevent 838
    DMD "_", CL("TRINITY MULTBALL"), "_", eNone, eNone, eNone, 2500, True, "vo_multiball1"
    AddMultiball 2
    For i = 1 to 8
        ArrowMultiPlier(i) = 1
    Next
    UpdateArrowLights
    li016.State = 2
    ChangeSong
End Sub

Sub StopJasonMultiball
	pupevent MultiballStop
	pupevent 864
    bJasonMBStarted = False
    BallsInLock(CurrentPlayer) = 0
    TurnOffArrows
    li016.State = 0
End Sub

Sub UpdateArrowLights 'sets the color of the arrows according to the jackpot multiplier
    SetLightColor li062, ArrowMultiPlier(1), 2
    SetLightColor li064, ArrowMultiPlier(2), 2
    SetLightColor li053, ArrowMultiPlier(3), 2
    SetLightColor li066, ArrowMultiPlier(4), 2
    SetLightColor li067, ArrowMultiPlier(5), 2
    SetLightColor li068, ArrowMultiPlier(6), 2
    SetLightColor li052, ArrowMultiPlier(7), 2
    SetLightColor li059, ArrowMultiPlier(8), 2
End Sub

'***********************************
' Freddy Kruger Multiball - Ramps
'***********************************
' starts after 3 counselors are killed
' shoot the blue arrows at the ramps to collect jackpots
' this will build the super jackpot at the right loop, shoot the cabin for a better shot

'===============================   DEBUG CODE ===========================================
Dim CurrTime,objIEDebugWindow
CurrTime = Timer
Sub Debug( myDebugText )

' Uncomment the next line to turn off debugging
Exit Sub


	If Not IsObject( objIEDebugWindow ) Then
		Set objIEDebugWindow = CreateObject( "InternetExplorer.Application" )

		objIEDebugWindow.Navigate "about:blank"
		objIEDebugWindow.Visible = True
		'objIEDebugWindow.AutoScroll = True
		objIEDebugWindow.ToolBar = False
		objIEDebugWindow.Width = 1400	
		objIEDebugWindow.Height = 500
		objIEDebugWindow.Left = 2100
		objIEDebugWindow.Top = 520
		Do While objIEDebugWindow.Busy
		Loop
		objIEDebugWindow.Document.Title = "My Debug Window"
		objIEDebugWindow.Document.Body.InnerHTML = "<b>Matrix Debug Window -TimeStamp: " & GameTime & "</b></br>"

	End If

objIEDebugWindow.Document.Body.InnerHTML = objIEDebugWindow.Document.Body.InnerHTML & myDebugText & " --TimeStamp:<b> " & GameTime & "</b><br>" & vbCrLf
objIEDebugWindow.Document.Body.scrollTop = objIEDebugWindow.Document.Body.scrollTop +objIEDebugWindow.Document.Body.scrollHeight
End Sub
'=========================================================================================
Sub Debug4( myDebugText )

' Uncomment the next line to turn off debugging
Exit Sub


	If Not IsObject( objIEDebugWindow ) Then
		Set objIEDebugWindow = CreateObject( "InternetExplorer.Application" )

		objIEDebugWindow.Navigate "about:blank"
		objIEDebugWindow.Visible = True
		'objIEDebugWindow.AutoScroll = True
		objIEDebugWindow.ToolBar = False
		objIEDebugWindow.Width = 1400	
		objIEDebugWindow.Height = 500
		objIEDebugWindow.Left = 2100
		objIEDebugWindow.Top = 520
		Do While objIEDebugWindow.Busy
		Loop
		objIEDebugWindow.Document.Title = "My Debug Window"
		objIEDebugWindow.Document.Body.InnerHTML = "<b>Matrix Debug Window -TimeStamp: " & CurrTime & "</b></br>"

	End If

objIEDebugWindow.Document.Body.InnerHTML = objIEDebugWindow.Document.Body.InnerHTML & myDebugText & " --TimeStamp:<b> " & CurrTime & "</b><br>" & vbCrLf
objIEDebugWindow.Document.Body.scrollTop = objIEDebugWindow.Document.Body.scrollTop +objIEDebugWindow.Document.Body.scrollHeight
End Sub

dim MultiballStart,MultiballStop,tmp

Sub SelectMultiballEventPairs
	tmp = RndNbr(2)
	Select Case tmp
		Case 1:
			Debug "Case1: " &tmp &":" &MultiballStart
			MultiballStart = "860"
			MultiballStop = "861"
		Case 2:
			Debug "Case2: " &tmp &":" &MultiballStart
			MultiballStart = "862"
			MultiballStop = "863"
	End Select

End Sub

Sub StartFreddyMultiball
    Dim i
    If bMultiBallMode Then Exit Sub 'do not start if already in a multiball mode
    bFreddyMBStarted = True
    ChangeSong
    EnableBallSaver 20
	SelectMultiballEventPairs
	pupevent MultiballStart
	pupevent 838
    DMD "_", CL("TRINITY MULTIBALL"), "_", eNone, eNone, eNone, 2500, True, "vo_multiball1"
    AddMultiball 2
    SetLightColor li064, blue, 2
    SetLightColor li068, blue, 2
    li015.State = 2
End Sub

Sub StopFreddyMultiball
	pupevent MultiballStop	
	pupevent 864
    bFreddyMBStarted = False
    li064.State = 0
    li068.State = 0
    li015.State = 0
End Sub

'***********************************
' Michael Myers Multiball - Spinners
'***********************************
' starts randomly after 3 counselors are killed
' shoot the blue arrows at the spinners to collect jackpots
' this will build the super jackpot at the lower right loop
' shoot it to collect it

Sub StartMichaelMultiball
    Dim i
    If bMultiBallMode Then Exit Sub 'do not start if already in a multiball mode
    bMichaelMBStarted = True
    ChangeSong
    EnableBallSaver 20
	SelectMultiballEventPairs
	pupevent MultiballStart
	pupevent 838
    DMD "_", CL("MORPHEUS MULTIBALL"), "_", eNone, eNone, eNone, 2500, True, "vo_multiball1"
    AddMultiball 2
    SetLightColor li062, blue, 2
    SetLightColor li067, blue, 2
    SetLightColor li059, blue, 2
    li075.State = 2
End Sub

Sub StopMichaelMultiball
	pupevent MultiballStop	
	pupevent 864
    bMichaelMBStarted = False
    li062.State = 0
    li067.State = 0
    li059.State = 0
    li075.State = 0
End Sub

'****************************
' Mystery award at the scoop
'****************************
' this is a kind of award after completing the inlane and outlanes

Sub CheckMystery 'if all the inlanes and outlanes are lit then lit the mystery award
    dim i
    If Mystery(CurrentPlayer, 1) + Mystery(CurrentPlayer, 2) + Mystery(CurrentPlayer, 3) + Mystery(CurrentPlayer, 4) = 4 Then
        DMD "_", CL("MYSTERY IS LIT"), "", eNone, eNone, eNone, 1000, True, "vo_waytogo"
        li078.State = 1
        ' and reset the lights
        For i = 1 to 4
            Mystery(CurrentPlayer, i) = 0
        Next
    End If
    UpdateMysteryLights
End Sub

Sub AwardMystery 'mostly points but sometimes it will lit the special or the extra ball
    Dim tmp
    FlashEffect 1
    Select Case RndNbr(20)
        Case 1:LitExtraBall                   'lit extra ball
        Case 2:LitSpecial                     'lit special
        Case Else:
            tmp = 250000 + RndNbr(25) * 10000 'add from 250.000 to 500.000
            Select Case RndNbr(4)
                Case 1:DMD CL("MYSTERY SCORE"), CL(FormatScore(tmp) ), "", eNone, eNone, eNone, 2000, True, "vo_notbad"
                Case 2:DMD CL("MYSTERY SCORE"), CL(FormatScore(tmp) ), "", eNone, eNone, eNone, 2000, True, "vo_excellentscore"
                Case 3:DMD CL("MYSTERY SCORE"), CL(FormatScore(tmp) ), "", eNone, eNone, eNone, 2000, True, "vo_nowyouaregettinghot"
                Case 4:DMD CL("MYSTERY SCORE"), CL(FormatScore(tmp) ), "", eNone, eNone, eNone, 2000, True, "vo_welldone"
            End Select
    End Select
    vpmtimer.addtimer 3500, "kickBallOut '"
End Sub

'**********************************
'   Modes - Hunting the counselors
'**********************************

' This table has 14 modes which will be selected at random
' After killing all counselors you'll a surprise :)

' current active Mode number is stored in Mode(CurrentPlayer,0)
' select a new random Mode if none is active
Sub SelectMode
    Dim i
    If Mode(CurrentPlayer, 0) = 0 Then
        ' reset the Modes that are not finished
        For i = 1 to 14
            If Mode(CurrentPlayer, i) = 2 Then Mode(CurrentPlayer, i) = 0
        Next
        NewMode = RndNbr(14)
        do while Mode(CurrentPlayer, NewMode) <> 0
            NewMode = RndNbr(14)
        loop
        Mode(CurrentPlayer, NewMode) = 2
        Li076.State = 2 'Start hunting light at the scoop
        UpdateModeLights
    'debug.print "newmode " & newmode
    End If
End Sub

' Update the lights according to the mode's state
Sub UpdateModeLights
    li014.State = Mode(CurrentPlayer, 1)
    li013.State = Mode(CurrentPlayer, 2)
    li012.State = Mode(CurrentPlayer, 3)
    li011.State = Mode(CurrentPlayer, 4)
    li010.State = Mode(CurrentPlayer, 5)
    li009.State = Mode(CurrentPlayer, 6)
    li008.State = Mode(CurrentPlayer, 7)
    li007.State = Mode(CurrentPlayer, 8)
    li006.State = Mode(CurrentPlayer, 9)
    li005.State = Mode(CurrentPlayer, 10)
    li004.State = Mode(CurrentPlayer, 11)
    li001.State = Mode(CurrentPlayer, 12)
    li002.State = Mode(CurrentPlayer, 13)
    li003.State = Mode(CurrentPlayer, 14)
End Sub

' Starting a mode means to setup some lights and variables, maybe timers
' Mode lights will always blink during an active mode
Sub StartMode
    Mode(CurrentPlayer, 0) = NewMode
    Li076.State = 0
    ChangeSong
    'PlaySound "vo_hahaha"&RndNbr(4)
    ReadyToKill = False
    Select Case NewMode
        Case 1 'A.J Mason = Super Spinners
			pupevent 814
			pupevent 849
            DMD CL("FOLLOW WHITE RABBIT"), CL("SHOOT THE SPINNERS"), "", eNone, eNone, eNone, 1500, True, "vo_shootthespinners"
            SpinCount = 0
            SetLightColor li062, amber, 2
            SetLightColor li067, amber, 2
            SetLightColor li059, amber, 2
            SpinNeeded = 50
        Case 2 'Adam = 5 Targets at semi random
			pupevent 819
			pupevent 853
            DMD CL("AGENT SMITH FIGHT"), CL("HIT THE LIT TARGETS"), "", eNone, eNone, eNone, 1500, True, "vo_shootthetargets"
            TargetModeHits = 0
            'Turn off all blue targets
            Li031.State = 0
            Li049.State = 0
            Li050.State = 0
            Li051.State = 0
            Li058.State = 0
            Li057.State = 0
            Li079.State = 0
            Select Case RndNbr(4) 'select 4 blue targets
                Case 1:Li031.State = 2:Li050.State = 2:li058.State = 2:Li079.State = 2
                Case 2:Li049.State = 2:Li050.State = 2:li058.State = 2:Li057.State = 2
                Case 3:Li031.State = 2:Li049.State = 2:li051.State = 2:Li057.State = 2
                Case 4:Li031.State = 2:Li049.State = 2:li050.State = 2:Li051.State = 2
            End Select
        Case 3 'Brandon = 5 Flashing Shots 90 seconds to complete
			pupevent 818
			pupevent 852
            DMD CL("ORACLE MEETING"), CL("SHOOT THE LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            DMD CL("ORACLE MEETING"), CL("YOU HAVE 90 SECONDS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            SetLightColor li062, amber, 2
            SetLightColor li064, amber, 2
            SetLightColor li066, amber, 2
            SetLightColor li067, amber, 2
            SetLightColor li068, amber, 2
            EndModeCountdown = 90
            EndModeTimer.Enabled = 1
        Case 4 'Chad = 5 Orbits 'uses trigger008 to detect the completed orbits
			pupevent 806
			pupevent 841
            DMD CL("KILL SMITH CLONES"), CL("SHOOT 5 ORBITS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            SetLightColor li067, amber, 2
            SetLightColor li052, amber, 2
        Case 5 'Deborah = Shoot 4 lights 60 seconds, all shots are lit, after 4 shot lit the cabin to kill her & collect jackpot
			pupevent 807
			pupevent 842
            DMD CL("SERAPH FIGHT "), CL("SHOOT 4 LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            DMD CL("SERAPH FIGHT"), CL("YOU HAVE 60 SECONDS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            TurnonArrows amber
            EndModeCountdown = 60
            EndModeTimer.Enabled = 1
        Case 6 'Eric = Shoot the ramps
			pupevent 811
			pupevent 846
            DMD CL("FIND KEYMAKER"), CL("SHOOT 5 RAMPS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            SetLightColor li064, amber, 2
            SetLightColor li068, amber, 2
        Case 7 'Jenny=  Target Frenzy
			pupevent 809
			pupevent 844
            DMD CL("MEROVINGIAN"), CL("SHOOT 4 LIT TARGETS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            'Turn off all blue targets
            TurnOffBlueTargets
            Select Case RndNbr(4) 'select 4 blue targets
                Case 1:Li031.State = 2:Li050.State = 2:li058.State = 2:Li079.State = 2
                Case 2:Li049.State = 2:Li050.State = 2:li058.State = 2:Li057.State = 2
                Case 3:Li031.State = 2:Li049.State = 2:li051.State = 2:Li057.State = 2
                Case 4:Li031.State = 2:Li049.State = 2:li050.State = 2:Li051.State = 2
            End Select
        Case 8 'Mitch = 5 Targets in rotation
			pupevent 808
			pupevent 843
            DMD CL("THE TWINS CHASE"), CL("SHOOT 5 LIT TARGETS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            'Turn off all blue targets
            TurnOffBlueTargets
            'Start timer to rotate through the targets
            BlueTargetsCount = RndNbr(7)
            BlueTargetsTimer_Timer
            BlueTargetsTimer.Enabled = 1
        Case 9 'Fox = Magnet
			pupevent 817
			pupevent 851
            DMD CL("MEET THE ARCHITECT"), CL("SHOOT THE MAGNET"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            li065.State = 2
        Case 10 'Victoria = Ramps and Orbits - all lights lit
			pupevent 822
			pupevent 855
            DMD CL("THE TRAINMAN"), CL("SHOOT RAMPS ORBITS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            SetLightColor li064, amber, 2
            SetLightColor li067, amber, 2
            SetLightColor li068, amber, 2
            SetLightColor li052, amber, 2
        Case 11 'Kenny = 5 Blue Targets at random
			pupevent 824
			pupevent 856
            DMD CL("DEUS EX MACHINA"), CL("SHOOT BLUE TARGETS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            'Turn off all blue targets
            TurnOffBlueTargets
            'Start timer to rotate through the targets
            BlueTargetsCount = RndNbr(7)
            BlueTargetsTimer_Timer
            BlueTargetsTimer.Enabled = 1
        Case 12 'Sheldon = Super Spinners at random
			pupevent 810
			pupevent 845
            DMD CL("KILL CYPHER"), CL("SHOOT THE SPINNERS"), "", eNone, eNone, eNone, 1500, True, ""
            SpinCount = 0
            SpinNeeded = 50
            SpinnersTimer_Timer
            SpinnersTimer.Enabled = 1
        Case 13 'Tiffany = Follow the Lights All main Shots 1 at a time in rotation
			pupevent 812
			pupevent 847
			DMD CL("ZION SPEECH"), CL("SHOOT LIT LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            ArrowsCount = 0
            SetLightColor li062, amber, 0
            SetLightColor li064, amber, 0
            SetLightColor li053, amber, 0
            SetLightColor li066, amber, 0
            SetLightColor li067, amber, 0
            SetLightColor li068, amber, 0
            SetLightColor li052, amber, 0
            SetLightColor li059, amber, 0
            FollowTheLights_Timer
            FollowTheLights.Enabled = 1
        Case 14 'Vanessa = Follow the Lights All main shots at random
			pupevent 813
			pupevent 848
            DMD CL("PERSEPHONES REVENGE"), CL("SHOOT LIT LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            TargetModeHits = 0
            ArrowsCount = 0
            SetLightColor li062, amber, 0
            SetLightColor li064, amber, 0
            SetLightColor li053, amber, 0
            SetLightColor li066, amber, 0
            SetLightColor li067, amber, 0
            SetLightColor li068, amber, 0
            SetLightColor li052, amber, 0
            SetLightColor li059, amber, 0
            FollowTheLights_Timer
            FollowTheLights.Enabled = 1
        Case 15 'the big final mode: 5 ball multiball all jackpots are lit, no timer, score jackpots until the last multiball
			pupevent 820
			pupevent 854
            DMD CL("DESTROY MATRIX"), CL("SHOOT THE JACKPOTS"), "", eNone, eNone, eNone, 1500, True, ""
            AddMultiball 5
            SetLightColor li062, red, 2
            SetLightColor li064, red, 2
            SetLightColor li053, red, 2
            SetLightColor li066, red, 2
            SetLightColor li067, red, 2
            SetLightColor li068, red, 2
            SetLightColor li052, red, 2
            SetLightColor li059, red, 2
            ChangeGi Red
            ChangeGIIntensity 2
    End Select
    ' kick out the ball
    If BallsinHole Then
        vpmtimer.addtimer 2500, "kickBallOut '"
    End If
End Sub

Sub CheckWinMode
    DOF 126, DOFPulse
    LightSeqInserts.StopPlay 'stop the light effects before starting again so they don't play too long.
    LightEffect 3
    Select Case NewMode
        Case 1
            Addscore 10000
            If SpinCount = SpinNeeded Then
                DMD CL("YOU CATCHED NOW"), CL("SHOOT THE SCOOP"), "", eNone, eNone, eNone, 1500, True, "vo_itstimetodie"
                li029.State = 2 'lit the Hunter Jackpot at the scoop
                li062.State = 0 'turn off the spinner lights
                li067.State = 0
                li059.State = 0
                ReadyToKill = True
            End If
        Case 2
            Addscore 150000
            If TargetModeHits = 4 Then
                DMD CL("YOU CATCHED"), CL("SHOOT THE MAGNET"), "", eNone, eNone, eNone, 1500, True, "vo_itstimetodie"
                li065.State = 2
                ReadyToKill = True
            End If
        Case 3
            Addscore 150000
            If TargetModeHits = 5 Then
                DMD CL("YOU CATCHED"), CL("SHOOT THE SCOOP"), "", eNone, eNone, eNone, 1500, True, "vo_itstimetodie"
                li029.State = 2
                ReadyToKill = True
            End If
        Case 4
            Addscore 150000
            If TargetModeHits = 5 Then WinMode:End if
        Case 5
            Addscore 150000
            If TargetModeHits = 4 Then      'lit the cabin for the kill & jackpot
                DMD CL("YOU CATCHED NOW"), CL("SHOOT THE PHONE"), "", eNone, eNone, eNone, 1500, True, "vo_itstimetodie"
                SetLightColor li066, red, 2 'change the color to red
                ReadyToKill = True
            End If
        Case 6
            Addscore 150000
            If TargetModeHits = 5 Then WinMode:End if
        Case 7
            Addscore 150000
            If TargetModeHits = 4 Then WinMode:End if
        Case 8
            Addscore 150000
            If TargetModeHits = 5 Then
                WinMode
            Else
                BlueTargetsTimer_Timer 'to lit another target
            End if
        Case 9
            Addscore 150000
            If TargetModeHits = 5 Then WinMode:End if
        Case 10
            Addscore 150000
            If TargetModeHits = 6 Then WinMode:End if
        Case 11
            Addscore 150000
            If TargetModeHits = 6 Then
                WinMode
            Else
                BlueTargetsTimer_Timer 'to lit another target
            End if
        Case 12
            Addscore 10000
            If SpinCount = SpinNeeded Then
                WinMode
                SpinnersTimer.Enabled = 0
            End If
        Case 13
            Addscore 150000
            If TargetModeHits = 6 Then
                WinMode
            Else
                FollowtheLights_Timer 'to lit another target
            End if
        Case 14
            Addscore 150000
            If TargetModeHits = 6 Then
                WinMode
            Else
                FollowtheLights_Timer 'to lit another target
            End if
    End Select
End Sub

'called after completing a mode
Sub WinMode
    CounselorsKilled(CurrentPlayer) = CounselorsKilled(CurrentPlayer) + 1
    Select Case NewMode
        Case 1, 5, 7, 10, 13, 14
            DMD "", "", "d_kill1", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill2", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill3", eNone, eNone, eNone, 240, False, "sfx_screamf" &RndNbr(5)
            DMD "", "", "d_kill4", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill5", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill6", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill7", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill8", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill9", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill10", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill11", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill12", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill13", eNone, eNone, eNone, 240, False, ""
            DMD CL("NICE MR.ANDERSON"), CL(FormatScore(3500000) ), "_", eNone, eBlink, eNone, 1200, False, ""
            DMD CL("NICE MR.ANDERSON"), CL(FormatScore(3500000) ), "_", eNone, eBlink, eNone, 2000, True, "vo_fantastic"
			pupevent 868
			pupevent 857
            Addscore2 3500000
        Case 2, 3, 4, 6, 8, 9, 11, 12
            DMD "", "", "d_kill1", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill2", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill3", eNone, eNone, eNone, 240, False, "sfx_screamm" &RndNbr(5)
            DMD "", "", "d_kill4", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill5", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill6", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill7", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill8", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill9", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill10", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill11", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill12", eNone, eNone, eNone, 240, False, ""
            DMD "", "", "d_kill13", eNone, eNone, eNone, 240, False, ""
            DMD CL("KEEP FIGHTING NEO"), CL(FormatScore(3500000) ), "_", eNone, eBlinkFast, eNone, 1200, False, ""
            DMD CL("KEEP FIGHTING NEO"), CL(FormatScore(3500000) ), "_", eNone, eBlinkFast, eNone, 2000, True, "vo_excellent"
			pupevent 868
			pupevent 857
            Addscore2 3500000
    End Select

' Update Scorbit
	Select Case NewMode
		Case 1: ScorbitBuildGameModes
		Case 2: ScorbitBuildGameModes
		Case 3: ScorbitBuildGameModes
		Case 4: ScorbitBuildGameModes
		Case 5: ScorbitBuildGameModes
		Case 6: ScorbitBuildGameModes
		Case 7: ScorbitBuildGameModes
		Case 8: ScorbitBuildGameModes
		Case 9: ScorbitBuildGameModes
		Case 10: ScorbitBuildGameModes
		Case 11: ScorbitBuildGameModes
		Case 12: ScorbitBuildGameModes
		Case 13: ScorbitBuildGameModes
		Case 14: ScorbitBuildGameModes
		Case 15: ScorbitBuildGameModes
	End Select
    ' eject the ball if it is in the scoop
    If BallsinHole Then
        vpmtimer.addtimer 5000, "kickBallOut '"
    End If
    DOF 139, DOFPulse
    Mode(CurrentPlayer, 0) = 0
    Mode(CurrentPlayer, NewMode) = 1
    UpdateModeLights
    FlashEffect 2
    LightEffect 2
    GiEffect 2
    ChangeSong
    StopMode2 'to stop specific lights, timers and variables of the mode
    ' start the police hurry up after 2 kills

    ' check for extra modes after each 3 completed kills
    IF CounselorsKilled(CurrentPlayer) MOD 6 = 0 Then
        StartMichaelMultiball
    ElseIF CounselorsKilled(CurrentPlayer) MOD 3 = 0 Then
        StartFreddyMultiball
    ' If all counselors er dead then start final mode
    ElseIF CounselorsKilled(CurrentPlayer) = 14 Then
        NewMode = 15
        StartMode
		ScorbitBuildGameModes
    End If
End Sub

Sub StopMode 'called at the end of a ball
    Dim i
    Mode(CurrentPlayer, 0) = 0
    For i = 0 to 14 'stop any counselor blinking light
        If Mode(CurrentPlayer, i) = 2 Then Mode(CurrentPlayer, i) = 0
    Next
    UpdateModeLights
    StopMode2
End Sub

Sub StopMode2 'stop any mode special lights, timers or variables, called after a win or end of ball
    ' stop some timers or reset Mode variables
    Select Case NewMode
        Case 1:TurnOffArrows:li029.State = 0
        Case 2:UpdateTargetLights:li065.State = 0
        Case 3:li029.State = 0:EndModeTimer.Enabled = 0
        Case 4:li067.State = 0:li052.State = 0
        Case 5:TurnOffArrows:EndModeTimer.Enabled = 0
        Case 6:TurnOffArrows
        Case 7:UpdateTargetLights
        Case 8:BlueTargetsTimer.Enabled = 0:UpdateTargetLights
        Case 9:li065.State = 0
        Case 10:TurnOffArrows
        Case 11:BlueTargetsTimer.Enabled = 0:UpdateTargetLights
        Case 12:TurnOffArrows:SpinnersTimer.Enabled = 0
        Case 13:TurnOffArrows:FollowTheLights.Enabled = 0
        Case 14:TurnOffArrows:FollowTheLights.Enabled = 0
        Case 15:ResetModes
    End Select
    ' restore multiball lights
    If bJasonMBStarted Then UpdateArrowLights:End If
    If bFreddyMBStarted Then SetLightColor li064, blue, 2:SetLightColor li068, blue, 2
    If bMichaelMBStarted Then SetLightColor li062, blue, 2:SetLightColor li067, blue, 2:SetLightColor li059, blue, 2
    NewMode = 0
    ReadyToKill = False
End Sub

Sub ResetModes
    Dim i, j
    For j = 0 to 4
        CounselorsKilled(j) = 0
        For i = 0 to 14
            Mode(CurrentPlayer, i) = 0
        Next
    Next
    NewMode = 0
    'reset Mode variables
    TurnOffArrows
    SpinCount = 0
    ReadyToKill = False
End Sub

Sub EndModeTimer_Timer '1 second timer to count down to end the timed modes
    EndModeCountdown = EndModeCountdown - 1
    Select Case EndModeCountdown
        Case 16:PlaySound "vo_timerunningout"
        Case 10:PlaySound "vo_ten"
        Case 9:PlaySound "vo_nine"
        Case 8:PlaySound "vo_eight"
        Case 7:PlaySound "vo_seven"
        Case 6:PlaySound "vo_six"
        Case 5:PlaySound "vo_five"
        Case 4:PlaySound "vo_four"
        Case 3:PlaySound "vo_three"
        Case 2:PlaySound "vo_two"
        Case 1:PlaySound "vo_one"
        Case 0:PlaySound "vo_timeisup":StopMode
    End Select
End Sub

Sub BlueTargetsTimer_Timer 'rotates though all the targets
    If NewMode = 8 Then
        BlueTargetsCount = (BlueTargetsCount + 1) MOD 7
    Else                                 'this will be mode 11
        BlueTargetsCount = RndNbr(7) - 1 'from 0 to 6
    End If
    Select Case BlueTargetsCount
        Case 0:Li031.State = 2:Li049.State = 0:Li050.State = 0:Li051.State = 0:Li058.State = 0:Li057.State = 0:Li079.State = 0
        Case 1:Li031.State = 0:Li049.State = 2:Li050.State = 0:Li051.State = 0:Li058.State = 0:Li057.State = 0:Li079.State = 0
        Case 2:Li031.State = 0:Li049.State = 0:Li050.State = 2:Li051.State = 0:Li058.State = 0:Li057.State = 0:Li079.State = 0
        Case 3:Li031.State = 0:Li049.State = 0:Li050.State = 0:Li051.State = 2:Li058.State = 0:Li057.State = 0:Li079.State = 0
        Case 4:Li031.State = 0:Li049.State = 0:Li050.State = 0:Li051.State = 0:Li058.State = 2:Li057.State = 0:Li079.State = 0
        Case 5:Li031.State = 0:Li049.State = 0:Li050.State = 0:Li051.State = 0:Li058.State = 0:Li057.State = 2:Li079.State = 0
        Case 6:Li031.State = 0:Li049.State = 0:Li050.State = 0:Li051.State = 0:Li058.State = 0:Li057.State = 0:Li079.State = 2
    End Select
End Sub

Sub SpinnersTimer_Timer
    Select Case RndNbr(3)
        Case 0:SetLightColor li062, amber, 2:SetLightColor li067, amber, 0:SetLightColor li059, amber, 0
        Case 1:SetLightColor li062, amber, 0:SetLightColor li067, amber, 2:SetLightColor li059, amber, 0
        Case 2:SetLightColor li062, amber, 0:SetLightColor li067, amber, 0:SetLightColor li059, amber, 2
    End Select
End Sub

Sub FollowTheLights_Timer 'rotates though all the targets
    If NewMode = 13 Then
        ArrowsCount = (ArrowsCount + 1) MOD 8
    Else                            'this will be mode 14
        ArrowsCount = RndNbr(8) - 1 'from 0 to 8
    End If
    Select Case ArrowsCount
        Case 0:li062.State = 2:Li064.State = 0:Li053.State = 0:Li066.State = 0:Li067.State = 0:Li068.State = 0:Li052.State = 0:Li059.State = 0
        Case 1:li062.State = 0:Li064.State = 2:Li053.State = 0:Li066.State = 0:Li067.State = 0:Li068.State = 0:Li052.State = 0:Li059.State = 0
        Case 2:li062.State = 0:Li064.State = 0:Li053.State = 2:Li066.State = 0:Li067.State = 0:Li068.State = 0:Li052.State = 0:Li059.State = 0
        Case 3:li062.State = 0:Li064.State = 0:Li053.State = 0:Li066.State = 2:Li067.State = 0:Li068.State = 0:Li052.State = 0:Li059.State = 0
        Case 4:li062.State = 0:Li064.State = 0:Li053.State = 0:Li066.State = 0:Li067.State = 2:Li068.State = 0:Li052.State = 0:Li059.State = 0
        Case 5:li062.State = 0:Li064.State = 0:Li053.State = 0:Li066.State = 0:Li067.State = 0:Li068.State = 2:Li052.State = 0:Li059.State = 0
        Case 6:li062.State = 0:Li064.State = 0:Li053.State = 0:Li066.State = 0:Li067.State = 0:Li068.State = 0:Li052.State = 2:Li059.State = 0
        Case 7:li062.State = 0:Li064.State = 0:Li053.State = 0:Li066.State = 0:Li067.State = 0:Li068.State = 0:Li052.State = 0:Li059.State = 2
    End Select
End Sub

'***********
'  torch
'***********
Dim i1, i2
i1 = 0
i2 = 4

Sub torchtimer_timer

    i1 = (i1 + 1) MOD 8
    i2 = (i2 + 1) MOD 8
End Sub


'******************************************************
'	ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners 				https://youtu.be/AXX3aen06FM
' Adding nFozzy roth physics : pt2 flipper physics 					https://youtu.be/VSBFuK2RCPE
' Adding nFozzy roth physics : pt3 other elements 					https://youtu.be/JN8HEJapCvs
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |






'******************************************************
'	ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
'	1. flippers with specific physics settings
'	2. custom triggers for each flipper (TriggerLF, TriggerRF)
'	3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
'	4. and, special scripting
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
'	for each x in a
'		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 80
'		x.DebugOn=False ' prints some info in debugger
'
'		x.AddPt "Polarity", 0, 0, 0
'		x.AddPt "Polarity", 1, 0.05, - 2.7
'		x.AddPt "Polarity", 2, 0.33, - 2.7
'		x.AddPt "Polarity", 3, 0.37, - 2.7
'		x.AddPt "Polarity", 4, 0.41, - 2.7
'		x.AddPt "Polarity", 5, 0.45, - 2.7
'		x.AddPt "Polarity", 6, 0.576, - 2.7
'		x.AddPt "Polarity", 7, 0.66, - 1.8
'		x.AddPt "Polarity", 8, 0.743, - 0.5
'		x.AddPt "Polarity", 9, 0.81, - 0.5
'		x.AddPt "Polarity", 10, 0.88, 0
'
'		x.AddPt "Velocity", 0, 0, 1
'		x.AddPt "Velocity", 1, 0.16, 1.06
'		x.AddPt "Velocity", 2, 0.41, 1.05
'		x.AddPt "Velocity", 3, 0.53, 1 '0.982
'		x.AddPt "Velocity", 4, 0.702, 0.968
'		x.AddPt "Velocity", 5, 0.95,  0.968
'		x.AddPt "Velocity", 6, 1.03, 0.945
'	Next
'
'	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
'	for each x in a
'		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 80
'		x.DebugOn=False ' prints some info in debugger
'
'		x.AddPt "Polarity", 0, 0, 0
'		x.AddPt "Polarity", 1, 0.05, - 3.7
'		x.AddPt "Polarity", 2, 0.33, - 3.7
'		x.AddPt "Polarity", 3, 0.37, - 3.7
'		x.AddPt "Polarity", 4, 0.41, - 3.7
'		x.AddPt "Polarity", 5, 0.45, - 3.7
'		x.AddPt "Polarity", 6, 0.576,- 3.7
'		x.AddPt "Polarity", 7, 0.66, - 2.3
'		x.AddPt "Polarity", 8, 0.743, - 1.5
'		x.AddPt "Polarity", 9, 0.81, - 1
'		x.AddPt "Polarity", 10, 0.88, 0
'
'		x.AddPt "Velocity", 0, 0, 1
'		x.AddPt "Velocity", 1, 0.16, 1.06
'		x.AddPt "Velocity", 2, 0.41, 1.05
'		x.AddPt "Velocity", 3, 0.53, 1 '0.982
'		x.AddPt "Velocity", 4, 0.702, 0.968
'		x.AddPt "Velocity", 5, 0.95,  0.968
'		x.AddPt "Velocity", 6, 1.03, 0.945
'
'	Next
'
'	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
''*******************************************
''  Late 80's early 90's
'
'Sub InitPolarity()
'	dim x, a : a = Array(LF, RF)
'	for each x in a
'		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 60
'		x.DebugOn=False ' prints some info in debugger
'
'		x.AddPt "Polarity", 0, 0, 0
'		x.AddPt "Polarity", 1, 0.05, - 5
'		x.AddPt "Polarity", 2, 0.4, - 5
'		x.AddPt "Polarity", 3, 0.6, - 4.5
'		x.AddPt "Polarity", 4, 0.65, - 4.0
'		x.AddPt "Polarity", 5, 0.7, - 3.5
'		x.AddPt "Polarity", 6, 0.75, - 3.0
'		x.AddPt "Polarity", 7, 0.8, - 2.5
'		x.AddPt "Polarity", 8, 0.85, - 2.0
'		x.AddPt "Polarity", 9, 0.9, - 1.5
'		x.AddPt "Polarity", 10, 0.95, - 1.0
'		x.AddPt "Polarity", 11, 1, - 0.5
'		x.AddPt "Polarity", 12, 1.1, 0
'		x.AddPt "Polarity", 13, 1.3, 0
'
'		x.AddPt "Velocity", 0, 0, 1
'		x.AddPt "Velocity", 1, 0.16, 1.06
'		x.AddPt "Velocity", 2, 0.41, 1.05
'		x.AddPt "Velocity", 3, 0.53, 1 '0.982
'		x.AddPt "Velocity", 4, 0.702, 0.968
'		x.AddPt "Velocity", 5, 0.95,  0.968
'		x.AddPt "Velocity", 6, 1.03,  0.945
'	Next
'
'	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'	LF.SetObjects "LF", LeftFlipper, TriggerLF
'	RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'*******************************************
' Early 90's and after

Sub InitPolarity()
	Dim x, a
	a = Array(LF, RF)
	For Each x In a
		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 60
		x.DebugOn=False ' prints some info in debugger
		
		x.AddPt "Polarity", 0, 0, 0
		x.AddPt "Polarity", 1, 0.05, -5.5
		x.AddPt "Polarity", 2, 0.4, -5.5
		x.AddPt "Polarity", 3, 0.6, -5.0
		x.AddPt "Polarity", 4, 0.65, -4.5
		x.AddPt "Polarity", 5, 0.7, -4.0
		x.AddPt "Polarity", 6, 0.75, -3.5
		x.AddPt "Polarity", 7, 0.8, -3.0
		x.AddPt "Polarity", 8, 0.85, -2.5
		x.AddPt "Polarity", 9, 0.9,-2.0
		x.AddPt "Polarity", 10, 0.95, -1.5
		x.AddPt "Polarity", 11, 1, -1.0
		x.AddPt "Polarity", 12, 1.05, -0.5
		x.AddPt "Polarity", 13, 1.1, 0
		x.AddPt "Polarity", 14, 1.3, 0
		
		x.AddPt "Velocity", 0, 0,	   1
		x.AddPt "Velocity", 1, 0.160, 1.06
		x.AddPt "Velocity", 2, 0.410, 1.05
		x.AddPt "Velocity", 3, 0.530, 1'0.982
		x.AddPt "Velocity", 4, 0.702, 0.968
		x.AddPt "Velocity", 5, 0.95,  0.968
		x.AddPt "Velocity", 6, 1.03,  0.945
	Next
	
	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
	LF.SetObjects "LF", LeftFlipper, TriggerLF
	RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'' Flipper trigger hit subs
'Sub TriggerLF_Hit()
'	LF.Addball activeball
'End Sub
'Sub TriggerLF_UnHit()
'	LF.PolarityCorrect activeball
'End Sub
'Sub TriggerRF_Hit()
'	RF.Addball activeball
'End Sub
'Sub TriggerRF_UnHit()
'	RF.PolarityCorrect activeball
'End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt		'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay		'delay before trigger turns off and polarity is disabled
	Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)
	Private Name
	
	Dim PolarityIn, PolarityOut
	Dim VelocityIn, VelocityOut
	Dim YcoefIn, YcoefOut
	Public Sub Class_Initialize
		ReDim PolarityIn(0)
		ReDim PolarityOut(0)
		ReDim VelocityIn(0)
		ReDim VelocityOut(0)
		ReDim YcoefIn(0)
		ReDim YcoefOut(0)
		Enabled = True
		TimeDelay = 50
		LR = 1
		Dim x
		For x = 0 To UBound(balls)
			balls(x) = Empty
			Set Balldata(x) = new SpoofBall
		Next
	End Sub
	
	Public Sub SetObjects(aName, aFlipper, aTrigger)
		
		If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
		If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
		If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
		If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
		Name = aName
		Set Flipper = aFlipper
		FlipperStart = aFlipper.x
		FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
		FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y
		
		Dim str
		str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
		ExecuteGlobal(str)
		str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
		ExecuteGlobal(str)
		
	End Sub
	
	' Legacy: just no op
	Public Property Let EndPoint(aInput)
		
	End Property
	
	Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
		Select Case aChooseArray
			Case "Polarity"
				ShuffleArrays PolarityIn, PolarityOut, 1
				PolarityIn(aIDX) = aX
				PolarityOut(aIDX) = aY
				ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity"
				ShuffleArrays VelocityIn, VelocityOut, 1
				VelocityIn(aIDX) = aX
				VelocityOut(aIDX) = aY
				ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef"
				ShuffleArrays YcoefIn, YcoefOut, 1
				YcoefIn(aIDX) = aX
				YcoefOut(aIDX) = aY
				ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
	End Sub
	
	Public Sub AddBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If IsEmpty(balls(x)) Then
				Set balls(x) = aBall
				Exit Sub
			End If
		Next
	End Sub
	
	Private Sub RemoveBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If TypeName(balls(x) ) = "IBall" Then
				If aBall.ID = Balls(x).ID Then
					balls(x) = Empty
					Balldata(x).Reset
				End If
			End If
		Next
	End Sub
	
	Public Sub Fire()
		Flipper.RotateToEnd
		processballs
	End Sub
	
	Public Property Get Pos 'returns % position a ball. For debug stuff.
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next
	End Property
	
	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub
	'Timer shutoff for polaritycorrect
	Private Function FlipperOn()
		If GameTime < FlipAt+TimeDelay Then
			FlipperOn = True
		End If
	End Function
	
	Public Sub PolarityCorrect(aBall)
		If FlipperOn() Then
			Dim tmp, BallPos, x, IDX, Ycoef
			Ycoef = 1
			
			'y safety Exit
			If aBall.VelY > -8 Then 'ball going down
				RemoveBall aBall
				Exit Sub
			End If
			
			'Find balldata. BallPos = % on Flipper
			For x = 0 To UBound(Balls)
				If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)								'find safety coefficient 'ycoef' data
				End If
			Next
			
			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)												'find safety coefficient 'ycoef' data
			End If
			
			'Velocity correction
			If Not IsEmpty(VelocityIn(0) ) Then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
				
				If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
				
				If Enabled Then aBall.Velx = aBall.Velx*VelCoef
				If Enabled Then aBall.Vely = aBall.Vely*VelCoef
			End If
			
			'Polarity Correction (optional now)
			If Not IsEmpty(PolarityIn(0) ) Then
				Dim AddX
				AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				
				If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
			End If
			If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	Dim x, aCount
	aCount = 0
	ReDim a(UBound(aArray) )
	For x = 0 To UBound(aArray)		'Shuffle objects in a temp array
		If Not IsEmpty(aArray(x) ) Then
			If IsObject(aArray(x)) Then
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	If offset < 0 Then offset = 0
	ReDim aArray(aCount-1+offset)		'Resize original array
	For x = 0 To aCount-1				'set objects back into original array
		If IsObject(a(x)) Then
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
	BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)		'Set up line via two points, no clamping. Input X, output Y
	Dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
	Public Property Let Data(aBall)
		With aBall
			x = .x
			y = .y
			z = .z
			velx = .velx
			vely = .vely
			velz = .velz
			id = .ID
			mass = .mass
			radius = .radius
		End With
	End Property
	Public Sub Reset()
		x = Empty
		y = Empty
		z = Empty
		velx = Empty
		vely = Empty
		velz = Empty
		id = Empty
		mass = Empty
		radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	Dim y 'Y output
	Dim L 'Line
	'find active line
	Dim ii
	For ii = 1 To UBound(xKeyFrame)
		If xInput <= xKeyFrame(ii) Then
			L = ii
			Exit For
		End If
	Next
	If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)		'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )
	
	If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )		 'Clamp lower
	If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )		'Clamp upper
	
	LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'	 - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'	 - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b
	   Dim BOT
	   BOT = GetBalls
	
	If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then
			For b = 0 To UBound(BOT)
				If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					Exit Sub
				End If
			Next
			For b = 0 To UBound(BOT)
				If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
					BOT(b).velx = BOT(b).velx / 1.3
					BOT(b).vely = BOT(b).vely - 0.5
				End If
			Next
		End If
	Else
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
	End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
	if velocity < 0.7 then exit sub		'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
		If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
		If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
		coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub
	



'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
	dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
	dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		End If
	ElseIf dx = 0 Then
		If dy = 0 Then
			Atn2 = 0
		Else
			Atn2 = Sgn(dy) * pi / 2
		End If
	End If
End Function

Function max(a,b)
	If a > b Then
		max = a
	Else
		max = b
	End If
End Function

Function min(a,b)
	If a > b Then
		min = b
	Else
		min = a
	End If
End Function


'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
	DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
	Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
	
	If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
	Case 0
		SOSRampup = 2.5
	Case 1
		SOSRampup = 6
	Case 2
		SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity
	
	Flipper.eostorque = EOST
	Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST * EOSReturn / FReturn
	
	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b, BOT
				BOT = GetBalls
		
		For b = 0 To UBound(BOT)
			If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper
	
	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup
			Flipper.endangle = FEndAngle - 3 * Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
		If FCount = 0 Then FCount = GameTime
		
		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
		If FState <> 3 Then
			Flipper.eostorque = EOST
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If
	End If
End Sub

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle)	'-1 for Right Flipper
	Dim LiveCatchBounce																														'If live catch is not perfect, it won't freeze ball totally
	Dim CatchTime
	CatchTime = GameTime - FCount
	
	If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
		If CatchTime <= LiveCatch * 0.5 Then												'Perfect catch only when catch time happens in the beginning of the window
			LiveCatchBounce = 0
		Else
			LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)		'Partial catch when catch happens a bit late
		End If
		
		If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
		ball.angmomx = 0
		ball.angmomy = 0
		ball.angmomz = 0
	Else
		If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
	End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************





'******************************************************
' 	ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
	RubbersD.dampen ActiveBall
	TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
	SleevesD.Dampen ActiveBall
	TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD				'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False	'shows info in textbox "TBPout"
RubbersD.Print = False	  'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1		 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967	'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64	   'there's clamping so interpolate up to 56 at least

Dim SleevesD	'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False	'shows info in textbox "TBPout"
SleevesD.Print = False	  'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn   'tbpOut.text
	Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
	End Sub
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If GameTime > 100 Then Report
	End Sub
	
	Public Sub Dampen(aBall)
		If threshold Then
			If BallSpeed(aBall) < threshold Then Exit Sub
		End If
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
		"actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
		If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)
		
		aBall.velx = aBall.velx * coef
		aBall.vely = aBall.vely * coef
		If debugOn Then TBPout.text = str
	End Sub
	
	Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
			aBall.velx = aBall.velx * coef
			aBall.vely = aBall.vely * coef
		End If
	End Sub
	
	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		Dim x
		For x = 0 To UBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
		Next
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
	Public ballvel, ballvelx, ballvely
	
	Private Sub Class_Initialize
		ReDim ballvel(0)
		ReDim ballvelx(0)
		ReDim ballvely(0)
	End Sub
	
	Public Sub Update()	'tracks in-ball-velocity
		Dim str, b, AllBalls, highestID
		allBalls = GetBalls
		
		For Each b In allballs
			If b.id >= HighestID Then highestID = b.id
		Next
		
		If UBound(ballvel) < highestID Then ReDim ballvel(highestID)	'set bounds
		If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)	'set bounds
		If UBound(ballvely) < highestID Then ReDim ballvely(highestID)	'set bounds
		
		For Each b In allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
Sub RDampen_Timer
	Cor.Update
End Sub


'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
' 	ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1	  '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7	 'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
	Dim zMultiplier, vel, vratio
	If TargetBouncerEnabled = 1 And aball.z < 30 Then
		'   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		vel = BallSpeed(aBall)
		If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
		Select Case Int(Rnd * 6) + 1
			Case 1
				zMultiplier = 0.2 * defvalue
			Case 2
				zMultiplier = 0.25 * defvalue
			Case 3
				zMultiplier = 0.3 * defvalue
			Case 4
				zMultiplier = 0.4 * defvalue
			Case 5
				zMultiplier = 0.45 * defvalue
			Case 6
				zMultiplier = 0.5 * defvalue
		End Select
		aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
		aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
		aBall.vely = aBall.velx * vratio
		'   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		'   debug.print "conservation check: " & BallSpeed(aBall)/vel
	End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
	TargetBouncer ActiveBall, 1
End Sub



'******************************************************
'	ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'	 - On the table, add the endpoint primitives that define the two ends of the Slingshot
'	 - Initialize the SlingshotCorrection objects in InitSlingCorrection
'	 - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
	LS.Object = LeftSlingshot
	LS.EndPoint1 = EndPoint1LS
	LS.EndPoint2 = EndPoint2LS
	
	RS.Object = RightSlingshot
	RS.EndPoint1 = EndPoint1RS
	RS.EndPoint2 = EndPoint2RS
	
	'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
	' These values are best guesses. Retune them if needed based on specific table research.
	AddSlingsPt 0, 0.00, - 4
	AddSlingsPt 1, 0.45, - 7
	AddSlingsPt 2, 0.48,	0
	AddSlingsPt 3, 0.52,	0
	AddSlingsPt 4, 0.55,	7
	AddSlingsPt 5, 1.00,	4
End Sub

Sub AddSlingsPt(idx, aX, aY)		'debugger wrapper for adjusting flipper script In-game
	Dim a
	a = Array(LS, RS)
	Dim x
	For Each x In a
		x.addpoint idx, aX, aY
	Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
'	dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
'	dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
'	dim rx, ry
'	rx = x*dCos(angle) - y*dSin(angle)
'	ry = x*dSin(angle) + y*dCos(angle)
'	RotPoint = Array(rx,ry)
'End Function

Class SlingshotCorrection
	Public DebugOn, Enabled
	Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2
	
	Public ModIn, ModOut
	
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
		Enabled = True
	End Sub
	
	Public Property Let Object(aInput)
		Set Slingshot = aInput
	End Property
	
	Public Property Let EndPoint1(aInput)
		SlingX1 = aInput.x
		SlingY1 = aInput.y
	End Property
	
	Public Property Let EndPoint2(aInput)
		SlingX2 = aInput.x
		SlingY2 = aInput.y
	End Property
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If GameTime > 100 Then Report
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
	
	
	Public Sub VelocityCorrect(aBall)
		Dim BallPos, XL, XR, YL, YR
		
		'Assign right and left end points
		If SlingX1 < SlingX2 Then
			XL = SlingX1
			YL = SlingY1
			XR = SlingX2
			YR = SlingY2
		Else
			XL = SlingX2
			YL = SlingY2
			XR = SlingX1
			YR = SlingY1
		End If
		
		'Find BallPos = % on Slingshot
		If Not IsEmpty(aBall.id) Then
			If Abs(XR - XL) > Abs(YR - YL) Then
				BallPos = PSlope(aBall.x, XL, 0, XR, 1)
			Else
				BallPos = PSlope(aBall.y, YL, 0, YR, 1)
			End If
			If BallPos < 0 Then BallPos = 0
			If BallPos > 1 Then BallPos = 1
		End If
		
		'Velocity angle correction
		If Not IsEmpty(ModIn(0) ) Then
			Dim Angle, RotVxVy
			Angle = LinearEnvelope(BallPos, ModIn, ModOut)
			'   debug.print " BallPos=" & BallPos &" Angle=" & Angle
			'   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
			RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
			If Enabled Then aBall.Velx = RotVxVy(0)
			If Enabled Then aBall.Vely = RotVxVy(1)
			'   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
			'   debug.print " "
		End If
	End Sub
End Class

'==================================================================================================================

Sub Trigger010_Hit() ' middle ramp booster
        Debug "10 Y Value: " &ActiveBall.VelY 
Debug "----------------------------"
    If ActiveBall.VelY < -8 And ActiveBall.VelY > -10 Then ActiveBall.Vely = -15: Debug "10 - Boosted"
End Sub

Sub Trigger019_Hit() ' top booster
        Debug "19 X Value: " &ActiveBall.VelX 
        Debug "19 Y Value " &ActiveBall.VelY 
Debug "----------------------------"
    If ActiveBall.VelX < -1  Then ActiveBall.VelX = -13: Debug "19 - Boosted"
End Sub

Sub Trigger020_Hit() '  lower booster
        Debug "20 Y Value: " &ActiveBall.VelY 
Debug "----------------------------"
    If ActiveBall.VelY < -8 And ActiveBall.VelY > -10 Then ActiveBall.Vely = -15: Debug "20 - Boosted"
End Sub


'**************************
'   PinUp Player USER Config
'**************************

dim PuPDMDDriverType: PuPDMDDriverType=0   ' 0=LCD DMD, 1=RealDMD 2=FULLDMD (large/High LCD)
dim useRealDMDScale : useRealDMDScale=1    ' 0 or 1 for RealDMD scaling.  Choose which one you prefer.
dim useDMDVideos    : useDMDVideos=true   ' true or false to use DMD splash videos.
Dim pGameName       : pGameName="TheMatrix"  'pupvideos foldername, probably set to cGameName in realworld







'********************* START OF PUPDMD FRAMEWORK v1.0 *************************
'******************** DO NOT MODIFY STUFF BELOW   THIS LINE!!!! ***************
'******************************************************************************
'*****   Create a PUPPack within PUPPackEditor for layout config!!!  **********
'******************************************************************************
'
'
'  Quick Steps:
'      1>  create a folder in PUPVideos with Starter_PuPPack.zip and call the folder "yourgame"
'      2>  above set global variable pGameName="yourgame"
'      3>  copy paste the settings section above to top of table script for user changes.
'      4>  on Table you need to create ONE timer only called pupDMDUpdate and set it to 250 ms enabled on startup.
'      5>  go to your table1_init or table first startup function and call PUPINIT function
'      6>  Go to bottom on framework here and setup game to call the appropriate events like pStartGame (call that in your game code where needed)...etc
'      7>  attractmodenext at bottom is setup for you already,  just go to each case and add/remove as many as you want and setup the messages to show.  
'      8>  Have fun and use pDMDDisplay(xxxx)  sub all over where needed.  remember its best to make a bunch of mp4 with text animations... looks the best for sure!
'
'
'Note:  for *Future Pinball* "pupDMDupdate_Timer()" timer needs to be renamed to "pupDMDupdate_expired()"  and then all is good.
'       and for future pinball you need to add the follow lines near top
'Need to use BAM and have com idll enabled.
'				Dim icom : Set icom = xBAM.Get("icom") ' "icom" is name of "icom.dll" in BAM\Plugins dir
'				if icom is Nothing then MSGBOX "Error cannot run without icom.dll plugin"
'				Function CreateObject(className)       
'   					Set CreateObject = icom.CreateObject(className)   
'				End Function


Const HasPuP = True   'dont set to false as it will break pup

Const pTopper=0
Const pDMD=1
Const pBackglass=2
Const pPlayfield=3
Const pMusic=4
Const pMusic2=5
Const pCallouts=6
Const pBackglass2=7
Const pTopper2=8
Const pPopUP=9
Const pPopUP2=10


'pages
Const pDMDBlank=0
Const pScores=1
Const pBigLine=2
Const pThreeLines=3
Const pTwoLines=4
Const pTargerLetters=5

'dmdType
Const pDMDTypeLCD=0
Const pDMDTypeReal=1
Const pDMDTypeFULL=2






dim PUPDMDObject  'for realtime mirroring.
Dim pDMDlastchk: pDMDLastchk= -1    'performance of updates
Dim pDMDCurPage: pDMDCurPage= 0     'default page is empty.
Dim pInAttract : pInAttract=false   'pAttract mode




'*************  starts PUP system,  must be called AFTER b2s/controller running so put in last line of table1_init
Sub PuPInit

'Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")   
'PuPlayer.B2SInit "", pGameName
PuPlayer.LabelInit pBackglass

pSetPageLayouts

pDMDSetPage(1)   'set blank text overlay page.


pDMDStartUP				 ' firsttime running for like an startup video..

	if Scorbitactive then 
'		if Scorbit.DoInit(4163, "PupOverlays", myVersion, "btlc-vpin") then 	' Staging
		if Scorbit.DoInit(4239, "PupOverlays", myVersion, "matrix-vpin") then 	' Prod
			tmrScorbit.Interval=2000
			tmrScorbit.UserValue = 0
			tmrScorbit.Enabled=True 
			Scorbit.UploadLog = ScorbitUploadLog
		End if 
	End if 
debug4 "In PUPINIT"
'PuPlayer.LabelSet pBackglass,"FinalScore","TESTING MESSAGING",1,""
End Sub 'end PUPINIT

sub CheckPairing
		Debug4 "In check pairing"

	if (Scorbit.bNeedsPairing) then 
		PuPEvent 998

		Debug4 "Scorbit Needs Pairing"
		if ScorbitQRLeft = 0 Then 
			pBackglassLabelSetSizeImage "ScorbitQR1",16,28
			pBackglassLabelSetPos "ScorbitQR1",9,84
		Else
			pBackglassLabelSetSizeImage "ScorbitQR1",16,28
			pBackglassLabelSetPos "ScorbitQR1",91,84
		End If
		pBackglasslabelshow "ScorbitQR1"
		pBackglasslabelshow "ScorbitQRIcon1"
		DelayQRClaim.Interval=6000
		DelayQRClaim.Enabled=True
	end if
End sub

Sub hideScorbit
	if usePUP Then
		pBackglasslabelhide "ScorbitQR1"
		pBackglasslabelhide "ScorbitQRIcon1"
		pBackglasslabelhide "ScorbitQR2"
		pBackglasslabelhide "ScorbitQRIcon2"
	end if
End Sub




'PinUP Player DMD Helper Functions

Sub pBackglassLabelHide(labName)
PuPlayer.LabelSet pBackglass,labName,"",0,""   
end sub

Sub pBackglassLabelShow(labName)
PuPlayer.LabelSet pBackglass,labName,"",1,""   
end sub

sub pBackglassLabelSetPos(labName, xpos, ypos)
   PuPlayer.LabelSet pBackglass,labName,"",1,"{'mt':2,'xpos':"&xpos& ",'ypos':"&ypos&"}"    
end sub

sub pBackglassLabelSetSizeImage(labName, lWidth, lHeight)
   PuPlayer.LabelSet pBackglass,labName,"",1,"{'mt':2,'width':"& lWidth & ",'height':"&lHeight&"}" 
end sub

Sub pDMDScrollBig(msgText,timeSec,mColor)
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
end sub

Sub pDMDScrollBigV(msgText,timeSec,mColor)
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'yps':1,'ype':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
end sub


Sub pDMDSplashScore(msgText,timeSec,mColor)
PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':1,'fq':250,'len':"& (timeSec*1000) &",'fc':" & mColor & "}"
end Sub

Sub pDMDSplashScoreScroll(msgText,timeSec,mColor)
PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':"& (timeSec*1000) &", 'mlen':"& (timeSec*1000) &",'tt':0, 'fc':" & mColor & "}"
end Sub

Sub pDMDZoomBig(msgText,timeSec,mColor)  'new Zoom
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':3,'hstart':5,'hend':80,'len':" & (timeSec*1000) & ",'mlen':" & (timeSec*500) & ",'tt':5,'fc':" & mColor & "}"
end sub

Sub pDMDTargetLettersInfo(msgText,msgInfo, timeSec)  'msgInfo = '0211'  0= layer 1, 1=layer 2, 2=top layer3.
'this function is when you want to hilite spelled words.  Like B O N U S but have O S hilited as already hit markers... see example.
PuPlayer.LabelShowPage pDMD,5,timeSec,""  'show page 5
Dim backText
Dim middleText
Dim flashText
Dim curChar
Dim i
Dim offchars:offchars=0
Dim spaces:spaces=" "  'set this to 1 or more depends on font space width.  only works with certain fonts
                          'if using a fixed font width then set spaces to just one space.

For i=1 To Len(msgInfo)
    curChar="" & Mid(msgInfo,i,1)
    if curChar="0" Then
            backText=backText & Mid(msgText,i,1)
            middleText=middleText & spaces
            flashText=flashText & spaces          
            offchars=offchars+1
    End If
    if curChar="1" Then
            backText=backText & spaces
            middleText=middleText & Mid(msgText,i,1)
            flashText=flashText & spaces
    End If
    if curChar="2" Then
            backText=backText & spaces
            middleText=middleText & spaces
            flashText=flashText & Mid(msgText,i,1)
    End If   
Next 

if offchars=0 Then 'all litup!... flash entire string
   backText=""
   middleText=""
   FlashText=msgText
end if  

PuPlayer.LabelSet pDMD,"Back5"  ,backText  ,1,""
PuPlayer.LabelSet pDMD,"Middle5",middleText,1,""
PuPlayer.LabelSet pDMD,"Flash5" ,flashText ,0,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) & "}"   
end Sub


Sub pDMDSetPage(pagenum)    
    PuPlayer.LabelShowPage pBackglass,pagenum,0,""   'set page to blank 0 page if want off
    PDMDCurPage=pagenum
end Sub

Sub pHideOverlayText(pDisp)
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& pDisp &", ""FN"": 34 }"             'hideoverlay text during next videoplay on DMD auto return
end Sub



Sub pDMDShowLines3(msgText,msgText2,msgText3,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,3,timeSec,""
PuPlayer.LabelSet pDMD,"Splash3a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash3b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash3c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowLines2(msgText,msgText2,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,4,timeSec,""
PuPlayer.LabelSet pDMD,"Splash4a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash4b",msgText2,vis,pLine2Ani
end Sub

Sub pDMDShowCounter(msgText,msgText2,msgText3,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,6,timeSec,""
PuPlayer.LabelSet pDMD,"Splash6a",msgText,vis, pLine1Ani
PuPlayer.LabelSet pDMD,"Splash6b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash6c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowBig(msgText,timeSec, mColor)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,vis,pLine1Ani
end sub


Sub pDMDShowHS(msgText,msgText2,msgText3,timeSec) 'High Score
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,7,timeSec,""
PuPlayer.LabelSet pDMD,"Splash7a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash7b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash7c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDSetBackFrame(fname)
  PuPlayer.playlistplayex pDMD,"PUPFrames",fname,0,1    
end Sub

Sub pDMDStartBackLoop(fPlayList,fname)
  PuPlayer.playlistplayex pDMD,fPlayList,fname,0,1
  PuPlayer.SetBackGround pDMD,1
end Sub

Sub pDMDStopBackLoop
  PuPlayer.SetBackGround pDMD,0
  PuPlayer.playstop pDMD
end Sub


Dim pNumLines

'Theme Colors for Text (not used currenlty,  use the |<colornum> in text labels for colouring.
Dim SpecialInfo
Dim pLine1Color : pLine1Color=8454143  
Dim pLine2Color : pLine2Color=8454143
Dim pLine3Color :  pLine3Color=8454143
Dim curLine1Color: curLine1Color=pLine1Color  'can change later
Dim curLine2Color: curLine2Color=pLine2Color  'can change later
Dim curLine3Color: curLine3Color=pLine3Color  'can change later


Dim pDMDCurPriority: pDMDCurPriority =-1
Dim pDMDDefVolume: pDMDDefVolume = 0   'default no audio on pDMD

Dim pLine1
Dim pLine2
Dim pLine3
Dim pLine1Ani
Dim pLine2Ani
Dim pLine3Ani

Dim PriorityReset:PriorityReset=-1
DIM pAttractReset:pAttractReset=-1
DIM pAttractBetween: pAttractBetween=2000 '1 second between calls to next attract page
DIM pDMDVideoPlaying: pDMDVideoPlaying=false


'************************ where all the MAGIC goes,  pretty much call this everywhere  ****************************************
'*************************                see docs for examples                ************************************************
'****************************************   DONT TOUCH THIS CODE   ************************************************************

Sub pupDMDDisplay(pEventID, pText, VideoName,TimeSec, pAni,pPriority)
' pEventID = reference if application,  
' pText = "text to show" separate lines by ^ in same string
' VideoName "gameover.mp4" will play in background  "@gameover.mp4" will play and disable text during gameplay.
' also global variable useDMDVideos=true/false if user wishes only TEXT
' TimeSec how long to display msg in Seconds
' animation if any 0=none 1=Flasher
' also,  now can specify color of each line (when no animation).  "sometext|12345"  will set label to "sometext" and set color to 12345

DIM curPos
if pDMDCurPriority>pPriority then Exit Sub  'if something is being displayed that we don't want interrupted.  same level will interrupt.
pDMDCurPriority=pPriority
if timeSec=0 then timeSec=1 'don't allow page default page by accident


pLine1=""
pLine2=""
pLine3=""
pLine1Ani=""
pLine2Ani=""
pLine3Ani=""


if pAni=1 Then  'we flashy now aren't we
pLine1Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"  
pLine2Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"  
pLine3Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"  
end If

curPos=InStr(pText,"^")   'Lets break apart the string if needed
if curPos>0 Then 
   pLine1=Left(pText,curPos-1) 
   pText=Right(pText,Len(pText) - curPos)
   
   curPos=InStr(pText,"^")   'Lets break apart the string
   if curPOS>0 Then
      pLine2=Left(pText,curPos-1) 
      pText=Right(pText,Len(pText) - curPos)

      curPos=InStr("^",pText)   'Lets break apart the string   
      if curPos>0 Then
         pline3=Left(pText,curPos-1) 
      Else 
        if pText<>"" Then pline3=pText 
      End if 
   Else 
      if pText<>"" Then pLine2=pText
   End if    
Else 
  pLine1=pText  'just one line with no break 
End if


'lets see how many lines to Show
pNumLines=0
if pLine1<>"" then pNumLines=pNumlines+1
if pLine2<>"" then pNumLines=pNumlines+1
if pLine3<>"" then pNumLines=pNumlines+1

if pDMDVideoPlaying Then 
			PuPlayer.playstop pDMD
			pDMDVideoPlaying=False
End if


if (VideoName<>"") and (useDMDVideos) Then  'we are showing a splash video instead of the text.
    
    PuPlayer.playlistplayex pDMD,"DMDSplash",VideoName,pDMDDefVolume,pPriority  'should be an attract background (no text is displayed)
    pDMDVideoPlaying=true
end if 'if showing a splash video with no text




if StrComp(pEventID,"shownum",1)=0 Then              'check eventIDs
    pDMDShowCounter pLine1,pLine2,pLine3,timeSec
Elseif StrComp(pEventID,"target",1)=0 Then              'check eventIDs
    pDMDTargetLettersInfo pLine1,pLine2,timeSec
Elseif StrComp(pEventID,"highscore",1)=0 Then              'check eventIDs
    pDMDShowHS pLine1,pLine2,pline3,timeSec
Elseif (pNumLines=3) Then                'depends on # of lines which one to use.  pAni=1 will flash.
    pDMDShowLines3 pLine1,pLine2,pLine3,TimeSec
Elseif (pNumLines=2) Then
    pDMDShowLines2 pLine1,pLine2,TimeSec
Elseif (pNumLines=1) Then
    pDMDShowBig pLine1,timeSec, curLine1Color
Else
    pDMDShowBig pLine1,timeSec, curLine1Color
End if

PriorityReset=TimeSec*1000
End Sub 'pupDMDDisplay message

Sub pupDMDupdate_Timer()
	pUpdateScores    

    if PriorityReset>0 Then  'for splashes we need to reset current prioirty on timer
       PriorityReset=PriorityReset-pupDMDUpdate.interval
       if PriorityReset<=0 Then 
            pDMDCurPriority=-1            
            if pInAttract then pAttractReset=pAttractBetween ' pAttractNext  call attract next after 1 second
			pDMDVideoPlaying=false			
			End if
    End if

    if pAttractReset>0 Then  'for splashes we need to reset current prioirty on timer
       pAttractReset=pAttractReset-pupDMDUpdate.interval
       if pAttractReset<=0 Then 
            pAttractReset=-1            
            if pInAttract then pAttractNext
			End if
    end if 
End Sub


'********************* END OF PUPDMD FRAMEWORK v1.0 *************************
'******************** DO NOT MODIFY STUFF ABOVE THIS LINE!!!! ***************
'****************************************************************************

'*****************************************************************
'   **********  PUPDMD  MODIFY THIS SECTION!!!  ***************
'PUPDMD Layout for each Table1
'Setup Pages.  Note if you use fonts they must be in FONTS folder of the pupVideos\tablename\FONTS  "case sensitive exact naming fonts!"
'*****************************************************************

Sub pSetPageLayouts

DIM dmddef
DIM dmdalt
DIM dmdscr
DIM dmdfixed

'labelNew <screen#>, <Labelname>, <fontName>,<size%>,<colour>,<rotation>,<xalign>,<yalign>,<xpos>,<ypos>,<PageNum>,<visible>
'***********************************************************************'
'<screen#>, in standard we’d set this to pDMD ( or 1)
'<Labelname>, your name of the label. keep it short no spaces (like 8 chars) although you can call it anything really. When setting the label you will use this labelname to access the label.
'<fontName> Windows font name, this must be exact match of OS front name. if you are using custom TTF fonts then double check the name of font names.
'<size%>, Height as a percent of display height. 20=20% of screen height.
'<colour>, integer value of windows color.
'<rotation>, degrees in tenths   (900=90 degrees)
'<xAlign>, 0= horizontal left align, 1 = center horizontal, 2= right horizontal
'<yAlign>, 0 = top, 1 = center, 2=bottom vertical alignment
'<xpos>, this should be 0, but if you want to ‘force’ a position you can set this. it is a % of horizontal width. 20=20% of screen width.
'<ypos> same as xpos.
'<PageNum> IMPORTANT… this will assign this label to this ‘page’ or group.
'<visible> initial state of label. visible=1 show, 0 = off.

	pupCreateLabelImage "ScorbitQRicon1","PuPOverlays\\QRcodeS.png",50,30,34,60,1,0
	pupCreateLabelImage "ScorbitQR1","PuPOverlays\\QRcode.png",50,30,34,60,1,0

	pupCreateLabelImage "ScorbitQRicon2","PuPOverlays\\QRcodeB.png",50,30,34,60,1,0
	pupCreateLabelImage "ScorbitQR2","PuPOverlays\\QRclaim.png",50,30,34,60,1,0

	PuPlayer.LabelNew pBackglass,"FinalScore",		dmddef,	15,255	,0,1,1,50,5,1,0



if PuPDMDDriverType=pDMDTypeReal Then 'using RealDMD Mirroring.  **********  128x32 Real Color DMD  
	dmdalt="PKMN Pinball"
    dmdfixed="Instruction"
    dmdscr="Impact"    'main scorefont
	dmddef="Zig"

	'Page 1 (default score display)
  		 PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,95,0,1,0
		 PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,21,33023   ,1,0,0,15,0,1,0
		 PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,21,33023   ,1,2,0,85,0,1,0
		 PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
		 PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,8454143   ,0,1,1, 0,0,1,0


	'Page 2 (default Text Splash 1 Big Line)
		 PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

	'Page 3 (default Text Splash 2 and 3 Lines)
		 PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
		 PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
	     PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,55,3,0


	'Page 4 (2 Line Gameplay DMD)
		 PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
	     PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

	'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
		PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

	'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
		PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
		PuPlayer.LabelNew pDMD,"Splash6b",dmddef,50,33023,0,1,0,60,0,6,0
		PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

 	'Page 7 (Show High Scores Fixed Fonts)
		PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
		PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
		PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0


END IF  ' use PuPDMDDriver

if PuPDMDDriverType=pDMDTypeLCD THEN  'Using 4:1 Standard ratio LCD PuPDMD  ************ lcd **************

	'dmddef="Impact"
	dmdalt="PKMN Pinball"    
    dmdfixed="Instruction"
	dmdscr="Impact"  'main score font
	dmddef="Impact"

	'Page 1 (default score display)
		PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,95,0,1,0
		PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,20,33023   ,1,0,0,15,0,1,0
		PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,20,33023   ,1,2,0,85,0,1,0
		PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
		PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,8454143   ,0,1,1, 0,0,1,0


	'Page 2 (default Text Splash 1 Big Line)
		PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

	'Page 3 (default Text 3 Lines)
		PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
		PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
		PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,57,3,0


	'Page 4 (default Text 2 Line)
		PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
		PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

	'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
		PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

	'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
		PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
		PuPlayer.LabelNew pDMD,"Splash6b",dmddef,50,33023,0,1,0,60,0,6,0
		PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

	'Page 7 (Show High Scores Fixed Fonts)
		PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
		PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
		PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0


END IF  ' use PuPDMDDriver

if PuPDMDDriverType=pDMDTypeFULL THEN  'Using FULL BIG LCD PuPDMD  ************ lcd **************

	'dmddef="Impact"
	dmdalt="PKMN Pinball"    
    dmdfixed="Instruction"
	dmdscr="Impact"  'main score font
	dmddef="Impact"

	'Page 1 (default score display)
		PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,95,0,1,0
		PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,20,33023   ,1,0,0,15,0,1,0
		PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,20,33023   ,1,2,0,85,0,1,0
		PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
		PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,8454143   ,0,1,1, 0,0,1,0		


	'Page 2 (default Text Splash 1 Big Line)
		PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

	'Page 3 (default Text 3 Lines)
		PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
		PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
		PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,57,3,0


	'Page 4 (default Text 2 Line)
		PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
		PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

	'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
		PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

	'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
		PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
		PuPlayer.LabelNew pDMD,"Splash6b",dmddef,50,33023,0,1,0,60,0,6,0
		PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

	'Page 7 (Show High Scores Fixed Fonts)
		PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
		PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
		PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0


END IF  ' use PuPDMDDriver




end Sub 'page Layouts


'*****************************************************************
'        PUPDMD Custom SUBS/Events for each Table1
'     **********    MODIFY THIS SECTION!!!  ***************
'*****************************************************************
'
'
'  we need to somewhere in code if applicable
'
'   call pDMDStartGame,pDMDStartBall,pGameOver,pAttractStart
'
'
'
'
'


Sub pDMDStartGame
pInAttract=false
pDMDSetPage(pScores)   'set blank text overlay page.

end Sub


Sub pDMDStartBall
end Sub

Sub pDMDGameOver
pAttractStart
end Sub

Sub pAttractStart
pDMDSetPage(pDMDBlank)   'set blank text overlay page.
'pCurAttractPos=0
'pInAttract=True          'Startup in AttractMode
'pAttractNext
end Sub

Sub pDMDStartUP
	vpmtimer.addtimer 2500, "CheckPairing ' "
 'pupDMDDisplay "attract","Welcome","@welcome.mp4",2,0,10
 'pInAttract=true
end Sub

DIM pCurAttractPos: pCurAttractPos=0


'********************** gets called auto each page next and timed already in DMD_Timer.  make sure you use pupDMDDisplay or it wont advance auto.
Sub pAttractNext
pCurAttractPos=pCurAttractPos+1

  Select Case pCurAttractPos

  Case 1 pupDMDDisplay "attract","Attract^1","",5,1,10
  Case 2 pupDMDDisplay "attract","Attract^2","",3,0,10
  Case 3 pupDMDDisplay "attract","Attract^3","",2,0,10
  Case 4 pupDMDDisplay "attract","Attract^4","",3,1,10
  Case 5 pupDMDDisplay "attract","Attract^5","",1,0,10
  Case 6 pupDMDDisplay "attract","Attract^6","",3,1,10
  Case 7 pupDMDDisplay "attract","Attract^7","",2,0,10
  Case 8 pupDMDDisplay "attract","Attract^8","",1,0,10
  Case 9 pupDMDDisplay "attract","Attract^9","",1,1,10
  Case 10 pupDMDDisplay "attract","Attract^10","",3,1,10
  Case Else
    pCurAttractPos=0
    pAttractNext 'reset to beginning
  end Select

end Sub


'************************ called during gameplay to update Scores ***************************
Dim CurTestScore:CurTestScore=100000
Sub pUpdateScores  'call this ONLY on timer 300ms is good enough
if pDMDCurPage <> pScores then Exit Sub

puPlayer.LabelSet pDMD,"CurScore","" & FormatNumber(CurTestScore,0),1,""
puPlayer.LabelSet pDMD,"Play1","play " & 2,1,""
puPlayer.LabelSet pDMD,"Ball","ball "  & 2,1,""
end Sub

Sub pupCreateLabel(lName, lValue, lFont, lSize, lColor, xpos, ypos, alignX, alignY, pagenum, lvis)
	PuPlayer.LabelNew pBackglass,lName ,lFont,lSize,lColor,0,alignX,alignY,1,1,pagenum,lvis
	PuPlayer.LabelSet pBackglass,lName,lValue,lvis,"{'mt':2,'xpos':"& xpos & ",'ypos':"&ypos&"}"
end Sub

Sub pupCreateLabelImage(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
	PuPlayer.LabelNew pBackglass,lName ,"",50,RGB(100,100,100),0,1,1,1,1,pagenum,lvis
	PuPlayer.LabelSet pBackglass,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&"}"
end Sub

Sub pupCreateLabelImageDMD(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
	PuPlayer.LabelNew pDMD,lName ,"",50,RGB(100,100,100),0,1,1,1,1,pagenum,lvis
	PuPlayer.LabelSet pDMD,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&"}"
end Sub

Sub pupPoPCreateLabelImage(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
	PuPlayer.LabelNew pPopUP,lName ,"",50,RGB(100,100,100),0,1,1,1,1,pagenum,lvis
	PuPlayer.LabelSet pPopUP,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&"}"
end Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  SCORBIT Interface
' To Use:
' 1) Define a timer tmrScorbit
' 2) Call DoInit at the end of PupInit or in Table Init if you are nto using pup with the appropriate parameters
'     Replace 389 with your TableID from Scorbit 
'     Replace GRWvz-MP37P from your table on OPDB - eg: https://opdb.org/machines/2103
'		if Scorbit.DoInit(389, "PupOverlays", "1.0.0", "GRWvz-MP37P") then 
'			tmrScorbit.Interval=2000
'			tmrScorbit.UserValue = 0
'			tmrScorbit.Enabled=True 
'		End if 
' 3) Customize helper functions below for different events if you want or make your own 
' 4) Call 
'		DoInit - After Pup/Screen is setup (PuPInit)
'		StartSession - When a game starts (ResetForNewGame)
'		StopSession - When the game is over (Table1_Exit, EndOfGame)
'		SendUpdate - called when Score Changes (AddScore)
'			SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
'			Example:  Scorbit.SendUpdate Score(0), Score(1), Score(2), Score(3), Balls, CurrentPlayer+1, PlayersPlayingGame
'		SetGameMode - When different game events happen like starting a mode, MB etc.  (ScorbitBuildGameModes helper function shows you how)
' 5) Drop the binaries sQRCode.exe and sToken.exe in your Pup Root so we can create session tokens and QRCodes.
'	- Drop QRCode Images (QRCodeS.png, QRcodeB.png) in yur pup PuPOverlays if you want to use those 
' 6) Callbacks 
'		Scorbit_Paired   	- Called when machine is successfully paired.  Hide QRCode and play a sound 
'		Scorbit_PlayerClaimed	- Called when player is claimed.  Hide QRCode, play a sound and display name 
'		ScorbitClaimQR		- Call before/after plunge (swPlungerRest_Hit, swPlungerRest_UnHit)
' 7) Other 
'		Set Pair QR Code	- During Attract
'			if (Scorbit.bNeedsPairing) then 
'				PuPlayer.LabelSet pDMDFull, "ScorbitQR_a", "PuPOverlays\\QRcode.png",1,"{'mt':2,'width':32, 'height':64,'xalign':0,'yalign':0,'ypos':5,'xpos':5}"
'				PuPlayer.LabelSet pDMDFull, "ScorbitQRIcon_a", "PuPOverlays\\QRcodeS.png",1,"{'mt':2,'width':36, 'height':85,'xalign':0,'yalign':0,'ypos':3,'xpos':3,'zback':1}"
'			End if 
'		Set Player Names 	- Wherever it makes sense but I do it here: (pPupdateScores)
'		   if ScorbitActive then 
'			if Scorbit.bSessionActive then
'				PlayerName=Scorbit.GetName(CurrentPlayer+1)
'				if PlayerName="" then PlayerName= "Player " & CurrentPlayer+1 
'			End if 
'		   End if 
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' TABLE CUSTOMIZATION START HERE 

Sub Scorbit_Paired()								' Scorbit callback when new machine is paired 
debug4 "Scorbit PAIRED"
	PlaySound "scorbit_login"
	PuPEvent 1000
	pbackglasslabelhide "ScorbitQR1"
	pbackglasslabelhide "ScorbitQRIcon1"
End Sub 

Sub Scorbit_PlayerClaimed(PlayerNum, PlayerName)	' Scorbit callback when QR Is Claimed 
debug4 "Scorbit LOGIN"
	PlaySound "scorbit_login"
'	PlaySoundGame "scorbit_login", 0, 1, 0 ,0, 0, 1
	PuPEvent 1000
	ScorbitClaimQR(False)
End Sub 


Sub ScorbitClaimQR(bShow)	
debug4 "In ScorbitClaimQR: " &bShow					'  Show QRCode on first ball for users to claim this position
	if Scorbit.bSessionActive=False then Exit Sub 
	if ScorbitShowClaimQR=False then Exit Sub
	if Scorbit.bNeedsPairing then exit sub 

	if bShow and balls=1 and bGameInPlay and Scorbit.GetName(CurrentPlayer+1)="" then 
		if ScorbitQRLeft=0 then ' Dispaly in lower left
			PuPEvent 999
			pBackglassLabelSetSizeImage "ScorbitQR2",16,28
			pBackglassLabelSetPos "ScorbitQR2",9,84
			pbackglasslabelshow "ScorbitQRIcon2"
			pbackglasslabelshow "ScorbitQR2"
		else  ' display in lower right
			Debug4 "Showing Claim QR"
			pBackglassLabelSetSizeImage "ScorbitQR2",16,28
			pBackglassLabelSetPos "ScorbitQR2",91,84
			
			PuPEvent 999
'			BallHandlingQueue.Add "ScorbitQR2","pbackglasslabelshow ""ScorbitQR2"" ",44,0,0,0,0,True
			pbackglasslabelshow "ScorbitQR2"
			pbackglasslabelshow "ScorbitQRIcon2"
		End if 
	Else 
		PuPEvent 1000
		debug4 "Hiding QR claim"
		pbackglasslabelhide "ScorbitQRIcon2"
		pbackglasslabelhide "ScorbitQR2"
	End if 
End Sub 

Sub StopScorbit
	Scorbit.StopSession Score(0), Score(1), Score(2), Score(3), PlayersPlayingGame   ' Stop updateing scores
End Sub

Sub ScorbitBuildGameModes()		' Custom function to build the game modes for better stats 
	dim GameModeStr
	if Scorbit.bSessionActive=False then Exit Sub 
	GameModeStr="NA:"

	if BallsRemaining(CurrentPlayer) <= 0 Then	'no balls left
		GameModeStr="NA{red}:YOU FAILED!!!"
	Else										'game on
		Debug4 "AMCombo Value - " 
		Select Case Mode(CurrentPlayer,0)
			Case 1:
				GameModeStr="NA{yellow}:FOLLOW WHITE RABBIT COMPLETED"
			Case 2:
				GameModeStr="NA{green}:AGENT SMITH FIGHT COMPLETED"
			Case 3:
				GameModeStr="NA{red}:ORACLE MEETING COMPLETED"
			Case 4:
				GameModeStr="NA{orange}:KILL SMITH CLONES COMPLETED"
			Case 5:
				GameModeStr="NA{blue}:SERAPH FIGHT COMPLETED"
			Case 6:
				GameModeStr="NA{white}:FIND KEYMAKER COMPLETED"
			Case 7:
				GameModeStr="NA{yellow}:MEROVINGIAN COMPLETED"
			Case 8:
				GameModeStr="NA{green}:THE TWINS CHASE COMPLETED"
			Case 9:
				GameModeStr="NA{blue}:MEET THE ARCHITECT COMPLETED"
			Case 10:
				GameModeStr="NA{white}:THE TRAINMAN COMPLETED"
			Case 11:
				GameModeStr="NA{yellow}:DEUS EX MACHINA COMPLETED"
			Case 12:
				GameModeStr="NA{red}:KILL CYPHER COMPLETED"
			Case 13:
				GameModeStr="NA{orange}:ZION SPEECH COMPLETED"
			Case 14:
				GameModeStr="NA{blue}:PERSEPHONES REVENGE COMPLETED"
			Case 15:
				GameModeStr="NA{green}:MATRIX DESTROYED"
		End Select


	End If ' endif balls remaining
	Scorbit.SetGameMode(GameModeStr)

End Sub 






' END ----------

Sub Scorbit_LOGUpload(state)	' Callback during the log creation process.  0=Creating Log, 1=Uploading Log, 2=Done 
	Select Case state 
		case 0:
			debug4 "CREATING LOG"
		case 1:
			debug4 "Uploading LOG"
		case 2:
			debug4 "LOG Complete"
	End Select 
End Sub 
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' TABLE CUSTOMIZATION END HERE - NO NEED TO EDIT BELOW THIS LINE


dim Scorbit : Set Scorbit = New ScorbitIF
' Workaround - Call get a reference to Member Function
Sub tmrScorbit_Timer()								' Timer to send heartbeat 
	Scorbit.DoTimer(tmrScorbit.UserValue)
	tmrScorbit.UserValue=tmrScorbit.UserValue+1
	if tmrScorbit.UserValue>5 then tmrScorbit.UserValue=0
End Sub 
Function ScorbitIF_Callback()
	Scorbit.Callback()
End Function 
Class ScorbitIF

	Public bSessionActive
	Public bNeedsPairing
	Private bUploadLog
	Private bActive
	Private LOGFILE(10000000)
	Private LogIdx

	Private bProduction

	Private TypeLib
	Private MyMac
	Private Serial
	Private MyUUID
	Private TableVersion

	Private SessionUUID
	Private SessionSeq
	Private SessionTimeStart
	Private bRunAsynch
	Private bWaitResp
	Private GameMode
	Private GameModeOrig		' Non escaped version for log
	Private VenueMachineID
	Private CachedPlayerNames(4)
	Private SaveCurrentPlayer

	Public bEnabled
	Private sToken
	Private machineID
	Private dirQRCode
	Private opdbID
	Private wsh

	Private objXmlHttpMain
	Private objXmlHttpMainAsync
	Private fso
	Private Domain

	Public Sub Class_Initialize()
		bActive="false"
		bSessionActive=False
		bEnabled=False 
	End Sub 

	Property Let UploadLog(bValue)
		bUploadLog = bValue
	End Property

	Sub DoTimer(bInterval)	' 2 second interval
		dim holdScores(4)
		dim i
		if bInterval=0 then 
			SendHeartbeat()
		elseif bRunAsynch And bSessionActive = True then ' Game in play (Updated for TNA to resolve stutter in CoopMode)
			Scorbit.SendUpdate Score(1), Score(2), Score(3), Score(4), Balls, CurrentPlayer, PlayersPlayingGame
		End if 
	End Sub 

	Function GetName(PlayerNum)	' Return Parsed Players name  
		if PlayerNum<1 or PlayerNum>4 then 
			GetName=""
		else 
			GetName=CachedPlayerNames(PlayerNum-1)
		End if 
	End Function 

	Function DoInit(MyMachineID, Directory_PupQRCode, Version, opdb)
		dim Nad
		Dim EndPoint
		Dim resultStr 
		Dim UUIDParts 
		Dim UUIDFile

		bProduction=1
'		bProduction=0
		SaveCurrentPlayer=0
		VenueMachineID=""
		bWaitResp=False 
		bRunAsynch=False 
		DoInit=False 
		opdbID=opdb
		dirQrCode=Directory_PupQRCode
		MachineID=MyMachineID
		TableVersion=version
		bNeedsPairing=False
		if bProduction then 
			domain = "api.scorbit.io"
		else 
			domain = "staging.scorbit.io"
			domain = "scorbit-api-staging.herokuapp.com"
		End if 
		Set fso = CreateObject("Scripting.FileSystemObject")
		dim objLocator:Set objLocator = CreateObject("WbemScripting.SWbemLocator")
		Dim objService:Set objService = objLocator.ConnectServer(".", "root\cimv2")
		Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
		Set objXmlHttpMainAsync = CreateObject("Microsoft.XMLHTTP")
		objXmlHttpMain.onreadystatechange = GetRef("ScorbitIF_Callback")
		Set wsh = CreateObject("WScript.Shell")

		' Get Mac for Serial Number 
		dim Nads: set Nads = objService.ExecQuery("Select * from Win32_NetworkAdapter where physicaladapter=true")
		for each Nad in Nads
			if not isnull(Nad.MACAddress) then
				if left(Nad.MACAddress, 6)<>"00090F" then ' Skip over forticlient MAC
debug4 "Using MAC Addresses:" & Nad.MACAddress & " From Adapter:" & Nad.description   
					MyMac=replace(Nad.MACAddress, ":", "")
					Exit For 
				End if 
			End if 
		Next
		Serial=eval("&H" & mid(MyMac, 5))
		if Serial<0 then Serial=eval("&H" & mid(MyMac, 6))		' Mac Address Overflow Special Case 
		if MyMachineID<>2108 then 			' GOTG did it wrong but MachineID should be added to serial number also
			Serial=Serial+MyMachineID
		End if 
'		Serial=123456
		debug4 "Serial:" & Serial

		' Get System UUID
		set Nads = objService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
		for each Nad in Nads
			debug4 "Using UUID:" & Nad.UUID   
			MyUUID=Nad.UUID
			Exit For 
		Next

		if MyUUID="" then 
			MsgBox "SCORBIT - Can get UUID, Disabling."
			Exit Function
		elseif MyUUID="03000200-0400-0500-0006-000700080009" or ScorbitAlternateUUID then
			If fso.FolderExists(UserDirectory) then 
				If fso.FileExists(UserDirectory & "ScorbitUUID.dat") then
					Set UUIDFile = fso.OpenTextFile(UserDirectory & "ScorbitUUID.dat",1)
					MyUUID = UUIDFile.ReadLine()
					UUIDFile.Close
					Set UUIDFile = Nothing
				Else 
					MyUUID=GUID()
					Set UUIDFile=fso.CreateTextFile(UserDirectory & "ScorbitUUID.dat",True)
					UUIDFile.WriteLine MyUUID
					UUIDFile.Close
					Set UUIDFile=Nothing
				End if
			End if 
		End if

		' Clean UUID
		UUIDParts=split(MyUUID, "-")
		MyUUID=LCASE(Hex(eval("&h" & UUIDParts(0))+MyMachineID) & UUIDParts(1) &  UUIDParts(2) &  UUIDParts(3) & UUIDParts(4))		 ' Add MachineID to UUID
		MyUUID=LPad(MyUUID, 32, "0")
'		MyUUID=Replace(MyUUID, "-",  "")
		debug4 "MyUUID:" & MyUUID 


		' Authenticate and get our token 
		if getStoken() then 
			bEnabled=True 
'			SendHeartbeat
			DoInit=True
		End if 
	End Function 

	Sub Callback()
		Dim ResponseStr
		Dim i 
		Dim Parts
		Dim Parts2
		Dim Parts3
		if bEnabled=False then Exit Sub 

		if bWaitResp and objXmlHttpMain.readystate=4 then 
'			debug4 "CALLBACK: " & objXmlHttpMain.Status & " " & objXmlHttpMain.readystate
			if objXmlHttpMain.Status=200 and objXmlHttpMain.readystate = 4 then 
				ResponseStr=objXmlHttpMain.responseText
				'debug3 "RESPONSE: " & ResponseStr

				' Parse Name 
				If bSessionActive = True Then
					if CachedPlayerNames(SaveCurrentPlayer-1)="" then  ' Player doesnt have a name
						if instr(1, ResponseStr, "cached_display_name") <> 0 Then	' There are names in the result
							Parts=Split(ResponseStr,",{")							' split it 
							if ubound(Parts)>=SaveCurrentPlayer-1 then 				' Make sure they are enough avail
								if instr(1, Parts(SaveCurrentPlayer-1), "cached_display_name")<>0 then 	' See if mine has a name 
									CachedPlayerNames(SaveCurrentPlayer-1)=GetJSONValue(Parts(SaveCurrentPlayer-1), "cached_display_name")		' Get my name
									CachedPlayerNames(SaveCurrentPlayer-1)=Replace(CachedPlayerNames(SaveCurrentPlayer-1), """", "")
									Scorbit_PlayerClaimed SaveCurrentPlayer, CachedPlayerNames(SaveCurrentPlayer-1)
	'								debug4 "Player Claim:" & SaveCurrentPlayer & " " & CachedPlayerNames(SaveCurrentPlayer-1)
								End if 
							End if
						End if 
					else												    ' Check for unclaim 
						if instr(1, ResponseStr, """player"":null")<>0 Then	' Someone doesnt have a name
							Parts=Split(ResponseStr,"[")						' split it 
	'debug4 "Parts:" & Parts(1)
							Parts2=Split(Parts(1),"}")							' split it 
							for i = 0 to Ubound(Parts2)
	'debug4 "Parts2:" & Parts2(i)
								if instr(1, Parts2(i), """player"":null")<>0 Then
									CachedPlayerNames(i)=""
								End if 
							Next 
						End if 
					End if
				End If

				'Check heartbeat
				HandleHeartbeatResp ResponseStr
			End if 
			bWaitResp=False
		End if 
	End Sub

	Public Sub StartSession()
		if bEnabled=False then Exit Sub 
		Debug4 "Scorbit Start Session" 
		CachedPlayerNames(0)=""
		CachedPlayerNames(1)=""
		CachedPlayerNames(2)=""
		CachedPlayerNames(3)=""
		bRunAsynch=True 
		bActive="true"
		bSessionActive=True
		SessionSeq=0
		SessionUUID=GUID()
		SessionTimeStart=GameTime
		LogIdx=0
		SendUpdate 0, 0, 0, 0, 1, 1, 1
	End Sub

	' Custom method for TNA to work around coop mode stuttering
	Public Sub ForceAsynch(enabled)
		if bEnabled=False then Exit Sub
		if bSessionActive=True then Exit Sub 'Sessions should always control asynch when active
		bRunAsynch=enabled
	End Sub

	Public Sub StopSession(P1Score, P2Score, P3Score, P4Score, NumberPlayers)
		StopSession2 P1Score, P2Score, P3Score, P4Score, NumberPlayers, False
	End Sub 

	Public Sub StopSession2(P1Score, P2Score, P3Score, P4Score, NumberPlayers, bCancel)
		Dim i
		dim objFile
		if bEnabled=False then Exit Sub 
		bRunAsynch=False 'Asynch might have been forced on in TNA to prevent coop mode stutter
		if bSessionActive=False then Exit Sub 
debug4 "Scorbit Stop Session" 

		bActive="false" 
		SendUpdate P1Score, P2Score, P3Score, P4Score, -1, -1, NumberPlayers
		bSessionActive=False
'		SendHeartbeat

		if bUploadLog and LogIdx<>0 and bCancel=False then 
			debug4 "Creating Scorbit Log: Size" & LogIdx
			Scorbit_LOGUpload(0)
			Set objFile = fso.CreateTextFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
			For i = 0 to LogIdx-1 
				objFile.Writeline LOGFILE(i)
			Next 
			objFile.Close
			LogIdx=0
			Scorbit_LOGUpload(1)
			pvPostFile "https://" & domain & "/api/session_log/", puplayer.getroot&"\" & cGameName & "\sGameLog.csv", False
			Scorbit_LOGUpload(2)
			on error resume next
			fso.DeleteFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
			on error goto 0
		End if 

	End Sub 

	Public Sub SetGameMode(GameModeStr)
		GameModeOrig=GameModeStr
		GameMode=GameModeStr
		GameMode=Replace(GameMode, ":", "%3a")
		GameMode=Replace(GameMode, ";", "%3b")
		GameMode=Replace(GameMode, " ", "%20")
		GameMode=Replace(GameMode, "{", "%7B")
		GameMode=Replace(GameMode, "}", "%7D")
	End sub 

	Public Sub SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
		SendUpdateAsynch P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bRunAsynch
	End Sub 

	Public Sub SendUpdateAsynch(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bAsynch)
		dim i
		Dim PostData
		Dim resultStr
		dim LogScores(4)

		if bUploadLog then 
			if NumberPlayers>=1 then LogScores(0)=P1Score
			if NumberPlayers>=2 then LogScores(1)=P2Score
			if NumberPlayers>=3 then LogScores(2)=P3Score
			if NumberPlayers>=4 then LogScores(3)=P4Score
			LOGFILE(LogIdx)=DateDiff("S", "1/1/1970", Now()) & "," & LogScores(0) & "," & LogScores(1) & "," & LogScores(2) & "," & LogScores(3) & ",,," &  CurrentPlayer & "," & CurrentBall & ",""" & GameModeOrig & """"
			LogIdx=LogIdx+1
		End if

		if bSessionActive=False then Exit Sub 
		if bEnabled=False then Exit Sub 
		if bWaitResp then exit sub ' Drop message until we get our next response 

'		msgbox "currentplayer: " & CurrentPlayer
		SaveCurrentPlayer=CurrentPlayer
		PostData = "session_uuid=" & SessionUUID & "&session_time=" & GameTime-SessionTimeStart+1 & _
					"&session_sequence=" & SessionSeq & "&active=" & bActive

		SessionSeq=SessionSeq+1
		if NumberPlayers > 0 then 
			for i = 0 to NumberPlayers-1
				PostData = PostData & "&current_p" & i+1 & "_score="
				if i <= NumberPlayers-1 then 
					if i = 0 then PostData = PostData & P1Score
					if i = 1 then PostData = PostData & P2Score
					if i = 2 then PostData = PostData & P3Score
					if i = 3 then PostData = PostData & P4Score
				else 
					PostData = PostData & "-1"
				End if 
			Next 

			PostData = PostData & "&current_ball=" & CurrentBall & "&current_player=" & CurrentPlayer
			if GameMode<>"" then PostData=PostData & "&game_modes=" & GameMode
		End if 
		resultStr = PostMsg("https://" & domain, "/api/entry/", PostData, bAsynch)
		'if resultStr<>"" then debug3 "SendUpdate Resp:" & resultStr    			'rtp12
	End Sub 

' PRIVATE BELOW 
	Private Function LPad(StringToPad, Length, CharacterToPad)
	  Dim x : x = 0
	  If Length > Len(StringToPad) Then x = Length - len(StringToPad)
	  LPad = String(x, CharacterToPad) & StringToPad
	End Function

	Private Function GUID()		
		Dim TypeLib
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		GUID = Mid(TypeLib.Guid, 2, 36)
	End Function

	Private Function GetJSONValue(JSONStr, key)
		dim i 
		Dim tmpStrs,tmpStrs2
		if Instr(1, JSONStr, key)<>0 then 
			tmpStrs=split(JSONStr,",")
			for i = 0 to ubound(tmpStrs)
				if instr(1, tmpStrs(i), key)<>0 then 
					tmpStrs2=split(tmpStrs(i),":")
					GetJSONValue=tmpStrs2(1)
					exit for
				End if 
			Next 
		End if 
	End Function

	Private Sub SendHeartbeat()
		Dim resultStr
		if bEnabled=False then Exit Sub 
		resultStr = GetMsgHdr("https://" & domain, "/api/heartbeat/", "Authorization", "SToken " & sToken)
		
		'Customized for TNA
		If bRunAsynch = False Then 
			debug4 "Heartbeat Resp:" & resultStr
			HandleHeartbeatResp ResultStr
		End If
	End Sub 

	'TNA custom method
	Private Sub HandleHeartbeatResp(resultStr)
		dim TmpStr
		Dim Command
		Dim rc
		'Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
		Dim QRFile:QRFile=puplayer.getroot & cGameName & "\" & dirQrCode
debug4 "QRFile: " &QRFile
		If VenueMachineID="" then
			If resultStr<>"" And Not InStr(resultStr, """machine_id"":" & machineID)=0 Then 'We Paired
				bNeedsPairing=False
				debug4 "Scorbit: Paired"
				Scorbit_Paired()
			ElseIf resultStr<>"" And Not InStr(resultStr, """unpaired"":true")=0 Then 'We Did not Pair
				debug4 "Scorbit: NOT Paired"
				bNeedsPairing=True
			Else
				' Error (or not a heartbeat); do nothing
			End If

			TmpStr=GetJSONValue(resultStr, "venuemachine_id")
			if TmpStr<>"" then 
				VenueMachineID=TmpStr
'debug4 "VenueMachineID=" & VenueMachineID			
				'Command = """" & puplayer.getroot&"\" & cGameName & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
				Command = """" & puplayer.getroot & cGameName & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
				rc = wsh.Run(Command, 0, False)
			End if 
		End if
	End Sub

	Private Function getStoken()
		Dim result
		Dim results
'		dim wsh
		Dim tmpUUID:tmpUUID="adc12b19a3504453a7414e722f58736b"
		Dim tmpVendor:tmpVendor="vscorbitron"
		Dim tmpSerial:tmpSerial="999990104"
		'Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
		Dim QRFile:QRFile=puplayer.getroot & cGameName & "\" & dirQrCode
		'Dim sTokenFile:sTokenFile=puplayer.getroot&"\" & cGameName & "\sToken.dat"
		Dim sTokenFile:sTokenFile=puplayer.getroot & cGameName & "\sToken.dat"

		' Set everything up
		tmpUUID=MyUUID
		tmpVendor="vpin"
		tmpSerial=Serial
		
		on error resume next
		fso.DeleteFile(sTokenFile)
		On error goto 0 

		' get sToken and generate QRCode
'		Set wsh = CreateObject("WScript.Shell")
		Dim waitOnReturn: waitOnReturn = True
		Dim windowStyle: windowStyle = 0
		Dim Command 
		Dim rc
		Dim objFileToRead

		'Command = """" & puplayer.getroot&"\" & cGameName & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
		Command = """" & puplayer.getroot & cGameName & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
debug4 "RUNNING Command:" & Command
		rc = wsh.Run(Command, windowStyle, waitOnReturn)
debug4 "Return:" & rc
		if FileExists(puplayer.getroot&"\" & cGameName & "\sToken.dat") and rc=0 then
			Set objFileToRead = fso.OpenTextFile(puplayer.getroot&"\" & cGameName & "\sToken.dat",1)
			result = objFileToRead.ReadLine()
			objFileToRead.Close
			Set objFileToRead = Nothing

			if Instr(1, result, "Invalid timestamp")<> 0 then 
				MsgBox "Scorbit Timestamp Error: Please make sure the time on your system is exact"
				getStoken=False
			elseif Instr(1, result, ":")<>0 then 
				results=split(result, ":")
				sToken=results(1)
				sToken=mid(sToken, 3, len(sToken)-4)
debug4 "Got TOKEN:" & sToken
				getStoken=True
			Else 
debug4 "ERROR:" & result
				getStoken=False
			End if 
		else 
debug4 "ERROR No File:" & rc
		End if 

	End Function 

	private Function FileExists(FilePath)
		If fso.FileExists(FilePath) Then
			FileExists=CBool(1)
		Else
			FileExists=CBool(0)
		End If
	End Function

	Private Function GetMsg(URLBase, endpoint)
		GetMsg = GetMsgHdr(URLBase, endpoint, "", "")
	End Function

	Private Function GetMsgHdr(URLBase, endpoint, Hdr1, Hdr1Val)
		Dim Url
		Url = URLBase + endpoint & "?session_active=" & bActive
'debug4 "Url:" & Url  & "  Async=" & bRunAsynch
		objXmlHttpMain.open "GET", Url, bRunAsynch
'		objXmlHttpMain.setRequestHeader "Content-Type", "text/xml"
		objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
		if Hdr1<> "" then objXmlHttpMain.setRequestHeader Hdr1, Hdr1Val

'		on error resume next
			err.clear
			objXmlHttpMain.send ""
			if err.number=-2147012867 then 
				MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
				bEnabled=False
			elseif err.number <> 0 then 
				debug3 "Server error: (" & err.number & ") " & Err.Description
			End if 
			if bRunAsynch=False then 
debug4 "Status: " & objXmlHttpMain.status
				If objXmlHttpMain.status = 200 Then
					GetMsgHdr = objXmlHttpMain.responseText
				Else 
					GetMsgHdr=""
				End if 
			Else 
				bWaitResp=True
				GetMsgHdr=""
			End if 
'		On error goto 0

	End Function

	Private Function PostMsg(URLBase, endpoint, PostData, bAsynch)
		Dim Url

		Url = URLBase + endpoint
'debug4 "PostMSg:" & Url & " " & PostData			'rtp12

		objXmlHttpMain.open "POST",Url, bAsynch
		objXmlHttpMain.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXmlHttpMain.setRequestHeader "Content-Length", Len(PostData)
		objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
		objXmlHttpMain.setRequestHeader "Authorization", "SToken " & sToken
		if bAsynch then bWaitResp=True 

		on error resume next
			objXmlHttpMain.send PostData
			if err.number=-2147012867 then 
				MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
				bEnabled=False
			elseif err.number <> 0 then 
				'debug3 "Multiplayer Server error (" & err.number & ") " & Err.Description
			End if 
			If objXmlHttpMain.status = 200 Then
				PostMsg = objXmlHttpMain.responseText
			else 
				PostMsg="ERROR: " & objXmlHttpMain.status & " >" & objXmlHttpMain.responseText & "<"
			End if 
		On error goto 0
	End Function

	Private Function pvPostFile(sUrl, sFileName, bAsync)
'debug4 "Posting File " & sUrl & " " & sFileName & " " & bAsync & " File:" & Mid(sFileName, InStrRev(sFileName, "\") + 1)
		Dim STR_BOUNDARY:STR_BOUNDARY  = GUID()
		Dim nFile  
		Dim baBuffer()
		Dim sPostData
		Dim Response

		'--- read file
		Set nFile = fso.GetFile(sFileName)
		With nFile.OpenAsTextStream()
			sPostData = .Read(nFile.Size)
			.Close
		End With


		'--- prepare body
		sPostData = "--" & STR_BOUNDARY & vbCrLf & _
			"Content-Disposition: form-data; name=""uuid""" & vbCrLf & vbCrLf & _
			SessionUUID & vbcrlf & _
			"--" & STR_BOUNDARY & vbCrLf & _
			"Content-Disposition: form-data; name=""log_file""; filename=""" & SessionUUID & ".csv""" & vbCrLf & _
			"Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
			sPostData & vbCrLf & _
			"--" & STR_BOUNDARY & "--"


		'--- post
		With objXmlHttpMain
			.Open "POST", sUrl, bAsync
			.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
			.SetRequestHeader "Authorization", "SToken " & sToken
			.Send sPostData ' pvToByteArray(sPostData)
			If Not bAsync Then
				Response= .ResponseText
				pvPostFile = Response
debug4 "Upload Response: " & Response
			End If
		End With

	End Function

	Private Function pvToByteArray(sText)
		pvToByteArray = StrConv(sText, 128)		' vbFromUnicode
	End Function

End Class 
'  END SCORBIT 
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub DelayQRClaim_Timer()
	if bOnTheFirstBall AND bBallInPlungerLane then ScorbitClaimQR(True)
	DelayQRClaim.Enabled=False
End Sub


Dim ColorLUT : ColorLUT = 0
Dim VRANIMATION

' Base options
Const Opt_LUT = 0
Const Opt_Freeplay = 1
Const Opt_MaxPlayers = 2
Const Opt_HIghQFlex = 3
'Pup
Const Opt_UsePup = 4
Const Opt_UseScorebit = 5
' VR options
Const Opt_VRRoomChoice = 6
Const Opt_VRplayAnimations = 7
' Informations
Const Opt_Info_1 = 8
Const Opt_Info_2 = 9

Const NOptions = 10

Const FlexDMD_RenderMode_DMD_GRAY_2 = 0
Const FlexDMD_RenderMode_DMD_GRAY_4 = 1
Const FlexDMD_RenderMode_DMD_RGB = 2
Const FlexDMD_RenderMode_SEG_2x16Alpha = 3
Const FlexDMD_RenderMode_SEG_2x20Alpha = 4
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9
Const FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10
Const FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11
Const FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12
Const FlexDMD_RenderMode_SEG_4x7Num10 = 13
Const FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14
Const FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15
Const FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const FlexDMD_Align_TopLeft = 0
Const FlexDMD_Align_Top = 1
Const FlexDMD_Align_TopRight = 2
Const FlexDMD_Align_Left = 3
Const FlexDMD_Align_Center = 4
Const FlexDMD_Align_Right = 5
Const FlexDMD_Align_BottomLeft = 6
Const FlexDMD_Align_Bottom = 7
Const FlexDMD_Align_BottomRight = 8

Const FlexDMD_Scaling_Fit = 0
Const FlexDMD_Scaling_Fill = 1
Const FlexDMD_Scaling_FillX = 2
Const FlexDMD_Scaling_FillY = 3
Const FlexDMD_Scaling_Stretch = 4
Const FlexDMD_Scaling_StretchX = 5
Const FlexDMD_Scaling_StretchY = 6
Const FlexDMD_Scaling_None = 7

Const FlexDMD_Interpolation_Linear = 0
Const FlexDMD_Interpolation_ElasticIn = 1
Const FlexDMD_Interpolation_ElasticOut = 2
Const FlexDMD_Interpolation_ElasticInOut = 3
Const FlexDMD_Interpolation_QuadIn = 4
Const FlexDMD_Interpolation_QuadOut = 5
Const FlexDMD_Interpolation_QuadInOut = 6
Const FlexDMD_Interpolation_CubeIn = 7
Const FlexDMD_Interpolation_CubeOut = 8
Const FlexDMD_Interpolation_CubeInOut = 9
Const FlexDMD_Interpolation_QuartIn = 10
Const FlexDMD_Interpolation_QuartOut = 11
Const FlexDMD_Interpolation_QuartInOut = 12
Const FlexDMD_Interpolation_QuintIn = 13
Const FlexDMD_Interpolation_QuintOut = 14
Const FlexDMD_Interpolation_QuintInOut = 15
Const FlexDMD_Interpolation_SineIn = 16
Const FlexDMD_Interpolation_SineOut = 17
Const FlexDMD_Interpolation_SineInOut = 18
Const FlexDMD_Interpolation_BounceIn = 19
Const FlexDMD_Interpolation_BounceOut = 20
Const FlexDMD_Interpolation_BounceInOut = 21
Const FlexDMD_Interpolation_CircIn = 22
Const FlexDMD_Interpolation_CircOut = 23
Const FlexDMD_Interpolation_CircInOut = 24
Const FlexDMD_Interpolation_ExpoIn = 25
Const FlexDMD_Interpolation_ExpoOut = 26
Const FlexDMD_Interpolation_ExpoInOut = 27
Const FlexDMD_Interpolation_BackIn = 28
Const FlexDMD_Interpolation_BackOut = 29
Const FlexDMD_Interpolation_BackInOut = 30

Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo



Sub Options_Open
	bOptionsMagna = False
	On Error Resume Next
	Set OptionDMD = CreateObject("FlexDMD.FlexDMD")
	On Error Goto 0
	If OptionDMD is Nothing Then
		Debug.Print "FlexDMD is not installed"
		Debug.Print "Option UI can not be opened"
		MsgBox "You need to install FlexDMD to access table options"
		Exit Sub
	End If
	If Table1.ShowDT Then OptionDMDFlasher.RotX = -(Table1.Inclination + Table1.Layback)
	bInOptions = True
	OptPos = 0
	OptSelected = False
	OptionDMD.Show = False
	OptionDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_2
	OptionDMD.Width = 128
	OptionDMD.Height = 32
	OptionDMD.Clear = True
	OptionDMD.Run = True
	Dim a, scene, font
	Set scene = OptionDMD.NewGroup("Scene")
	Set OptFontHi = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
	Set OptFontLo = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
	Set OptSel = OptionDMD.NewGroup("Sel")
	Set a = OptionDMD.NewLabel(">", OptFontLo, ">>>")
	a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
	OptSel.AddActor a
	Set a = OptionDMD.NewLabel(">", OptFontLo, "<<<")
	a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
	OptSel.AddActor a
	scene.AddActor OptSel
	OptSel.SetBounds 0, 0, 128, 32
	OptSel.Visible = False
	
	Set a = OptionDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
	a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
	scene.AddActor a
	Set a = OptionDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
	a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
	scene.AddActor a
	Set OptN = OptionDMD.NewLabel("Pos", OptFontLo, "LINE 1")
	Set OptTop = OptionDMD.NewLabel("Top", OptFontLo, "LINE 1")
	Set OptBot = OptionDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
	scene.AddActor OptN
	scene.AddActor OptTop
	scene.AddActor OptBot
	Options_OnOptChg
	OptionDMD.LockRenderThread
	OptionDMD.Stage.AddActor scene
	OptionDMD.UnlockRenderThread
	OptionDMDFlasher.Visible = True
	DisableStaticPrerendering = True
	If NOT usePUP and (RenderingMode = 2 or Table1.ShowFSS or VRTest) Then 'for some reason the morror part will take presence on this stage
		VR_Backglass.visible = True
		VR_PupBackglass.visible = False
		VRPupTopper.visible=False
	End If
End Sub

Sub Options_UpdateDMD
	If OptionDMD is Nothing Then Exit Sub
	Dim DMDp: DMDp = OptionDMD.DmdPixels
	If Not IsEmpty(DMDp) Then
		OptionDMDFlasher.DMDWidth = OptionDMD.Width
		OptionDMDFlasher.DMDHeight = OptionDMD.Height
		OptionDMDFlasher.DMDPixels = DMDp
	else 

	
	End If
End Sub

Sub Options_Close
	bInOptions = False
	OptionDMDFlasher.Visible = False
	If OptionDMD is Nothing Then Exit Sub
	OptionDMD.Run = False
	Set OptionDMD = Nothing
	DisableStaticPrerendering = False 
	
End Sub

Function Options_OnOffText(opt)
	If opt Then
		Options_OnOffText = "ON"
	Else
		Options_OnOffText = "OFF"
	End If
End Function

Sub Options_OnOptChg
	If OptionDMD is Nothing Then Exit Sub
	OptionDMD.LockRenderThread
	OptN.Text = (OptPos+1) & "/" & (NOptions)
	If OptSelected Then
		OptTop.Font = OptFontLo
		OptBot.Font = OptFontHi
		OptSel.Visible = True
	Else
		OptTop.Font = OptFontHi
		OptBot.Font = OptFontLo
		OptSel.Visible = False
	End If
	If OptPos = Opt_LUT Then
		OptTop.Text = "COLOR SATURATION"
		if ColorLUT = 1 Then OptBot.text = "DISABLED"
		if ColorLUT = 2 Then OptBot.text = "LUTIMAGE 1"
		if ColorLUT = 3 Then OptBot.text = "LUTIMAGE 2"
		if ColorLUT = 4 Then OptBot.text = "LUTIMAGE 3"
		if ColorLUT = 5 Then OptBot.text = "LUTIMAGE 4"
		if ColorLUT = 6 Then OptBot.text = "LUTIMAGE 5"
		if ColorLUT = 7 Then OptBot.text = "LUTIMAGE 6"
		if ColorLUT = 8 Then OptBot.text = "LUTIMAGE 7"
		if ColorLUT = 9 Then OptBot.text = "LUTIMAGE 8"
		if ColorLUT = 10 Then OptBot.text = "LUTIMAGE 9"
		SaveValue cGameName, "LUT", ColorLUT
	ElseIf OptPos = Opt_Freeplay Then
		OptTop.Text = "FREEPLAY (RESTART REQUIRED)"
		OptBot.Text = Options_OnOffText(bFreePlay)
		SaveValue cGameName, "FREEPLAY", bFreePlay
	ElseIf OptPos = Opt_MaxPlayers Then
		OptTop.Text = "MAX PLAYERS"
		if MaxPlayers = 1 Then OptBot.text = "1"
		if MaxPlayers = 2 Then OptBot.text = "2"
		if MaxPlayers = 3 Then OptBot.text = "3"
		if MaxPlayers = 4 Then OptBot.text = "4"
		SaveValue cGameName, "MAXPLAYERS", MaxPlayers
	ElseIf OptPos = Opt_HIghQFlex Then
		OptTop.Text = "HIGHRES DMD(RESTART REQ)"
		OptBot.Text = Options_OnOffText(FlexDMDHighQuality)
		SaveValue cGameName, "HIGHRESDMD", FlexDMDHighQuality
	ElseIf OptPos = Opt_UsePup Then
		OptTop.Text = "USE PUP (RESTART REQUIRED)"
		OptBot.Text = Options_OnOffText(Usepup)
		SaveValue cGameName, "USEPUP", Usepup		
	ElseIf OptPos = Opt_UseScorebit Then
		OptTop.Text = "SCOREBIT (RESTART REQUIRED)"
		OptBot.Text = Options_OnOffText(ScorbitActive)
		SaveValue cGameName, "SCORBITACTIVE", ScorbitActive
	ElseIf OptPos = Opt_VRRoomChoice Then
		OptTop.Text = "VR ROOM (VR AND FSS ONLY)"
		If VRRoomChoice = 1 Then OptBot.Text = "DOJO ROOM"
		If VRRoomChoice = 2 Then OptBot.Text = "WHITE ROOM"
		If VRRoomChoice = 3 Then OptBot.Text = "MINIMAL ROOM"
		If VRRoomChoice = 4 Then OptBot.Text = "MR GREEN RGB 0,153,51"
		If VRRoomChoice = 5 Then OptBot.Text = "MR BLUE RGB 0,71,187"
		SaveValue cGameName, "VRROOM", VRRoomChoice
	ElseIf OptPos = Opt_VRplayAnimations Then
		OptTop.Text = "SHOW VR ANIMATION"
		OptBot.Text = Options_OnOffText(VRANIMATION)
		SaveValue cGameName, "VRANIMATION", VRANIMATION
	ElseIf OptPos = Opt_Info_1 Then
		OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
		OptBot.Text = "The Matrix (Original 2024) " & myVersion
	ElseIf OptPos = Opt_Info_2 Then
		OptTop.Text = "RENDER MODE"
		If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
		If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
		If RenderingMode = 2 Then OptBot.Text = "VR"
	End If
	OptTop.Pack
	OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight
	OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
	OptionDMD.UnlockRenderThread
	UpdateMods
End Sub

Sub Options_Toggle(amount)
	If OptionDMD is Nothing Then Exit Sub
	If OptPos = Opt_LUT Then
		ColorLUT = ColorLUT + amount * 1
		If ColorLUT < 1 Then ColorLUT = 10
		If ColorLUT > 10 Then ColorLUT = 1
	ElseIf OptPos = Opt_VRRoomChoice Then
		VRRoomChoice = VRRoomChoice + amount
		If VRRoomChoice < 1 Then VRRoomChoice = 5
		If VRRoomChoice > 5 Then VRRoomChoice = 1
	ElseIf OptPos = 	Opt_MaxPlayers Then
		MaxPlayers = MaxPlayers + amount
		If MaxPlayers < 1 Then MaxPlayers = 4
		If MaxPlayers > 4 Then MaxPlayers = 1
	Elseif OptPos = Opt_UsePup  then
		USEPUP = Not USEPUP
	Elseif OptPos = Opt_freeplay  then
		bFreePlay = Not bFreePlay
	Elseif OptPos = Opt_HIghQFlex  then
		FlexDMDHighQuality = Not FlexDMDHighQuality
	ElseIf OptPos =Opt_UseScorebit Then
		ScorbitActive = ScorbitActive + amount
		If ScorbitActive < 0 Then ScorbitActive = 1
		If ScorbitActive > 1 Then ScorbitActive = 0
	ElseIf OptPos =Opt_VRplayAnimations Then
		VRANIMATION = VRANIMATION + amount
		If VRANIMATION < 0 Then VRANIMATION = 1
		If VRANIMATION > 1 Then VRANIMATION = 0
	End If


End Sub

Sub Options_KeyDown(ByVal keycode)
	
	If OptSelected Then
		If keycode = LeftMagnaSave Then ' Exit / Cancel
			OptSelected = False
		ElseIf keycode = RightMagnaSave Then ' Enter / Select
			OptSelected = False
		ElseIf keycode = LeftFlipperKey Then ' Next / +
			Options_Toggle	-1
		ElseIf keycode = RightFlipperKey Then ' Prev / -
			Options_Toggle	1
		End If
	Else
		If keycode = LeftMagnaSave Then ' Exit / Cancel
			Options_Close
		ElseIf keycode = RightMagnaSave Then ' Enter / Select
			If OptPos < Opt_Info_1 Then OptSelected = True
			if OptPos = Opt_UseScorebit and Not USEPUP Then OptSelected = false : 
			if OptPos = Opt_VRplayAnimations and Not (RenderingMode = 2 or Table1.ShowFSS or VRTest) Then OptSelected = false
			if OptPos = Opt_VRRoomChoice and Not (RenderingMode = 2 or Table1.ShowFSS or VRTest) Then OptSelected = false
		ElseIf keycode = LeftFlipperKey Then ' Next / +
			OptPos = OptPos - 1
			If OptPos < 0 Then OptPos = NOptions - 1
		ElseIf keycode = RightFlipperKey Then ' Prev / -
			OptPos = OptPos + 1
			If OptPos >= NOPtions Then OptPos = 0
		End If
	End If

	'If bInOptions Then
'		Options_KeyDown keycode
'		Exit Sub
'	End If
	Options_OnOptChg
End Sub

Sub Options_Load
	Dim x
    x = LoadValue(cGameName, "LUT") : If x <> "" Then ColorLUT = CInt(x) Else ColorLUT = 1
	x = LoadValue(cGameName, "FREEPLAY") : If x <> "" Then bfreeplay = Cbool(x) Else bfreeplay = True
	x = LoadValue(cGameName, "MAXPLAYERS") : If x <> "" Then MaxPlayers = Cint(x) Else MaxPlayers = 4
	x = LoadValue(cGameName, "HIGHRESDMD") : If x <> "" Then FlexDMDHighQuality = Cbool(x) Else FlexDMDHighQuality = True
	x = LoadValue(cGameName, "VRROOM") : If x <> "" Then VRRoomChoice = CInt(x) Else VRRoomChoice = 1
	x = LoadValue(cGameName, "USEPUP") : If x <> "" Then USEPUP = Cbool(x) Else USEPUP = True
	x = LoadValue(cGameName, "SCORBITACTIVE") : If x <> "" Then ScorbitActive = CInt(x) Else ScorbitActive = 0
	x = LoadValue(cGameName, "VRANIMATION") : If x <> "" Then VRANIMATION = CInt(x) Else VRANIMATION = 1
	if USEPUP = False then ScorbitActive = 0 'disable scorebit when Pup is not used
	UpdateMods
End Sub

Sub UpdateMods
	Dim BL, LM, x, y, c, enabled

	'*********************
	'Color LUT
	'*********************

	if ColorLUT = 1 Then Table1.ColorGradeImage = ""
	if ColorLUT = 2 Then Table1.ColorGradeImage = "LUT0"
	if ColorLUT = 3 Then Table1.ColorGradeImage = "LUT1"
	if ColorLUT = 4 Then Table1.ColorGradeImage = "LUT2"
	if ColorLUT = 5 Then Table1.ColorGradeImage = "LUT3"
	if ColorLUT = 6 Then Table1.ColorGradeImage = "LUT4"
	if ColorLUT = 7 Then Table1.ColorGradeImage = "LUT5"
	if ColorLUT = 8 Then Table1.ColorGradeImage = "LUT6"
	if ColorLUT = 9 Then Table1.ColorGradeImage = "LUT7"
	if ColorLUT = 10 Then Table1.ColorGradeImage = "LUT8"
	if ColorLUT = 11 Then Table1.ColorGradeImage = "LUT9"

	If RenderingMode = 2 or Table1.ShowFSS or VRTest Then
		UseFlexDMD = False
		VR_Dogo047AgentSmith.visible = False
		VR_Dogo048Neo.visible = False
		VR_Dogo047AgentSmith.StopAnim()
		VR_Dogo048Neo.stopanim()
		If VRRoomChoice = 1 Then
			For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRWhiteroom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRDojo : VR_Obj.Visible = true : Next
			if VRANIMATION = 1  Then
				VR_Dogo047AgentSmith.visible = true
				VR_Dogo048Neo.visible = true
				VR_Dogo047AgentSmith.PlayAnimEndless (0.10)
				VR_Dogo048Neo.PlayAnimEndless (0.10)
			Else
				VR_Dogo047AgentSmith.visible = False
				VR_Dogo048Neo.visible = False
				VR_Dogo047AgentSmith.StopAnim()
				VR_Dogo048Neo.stopanim()
			End If
		ElseIf VRRoomChoice = 2 Then
			For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRWhiteroom : VR_Obj.Visible = true : Next
			For Each VR_Obj in VRDojo : VR_Obj.Visible = false : Next
			VRSphere.color = RGB(255,255,255)
		ElseIf VRRoomChoice = 4 Then
			For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRWhiteroom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRDojo : VR_Obj.Visible = false : Next
			VRSphere.visible = true
			VRSphere.color = RGB(0,153,51)
		ElseIf VRRoomChoice = 5 Then
			For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRWhiteroom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRDojo : VR_Obj.Visible = false : Next
			VRSphere.visible = true
			VRSphere.color = RGB(0,71,187)
		else
			For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = true : Next
			For Each VR_Obj in VRWhiteroom : VR_Obj.Visible = false : Next
			For Each VR_Obj in VRDojo : VR_Obj.Visible = false : Next
		End If
		For Each VR_Obj in VRCab : VR_Obj.Visible = True : Next

	Else
		For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = False : Next
		For Each VR_Obj in VRWhiteroom : VR_Obj.Visible = false : Next
		For Each VR_Obj in VRDojo : VR_Obj.Visible = false : Next
	end If

	
End Sub

Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then     
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function

Sub VRPupTopper_Timer()
		VRPupTopper.VideoCapUpdate="PUPSCREEN0"
end Sub	

'***************************************************************************
' VR Plunger Animation Code 
'***************************************************************************

Dim VRPlungerYstart: VRPlungerystart = PinCab_Shooter.Y 
Dim PlungerYstart : Plungerystart= Plunger.Y


Sub TimerVRPlunger_Timer
	if PinCab_Shooter.Y - VRPlungerYstart < 100 then 
	  PinCab_Shooter.Y = PinCab_Shooter.y +6
     End If
End Sub

Sub TimerVRPlunger2_Timer
	PinCab_Shooter.Y = (VRPlungerYstart + PlungerYstart - Plunger.y) + (5 * Plunger.Position) 
end sub

'***************************************************************************
' VR Character Animation Code
'***************************************************************************

