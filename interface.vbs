Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
Set colItems = objWMIService.ExecQuery ("Select * From Win32_DisplayConfiguration")
For Each objItem in colItems
SarahID = "Sarah"
ScriptComplete = False
agFl = "C:\SARAH\plugins\interface\interface.acs"
Set AgentControl = WScript.CreateObject("Agent.Control.2", "AgentControl_")
AgentControl.Connected = True
AgentControl.Characters.Load SarahID, agFl
Set Sarah = AgentControl.Characters(SarahID)
Sarah.MoveTo (objItem.PelsWidth/2)-310,(objItem.PelsHeight/2)-360
Sarah.Balloon.Style = 1
Sarah.Show
Sarah.Speak ("Vous avez reçu un nouvel email, voici le titre de votre message, ; En direct d'AlloVoisins à Thiel sur acolin")
Set EndReq = Sarah.Speak("\mrk=999999999\")
Do While Sarah.Visible=True
Loop
next
