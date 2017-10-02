exports.speak = function(tts, SARAH){
    
	if (tts != false)
   {
	var itts=tts.replace("[Charlie]","").replace("\"",' ').replace(/\n/g, ' ')
  itts=itts.replace('"',' ')
	console.log(itts)
  var long=(tts.toString().length)*105 
  
	var fs = require('fs');
	
	var stri='Set objWMIService = GetObject("winmgmts:\\\\" & "." & "'+'\\root\\cimv2")\r\n'+'Set colItems = objWMIService.ExecQuery ("Select * From Win32_DisplayConfiguration")\r\n'+'For Each objItem in colItems\r\n'+'SarahID = "Sarah"\n'+'ScriptComplete = False\n'+'agFl = "'+__dirname+"\\interface"+'.acs"\n'+'Set AgentControl = WScript.CreateObject("Agent.Control.2", "AgentControl_")\n'+'AgentControl.Connected = True\n'+'AgentControl.Characters.Load SarahID, agFl\n'+	'Set Sarah = AgentControl.Characters(SarahID)\n'+	'Sarah.MoveTo '+'(objItem.PelsWidth/2)-310'+','+ '(objItem.PelsHeight/2)-360'+'\r\n'+'Sarah.Balloon.Style = 1\r\n'+'Sarah.Show\n'+'Sarah.Speak ("'+itts+'")\n'+'Set EndReq = Sarah.Speak("\\mrk=999999999\\")\n'+'Do While Sarah.Visible=True\n'+'Loop\n'+'next\n';;
	
	fs.writeFile(__dirname+"/interface.vbs", stri, 'ascii')
   
   setTimeout((function() {
 

var exec = require('child_process').exec;
  var process = __dirname+'/interface.vbs';
  var child = exec(process,
  	function (error, stdout, stderr) {
        
		
		});
}), 100); 
   }
   
  
  return true ;   
  
};
