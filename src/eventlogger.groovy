import java.io.*
import java .text.*
import groovy.time.*

def args = ["evtSrc":"system"]
pout = pout == null ? System.out : pout
perr = perr == null ? System.err : perr

evtSrc = evtSrc == null ? "system" : evtSrc 
args["evtSrc"]=evtSrc

Logger.StdErr = perr
Logger.StdOut = pout 

def propertyMissing(String name){} 

def arguments(defaults){
	def parameters = defaults
	args.each{ argument -> 
		if ( argument.contains("=") )
		{
			(key,value)= argument.split("=")
			parameters[key] = value.replaceAll('"','') 
		}
	}
	return parameters 
}



class EventLogQuery{

	def evtQueryCmd
	def startTime
	def endTime 
	def evtSource
	public EventLogQuery(src){


		//# 2012-08-25T16:44:15.563
		//# yyyy-MM-ddTHH:MM:ss.SSS
		//evtQueryCmd= "wevtutil qe #evtSrc /f:text \"/q:*[System[TimeCreated[@SystemTime>='#t1' and @SystemTime<='#t2']]]\""
		evtQueryCmd='wevtutil qe #evtSrc /rd:true /f:text /q:"*[System[TimeCreated[timediff(@SystemTime) <= 65000]]]'
		def t2 = new Date()
		def t1
		use (TimeCategory){ 
			t1 = 1.minutes.ago 
		}

		startTime = t1.format("yyyy-MM-dd#HH:mm:ss.SSS").replace("#","T")
		endTime = t2.format("yyyy-MM-dd#HH:mm:ss.SSS").replace("#","T")
		
		evtSource = src

	}

	def BufferedReader execute() {
		def cmd = "wevtutil qe System /f:text"
		evtQueryCmd = evtQueryCmd.replace("#t1",startTime)	
		evtQueryCmd = evtQueryCmd.replace("#t2",endTime)	
		evtQueryCmd = evtQueryCmd.replace("#evtSrc",evtSource)	


		String[] cmdArgs = evtQueryCmd.split(" ");
		ProcessBuilder builder = new ProcessBuilder(cmdArgs)
		builder.redirectErrorStream(true)
		Process p = builder.start()
		p.getOutputStream().flush()
		p.getOutputStream().close()
		return  new BufferedReader (new InputStreamReader ( p.getInputStream() ) ) 
	}
}
class EventBuilder{
	def buffer = "" 
	def sep = ";"
	def evtDate  
	def addEventInfoLine(line){
		if (line.contains("Date"))
		{
			
			def dtStr = line.split("Date:")[1] .trim() 
			evtDate = new Date().parse("yyyy-MM-dd'T'HH:mm:ss.SSS",dtStr)
		}	
		buffer = buffer + line.trim() + sep
	}

	def addDescriptionLine(line){
		buffer = buffer + line + "\n"
	}

	public String ToString(){
		return buffer
	}
}

class EventParser{
	def parseStates = ["EntryStart","EntryDescription","EntryEnd"]
	def stateId = -1 
	def buffer = ""
	def sep = ";"
	def eb 

	public EventParser(){
		eb = new EventBuilder()
	}

	def parse(line){
		if (line.startsWith("Event")){
			this.stateId = this.stateId + 1 			
			
			if (this.stateId >= 2){
				Logger.log(eb.ToString(),eb.evtDate)
				this.eb = new EventBuilder() 
				this.stateId = 0
			}
		}
		if (line.contains("Description")){
			this.stateId = this.stateId + 1
		}
		process(line)
	}

	def process(line){
		if (this.stateId ==0)
		{
			eb.addEventInfoLine(line) 
		}
		if(this.stateId == 1 )
		{

			eb.addDescriptionLine(line) 	
		}
	}
}

class Logger { 

	public static PrintStream StdErr  = System.err
	public static PrintStream StdOut  = System.out 
	
	public static void  log(line)
	{
		log(line, new Date().format("yyyy-MM-dd HH:mm:ss" )  ) 	
	}

	public static void  log(line,timestamp)
	{
		def sep = ";" 
		def dt = timestamp 

		def dtStr = dt.format("yyyy-MM-dd HH:mm:ss")


		def message =  "" + dtStr + sep +  line 

		Logger.StdOut << message 

	}
}




//def static fileHandles = [:]
//Set Argument Defaults
//parameters = arguments( ["evtSrc":"System"] ) 
parameters = arguments( args  ) 
def evtSrc = parameters["evtSrc"]
def elq = new EventLogQuery(evtSrc)

def ep = new EventParser() 
reader = elq.execute()

reader.eachLine(){ line -> 
	ep.parse(line)
}

Logger.StdErr << new Date() << " "
Logger.StdErr << "Used event source:" << evtSrc 
Logger.StdErr << "\n"




