import java.io.*
import java .text.*
import groovy.time.*

def args = ["evtSrc":"system"]
pout = pout == null ? System.out : pout
perr = perr == null ? System.err : perr
bundleDir = bundleDir== null ? "" : ""
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
	public EventLogQuery(src,bundleDir){

		


		//# 2012-08-25T16:44:15.563
		//# yyyy-MM-ddTHH:MM:ss.SSS
		//cscript equery.vbs /fi "Datetime gt 07/05/2012,10:00:00AM" /fo csv
		// /nh //nologo /v /fo csv
		evtQueryCmd='cscript '+ bundleDir +'equery.vbs /fi "Datetime gt #t1" /nh //nologo /v /fo csv /L #evtSrc'
		def t2 = new Date()
		def t1
		use (TimeCategory){ 
			t1 = 1.minutes.ago 
		}

		startTime = t1.format("M/dd/yyyy,hh:mm:ssa").replace("#","T")
		endTime = t2.format("yyyy-MM-dd#hh:mm:ss.SSS").replace("#","T")

		
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


class Logger { 

	public static PrintStream StdErr  = System.err
	public static PrintStream StdOut  = System.out 
	
	public static void  log(line)
	{
		log(line, new Date().format("yyyy-MMM-dd HH:mm:ss zzz" )  ) 	
	}

	public static void  log(line,timestamp)
	{
		def sep = " " 
		def dt = timestamp 

		def dtStr = dt.format("yyyy-MMM-dd HH:mm:ss zzz")
		def message =  "" + dtStr + sep +  line 

		Logger.StdOut << message 

	}
}

def isRecord(line){
	tokens=line.split(",")
	if(tokens.size() < 4){ return false}
	if(tokens[0][0] != '"') {return false}
	
	return true
}

def isMultiLineEntry(line){
	if (line.indexOf("-----") == 0){
		return false
	}
	if (line.contains("Listing the") == true){
		return false
	}
	
	tokens=line.split(",")
	if (tokens.size() < 3){
		return true
	}
	return false
}


//def static fileHandles = [:]
//Set Argument Defaults
parameters = arguments( args  ) 
def evtSrc = parameters["evtSrc"]
def elq = new EventLogQuery(evtSrc,bundleDir)

 reader = elq.execute()  
/*Tue Jul 09 15:12:05 BST 2013
cscript equery.vbs /fi "Datetime gt 7/09/2013,03:12:05PM" /fo csv /L system
Microsoft (R) Windows Script Host Version 5.6
Copyright (C) Microsoft Corporation 1996-2001. All rights reserved.

 
INFO: No records available for the 'system' log with the specified criteria.
 
*/
Logger logger=new Logger()
bufferedLines=[]
reader.eachLine(){ line -> 
	tokens=line.split(",")
	//if (tokens.size() >=  4){
	if (isRecord(line)){
		if(bufferedLines.size()>0){
			Logger.StdOut << new Date()  << " " << line; 
			Logger.log(bufferedLines.join("\n"),new Date())
			Logger.StdOut << "\n";		
		}
		bufferedLines=[]
		bufferedLines.add(line)
	}
	
	if(isMultiLineEntry(line)){
		bufferedLines.add(line)
	}
}


