using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.VisualBasic;

/*
 History of updates:
 * 2013-0....: added  count of seconds over 10% hrr, 20% hrr, 30% hrrr. Requested by Briar.
 * 2013-07-17: added  "&& cadence > 0" to all the decisions where seconds over certain hrr is counted. Requested by Briar.
 */

namespace LiberiLogsAnalyser
{
    class Program
    {
        //parameters that are available in the Configuration file (ConfigurationPArameters.txt)
        enum Parameters
        {   none = 0,
            skipReportFileCreation ,
            logEventsInMergeFile,
            overWriteMergedLogs,
            omitSimulatedHRFiles,
            startDateValues,
            targetHRPercentage,
            logsInputPath,
            projectInitials,
            reportsOutputPath,
            simulatedHRColumn,
            trialLengthInDays};

        static string dateToAnalyze;

        static Hashtable participantsProfiles;
        private static bool skipReportFileCreation = false;
        private static bool logEventsInMergeFile = true;
        private static bool overWriteMergedLogs = false;
        private static bool omitSimulatedHRFiles = true;
        private static int[] startDateValues = new int[3];
        private static float targetHRPercentage = 0.4f;
        
        private static string logsInputPath ;
        private static string projectInitials ;
        private static string reportsOutputPath ;
        private static int simulatedHRColumn = 7;
        private static int trialLengthInDays = 0;


        //*******Gets date from user then creates a log file for each player for each date**********//
        static void Main(string[] args)
        {
            //Create a log file for each participant if they played on the above date, store which participants you made logs for
            //to display at end of program execution
            getParamaterValues();
            getParticipantsProfiles();
            checkConfigParamsSet();
            DateTime today = DateTime.Now;
            string todayString;
            todayString = today.Year + "-";
            todayString += today.Month < 10 ? "0" + today.Month : "" + today.Month;
            todayString += "-";
            todayString += today.Day < 10 ? "0" + today.Day : "" + today.Day;

            string input = Microsoft.VisualBasic.Interaction.InputBox("Enter the date you want the report for.\nWrite 'all' for every day since the beginning of the trial until today):","Date for report",todayString);
            if (input.ToLower() == "all")
                analyzeAllDates(todayString);
            else
                try
                {
                    DateTime.Parse(input);
                    analizeDate(input);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Wrong date");
                }
        }

        private static void checkConfigParamsSet()
        {
           if( logsInputPath == null || projectInitials == null || reportsOutputPath == null || 
                startDateValues == null || trialLengthInDays == 0 || participantsProfiles == null)
           {
               Console.WriteLine("A critical parameter is missing.");
               Environment.Exit(0);
           }
        }

        //Load the parameters required from a configuration file
        private static void getParamaterValues()
        {
            String pathToUserProfile = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            StreamReader sr = new StreamReader("ConfigurationParameters.txt");
            String line;
            Parameters parameter = (Parameters)1; //use this as place holder any incorrect value will be overwritten.

            //add each lines parameter values to the associated variable 
            while ((line = sr.ReadLine()) != null)
            {
                String[] lineData = line.Split('=');
                lineData[1].Trim();
                try
                {
                    parameter = (Parameters)Enum.Parse(typeof(Parameters), lineData[0].ToString());
                }
                catch (ArgumentException) {//line empty-clear paramater value so no variable is overwritten
                    parameter = (Parameters)Enum.Parse(typeof(Parameters),"none");
                }
                try { 
                    switch (parameter)
                    {
                        case Parameters.skipReportFileCreation:
                            skipReportFileCreation = System.Convert.ToBoolean(lineData[1]);
                            break;
                        case Parameters.logEventsInMergeFile:
                            logEventsInMergeFile = System.Convert.ToBoolean(lineData[1]);
                            break;
                        case Parameters.overWriteMergedLogs:
                            overWriteMergedLogs = System.Convert.ToBoolean(lineData[1]);
                            break;
                        case Parameters.omitSimulatedHRFiles:
                            omitSimulatedHRFiles = System.Convert.ToBoolean(lineData[1]);
                            break;
                        case Parameters.targetHRPercentage:
                            float.TryParse(lineData[1], out targetHRPercentage);
                            break;
                        case Parameters.logsInputPath:
                            //not sure about @
                            logsInputPath = pathToUserProfile + lineData[1];
                            break;
                        case Parameters.projectInitials:
                            //not sure about @
                            projectInitials =  lineData[1].TrimStart();
                            break;
                        case Parameters.reportsOutputPath:
                            //not sure about @
                            reportsOutputPath = pathToUserProfile+ lineData[1];
                            break;
                        case Parameters.simulatedHRColumn:
                            int.TryParse(lineData[1], out simulatedHRColumn);
                            break;
                        case Parameters.trialLengthInDays:
                            int.TryParse(lineData[1], out trialLengthInDays);
                            break;
                        case Parameters.startDateValues:
                            string[] temp = lineData[1].Split(',');  //do more
                            startDateValues[0] = System.Convert.ToInt32(temp[0]);
                            startDateValues[1] = System.Convert.ToInt32(temp[1]);
                            startDateValues[2] = System.Convert.ToInt32(temp[2]);

                            break;
                        default:
                          //parameter is unknown - do nothing
                            break;
                    }
                }catch(Exception e){
                     Console.WriteLine(e);
                }
            }
            sr.Close();

        }
        
        //Loads participants information
        private static void getParticipantsProfiles()
        {
            StreamReader sr = new StreamReader("ParticipantsProfiles.txt");
            sr.ReadLine(); //drop first line with titles
            participantsProfiles = new Hashtable();

            String line;
            while ((line = sr.ReadLine()) != null)
            {
                String[] lineData = line.Split(',');
                participantsProfiles.Add(lineData[0], lineData);
            }
            sr.Close();
        }

        /*
         * Analyzes all the log files until the date sent by parameter "today"
         */
        private static void analyzeAllDates(string today)
        {
            string result = "";
            DateTime startDate = new DateTime(startDateValues[0],startDateValues[1],startDateValues[2]);  //beginning of the RCT trial
            List<string> dates = new List<string>();
            string dateString;
            
            //creates the list of dates to analyze
            for (int i = 0; i < trialLengthInDays; i++)    
            {
                dateString = startDate.Year+"-";
                dateString += startDate.Month < 10 ? "0" + startDate.Month : "" + startDate.Month;
                dateString += "-";
                dateString += startDate.Day < 10 ? "0" + startDate.Day : ""+startDate.Day;
                dates.Add(dateString);
                if(dateString == today)
                    break;
                startDate = startDate.AddDays(1);
            }
            //Analyzes all dates
            foreach (string theDate in dates)
            {
                dateToAnalyze = theDate;
                result = "";
                //analyzes all participants read from the profiles file.
                foreach (string id in participantsProfiles.Keys)
                {
                    result = analizeParticipant(id) + " ";
                }
                Console.WriteLine(dateToAnalyze + " daily reports created.");
            }
            Console.WriteLine("Done");
            Microsoft.VisualBasic.Interaction.MsgBox(dates.Count + " daily report files created.");
        }
        
        //Analyzes all participants' logs for an specific date
        private static void analizeDate(string theDate)
        {
            string result = "";
            dateToAnalyze = theDate;
            result = "";
            //analyzes all participants read from the profiles file.
            foreach (string id in participantsProfiles.Keys)
            {
                result = analizeParticipant(id) + " ";
            }
            Console.WriteLine(dateToAnalyze + " daily reports created");
            Console.WriteLine("Done");
            Microsoft.VisualBasic.Interaction.MsgBox("Daily report files created");
        }

        //********Analyzes the log files from a participant "id" for the date "dateToAnalyze" and creates a report file for that date ************//
        static string analizeParticipant(string id)
        {
            bool inDozo = false;
            bool inGekku = false;
            bool inBiri = false;
            bool inBobo = false;
            bool inWiskin = false;
            bool inPogi = false;
            bool inIsland = false;
            bool inWarmup = false;
            bool inCooldown = false;
            bool inMiniCooldown = false;
            bool browseHeadGear = false;
            bool browseHeadShop = false;
            bool browseGorgeousGekku = false;
            bool browseInventory = false;
            bool browseBodyShop = false;
            bool browseHandShop = false;
            bool browseFeetShop = false;
            bool browseAnything = false;
            bool switchedZones = false;
            bool beforeFirstGame = false;

            int tripsInvolvingZoneTransfer = 0;
            int secondsToFirstGameTotal = 0;
            int secondsToFirstGamePedaling = 0;
            int zoneTransferTripTimeTotal = 0;
            int zoneTransferTripTimePedal = 0;
            int tempZoneTransferTimeTotal = 0;
            int tempZoneTransferTimePedal = 0;
            string lastEvent = ""; 

            int totalBrowseSeconds = 0; 
            int feetShopSeconds = 0; 
            int bodyShopSeconds = 0;
            int headShopSeconds = 0;
            int handShopSeconds = 0;
            int inventorySeconds = 0;
            int headGearSeconds = 0;
            int gorgeousGekkuSeconds = 0;

            int dozoTotalSeconds = 0;
            int gekkuTotalSeconds = 0;
            int biriTotalSeconds = 0;
            int boboTotalSeconds = 0;
            int wiskinTotalSeconds = 0;
            int pogiTotalSeconds = 0;
            int warumpTotalSeconds = 0;
            int cooldownTotalSeconds = 0;
            int miniCooldownSeconds = 0;
            int islandTotalSeconds = 0;


            int numPlayersAtWarmUpEnd = 0;
            int totalPlayersColumn;
            int eventColumn;
            int minigameColumn;

            int secondsOverTHR = 0;
            int dozoTimeOverTHR = 0;
            int gekkuTimeOverTHR = 0;
            int biriTimeOverTHR = 0;
            int boboTimeOverTHR = 0;
            int wiskinTimeOverTHR = 0;
            int pogiTimeOverTHR = 0;
            int islandTimeOverTHR = 0;

            int reachedPlayTimeLimit = 0;
            int reachedTimeAtTHRLimit = 0;

            int timesSentToCooldown = 0;
            int secondsConnected = 0;
            int secondsPedaled = 0;

            string gamesPlayed = "";
            string[] profileData = (string[])(participantsProfiles[id]);
            int rhr = Int32.Parse(profileData[1]);
            int maxHR = Int32.Parse(profileData[2]); ;
            int hrr = (maxHR - rhr);
            
            float targetHR = rhr + targetHRPercentage * hrr;

            //Levels of HR reserve
            float hrr10 = rhr + 0.10f * hrr,
                hrr20 = rhr + 0.2f * hrr,
                hrr30 = rhr + 0.3039f * hrr,
                hrr40 = rhr + 0.40f * hrr,
                hrr45 = rhr + 0.45f * hrr,
                hrr50 = rhr + 0.50f * hrr,
                hrr60 = rhr + 0.6f * hrr,
                hrr65 = rhr + 0.65f * hrr,
                hrr70 = rhr + 0.70f * hrr,
                hrr80 = rhr + 0.80f * hrr,
                hrr90 = rhr + 0.9f * hrr;
            
            //Levels of HR max
            float hrMax30 = 0.3f * maxHR,
                hrMax40 = 0.4f * maxHR,
                hrMax50 = 0.5f * maxHR,
                hrMax60 = 0.6f * maxHR,
                hrMax70 = 0.70f * maxHR,
                hrMax80 = 0.80f * maxHR,
                hrMax90 = 0.9f * maxHR;

            //time over each level of hr reserve and hr max
            int secondsUnder30hrr = 0,
                secondsOver10hrr = 0,
                secondsOver20hrr = 0,
                secondsOver30hrr = 0,
                secondsOver40hrr = 0,
                secondsOver45hrr = 0,
                secondsOver50hrr = 0,
                secondsOver60hrr = 0,
                secondsOver65hrr = 0,
                secondsOver70hrr = 0,
                secondsOver80hrr = 0,
                secondsOver90hrr = 0,
                secondsOver30Max = 0,
                secondsOver40Max = 0,
                secondsOver50Max = 0,
                secondsOver60Max = 0,
                secondsOver70Max = 0,
                secondsOver80Max = 0,
                secondsOver90Max = 0;

            List<float> hr = new List<float>();
            List<float> rpm = new List<float>();
            
            string logFilesPath = logsInputPath + projectInitials + id + "\\";
            string outputFilePath = reportsOutputPath + id + "\\";
            string pattern = id + "_" + dateToAnalyze + "_1*.csv"; //the 1 represents the first character of hours after between 10:00 and 19:59 (7:59pm) This is to omit analyzing Reports files
            List<string> filesPaths = new List<string>(Directory.GetFiles(logFilesPath, pattern));

            //We probably want to merge when an analysis has not been made for this date yet.
            //Merging puts all the log files from the same day in one single log file.
            bool mergeLogs = true;
            if (File.Exists(outputFilePath + id + "_" + dateToAnalyze + ".csv") && !overWriteMergedLogs)
                mergeLogs = false;
            StreamWriter mergedLogsFile=null;
            if(mergeLogs)
                mergedLogsFile = new StreamWriter(outputFilePath + id + "_" + dateToAnalyze + ".csv");

            bool wroteTitles = false;
            //loop through every log file created during the game session
            foreach (string logFile in filesPaths)
            {
                bool simulatedHR = false;     
                browseGorgeousGekku = browseHeadGear = browseBodyShop = browseFeetShop = browseHandShop = browseHeadGear = browseHeadShop = browseInventory = false;
                inDozo = inGekku = inBiri = inBobo = inWiskin = inPogi = inIsland = inCooldown = inWarmup = inMiniCooldown = false;
                try
                {
                    StreamReader sr = new StreamReader(logFile);
                    String line = sr.ReadLine(); //first line with titles
                    if (mergeLogs && wroteTitles == false)
                            mergedLogsFile.WriteLine(line);

                    if (line.Contains("Sim. Heart Rate"))    //earlier logs did not have this column
                    {
                        line = sr.ReadLine();  //second line with some data, like cadence cap.
                        String[] titles = line.Split(',');
                        if (titles[simulatedHRColumn] == "TRUE")    //if columns order changes, this index has to change
                        {
                            simulatedHR = true;
                            if (omitSimulatedHRFiles)
                            {
                                sr.Close();
                                continue;
                            }
                        }
                    }
                    else
                    {
                        line = sr.ReadLine();  //second line with some data, like cadence cap.
                    }
                    if(mergeLogs && wroteTitles==false)
                            mergedLogsFile.WriteLine(line);
                    line = sr.ReadLine();// //third line with titles for stream of data
                    String[] titleList = line.Split(',');
                    if (mergeLogs && wroteTitles == false)
                            mergedLogsFile.WriteLine(line);
                    wroteTitles = true;


                    //find the column number of specific titles
                    totalPlayersColumn = Array.IndexOf(titleList, "Total Players");
                    eventColumn = Array.IndexOf(titleList, "Event");
                    minigameColumn = eventColumn + 1;
                    
                    //------------------------------------------------------------------------------------------------------
                    //Format of each line should be:
                    //Type, ServerTime, ClientTime, HR,  Cadence, X,    Y,    Z,    LookX, LookY, Energy, HasBuff,    Event, Minigame, MinigamePlayers, TotalPlayers
                    // M/E, HH:MM:SS,   HH:MM:SS,   ###, ###,     ##.#, ##.#, ##.#, ##.#,  ##.#,  #.#,    TRUE/FALSE, abcde, abcde,    #,               #
                    //------------------------------------------------------------------------------------------------------
                    bool reportedExceptionAboutTimeInMinigames = false;
                    inIsland = true;

                    while ((line = sr.ReadLine()) != null)
                    {
                        String[] data = line.Split(',');
                        if (data[0].Equals("M"))    //only count the measure ("M") lines to calculate time. Event ("E") lines can happen more than once a second.
                        {
                            secondsConnected++;
                            float singleHR = float.Parse(data[3]);
                            if (simulatedHR)
                            {
                                singleHR = 0;
                            }                                      
                            float cadence = float.Parse(data[4]);


                            if (mergeLogs)
                               // if (singleHR > 0)
                                    mergedLogsFile.WriteLine(line);

                           
                            
                            if (cadence > 0)
                            {
                                secondsPedaled++;
                                tempZoneTransferTimePedal++;
                            }
                            if (singleHR > 0)
                            {
                                //this part can be improved with a list or Hashtable
                                if (singleHR >= hrr10)
                                {
                                    secondsOver10hrr++;
                                    if (singleHR >= hrr20)
                                    {
                                        secondsOver20hrr++;
                                        if (singleHR >= hrr30)
                                        {
                                            secondsOver30hrr++;
                                            if (singleHR >= hrr40)
                                            {
                                                secondsOver40hrr++;
                                                if (singleHR >= hrr45)
                                                {
                                                    secondsOver45hrr++;
                                                    if (singleHR >= hrr50)
                                                    {
                                                        secondsOver50hrr++;
                                                        if (singleHR >= hrr60)
                                                        {
                                                            secondsOver60hrr++;
                                                            if (singleHR >= hrr65)
                                                            {
                                                                secondsOver65hrr++;
                                                                if (singleHR >= hrr70)
                                                                {
                                                                    secondsOver70hrr++;
                                                                    if (singleHR >= hrr80)
                                                                    {
                                                                        secondsOver80hrr++;
                                                                        if (singleHR >= hrr90)
                                                                        {
                                                                            secondsOver90hrr++;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (singleHR < hrr30)
                                {
                                    secondsUnder30hrr++;
                                }
                                //Calculations with HR MAx
                                if (singleHR >= hrMax30)
                                {
                                    secondsOver30Max++;
                                    if (singleHR >= hrMax40)
                                    {
                                        secondsOver40Max++;
                                        if (singleHR >= hrMax50)
                                        {
                                            secondsOver50Max++;
                                            if (singleHR >= hrMax60)
                                            {
                                                secondsOver60Max++;
                                                if (singleHR >= hrMax70)
                                                {
                                                    secondsOver70Max++;
                                                    if (singleHR >= hrMax80)
                                                    {
                                                        secondsOver80Max++;
                                                        if (singleHR >= hrMax90)
                                                        {
                                                            secondsOver90Max++;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            //Count seconds over target heart rate even if cadence is zero
                            if (singleHR >= targetHR)
                            {
                                secondsOverTHR++;
                                if (inDozo)
                                    dozoTimeOverTHR++;
                                else if (inGekku)
                                    gekkuTimeOverTHR++;
                                else if (inBiri)
                                    biriTimeOverTHR++;
                                else if (inBobo)
                                    boboTimeOverTHR++;
                                else if (inWiskin)
                                    wiskinTimeOverTHR++;
                                else if (inPogi)
                                    pogiTimeOverTHR++;
                                else if (inIsland)
                                    islandTimeOverTHR++;
                            }
                            if(!beforeFirstGame){
                                tempZoneTransferTimeTotal++;
                            }
                            // if (singleHR > 0)
                            {
                                hr.Add(singleHR);
                                rpm.Add(cadence);   //only used for writing an average cadence

                                if (inDozo)
                                    dozoTotalSeconds++;
                                else if (inGekku)
                                    gekkuTotalSeconds++;
                                else if (inBiri)
                                    biriTotalSeconds++;
                                else if (inBobo)
                                    boboTotalSeconds++;
                                else if (inWiskin)
                                    wiskinTotalSeconds++;
                                else if (inPogi)
                                    pogiTotalSeconds++;
                                else if (inIsland)
                                    islandTotalSeconds++;
                                else if (inWarmup)
                                    warumpTotalSeconds++;
                                else if (inCooldown)
                                    cooldownTotalSeconds++;
                                else if (inMiniCooldown)
                                    miniCooldownSeconds++;
                                else if (beforeFirstGame)
                                {
                                    secondsToFirstGameTotal++;
                                    if (cadence > 0)
                                        secondsToFirstGamePedaling++;
                                }
                             }
                            if (browseAnything){
                                if (browseHeadGear) {
                                    totalBrowseSeconds++;
                                    headGearSeconds++;
                                }else if (browseBodyShop)  {
                                    totalBrowseSeconds++;
                                    bodyShopSeconds++;
                                }else if (browseFeetShop)  {
                                    totalBrowseSeconds++;
                                    feetShopSeconds++;
                                }else if (browseGorgeousGekku){
                                    totalBrowseSeconds++;
                                    gorgeousGekkuSeconds++;
                                }else if (browseHandShop) {
                                    totalBrowseSeconds++;
                                    handShopSeconds++;
                                }else if (browseHeadShop) {
                                    totalBrowseSeconds++;
                                    headShopSeconds++;
                                }else if (browseInventory) {
                                    totalBrowseSeconds++;
                                    inventorySeconds++;
                                }

                            }

                        }
                        else if (data[0].Equals("E")) //check for times in minigames
                        {
                            //float singleHR = float.Parse(data[3]);
                            if (logEventsInMergeFile && mergeLogs )// && singleHR > 0)
                            {
                                mergedLogsFile.WriteLine(line);
                            }
                            try
                            {
                                if (data[eventColumn].Contains("Entered zone"))
                                {
                                    switchedZones = true;
                                }
                                else if (data[eventColumn].Contains("Started minigame:"))
                                {
                                    
                                    if (switchedZones || !lastEvent.Equals(data[minigameColumn]) )
                                    {
                                        switchedZones = false;
                                        tripsInvolvingZoneTransfer++;
                                        zoneTransferTripTimeTotal = zoneTransferTripTimeTotal +tempZoneTransferTimeTotal;
                                        zoneTransferTripTimePedal = zoneTransferTripTimePedal +tempZoneTransferTimePedal;
                                    }
                                    if (switchedZones)
                                    { //if player switched zones before entering increase number of island trips
                                        switchedZones = false;
                                        tripsInvolvingZoneTransfer++;
                                    }
                                    if (beforeFirstGame)
                                        beforeFirstGame = false;

                                    inIsland = false;
                                    if (data[minigameColumn].Equals("Dozo Quest")){
                                        inDozo = true;
                                        lastEvent = "Dozo Quest";
                                        if (!gamesPlayed.Contains("Dozo Quest")){
                                            gamesPlayed  = gamesPlayed +"Dozo Quest, ";
                                        }
                                    }else if (data[minigameColumn].Equals("Gekku Race")){
                                        inGekku = true;
                                        lastEvent = "Gekku Race";
                                        if (!gamesPlayed.Contains("Gekku Race"))
                                        {
                                            gamesPlayed = gamesPlayed + "Gekku Race, ";
                                        }
                                    }else if (data[minigameColumn].Equals("Biri Brawl")){
                                        inBiri = true;
                                        lastEvent = "Biri Brawl";
                                        if (!gamesPlayed.Contains("Biri Brawl"))
                                        {
                                            gamesPlayed = gamesPlayed + "Biri Brawl, ";
                                        }
                                    }else if (data[minigameColumn].Equals("Round Up")){
                                        inBobo = true;
                                        lastEvent = "Round Up";
                                        if (!gamesPlayed.Contains("Round Up"))
                                        {
                                            gamesPlayed = gamesPlayed + "Round Up, ";
                                        }
                                    }else if (data[minigameColumn].Equals("Wiskin Defence")){
                                        inWiskin = true;
                                        lastEvent = "Wiskin Defense";
                                        if (!gamesPlayed.Contains("Wiskin Defence"))
                                        {
                                            gamesPlayed = gamesPlayed + "Wiskin Defence, ";
                                        }
                                    }
                                    else if (data[minigameColumn].Equals("Pogi Pong")) { 
                                        inPogi = true;
                                        lastEvent = "Pogi Pong";
                                        if (!gamesPlayed.Contains("Pogi Pong"))
                                        {
                                            gamesPlayed = gamesPlayed + "Pogi Pong, ";
                                        }
                                    }
                                }
                                else if (data[eventColumn].Contains("Stopped minigame:"))
                                {
                                        inDozo = false;
                                        inGekku = false;
                                        inBiri = false;
                                        inBobo = false;
                                        inWiskin = false;
                                        inPogi = false;
                                        inIsland = true;
                                        tempZoneTransferTimePedal = 0;
                                        tempZoneTransferTimeTotal = 0;
                                }
                                else if(data[eventColumn].Contains("Entered Pedal Bear session: MiniCooldown"))
                                {
                                    timesSentToCooldown++;
                                    inMiniCooldown = true;
                                }
                                else if(data[eventColumn].Contains("Left Pedal Bear session: MiniCooldown"))
                                    inMiniCooldown = false;
                                else if (data[eventColumn].Contains("Entered Pedal Bear session: Warmup"))
                                    inWarmup= true;
                                else if (data[eventColumn].Contains("Left Pedal Bear session: Warmup"))  {
                                    inWarmup = false;
                                    beforeFirstGame = true;
                                    numPlayersAtWarmUpEnd = Convert.ToInt32(data[totalPlayersColumn]);
                                }
                                else if (data[eventColumn].Contains("Entered Pedal Bear session: QuitCooldown") || data[eventColumn].Contains("Entered Pedal Bear session: KickCooldown"))
                                    inCooldown = true;
                                else if (data[eventColumn].Contains("Left Pedal Bear session: QuitCooldown"))
                                    inCooldown = false;
                                else if (data[eventColumn].Contains("Started browsing"))
                                {
                                    inIsland = false;
                                    zoneTransferTripTimePedal += tempZoneTransferTimePedal;
                                    zoneTransferTripTimeTotal += tempZoneTransferTimeTotal;
                                    browseAnything = true;
                                    if (data[eventColumn].Contains("Gorgeous Gekku"))
                                    {
                                        lastEvent = "Gorgeous Gekku";
                                        browseGorgeousGekku = true;
                                    }
                                    else if (data[eventColumn].Contains("Headgear Shop"))
                                    {
                                        lastEvent = "Headgear Shop";
                                        browseHeadGear = true;
                                    }
                                    else if (data[eventColumn].Contains("inventory"))
                                    {
                                        browseInventory = true;
                                    }
                                    else if (data[eventColumn].Contains("Head Shop"))
                                    {
                                        lastEvent = "Head Shop";
                                        browseHeadShop = true;
                                    }
                                    else if (data[eventColumn].Contains("Body Shop"))
                                    {
                                        lastEvent = "Body Shop";
                                        browseBodyShop = true;
                                    }
                                    else if (data[eventColumn].Contains("Hands Shop"))
                                    {
                                        lastEvent = "Hands Shop";
                                        browseHandShop = true;
                                    }
                                    else if (data[eventColumn].Contains("Feet Shop"))
                                    {
                                        lastEvent = "Feet Shop";
                                        browseFeetShop = true;
                                    }


                                }
                                else if (data[eventColumn].Contains("Stopped browsing"))
                                {
                                    browseAnything = false;
                                    browseBodyShop = false;
                                    browseFeetShop = false;
                                    browseGorgeousGekku = false;
                                    browseHandShop = false;
                                    browseHeadGear = false;
                                    browseHeadShop = false;
                                    browseInventory = false;
                                    inIsland = true;
                                    tempZoneTransferTimeTotal = 0;
                                    tempZoneTransferTimePedal = 0;

                                }

                                if (data[eventColumn] == "Entered Pedal Bear session: KickCooldown - reached THR time limit.")
                                    reachedTimeAtTHRLimit = 1;
                                if (data[eventColumn] == "Entered Pedal Bear session: KickCooldown - reached play time limit.")
                                    reachedPlayTimeLimit = 1;

                                if (data[eventColumn].Contains("Entered Pedal Bear session"))
                                {
                                    if (switchedZones)
                                    { //if player switched zones before entering increase number of island trips 
                                        switchedZones = false;
                                        tripsInvolvingZoneTransfer++;
                                    }
                                    inDozo = false;
                                    inGekku = false;
                                    inBiri = false;
                                    inBobo = false;
                                    inWiskin = false;
                                    inPogi = false;
                                    inIsland = false;
                                }
                            }
                            catch (Exception e)
                            {
                                if (reportedExceptionAboutTimeInMinigames == false)
                                {
                                    reportedExceptionAboutTimeInMinigames = true;
                                    Console.WriteLine("Exception counting time in minigames. Probably minigame field is missing. (in " + logFile + ")\n"+e.Message);
                                }
                            }
                        }
                    }
                    sr.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Tried to read file " + logFile + " but was not found.\n"+e.Message);
                }
                Console.Write(".");
            }
            if (mergeLogs)
                mergedLogsFile.Close();
            
            if (skipReportFileCreation)
                return "";
            
            //----------------------------------------------------------------------------------
            // Do calculations and write the output file
            //----------------------------------------------------------------------------------
            float avgHR = 0;
            float maxReachedHR = 0;
            float minHR = 0;
            if (hr.Count > 0)
            {
                avgHR = hr.Average();
                maxReachedHR = hr.Max();
                try
                {
                    hr.RemoveAll(isZero);
                    minHR = hr.Min();
                }
                catch (Exception e)
                {   
                }
            }
            
            float maxRPM = 0;
            float avgRPM = 0;
            if (rpm.Count > 0)
            {
                maxRPM = rpm.Max();
                //rpm.RemoveAll(isZero);    //Not sure we should do this
                //Just in case the list is now empty
                if (rpm.Count > 0)
                {
                    avgRPM = rpm.Average();
                }
            }

            //find the number of games listed
            int numberOfGames = gamesPlayed.Split(',').Length -1;


            //StreamWriter outfile = new StreamWriter(outputFilePath + id + "_" + date + "_AnalysisForHam.csv");
            StreamWriter outfile = new StreamWriter(outputFilePath + id + "_" + dateToAnalyze + "_DailyReports.csv");
            outfile.WriteLine("Participant, " + id);
            outfile.WriteLine("Date, " + dateToAnalyze);
            outfile.WriteLine("Target HR (RHR+"+targetHRPercentage*100+"% HRR), " + (int)targetHR);
            outfile.WriteLine("");
            if (secondsConnected > 0)
                outfile.WriteLine("Number of days played, " + 1);
            else
                outfile.WriteLine("Number of days played, " + 0);
            outfile.WriteLine("Max reached HR, " + maxReachedHR);
            outfile.WriteLine("Minimum HR, " + minHR);
            outfile.WriteLine("Average HR, " + (int)avgHR);
            outfile.WriteLine("");
            //outfile.WriteLine("Minutes under 30% HRR, " + secondsUnder30 / 60f);
            outfile.WriteLine("Minutes over THR, " + secondsOverTHR / 60f);
            outfile.WriteLine("Minutes over 10% HRR, " + secondsOver10hrr / 60f);
            outfile.WriteLine("Minutes over 20% HRR, " + secondsOver20hrr / 60f);
            outfile.WriteLine("Minutes over 30% HRR, " + secondsOver30hrr / 60f);
            outfile.WriteLine("Minutes over 40% HRR, " + secondsOver40hrr / 60f);
            outfile.WriteLine("Minutes over 45% HRR, " + secondsOver45hrr / 60f);
            outfile.WriteLine("Minutes over 50% HRR, " + secondsOver50hrr / 60f);
            outfile.WriteLine("Minutes over 60% HRR, " + secondsOver60hrr / 60f);
            outfile.WriteLine("Minutes over 65% HRR, " + secondsOver65hrr / 60f);
            outfile.WriteLine("Minutes over 70% HRR, " + secondsOver70hrr / 60f);
            outfile.WriteLine("Minutes over 80% HRR, " + secondsOver80hrr / 60f);
            outfile.WriteLine("Minutes over 90% HRR, " + secondsOver90hrr / 60f);
            outfile.WriteLine("");
            outfile.WriteLine("Minutes over 30% HRMax, " + secondsOver30Max / 60f);
            outfile.WriteLine("Minutes over 40% HRMax, " + secondsOver40Max / 60f);
            outfile.WriteLine("Minutes over 50% HRMax, " + secondsOver50Max / 60f);
            outfile.WriteLine("Minutes over 60% HRMax, " + secondsOver60Max / 60f);
            outfile.WriteLine("Minutes over 70% HRMax, " + secondsOver70Max / 60f);
            outfile.WriteLine("Minutes over 80% HRMax, " + secondsOver80Max / 60f);
            outfile.WriteLine("Minutes over 90% HRMax, " + secondsOver90Max / 60f);
            outfile.WriteLine("");
            outfile.WriteLine("Minutes btwn 10% and 19% HRR, " + (secondsOver10hrr - secondsOver20hrr) / 60f);
            outfile.WriteLine("Minutes btwn 20% and 29% HRR, " + (secondsOver20hrr - secondsOver30hrr) / 60f);
            outfile.WriteLine("Minutes btwn 30% and 39% HRR, " + (secondsOver30hrr - secondsOver40hrr) / 60f);
            outfile.WriteLine("Minutes btwn 40% and 49% HRR, " + (secondsOver40hrr - secondsOver50hrr) / 60f);
            outfile.WriteLine("Minutes btwn 50% and 59% HRR, " + (secondsOver50hrr - secondsOver60hrr) / 60f);
            outfile.WriteLine("Minutes btwn 60% and 69% HRR, " + (secondsOver60hrr - secondsOver70hrr) / 60f);
            outfile.WriteLine("Minutes btwn 70% and 79% HRR, " + (secondsOver70hrr - secondsOver80hrr) / 60f);
            outfile.WriteLine("Minutes btwn 80% and 89% HRR, " + (secondsOver80hrr - secondsOver90hrr) / 60f);
            outfile.WriteLine("");
            outfile.WriteLine("Times sent to mini cool down, " + timesSentToCooldown);
            outfile.WriteLine("Times reached play time limit, " + reachedPlayTimeLimit );
            outfile.WriteLine("Times reached time at thr limit, " + reachedTimeAtTHRLimit );
            outfile.WriteLine("");
            outfile.WriteLine("Minutes pedaling, " + secondsPedaled / 60f);
            outfile.WriteLine("Minutes connected, " + secondsConnected / 60f);
            outfile.WriteLine("Max cadence, " + maxRPM);
            outfile.WriteLine("Average cadence, " + avgRPM);
            outfile.WriteLine("");
            outfile.WriteLine("Warmup Minutes, " + warumpTotalSeconds / 60f);
            outfile.WriteLine("Cooldown Minutes, " + cooldownTotalSeconds / 60f);
            outfile.WriteLine("Mini cooldown Minutes, " + miniCooldownSeconds / 60f);
            outfile.WriteLine("");
            outfile.WriteLine("Gekku Race Minutes, " + gekkuTotalSeconds / 60f);
            outfile.WriteLine("Dozo Quest Minutes, " + dozoTotalSeconds / 60f);
            outfile.WriteLine("Biri Brawl Minutes, " + biriTotalSeconds / 60f);
            outfile.WriteLine("Round Up Minutes, " + boboTotalSeconds / 60f);
            outfile.WriteLine("Wiskin Defence Minutes, " + wiskinTotalSeconds / 60f);
            outfile.WriteLine("Pogi Pong Minutes, " + pogiTotalSeconds / 60f);
            outfile.WriteLine("Island Minutes, " + islandTotalSeconds / 60f);
            outfile.WriteLine("Total time browsing shops and inventory, " + totalBrowseSeconds / 60f);
            outfile.WriteLine("");
            outfile.WriteLine("% of Gekku Race time over THR, " + (gekkuTotalSeconds > 0 ? ((float)gekkuTimeOverTHR) / gekkuTotalSeconds : 0));
            outfile.WriteLine("% of Dozo Quest time over THR, " + (dozoTotalSeconds > 0 ? ((float)dozoTimeOverTHR) / dozoTotalSeconds : 0));
            outfile.WriteLine("% of Biri Brawl time over THR, " + (biriTotalSeconds > 0 ? ((float)biriTimeOverTHR) / biriTotalSeconds : 0));
            outfile.WriteLine("% of Round Up time over THR, " + (boboTotalSeconds > 0 ? ((float)boboTimeOverTHR) / boboTotalSeconds : 0));
            outfile.WriteLine("% of Wiskin Defence time over THR, " + (wiskinTotalSeconds > 0 ? ((float)wiskinTimeOverTHR) / wiskinTotalSeconds : 0));
            outfile.WriteLine("% of Pogi Pong time over THR, " + (pogiTotalSeconds > 0 ? ((float)pogiTimeOverTHR) / pogiTotalSeconds : 0));
            outfile.WriteLine("% of Island time over THR, " + (islandTotalSeconds > 0 ? ((float)islandTimeOverTHR) / islandTotalSeconds : 0));
            outfile.WriteLine("");

            outfile.WriteLine("Players on island when warmup finished, " + numPlayersAtWarmUpEnd);
            outfile.WriteLine("Number of different games played, " + numberOfGames);
            outfile.WriteLine("Trips involving zone transfer, " + tripsInvolvingZoneTransfer);
            outfile.WriteLine("Seconds to start first game, " + secondsToFirstGameTotal); 
            outfile.WriteLine("Seconds to start first game while cadence>0, " + secondsToFirstGamePedaling);
            outfile.WriteLine("Total seconds spent travelling between shops/games, " + zoneTransferTripTimeTotal);
            outfile.WriteLine("Pedaling Seconds spent travelling between shops/games,"+ zoneTransferTripTimePedal);
            outfile.WriteLine("");
            outfile.WriteLine("Time spent browsing Gorgeous Gekku, " + gorgeousGekkuSeconds / 60f);
            outfile.WriteLine("Time spent browsing Hand Shop, " + handShopSeconds / 60f);
            outfile.WriteLine("Time spent browsing Head Shop, " + headShopSeconds / 60f);
            outfile.WriteLine("Time spent browsing Head Gear Shop, " + headGearSeconds / 60f);
            outfile.WriteLine("Time spent browsing Feet Shop, " + feetShopSeconds / 60f);
            outfile.WriteLine("Time spent browsing Body Shop, " + bodyShopSeconds / 60f);
            outfile.WriteLine("Time spent browsing inventory, " + inventorySeconds / 60f);
            outfile.WriteLine("");
            outfile.WriteLine("Gekku Race time over THR, " + gekkuTimeOverTHR / 60f);
            outfile.WriteLine("Dozo Quest time over THR, " + dozoTimeOverTHR / 60f );
            outfile.WriteLine("Biri Brawl time over THR, " + biriTimeOverTHR / 60f);
            outfile.WriteLine("Round Up time over THR, " + boboTimeOverTHR / 60f);
            outfile.WriteLine("Wiskin Defence time over THR, " + wiskinTimeOverTHR / 60f);
            outfile.WriteLine("Pogi Pong time over THR, " + pogiTimeOverTHR / 60f);
            outfile.WriteLine("Island time over THR, " + islandTimeOverTHR / 60f);

            outfile.Close();
            return id + " ";
        }//end analizeOneParticipant


        //Predicate for remove
        private static bool isZero(float num)
        {
            if (num == 0)
                return true;
            else
                return false;
        }
    }
}
