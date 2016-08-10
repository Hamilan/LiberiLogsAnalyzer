using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualBasic;

namespace LiberiLogsAnalyser
{
    /*Merges all the daily reports into one summarized accumulated report, adding a weekly report every week (7 days)
     */
    class ProgramMergeAnalysis
    {
        enum Parameters
        {
            none = 0,
            reportsOutputPath = 1,
            participants = 2,
            produceWeekReportColum = 3,
            produceDailyReportColumn = 4
        };


        static string reportsFolder;
        static string participants;
        static bool produceWeekReportColum = true;
        static bool produceDailyReportColumn = true;

        static void Main(string[] args)
        {
            getParamaterValues();
            String[] idsList = participants.Split(','); 
            foreach (string id in idsList)
            {
                mergeAnalysisFiles(id);
                Console.WriteLine("Reports for "+id+" done.");
            }
            Microsoft.VisualBasic.Interaction.MsgBox("All analysis files have been merged for participants "+participants);
        }
      

        private static void getParamaterValues()
        {
            
            String pathToUserProfile = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            StreamReader sr = new StreamReader("ConfigurationParameters.txt");
            String line;
            Parameters parameter = (Parameters)0; //use this as place holder value will be overwritten.

            //add each lines parameter values to the associated variable 
            while ((line = sr.ReadLine()) != null)
            {
                String[] lineData = line.Trim().Split('=');
                try
                {
                    parameter = (Parameters)Enum.Parse(typeof(Parameters), lineData[0].ToString());
                }
                catch (ArgumentException) {//line empty-clear paramater value so no variable is overwritten
                    parameter = (Parameters)Enum.Parse(typeof(Parameters),"none");
                }
                try
                {
                    switch (parameter)
                    {
                        case Parameters.reportsOutputPath:
                            reportsFolder = pathToUserProfile + lineData[1];
                            break;
                        case Parameters.participants:
                            participants = lineData[1];
                            break;
                        case Parameters.produceWeekReportColum:
                            produceWeekReportColum = System.Convert.ToBoolean(lineData[1]);
                            break;
                        case Parameters.produceDailyReportColumn:
                            produceDailyReportColumn = System.Convert.ToBoolean(lineData[1]);
                            break;
                        default:
                            //parameter is unknown - do nothing
                            break;
                    }
                } catch (Exception e)   {
                    Console.WriteLine(e);
                }
            }
            sr.Close();
        }

        //check if titles week reoprt is an average
        private static bool needsToBeAveraged(String title)
        {
            if (title.Contains("Average ") || title == "Number of different games played"
                    || title == "Seconds to start first game" || title == "Players on island when warmup finished"
                    || title == "Trips involving zone transfer" || title == "Seconds to start first game" || title == "Seconds to start first game while cadence>0"  ||title.Contains("% of") || title.Contains("Target HR"))
            {
                return true;
            }else {
                return false;
            }
        }

        //check if titles week reoprt is an average
        private static bool daysNotPlayIncreased(String title)
        {
            if (title.Contains("Max ") || title == "Minimum HR" || title.Contains("Average ") || title == "Number of different games played" || title == "Seconds to start first game" || title == "Players on island when warmup finished" || title == "Trips involving zone transfer" || title == "Total seconds spent travelling between shops/games" || title == "Pedaling Seconds spent travelling between shops/games"
                || title == "Seconds to start first game while cadence>0")
            {
                return true;
            } else {
                return false;
            }
        }

        private static void printIfDayReportsNeeded(StreamWriter outfile , String value )
        {
            if(produceDailyReportColumn)
                outfile.Write(value);
        }

        //Creates a log file for one participant "id"
        static void mergeAnalysisFiles(string id)
        {
            List<string> titles = new List<string>();
            List<List<string>> dataMatrix = new List<List<string>>();
            string logInputFilesPath = reportsFolder + "\\" + id + "\\";


            string pattern = "*_DailyReports.csv";
            List<string> filePaths = new List<string>(Directory.GetFiles(logInputFilesPath, pattern));
            //loop through every log file created during the game session
            int counter = 0;
            foreach (string fileIn in filePaths)
            {
                try
                {
                    StreamReader sr = new StreamReader(fileIn);
                    String line;
                    int i = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        String[] data = line.Split(',');
                        if (data[0]!="")
                        {
                            if (titles.Contains(data[0]) == false)
                            {
                                titles.Add(data[0]);
                            }
                            if (dataMatrix.Count <= i)
                            {
                                dataMatrix.Add(new List<string>());
                            }
                            if(data.Length>1)
                                dataMatrix[i].Add(data[1]);
                            i++;
                        }
                    }
                    if ( titles.Contains("Week") == false)   
                    {
                        titles.Add("Week");
                    }
                    if (dataMatrix.Count <= i)
                    {
                        dataMatrix.Add(new List<string>());
                    }
                    dataMatrix[i].Add("" + (counter / 7 + 1));
                    counter++;
                    i++;
                    sr.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception merging analysis file " + fileIn +e);
                }
            }

            String outputFile = reportsFolder + "\\" + id + "\\" + id + "_WeeklyReport.csv";
            StreamWriter outfile;
            outfile = new StreamWriter(outputFile);
            for (int i = 0; i < titles.Count; i++)
            {
                int column = 0;
                outfile.Write(titles[i] + ",");
                    
                float max = 0;
                float sum = 0;
                float min = 999;

                int daysCounter = 0;
                int daysDidNotPlay = 0;
                float fValue = 0;
                foreach (string value in dataMatrix[i])
                {
                    fValue = 0; 
                    daysCounter++;
                    if (titles[i] != "Date")    //only row that is not a number
                    {
                        fValue = float.Parse(value);
                        sum += fValue;
                        if (fValue > 0 && fValue < min)
                            min = fValue;
                        if (fValue > max)
                            max = fValue;

                        if (daysNotPlayIncreased(titles[i]))
                        {
                            if (fValue == 0)
                            {
                                daysDidNotPlay++;
                            }
                            printIfDayReportsNeeded(outfile, value + ",");
                        }
                        else if (titles[i].Contains("% of"))
                        {
                            //------------------SENSITIVE CODE:  If the rows in the Analysis files change, the number to substract might change too.
                            //This is for averages
                            if (fValue == 0 && float.Parse(dataMatrix[ i - 8 ][column])==0)
                            {
                                daysDidNotPlay++;
                                printIfDayReportsNeeded(outfile, ",");
                            }
                            else
                                printIfDayReportsNeeded(outfile, value + ",");
                        }
                        else
                            printIfDayReportsNeeded(outfile, value + ",");
                    }
                    else
                        printIfDayReportsNeeded(outfile, value + ",");

                        
                    if (daysCounter == 7)
                    {
                        if (produceWeekReportColum)
                        {
                            if (titles[i].Contains("Date"))
                            {
                                outfile.Write("WEEK,");
                            }
                            else if (titles[i].Contains("Max "))
                            {
                                outfile.Write(max + ",");
                            }
                            else if (titles[i] == "Minimum HR")
                            {
                                outfile.Write(min + ",");
                            }
                            else if (needsToBeAveraged(titles[i]))
                            {
                               
                                    outfile.Write(sum / (daysCounter - daysDidNotPlay) + ",");
                                
                            }
                            else if (titles[i].Contains("Minutes") || titles[i].Contains("Times ") || titles[i].Contains("browsing") || titles[i] == "Number of days played" || titles[i] == "Total seconds spent travelling between shops/games" || titles[i] == "Pedaling Seconds spent travelling between shops/games")             
                            {   

                                outfile.Write(sum + ",");
                                
                            }
                            else if (titles[i] == "Week")
                            {
                                outfile.Write((column / 7) + 1 + ",");
                            }
                            else
                            {
                                outfile.Write(",");
                            }
                        }
                        daysCounter = 0;
                        daysDidNotPlay = 0;
                        max = 0;
                        min = 999;
                        sum = 0;
                    }
                    column++;
                }
                if(daysCounter>0)
                {
                    if (produceWeekReportColum)
                    {
                        if (titles[i].Contains("Date"))
                        {
                            outfile.Write("WEEK,");
                        }
                        else if (titles[i].Contains("Max "))
                        {
                            outfile.Write(max + ",");
                        }
                        else if (titles[i] == "Minimum HR")
                        {
                            outfile.Write(min + ",");
                        }
                        else if (needsToBeAveraged(titles[i]))
                        {
                           
                                outfile.Write(sum / (daysCounter - daysDidNotPlay) + ",");
                          
                        }

                        else if (titles[i].Contains("Minutes") || titles[i].Contains("Times ") || titles[i].Contains("browsing") || titles[i] == "Number of days played")
                        {
                            outfile.Write(sum + ",");
                          
                        }
                        else if (titles[i] == "Week")
                        {
                            outfile.Write((column / 7) + 1 + ",");
                        }
                        else
                        {
                            outfile.Write(",");
                        }
                    }
                    daysCounter = 0;
                    daysDidNotPlay = 0;
                    max = 0;
                    min = 999;
                    sum = 0;
                }
                outfile.WriteLine(); //finish Line
                
                //add an empty line after the next rows
                if (titles[i].Contains("Target HR ") || titles[i] == "Minutes over 90% HRR" || titles[i] == "Minutes over 90% HRMax" || titles[i] == "Minutes btwn 80% and 89% HRR" || titles[i] == "Average cadence" || titles[i] == "Total time browsing shops and inventory" || titles[i] == "% of Island time over THR" || titles[i] == "Time spent browsing inventory" || titles[i] == "Pedaling Seconds spent travelling between shops/games")
                {
                    outfile.WriteLine();
                }
            }
            outfile.Close();
        }
    }
}
