using System;
using System.Collections;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelScripts
{
    class Program
    {
        //parameters that are available in the Configuration file (ConfigurationPArameters.txt)
        enum Parameters
        {
            none = 0,
            reportsOutputPath = 1,
            participants = 2,
            numberOfWeeksSoFar = 3,
            outFile = 4
            
        };


        static string reportsFolder;
        private static string outFile;
        private static string[] participants;
        private static int numberOfWeeksSoFar;
        private static System.Object cell1PS;
        private static System.Object cell2PS;
        private static Excel.Range cellMinsInMinigameFromPS;
        private static Excel.Range cellMinsInMinigameToPS;
        private static Excel.Range cellPercentoverTHRFromPS;
        private static Excel.Range cellPercentoverTHRToPS;

        private static Hashtable rowsToSum = new Hashtable();
        private static  Hashtable rowsToAvg = new Hashtable();
      

        static void Main(string[] args)
        { 
            getParamaterValues();

            Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook newWB = app.ActiveWorkbook;
          
            CreateRowsSumForParticipantSheets(rowsToSum);
            CreateRowsAvgForParticipantSheets(rowsToAvg);
            copyLogInfoToResults(app, out newWB);
           createPedalingAndTHRSheet(newWB);
           createMinigameSheet(newWB);
            newWB.Save();
            newWB.Close(1);
            app.DisplayAlerts = false;
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

        }
        //which rows of the participant sheets need averaging
        private static Hashtable CreateRowsSumForParticipantSheets( Hashtable sum){
            sum.Add("Warmup Minutes", "Warmup Minutes");
            sum.Add("Cooldown Minutes","Cooldown Minutes");
            sum.Add("Mini cooldown Minutes", "Mini cooldown");
            sum.Add("Gekku Race Minutes","Gekku Race");
            sum.Add("Dozo Quest Minutes","Dozo Quest");            
            sum.Add("Biri Brawl Minutes","Biri Brawl");                      
            sum.Add("Round Up Minutes","Round Up");                      
            sum.Add("Wiskin Defence Minutes","Wiskin Defence");                    
            sum.Add("Pogi Pong Minutes", "Pogi Pong");  
           sum.Add("Island Minutes", "Island");                    
           sum.Add( "Total time browsing shops and inventory", "Browsing shops and inventory");                    
           sum.Add(  "Times reached play time limit",  "Times reached play time limit");                    
           sum.Add( "Times reached time at thr limit",  "Times reached time at thr limit");                    
            sum.Add(   "Total seconds spent travelling between shops/games", "Travelling between shops/games");                    
           sum.Add(   "Pedaling Seconds spent travelling between shops/games", "Pedaling between shops/games");                    
                                 

            return sum;
        }

        
        //which rows of the participant sheets need averaging
       private static Hashtable CreateRowsAvgForParticipantSheets( Hashtable avg){
            avg.Add("% of Gekku Race time over THR", "% of Gekku Race time over THR");
            avg.Add("% of Dozo Quest time over THR",  "% of Dozo Quest time over THR");
            avg.Add("% of Biri Brawl time over THR",   "% of Biri Brawl time over THR");
            avg.Add("% of Round Up time over THR",   "% of Round Up time over THR");
            avg.Add("% of Wiskin Defence time over THR", "% of Wiskin Defence time over THR");
            avg.Add("% of Pogi Pong time over THR",   "% of Pogi Pong time over THR");
            avg.Add("% of Island time over THR",      "% of Island time over THR");

            return avg;
        }


       private static void createDataAreaTHRandPedal(Excel.Worksheet ws, int topLeftRow, int topLeftCol, string formula)
       {
           try { 
           ws.Cells[topLeftRow+1, numberOfWeeksSoFar + 1 +topLeftCol] = "Average";
           ws.Cells[topLeftRow+1, numberOfWeeksSoFar + 2+ topLeftCol] = "StdDev";
           ws.Cells[participants.Length + topLeftRow +2, topLeftCol] = "Average";
           ws.Cells[participants.Length + topLeftRow + 3, topLeftCol] = "StdDev";
           string temp = formula;
           for (int p = 0; p < participants.Length; p++)
           {
               
               
               //set formulas
               ws.Cells[topLeftRow+2+p, topLeftCol] = "P" + participants[p];
               for (int w = 1; w <= numberOfWeeksSoFar; w++)
               {
                   temp = temp.Replace("{ID}", participants[p].ToString());
                   temp = temp.Replace("{Letter}",pickLetter(w));
                   //setFormulas in cells
                   ws.Cells[2+topLeftRow + p, topLeftCol + w].Formula = temp.ToString();

                   //add week titles
                   ws.Cells[topLeftRow+1, w + topLeftCol] = "W" + w;

                   //average/stdev the weeks
                   string weekfromCell = ((Excel.Range)ws.Cells[topLeftRow+2, w + topLeftCol]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                   string weektoCell = ((Excel.Range)ws.Cells[participants.Length + 1+topLeftRow, w + topLeftCol]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                   ws.Cells[participants.Length+topLeftRow + 2, w + topLeftCol].Formula = "=AVERAGE(" + weekfromCell + ":" + weektoCell + ")";
                   ws.Cells[participants.Length+topLeftRow + 3, w + topLeftCol].Formula = "=STDEV.P(" + weekfromCell + ":" + weektoCell + ")";
                   temp = formula;
               }

               //avg/stdev the participants
               string participantFromCell = ((Excel.Range)ws.Cells[topLeftRow +2 + p, topLeftCol+1]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
               string participantToCell = ((Excel.Range)ws.Cells[topLeftRow+2 + p, numberOfWeeksSoFar + topLeftCol]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
               ws.Cells[topLeftRow + 2 + p, numberOfWeeksSoFar + 1 + topLeftCol].Formula = "=AVERAGE(" + participantFromCell + ":" + participantToCell + ")";
               ws.Cells[topLeftRow + 2 + p, numberOfWeeksSoFar + 2+topLeftCol].Formula = "=STDEV.P(" + participantFromCell + ":" + participantToCell + ")";
           }
           }
           catch (Exception e)
           {
               ((Excel.Application)ws.Parent.Parent).Quit();
               Console.WriteLine(e);
               
           }

        }

       private static void createMinigameSheet(Excel.Workbook newWB)
       {
           try {
               Console.WriteLine("Creating sheet for minigame data");
           Excel.Worksheet ws = newWB.Sheets.Add();

           ws.Name = "Minigames";
           //create Data areas
           ws.Cells[1, 2] = "Total minutes in minigames";
           ws.Cells[3, 1] = "Warm Up"; ws.Cells[3, participants.Length + 2] = "Warm Up";
           ws.Cells[4, 1] = "Cool down"; ws.Cells[4, participants.Length + 2] = "Cool down";
           ws.Cells[5, 1] = "Mini Cooldown"; ws.Cells[5, participants.Length + 2] = "Mini Cooldown";
           ws.Cells[6, 1] = "Gekku Race"; ws.Cells[6, participants.Length + 2] = "Gekku Race";
           ws.Cells[7, 1] = "Dozo Quest"; ws.Cells[7, participants.Length + 2] = "Dozo Quest";
           ws.Cells[8, 1] = "Biri Brawl"; ws.Cells[8, participants.Length + 2] = "Biri Brawl";
           ws.Cells[9, 1] = "Bobo Ranch"; ws.Cells[9, participants.Length + 2] = "Bobo Ranch";
           ws.Cells[10, 1] = "Wiskin Defence"; ws.Cells[10, participants.Length + 2] = "Wiskin Defence";
           ws.Cells[11, 1] = "Pogi Pong"; ws.Cells[11, participants.Length + 2] = "Pogi Pong";
           ws.Cells[12, 1] = "Island"; ws.Cells[12, participants.Length + 2] ="Island";
           ws.Cells[13, 1] = "Shops and Inventory"; ws.Cells[13, participants.Length + 2] = "Shops and Inventory";
           createDataAreaMinigames(ws, 1, 1, "='{ID}'!G{Row}", 47,10);
           ws.Cells[18, 1] = "Percentage of time over THR"; ws.Cells[18, participants.Length + 2] = "Percentage of time over THR";
           ws.Cells[20, 1] = "Gekku Race"; ws.Cells[20, participants.Length + 2] = "Gekku Race";
           ws.Cells[21, 1] = "Dozo Quest"; ws.Cells[21, participants.Length + 2] = "Dozo Quest";
           ws.Cells[22, 1] = "Biri Brawl"; ws.Cells[22, participants.Length + 2] = "Biri Brawl";
           ws.Cells[23, 1] = "Bobo Ranch"; ws.Cells[23, participants.Length + 2] = "Bobo Ranch";
           ws.Cells[24, 1] = "Wiskin Defence"; ws.Cells[24, participants.Length + 2] = "Wiskin Defence";
           ws.Cells[25, 1] = "Pogi Pong"; ws.Cells[25, participants.Length + 2] = "Pogi Pong";
           ws.Cells[26, 1] = "Island"; ws.Cells[26, participants.Length + 2] = "Island";
           createDataAreaMinigames(ws, 18, 1, "='{ID}'!G{Row}", 59,7);

           //piegraph
           Microsoft.Office.Interop.Excel.ChartObjects chartObjs = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
           Microsoft.Office.Interop.Excel.ChartObject chartObj = chartObjs.Add(800, 20, 300, 300);
           Microsoft.Office.Interop.Excel.Chart avgMinInGamesPie = chartObj.Chart;
           string fromLetter = pickLetter(participants.Length+1);
           string toLetter = pickLetter(participants.Length + 2);
           Microsoft.Office.Interop.Excel.Range rg = ws.get_Range("=Minigames!$"+fromLetter+"$4:$"+toLetter+"$12");
           avgMinInGamesPie.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;
           avgMinInGamesPie.SetSourceData(rg, Type.Missing);
           avgMinInGamesPie.HasTitle = true;
           avgMinInGamesPie.ChartTitle.Text = "Average minutes in minigames ";
           avgMinInGamesPie.ApplyChartTemplate("participantTotalMinutesInMinigames.crtx");
           //bargraph
           Microsoft.Office.Interop.Excel.ChartObjects chartObjs4 = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
           Microsoft.Office.Interop.Excel.ChartObject chartObj4 = chartObjs4.Add(1100, 320, 300, 250);
           Microsoft.Office.Interop.Excel.Chart avgTimeoverTHR = chartObj4.Chart;
           avgTimeoverTHR.SetSourceData(rg, Type.Missing);
           avgTimeoverTHR.HasTitle = true;
           avgTimeoverTHR.ChartTitle.Text = "Average Time over THR ";
           avgTimeoverTHR.ApplyChartTemplate("participantPercentageOfTImeOverTHR.crtx");
           avgTimeoverTHR.Axes(Excel.XlAxisType.xlValue).MinimumScale = 50;
           //bar graph
           Microsoft.Office.Interop.Excel.ChartObjects chartObjs3 = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
           Microsoft.Office.Interop.Excel.ChartObject chartObj3 = chartObjs3.Add(800, 320, 300, 250);
           Microsoft.Office.Interop.Excel.Chart PercentoverTHR = chartObj3.Chart;
            rg = ws.get_Range("=Minigames!$" + fromLetter + "$18:$" + toLetter + "$26");
           PercentoverTHR.SetSourceData(rg, Type.Missing);
           PercentoverTHR.HasTitle = true;
           PercentoverTHR.ChartTitle.Text = "Percentage of Time over THR ";
           PercentoverTHR.ApplyChartTemplate("participantPercentageOfTImeOverTHR.crtx");
         

           //piegraph
           Microsoft.Office.Interop.Excel.ChartObjects chartObjs2 = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
           Microsoft.Office.Interop.Excel.ChartObject chartObj2 = chartObjs2.Add(1100, 20, 300, 300);
           Microsoft.Office.Interop.Excel.Chart totalMinsInGames = chartObj2.Chart;
          //  fromLetter = pickLetter(participants.Length + 1);
            toLetter = pickLetter(participants.Length + 4);
            rg = ws.get_Range("=Minigames!$"+fromLetter+"$4:$"+fromLetter+"$12,Minigames!$"+toLetter+"$4:$"+toLetter+"$12");
           totalMinsInGames.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;
           totalMinsInGames.SetSourceData(rg, Type.Missing);
           totalMinsInGames.HasTitle = true;
           totalMinsInGames.ChartTitle.Text = "Total minutes in minigames ";
           totalMinsInGames.ApplyChartTemplate("participantTotalMinutesInMinigames.crtx");
           }
           catch (Exception e)
           {
               Console.WriteLine("Creating the minigame sheet");
               Console.WriteLine(e);
           }

        }
          
        private static void createDataAreaMinigames(Excel.Worksheet ws, int topLeftRow, int topLeftCol, string formula, int formulaRowInc, int numberOfMiniGames)
        {
            try
            {


                
                // int numberOfMiniGames = 10;
                ws.Cells[topLeftRow + 1, participants.Length + 2 + topLeftCol] = "Average";
                ws.Cells[topLeftRow + 1, participants.Length + 3 + topLeftCol] = "StdDev";
                ws.Cells[topLeftRow + 1, participants.Length + 4 + topLeftCol] = "Total";
                ws.Cells[topLeftRow + numberOfMiniGames + 2, topLeftCol] = "Total";


                string temp = formula;
                int tempRow = formulaRowInc;
                for (int p = 0; p < participants.Length; p++)
                {
                    //set formulas
                    tempRow = formulaRowInc;

                    ws.Cells[topLeftRow + 1, topLeftCol + p + 1] = "P" + participants[p];
                    for (int i = 0; i < numberOfMiniGames; i++)
                    {
                        temp = temp.Replace("{ID}", participants[p].ToString());
                        //  temp = temp.Replace("{Letter}", pickLetter(p + 1));
                        temp = temp.Replace("{Row}", tempRow.ToString());
                        tempRow++;
                        //setFormulas in cells
                        ws.Cells[topLeftRow + i + 2, topLeftCol + p + 1].Formula = temp.ToString();
                        temp = formula;

                        //avg/stdev the participants
                        string gameFromCell = ((Excel.Range)ws.Cells[topLeftRow + 2 + i, topLeftCol + 1]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        string gameToCell = ((Excel.Range)ws.Cells[topLeftRow + 2 + i, participants.Length + topLeftCol]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        ws.Cells[topLeftRow + 2 + i, participants.Length + 2 + topLeftCol].Formula = "=AVERAGE(" + gameFromCell + ":" + gameToCell + ")";
                        ws.Cells[topLeftRow + 2 + i, participants.Length + 3 + topLeftCol].Formula = "=STDEV.P(" + gameFromCell + ":" + gameToCell + ")";
                        ws.Cells[topLeftRow + 2 + i, participants.Length + 4 + topLeftCol].Formula = "=SUM(" + gameFromCell + ":" + gameToCell + ")";

                    }
                    string participantFromCell = ((Excel.Range)ws.Cells[topLeftRow + 2, topLeftCol + p + 1]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    string participantToCell = ((Excel.Range)ws.Cells[topLeftRow + 1 + numberOfMiniGames, p + topLeftCol + 1]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    ws.Cells[topLeftRow + 2 + numberOfMiniGames, p + 1 + topLeftCol].Formula = "=SUM(" + participantFromCell + ":" + participantToCell + ")";


                }
             }catch(Exception e){
                 ((Excel.Application)ws.Parent.Parent).Quit();
            }
        }


        private static void createPedalingAndTHRSheet(Excel.Workbook newWB){
            try
            {
                Console.WriteLine("Creating sheet for pedaling and THR data");
                Excel.Worksheet ws = newWB.Sheets.Add();

                ws.Name = "Pedaling and THR";
                //create Data areas
                ws.Cells[1, 2] = "MINUTES PEDALING";
                createDataAreaTHRandPedal(ws, 1, 1, "='{ID}'!{Letter}42");
                ws.Cells[35, 2] = "MINUTES OVER THR";
                createDataAreaTHRandPedal(ws, 35, 1, "='{ID}'!{Letter}$9");

                //create Chart avg Minutes Pedaling
                Microsoft.Office.Interop.Excel.ChartObjects chartObjs = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
                Microsoft.Office.Interop.Excel.ChartObject chartObj = chartObjs.Add(500, 20, 300, 250);
                Microsoft.Office.Interop.Excel.Chart xlChart = chartObj.Chart;
                string toLetter = pickLetter(numberOfWeeksSoFar + 1);
                Microsoft.Office.Interop.Excel.Range rg = ws.get_Range("='Pedaling and THR'!$A$3:$A$" + (2 + participants.Length).ToString() + ",'Pedaling and THR'!$" + toLetter + "$3:$" + toLetter + "$" + (2 + participants.Length).ToString());
                xlChart.SetSourceData(rg, Type.Missing);
                xlChart.HasTitle = true;
                xlChart.ChartTitle.Text = "Average of minutes pedaling per participant";
                xlChart.ApplyChartTemplate("AvgMinPedal.crtx");

                Microsoft.Office.Interop.Excel.ChartObjects chartObjs2 = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
                Microsoft.Office.Interop.Excel.ChartObject chartObj2 = chartObjs2.Add(500, 270, 300, 250);
                Microsoft.Office.Interop.Excel.Chart xlChart2 = chartObj2.Chart;
                toLetter = pickLetter(numberOfWeeksSoFar);
                string row = (participants.Length + 3).ToString();
                rg = ws.get_Range("='Pedaling and THR'!$B$2:$" + toLetter + "$2,'Pedaling and THR'!$B$" + row + ":$" + toLetter + "$" + row);
                xlChart2.SetSourceData(rg, Type.Missing);
                xlChart2.HasTitle = true;
                xlChart2.ChartTitle.Text = "Average of minutes pedaling per week";
                xlChart2.ApplyChartTemplate("AvgMinPedal.crtx");


                Microsoft.Office.Interop.Excel.ChartObjects chartObjs3 = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
                Microsoft.Office.Interop.Excel.ChartObject chartObj3 = chartObjs3.Add(500, 770, 300, 250);
                Microsoft.Office.Interop.Excel.Chart xlChart3 = chartObj3.Chart;
                toLetter = pickLetter(numberOfWeeksSoFar + 1);
                row = (participants.Length + 36).ToString();
                rg = ws.get_Range("='Pedaling and THR'!$A$36:$A$" + row + ",'Pedaling and THR'!$" + toLetter + "$36:$" + toLetter + "$" + row);
                xlChart3.SetSourceData(rg, Type.Missing);
                xlChart3.HasTitle = true;
                xlChart3.ChartTitle.Text = "Average of minutes over THR per week";
                xlChart3.ApplyChartTemplate("AvgMinPedal.crtx");

                Microsoft.Office.Interop.Excel.ChartObjects chartObjs4 = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
                Microsoft.Office.Interop.Excel.ChartObject chartObj4 = chartObjs4.Add(500, 520, 300, 250);
                Microsoft.Office.Interop.Excel.Chart xlChart4 = chartObj4.Chart;
                toLetter = pickLetter(numberOfWeeksSoFar);
                row = (participants.Length + 36).ToString();
                rg = ws.get_Range("='Pedaling and THR'!$A$36:$" + toLetter + "$" + row);
                xlChart4.SetSourceData(rg, Type.Missing);
                xlChart4.HasTitle = true;
                xlChart4.ChartTitle.Text = "Minutes over THR per week";
                xlChart4.ApplyChartTemplate("minutesOverTHRperWeek.crtx");

                newWB.Save();
            }catch(Exception e){
                ((Excel.Application)newWB.Parent).Quit();
            }
        }




        //inputs week# ;returns the column letter that corresponds
        private static string pickLetter(int week)
        {
            switch (week)
            {
                case 1:
                    return "B";
                    
                case 2:
                    return "C";
                   
                case 3:
                    return "D";
                case 4:
                    return "E";
                case 5:
                    return "F";
                case 6:
                    return "G";
                case 7:
                    return "H";
                case 8:
                    return "I";
                case 9:
                    return "J";
                case 10:
                    return "K";
                case 11:
                    return "L";
                case 12:
                    return "M";
                case 13:
                    return "N";
                case 14:
                    return "O";
                case 15:
                    return "P";

            }
            return "";
        }

        //creates the participant sheets in the application by copying log files
        private static void copyLogInfoToResults(Excel.Application app, out Excel.Workbook newWB)
        {
            
           
            Excel.Workbook wb = app.ActiveWorkbook;
           
            Excel.Worksheet ws, ws2;
            string ResultsPath = reportsFolder+outFile;
           newWB  = app.Workbooks.Add();
            
            for (int i = 0; i < participants.Length; i++)
            {
                string id = participants[i];
                string firstPath = reportsFolder + id + "\\" + id + "_WeeklyReport";
                Console.WriteLine("Creating datasheet for participant"+id);


                try
                {
                    wb = app.Workbooks.Open(firstPath, false, false, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing);
                }
                catch (Exception e)
                {
                    
                    Console.WriteLine(e);
                    //wb.Close(false);
                    app.Quit();
                }

                //copy data
                ws = (Excel.Worksheet)wb.Worksheets[1];
                ws2 = newWB.Sheets.Add();
                ws2.Name = id;
                Excel.Range sourceRange = ws.get_Range(cell1PS, cell2PS);
                Excel.Range destinationRange = ws2.get_Range(cell1PS, cell2PS);

                sourceRange.Copy(Type.Missing);
                destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                sumPlayerSheets(ws2);
                averagePlayerSheets(ws2);
                participantSheetMinutesPerGamePie(ws2);
                participantSheetPercentoverTHR(ws2);
               newWB.Save();
              //  wb.Close(false);
            }
            try { 
              //   wb.Close(false);
                newWB.SaveAs(ResultsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
              
            }
            catch (Exception e)
            {

                ((Excel.Application)newWB.Parent).Quit();
            }
            }

            
            //Create teh sum of rows for participant sheets
           private static void sumPlayerSheets(Excel.Worksheet sheet) {

               foreach (DictionaryEntry label in rowsToSum)
               {
                   Excel.Range found = sheet.get_Range("A2", "A2").Find(label.Key.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, true, Type.Missing, Type.Missing);

                   if (!(found == null))
                   {
                       int Row_Number = found.Row;
                       sheet.Cells[Row_Number, numberOfWeeksSoFar + 2] = label.Value;
                       string fromCell = ((Excel.Range)sheet.Cells[Row_Number, 2]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
           Type.Missing, Type.Missing);
                       string toCell = ((Excel.Range)sheet.Cells[Row_Number, numberOfWeeksSoFar + 1]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
           Type.Missing, Type.Missing);

                      
                       sheet.Cells[Row_Number, numberOfWeeksSoFar + 3].Formula = "=SUM(" + fromCell + ":" + toCell + ")";
                       if (label.Key.ToString() == "Warmup Minutes")
                       {
                           sheet.Cells[Row_Number - 1, numberOfWeeksSoFar + 2] = "Total minutes spent in minigames";
                           cellMinsInMinigameFromPS = ((Excel.Range)sheet.Cells[Row_Number, numberOfWeeksSoFar + 2]);
                       }
                       else if (label.Key.ToString() == "Total time browsing shops and inventory")
                       {
                           cellMinsInMinigameToPS = ((Excel.Range)sheet.Cells[Row_Number, numberOfWeeksSoFar + 3]);
                      }

                   }
                   
               }
               


            }



           //Create the average for rows of player sheets
           private static void averagePlayerSheets(Excel.Worksheet sheet)
           {

               foreach (DictionaryEntry label in rowsToAvg)
               {
                   Excel.Range found = sheet.get_Range("A2", "A2").Find(label.Key.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, true, Type.Missing, Type.Missing);

                   if (!(found == null))
                   {
                       int Row_Number = found.Row;
                       sheet.Cells[Row_Number, numberOfWeeksSoFar + 2] = label.Value;
                       string fromCell = ((Excel.Range)sheet.Cells[Row_Number, 2]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
           Type.Missing, Type.Missing);
                       string toCell = ((Excel.Range)sheet.Cells[Row_Number, numberOfWeeksSoFar + 1]).get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
           Type.Missing, Type.Missing);


                       sheet.Cells[Row_Number, numberOfWeeksSoFar + 3].Formula = "=AVERAGE(" + fromCell + ":" + toCell + ")";
                       if (label.Key.ToString() == "% of Gekku Race time over THR")
                       {
                           (sheet.Cells[Row_Number-1, numberOfWeeksSoFar + 2]) = "Percentage of time over THR";
                           cellPercentoverTHRFromPS = ((Excel.Range)sheet.Cells[Row_Number, numberOfWeeksSoFar + 2]);
                       }
                       else if (label.Key.ToString() == "% of Island time over THR")
                       {
                           cellPercentoverTHRToPS = ((Excel.Range)sheet.Cells[Row_Number, numberOfWeeksSoFar + 3]);
                       }
                   }
               }


           }

        //creates pie graphs in participant sheets
           public static void participantSheetMinutesPerGamePie(Excel.Worksheet sheet)
           {
               Microsoft.Office.Interop.Excel.ChartObjects chartObjs = (Microsoft.Office.Interop.Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
               Microsoft.Office.Interop.Excel.ChartObject chartObj = chartObjs.Add(600, 20, 400, 400);
               Microsoft.Office.Interop.Excel.Chart xlChart = chartObj.Chart;
               Microsoft.Office.Interop.Excel.Range rg = sheet.get_Range(cellMinsInMinigameFromPS, cellMinsInMinigameToPS);
               xlChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;
               xlChart.SetSourceData(rg, Type.Missing);
               xlChart.HasTitle = true;
               xlChart.ChartTitle.Text = "Total Minutes in Minigames";
               xlChart.ApplyChartTemplate("participantTotalMinutesInMinigames.crtx");
              
           }

            //creates bar graph in participant sheets
           public static void participantSheetPercentoverTHR(Excel.Worksheet sheet)
           {
               Microsoft.Office.Interop.Excel.ChartObjects chartObjs = (Microsoft.Office.Interop.Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
               Microsoft.Office.Interop.Excel.ChartObject chartObj = chartObjs.Add(600, 500, 400, 400);

               Microsoft.Office.Interop.Excel.Chart xlChart = chartObj.Chart;
               Microsoft.Office.Interop.Excel.Range rg = sheet.get_Range(cellPercentoverTHRFromPS, cellPercentoverTHRToPS);
               xlChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered;
               xlChart.SetSourceData(rg, Type.Missing);
               xlChart.HasTitle = true;
               xlChart.ChartTitle.Text = "Percent of Time over THR";
               xlChart.ApplyChartTemplate("participantPercentageOfTImeOverTHR.crtx");

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
                    catch (ArgumentException)
                    {//line empty-clear paramater value so no variable is overwritten
                        parameter = (Parameters)Enum.Parse(typeof(Parameters), "none");
                    }
                    try
                    {
                        switch (parameter)
                        {
                            case Parameters.participants:
                                participants = lineData[1].Split(',');
                                break;
                            case Parameters.reportsOutputPath:
                                reportsFolder = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)+ lineData[1].Trim();
                                break;
                            case Parameters.numberOfWeeksSoFar:
                                numberOfWeeksSoFar = System.Convert.ToInt32(lineData[1].Trim());
                                break;
                            case Parameters.outFile:
                                outFile = lineData[1];
                                break;
                            default:
                                //parameter is unknown - do nothing
                                break;
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
               
                sr.Close();
                //set range of cells to copy for each participant
                cell1PS = "A1";
                int numberOfRows = 90;//changes if new rows are added
                cell2PS = pickLetter(numberOfWeeksSoFar) + numberOfRows.ToString();
            }
        }
    
}
