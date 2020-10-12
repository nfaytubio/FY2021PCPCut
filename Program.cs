using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Regex = System.Text.RegularExpressions.Regex;
using System.IO;
using System.IO.Compression;


namespace attempt2
{
    class Program
    {
        static void copypaste(Excel.Application excelapp, Excel.Workbook source, Excel.Workbook destination, string worksheetname, string filter, int filtercolumn, string lastcolumn)
        {
            Excel.Worksheet sourceworksheet = source.Worksheets[worksheetname];

            sourceworksheet.Copy(destination.Worksheets[1]);

            Excel.Worksheet destinationworksheet = destination.Worksheets[worksheetname];

            long rows = sourceworksheet.UsedRange.Rows.Count;

            destinationworksheet.Rows["2:" + rows + 1].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

            string filterrange = "A1:" + lastcolumn + rows;
            string copyrange = "A2:" + lastcolumn + rows;

            try
            {
                sourceworksheet.Range[filterrange].AutoFilter(filtercolumn, "=*" + filter + "*", Excel.XlAutoFilterOperator.xlAnd);
                sourceworksheet.Range[copyrange].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                destinationworksheet.Range[copyrange].PasteSpecial();
                //destinationworksheet.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }

            catch (Exception)
            {
                destinationworksheet.Delete();
                Console.WriteLine("Cannot filter " + worksheetname + ", sheet removed for "+ destination.Path);
            }
        }

        static void Main(string[] args)
        {
            //static variables
            //only need 1 tmTerritory to cut both sheets
            List<string> tmTerritory = new List<string>() { "21101-Long Island-KYLE PROVOST", "21102-Newark-KYLE RITTER", "21103-New Jersey NW-BRANDON CONNER", "21104-Philadelphia-JEFFREY KARCZEWSKI", "21105-DCBaltimore-DUSTIN HARTWIGSEN", "21106-Harrisburg-CHRISTINA STENCOVAGE", "21107-Hartford-FRANCIS REGALA", "21201-Charlotte-HILLARY WILSON", "21204-Wilmington-JOSEPH CURTIS", "21206-Raleigh-TODD SUTTON", "21302-Baton Rouge-JESSICA CARD", "21303-Alabama-SHELBY ANDERSON", "21305-Orlando-TAYLOR ROSSMOORE", "21306-Tampa-KATIE JARVIS", "21307-Atlanta-TYLER EARLEY", "21301-Shreveport-TINA PHILLIPS", "21402-San Antonio-KELSEY LOCKLEAR", "21403-Fort Worth-BROOKE SHOEMAKER", "21404-Los Angeles-AMANDA CRAVEN", "21405-San Jose-NICKI DERKSEN", "21406-South Texas-MICHAEL DE LOS SANTOS", "21407-Dallas-TREY GOODCHILD", "21408-Houston-MCKENZIE BRUNNEMANN", "11104-Northern PA-MISSY PERRY", "11106-Carolinas-KAZON NEELY", "11108-New Jersey-DANIELLE KARACICA", "11202-South Florida-MICHAEL WOLSTENCROFT", "11211-St Louis-CORBIN SELLERS", "11303-Central Texas-MARC GALLAGHER", "11304-Houston-AMY EVANICH", "11306-North Texas-BOBBI LEWIS", "11311-SocalVegas-MACKENZIE HALL", "11410-SocalVegas-MACKENZIE HALL"};
            List<string> rsdRegion = new List<string>() { "Northeast", "Southeast", "West", "Specialty"};

            //List<string> copyworksheets = new List<string>() { "Aciphex Targets_National", "Karbinal ER Targets_National", "PVFTVF Targets_National", "Tuzistra XR Targets_National", "Zolpimist Targets_National", "Natesto Targets_National" };
            //List<string> copycolumns = new List<string>() { "V", "V", "W", "T", "Y", "AD" };

            var copyworksheetcolumns = new Dictionary<string, string>() {
            {"Aciphex Targets","V"},
            {"Karbinal ER Targets","V"},
            {"PVFTVF Targets","W"},
            {"Tuzistra XR Targets","T"},
            {"Zolpimist Targets","Y"},
            {"Natesto Targets","AD"}};

            List<string> staticRSD = new List<string>() { "RSD Territory Summary","RSD Region Summary", "Table of Contents" };
            //List<string> staticWSRSD = new List<string>() { "Q3 NonReportingICAdj", "MBO_PCPTeamQuarterSummary", "MBO - Sean Mangelson", "TMOQ", "TMOHY", "KeyAssumptionsDefs", "RSD Payout Rules", "blank" };
            //List<string> staticWSNSD = new List<string>() { "Q3 NonReportingICAdj", "MBO_PCPTeamQuarterSummary", "MBO - Sean Mangelson", "TMOQ", "TMOHY", "KeyAssumptionsDefs", "NSD Payout Rules", "blank", "Q4 20 ICTerritoryProductSummary", "Q4 20 IC Territory Summary" };
            string rawDate = DateTime.Today.ToShortDateString();
            string cleanDate = Regex.Replace(rawDate, "[^A-Za-z0-9 ]", "");
            string JDScorecard = @"C:\Users\nfisher\Documents\Target List Cuts\National - Top 100 Report - July 2020.xlsx";


            foreach (var territory in tmTerritory)
            {
                Excel.Application excelapp = new Excel.Application();
                excelapp.Visible = false;

                excelapp.DisplayAlerts = false;

                Excel.Workbook wb1 = excelapp.Workbooks.Open(JDScorecard);

                Excel.Workbook output = excelapp.Workbooks.Add();

                foreach (var kvp in copyworksheetcolumns)
                {
                    copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: kvp.Key, filter: territory, filtercolumn: 2, lastcolumn: kvp.Value);
                }

                output.Worksheets["Sheet1"].Delete();
                wb1.Worksheets["Table of Contents"].Copy(output.Worksheets[1]);
                output.Worksheets["Table of Contents"].Range["B2"].Select();

                string path = @"C:\Users\nfisher\Documents\Target List Cuts\Top 100\TM\";
                string workbookpath = path + territory + " TM - Top 100 Report - July 2020 " + cleanDate + ".xlsx";

                if (!File.Exists(workbookpath))
                {
                    output.SaveAs(workbookpath);
                }
                object misValue = System.Reflection.Missing.Value;

                wb1.Close(false, misValue, misValue);
                output.Close(false, misValue, misValue);
                excelapp.Quit();

                //worksheet variables
                //make this into a function on refactor
                //Excel.Worksheet NatestoTargetsNational = wb1.Worksheets["Natesto Targets_National"];
                //Excel.Worksheet ZolpimistTargetsNational = wb1.Worksheets["Zolpimist Targets_National"];
                //Excel.Worksheet TuzistraTargetsNational = wb1.Worksheets["Tuzistra XR Targets_National"];
                //Excel.Worksheet PVFTVFTargetsNational = wb1.Worksheets["PVFTVF Targets_National"];
                //Excel.Worksheet KarbinalERTargetsNational = wb1.Worksheets["Karbinal ER Targets_National"];
                //Excel.Worksheet AciphexTargetsNational = wb1.Worksheets["Aciphex Targets_National"];

                ////make this into a function on refactor
                //long natrows = NatestoTargetsNational.UsedRange.Rows.Count;
                //long zolrows = ZolpimistTargetsNational.UsedRange.Rows.Count;
                //long tuzrows = TuzistraTargetsNational.UsedRange.Rows.Count;
                //long ptrows = PVFTVFTargetsNational.UsedRange.Rows.Count;
                //long karrows = KarbinalERTargetsNational.UsedRange.Rows.Count;
                //long acirows = AciphexTargetsNational.UsedRange.Rows.Count;

                //Excel.Workbook output = excelapp.Workbooks.Add();

                ////foreach (string ws in staticWSTM)
                ////{
                ////    wb1.Worksheets[ws].Copy(output.Worksheets[1]);
                ////}

                ////make this into a function on refactor
                //NatestoTargetsNational.Copy(output.Worksheets[1]);
                //ZolpimistTargetsNational.Copy(output.Worksheets[1]);
                //TuzistraTargetsNational.Copy(output.Worksheets[1]);
                //PVFTVFTargetsNational.Copy(output.Worksheets[1]);
                //KarbinalERTargetsNational.Copy(output.Worksheets[1]);
                //AciphexTargetsNational.Copy(output.Worksheets[1]);

                ////output worksheet variables
                //Excel.Worksheet NatestoTargetsNationalOut = output.Worksheets["Natesto Targets_National"];
                //Excel.Worksheet ZolpimistTargetsNationalOut = output.Worksheets["Zolpimist Targets_National"];
                //Excel.Worksheet TuzistraTargetsNationalOut = output.Worksheets["Tuzistra XR Targets_National"];
                //Excel.Worksheet PVFTVFTargetsNationalOut = output.Worksheets["PVFTVF Targets_National"];
                //Excel.Worksheet KarbinalERTargetsNationalOut = output.Worksheets["Karbinal ER Targets_National"];
                //Excel.Worksheet AciphexTargetsNationalOut = output.Worksheets["Aciphex Targets_National"];

                //NatestoTargetsNationalOut.Rows["2:" + natrows + 1].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                ////NatestoTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                ////output.Worksheets["Q4 20 IC Territory Summary"].Rows["5:60"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                //ZolpimistTargetsNationalOut.Rows["2:" + zolrows + 1].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                ////ZolpimistTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                //TuzistraTargetsNationalOut.Rows["2:" + tuzrows + 1].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                ////TuzistraTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                ////PVFTVFTargetsNationalOut.Rows["2:" + ptrows + 1].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                ////PVFTVFTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                //KarbinalERTargetsNationalOut.Rows["2:" + karrows + 1].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                ////KarbinalERTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                //AciphexTargetsNationalOut.Rows["2:" + acirows + 1].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                ////AciphexTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                ////make this into a function on refactor
                ////ICTerritoryProductSummary.Outline.ShowLevels(0, 2);

                ////natesto copy paste
                //string natrange = "A2:AD" + natrows;

                //NatestoTargetsNational.Range[natrange].AutoFilter(2, "=*" + territory + "*", Excel.XlAutoFilterOperator.xlAnd);
                //NatestoTargetsNational.Range["A3:AD" + natrows].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                //NatestoTargetsNationalOut.Range["A3:AD" + natrows].PasteSpecial();
                //NatestoTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                ////zolpimist product copy paste
                //string zolrange = "A2:Y" + zolrows;

                //ZolpimistTargetsNational.Range[zolrange].AutoFilter(2, "=*" + territory + "*", Excel.XlAutoFilterOperator.xlAnd);
                //ZolpimistTargetsNational.Range["A3:Y" + zolrows].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                //ZolpimistTargetsNationalOut.Range["A3:Y" + zolrows].PasteSpecial();
                //ZolpimistTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                ////tuzistra product copy paste
                //string tuzrange = "A2:T" + tuzrows;

                //TuzistraTargetsNational.Range[tuzrange].AutoFilter(2, "=*" + territory + "*", Excel.XlAutoFilterOperator.xlAnd);
                //TuzistraTargetsNational.Range["A3:T" + tuzrows].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                //TuzistraTargetsNationalOut.Range["A3:T" + tuzrows].PasteSpecial();
                //TuzistraTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                ////pvftvf product copy paste

                //try
                //{
                //    string ptrange = "A2:W" + ptrows;


                //    PVFTVFTargetsNational.Range[ptrange].AutoFilter(2, "=*" + territory + "*", Excel.XlAutoFilterOperator.xlAnd);
                //    PVFTVFTargetsNational.Range["A3:W" + ptrows].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                //    PVFTVFTargetsNationalOut.Range["A3:W" + ptrows].PasteSpecial();
                //    PVFTVFTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                //}

                //catch (Exception)
                //{
                //    PVFTVFTargetsNationalOut.Delete();
                //    Console.WriteLine("Cannot filter, sheet removed.");
                //}


                ////karbinal product copy paste
                //try
                //{
                //    string karrange = "A2:V" + karrows;

                //    KarbinalERTargetsNational.Range[karrange].AutoFilter(2, "=*" + territory + "*", Excel.XlAutoFilterOperator.xlAnd);
                //    KarbinalERTargetsNational.Range["A3:V" + karrows].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                //    KarbinalERTargetsNationalOut.Range["A3:V" + karrows].PasteSpecial();
                //    KarbinalERTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                //}

                //catch (Exception)
                //{
                //    KarbinalERTargetsNationalOut.Delete();
                //    Console.WriteLine("Cannot filter, sheet removed.");
                //}

                ////aciphex product copy paste
                //try
                //{
                //    string acirange = "A2:V" + acirows;

                //    AciphexTargetsNational.Range[acirange].AutoFilter(2, "=*" + territory + "*", Excel.XlAutoFilterOperator.xlAnd);
                //    AciphexTargetsNational.Range["A3:V" + acirows].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                //    AciphexTargetsNationalOut.Range["A3:V" + acirows].PasteSpecial();
                //    AciphexTargetsNationalOut.Rows["2"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                //}

                //catch (Exception)
                //{
                //    AciphexTargetsNationalOut.Delete();
                //    Console.WriteLine("Cannot filter, sheet removed.");
                //}

                //collapse groups
                //output.Worksheets["Q4 20 ICTerritoryProductSummary"].Outline.ShowLevels(0, 1);
            }

            foreach (var region in rsdRegion)
            {
                Excel.Application excelapp = new Excel.Application();
                excelapp.Visible = false;

                excelapp.DisplayAlerts = false;

                Excel.Workbook wb1 = excelapp.Workbooks.Open(JDScorecard);

                Excel.Workbook output = excelapp.Workbooks.Add();

                foreach (var kvp in copyworksheetcolumns)
                {
                    copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: kvp.Key, filter: region, filtercolumn: 1, lastcolumn: kvp.Value);
                }

                output.Worksheets["Sheet1"].Delete();

                foreach (string ws in staticRSD)
                {
                    wb1.Worksheets[ws].Copy(output.Worksheets[1]);
                }

                output.Worksheets["Table of Contents"].Range["B2"].Select();

                string path = @"C:\Users\nfisher\Documents\Target List Cuts\Top 100\RSD\";
                string workbookpath = path + region + " RSD - Top 100 Report - July 2020 " + cleanDate + ".xlsx";

                if (!File.Exists(workbookpath))
                {
                    output.SaveAs(workbookpath);
                }
                object misValue = System.Reflection.Missing.Value;

                wb1.Close(false, misValue, misValue);
                output.Close(false, misValue, misValue);
                excelapp.Quit();
            }

            //foreach (var sd in nsd)
            //{
            //    Excel.Application excelapp = new Excel.Application();
            //    excelapp.Visible = false;

            //    excelapp.DisplayAlerts = false;

            //    Excel.Workbook wb1 = excelapp.Workbooks.Open(Q420ICScorecard);

            //    Excel.Workbook output = excelapp.Workbooks.Add();

            //    foreach (string ws in staticWSNSD)
            //    {
            //        wb1.Worksheets[ws].Copy(output.Worksheets[1]);
            //    }

            //    output.Worksheets["Sheet1"].Delete();
            //    output.Worksheets["Q4 20 IC Territory Summary"].Select();

            //    string path = @"C:\Users\nfisher\Documents\Target List Cuts\IC\";
            //    string workbookpath = path + sd + "_" + "Q420IC" + "_Scorecard_v" + cleanDate + ".xlsx";

            //    if (!File.Exists(workbookpath))
            //    {
            //        output.SaveAs(workbookpath);
            //    }

            //    object misValue = System.Reflection.Missing.Value;

            //    wb1.Close(false, misValue, misValue);
            //    output.Close(false, misValue, misValue);
            //    excelapp.Quit();

            //}

            /*lastrow = xs1.Rows.End[Excel.XlDirection.xlDown];
            lastcol = xs1.Columns.End[Excel.XlDirection.xlToRight];
            xs1.Cells[lastrow.Row, lastcol.Column]*/

            //foreach (string team in tmTerritory)
            //{
            //    string path = @"C:\Users\nfisher\Documents\Target List Cuts\IC\" + team;
            //    if (!Directory.Exists(path))
            //    {
            //        Directory.CreateDirectory(path);
            //    }
            //}

            //ONLY USED FOR COPYING TERRITORY PNGS

            //var tregions = teams.Zip(territories, (a, b) => new { team = a, territory = b });

            //foreach (var item in tregions)
            //{
            //    int dash = item.territory.IndexOf("-");
            //    int dash2 = item.territory.IndexOf("-",dash+1);

            //    string sourceFile = @"C:\Users\nfisher\Documents\Target List Cuts\Maps\"+item.territory.Substring(0, dash2)+".png";
            //    string destinationFile = @"C:\Users\nfisher\Documents\Target List Cuts\"+item.team+ @"\"+item.territory.Substring(0, dash2) + ".png";

            //    // To move a file or folder to a new location:
            //    System.IO.File.Copy(sourceFile, destinationFile, true);

            //}

            //foreach (string team in teams)
            //{
            //    string startPath = @"C:\Users\nfisher\Documents\Target List Cuts\" + team;
            //    string zipPath = @"C:\Users\nfisher\Documents\Target List Cuts\Q121_" + team + "TeamPlanningFile.zip";

            //    if (!File.Exists(zipPath))
            //    {
            //        ZipFile.CreateFromDirectory(startPath, zipPath);
            //    }
            //}
        }

        /*static Excel Worksheet ExcelWorksheet(string[] args)
        {
            
        }*/
    }
}
