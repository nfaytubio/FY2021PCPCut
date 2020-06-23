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
        static void Main(string[] args)
        {
            //static variables
            List<string> teams = new List<string>() { "Specialty", "Specialty", "Specialty", "Specialty", "Specialty", "Specialty", "Specialty", "Specialty", "Specialty", "Northeast", "Northeast", "Northeast", "Northeast", "Northeast", "Northeast", "Northeast", "Southeast", "Southeast", "Southeast", "Southeast", "Southeast", "Southeast", "Southeast", "Southeast", "West", "West", "West", "West", "West", "West", "West", "West" };
            List<string> regions = new List<string>() { "Carolinas", "Central Texas", "Houston", "New Jersey", "North Texas", "Northern PA", "Socal/Vegas", "South Florida", "St Louis", "Long Island", "Newark", "New Jersey NW", "Philadelphia", "DC/Baltimore", "Harrisburg", "Hartford", "Charlotte", "Wilmington", "Raleigh", "Baton Rouge", "Alabama", "Orlando", "Tampa", "Atlanta", "Shreveport", "San Antonio", "Fort Worth", "Los Angeles", "San Jose", "South Texas", "Dallas", "Houston" };
            List<string> ids = new List<string>() { "Carolinas", "Central Texas", "Houston", "New Jersey", "North Texas", "Northern PA", "Socal/Vegas", "South Florida", "St Louis", "21101", "21102", "21103", "21104", "21105", "21106", "21107", "21201", "21204", "21206", "21302", "21303", "21305", "21306", "21307", "21301", "21402", "21403", "21404", "21405", "21406", "21407", "21408" };
            string rawDate = DateTime.Today.ToShortDateString();
            string cleanDate = Regex.Replace(rawDate, "[^A-Za-z0-9 ]", "");
            string specialty = @"C:\Users\nfisher\Documents\Target List Cuts\SpecialtyNatestoZolpimist.xlsx";
            string natforecast = @"C:\Users\nfisher\Documents\Target List Cuts\Fiscal21NationalForecastv4.xlsx";
                      

            /*lastrow = xs1.Rows.End[Excel.XlDirection.xlDown];
            lastcol = xs1.Columns.End[Excel.XlDirection.xlToRight];
            xs1.Cells[lastrow.Row, lastcol.Column]*/
            
            foreach (string team in teams)
            {
                string path = @"C:\Users\nfisher\Documents\Target List Cuts\" + team;
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }

            var tregions = teams.Zip(regions, (a, b) => new { team = a, region = b });

            foreach (var item in tregions)
            {
                Excel.Application excelapp = new Excel.Application();
                excelapp.Visible = false;

                Excel.Workbook wb1 = excelapp.Workbooks.Open(specialty);
                Excel.Workbook wb2 = excelapp.Workbooks.Open(natforecast);

                //worksheet variables
                Excel.Worksheet natestotargetlist = wb1.Worksheets["NatestoTargetList"];
                Excel.Worksheet zolpimisttargetlist = wb1.Worksheets["ZolpimistTargetList"];
                Excel.Worksheet targetlistfull = wb1.Worksheets["TargetListFull"];

                long ntlrows = natestotargetlist.UsedRange.Rows.Count;
                long ztlrows = zolpimisttargetlist.UsedRange.Rows.Count;
                long tgfrows = targetlistfull.UsedRange.Rows.Count;

                Excel.Workbook output = excelapp.Workbooks.Add();

                natestotargetlist.Copy(output.Worksheets[1]);
                zolpimisttargetlist.Copy(output.Worksheets[1]);

                output.Worksheets["NatestoTargetList"].Rows["9:" + ntlrows].ClearContents();
                output.Worksheets["ZolpimistTargetList"].Rows["7:" + ztlrows].ClearContents();

                natestotargetlist.Outline.ShowLevels(0, 2);
                zolpimisttargetlist.Outline.ShowLevels(0, 2);

                targetlistfull.Range["A2:DC" + tgfrows].AutoFilter(1, item.region);

                targetlistfull.Range["A3:I" + tgfrows].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                //need to specify actual paste range here, B9 only pastes cells to row B
                output.Worksheets["NatestoTargetList"].Range["B9"].PasteSpecial();

                output.Worksheets["NatestoTargetList"].Outline.ShowLevels(0, 1);
                output.Worksheets["ZolpimistTargetList"].Outline.ShowLevels(0, 1);

                output.Worksheets["Sheet1"].Delete();

                string path = @"C:\Users\nfisher\Documents\Target List Cuts\" + item.team;
                string workbookpath = path + @"\" + item.region + "_" + item.team + "Team" + "_PlanningFile_v" + cleanDate + ".xlsx";

                if (!File.Exists(workbookpath))
                {
                    output.SaveAs(workbookpath);
                }
                object misValue = System.Reflection.Missing.Value;

                wb1.Close(false, misValue, misValue);
                wb2.Close(false, misValue, misValue);
                output.Close(false, misValue, misValue);
                excelapp.Quit();
            }

            foreach (string team in teams)
            {
                string startPath = @"C:\Users\nfisher\Documents\Target List Cuts\" + team;
                string zipPath = @"C:\Users\nfisher\Documents\Target List Cuts\Q121_" + team + "TeamPlanningFile.zip";

                if (!File.Exists(zipPath))
                {
                    ZipFile.CreateFromDirectory(startPath, zipPath);
                }
            }
        }

        /*static void CutExcel(string[] args)
        {
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = false;

            Excel.Workbook wb1 = excelapp.Workbooks.Open(specialty);
            Excel.Workbook wb2 = excelapp.Workbooks.Open(natforecast);

            //worksheet variables
            Excel.Worksheet natestotargetlist = wb1.Worksheets["NatestoTargetList"];
            Excel.Worksheet zolpimisttargetlist = wb1.Worksheets["ZolpimistTargetList"];
            Excel.Worksheet targetlistfull = wb1.Worksheets["TargetListFull"];

            long ntlrows = natestotargetlist.UsedRange.Rows.Count;
            long ztlrows = zolpimisttargetlist.UsedRange.Rows.Count;
            long tgfrows = targetlistfull.UsedRange.Rows.Count;

            Excel.Workbook output = excelapp.Workbooks.Add();

            natestotargetlist.Copy(output.Worksheets[1]);
            zolpimisttargetlist.Copy(output.Worksheets[1]);

            output.Worksheets["NatestoTargetList"].Rows["9:" + ntlrows].ClearContents();
            output.Worksheets["ZolpimistTargetList"].Rows["7:" + ztlrows].ClearContents();

            natestotargetlist.Outline.ShowLevels(0, 2);
            zolpimisttargetlist.Outline.ShowLevels(0, 2);

            targetlistfull.Range["A2:DC" + tgfrows].AutoFilter(2, item.team);

            targetlistfull.Range["A3:I64263"].Copy();
            output.Worksheets["NatestoTargetList"].PasteSpecial(Excel.XlCellType.xlCellTypeVisible);

            string path = @"C:\Users\nfisher\Documents\Target List Cuts\" + item.team;
            string workbookpath = path + @"\" + item.region + "_" + item.team + "Team" + "_PlanningFile_v" + cleanDate + ".xlsx";

            output.SaveAs(workbookpath);


            output.Close();
        }*/
    }
}
