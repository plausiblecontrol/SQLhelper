using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smarp_o_Matic_XL_Gold_Premium {
  class Program {

    static void Main(string[] args) {

      DateTime newFile = File.GetLastWriteTime(@"\\file\Engineering\Systems_Engineering\Jason\SMARP-o-matic\Bin\New\New_Smarp-o-Matic_XL-Gold-Premium.exe");
      if (newFile > Convert.ToDateTime("6/17/15")) {
        Console.WriteLine("There is a new SMARP tool available, please use that one.");
        Console.WriteLine(@"\\SMARP-o-matic\Bin\New\");
        return;
      }

      string user = System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName.Split(',')[1].Replace(" ", "") + " " + System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName.Split(',')[0];
      string templateFile = Directory.GetCurrentDirectory() + "\\Sprint_(Suite)_(Release).xlsx";
      List<string> validSuites = new List<string>() { "Platform Server", "Platform Visualization", "System Management", "SCADA", "COMMS", "ISR", "GMS", "EMS", "DMS", "Gas" };
      
      if (args.Length < 1 || args.Length > 2) {
        Console.WriteLine("Provide proper suite name.");
        Console.WriteLine("<suite name> [<different reference date (defaults to today)>]");
        return;
      }


      DateTime requestedDate = DateTime.Now;
      if (args.Length == 2) {
        try { requestedDate = Convert.ToDateTime(args[1]); } catch {
          Console.WriteLine("Enter a valid date, or none at all");
          return;
        }
      }
 

      string thisSuite = args[0];

      if (thisSuite == "-?" || thisSuite == "?" || thisSuite == "help" || thisSuite == "-help" || thisSuite == "-h") {
        Console.WriteLine("<suite name> [<different reference date (defaults to today)>]");
        return;
      }


      if (!validSuites.Contains(thisSuite)) {
        Console.WriteLine("{0} is not a valid suite", thisSuite);
        Console.WriteLine("{0}", string.Join(", ", validSuites.ToArray()));
        return;
      }

      List<string> Exceptions = new List<string>();
      switch (thisSuite) {
        case "Platform Visualization":
          Exceptions.Add("Voyager");
          Exceptions.Add("Voyager Toolkit");
          Exceptions.Add("OSI Expedition");
          break;
        case "System Management":
          Exceptions.Add("Security Profiler");
          Exceptions.Add("monarchNET Tweak");
          break;
        case "ISR":
          Exceptions.Add("CHRONUS HRS");
          break;
        case "GMS":
          Exceptions.Add("Enerprise ITK");
          Exceptions.Add("OpenSVC");
          break;
        case "DMS":
          Exceptions.Add("Continua WSM");
          Exceptions.Add("Spectra OMS");
          break;
      }

      //initialize the query
      SuiteQry DBCon = new SuiteQry();
      //did not oo it b/c was doing quick, could clean up the array references to make it easier to understand
      List<string[]> suiteAll = DBCon.getSuites(thisSuite, requestedDate);
      if (suiteAll.Count == 0) {
        Console.WriteLine("No suite releases found in the last 14 days. Try adding an earlier reference date.");
          return;
      }
      //find previous suites for each suite in array (array is sorted by suite)
      //compare product releases between 2 suites
        //find products that have changed
          //get all versions of that product
        


      
      Excel.Application excelApp = null;
      Excel.Workbook NewBook = null;
      Excel.Worksheet NewSheet = null;
      Excel.Range xlR = null;
      excelApp = new Excel.Application();
      excelApp.DisplayAlerts = false;
      
      //1 spreadsheet for each suite found in last 14 days - COMMS 3.2.4, COMMS 1.2.3, and COMMS 3.3.5
      for (int x = 0; x < suiteAll.Count; x++) {//for each product sorted by suite
        NewBook = excelApp.Workbooks.Open(templateFile, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        NewSheet = NewBook.Sheets[2];
        string version = suiteAll[x][2] + "." + suiteAll[x][3] + "." + suiteAll[x][4] + "." + suiteAll[x][5];
        string exitFile = "Sprint_" + thisSuite.Replace(" ", "_") + "_" + version + ".xlsx";
        
        //add some new columns to put release info
        xlR = NewSheet.get_Range("F3", "H3").EntireColumn;
        xlR.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
        object[,] newHeaders = xlSafe(new string[] { "Investigation Item / Notes / Comments / Requirements", "Release Notes", "Engineering Notes", "Developer Notes" });
        NewSheet.get_Range("E3", "H3").Value2 = newHeaders;
        int c = 0;
        while (version == suiteAll[x][2] + "." + suiteAll[x][3] + "." + suiteAll[x][4] + "." + suiteAll[x][5]) {//while same suite
          //ugly array reference that could be cleaned up a bit
          //SuiteReleaseDate0	SName1	MajorRelease2	MajorRevision3	MinorRevision4	Patch5	ProductName6	PReleaseDate7	Release8	Requirements9	ReleaseNotes10	PENotes11	DeveloperNotes12 main/patch13
          object[,] xlLine = xlSafe(new string[] { user, suiteAll[x][6], suiteAll[x][8], suiteAll[x][9], suiteAll[x][10], suiteAll[x][11], suiteAll[x][12] });
          NewSheet.Range["B" + (4 + c), "H" + (4 + c)].Value2 = xlLine;
          string reportURL = "http://reports:8091/ReportServer/Pages/ReportViewer.aspx?http%3a%2f%2fosiinet%2freports%2freportslibrary%2fSoftware+Development%2fissued_release_notes.rdl&Product=" + suiteAll[x][6] + "&Release=" + suiteAll[x][8];
          NewSheet.Hyperlinks.Add(NewSheet.Range["D" + (4 + c), Type.Missing], reportURL, Type.Missing, "Report Page", suiteAll[x][8]);
          x++;
          c++;
          if (x == suiteAll.Count)
            break;
        }
        x--;
        NewSheet.Cells[2, 2] = "SMARP-O-MATIC"; //B2

        try {
          NewBook.SaveAs(Directory.GetCurrentDirectory() + "\\" + exitFile, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
        } catch {
          //Could try to do a name change? Or fail nicely? I don't like having to close ALL Excel documents like Andrew implemented originally...
          Console.WriteLine("File {0} is already opened and could not be overwritten - close it and re-run this tool", exitFile);
          continue;
        }
        NewBook.Close(false, Type.Missing, Type.Missing);
      }



      if (Exceptions.Count > 0) {
        NewBook = excelApp.Workbooks.Open(templateFile, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        NewSheet = NewBook.Sheets[2];
        string exitFile = "Sprint_" + thisSuite.Replace(" ", "_") + "_Exceptions.xlsx";
        
        //insert 3 new columns @ F to be used as the different release note fields.
        xlR = NewSheet.get_Range("F3", "H3").EntireColumn;
        xlR.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
        object[,] newHeaders = xlSafe(new string[] { "Investigation Item / Notes / Comments / Requirements", "Release Notes", "Engineering Notes", "Developer Notes" });
        NewSheet.get_Range("E3", "H3").Value2 = newHeaders;
        for (int i = 0; i < Exceptions.Count; i++) {
          //string[] prodNotes = DBCon.getExceptions(Exceptions[i], releaseTime);
          string[] prodNotes = DBCon.getExceptions(Exceptions[i], requestedDate.ToString("M/d/yyyy"));
          if (prodNotes[0] == Exceptions[i]) {
            //prodname, release, relNotes, eNotes, devNotes, custNotes
            object[,] xlLine = xlSafe(new string[] { user, prodNotes[0], prodNotes[1], prodNotes[5], prodNotes[2], prodNotes[3], prodNotes[4] });
            NewSheet.Range["B" + (4 + i), "H" + (4 + i)].Value2 = xlLine;
            string reportURL = "http://reports:8091/ReportServer/Pages/ReportViewer.aspx?http%3a%2f%2fosiinet%2freports%2freportslibrary%2fSoftware+Development%2fissued_release_notes.rdl&Product=" + prodNotes[0] + "&Release=" + prodNotes[1];
            NewSheet.Hyperlinks.Add(NewSheet.Range["D" + (4 + i), Type.Missing], reportURL, Type.Missing, "Report Page", prodNotes[1]);
          } else {
            object[,] xlLine = xlSafe(new string[] { user, Exceptions[i], "No changes", "Same release as previous sprint", "", "", "" });
            NewSheet.Range["B" + (4 + i), "H" + (4 + i)].Value2 = xlLine;
          }
        }

        //provide more information
        NewSheet.Cells[2, 2] = "SMARP-O-MATIC"; //B2
        NewSheet.Cells[2, 3] = "Latest release in last 2 weeks";

        try {
          NewBook.SaveAs(Directory.GetCurrentDirectory() + "\\" + exitFile, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
        } catch {
          //Could try to do a name change? Or fail nicely? I don't like having to close ALL Excel documents like Andrew implemented originally...
          Console.WriteLine("File {0} is already opened and could not be overwritten - close it and re-run this tool", exitFile);
          //continue;
        }
        NewBook.Close(false, Type.Missing, Type.Missing);
      }


      try { excelApp.Quit(); } catch { }
      //cleanup Excel objects
      releaseObject(xlR);
      releaseObject(NewSheet);
      releaseObject(NewBook);
      releaseObject(excelApp);

    }

    static void releaseObject(object obj) {
      try {//MS COM objects are dumb
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      } catch {
        obj = null;
      } finally {
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
    }
    static object[,] xlSafe(string[] line) {
      object[,] row = new object[1, line.Length];
      for (int i = 0; i < line.Length; i++) {
        row[0, i] = line[i];
      }
      return row;
    }
  }
}
