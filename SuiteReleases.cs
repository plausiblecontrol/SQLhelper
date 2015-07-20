using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace Smarp_o_Matic_XL_Gold_Premium {
  class SuiteQry {
    /*
     * 
     * select ProductName, Release, ReleaseDate, LastRelease, PENotes, ReleaseNotes, DeveloperNotes, CustomerReadmeNotes 
          from OSI_ReleaseInfo 
          where ProductName = 'OpenLSR' and ReleaseDate > '2015-04-02 00:00:00.000' and ReleaseDate <= '2015-05-14'
     * */
    public List<string[]> getPatch(string SuiteName, string majorrel, string majrev, string oldSuiteD, string newSuiteD) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> prodNotes = new List<string[]>();
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "SELECT s.Name Suite, sr.MajorRelease MR, sr.MajorRevision Mr, sr.MinorRevision mr, sr.Patch p, sr.ReleaseDate " +
         "from OSI_RMT_SuiteReleases sr " +
         "join OSI_RMT_Suites s on sr.SuiteID = s.ID " +
         "join OSI_RMT_SuiteReleaseStates srs on srs.ID = sr.StateID " +
         "join OSI_RMT_SuiteReleaseVisibility srv on srv.ID = sr.VisibilityID and s.Name = '" + SuiteName + "' and ReleaseDate > '" + oldSuiteD + "' and ReleaseDate < '" + newSuiteD + "'  " +
         "order by s.Name, MajorRelease, MajorRevision, MinorRevision, Patch";

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          //prodname[0], release[1], requirements[4], relnotes[5], eNotes[6], devNotes[7]
          prodNotes.Add(new string[] { rdr[0].ToString(), rdr[1].ToString(), rdr[2].ToString(), rdr[3].ToString(), rdr[4].ToString(), rdr[5].ToString() });
        }
      } catch { }
      return prodNotes;

    }

    public List<string[]> getProdNotes(string Product, string release, string relDate) {      
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> prodNotes = new List<string[]>();
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "select ProductName, Release, ReleaseDate, LastRelease, Requirements, ReleaseNotes, PENotes, DeveloperNotes " +
          "from OSI_ReleaseInfo "+
          "where ProductName = '" + Product + "' and Release = '" + release + "' and ReleaseDate = '"+relDate+"'" ;

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          //prodname[0], release[1], requirements[4], relnotes[5], eNotes[6], devNotes[7]
          prodNotes.Add(new string[] { rdr[0].ToString(), rdr[1].ToString(), rdr[4].ToString(), rdr[5].ToString(), rdr[6].ToString(), rdr[7].ToString()});
        }
      } catch { }
      return prodNotes;

    }

    public List<string[]> getNotes(string Suite, string majRel, string majRev, string minRev, string p) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> prodNotes = new List<string[]>();
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "select ProductName, Release, ri.ReleaseDate, ri.ReleaseNotes, PENotes, DeveloperNotes, Requirements, CustomerReadmeNotes, " +
          "s.Name, MajorRelease, MajorRevision, MinorRevision, Patch, srs.State, sr.ReleaseDate " +
          "from OSI_RMT_SuiteReleases sr " +
          "join OSI_RMT_Suites s on s.ID = sr.SuiteID " +
          "join OSI_RMT_SuiteReleaseProductRelease srpr on srpr.SuiteReleaseID = sr.ID " +
          "join OSI_ReleaseInfo ri on ri.ID = srpr.ProductReleaseID " +
          "join OSI_RMT_SuiteReleaseStates srs on srs.ID = sr.StateID " +
          "join OSI_RMT_SuiteReleaseVisibility srv on srv.ID = sr.VisibilityID " +
          "and Name like '%" + Suite + "%' and MajorRelease = '" + majRel + "' and MajorRevision = '" + majRev + "' and MinorRevision = '" + minRev + "' and Patch = '" + p + "'";


        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          //prodname[0], release[1], relNotes[3], eNotes[5], devNotes[6], custNotes[8]
          prodNotes.Add(new string[] { rdr[0].ToString(), rdr[1].ToString(), rdr[3].ToString(), rdr[4].ToString(), rdr[5].ToString(), rdr[6].ToString() + Environment.NewLine + rdr[7].ToString() });
        }
      } catch { }
      return prodNotes;

    }

    public string[] getExceptions(string product, string releasedate) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      DateTime oD = Convert.ToDateTime(releasedate).AddDays(-14);
      string[] prodNotes = new string[7];
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "select ProductName, Release, ReleaseDate, LastRelease, PENotes, ReleaseNotes, DeveloperNotes, CustomerReadmeNotes  " +
          "from OSI_ReleaseInfo " +
          "where ProductName = '" + product + "' and ReleaseDate > '" + oD.ToString("M/d/yyyy") + "' and ReleaseDate <= '" + releasedate + "'" +
          "order by ReleaseDate asc";

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {//get latest
          //prodname[0], release[1], relNotes[5], eNotes[4], devNotes[6], custNotes[7]
          prodNotes = new string[] { rdr[0].ToString(), rdr[1].ToString(), rdr[5].ToString(), rdr[4].ToString(), rdr[6].ToString(), rdr[7].ToString()};
        }
      } catch { }
      return prodNotes;
    }


    /*Select sr.ReleaseDate, s.Name, sr.MajorRelease, sr.MajorRevision, sr.MinorRevision, sr.Patch, ri.ProductName, ri.ReleaseDate, ri.Release, ri.Requirements, ri.ReleaseNotes, ri.PENotes, ri.DeveloperNotes
          from OSI_RMT_SuiteReleases sr 
          join OSI_RMT_Suites s on s.ID = sr.SuiteID 
          join OSI_RMT_SuiteReleaseProductRelease srpr on srpr.SuiteReleaseID = sr.ID 
          join OSI_ReleaseInfo ri on ri.ID = srpr.ProductReleaseID 
          where s.Name = 'Platform Server' and sr.ReleaseDate >= '4/20/15' and sr.ReleaseDate < '5/1/15'
     * */
    public List<string[]> getSuites(string SuiteName, DateTime todayDate) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> info = new List<string[]>();
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();
        string a = todayDate.AddDays(-14).ToString("M/d/yyyy");
        string b = todayDate.ToString("M/d/yyyy");
        string query = "Select sr.ReleaseDate, s.Name, sr.MajorRelease, sr.MajorRevision, sr.MinorRevision, sr.Patch, ri.ProductName, ri.ReleaseDate, ri.Release, ri.Requirements, ri.ReleaseNotes, ri.PENotes, ri.DeveloperNotes, ri.CustomerReadmeNotes  " +
          "from OSI_RMT_SuiteReleases sr  " +
          "join OSI_RMT_Suites s on s.ID = sr.SuiteID  " +
          "join OSI_RMT_SuiteReleaseProductRelease srpr on srpr.SuiteReleaseID = sr.ID  " +
          "join OSI_ReleaseInfo ri on ri.ID = srpr.ProductReleaseID  " +
          "where s.Name = '" + SuiteName + "' and sr.ReleaseDate >= '" + a + "' and sr.ReleaseDate < '" + b + "' " +
          "order by sr.MajorRelease, sr.MajorRevision, sr.MinorRevision, sr.Patch";

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          //SuiteReleaseDate0	SName1	MajorRelease2	MajorRevision3	MinorRevision4	Patch5	ProductName6	PReleaseDate7	Release8	Requirements9	ReleaseNotes10	PENotes11	DeveloperNotes12
          string[] line = new string[] { rdr[0].ToString(), rdr[1].ToString(), rdr[2].ToString(), rdr[3].ToString(), rdr[4].ToString(), rdr[5].ToString(), rdr[6].ToString(), rdr[7].ToString(), rdr[8].ToString(), "Req: "+rdr[9].ToString()+Environment.NewLine+"Readme: "+rdr[13].ToString(), rdr[10].ToString(), rdr[11].ToString(), rdr[12].ToString(), "" };
          info.Add(line);
        }
      } catch { }

      for (int i = 0; i < info.Count; i++) {
        DateTime sDate = Convert.ToDateTime(info[i][0]);
        DateTime bot = Convert.ToDateTime("4/30/2015");
        while (sDate > bot) {
          sDate = sDate.AddDays(-14);          
        }
        if (sDate == bot ){
          info[i][13] = "Main";
        } else {
          info[i][13] = "Patch";
        }
      }

      //for (int i = 0; i < info.Count; i++) {
      //  DateTime sDate = Convert.ToDateTime(info[i][0]);
      //  DateTime rDate = Convert.ToDateTime(info[i][7]);
      //  if (info[i][13] == "Main") {
      //    //if (rDate != sDate && rDate.AddDays(+1) != sDate && rDate.AddDays(+2) != sDate) {
      //    DateTime cDate = rDate.AddDays(13);
      //    if (rDate != sDate && cDate < sDate) {
      //      info[i][9] = "Same as previous sprint";
      //      info[i][10] = "";
      //      info[i][11] = "";
      //      info[i][12] = "";
      //    }
      //  } else {
      //    if (rDate != sDate) {
      //      info[i][9] = "Same as previous sprint";
      //      info[i][10] = "";
      //      info[i][11] = "";
      //      info[i][12] = "";
      //    }
      //  }
      //}
        return info;
    }




    public List<string[]> difSuites(string[] SuiteA, string[] SuiteB) {
      List<string[]> differences = new List<string[]>();
      List<string[]> ProdsA = getProds(SuiteA[0], SuiteA[1], SuiteA[2], SuiteA[3], SuiteA[4]);
      List<string[]> ProdsB = getProds(SuiteB[0], SuiteB[1], SuiteB[2], SuiteB[3], SuiteB[4]);
      foreach (string[] x in ProdsA) {
        foreach (string[] y in ProdsB) {
          if (x[0] == y[0]) { //OpCalc == OpCalc
            if (x[1] == y[1]) {//1.2.3p1 == 1.2.3p1
              differences.Add(new string[] { x[0], x[1], "Same release as previous sprint" });
              break;
            } else {
              differences.Add(new string[] { x[0], x[1], "See notes" });
              break;
            }
          }
        }
      }
      return differences;
    }

    public List<string[]> getProds(string Suite, string majRel, string majRev, string minRev, string p) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> products = new List<string[]>();
      string[] product = new string[2];
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "Select ri.ProductName, ri.Release " +
          "from OSI_RMT_SuiteReleases sr " +
          "join OSI_RMT_Suites s on s.ID = sr.SuiteID " +
          "join OSI_RMT_SuiteReleaseProductRelease srpr on srpr.SuiteReleaseID = sr.ID " +
          "join OSI_ReleaseInfo ri on ri.ID = srpr.ProductReleaseID " +
          "where s.Name = '" + Suite + "' and " +
          " MajorRelease = '" + majRel + "' and MajorRevision = '" + majRev + "' and MinorRevision = '" + minRev + "' and Patch = '" + p + "'";

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          products.Add(new string[] { rdr[0].ToString(), rdr[1].ToString() });
        }
      } catch { }
      return products;
    }

    public List<string[]> getPrevAvail(string SuiteName, DateTime requested, string PreviousMatches) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> LatestRels = new List<string[]>();
      string[] LatestRel = new string[8];
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "SELECT sr.ID, s.Name Suite, sr.MajorRelease MR, sr.MajorRevision Mr, sr.MinorRevision mr, sr.Patch p, srs.State, sr.ReleaseDate, srv.Visibility " +
         "from OSI_RMT_SuiteReleases sr " +
         "join OSI_RMT_Suites s on sr.SuiteID = s.ID " +
         "join OSI_RMT_SuiteReleaseStates srs on srs.ID = sr.StateID " +
         "join OSI_RMT_SuiteReleaseVisibility srv on srv.ID = sr.VisibilityID and Patch = '1' " +
         "order by s.Name, MajorRelease, MajorRevision, MinorRevision, Patch";

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          string[] row = new string[8];
          if (rdr[1].ToString() == SuiteName) {
            for (int i = 0; i < 8; i++) {
              row[i] = rdr[i + 1].ToString();
            }
            //COMMS	6	2	19	1	Issued	2015-02-05 00:00:00.000	Available
            if (PreviousMatches == "") {
              if ((Convert.ToDateTime(LatestRels[0][6]) < Convert.ToDateTime(row[6])) && (Convert.ToDateTime(row[6]) < requested)) {
                LatestRel = row;
                LatestRels.Clear();
                LatestRels.Add(LatestRel);
              } else if (Convert.ToDateTime(LatestRels[0][6]) == Convert.ToDateTime(row[6])) {
                LatestRels.Add(row);
              }
            } else {
              if (PreviousMatches.Contains(row[1] + row[2])) {
                if (LatestRels.Count == 0) {
                  LatestRels.Add(row);
                } else {
                  for (int i = 0; i < LatestRels.Count; i++) {
                    if (row[0] == LatestRels[i][0] && row[1] == LatestRels[i][1] && row[2] == LatestRels[i][2] && Convert.ToInt32(LatestRels[i][3]) < Convert.ToInt32(row[3]) && (Convert.ToDateTime(row[6]) < requested)) {
                      LatestRels.RemoveAt(i);
                      LatestRels.Add(row);
                      break;
                    } else if ((i + 1) == LatestRels.Count && (Convert.ToDateTime(row[6]) < requested)) {
                      LatestRels.Add(row);
                      break;
                    }
                  }
                }
              }
            }
          }
        }
      } catch (Exception e) {
        string mrgs = e.Message;
      }
      return LatestRels;
    }

    public List<string[]> GetSuiteBunch(string SuiteName, string reqDate, string prevDate) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> LatestRels = new List<string[]>();
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "SELECT s.Name Suite, sr.MajorRelease MR, sr.MajorRevision Mr, sr.MinorRevision mr, sr.Patch p, sr.ReleaseDate " +
         "from OSI_RMT_SuiteReleases sr " +
         "join OSI_RMT_Suites s on sr.SuiteID = s.ID " +
         "join OSI_RMT_SuiteReleaseStates srs on srs.ID = sr.StateID " +
         "join OSI_RMT_SuiteReleaseVisibility srv on srv.ID = sr.VisibilityID and s.Name = '" + SuiteName + "' and ReleaseDate > '" + prevDate + "' and ReleaseDate <= '" + reqDate + "'  " + 
         "order by s.Name, MajorRelease, MajorRevision, MinorRevision, Patch";

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          LatestRels.Add(new string[] { rdr[0].ToString(), rdr[1].ToString(), rdr[2].ToString(), rdr[3].ToString(), rdr[4].ToString(), rdr[5].ToString() });
        }
      } catch (Exception e) {
        string mrgs = e.Message;
      }
      
      return LatestRels;
    }



    public List<string[]> getLatestAvail(string SuiteName, DateTime requested) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<string[]> LatestRels = new List<string[]>();
      string[] LatestRel = new string[8];
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();

        string query = "SELECT sr.ID, s.Name Suite, sr.MajorRelease MR, sr.MajorRevision Mr, sr.MinorRevision mr, sr.Patch p, srs.State, sr.ReleaseDate, srv.Visibility " +
         "from OSI_RMT_SuiteReleases sr " +
         "join OSI_RMT_Suites s on sr.SuiteID = s.ID " +
         "join OSI_RMT_SuiteReleaseStates srs on srs.ID = sr.StateID " +
         "join OSI_RMT_SuiteReleaseVisibility srv on srv.ID = sr.VisibilityID "+
         "order by s.Name, MajorRelease, MajorRevision, MinorRevision, Patch";

        cmd = new SqlCommand(query);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        DateTime later = Convert.ToDateTime("1/1/1900");
        LatestRel[6] = later.ToString();
        LatestRels.Add(LatestRel);
        while (rdr.Read()) {
          string[] row = new string[8];
          if (rdr[1].ToString() == SuiteName) {
            for (int i = 0; i < 8; i++) {
              row[i] = rdr[i + 1].ToString();
            }
            //COMMS	6	2	19	1	Issued	2015-02-05 00:00:00.000	Available
            if ((Convert.ToDateTime(LatestRels[0][6]) < Convert.ToDateTime(row[6])) && (Convert.ToDateTime(row[6]) < requested)) {
              LatestRel = row;
              LatestRels.Clear();
              LatestRels.Add(LatestRel);
            } else if (Convert.ToDateTime(LatestRels[0][6]) == Convert.ToDateTime(row[6])) {
              LatestRels.Add(row);
            }
          }
        }
      } catch (Exception e) {
        string mrgs = e.Message;
      }
      bool oldnews = false;
      List<string[]> exitRels = new List<string[]>();
      foreach (string[] x in LatestRels) {
        for (int i = 0; i < LatestRels.Count; i++) {
          if (x[0] == LatestRels[i][0] && x[1] == LatestRels[i][1] && x[2] == LatestRels[i][2] && x[3] == LatestRels[i][3] && Convert.ToInt32(LatestRels[i][4]) > Convert.ToInt32(x[4])) {
            oldnews = true;
            break;
          }
        }
        if (!oldnews) {
          exitRels.Add(x);
        }
        oldnews = false;
      }
      return exitRels;
    }

  }
}
