using Algenta.Colectica.Model.Ddi;
using Algenta.Colectica.Model.Ddi.Serialization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace Soc2010Ddi
{
    public class Soc2010ToDdiConverter
    {
        static string sheetName = "Structure";
        static string agencyId = "int.example";

        public void Convert(string inputFileName, string outputFileName)
        {
            // Make sure we have valid inputs.
            if (string.IsNullOrWhiteSpace(inputFileName))
            {
                throw new ArgumentNullException("inputFileName");
            }

            if (string.IsNullOrWhiteSpace(outputFileName))
            {
                throw new ArgumentNullException("outputFileName");
            }

            // Create an empty code list.
            var codeList = new CodeList() { AgencyId = agencyId };

            // Track the most recently added code at each level of the hierarchy.
            var lastCodeByLevel = new Dictionary<int, Code>();

            // Track the current level in the code list hierarchy.
            //  1 = Major Group
            //  2 = Sub-Major Group
            //  3 = Minor Gruop
            //  4 = Unit Group
            int currentLevel = 0;

            // Go through each row of the spreadsheet and add a code 
            // for each row, at the appropriate level of the hierarchical
            // code list.
            var rows = GetRows(inputFileName);
            foreach (SocSpreadsheetRow row in rows)
            {
                int lastLevel = currentLevel;

                var code = new Code { AgencyId = agencyId };

                // Determine which level of code this row represents.
                if (row.IsMajor)
                {
                    currentLevel = 1;
                    code.Value = row.MajorGroup;
                }
                else if (row.IsSubMajor)
                {
                    currentLevel = 2;
                    code.Value = row.SubMajorGroup;
                }
                else if (row.IsMinor)
                {
                    currentLevel = 3;
                    code.Value = row.MinorGroup;
                }
                else if (row.IsUnit)
                {
                    currentLevel = 4;
                    code.Value = row.UnitGroup;
                }

                var category = new Category { AgencyId = agencyId };
                category.Label["en-GB"] = row.GroupTitle;
                code.Category = category;

                // If there are no codes for the level above this, 
                // add the code at the top of the code list.
                if (currentLevel == 1)
                {
                    codeList.Codes.Add(code);
                }
                else
                {
                    // If anything is in the stack, add the new code 
                    // as a child of appropriate code.
                    var parentCode = lastCodeByLevel[currentLevel - 1];
                    parentCode.ChildCodes.Add(code);
                }

                lastCodeByLevel[currentLevel] = code;
            }


            // Serialize the code list to DDI 3.2 and save it to disk.
            var serializer = new Ddi32Serializer();
            var xml = serializer.SerializeAsFragment(codeList);
            xml.Save(outputFileName);
        }

        /// <summary>
        /// Read each row of the spreadsheet into a simple data structure.
        /// </summary>
        /// <param name="inputFileName"></param>
        /// <returns></returns>
        List<SocSpreadsheetRow> GetRows(string inputFileName)
        {
            var results = new List<SocSpreadsheetRow>();

            DataTable table = GetSpreadsheetContents(inputFileName);
            foreach (DataRow row in table.Rows)
            {
                var code = new SocSpreadsheetRow
                {
                    MajorGroup = row["Major Group"].ToString(),
                    SubMajorGroup = row["Sub-Major Group"].ToString(),
                    MinorGroup = row["Minor Group"].ToString(),
                    UnitGroup = row["Unit   Group"].ToString(),
                    GroupTitle = row["Group Title"].ToString()
                };

                if (!code.IsEmpty)
                {
                    results.Add(code);
                }
            }

            return results;
        }

        /// <summary>
        /// Read the entire spreadsheet into a DataTable.
        /// </summary>
        /// <param name="inputFileName"></param>
        /// <returns></returns>
        DataTable GetSpreadsheetContents(string inputFileName)
        {
            // Load the spreadsheet.
            string connectionString = string.Format(
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";",
                inputFileName);

            var adapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}$]", sheetName), connectionString);
            var ds = new DataSet();

            adapter.Fill(ds, sheetName);

            return ds.Tables[0];
        }
    }

}