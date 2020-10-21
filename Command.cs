using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using CivilDB = Autodesk.Civil.DatabaseServices;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using civil3dCogoPoints;


[assembly: CommandClass(typeof(CornerRecordExtract.Commands))]
[assembly: ExtensionApplication(null)]
namespace CornerRecordExtract
{
    #region Commands
    public class Commands
    {
        [CommandMethod("OCPWBR")]
        public void ListAttributes()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            var doc = Application.DocumentManager.MdiActiveDocument;

            try
            {
                var acDB = doc.Database;

                using (var trans = acDB.TransactionManager.StartTransaction())
                {
                    DBDictionary layoutPages = (DBDictionary)trans.GetObject(acDB.LayoutDictionaryId,
                        OpenMode.ForRead);

                    // Handle Corner Record meta data dictionary extracted from Properties and Content
                    Dictionary<String, object> cornerRecordForms = new Dictionary<string, object>();

                    CivilDB.CogoPointCollection cogoPointsColl = CivilDB.CogoPointCollection.GetCogoPoints(doc.Database);
                    var cogoPointCollected = CogoPointJson.geolocationCapture(cogoPointsColl);


                    List<string> layoutNamesList = new List<string>();

                    foreach (DBDictionaryEntry layoutPage in layoutPages)
                    {
                        var crFormItems = layoutPage.Value.GetObject(OpenMode.ForRead) as Layout;
                        var isModelSpace = crFormItems.ModelType;

                        ObjectIdCollection textObjCollection = new ObjectIdCollection();

                        Dictionary<String, String> crAttributes = new Dictionary<String, String>();

                        if (isModelSpace != true)
                        {
                            BlockTableRecord blkTblRec = trans.GetObject(crFormItems.BlockTableRecordId,
                                OpenMode.ForRead) as BlockTableRecord;

                            layoutNamesList.Add(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ", ""));

                            foreach (ObjectId blkId in blkTblRec)
                            {
                                textObjCollection.Add(blkId);

                                var blkRef = trans.GetObject(blkId, OpenMode.ForRead) as BlockReference;

                                if (blkRef != null && blkRef.IsDynamicBlock)
                                {
                                    //ed.WriteMessage("\nBlock: " + blkRef.Name);
                                    //btr.Dispose();

                                    AttributeCollection attCol = blkRef.AttributeCollection;

                                    foreach (ObjectId attId in attCol)
                                    {
                                        AttributeReference attRef = (AttributeReference)trans.GetObject(attId, OpenMode.ForRead);

                                        //ed.WriteMessage("\nAttribute Tag: {0} \nAttribute String: {1}", attRef.Tag.ToString(), attRef.TextString.ToString());

                                        if (attRef.Tag.ToString() == "CITY_NAME")
                                        {
                                            crAttributes.Add("CRCity_c", attRef.TextString.ToString());
                                        }
                                        else if (attRef.Tag.ToString() == "LEGAL_DESCRIPTION")
                                        {
                                            crAttributes.Add("Corner_Type_c", "Lot");
                                            crAttributes.Add("Legal_Description_c", attRef.TextString.ToString());
                                        }
                                    }
                                }
                            }

                            // Build Final JSON File with LayoutName and Attributes
                            cornerRecordForms.Add(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ", ""), crAttributes);
                        }
                    }

                    // Checks to see whether the points from the cogo point collection exist within 
                    // the layout by searching for the correct collection key and layout name
                    List<string> cogoPointCollectedCheck = cogoPointCollected.Keys.ToList();
                    List<bool> boolCheckResults = new List<bool>();

                    IEnumerable<string> cogoPointNameCheck = layoutNamesList.Except(cogoPointCollectedCheck);
                    List<string> cogoPointNameCheckResults = cogoPointNameCheck.ToList();
                    var layoutNameChecker = new Regex("^(\\s*cr\\s*\\d\\d*)$");

                    if (!cogoPointNameCheckResults.Where(f => layoutNameChecker.IsMatch(f)).ToList().Any())
                    {
                        boolCheckResults.Add(true);
                    }
                    else
                    {
                        foreach (string cogoPointNameResultItem in cogoPointNameCheckResults)
                        {
                            Match layoutNameMatch = Regex.Match(cogoPointNameResultItem, "^(\\s*cr\\s*\\d\\d*)$",
                                RegexOptions.IgnoreCase);

                            if (layoutNameMatch.Success)
                            {
                                string layoutNameX = layoutNameMatch.Value;
                                ed.WriteMessage("\nLayout Named {0} does not have an associated cogo point", layoutNameX);
                            }
                        }
                        boolCheckResults.Add(false);
                    }


                    IEnumerable<string> layoutNameCheck = cogoPointCollectedCheck.Except(layoutNamesList);
                    List<string> layoutNameCheckResults = layoutNameCheck.ToList();
                    var cogoNameChecker = new Regex("^(\\s*cr\\s*\\d\\d*)$");


                    // If the layout name has any value other than CR == PASS
                    // If CR point exists and does not match then throw an error for user to fix
                    if (!layoutNameCheckResults.Where(f => cogoNameChecker.IsMatch(f)).ToList().Any())
                    {
                        boolCheckResults.Add(true);
                    }
                    else // Found a CR point that DID NOT match a layout name 
                    {
                        foreach (string layoutNameCheckResultItem in layoutNameCheckResults)
                        {
                            Match cogoNameMatch = Regex.Match(layoutNameCheckResultItem, "^(\\s*cr\\s*\\d\\d*)$",
                                RegexOptions.IgnoreCase);

                            if (cogoNameMatch.Success)
                            {
                                string cogoNameX = cogoNameMatch.Value;
                                ed.WriteMessage("\nCorner Record point named {0} does not have an associated Layout",
                                    cogoNameX);
                            }
                        }
                        boolCheckResults.Add(false);
                    }

                    // Output JSON file to BIN folder
                    // IF there are two true booleans in the list then add the data to the corresponding keys (cr1 => cr1)
                    if ((boolCheckResults.Count(v => v == true)) == 2)
                    {
                        foreach (string cornerRecordFormKey in cornerRecordForms.Keys)
                        {
                            if (cogoPointCollected.ContainsKey(cornerRecordFormKey))
                            {
                                //ed.WriteMessage("WORKS FOR {0}", cornerRecordFormKey);
                                var cogoFinal = (Dictionary<string, string>)cornerRecordForms[cornerRecordFormKey];

                                //var cogoFinalType = ((Dictionary<string, string>)cogoPointCollected[cornerRecordFormKey])["Corner_Type_c"];

                                var cogoFinalLong = ((Dictionary<string, object>)cogoPointCollected[cornerRecordFormKey])
                                   ["Geolocation_Longitude_s"];
                                var cogoFinalLat = ((Dictionary<string, object>)cogoPointCollected[cornerRecordFormKey])
                                    ["Geolocation_Latitude_s"];

                                //cogoFinal.Add("Corner_Type_c", cogoFinalType);
                                cogoFinal.Add("Geolocation_Longitude_s", cogoFinalLong.ToString());
                                cogoFinal.Add("Geolocation_Latitude_s", cogoFinalLat.ToString());
                            }
                        }

                        using (var writer = File.CreateText("CornerRecordForms.json"))
                        {
                            string strResultJson = JsonConvert.SerializeObject(cornerRecordForms,
                                Formatting.Indented);
                            writer.WriteLine(strResultJson);
                        }
                    }
                    trans.Commit();
                }
            }

            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
        }
    }
    #endregion
}