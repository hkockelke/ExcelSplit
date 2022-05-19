using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;
using System.Collections;

namespace ExcelSplit
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                throw new Exception("Missing arguments: input-CSV-File output-directory ");
            }
            
            string csvFile = args[0];
            string outDir = args[1];

            string csvFilename = Path.GetFileName(csvFile);
            string outname = csvFilename.Replace("_load", "");

            DateTime utcDate = DateTime.UtcNow;
            string CurrentDate = utcDate.ToString("yyMM");

            string KAToutcsvfile = Path.Combine(outDir, "Kat" + outname);
            string MAToutcsvfile = Path.Combine(outDir, "Mat" + outname);
            string SERoutcsvfile = Path.Combine(outDir, "Ser" + outname);
            string STRoutcsvfile = Path.Combine(outDir, "Str" + outname);

            StringBuilder outputKAT = new StringBuilder();
            StringBuilder outputMAT = new StringBuilder();
            StringBuilder outputSER = new StringBuilder();
            StringBuilder outputSTR = new StringBuilder();
            int icounter = 0;
            List<string> itemIds = new List<string>();
            Hashtable itemCounter = new Hashtable();
            Hashtable columns = new Hashtable();

            // perform first check of the data to count the numbers
            using (TextFieldParser csvParser = new TextFieldParser(csvFile))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { ";" });
                csvParser.HasFieldsEnclosedInQuotes = true;
                // The row with the column names:
                //string HeaderLine = csvParser.ReadLine();
                string[] HeaderFields = csvParser.ReadFields();
                foreach (string ColumnName in HeaderFields)
                {
                    if (columns.ContainsKey(ColumnName))
                    {
                        throw new Exception("Column Name duplicated: " + ColumnName);
                    }
                    else
                    {
                        columns.Add(ColumnName, icounter);
                    }
                    icounter++;
                }

                int i_level = Convert.ToInt32(columns["Level"]);
                int i_ID = Convert.ToInt32(columns["item_id"]);


                while (!csvParser.EndOfData)
                {
                    string[] fields = csvParser.ReadFields();
                    string Level = fields[i_level];
                    string ID = fields[i_ID];

                    if (Level == "1")
                    {
                        if (itemCounter.ContainsKey(ID))
                        {
                            string snbItem = itemCounter[ID].ToString();
                            int nbItem = Convert.ToInt32(snbItem);
                            nbItem++;
                            itemCounter[ID] = nbItem.ToString();
                        }
                        else
                        {
                            itemCounter.Add(ID, "1");
                        }
                    }
                }
            }

            icounter = 0;
            using (TextFieldParser csvParser = new TextFieldParser(csvFile))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { ";" });
                csvParser.HasFieldsEnclosedInQuotes = true;


                // Skip the row with the column names
                string HeaderLine = csvParser.ReadLine();
                outputKAT.AppendLine("LFD;Baugruppe;Teile-Nr;Verweis_auf_Baugruppe;POS;Menge;ME;Stue_Typ;Zeichnung");
                outputMAT.AppendLine("MATNR;Groesse_DIM;Text_DE;Text_EN;Text_ES;Text_FR;B;K;E;M;WN;Dru_EG;Dru_Rest;Drucken_Materialien");
                outputSER.AppendLine("Masch-Nr;EB-Nr;K_VARI;Text_DE;Text_EN;Text_ES;Text_FR;Text_ZH;TextNr;GrößeAbmessung");
                outputSTR.AppendLine("Strukturknoten;SKnoten Verweis;Verweis Nr;Laufende Nummer;T;Text_DE;Text_EN;Text_ES;Text_FR;Icon;Text_ZH;TextNr;GrößeAbmessung");

                int i_level = Convert.ToInt32(columns["Level"]);
                int i_SequenceNumber = Convert.ToInt32(columns["SequenceNumber"]);
                int i_bl_quantity = Convert.ToInt32(columns["bl_quantity"]);
                int i_ID = Convert.ToInt32(columns["item_id"]);
                //int i_SAP_Material_Number = Convert.ToInt32(columns["SAP Material Number"]);
                int i_SAP_Material_Number = Convert.ToInt32(columns["pm5_dr_rtp_mat_no_res"]);
                int i_item_revision_id = Convert.ToInt32(columns["item_revision_id"]);
                int i_object_name = Convert.ToInt32(columns["object_name"]);
                int i_pm5_dr_spare_part = Convert.ToInt32(columns["pm5_dr_spare_part"]);
                int i_pm5_dr_basic_material = Convert.ToInt32(columns["pm5_dr_basic_material"]);
                int i_pm5_dr_branch = Convert.ToInt32(columns["pm5_dr_branch"]);
                int i_pm5_dr_kc_code = Convert.ToInt32(columns["pm5_dr_kc_code"]);
                int i_pm5_dr_productlabel = Convert.ToInt32(columns["pm5_dr_productlabel"]);
                int i_pm5_dr_productlabel_add = Convert.ToInt32(columns["pm5_dr_productlabel_add"]);
                int i_pm5_dr_shorttext = Convert.ToInt32(columns["pm5_dr_shorttext"]);
                int i_pm5_dr_srp_code = Convert.ToInt32(columns["pm5_dr_srp_code"]);
                int i_pm5_dr_supplier_name1 = Convert.ToInt32(columns["pm5_dr_supplier_name1"]);
                int i_pm5_dr_surface_finish = Convert.ToInt32(columns["pm5_dr_surface_finish"]);
                int i_pm5_dr_welding_ge_eb_pa = Convert.ToInt32(columns["pm5_dr_welding_ge_eb_pa"]);
                int i_pm5_dr_welding_length = Convert.ToInt32(columns["pm5_dr_welding_length"]);
                int i_ics_1001 = Convert.ToInt32(columns["1001"]);
                int i_ics_1002 = Convert.ToInt32(columns["1002"]);
                int i_ics_1003 = Convert.ToInt32(columns["1003"]);
                int i_ics_1004 = Convert.ToInt32(columns["1004"]);
                int i_ics_1005 = Convert.ToInt32(columns["1005"]);
                int i_ics_1007 = Convert.ToInt32(columns["1007"]);
                int i_ics_1011 = Convert.ToInt32(columns["1011"]);
                int i_ics_1012 = Convert.ToInt32(columns["1012"]);
                int i_ics_1135 = Convert.ToInt32(columns["1135"]);
                int i_ics_1136 = Convert.ToInt32(columns["1136"]);
                int i_ics_1137 = Convert.ToInt32(columns["1137"]);
                int i_ics_1138 = Convert.ToInt32(columns["1138"]);
                int i_ics_1139 = Convert.ToInt32(columns["1139"]);
                int i_ics_1148 = Convert.ToInt32(columns["1148"]);

                while (!csvParser.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvParser.ReadFields();
                    int nbField = fields.Length;

                    if (nbField < 88)
                    {
                        throw new Exception("nbFields<88: " + nbField.ToString());
                    }

                    string Level = fields[i_level];
                    string s_item_id = fields[i_ID];
                    string s_SequenceNumber = fields[i_SequenceNumber];
                    string s_item_revision_id = fields[i_item_revision_id];
                    string s_ics_1001 = fields[i_ics_1001];
                    string s_ics_1007 = fields[i_ics_1007];
                    string s_ics_1011 = fields[i_ics_1011];
                    string s_ics_1012 = fields[i_ics_1012];
                    string s_ics_1135 = fields[i_ics_1135];
                    string s_ics_1136 = fields[i_ics_1136];
                    string s_ics_1137 = fields[i_ics_1137];
                    string s_ics_1138 = fields[i_ics_1138];
                    string s_ics_1139 = fields[i_ics_1139];
                    string s_ics_1148 = fields[i_ics_1148];
                    string SAP_Material_Number = fields[i_SAP_Material_Number]; 

                    if (Level == "0")
                    {
                        string Baugruppe_Zeile2 = s_ics_1138 + "." + s_ics_1139 + "." + s_ics_1148;
                        string Baugruppe_Zeile3 = "EB" + s_ics_1138 + "-" + s_ics_1139 + "-" + s_item_id + "-" + CurrentDate;
                        string Teile_Nr_Zeile2 = Baugruppe_Zeile3;
                        // 2nd line
                        outputKAT.AppendLine(";" + Baugruppe_Zeile2 + ";" + Teile_Nr_Zeile2 + Teile_Nr_Zeile2 + ";" + ";" + ";" + "EZ;");
                        // 3rd line
                        outputKAT.AppendLine(";" + Baugruppe_Zeile3 + ";" + Teile_Nr_Zeile2 + ";" + Teile_Nr_Zeile2 + ";" + ";" + ";" + ";EB;" + Teile_Nr_Zeile2 + "*.*");
                        // 4th line
                        outputKAT.AppendLine(icounter.ToString() + ";" + Baugruppe_Zeile3 + ";" + SAP_Material_Number + ";" + SAP_Material_Number + ";0;1;"+ s_ics_1001);

                        icounter++;
                    }
                    else if (Level == "1")
                    {
                       
                        if (!itemIds.Contains(s_item_id))
                        {
                            itemIds.Add(s_item_id);
                            string Quantity = itemCounter[s_item_id].ToString();
                            string EB_Nr = "EB" + s_ics_1138 + "-" + s_ics_1139 + "-" + CurrentDate;
                            string Text_DE = s_ics_1135 + s_ics_1136 + s_ics_1137;
                            string Text_EN = s_ics_1135 + s_ics_1136;
                            string Strukturknoten = s_ics_1138 + "." + s_ics_1139 + "." + s_ics_1148;
                            string SKnoten_Verweis = "EB" + s_ics_1138 + "-" + s_ics_1139 + "-" + s_item_id + "-" + CurrentDate;
                            string Laufende_Nummer = SKnoten_Verweis;
                            string Text_DE_Str = SKnoten_Verweis + " " + s_ics_1135 + s_ics_1136;
                            outputKAT.AppendLine(icounter.ToString() + ";EB" + s_item_id + ";" + s_item_revision_id + ";" + SAP_Material_Number + ";" + s_SequenceNumber + ";" + Quantity + ";;");
                            
                            outputMAT.AppendLine(SAP_Material_Number + ";" + s_ics_1137 + ";" + Text_EN + ";" + Text_EN + ";" + Text_EN + ";" + Text_EN + ";J;" + s_ics_1007 + ";" + s_ics_1011 + ";" + s_ics_1012 + ";WN000000;      ;        ;  ");
                            
                            outputSER.AppendLine(";"+ EB_Nr + ";" + EB_Nr + ";" + Text_DE + ";EN-Default;ES-Default;FR-Default;ZH-Default;" + s_ics_1136 + ";" + s_ics_1137);
                            
                            outputSTR.AppendLine(Strukturknoten + ";" + SKnoten_Verweis + ";;" + Laufende_Nummer + ";B;" + Text_DE_Str + ";;;;;;;");
                            
                            icounter++;
                        }
                    }
                    //all other levels are intentionally omitted

                }

            }

            // output: write to CSV files
            File.WriteAllText(KAToutcsvfile, outputKAT.ToString(), Encoding.UTF8);
            File.WriteAllText(MAToutcsvfile, outputMAT.ToString(), Encoding.UTF8);
            File.WriteAllText(SERoutcsvfile, outputSER.ToString(), Encoding.UTF8);
            File.WriteAllText(STRoutcsvfile, outputSTR.ToString(), Encoding.UTF8);
        }
    }
}
