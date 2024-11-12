using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
    public class BOM_dt
    {
        static IEdmVault5 vault1 = new EdmVault5();

        public static void BOM(IEdmFile7 aFile, string config, int version, int BomFlag, ref DataTable dt)

        {
            IEdmBomView bomView;

            bomView = aFile.GetComputedBOM(Root.strFullBOM, version, config, BomFlag); //1//(int)EdmBomFlag.EdmBf_AsBuilt + //2// (int)EdmBomFlag.EdmBf_ShowSelected);
            bomView.GetRows(out object[] ppoRows);
            bomView.GetColumns(out EdmBomColumn[] ppoColumns);

            dt.Columns.Add(Root.strFileID, typeof(int));
            dt.Columns.Add(Root.strFolderID, typeof(int));
            dt.Columns.Add(Root.strFileName, typeof(string));
            dt.Columns.Add(Root.strFoundIn, typeof(string));
            dt.Columns.Add(Root.strLatestVer, typeof(int));
            dt.Columns.Add(Root.strRev, typeof(int));
            dt.Columns.Add(Root.strDrawState, typeof(string));
            dt.Columns.Add(Root.strNeedsRegeneration, typeof(bool));
            dt.Columns.Add(Root.strSection, typeof(string));
            dt.Columns.Add(Root.strConfig, typeof(string));

            if (ppoRows.Length > 0)
            {             
                foreach (IEdmBomCell ppoRow in ppoRows)
                {
                    ForColi(ppoRow, ppoColumns, aFile.Name.ToString(), ref dt);
                }
            }

        }

        static void ForColi(IEdmBomCell Row, EdmBomColumn[] ppoColumns, string aFileName,ref DataTable dt)
        {
            string f = "";//Found In
            IEdmFile7 bFile;



            bool TrueRowFlag = false;
            DataRow workRow = dt.NewRow();
            object poValue = null;
            object poComputedValue = null;
            string pbsConfiguration = "";
            bool pbReadOnly = false;
            int refDrToModel = -1;
            bool NeedsRegeneration = false;
            IEdmFile7 modelFile = null;
            string p="";
            string d="";

            if (Row.GetTreeLevel() == 1 || Row.GetTreeLevel() == 0)
            {
                for (int Coli = 0; Coli < ppoColumns.Length; Coli++)

                {
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strFoundIn))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                        f = poComputedValue.ToString();

                    }
                    /*
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strLatestVer))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                        workRow[Root.strLatestVer] = (int)poComputedValue;

                    }
                    */

                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strConfig))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                        workRow[Root.strConfig] = poComputedValue.ToString();

                    }
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strSection))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);

                        workRow[Root.strSection] = poComputedValue.ToString();

                    }


                    //Если деталь или сборка, то вносим в таблицу с информацией о наличии чертежа и dxf, иначе игнорим
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strFileName))
                    {

                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                         p = f + "\\" + poComputedValue.ToString();       //Путь к файлу детали или сборки
                         d = "";                                          //Путь к файлу чертежа



                        if (poComputedValue.ToString().Contains(".sldasm"))
                        {
                            TrueRowFlag = true; d = p.Replace(".sldasm", ".SLDDRW");
                        }
                        else if (poComputedValue.ToString().Contains(".SLDASM"))
                        {
                            TrueRowFlag = true; d = p.Replace(".SLDASM", ".SLDDRW");
                        }
                        else if (poComputedValue.ToString().Contains(".sldprt"))
                        {
                            TrueRowFlag = true; d = p.Replace(".sldprt", ".SLDDRW");

                        }
                        else if (poComputedValue.ToString().Contains(".SLDPRT"))
                        {
                            TrueRowFlag = true; d = p.Replace(".SLDPRT", ".SLDDRW");

                        }

                    }
                }


                //Проверяем есть ли зачекиненный чертеж в папке с деталью с именем соответствующим детали
                if (!vault1.IsLoggedIn) { vault1.LoginAuto(Root.pdmName, 0); }
                modelFile = (IEdmFile7)vault1.GetFileFromPath(p, out IEdmFolder5 modelFolder);

                bFile = (IEdmFile7)vault1.GetFileFromPath(d, out IEdmFolder5 bFolder);

                if ((bFile != null) && (!bFile.IsLocked)) //true если файл не пусто и зачекинен                                           
                {
                    try
                    {
                        workRow[Root.strFileID] = bFile.ID;
                        workRow[Root.strFolderID] = bFolder.ID;
                        workRow[Root.strFileName] = d;

                        workRow[Root.strFoundIn] = f;
                        workRow[Root.strDrawState] = bFile.CurrentState.Name.ToString();


                        int versionDraiwing = bFile.CurrentVersion;


                        NeedsRegeneration = bFile.NeedsRegeneration(versionDraiwing, bFolder.ID);

                        // Достаем из чертежа версию ссылки на родителя (VersionRef)
                        IEdmReference5 ref5 = bFile.GetReferenceTree(bFolder.ID);
                        IEdmReference10 ref10 = (IEdmReference10)ref5;
                        IEdmPos5 pos = ref10.GetFirstChildPosition3("A", true, true, (int)EdmRefFlags.EdmRef_File, "", 0);
                        while (!pos.IsNull)
                        {

                            IEdmReference10 @ref = (IEdmReference10)ref5.GetNextChild(pos);
                            //
                            string extension = Path.GetExtension(@ref.Name);
                            if (extension == ".sldasm" || extension == ".sldprt" || extension == ".SLDASM" || extension == ".SLDPRT")
                            {
                                refDrToModel = @ref.VersionRef;
                                break;
                            }
                            else
                            {
                                ref5.GetNextChild(pos);
                            }
                        }

                         dt.Rows.Add(workRow);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("error:" + bFile.Name);

                    }
                    //  if (TrueRowFlag == true && (!(refDrToModel == modelFile.CurrentVersion) || NeedsRegeneration))
                            //   {
                      
                            //   }
                }

            }

            else

            { TrueRowFlag = false; }//если не первый левел или не 0 левел


            

            if (Row.GetTreeLevel() == 1)
            {
                if (workRow[dt.Columns.IndexOf(Root.strSection)].ToString().Contains("Сборочные единицы"))
                {
                    string conf = workRow[dt.Columns.IndexOf(Root.strConfig)].ToString();
                    int vers = Convert.ToInt16(workRow[dt.Columns.IndexOf(Root.strLatestVer)]);//Последняя версия файла в PDM
                    if (modelFile != null) { BOM(modelFile, conf, vers, 0, ref dt); }//1//(int)EdmBomFlag.EdmBf_AsBuilt + //2// (int)EdmBomFlag.EdmBf_ShowSelected);
                }
            }


        }

    }
}
    
