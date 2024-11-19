using EPDM.Interop.epdm;
using System;
using System.IO;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
    public class BOM_dt
    {
        static IEdmVault5 vault1 = new EdmVault5();
        static SldApp sldApp = null;

        public static void BOM(IEdmFile7 aFile, string config, int version, int BomFlag)

        {
            IEdmBomView bomView;

            bomView = aFile.GetComputedBOM(Root.strFullBOM, version, config, BomFlag); //1//(int)EdmBomFlag.EdmBf_AsBuilt + //2// (int)EdmBomFlag.EdmBf_ShowSelected);
            bomView.GetRows(out object[] ppoRows);
            bomView.GetColumns(out EdmBomColumn[] ppoColumns);
    
            if (ppoRows.Length > 0)
            {             
                foreach (IEdmBomCell ppoRow in ppoRows)
                {
                    ForColi(ppoRow, ppoColumns, aFile.Name.ToString());
                }
            }

            if (Root.drawings.Count > 0)
            {
                sldApp = new SldApp();
                sldApp.Run();
            }

        }

        static void ForColi(IEdmBomCell Row, EdmBomColumn[] ppoColumns, string aFileName)
        {
            string f = "";//Found In
            IEdmFile7 bFile;

            object poValue = null;
            object poComputedValue = null;
            string pbsConfiguration = "";
            bool pbReadOnly = false;
            int refDrToModel = -1;
            bool NeedsRegeneration = false;
            int LatestVer = -1;
            string Config = "";
            string Section = "";
            IEdmFile7 modelFile = null;
            string p="";
            string d="";
            Drawing draw = null;

            if (Row.GetTreeLevel() == 1 || Row.GetTreeLevel() == 0)
            {
                for (int Coli = 0; Coli < ppoColumns.Length; Coli++)

                {
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strFoundIn))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                        f = poComputedValue.ToString();

                    }
                  
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strLatestVer))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                        LatestVer = Convert.ToInt16(poComputedValue);

                    }
                   
                  

                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strConfig))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                        Config = poComputedValue.ToString();

                    }
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strSection))
                    {
                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);

                        Section = poComputedValue.ToString();
                    
                        if(!(Section == "Сборочные единицы"|| Section == "Детали"))
                        {
                            return;
                        }
                    }

                    //Если деталь или сборка
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strFileName))
                    {

                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                         p = f + "\\" + poComputedValue.ToString();       //Путь к файлу детали или сборки
                         d = "";                                          //Путь к файлу чертежа



                        if (poComputedValue.ToString().Contains(".sldasm"))
                        {
                             d = p.Replace(".sldasm", ".SLDDRW");
                        }
                        else if (poComputedValue.ToString().Contains(".SLDASM"))
                        {
                            d = p.Replace(".SLDASM", ".SLDDRW");
                        }
                        else if (poComputedValue.ToString().Contains(".sldprt"))
                        {
                            d = p.Replace(".sldprt", ".SLDDRW");

                        }
                        else if (poComputedValue.ToString().Contains(".SLDPRT"))
                        {
                             d = p.Replace(".SLDPRT", ".SLDDRW");

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

                           if (!(refDrToModel == modelFile.CurrentVersion) || NeedsRegeneration)
                               {
                                  draw = new Drawing(bFile.ID, bFolder.ID, d);
                                  Root.drawings.Add(draw);
                               }
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("error:" + bFile.Name + ex.Message);

                    }
                    
                }

            }

            if (Row.GetTreeLevel() == 1 && Section == "Сборочные единицы")
            {
               
              
                    if (modelFile != null) { BOM(modelFile, Config, LatestVer, 0); }
                
            }


        }



    }
}
    
