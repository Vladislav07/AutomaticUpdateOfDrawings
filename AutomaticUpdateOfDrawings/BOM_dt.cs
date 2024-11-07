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

        IEdmBatchUnlock2 batchUnlocker;
        IEdmSelectionList6 fileList = null;
        EdmSelectionObject poSel;

        int fileCount = 0;
        static int i = 0;


        public static void BOM(IEdmFile7 aFile, string config, int version, int BomFlag, ref DataTable dt)

        {
            IEdmBomView bomView;

            bomView = aFile.GetComputedBOM(Root.strFullBOM, version, config, BomFlag); //1//(int)EdmBomFlag.EdmBf_AsBuilt + //2// (int)EdmBomFlag.EdmBf_ShowSelected);
            bomView.GetRows(out object[] ppoRows);
            bomView.GetColumns(out EdmBomColumn[] ppoColumns);


            //Заполняем таблицу dt данными из BOM
            if (ppoRows.Length > 0)//Отбрасываем пустые сборки
            {
    

                //Построчно добавляем строки из BOM в таблицу dt с нужной доп информацией 
                foreach (IEdmBomCell ppoRow in ppoRows)
                {
                    ForColi(ppoRow, ppoColumns, aFile.Name.ToString(), dt);
                }
            }

        }

        static void ForColi(IEdmBomCell Row, EdmBomColumn[] ppoColumns, string aFileName,  DataTable dt)
        {
            string f = "";//Found In
            IEdmFile7 bFile;



            bool TrueRowFlag = false;
            DataRow workRow = dt.NewRow();
            object poValue = null;
            object poComputedValue = null;
           // string pbsConfiguration = "";
           // bool pbReadOnly = false;
 

            if (Row.GetTreeLevel() == 1 || Row.GetTreeLevel() == 0)
            {
                for (int Coli = 0; Coli < ppoColumns.Length; Coli++)

                {
                  

                    //Если деталь или сборка, то вносим в таблицу с информацией о наличии чертежа и dxf, иначе игнорим
                    if (ppoColumns[Coli].mbsCaption.Contains(Root.strFileName))
                    {

                        Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                        string p = f + "\\" + poComputedValue.ToString();       //Путь к файлу детали или сборки
                        string d = "";                                          //Путь к файлу чертежа


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


                        //Проверяем есть ли зачекиненный чертеж в папке с деталью с именем соответствующим детали
                        if (!vault1.IsLoggedIn) { vault1.LoginAuto(Root.pdmName, 0); }
                        bFile = (IEdmFile7)vault1.GetFileFromPath(d, out IEdmFolder5 bFolder);

                        int refDrToModel = -1;
                        bool isValiddrawing = false;


                        if ((bFile != null) && (!bFile.IsLocked)) //true если файл не пусто и зачекинен                                           
                        {
                            workRow[Root.strDraw] = true;
                            workRow[Root.strDrawState] = bFile.CurrentState.Name.ToString();
                            try
                            {
                                int versionDraiwing = bFile.CurrentVersion;

                                List<string> listDrawings = new List<string>();

                                isValiddrawing = bFile.NeedsRegeneration(versionDraiwing, bFolder.ID);

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
                                        workRow[Root.strRev] = @ref.VersionRef.ToString();
                                        refDrToModel = @ref.VersionRef;
                                        break;
                                    }
                                    else
                                    {
                                        ref5.GetNextChild(pos);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("error:" + bFile.Name);

                            }
                            IEdmFile7 modelFile = (IEdmFile7)vault1.GetFileFromPath(p, out IEdmFolder5 modelFolder);
                            if (!(refDrToModel == modelFile.CurrentVersion) || isValiddrawing)
                            {
                               // writeNameToTxt(bFile.Name + "vers :" + workRow[Root.strRev] + " - " + workRow[Root.strLatestVer] + " Is - " + isValiddrawing.ToString());
                                EdmSelItem selItem = new EdmSelItem();
                                selItem.mlDocID = bFile.ID;
                                selItem.mlProjID = bFolder.ID;
                                Root.SelectionDrawings.Add(selItem);
                                Root.listdrawings.Add(bFile.GetLocalPath(bFolder.ID));


                            }
                            else
                            {
                                // MessageBox.Show(bFile.Name);
                            }

                        }

                    }


                 

                    //Заполняем колонки
                    Row.GetVar(ppoColumns[Coli].mlVariableID, ppoColumns[Coli].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                    workRow[Coli] = poComputedValue.ToString();

            
                }
            }

            else

            { TrueRowFlag = false; }//если не первый левел или не 0 левел


            if (TrueRowFlag == true)
            {
                //ПРОВЕРКИ
                workRow[Root.strErC] = 0;//Количество ошибок в строке

                string regCuby = @"^CUBY-\d{8}$";
                string fileName = workRow[Root.strFileName].ToString();
                string[] parts = fileName.Split('.');
                string cuteFileName = parts[0].ToString();
                bool IsCUBY = Regex.IsMatch(cuteFileName, regCuby);

                //1. Проверка на наличие чертежа
                if (workRow[Root.strDraw].ToString() == ""
                    && (workRow[Root.strSection].ToString() == "Детали"
                    || workRow[Root.strSection].ToString() == "Сборочные единицы")
                    && (workRow[Root.strState].ToString() != Root.strPrelim)
                    && workRow[Root.strNoSHEETS].ToString() != "1")
                { workRow[Root.strErC] = Convert.ToInt16(workRow[Root.strErC]) + 1; } //Количество ошибок в строке


                //8. Проверка деталей и сборок в статусе Initiated
                if ((workRow[Root.strSection].ToString() == "Детали" || workRow[Root.strSection].ToString() == "Сборочные единицы")
                    && workRow[Root.strState].ToString() == Root.strInitiated)
                { workRow[Root.strErC] = Convert.ToInt16(workRow[Root.strErC]) + 1; }//Количество ошибок в строке

                //8.1 Проверка деталей и сборок в статусах "r_In Work" "r_На проверке" "Подтверждение ведущего" (статусы потока удаленщиков, не должно быть в таком)
                if ((workRow[Root.strSection].ToString() == "Детали" || workRow[Root.strSection].ToString() == "Сборочные единицы")
                    && (workRow[Root.strState].ToString() == "r_In Work" ||
                        workRow[Root.strState].ToString() == "r_На проверке" ||
                        workRow[Root.strState].ToString() == "Подтверждение Ведущего"))
                { workRow[Root.strErC] = Convert.ToInt16(workRow[Root.strErC]) + 1; }//Количество ошибок в строке


                //9. Проверка покупных в статусе In Work
                if ((workRow[Root.strSection].ToString() == "Прочие изделия"
                    || workRow[Root.strSection].ToString() == "Стандартные изделия"
                    || workRow[Root.strSection].ToString() == "Материалы")
                        && workRow[Root.strState].ToString() == Root.strInWork)
                { workRow[Root.strErC] = Convert.ToInt16(workRow[Root.strErC]) + 1; }//Количество ошибок в строке


                //10. Проверка равенства состояния чертежа и детали или сборки
                if (workRow[Root.strDraw].ToString() == "True"
                    && workRow[Root.strState].ToString() != workRow[Root.strDrawState].ToString())
                { workRow[Root.strErC] = Convert.ToInt16(workRow[Root.strErC]) + 1; }//Количество ошибок в строке






       

           


        

 

               
            

                dt.Rows.Add(workRow);
            }

            if (Row.GetTreeLevel() == 1)
            {
                if (workRow[dt.Columns.IndexOf(Root.strFileName)].ToString().Contains(".sldasm") || workRow[dt.Columns.IndexOf(Root.strFileName)].ToString().Contains(".SLDASM"))
                {
                    string conf = workRow[dt.Columns.IndexOf(Root.strConfig)].ToString();
                    int vers = Convert.ToInt16(workRow[dt.Columns.IndexOf(Root.strLatestVer)]);//Последняя версия файла в PDM
                                                                                                        // int vers = Convert.ToInt16(workRow[dt.Columns.IndexOf(Root.strFoundInVer)]); //Версия используемая в сборке
                    string pathFile = workRow[dt.Columns.IndexOf(Root.strFoundIn)].ToString() + "\\" + workRow[dt.Columns.IndexOf(Root.strFileName)].ToString();
                    int qtyZfile = Convert.ToInt16(workRow[dt.Columns.IndexOf(Root.strTQTY)]);
                    if (!vault1.IsLoggedIn) { vault1.LoginAuto(Root.pdmName, 0); }
                    IEdmFile7 zFile = (IEdmFile7)vault1.GetFileFromPath(pathFile, out IEdmFolder5 zFolder);

                    if (zFile != null) { BOM(zFile, conf, vers, 0, qtyZfile, ref dt); }//1//(int)EdmBomFlag.EdmBf_AsBuilt + //2// (int)EdmBomFlag.EdmBf_ShowSelected);
                }
            }

       

  

         
        
     

        public void SldOpenFile()
        {
            try
            {

                SldApp sldApp = null;
                sldApp = new SldApp();
                sldApp.AddDrawingsToBatchGet();
                sldApp.BatchGet();
                sldApp.Metod();
                sldApp.DrawingsBatchUnLock();
            }
            catch (Exception)
            {

                throw;
            }

        }
    }
