using EPDM.Interop.epdm;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
    public class SldApp
    {
        SldWorks swApp;


        IEdmBatchGet batchGetter;
        static EdmSelItem[] ppoSelection = null;
        static IEdmVault5 vault1 =null;
        IEdmVault7 vault2 = null;
        IEdmBatchUnlock2 batchUnlocker;
       

        public SldApp()
        {
           // IEdmVault5 vault1 = vau;
           // IEdmVault7 vault2 = (IEdmVault7)vault1;
            swApp = new SldWorks();
            swApp.Visible = true;
            ppoSelection = new EdmSelItem[Root.drawings.Count];
            Run();
        }

        public void Run()
        {
            Message();
            AddSelItemToList();
            DrawingsBatchGet();
            OpenAndRefresh();
            DrawingsBatchUnLock();
           // swApp.ExitApp();

        }

        void AddSelItemToList()
        {
            int i = 0;

            try
            {
                Array.Resize(ref ppoSelection, Root.drawings.Count);
                foreach (Drawing item in Root.drawings)
                {
                    ppoSelection[i] = new EdmSelItem();
                    ppoSelection[i].mlDocID = item.ID_File;
                    ppoSelection[i].mlProjID = item.ID_Folder;
                    i++;
                }

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void Message()
        {
            string str = "Файлы для перестроения:" + '\n';
            foreach (Drawing item in Root.drawings)
            {
                str = str +  Path.GetFileName(item.NameDraw) + '\n';
            }
            MessageBox.Show(str);
        }

        public void DrawingsBatchGet()
        {
            try
            {
               


                batchGetter = (IEdmBatchGet)Root.v2.CreateUtility(EdmUtility.EdmUtil_BatchGet);
                batchGetter.AddSelection((EdmVault5)Root.v, ref ppoSelection);
                if ((batchGetter != null))
                {

                    batchGetter.CreateTree(0, (int)EdmGetCmdFlags.Egcf_Lock + (int)EdmGetCmdFlags.Egcf_SkipOpenFileChecks);// + (int)EdmGetCmdFlags.Egcf_IncludeAutoCacheFiles);
                   bool retVal = batchGetter.ShowDlg(0);
                    if ((retVal))
                    {
                        batchGetter.GetFiles(0, null);
                    }

                }

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void DrawingsBatchUnLock()
        {   
            try
            {
               
                batchUnlocker = (IEdmBatchUnlock2)Root.v2.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);
                batchUnlocker.AddSelection((EdmVault5)Root.v, ref ppoSelection);
                batchUnlocker.CreateTree(0, (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock);

                batchUnlocker.Comment = "Refresh";
                bool retVal = batchUnlocker.ShowDlg(0);
                if ((retVal))
                {
                    batchUnlocker.UnlockFiles(0, null);
                }


            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

      
     

        public void OpenAndRefresh()
        {
            ModelDoc2 swModelDoc = default(ModelDoc2);
            int errors = 0;
            int warnings = 0;
            int lErrors = 0;
            int lWarnings = 0;
            ModelDocExtension extMod;
            string fileName = null;
            DrawingDoc swDraw = default(DrawingDoc);
            object[] vSheetName = null;
            string sheetName;
            int i = 0;
            bool bRet = false;

            try
            {
                foreach (Drawing item in Root.drawings)
                {
                    fileName = item.NameDraw;
                    swModelDoc = (ModelDoc2)swApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                    swDraw = (DrawingDoc)swModelDoc;
                    extMod = swModelDoc.Extension;
                    vSheetName = (object[])swDraw.GetSheetNames();
                    for (i = 0; i < vSheetName.Length; i++)

                    {

                        sheetName = (string)vSheetName[i];

                        bRet = swDraw.ActivateSheet(sheetName);

                         extMod.Rebuild((int)swRebuildOptions_e.swCurrentSheetDisp);
                        swModelDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, ref lErrors, ref lWarnings);
                        Sheet swSheet = default(Sheet);

                         swSheet = (Sheet)swDraw.GetCurrentSheet();
                         MessageBox.Show(sheetName);

                    }
 
                    swModelDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, ref lErrors, ref lWarnings);
                    MessageBox.Show(lWarnings.ToString());
                    swApp.CloseDoc(fileName);
                    swModelDoc = null;

                }
            }
            catch (Exception)
            {
                MessageBox.Show(errors.ToString());

            }
        }
    }
}
