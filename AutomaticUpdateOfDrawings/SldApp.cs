using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;
using EPDM.Interop.epdm;
using Microsoft.VisualBasic;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace AutomaticUpdateOfDrawings
{
    public class SldApp
    {
        SldWorks swApp;


        IEdmBatchGet batchGetter;
        EdmSelItem[] ppoSelection = null;
        IEdmVault5 vault1 = new EdmVault5();
        IEdmVault7 vault2 = null;
        EdmSelectionObject poSel;
        bool retVal;
        IEdmBatchUnlock2 batchUnlocker;


        public SldApp()
        {
            swApp = new SldWorks();
            swApp.Visible = true;
            ppoSelection = new EdmSelItem[Root.drawings.Count];
        }

        public void Run()
        {
            Message();
        }

        void Message()
        {
            string str = "";
            foreach (Drawing item in Root.drawings)
            {
                str = str + Path.GetFileName(item.NameDraw);
            }
            MessageBox.Show(str);
        }

        public void AddDrawingsToBatchGet()
        {
            try
            {
                if (vault2 == null)
                {
                    ConnectPDM();
                }


                batchGetter = (IEdmBatchGet)vault2.CreateUtility(EdmUtility.EdmUtil_BatchGet);

                foreach (Drawing item in Root.drawings)
                {
                    batchGetter.AddSelectionEx((EdmVault5)vault1, item.ID_File, item.ID_Folder, 0);
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
            int i = 0;


            try
            {

                foreach (Drawing item in Root.drawings)
                {
                    ppoSelection[i] = new EdmSelItem();
                    ppoSelection[i].mlDocID = item.ID_File;
                    ppoSelection[i].mlProjID = item.ID_Folder;
                    i++;
                }

                batchUnlocker = (IEdmBatchUnlock2)vault2.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);
                batchUnlocker.AddSelection((EdmVault5)vault1, ref ppoSelection);
                batchUnlocker.CreateTree(0, (int)EdmUnlockBuildTreeFlags.Eubtf_Nothing);

                batchUnlocker.Comment = "Refresh";
                retVal = batchUnlocker.ShowDlg(0);
                object statuses = null;
                if ((retVal))
                {
                    batchUnlocker.UnlockFiles(0, null);
                    statuses = batchUnlocker.GetStatus((int)EdmUnlockStatusFlag.Eusf_CloseAfterCheckinFlag);
                    Interaction.MsgBox("Close Files after Check In selected? " + statuses);
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

        void ConnectPDM()
        {
            try
            {
                if (vault1 == null)
                {
                    vault1 = new EdmVault5();
                }


                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto(Root.pdmName, 0);
                }

                vault2 = (IEdmVault7)vault1;
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
        public void BatchGet()
        {

            try
            {
                if (vault2 == null)
                {
                    ConnectPDM();
                }

                if ((batchGetter != null))
                {

                    batchGetter.CreateTree(0, (int)EdmGetCmdFlags.Egcf_Lock + (int)EdmGetCmdFlags.Egcf_SkipOpenFileChecks);
                    retVal = batchGetter.ShowDlg(0);
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

        public void OpenAndRefresh()
        {
            ModelDoc2 swModelDoc = default(ModelDoc2);

            int errors = 0;
            int warnings = 0;
            int lErrors = 0;
            int lWarnings = 0;

            string fileName = null;

      
            try
            {
                foreach (Drawing item in Root.drawings)
                {
                    fileName = item.NameDraw;
                    swModelDoc = (ModelDoc2)swApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

                    swApp.CreateNewWindow();
                    swModelDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, ref lErrors, ref lWarnings);
                    swModelDoc.Close();
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
