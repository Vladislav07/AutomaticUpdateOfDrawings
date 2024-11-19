using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
    public class Root : IEdmAddIn5
    {

        public static int ASMID;
        public static int ASMFolderID;
        public static string name0;
        public static string pdmName = "CUBY_PDM";
        public static string strFullBOM = "FullBOM";
        public static string strConfig = "Configuration";
        public static string strFoundIn = "Found_In"; 
        public static string strFileName = "File_Name";
        public static string strLatestVer = "Latest_Version";
        public static string strConfigPoint = ".";
        public static string strSection = "Section";
        public static EdmSelItem[] ppoSelection;
        public static List<EdmSelItem> SelectionDrawings;
        public static List<string> listdrawings;
        public static IEdmBatchUnlock2 batchUnlocker;
        public static List<Drawing> drawings;
        public IEdmVault7 vault2 = null;
        IEdmFile7 aFile;
        IEdmFolder5 aFolder;
        string config;
        int version = -1;
        public string pathname0;
      
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            //Specify information to display in the add-in's Properties dialog box
            poInfo.mbsAddInName = "AutomaticUpdateOfDrawings_1.0.0";
            poInfo.mbsCompany = "CUBY";
            poInfo.mbsDescription = "Create specification FullBOM";
            poInfo.mlAddInVersion = 050419;
            poInfo.mlRequiredVersionMajor = 27;
            poInfo.mlRequiredVersionMinor = 1;

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PreState);
           
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            string FileName;
            string e;
            IEdmFile5 file5 = null;
            IEdmFolder5 folder5 = null;
            EdmVault5 v = new EdmVault5();
            drawings = new List<Drawing>();

            try
            {
                switch (poCmd.meCmdType)
                {

                    case EdmCmdType.EdmCmd_PreState:
                        foreach (EdmCmdData AffectedFile in ppoData)
                        {

                            if (AffectedFile.mbsStrData2 == "Pending Express Manufacturing")
                            {
                                
                                FileName = ((EdmCmdData)ppoData.GetValue(0)).mbsStrData1;
                                e = System.IO.Path.GetExtension(FileName);
                                name0 = System.IO.Path.GetFileNameWithoutExtension(FileName);
                                ASMID = ((EdmCmdData)ppoData.GetValue(0)).mlObjectID1;
                                ASMFolderID = ((EdmCmdData)ppoData.GetValue(0)).mlObjectID2;

                                if (!v.IsLoggedIn)
                                {
                                    v.LoginAuto("CUBY_PDM", 0);
                                    // MessageBox.Show("Ok!!!");
                                }

                                file5 = v.GetFileFromPath(FileName, out folder5);

                                if ((e == ".sldasm") || (e == ".SLDASM"))   //replace slddrw
                                {


                                    ASMID = file5.ID;
                                    ASMFolderID = folder5.ID;


                                    vault2 = (IEdmVault7)v;

                                    aFile = (IEdmFile7)vault2.GetObject(EdmObjectType.EdmObject_File, ASMID);
                                    aFolder = (IEdmFolder5)vault2.GetObject(EdmObjectType.EdmObject_Folder, ASMFolderID);
                                    config = strConfig;
                                    version = aFile.CurrentVersion;
                                    pathname0 = aFolder.LocalPath;

                                    BOM_dt.BOM(aFile, config, version, 2);
                                }
                            }

                        }

                        break;
                    //The event isn't registered
                    default:
                        ((EdmVault5)(poCmd.mpoVault)).MsgBox(poCmd.mlParentWnd, "An unknown command type was issued.");
                        break;
                }
            }


            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
    }
}
