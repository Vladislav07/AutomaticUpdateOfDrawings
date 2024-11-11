using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
   public class Root
    {

        public static int ASMID;
        public static int ASMFolderID;
        public static string name0;
        public static string pdmName = "CUBY_PDM";
        public static string strFullBOM = "FullBOM";
        public static string strConfig = ".";
        public static string strFoundIn = "Found In"; 
        public static string strFileName = "File_Name";
        public static string strDrawState = "Drawing State";
        public static string strNeedsRegeneration = "NeedsRegeneration";
        public static string strLatestVer = "Latest Version";
        public static string strRev = "RefVersion";
        public static string strPartNumber = "Обозначение";
        public static string strDescription_RUS = "Наименование";
        public static string strConfigPoint = ".";
        public static string strErC;
        public static string strFileID;
        public static string strFolderID;
        public static string strSection = "Раздел";
        public static EdmSelItem[] ppoSelection;
        public static List<EdmSelItem> SelectionDrawings;
        public static List<string> listdrawings;
        public static IEdmBatchUnlock2 batchUnlocker;
   

        public void OnCmd(string pathAss)
        {
            string FileName = pathAss;
            string e = System.IO.Path.GetExtension(FileName);
            name0 = System.IO.Path.GetFileNameWithoutExtension(FileName);

            IEdmFile5 file5 = null;
            IEdmFolder5 folder5 = null;
            EdmVault5 v = new EdmVault5();
            try 
                { 
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

                        Application.Run(new Form1());

                    }
                    else
                    {

                        MessageBox.Show("Select the assembly model file (.SLDASM)");
                        Application.Exit();
                    }
                }

    

            catch (System.Runtime.InteropServices.COMException ex)
                { MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message); }
            catch (Exception ex)
                { MessageBox.Show(ex.Message);
            }
        }
    }
}
