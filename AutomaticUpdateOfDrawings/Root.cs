using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
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
        public static string pdmName;
        public static string strFullBOM;
        public static string strFileName;
        public static string strDraw;           //= "Drawing";
        public static string strDrawState;      //= "Drawing State";
        public static string strRev;
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
