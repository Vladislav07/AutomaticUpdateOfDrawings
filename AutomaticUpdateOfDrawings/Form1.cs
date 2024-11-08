using System;
using System.Data;
using EPDM.Interop.epdm;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace AutomaticUpdateOfDrawings
{
    public partial class Form1 : Form
    {
        public string pathname0;

        public System.Data.DataTable dt2;

        IEdmFile7 aFile;
        IEdmFolder5 aFolder;
        IEdmBomMgr bomMgr;
   

        public IEdmVault5 vault1 = new EdmVault5();
        public IEdmVault7 vault2 = null;

        int rowerfilt = 0;

        public Form1()
        {    
            vault2 = (IEdmVault7)vault1;

            if (!vault1.IsLoggedIn) { vault1.LoginAuto(Root.pdmName, this.Handle.ToInt32()); }

            aFile = (IEdmFile7)vault1.GetObject(EdmObjectType.EdmObject_File, Root.ASMID);
            aFolder = (IEdmFolder5)vault1.GetObject(EdmObjectType.EdmObject_Folder, Root.ASMFolderID);

            pathname0 = aFolder.LocalPath;
            this.Text = (pathname0 + "\\" + aFile.Name);

            bomMgr = (IEdmBomMgr)vault2.CreateUtility(EdmUtility.EdmUtil_BomMgr);
           // bomMgr.GetBomLayouts(out EdmBomLayout[] ppoRetLayouts);       
        }
      
   


        void Data_output()
        {
        

            BOM_dt.BOM(aFile, config, version, 2, 1, ref dt2);///1//(int)EdmBomFlag.EdmBf_AsBuilt + //2// (int)EdmBomFlag.EdmBf_ShowSelected);
    
   
        }
     


        public void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                Data_output();
            }

            catch (System.Runtime.InteropServices.COMException ex)
            { MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

        }

 
    }
}
