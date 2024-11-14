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

        public System.Data.DataTable dt;

        IEdmFile7 aFile;
        IEdmFolder5 aFolder;
        IEdmBomMgr bomMgr;
        string config;
        int version = -1;

        public IEdmVault5 vault1 = new EdmVault5();
        public IEdmVault7 vault2 = null;

        int rowerfilt = 0;

        public Form1()
        {    
           

            if (!vault1.IsLoggedIn) { vault1.LoginAuto(Root.pdmName, this.Handle.ToInt32()); }
             vault2 = (IEdmVault7)vault1;

            aFile = (IEdmFile7)vault1.GetObject(EdmObjectType.EdmObject_File, Root.ASMID);
            aFolder = (IEdmFolder5)vault1.GetObject(EdmObjectType.EdmObject_Folder, Root.ASMFolderID);
            config = Root.strConfig;
            version = aFile.CurrentVersion;
            pathname0 = aFolder.LocalPath;
            dt = new System.Data.DataTable();
            BOM_dt.BOM(aFile, config, version, 1);///1//(int)EdmBomFlag.EdmBf_AsBuilt + //2// (int)EdmBomFlag.EdmBf_ShowSelected);
        }
      
   


        void Data_output()
        {

            this.Cursor = Cursors.WaitCursor;
            
            this.bindingSource1.DataSource = dt;
           
            this.Cursor = Cursors.Arrow;

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
