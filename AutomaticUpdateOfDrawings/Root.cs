using EPDM.Interop.epdm;
using System.IO;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
    public class Root : IEdmAddIn5
    {
        public static EdmVault5 v = null;
        public static EdmSelItem[] ppoSelection;
        public static List<EdmSelItem> SelectionDrawings;
        public static List<string> listdrawings;
        public static IEdmBatchUnlock2 batchUnlocker;
        public static List<Drawing> drawings;
        public IEdmVault7 vault2 = null;
        public string pathname0;

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            //Specify information to display in the add-in's Properties dialog box
            poInfo.mbsAddInName = "AutomaticUpdateOfDrawings_1.0.0";
            poInfo.mbsCompany = "CUBY";
            poInfo.mbsDescription = "Rebuild drawings";
            poInfo.mlAddInVersion = 050419;
            poInfo.mlRequiredVersionMajor = 27;
            poInfo.mlRequiredVersionMinor = 1;

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PreState);

        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            string FileName;
            string e;
            string designation;
            Dictionary<string, string> model = new Dictionary<string, string>();
            Dictionary<string, string> drawing = new Dictionary<string, string>();

            v = (EdmVault5)poCmd.mpoVault;
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

                                FileName = AffectedFile.mbsStrData1;
                                e = Path.GetExtension(FileName);
                                designation = Path.GetFileNameWithoutExtension(FileName);

                                string regCuby = @"^CUBY-\d{8}$";
                                bool IsCUBY = Regex.IsMatch(designation, regCuby);
                                if (!IsCUBY) continue;

                                switch (e)
                                {
                                    case "sldasm":
                                        model.Add(designation, FileName);
                                        break;
                                    case ".sldprt":
                                        model.Add(designation, FileName);
                                        break;
                                    case "SLDASM":
                                        model.Add(designation, FileName);
                                        break;
                                    case "SLDPRT":
                                        model.Add(designation, FileName);
                                        break;
                                    case ".SLDDRW":
                                        drawing.Add(designation, FileName);
                                        break;
                                    case ".slddrw":
                                        drawing.Add(designation, FileName);
                                        break;
                                    default:
                                        break;
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

            foreach (var item in drawing)
            {
                string key = item.Key;
                string valueD = item.Value;
                string valueM;
                bool isGet = false;
                if (model.ContainsKey(key))
                {
                    isGet = model.TryGetValue(key, out valueM);
                
                    if(isGet && valueM != null)
                    {
                        IsDrawingsRebuild(valueM, valueD);
                    }
                }
            }


        }

        void IsDrawingsRebuild(string p, string d)
        {
            int refDrToModel = -1;
            bool NeedsRegeneration = false;

            IEdmFile7 modelFile = null;
            IEdmFile7 bFile = null;

            Drawing draw = null;

            if ((modelFile != null) && (!modelFile.IsLocked))
            {
                modelFile = (IEdmFile7)v.GetFileFromPath(p, out IEdmFolder5 modelFolder);
                refDrToModel = modelFile.CurrentVersion;
            }


            bFile = (IEdmFile7)v.GetFileFromPath(d, out IEdmFolder5 bFolder);

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
    }
}
