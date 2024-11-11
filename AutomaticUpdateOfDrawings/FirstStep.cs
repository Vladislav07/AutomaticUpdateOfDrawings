using System;
using EPDM.Interop.epdm;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
    public class FirstStep
    {
        private IEdmVault5 vault1 = null;
        string path;
        public FirstStep()
        {
    
            try
            {
            

                path = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\Склад из контейнеров\Лестница\CUBY-00259479.sldasm";
               Root ass = new Root();
                ass.OnCmd(path);

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
    }
}
