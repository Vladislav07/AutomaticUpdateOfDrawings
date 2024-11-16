using System;
using EPDM.Interop.epdm;
using System.Windows.Forms;

namespace AutomaticUpdateOfDrawings
{
    public class FirstStep
    {
        string path;
        public FirstStep()
        {
    
            try
            {

                path = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\Участок сварочный\Кран балка\CUBY-00170130.sldasm";
              //  path = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\Склад из контейнеров\Лестница\CUBY-00259479.sldasm";
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
