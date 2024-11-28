using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomaticUpdateOfDrawings
{
   public class Program
    {
      static IEdmVault5 vault1 = new EdmVault5();
      static EdmCmdData[] ppoData = new EdmCmdData[4];
      static EdmCmd poCmd = new EdmCmd();
      public  static void Main()
        {
           
            if (!vault1.IsLoggedIn) { vault1.LoginAuto("CUBY_PDM", 0); }

            poCmd.mpoVault = vault1;
            poCmd.meCmdType = EdmCmdType.EdmCmd_PreState;
            ppoData[0].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\Участок сварочный\Кран балка\CUBY-00170132.sldprt";
            ppoData[0].mbsStrData2 = @"Pending Express Manufacturing";
            ppoData[1].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\Участок сварочный\Кран балка\CUBY-00170132.SLDDRW";
            ppoData[1].mbsStrData2 = @"Pending Express Manufacturing";
            ppoData[2].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\Участок сварочный\Кран балка\CUBY-00170130.sldasm";
            ppoData[2].mbsStrData2 = @"Pending Express Manufacturing";
            ppoData[3].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\Участок сварочный\Кран балка\CUBY-00170130.SLDDRW";
            ppoData[3].mbsStrData2 = @"Pending Express Manufacturing";

            Root test = new Root();
            test.OnCmd(poCmd, ppoData);
        }
    }
}
