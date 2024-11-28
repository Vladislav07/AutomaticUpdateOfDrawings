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
            ppoData[0].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\14 Участок изготовления сэндвич панелей\Кран балка для штрипсы\CUBY-00266509.SLDASM";
            ppoData[0].mbsStrData2 = @"Pending Express Manufacturing";
            ppoData[1].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\14 Участок изготовления сэндвич панелей\Кран балка для штрипсы\CUBY-00266731.SLDASM";
            ppoData[1].mbsStrData2 = @"Pending Express Manufacturing";
            ppoData[2].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\14 Участок изготовления сэндвич панелей\Кран балка для штрипсы\CUBY-00266509.SLDDRW";
            ppoData[2].mbsStrData2 = @"Pending Express Manufacturing";
            ppoData[3].mbsStrData1 = @"C:\CUBY_PDM\Work\Other\Без проекта\CUBY-V1.1\CAD\Завод контейнер\14 Участок изготовления сэндвич панелей\Кран балка для штрипсы\CUBY-00266731.SLDDRW";
            ppoData[3].mbsStrData2 = @"Pending Express Manufacturing";

            Root test = new Root();
            test.OnCmd(poCmd, ppoData);
        }
    }
}
