using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPDM.Interop.epdm;

namespace AutomaticUpdateOfDrawings
{
    public class CollBack:IEdmCallback
    {
        public void SetModifiedFlag()
        {
            throw new NotImplementedException();
        }

        public void SetProgressRange(int lMin, int lMax)
        {
            throw new NotImplementedException();
        }

        public void SetProgressPos(int lPos)
        {
            throw new NotImplementedException();
        }

        public void SetStatusMessage(string bsMessage)
        {
            throw new NotImplementedException();
        }
    }
}
