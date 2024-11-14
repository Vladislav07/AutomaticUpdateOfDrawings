﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomaticUpdateOfDrawings
{
    public class Drawing
    {
        public Drawing(int idFile, int idFolder, string n)
        {
            ID_File = idFile;
            ID_Folder = idFolder;
            NameDraw = n;    
        }

        public int ID_File { get; }
        public int ID_Folder { get; }
        public string NameDraw { get; }

    }
}
