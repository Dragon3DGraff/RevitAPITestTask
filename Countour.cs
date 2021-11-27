using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BergmannTestTask
{
    class Countour
    {
        static int n = 4;

        static List<XYZ> _countour = new List<XYZ>(n)
  {
    new XYZ( 0 , -75 , 0 ),
    new XYZ( 508, -75 , 0 ),
    new XYZ( 508, 75 , 0 ),
    new XYZ( 0, 75 , 0 )
  };

        const double _thicknessMm = 20.0;
    }
}
