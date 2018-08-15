using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace libMIDI
{
    public interface ISimpleNote
    {
        string Pitch { set; get; }
        int Octave { set; get; }
        byte Velocity { set; get; }
    }

}
