using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoverBlock
{
    class Block
    {
        public String ID;
        public String Name;
        public int aSlots;
        public int bSlots;

        public Block(String ID, String Name, int aSlots, int bSlots)
        {
            this.ID = ID;
            this.Name = Name;
            this.aSlots = aSlots;
            this.bSlots = bSlots;
        }
    }
}
