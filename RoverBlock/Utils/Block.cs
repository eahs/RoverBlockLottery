using System;

namespace RoverBlock
{
    [Serializable]
    public class Block
    {
        public string ID;
        public string Name;
        public int aSlots;
        public int bSlots;

        public Block(string ID, string Name, int aSlots, int bSlots)
        {
            this.ID = ID;
            this.Name = Name;
            this.aSlots = aSlots;
            this.bSlots = bSlots;
        }
    }
}
