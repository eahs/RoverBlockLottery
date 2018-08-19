using System;

namespace RoverBlock
{
    [Serializable]
    public class Block
    {
        public string ID;
        public string Name;
        public int Slots;

        public Block(string ID, string Name, int slots)
        {
            this.ID = ID;
            this.Name = Name;
            this.Slots = slots;
        }
    }
}
