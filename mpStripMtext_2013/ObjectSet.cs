namespace mpStripMtext
{
    using System.Collections.Generic;
    using Autodesk.AutoCAD.DatabaseServices;

    public class ObjectSet
    {
        public List<MText> Mtextobjlst = new List<MText>();
        public List<MLeader> Mldrobjlst = new List<MLeader>();
        public List<AttributeReference> Mattobjlst = new List<AttributeReference>();
        public List<Dimension> Dimobjlst = new List<Dimension>();
        public List<Table> Tableobjlst = new List<Table>();

        public bool IsEmpty
        {
            get
            {
                if (Mtextobjlst.Count == 0 &&
                    Mldrobjlst.Count == 0 &&
                    Mattobjlst.Count == 0 &&
                    Dimobjlst.Count == 0 &&
                    Tableobjlst.Count == 0)
                    return true;

                return false;
            }
        }
    }
}
