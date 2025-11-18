using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class CellCollection: ICollection<Cell>
    {
        public int Count => cells.Count();
        public bool IsReadOnly => throw new NotImplementedException();


        private List<Cell> cells;
                
        public Cell this[int index]
        {
            get => cells[index];
            set => cells[index] = value;
        }
        

        public CellCollection(params Cell[] cells) => this.cells = cells.ToList();

        internal CellCollection(IEnumerable<OpenXML.TableCell> cells)
        {
            //this.cells = cells.ToList();
        }

        #region functions
        public void Add(Cell item) => cells.Add(item);

        public void Clear() => cells.Clear();

        public bool Contains(Cell item) => cells.Contains(item);

        public void CopyTo(Cell[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<Cell> GetEnumerator() => cells.GetEnumerator();

        public bool Remove(Cell item) => cells.Remove(item);

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        #endregion
    }
}
