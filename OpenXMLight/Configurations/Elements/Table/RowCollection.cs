using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class RowCollection : ICollection<Row>
    {
        public int Count => rows.Count();

        public bool IsReadOnly => throw new NotImplementedException();


        private List<Row> rows;

        public Row this[int index]
        {
            get => rows[index];
            set => rows[index] = value;
        }

        public RowCollection()
        {
            rows = new List<Row>();
        }

        public RowCollection(params CellCollection[] cellColletion)
        {
            rows = new List<Row>();

            var tmpCollection = cellColletion.ToList();
            foreach(var colletionItem in tmpCollection)
            {
                Row row = new Row();
                row.Cells = colletionItem;

                rows.Add(row);
            }
        }

        internal RowCollection(IEnumerable<OpenXML.TableRow> rows)
        {
            //this.rows = rows.ToList();
        }

        #region functions
        public void Add(Row item) => rows.Add(item);

        public void Clear() => rows.Clear();

        public bool Contains(Row item) => rows.Contains(item);

        public void CopyTo(Row[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<Row> GetEnumerator() => rows.GetEnumerator();

        public bool Remove(Row item) => rows.Remove(item);

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        #endregion
    }
}
