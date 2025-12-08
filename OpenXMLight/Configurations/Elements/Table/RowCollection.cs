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

        public bool IsReadOnly => false;


        private List<Row> rows;
        internal Table? ParentTable { get; set; }
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
            this.rows = new();
            foreach(OpenXML.TableRow rowXml in rows)
            {
                Row row = new(rowXml);
                
                this.rows.Add(row);
            }
        }

        #region functions
        public void Add(Row item)
        {
            ParentTable?.TableXml.AppendChild(item.RowXml);
            
            rows.Add(item);
        }

        public void Clear()
        {
            ParentTable?.TableXml.RemoveAllChildren<OpenXML.TableRow>();
            
            rows.Clear();
        }

        public bool Contains(Row item) => rows.Contains(item);

        public void CopyTo(Row[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<Row> GetEnumerator() => rows.GetEnumerator();

        public bool Remove(Row item)
        {
            ParentTable?.TableXml.RemoveChild(item.RowXml);
            
            rows.Remove(item);

            return true;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        #endregion
    }
}
