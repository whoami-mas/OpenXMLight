using OpenXMLight.Configurations.Elements.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlElement = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements
{
    public class ElementCollection<T> : ICollection<T> where T : IElement
    {
        public int Count => _collection.Count;
        public bool IsReadOnly => false;

        public T this[int index]
        {
            get => _collection[index];
        }


        List<T> _collection = new();
        internal OpenXml.OpenXmlElement Parent { get; set; }
        internal ElementCollection(IEnumerable<T> _collection) => this._collection = _collection.ToList();


        public void Add(T item)
        {
            Parent.AppendChild(item.XmlElement);
            
            _collection.Add(item);
        }
        public void Add(T item, bool cloneNode = false)
        {
            if (cloneNode)
            {
                Parent.AppendChild(item.XmlElement.CloneNode(true));

                _collection.Add(item);
            }
            else
                Add(item);
        }
        public void Clear()
        {
            try
            {
                var typeOpenXml = GetTypeOpenXml();

                if (typeOpenXml == null)
                    throw new ArgumentException("Неизвестный тип");


                var method = Parent?.GetType()?.GetMethods()
                    .FirstOrDefault(f => f.Name == "RemoveAllChildren" &&
                                         f.IsGenericMethodDefinition)?
                    .MakeGenericMethod(typeOpenXml);

                method?.Invoke(Parent, null);

                if (string.Equals(Parent.LocalName, "tc"))
                    Parent.AppendChild(new OpenXmlElement.Paragraph());

                _collection.Clear();
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public bool Contains(T item) => _collection.Contains(item);

        public void CopyTo(T[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<T> GetEnumerator() => _collection.GetEnumerator();

        public bool Remove(T item)
        {
            try
            {
                Parent.RemoveChild(item.XmlElement);

                _collection.Remove(item);
                return true;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private Type? GetTypeOpenXml()
        {
            var currentType = typeof(T);
            var baseType = currentType.BaseType;

            if(baseType != null && baseType.IsGenericType)
            {
                var genericArgs = baseType.GetGenericArguments();

                if(genericArgs.Length >= 1)
                    return genericArgs[0];
            }

            return null;
        }
    }
}
