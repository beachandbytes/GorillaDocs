using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;

namespace GorillaDocs
{
    public class StringCollectionEnumerable : IEnumerable<string>
    {
        readonly StringCollection underlyingCollection;

        public StringCollectionEnumerable(StringCollection underlyingCollection)
        {
            if (underlyingCollection == null)
                underlyingCollection = new StringCollection();
            this.underlyingCollection = underlyingCollection;
        }

        public IEnumerator<string> GetEnumerator()
        {
            return new StringEnumeratorWrapper(underlyingCollection.GetEnumerator());
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
