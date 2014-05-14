using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;

namespace GorillaDocs
{
    public class StringEnumeratorWrapper : IEnumerator<string>
    {
        readonly StringEnumerator underlyingEnumerator;

        public StringEnumeratorWrapper(StringEnumerator underlyingEnumerator)
        {
            this.underlyingEnumerator = underlyingEnumerator;
        }

        public string Current
        {
            get
            {
                return this.underlyingEnumerator.Current;
            }
        }

        public void Dispose()
        {
            // No-op 
        }

        object System.Collections.IEnumerator.Current
        {
            get
            {
                return this.underlyingEnumerator.Current;
            }
        }

        public bool MoveNext()
        {
            return this.underlyingEnumerator.MoveNext();
        }

        public void Reset()
        {
            this.underlyingEnumerator.Reset();
        }
    }
}
