namespace EPPlus.ComponentModel.Exceptions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.Serialization;
    using System.Text;
    using System.Threading.Tasks;

    [Serializable]
    public class SheetNameExistsException : Exception
    {
        //
        // For guidelines regarding the creation of new exception types, see
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp
        // and
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp
        //

        public SheetNameExistsException()
        {
            
        }

        public SheetNameExistsException(string message) : base(message)
        {
        }

        public SheetNameExistsException(string message, Exception inner) : base(message, inner)
        {
        }

        protected SheetNameExistsException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
