using System;
using System.Runtime.Serialization;

namespace ClosedXML.Report
{
    public class InvalidNestedRangeException : Exception
    {
        public InvalidNestedRangeException()
        {
        }

        protected InvalidNestedRangeException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public InvalidNestedRangeException(string message) : base(message)
        {
        }

        public InvalidNestedRangeException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
