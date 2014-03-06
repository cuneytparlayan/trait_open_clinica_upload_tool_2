using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace OCDataImporter
{
    /// <summary>
    /// Custom exception for the data importer
    /// </summary>     
    class OCDataImporterException : Exception, ISerializable
    {
        public OCDataImporterException() : base()
        {
            // Add implementation.
        }
        public OCDataImporterException(string message) : base(message)
        {
            // Add implementation.
        }
        public OCDataImporterException(string message, Exception inner) : base(message, inner)
        {
            // Add implementation.
        }

        // This constructor is needed for serialization.
        protected OCDataImporterException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            // Add implementation.
        }
    }
}
