using System;
using myonemap.core;

namespace myonemap
{
    [global::System.Serializable]
    public class OneMapException : Exception
    {
        private string _InnerExceptionInfo = string.Empty;
        private Exception _BaseException = null;
        #region Constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public OneMapException() { }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="message"></param>
        public OneMapException(string message) : base(message) { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="message"></param>
        /// <param name="inner"></param>
        public OneMapException(string message, Exception inner)
            : base(message, inner)
        {
            _InnerExceptionInfo = inner.ToStringReflection();
        }

        public OneMapException(string message, Exception inner,Exception baseException)
            : base(message, inner)
        {
            _BaseException = baseException;
        }
        /// <summary>
        /// Serialization constructor
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected OneMapException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }

        #endregion

        public string InnerExceptionInfo
        {
            get { return _InnerExceptionInfo; }
        }

        public Exception BaseException
        {
            get { return _BaseException; }
        }

    }

}
