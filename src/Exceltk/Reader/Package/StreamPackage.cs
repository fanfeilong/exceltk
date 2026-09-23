using System;
using System.IO;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// Thin package wrapper around an already-opened content stream.
    /// </summary>
    public sealed class StreamPackage : IStreamPackage {
        private readonly bool m_ownsStream;
        private bool m_disposed;

        public StreamPackage(Stream content, bool ownsStream = true) {
            if (content == null) {
                throw new ArgumentNullException("content");
            }
            Content = content;
            m_ownsStream = ownsStream;
            IsValid = true;
            ExceptionMessage = null;
        }

        public Stream Content {
            get;
            private set;
        }

        public bool IsValid {
            get;
            private set;
        }

        public string ExceptionMessage {
            get;
            private set;
        }

        public void Dispose() {
            if (m_disposed) {
                return;
            }
            m_disposed = true;
            if (m_ownsStream && Content != null) {
                Content.Dispose();
                Content = null;
            }
        }
    }
}
