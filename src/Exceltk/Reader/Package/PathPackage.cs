using System;
using System.IO;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// Package that points at a file path without loading the whole file into memory.
    /// </summary>
    public sealed class PathPackage : IPathPackage {
        public PathPackage(string path) {
            if (string.IsNullOrEmpty(path)) {
                throw new ArgumentException("path");
            }
            Path = path;
            if (File.Exists(path)) {
                IsValid = true;
                ExceptionMessage = null;
            } else {
                IsValid = false;
                ExceptionMessage = "File not found: " + path;
            }
        }

        public string Path {
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
            // Nothing to release; path is metadata only.
        }
    }
}
