using System.Collections.Generic;
using System.IO;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Streaming parser: reads an input and yields package entities incrementally.
    /// </summary>
    public interface IPackageParser<out TPackage> where TPackage : IPackage {
        /// <summary>
        /// Parse from the current position, yielding packages one at a time.
        /// </summary>
        IEnumerable<TPackage> Parse();
    }
}
