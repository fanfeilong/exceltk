using System.Collections.Generic;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Streaming parser: recognize local format units and emit packages as soon as each is complete.
    /// Intended for progressive parse/render (e.g. bytes arriving over the network).
    /// </summary>
    public interface IPackageParser<out TPackage> where TPackage : IPackage {
        /// <summary>Pull all remaining packages from the current cursor.</summary>
        IEnumerable<TPackage> Parse();
    }

    /// <summary>Pull-style extension for parsers that can emit one package at a time.</summary>
    public interface IPullPackageParser<TPackage> : IPackageParser<TPackage> where TPackage : IPackage {
        bool TryRead(out TPackage package);
    }
}
