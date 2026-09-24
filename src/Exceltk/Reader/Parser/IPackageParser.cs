using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Byte-stream package framer (XUdt-style). Accumulates header → create empty package
    /// by type → accumulate body → <see cref="IPackage.DecodeBody"/> → queue.
    /// Does not invent field layout; does not own workbook semantics.
    /// </summary>
    public interface IPackageParser<TPackage> where TPackage : IPackage {
        /// <summary>Feed the next chunk of the wire stream.</summary>
        void PushData(byte[] data, int offset, int count);

        /// <summary>True when at least one fully decoded package is queued.</summary>
        bool HavePackage {
            get;
        }

        /// <summary>Dequeue the next decoded package. Caller must check <see cref="HavePackage"/>.</summary>
        TPackage PopPackage();

        /// <summary>True when the framer still holds incomplete header/body bytes.</summary>
        bool HaveUnParsedData {
            get;
        }

        /// <summary>Clear queue, partial buffer, and framing state.</summary>
        void Reset();
    }
}
