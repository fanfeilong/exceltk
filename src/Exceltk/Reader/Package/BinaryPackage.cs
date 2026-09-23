using Exceltk.Reader.Binary;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// One BIFF record as a streamed package entity.
    /// </summary>
    internal sealed class BinaryPackage : IPackage {
        public BinaryPackage(XlsBiffRecord record) {
            Record = record;
        }

        public XlsBiffRecord Record {
            get;
            private set;
        }

        public BIFFRECORDTYPE Id {
            get {
                return Record.ID;
            }
        }
    }
}
