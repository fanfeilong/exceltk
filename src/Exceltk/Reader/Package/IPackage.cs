namespace Exceltk.Reader.Package {
    /// <summary>
    /// A complete local format unit that a streaming parser can emit as soon as
    /// enough bytes/tokens arrive (e.g. one BIFF record, one XML cell/row).
    /// Suitable for progressive render over a network transfer.
    /// </summary>
    public interface IPackage {
    }

    /// <summary>
    /// Binary (BIFF) package base. Concrete record types in <c>Reader.Binary</c> inherit this.
    /// </summary>
    public abstract class BinaryPackage : IPackage {
    }

    /// <summary>
    /// OpenXML / worksheet XML package base. Concrete units live in <c>Reader.Xml</c>.
    /// </summary>
    public abstract class XmlPackage : IPackage {
    }
}
