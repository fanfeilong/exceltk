namespace Exceltk.Reader.Package {
    /// <summary>
    /// Package identified by a filesystem path (e.g. CSV).
    /// </summary>
    public interface IPathPackage : IExcelPackage {
        string Path {
            get;
        }
    }
}
