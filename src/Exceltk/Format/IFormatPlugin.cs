using System.Collections.Generic;
using System.IO;

using Exceltk.Reader;

namespace Exceltk.Format {
    /// <summary>
    /// Pluggable conversion format. Every format supports export (Excel → format)
    /// and import (format → Excel workbook DataSet).
    /// </summary>
    public interface IFormatPlugin {
        /// <summary>CLI target id, e.g. md / json / tex / img.</summary>
        string Name { get; }

        /// <summary>Default file extension without dot.</summary>
        string FileExtension { get; }

        /// <summary>Human-readable summary for help text.</summary>
        string Description { get; }

        bool SupportsExport { get; }
        bool SupportsImport { get; }

        IEnumerable<FormatArtifact> Export(DataSet dataSet, string sheetFilter);
        DataSet Import(Stream input, string sourcePath);
    }
}
