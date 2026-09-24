using System;
using System.Collections.Generic;
using System.Linq;

using Exceltk.Format.Plugins;

namespace Exceltk.Format {
    /// <summary>
    /// Discovers and resolves format plugins by CLI target name.
    /// </summary>
    public static class FormatRegistry {
        private static readonly object Gate = new object();
        private static Dictionary<string, IFormatPlugin> _plugins;

        public static IEnumerable<IFormatPlugin> All {
            get {
                EnsureRegistered();
                return _plugins.Values.OrderBy(p => p.Name, StringComparer.Ordinal);
            }
        }

        public static bool TryGet(string name, out IFormatPlugin plugin) {
            EnsureRegistered();
            if (string.IsNullOrEmpty(name)) {
                plugin = null;
                return false;
            }
            return _plugins.TryGetValue(name.Trim().ToLowerInvariant(), out plugin);
        }

        public static IFormatPlugin GetRequired(string name) {
            if (!TryGet(name, out IFormatPlugin plugin)) {
                throw new ArgumentException("Unknown format plugin: " + name);
            }
            return plugin;
        }

        public static void Register(IFormatPlugin plugin) {
            if (plugin == null) {
                throw new ArgumentNullException(nameof(plugin));
            }
            EnsureRegistered();
            lock (Gate) {
                _plugins[plugin.Name.ToLowerInvariant()] = plugin;
            }
        }

        private static void EnsureRegistered() {
            if (_plugins != null) {
                return;
            }
            lock (Gate) {
                if (_plugins != null) {
                    return;
                }
                var map = new Dictionary<string, IFormatPlugin>(StringComparer.OrdinalIgnoreCase);
                foreach (IFormatPlugin plugin in BuiltInPlugins()) {
                    map[plugin.Name.ToLowerInvariant()] = plugin;
                }
                _plugins = map;
            }
        }

        private static IEnumerable<IFormatPlugin> BuiltInPlugins() {
            yield return new MarkdownFormatPlugin();
            yield return new JsonFormatPlugin();
            yield return new TexFormatPlugin();
            yield return new MarkedImageFormatPlugin();
        }
    }
}
