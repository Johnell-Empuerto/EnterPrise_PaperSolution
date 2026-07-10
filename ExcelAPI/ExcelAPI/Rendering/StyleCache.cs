namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Thread-safe cache of resolved cell styles.
    /// Each CellFormat index is resolved exactly once and cached.
    /// All render layers consume ResolvedCellStyle from this cache.
    /// </summary>
    public class StyleCache
    {
        private readonly Dictionary<int, ResolvedCellStyle> _cache = new();
        private readonly object _lock = new();

        /// <summary>
        /// Get or create a resolved style for the given CellFormat index.
        /// </summary>
        public ResolvedCellStyle GetOrAdd(int styleIndex, Func<int, ResolvedCellStyle> factory)
        {
            lock (_lock)
            {
                if (_cache.TryGetValue(styleIndex, out var existing))
                    return existing;

                var resolved = factory(styleIndex);
                _cache[styleIndex] = resolved;
                return resolved;
            }
        }

        /// <summary>
        /// Try to get a previously resolved style.
        /// </summary>
        public bool TryGet(int styleIndex, out ResolvedCellStyle? style)
        {
            lock (_lock)
            {
                return _cache.TryGetValue(styleIndex, out style);
            }
        }

        /// <summary>
        /// Number of cached styles.
        /// </summary>
        public int Count
        {
            get { lock (_lock) { return _cache.Count; } }
        }

        /// <summary>
        /// Clear all cached styles (e.g., when switching workbooks).
        /// </summary>
        public void Clear()
        {
            lock (_lock)
            {
                _cache.Clear();
            }
        }
    }
}
