using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Server-side session storage for uploaded Excel workbooks.
    ///
    /// When a user uploads a workbook, the server:
    ///   1. Saves the original XLSX to TempWorkbooks/{sessionId}/original.xlsx
    ///   2. Records metadata (original filename, capture result references)
    ///   3. Returns only the sessionId to the browser
    ///
    /// The browser never knows the physical file path.
    /// The browser never needs to remember a filename.
    ///
    /// This mirrors the legacy ConMas Designer behavior, where the server
    /// owned the workbook (stored in def_top_file BLOB) and the browser
    /// only tracked a defTopId/sessionId.
    ///
    /// Sessions are stored in memory and on disk. Sessions expire and can
    /// be cleaned up.
    /// </summary>
    public interface ISessionWorkbookStore
    {
        /// <summary>Create a new session and store the workbook.</summary>
        SessionInfo CreateSession(string originalFilePath, string uploadFileName);

        /// <summary>Get session info by ID. Returns null if expired/invalid.</summary>
        SessionInfo? GetSession(string sessionId);

        /// <summary>Get the full path to the original workbook for a session.</summary>
        string? ResolveWorkbookPath(string sessionId);

        /// <summary>Remove a session and its files.</summary>
        void RemoveSession(string sessionId);
    }

    /// <summary>
    /// Information about an active editing session.
    /// </summary>
    public class SessionInfo
    {
        /// <summary>Unique session identifier (GUID).</summary>
        public string SessionId { get; set; } = "";

        /// <summary>Full path to the saved original workbook.</summary>
        public string WorkbookPath { get; set; } = "";

        /// <summary>Original filename as uploaded by the user.</summary>
        public string UploadFileName { get; set; } = "";

        /// <summary>When the session was created.</summary>
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        /// <summary>When the session was last accessed.</summary>
        public DateTime LastAccessedAt { get; set; } = DateTime.UtcNow;
    }

    public class SessionWorkbookStore : ISessionWorkbookStore
    {
        private readonly ConcurrentDictionary<string, SessionInfo> _sessions = new();
        private readonly string _sessionsRoot;
        private readonly ILogger<SessionWorkbookStore> _logger;

        /// <summary>
        /// Default session lifetime: 24 hours.
        /// After this, the session is eligible for cleanup.
        /// </summary>
        public static readonly TimeSpan SessionLifetime = TimeSpan.FromHours(24);

        public SessionWorkbookStore(ILogger<SessionWorkbookStore> logger)
        {
            _logger = logger;
            _sessionsRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempWorkbooks");
            Directory.CreateDirectory(_sessionsRoot);
            _logger.LogInformation("SessionWorkbookStore initialized at {Root}", _sessionsRoot);
        }

        public SessionInfo CreateSession(string originalFilePath, string uploadFileName)
        {
            var sessionId = Guid.NewGuid().ToString("N");
            var sessionDir = Path.Combine(_sessionsRoot, sessionId);
            Directory.CreateDirectory(sessionDir);

            var targetPath = Path.Combine(sessionDir, "original.xlsx");

            // Copy the original workbook into the session folder
            File.Copy(originalFilePath, targetPath, overwrite: true);

            var info = new SessionInfo
            {
                SessionId = sessionId,
                WorkbookPath = targetPath,
                UploadFileName = uploadFileName,
                CreatedAt = DateTime.UtcNow,
                LastAccessedAt = DateTime.UtcNow
            };

            _sessions[sessionId] = info;

            _logger.LogInformation(
                "[SESSION] Created {SessionId} for '{FileName}' at {Path}",
                sessionId, uploadFileName, targetPath);

            return info;
        }

        public SessionInfo? GetSession(string sessionId)
        {
            if (string.IsNullOrEmpty(sessionId)) return null;

            if (_sessions.TryGetValue(sessionId, out var info))
            {
                // Check expiry
                if (DateTime.UtcNow - info.CreatedAt > SessionLifetime)
                {
                    _logger.LogWarning("[SESSION] {SessionId} expired (created {Created})", sessionId, info.CreatedAt);
                    RemoveSession(sessionId);
                    return null;
                }

                info.LastAccessedAt = DateTime.UtcNow;
                return info;
            }

            // Session not in memory — check disk (survives app restart for in-progress uploads)
            var sessionDir = Path.Combine(_sessionsRoot, sessionId);
            var workbookPath = Path.Combine(sessionDir, "original.xlsx");
            if (File.Exists(workbookPath))
            {
                var diskInfo = new SessionInfo
                {
                    SessionId = sessionId,
                    WorkbookPath = workbookPath,
                    UploadFileName = $"{sessionId}.xlsx",
                    CreatedAt = File.GetCreationTimeUtc(workbookPath),
                    LastAccessedAt = DateTime.UtcNow
                };

                // Check disk-based expiry
                if (DateTime.UtcNow - diskInfo.CreatedAt > SessionLifetime)
                {
                    _logger.LogWarning("[SESSION] {SessionId} disk session expired", sessionId);
                    RemoveSession(sessionId);
                    return null;
                }

                _sessions[sessionId] = diskInfo;
                return diskInfo;
            }

            _logger.LogWarning("[SESSION] {SessionId} not found", sessionId);
            return null;
        }

        public string? ResolveWorkbookPath(string sessionId)
        {
            return GetSession(sessionId)?.WorkbookPath;
        }

        public void RemoveSession(string sessionId)
        {
            _sessions.TryRemove(sessionId, out _);

            var sessionDir = Path.Combine(_sessionsRoot, sessionId);
            try
            {
                if (Directory.Exists(sessionDir))
                {
                    Directory.Delete(sessionDir, recursive: true);
                    _logger.LogInformation("[SESSION] Removed {SessionId}", sessionId);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[SESSION] Failed to clean up {SessionId}", sessionId);
            }
        }

        /// <summary>
        /// Clean up all expired sessions. Call periodically from a background task.
        /// </summary>
        public void CleanupExpiredSessions()
        {
            var cutoff = DateTime.UtcNow - SessionLifetime;
            var expired = _sessions.Values
                .Where(s => s.CreatedAt < cutoff)
                .ToList();

            foreach (var session in expired)
            {
                RemoveSession(session.SessionId);
            }

            // Also check disk for orphaned sessions
            try
            {
                foreach (var dir in Directory.GetDirectories(_sessionsRoot))
                {
                    var dirName = Path.GetFileName(dir);
                    if (_sessions.ContainsKey(dirName)) continue;

                    var origPath = Path.Combine(dir, "original.xlsx");
                    if (File.Exists(origPath) && File.GetCreationTimeUtc(origPath) < cutoff)
                    {
                        try { Directory.Delete(dir, recursive: true); }
                        catch { /* best effort */ }
                    }
                }
            }
            catch { /* best effort */ }

            _logger.LogInformation("[SESSION] Cleanup: removed {Count} expired sessions", expired.Count);
        }
    }
}
