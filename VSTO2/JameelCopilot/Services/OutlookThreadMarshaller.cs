using System;
using System.Threading;
using System.Threading.Tasks;
using JameelCopilot.Diagnostics;

namespace JameelCopilot.Services
{
    /// <summary>
    /// Marshals callbacks onto a specific SynchronizationContext (the Outlook
    /// UI thread in production). Used by LiveOutlookSurface to make every COM
    /// access STA-correct regardless of which thread the chat service runs on.
    ///
    /// Reentrancy check is by <b>managed thread id</b>, not by SyncContext
    /// reference. VSTO/Outlook can swap the SyncContext instance on the UI
    /// thread between startup and first use (observed in a Phase 2 trace),
    /// so a captured reference becomes stale; the thread id stays stable.
    /// </summary>
    public sealed class OutlookThreadMarshaller
    {
        private readonly object _lock = new object();
        private SynchronizationContext _context;
        private readonly int _uiThreadId;

        /// <summary>
        /// Captures the managed thread id of the constructing thread as the
        /// target UI thread. Production callers should construct this on the
        /// Outlook UI thread (typically inside <c>ThisAddIn_Startup</c>).
        /// The SyncContext is captured eagerly here, but <see cref="RunAsync"/>
        /// will re-capture the live <c>SynchronizationContext.Current</c> on
        /// the next call that happens on the UI thread - because VSTO can
        /// swap the SyncContext between startup and first use.
        /// </summary>
        public OutlookThreadMarshaller(SynchronizationContext context)
            : this(context, Thread.CurrentThread.ManagedThreadId)
        {
        }

        /// <summary>
        /// Explicit-thread overload for tests where the marshaller is built
        /// off the target thread (e.g. xUnit constructs it on the test thread
        /// but the target is the test's worker thread).
        /// </summary>
        public OutlookThreadMarshaller(SynchronizationContext context, int uiThreadId)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _uiThreadId = uiThreadId;
        }

        /// <summary>The managed thread id this marshaller treats as "the UI thread".</summary>
        public int UiThreadId => _uiThreadId;

        /// <summary>
        /// Returns the SyncContext currently used for Post. Visible for trace
        /// logging and for tests that want to verify the lazy-recapture path.
        /// </summary>
        public SynchronizationContext CurrentContext
        {
            get { lock (_lock) return _context; }
        }

        // Called on the UI thread, this keeps the thread's SynchronizationContext
        // one that reaches the UI thread. If the live one differs from ours and
        // is a real context, ours is an orphan (installed before VSTO/WinForms
        // installed theirs): adopt the live one. But WinForms also swaps the
        // context under us: when a modal dialog (About Jameel Copilot) or DoEvents ends, it
        // leaves a plain SynchronizationContext, which "posts" to the thread
        // pool. Adopting that one sent every later Outlook and page call off
        // the UI thread, and an await on the UI thread resumed there too. A
        // plain or missing context is never adopted: ours is put back.
        private void RecaptureContextIfStale()
        {
            if (Thread.CurrentThread.ManagedThreadId != _uiThreadId) return;
            var live = SynchronizationContext.Current;
            lock (_lock)
            {
                if (ReferenceEquals(live, _context)) return;
                if (ReachesAThread(live))
                {
                    _context = live;
                }
                else if (ReachesAThread(_context))
                {
                    SynchronizationContext.SetSynchronizationContext(_context);
                }
            }
        }

        // The base class posts to the thread pool; every real UI context derives from it.
        private static bool ReachesAThread(SynchronizationContext context)
        {
            return context != null && context.GetType() != typeof(SynchronizationContext);
        }

        public Task RunAsync(Action action, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RecaptureContextIfStale();
            // Reentrancy guard: when the caller is already on the target UI
            // thread, Post+wait would deadlock. Invoke synchronously instead.
            if (IsOnTargetContext())
            {
                TraceLog.Write("RunAsync reentrant branch (sync invoke)", "Marshaller");
                var tcsSync = new TaskCompletionSource<bool>();
                try { action(); tcsSync.TrySetResult(true); }
                catch (Exception ex) { tcsSync.TrySetException(ex); }
                return tcsSync.Task;
            }
            SynchronizationContext ctx;
            lock (_lock) ctx = _context;
            TraceLog.Write("RunAsync posting (off-thread caller) to " + ctx?.GetType().Name + " hash=" + ctx?.GetHashCode(), "Marshaller");
            var tcs = new TaskCompletionSource<bool>();
            ctx.Post(_ =>
            {
                TraceLog.Write("RunAsync posted action running on T" + Thread.CurrentThread.ManagedThreadId, "Marshaller");
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }
                try { action(); tcs.TrySetResult(true); }
                catch (Exception ex) { tcs.TrySetException(ex); }
            }, null);
            return tcs.Task;
        }

        public Task<T> RunAsync<T>(Func<T> func, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RecaptureContextIfStale();
            // See reentrancy comment in the non-generic overload above.
            if (IsOnTargetContext())
            {
                TraceLog.Write("RunAsync<T> reentrant branch (sync invoke)", "Marshaller");
                var tcsSync = new TaskCompletionSource<T>();
                try { tcsSync.TrySetResult(func()); }
                catch (Exception ex) { tcsSync.TrySetException(ex); }
                return tcsSync.Task;
            }
            SynchronizationContext ctx;
            lock (_lock) ctx = _context;
            TraceLog.Write("RunAsync<T> posting (off-thread caller) to " + ctx?.GetType().Name + " hash=" + ctx?.GetHashCode(), "Marshaller");
            var tcs = new TaskCompletionSource<T>();
            ctx.Post(_ =>
            {
                TraceLog.Write("RunAsync<T> posted action running on T" + Thread.CurrentThread.ManagedThreadId, "Marshaller");
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }
                try { tcs.TrySetResult(func()); }
                catch (Exception ex) { tcs.TrySetException(ex); }
            }, null);
            return tcs.Task;
        }

        private bool IsOnTargetContext()
        {
            // Thread-id comparison instead of SyncContext reference equality:
            // VSTO can swap the SyncContext instance under us between startup
            // and first marshalled call. The thread id is the only stable
            // identity for "the UI thread".
            return Thread.CurrentThread.ManagedThreadId == _uiThreadId;
        }

        /// <summary>
        /// Lets Outlook's UI thread pump its message loop between chunks of a
        /// long marshalled operation. Every <see cref="RunAsync"/> work unit
        /// blocks the UI thread for as long as it runs (COM demands it), so a
        /// coarse unit (a folder scan, a folder-tree walk) freezes the whole
        /// mail client. Awaiting this between small units releases the thread
        /// long enough for Outlook to paint and answer the user; the caller
        /// keeps working on its own thread meanwhile.
        ///
        /// The posted no-op sits behind whatever Outlook has already queued:
        /// by the time it runs, the UI thread has serviced its own messages,
        /// which is exactly the turn we want to give it. Never blocks: the
        /// caller is off the UI thread already, and on it there is nothing to
        /// yield to (and posting + waiting there would deadlock).
        /// </summary>
        public Task YieldAsync(CancellationToken cancellationToken)
        {
            if (cancellationToken.IsCancellationRequested) return Task.FromResult(false);
            RecaptureContextIfStale();
            if (IsOnTargetContext()) return Task.FromResult(false);
            SynchronizationContext ctx;
            lock (_lock) ctx = _context;
            if (ctx == null) return Task.FromResult(false);
            var tcs = new TaskCompletionSource<bool>();
            ctx.Post(_ => tcs.TrySetResult(true), null);
            return tcs.Task;
        }
    }
}
