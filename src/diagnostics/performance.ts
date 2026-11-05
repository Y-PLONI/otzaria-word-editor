/**
 * Optional hooks for scripts/performance-probe.mjs. No recorder, clock read,
 * listeners, serialization or document inspection is installed in normal use.
 * A diagnostic failure must never change opening, editing or saving behavior.
 */
export type PerformanceDetails = Readonly<Record<string, string | number | boolean | null>>;

export interface PerformanceSpan {
  mark(name: string): void;
  bind(editor: object): void;
  end(status: string): void;
}

export interface PerformanceRecorder {
  mark(name: string, details?: PerformanceDetails): void;
  begin(name: string, details?: PerformanceDetails): PerformanceSpan;
  api(scope: string, method: string, request: unknown): (response: unknown, status: 'ok' | 'error') => void;
}

declare global {
  interface Window {
    __otzariaPerformance?: PerformanceRecorder;
  }
}

function recorder(): PerformanceRecorder | undefined {
  return typeof window === 'undefined' ? undefined : window.__otzariaPerformance;
}

export function markPerformance(name: string, details?: PerformanceDetails): void {
  try { recorder()?.mark(name, details); } catch { /* diagnostics only */ }
}

export function beginPerformanceSpan(name: string, details?: PerformanceDetails): PerformanceSpan | undefined {
  try {
    const span = recorder()?.begin(name, details);
    if (!span) return undefined;
    return {
      mark(stage) { try { span.mark(stage); } catch { /* diagnostics only */ } },
      bind(editor) { try { span.bind(editor); } catch { /* diagnostics only */ } },
      end(status) { try { span.end(status); } catch { /* diagnostics only */ } },
    };
  } catch {
    return undefined;
  }
}

export function beginPerformanceApi(
  scope: string,
  method: string,
  request: unknown,
): ((response: unknown, status: 'ok' | 'error') => void) | undefined {
  try {
    const end = recorder()?.api(scope, method, request);
    if (!end) return undefined;
    return (response, status) => { try { end(response, status); } catch { /* diagnostics only */ } };
  } catch {
    return undefined;
  }
}
