export interface KeySample {
  sequence: number; phase: string; documentId: number; downAt: number;
  visibleDomAt: number | null; renderOpportunityAt: number | null;
  coalesced: boolean; saveState: string; validated: boolean;
}
export interface Recording {
  schemaVersion: number; boundary: string; hostClickIncluded: false; inputSource: string;
  clock: string; timeOrigin: number; renderMetric: string; dropped: number; stopped: boolean;
  firstEdit: null | { documentId: number; openingId: number | null; inputAt: number;
    visibleDomAt: number; renderOpportunityAt: number; documentApiValidated: true };
  events: Array<{ name: string; at: number; [key: string]: string | number | boolean | null }>;
  keys: KeySample[]; longTasks: Array<{ start: number; duration: number }>; longTasksSupported: boolean;
  api: Array<{ scope: string; method: string; calls: number; completed: number; failed: number;
    totalMs: number; requestBytes: number | null; responseBytes: number | null; unknownBytes: number }>;
  apiCoverage: Array<{ documentId: number; method: string; status: string }>; apiBytes: string;
}
export interface Recorder {
  readonly ownerId: string | null;
  mark(name: string, details?: Record<string, string | number | boolean | null>): void;
  begin(name: string, details?: Record<string, string | number | boolean | null>): {
    mark(stage: string): void; bind(editor: object): void; end(status: string): void;
  };
  api(scope: string, method: string, request: unknown): (response: unknown, status: 'ok' | 'error') => void;
  bind(editor: object): object | undefined;
  prepare(phase: string, text: string): Promise<{ documentId: number; openingId: number | null; inputCount: number }>;
  finishBurst(): Promise<{ valid: boolean; received: number; expected: number; displayed: number }>;
  snapshot(): Recording;
  stop(): void;
  pending(): null | { received: number; displayed: number };
}
export function installPerformanceRecorder(options?: {
  apiBytes?: boolean; limit?: number; boundary?: 'plugin-navigation-start' | 'late-attach'; ownerId?: string;
}): Recorder;
export function validateStartupOpening(
  recording: Pick<Recording, 'boundary' | 'events'>,
  openingId: number | null,
  expectedKind?: 'file' | 'new' | null,
): boolean;
