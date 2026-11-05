import { getDocumentReadRevision, type DocumentReadRevision } from './document-read-state';
import {
  readPageSections,
  type PageSections,
  type PageSetupDocumentApi,
  type PageSetupHost,
  type PageSetupTarget,
  type ReadPageSections,
} from './page-setup';

export interface SectionSnapshot {
  read: ReadPageSections;
  dispose(): void;
}

interface SnapshotEntry {
  doc: PageSetupDocumentApi;
  revision: DocumentReadRevision;
  pending: Promise<PageSections | null>;
}

/**
 * מקור אחד לכל EditorSession: ארבעת קוראי התצוגה חולקים גם בקשה בתהליך.
 * נשמרת רק תשובת sections.list; מדידות הפריסה נקראות אצל הצרכן בכל רענון.
 * הכתיבות אינן משתמשות במקור הזה — הן חייבות לקרוא את המקטעים מחדש.
 */
export function createSectionSnapshot(host: PageSetupTarget): SectionSnapshot {
  let entry: SnapshotEntry | undefined;
  let disposed = false;

  const currentDocument = () => (host as PageSetupHost | null | undefined)?.activeEditor?.doc;

  async function read(): Promise<PageSections | null> {
    if (disposed) return null;
    const doc = currentDocument();
    if (!doc) return null;
    const revision = getDocumentReadRevision(doc);

    if (entry?.doc !== doc || entry.revision !== revision) {
      const next: SnapshotEntry = {
        doc,
        revision,
        // רושמים את הבקשה לפני שנכנסים למנוע, גם אם ה-API סינכרוני.
        pending: Promise.resolve().then(() => {
          if (disposed || currentDocument() !== doc || getDocumentReadRevision(doc) !== revision) return null;
          return readPageSections(host);
        }),
      };
      entry = next;
      // כשל או היעדר מקטעים אינם snapshot שימושי לתצוגה; הרענון הבא רשאי
      // לנסות שוב גם בלי שינוי מסמך (למשל כשהמנוע עדיין משלים את הפתיחה).
      // בקשה ישנה
      // אינה יכולה למחוק את מקומה של בקשה שהתחילה בינתיים לגרסה חדשה.
      void next.pending.then((result) => {
        if (!result?.items?.length && entry === next) entry = undefined;
      });
    }

    const result = await entry.pending;
    // הבדיקה נדרשת גם בפגיעה במטמון שכבר נפתר: בין הקריאה ל-await יכול
    // להתחיל שינוי מסמך, החלפת עורך או פירוק ה-session.
    if (disposed || currentDocument() !== doc || getDocumentReadRevision(doc) !== revision) return null;
    return result;
  }

  return {
    read,
    dispose() {
      disposed = true;
      entry = undefined;
    },
  };
}
