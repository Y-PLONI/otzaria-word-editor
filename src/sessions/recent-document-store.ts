import {
  forgetRecent,
  normalizeRecents,
  rememberRecent,
  setRecentPinned,
  sortedRecents,
  type RecentDocument,
  type RecentOpen,
} from './recent-documents';

type Change = (list: readonly RecentDocument[]) => RecentDocument[];

interface RecentDocumentStoreOptions {
  load(): Promise<unknown>;
  save(list: readonly RecentDocument[]): Promise<void>;
  onChange(list: RecentDocument[]): void;
  onLoading(loading: boolean): void;
  onError(error: unknown): void;
}

/**
 * פתיחת מסמך אינה מחכה להיסטוריה. שינויים מוקדמים מוצגים מיד, אך נכתבים
 * רק אחרי שהרשימה השמורה הגיעה והשינויים הוחלו עליה לפי סדרם. שרשור הכתיבות
 * מונע מתשובת storage.set ישנה לדרוס פתיחה, הצמדה או הסרה מאוחרת יותר.
 */
export function createRecentDocumentStore(options: RecentDocumentStoreOptions) {
  let current: RecentDocument[] = [];
  let loaded = false;
  let loading: Promise<void> | null = null;
  let pending: Change[] = [];
  let writes: Promise<void> = Promise.resolve();

  function persist(): void {
    const snapshot = current;
    writes = writes.then(() => options.save(snapshot)).catch(options.onError);
  }

  function load(): Promise<void> {
    if (loading) return loading;
    options.onLoading(true);
    loading = (async () => {
      try {
        const stored = sortedRecents(normalizeRecents(await Promise.resolve().then(() => options.load())));
        current = pending.reduce((list, change) => change(list), stored);
        const changed = pending.length > 0;
        pending = [];
        loaded = true;
        options.onChange(current);
        if (changed) persist();
      } catch (error) {
        // אין לדרוס היסטוריה שלא הצלחנו לקרוא. ניסיון נוסף יחיל גם את
        // הפעולות הממתינות; בינתיים הרשימה המקומית נשארת מוצגת.
        loading = null;
        options.onError(error);
      } finally {
        options.onLoading(false);
      }
    })();
    return loading;
  }

  function change(update: Change): void {
    current = update(current);
    options.onChange(current);
    if (loaded) persist();
    else {
      pending.push(update);
      void load();
    }
  }

  return {
    load,
    remember(entry: RecentOpen) {
      const snapshot = { ...entry };
      change((list) => rememberRecent(list, snapshot));
    },
    setPinned(token: string, pinned: boolean) {
      change((list) => setRecentPinned(list, token, pinned));
    },
    forget(token: string) {
      change((list) => forgetRecent(list, token));
    },
  };
}
