/**
 * גרסה משותפת לקריאות התצוגה של מסמך. המפתח הוא ה-Document API עצמו,
 * ולא הטאב הפעיל: תשובה של מסמך ברקע אינה יכולה לעדכן את המטמון של שכנו.
 *
 * הגרסה היא אסימון אטום. קורא שומר את זהותו לפני await ומשווה אחריו;
 * invalidation או פירוק מבטלים כך גם תשובה שכבר יצאה למנוע. הם אינם
 * מבטלים את עבודת ה-API שכבר נשלחה.
 */
export type DocumentReadRevision = object;

interface DocumentReadState {
  revision: DocumentReadRevision;
  listeners: Set<() => void>;
}

const states = new WeakMap<object, DocumentReadState>();

function stateFor(doc: object): DocumentReadState {
  let state = states.get(doc);
  if (!state) {
    state = { revision: {}, listeners: new Set() };
    states.set(doc, state);
  }
  return state;
}

/** להשוואת זהות בלבד; אין על האסימון נתוני מסמך. */
export function getDocumentReadRevision(doc: object): DocumentReadRevision {
  return stateFor(doc).revision;
}

/** המנוי מתזמן רענון אצל הצרכן; ההודעה עצמה אינה מפעילה קריאת מנוע. */
export function observeDocumentReadInvalidation(
  doc: object | null | undefined,
  listener: () => void,
): () => void {
  if (!doc) return () => {};
  const state = stateFor(doc);
  // רישומים נפרדים של אותו callback מתפרקים בנפרד.
  const notify = () => listener();
  state.listeners.add(notify);
  return () => { state.listeners.delete(notify); };
}

/** גם כשל חלקי יכול לשנות מסמך: הקורא למוטציה פוסל ב-finally. */
export function invalidateDocumentReads(doc: object | null | undefined): void {
  if (!doc) return;
  const state = stateFor(doc);
  state.revision = {};
  for (const listener of [...state.listeners]) {
    try { listener(); }
    catch (error) { console.warn('[otzaria-word] רענון חיווי המסמך נכשל', error); }
  }
}

/** פירוק מבטל תוצאות ממתינות ומנתק מנויים, בלי לתזמן עבודה למסמך שנסגר. */
export function forgetDocumentReads(doc: object | null | undefined): void {
  if (!doc) return;
  states.get(doc)?.listeners.clear();
  states.delete(doc);
}
