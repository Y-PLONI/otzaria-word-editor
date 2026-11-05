/**
 * מערכת המאקרו — חיבור חבילת superdoc-macros ל-session של העורך.
 *
 * מה המשתמש מקבל:
 *   - **מקליט מאקרו** בסגנון Word: Ctrl+Alt+R מתחיל/עוצר, Ctrl+Alt+P מנגן את
 *     ההקלטה האחרונה, Alt+F8 פותח את דיאלוג הניהול (הצירוף של Word עצמו).
 *     המקליט מתעד פקודות (דרך ה-controller) והקלדה (דרך beforeinput) — לא
 *     מיקומי סמן, ולכן ניגון פועל מהמקום שבו הסמן עומד, בדיוק כמו ב-Word.
 *   - **קטעי טקסט** עם השלמה אוטומטית: הקלדת מילת הפעלה ואחריה רווח מחליפה
 *     אותה בתוכן הקטע (למשל `בסד` ⟵ `בס"ד`). קטע לדוגמה נשתל בהפעלה
 *     הראשונה כדי שהיכולת תהיה גלויה.
 *   - **מאקרו כתובים** ב-JavaScript שרצים בארגז חול (iframe מבודד) מול API
 *     מצומצם — נכתבים ומורצים מדיאלוג הניהול (MacrosDialog.vue).
 *
 * הקיצורים עצמם **אינם נקשרים כאן**: הם רשומות ברג'יסטרי
 * (ui/shortcuts/registry.ts) ופעולות מעטפת (actions.ts) — מסלול אחד לכל
 * קיצור, כמו כל השאר. מה שכן נקשר כאן: הקיצורים שהמשתמש הצמיד לפריטים
 * שמורים (`kit.attachShortcuts`) — אלה נתונים, לא רשומות סטטיות, ואין להם
 * מקום ברג'יסטרי שנגזר ממנו טיפוס בזמן בנייה.
 *
 * המערכת שייכת ל-session: ההקלטה עוטפת את `ui.commands` של **המופע הפתוח**,
 * וההחלפה עוברת דרך ה-search handle שלו. לכן היא מותקנת בפתיחת מסמך ומפורקת
 * ב-`onDispose` — אותה תבנית כמו אדפטר החיפוש והמודד (ראו App.vue).
 *
 * השמירה ב-localStorage (ברירת המחדל של החבילה): המאקרו של המשתמש שורדים
 * החלפת מסמך והפעלה מחדש, בלי לגעת ב-workspace של הקבצים.
 */
import { shallowRef, type InjectionKey, type Ref } from 'vue';
import type { EditorSession } from './create-editor';
import { SHORTCUTS } from '../ui/shortcuts/registry';
import {
  HEBREW_MESSAGES,
  MacroError,
  MacroKit,
  createMemoryStorage,
  createSuperdocHost,
  setMacroMessages,
  type SuperdocLike,
  type SuperdocMacroHost,
} from 'superdoc-macros';
import type { CommandAdapter } from './command-adapter';

// הודעות הריצה של החבילה בעברית — פעם אחת, בטעינת המודול. ברירת המחדל שלה
// (מ-0.3.0) אנגלית, וכשל מאקרו מגיע לשורת המצב כמו כל כשל אחר בממשק — קול
// אחד, עברית. לפני installMacros ולא בתוכו: ההודעות הן מצב מודול בחבילה,
// לא מצב של session.
setMacroMessages(HEBREW_MESSAGES);

/** ההודעות שהמערכת מדווחת לשורת המצב. מרוכזות כדי שהבדיקות יטענו נוסח אחד. */
export const MACRO_STATUS = {
  recordingStarted: 'מקליט מאקרו… (Ctrl+Alt+R לעצירה)',
  recordingEmpty: 'ההקלטה הופסקה — לא הוקלטו פעולות',
  noRecordings: 'אין מאקרו מוקלט לניגון',
  replayDone: 'המאקרו נוגן',
  /** localStorage חסום או מלא — עובדים מהזיכרון, להפעלה הזאת בלבד. */
  storageUnavailable: 'שמירת מאקרו קבועה אינה זמינה — המאקרו יישמרו להפעלה הזאת בלבד',
  /** שמירת הקלטה נכשלה (למשל quota) — הצעדים שמורים בצד לניסיון נוסף. */
  recordingKept: 'ההקלטה נשמרה בצד — לאחר תיקון הבעיה לחצו שוב Ctrl+Alt+R',
  /** המשתמש סירב לשמור הקלטה לא-שלמה. */
  incompleteDiscarded: 'ההקלטה בוטלה',
  /** startRecording סירב משום שמאקרו אחר עדיין רץ. */
  recordingUnavailable: 'אי אפשר להתחיל הקלטה בזמן שמאקרו אחר רץ',
} as const;

/** כותרת דיאלוג האישור להקלטה לא-שלמה. */
export const INCOMPLETE_CONFIRM_TITLE = 'לשמור הקלטה חלקית?';

export function incompleteConfirmContent(detail: string): string {
  return `${detail}. הניגון לא יכלול את הפעולות האלה. לשמור בכל זאת?`;
}

export function recordingSavedText(name: string): string {
  return `המאקרו "${name}" נשמר (Ctrl+Alt+P לניגון)`;
}

export function recordingAutoStoppedText(name: string): string {
  return `ההקלטה הגיעה לתקרת הצעדים ונשמרה — "${name}"`;
}

export function replayFailedText(message: string): string {
  return `ניגון המאקרו נכשל: ${message}`;
}

/**
 * דגל הסקריפטים. המקליט וקטעי הטקסט פועלים תמיד; מאקרו כתובים — קוד שרץ,
 * גם אם בארגז חול — נשארים מאחורי דגל עד שההקשחה תוכרע סופית (סבב הסקירה
 * הראשון של ה-PR). ההדלקה מפורשת, מ-console:
 *
 *   localStorage.setItem('otzaria-word:macros-scripts', 'on')
 *
 * localStorage ולא ה-storage של אוצריא, בכוונה: הדגל נקרא סינכרונית בהתקנת
 * ה-session, והוא כלי למי שיודע מה הוא עושה — לא הגדרה בממשק.
 */
export const SCRIPTS_FLAG_KEY = 'otzaria-word:macros-scripts';

export function scriptsEnabled(): boolean {
  try {
    return globalThis.localStorage?.getItem(SCRIPTS_FLAG_KEY) === 'on';
  } catch {
    return false;
  }
}

export interface MacrosHandle {
  /** הערכה עצמה — הדיאלוג עובד מולה. כל התוצאות שלה כבר בעברית. */
  kit: MacroKit;
  /** האם לשונית הסקריפטים מוצגת. ראו SCRIPTS_FLAG_KEY. */
  scriptsEnabled: boolean;
  /**
   * האם הקלטה רצה, כ-ref: `kit.isRecording` הוא getter רגיל ש-Vue אינו
   * רואה, וכפתור „הקלט/עצור” צריך להתעדכן בלי שאיש ילחץ עליו שוב.
   */
  recording: Readonly<Ref<boolean>>;
  /** מתחיל או עוצר הקלטה, ומדווח לשורת המצב. */
  toggleRecording(): void;
  /** מנגן את ההקלטה האחרונה, ומדווח לשורת המצב. */
  replayLast(): void;
  /** מפרקת את המאזינים ואת עטיפת ה-controller. אידמפוטנטית. */
  dispose(): void;
}

/**
 * המערכת של המסמך הפתוח, או `null` כשאין מסמך. `shallowRef` בצד המספק —
 * אותו דפוס כמו `ACTIVE_SUPERDOC` ב-document-api.ts, ומאותו טעם.
 */
export const ACTIVE_MACROS: InjectionKey<Ref<MacrosHandle | null>> = Symbol('activeMacros');

/**
 * מה שנצרך מה-session: המופע בלבד. `Pick` ולא הטיפוס המלא — כדי שבדיקה תוכל
 * לכפול רק אותו (אותה תבנית כמו `SearchHost` ב-engine/search.ts).
 */
export type MacrosSession = Pick<EditorSession, 'superdoc'>;

/**
 * קטע לדוגמה, פעם אחת אי-פעם: בלי שום קטע ההשלמה האוטומטית בלתי נראית.
 * הדגל נפרד מרשימת הקטעים בכוונה — משתמש שמחק את כל הקטעים אמר משהו,
 * ושתילה חוזרת בכל פתיחת מסמך הייתה מתווכחת איתו.
 */
export const SEEDED_FLAG_KEY = 'otzaria-word:macros-seeded';

function seedDefaultSnippet(kit: MacroKit, persistMarker = true): void {
  try {
    if (globalThis.localStorage?.getItem(SEEDED_FLAG_KEY) === 'yes') return;
  } catch {
    /* בלי localStorage אין גם persistence לקטעים — שתילה לזיכרון עדיין מועילה. */
  }
  if (kit.listSnippets().length === 0) {
    kit.saveSnippet({ name: 'בס"ד', text: 'בס"ד', trigger: 'בסד' });
  }
  // A memory-only fallback must not mark persistent seeding as complete. The
  // next healthy startup still needs to write the example to durable storage.
  if (!persistMarker) return;
  try {
    globalThis.localStorage?.setItem(SEEDED_FLAG_KEY, 'yes');
  } catch {
    /* הדגל הוא נוחות; כישלון בכתיבתו אינו כשל של השתילה. */
  }
}

export interface InstallMacrosOptions {
  /** Shared dispatcher: even a replay already in flight must respect an opening lock. */
  commands?: Pick<CommandAdapter, 'run'>;
  /**
   * דיאלוג האישור להקלטה לא-שלמה (פעולה שאינה ניתנת להקלטה — למשל הכנסת
   * תמונה, שה-payload שלה הוא הקובץ כולו). בלעדיו — או כשהוא מחזיר false —
   * ההקלטה מבוטלת; שמירה חלקית דורשת הסכמה מפורשת, לא ברירת מחדל.
   */
  confirmIncomplete?: (title: string, content: string) => Promise<boolean>;
  /**
   * הצירופים של הקיצורים האישיים (ui/shortcuts/custom-shortcuts.ts), כמחרוזות
   * שהפרסר של החבילה מבין — `reservationText`.
   *
   * הם נוספים ל-`reservedShortcuts` יחד עם תוויות הרג'יסטרי, ומאותו טעם
   * בדיוק: קשירת מאקרו יושבת על שלב הלכידה של מכל המסמך, כלומר היא
   * **מקדימה** את המנתב שלנו. בלי ההצהרה הזאת אפשר היה להצמיד מאקרו לצירוף
   * שקיצור אישי כבר מחזיק, והתוצאה אינה שגיאה אלא שהקיצור האישי שותק.
   *
   * **מגבלה, ומדודה:** `reservedShortcuts` נקרא פעם אחת בבניית ה-kit, ואין
   * לחבילה API לרשימה חיה. לכן מוגנים הקיצורים שהיו מוגדרים כשהמסמך נפתח.
   * קיצור אישי שנוסף **אחרי** פתיחת המסמך אינו מוגן עד הפתיחה הבאה; הכיוון
   * ההפוך — דיאלוג „ניהול קיצורים” — כן קורא את רשימת המאקרו בזמן אמת.
   */
  reservedCombos?: readonly string[];
}

export function installMacros(
  editor: MacrosSession,
  container: HTMLElement,
  onStatus: (message: string, isError?: boolean) => void,
  options: InstallMacrosOptions = {},
): MacrosHandle {
  /* ההצרה ל-SuperdocLike מבנית: החבילה מגדירה בעצמה את תת-הצורה שהיא צורכת
     (commands / doc / search / view), במקום לייבא את הטיפוסים הפנימיים של
     superdoc — אותו טעם כמו ב-document-api.ts. `unknown` באמצע כי הטיפוס
     המלא של SuperDoc מדויק יותר בשדות שהחבילה מכריזה כאופציונליים. */
  const host: SuperdocMacroHost = createSuperdocHost({
    superdoc: editor.superdoc as unknown as SuperdocLike,
    container,
  });
  const commands = options.commands;
  if (commands) host.commands.execute = (id, payload) => commands.run(id, payload);

  const allowScripts = scriptsEnabled();

  const buildKit = (storage?: ConstructorParameters<typeof MacroKit>[0]['storage']): MacroKit =>
    new MacroKit({
      host,
      ...(storage ? { storage } : {}),
      onLog: (line) => onStatus(line),
      // כל קיצורי הרג'יסטרי חסומים לקיצורים אישיים: קיצור מאקרו נקשר בשלב
      // הלכידה, ולכן התנגשות הייתה מאפילה בשקט על קיצור של העורך. התוויות
      // נמסרות כמו שהן — מה שאינו ניתן לפירוק ("Ctrl + Shift ימני") פשוט
      // אינו ניתן גם להתנגשות, והחבילה מתעלמת ממנו.
      reservedShortcuts: [
        ...SHORTCUTS.map((shortcut) => shortcut.label),
        ...(options.reservedCombos ?? []),
      ],
      // ה-gate האמיתי יושב ב-kit ולא בדיאלוג: כשהדגל כבוי runScript/runSource
      // מסרבים וקיצורי סקריפטים אינם נקשרים — סקריפט קיים או מיובא לא ירוץ
      // בשום מסלול. הסתרת הלשונית היא רק הצד הקוסמטי של אותו מתג.
      scriptsEnabled: allowScripts,
      // הקלטה שנעצרה בתקרת הצעדים: ה-callback הוא מה שמעדכן את מחוון
      // ה„מקליט” ושומר את מה שהוקלט — בלעדיו הכפתור היה ממשיך להבטיח הקלטה
      // שכבר אינה קורית, והלחיצה הבאה הייתה זורקת את הצעדים.
      onRecordingAutoStop: () => void finishRecording(true),
    });

  /* localStorage חסום, חסר או מלא אסור לו להפיל את פתיחת המסמך: המאקרו הם
     פיצ'ר, לא תנאי. נפילה לאחסון-זיכרון משאירה את כל היכולות עובדות להפעלה
     הנוכחית, וההודעה אומרת בדיוק מה אבד — הקביעות. */
  let kit: MacroKit;
  let loadedKit: MacroKit | null = null;
  try {
    loadedKit = buildKit();
    seedDefaultSnippet(loadedKit);
    kit = loadedKit;
  } catch {
    // The store may be readable but become unwritable on the first seed write.
    // Preserve everything already loaded when falling back to session memory.
    const loadedState = loadedKit?.exportState() ?? null;
    kit = buildKit(createMemoryStorage());
    if (loadedState) {
      const restored = kit.importState(loadedState);
      if (!restored.ok) {
        console.warn('[otzaria-word] שחזור המאקרו לאחסון זמני נכשל:', restored.message);
      }
    }
    seedDefaultSnippet(kit, false);
    onStatus(MACRO_STATUS.storageUnavailable, true);
  }

  const recording = shallowRef(false);
  let disposed = false;
  let finishRecordingInFlight = false;
  // The generated name and consent belong to the pending snapshot. Retry the
  // same logical save without renaming it or asking twice after a recoverable
  // capacity/storage failure.
  let pendingIncompleteApproved = false;
  let pendingRecordingName: string | null = null;

  function clearPendingFinalizationState(): void {
    pendingIncompleteApproved = false;
    pendingRecordingName = null;
  }

  function reportRetainedRecording(error: unknown): void {
    const detail = error instanceof Error ? error.message : 'שמירת ההקלטה נכשלה';
    onStatus(`${detail}. ${MACRO_STATUS.recordingKept}`, true);
  }

  /**
   * עצירה ושמירה — מסלול אחד ללחיצת המשתמש, לעצירה האוטומטית בתקרה ולניסיון
   * חוזר אחרי כשל שמירה (ההקלטה שנעצרה נשמרת בחבילה בצד עד discard).
   */
  async function finishRecording(auto: boolean): Promise<void> {
    if (disposed || finishRecordingInFlight) return;
    finishRecordingInFlight = true;
    recording.value = false;
    const name = pendingRecordingName ?? `מאקרו ${kit.listRecordings().length + 1}`;
    pendingRecordingName = name;
    try {
      try {
        const saved = kit.stopRecording(name, undefined, {
          allowIncomplete: pendingIncompleteApproved,
        });
        clearPendingFinalizationState();
        if (saved) onStatus(auto ? recordingAutoStoppedText(saved.name) : recordingSavedText(saved.name));
        else onStatus(MACRO_STATUS.recordingEmpty, auto);
        return;
      } catch (error) {
        if (error instanceof MacroError && error.reason === 'recording-uncapturable') {
          // There is no replayable step. A partial save would be an empty
          // macro, so report the exact problem and discard explicitly.
          clearPendingFinalizationState();
          kit.cancelRecording();
          onStatus(error.message, true);
          return;
        }
        if (error instanceof MacroError && error.reason === 'recording-incomplete') {
          let consent = false;
          try {
            consent =
              (await options.confirmIncomplete?.(
                INCOMPLETE_CONFIRM_TITLE,
                incompleteConfirmContent(error.message),
              )) === true;
          } catch (confirmError) {
            // A host-dialog failure must not destroy the pending recording or
            // surface as an unhandled rejection. A later retry may ask again.
            reportRetainedRecording(confirmError);
            return;
          }
          if (disposed) return;
          if (consent) {
            pendingIncompleteApproved = true;
            try {
              const saved = kit.stopRecording(name, undefined, { allowIncomplete: true });
              clearPendingFinalizationState();
              if (saved) onStatus(recordingSavedText(saved.name));
              return;
            } catch (retryError) {
              if (
                retryError instanceof MacroError &&
                (retryError.reason === 'recording-too-large' ||
                  retryError.reason === 'recording-uncapturable')
              ) {
                clearPendingFinalizationState();
                kit.cancelRecording();
                onStatus(retryError.message, true);
                return;
              }
              // Capacity and persistence failures are retryable. The package
              // retains the stopped capture; cancelling here would defeat it.
              // Consent and name stay attached to this pending snapshot.
              reportRetainedRecording(retryError);
              return;
            }
          }
          clearPendingFinalizationState();
          kit.cancelRecording();
          onStatus(MACRO_STATUS.incompleteDiscarded);
          return;
        }
        if (error instanceof MacroError && error.reason === 'recording-too-large') {
          clearPendingFinalizationState();
          kit.cancelRecording();
          onStatus(error.message, true);
          return;
        }
        reportRetainedRecording(error);
      }
    } finally {
      finishRecordingInFlight = false;
    }
  }

  const unbindShortcuts = kit.attachShortcuts(container);
  const disableAutoText = kit.enableAutoText();
  return {
    kit,
    scriptsEnabled: allowScripts,
    recording,

    toggleRecording() {
      if (disposed || finishRecordingInFlight) return;
      if (kit.isRecording) {
        void finishRecording(false);
      } else if (kit.hasPendingRecording) {
        // הקלטה שנעצרה ושמירתה נכשלה ממתינה בצד — הלחיצה מנסה שוב במקום
        // להתחיל הקלטה חדשה שהייתה זורקת אותה.
        void finishRecording(false);
      } else {
        clearPendingFinalizationState();
        const started = kit.startRecording();
        recording.value = started;
        if (started) onStatus(MACRO_STATUS.recordingStarted);
        else onStatus(MACRO_STATUS.recordingUnavailable, true);
      }
    },

    replayLast() {
      const recordings = kit.listRecordings();
      const last = recordings[recordings.length - 1];
      if (!last) {
        onStatus(MACRO_STATUS.noRecordings, true);
        return;
      }
      void kit.replayRecording(last.id).then((result) => {
        if (result.ok) onStatus(MACRO_STATUS.replayDone);
        else onStatus(replayFailedText(result.failures[0]?.message ?? 'כשל לא ידוע'), true);
      });
    },

    dispose() {
      if (disposed) return;
      disposed = true;
      clearPendingFinalizationState();
      // הקלטה שנשארה פתוחה בהחלפת מסמך נזרקת: הצעדים הבאים היו מוקלטים
      // מהמסמך הלא נכון.
      kit.cancelRecording();
      recording.value = false;
      disableAutoText();
      unbindShortcuts();
      host.dispose();
    },
  };
}
