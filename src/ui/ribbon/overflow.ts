/**
 * כיווץ קבוצות הרצועה כשאין מקום — ההתנהגות של Word.
 *
 * במקום פס גלילה, הרצועה מתכווצת בהדרגה, שלב אחר שלב, כמו ב-Word:
 *
 *   1. **בינוני** — הכפתורים הגדולים של הקבוצה נעשים קטנים, בטורים של שלושה.
 *   2. **קטן** — הכפתורים הקטנים מאבדים את התווית ונשארים אייקון.
 *   3. **מכווץ** — הקבוצה כולה מתקפלת לכפתור אחד שפותח את תוכנה בפופאובר.
 *
 * כל שלב עובר על הקבוצות מהסוף להתחלה, והשלב הבא מתחיל רק כשהקודם נגמר
 * ועדיין אין מקום — כלומר אף קבוצה אינה מתקפלת כל עוד לאחרת יש עוד כפתור
 * גדול להקטין. „לוח” היא האחרונה שנכנעת.
 *
 * שני השלבים הראשונים הם CSS בלבד (`data-scale` על שורש הקבוצה, ribbon.css),
 * ולכן הבקר מודד את רוחב כל קבוצה בכל שלב במשימה סינכרונית אחת — בלי שאף
 * שלב ביניים יצויר. המכווץ דורש רינדור (הצ'יפ), ולכן הוא נמדד בפעימה שאחרי.
 */
import { inject, provide, type InjectionKey, type Ref } from 'vue';

/** רוחב שמניחים לצ'יפ עד שנמדד אחד אמיתי. */
export const CHIP_ESTIMATE_PX = 76;

/** השלבים שלפני הכיווץ: פרוש (0), בינוני (1), קטן (2). */
export const SCALE_STEPS = 3;

export interface GroupWidth {
  /**
   * רוחב הקבוצה בכל שלב, לפי הסדר: פרושה, בינונית, קטנה. שלב שחסר, או שאינו
   * צר מהשלב הנוכחי, אינו קונה כלום ומדלגים עליו.
   */
  steps: readonly number[];
  /** רוחב הצ'יפ שיחליף אותה. */
  chip: number;
}

export interface GroupPlan {
  /** השלב של הקבוצה כשהיא פרושה — 0 עד `SCALE_STEPS - 1`. */
  scale: number;
  collapsed: boolean;
}

/**
 * השלב של כל קבוצה ברוחב הזה — פונקציה טהורה, ולכן נבדקת בלי DOM.
 *
 * דטרמיניסטית לחלוטין ביחס לרוחב: אין כאן זיכרון של המצב הקודם, ולכן אין
 * מצב שבו כיווץ אחד גורר את הבא ושניהם מהבהבים.
 */
export function planScale(groups: readonly GroupWidth[], available: number): GroupPlan[] {
  const plan = groups.map(() => ({ scale: 0, collapsed: false }));
  if (!groups.length || available <= 0) return plan;

  const width = (i: number): number =>
    plan[i].collapsed ? groups[i].chip : groups[i].steps[plan[i].scale];

  let total = groups.reduce((sum, _, i) => sum + width(i), 0);
  for (let stage = 1; stage <= SCALE_STEPS; stage++) {
    const collapse = stage === SCALE_STEPS;
    for (let i = groups.length - 1; i >= 0; i--) {
      if (total <= available) return plan;
      const next = collapse ? groups[i].chip : groups[i].steps[stage];
      const current = width(i);
      // שלב שאינו צר מהנוכחי (או שלא נמדד) אינו קונה כלום, וכיווץ כזה רק
      // מסתיר פקדים. `!(<)` ולא `>=`, כדי ש-undefined ייפול כאן גם הוא.
      if (!(next < current)) continue;
      total += next - current;
      if (collapse) plan[i].collapsed = true;
      else plan[i].scale = stage;
    }
  }
  return plan;
}

export interface RibbonGroupEntry {
  el: HTMLElement;
  collapsed: Ref<boolean>;
  /** סוגר את הפופאובר של הקבוצה — נקרא כשהיא נפרשת בחזרה. */
  close: () => void;
}

export interface RibbonOverflow {
  register(entry: RibbonGroupEntry): () => void;
  /** מדידה מחדש לפי בקשה — פקד ששינה רוחב. */
  remeasure(): void;
}

const RIBBON_OVERFLOW: InjectionKey<RibbonOverflow> = Symbol('ribbon-overflow');

interface Tracked extends RibbonGroupEntry {
  steps: number[];
  chip: number;
}

/**
 * השלב נכתב ישירות על ה-DOM ולא דרך Vue: זה מה שמאפשר למדוד את כל השלבים
 * באותה משימה. Vue אינו נוגע בתכונה שאינה בתבנית, ולכן רינדור של הקבוצה
 * אינו מוחק אותה.
 */
function setScale(el: HTMLElement, scale: number): void {
  if (scale > 0) el.dataset.scale = String(scale);
  else delete el.dataset.scale;
}

/**
 * מתקין את הבקר על גוף הרצועה. הקבוצות מוצאות אותו ב-`useRibbonOverflow`.
 */
export function provideRibbonOverflow(host: Ref<HTMLElement | null>): RibbonOverflow {
  const entries: Tracked[] = [];
  let scheduled = 0;
  let passes = 0;

  function schedule(): void {
    if (scheduled) return;
    scheduled = requestAnimationFrame(() => {
      scheduled = 0;
      measure();
    });
  }

  function measure(): void {
    const container = host.value;
    // רצועה מכווצת או שטרם נפרסה — `clientWidth` אפס, וכל מדידה כאן הייתה
    // מכווצת את הכול על סמך כלום.
    if (!container || container.clientWidth === 0 || !entries.length) return;

    entries.sort((a, b) =>
      a.el.compareDocumentPosition(b.el) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1,
    );

    // רוחב כל שלב, לכל הקבוצות הפרושות יחד: כתיבה של השלב לכולן ואז קריאה של
    // כולן, פעם לכל שלב. הכול באותה משימה — הדפדפן אינו מצייר באמצע, ולכן אף
    // שלב ביניים אינו נראה. קבוצה מכווצת שומרת את מה שנמדד כשהייתה פרושה.
    const expanded = entries.filter((entry) => !entry.collapsed.value);
    for (let step = 0; step < SCALE_STEPS; step++) {
      for (const entry of expanded) setScale(entry.el, step);
      for (const entry of expanded) {
        const width = entry.el.offsetWidth;
        if (width) entry.steps[step] = width;
      }
    }
    for (const entry of entries) {
      if (!entry.collapsed.value) continue;
      const width = entry.el.offsetWidth;
      if (width) entry.chip = width;
    }

    const style = getComputedStyle(container);
    const available =
      container.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);

    const plan = planScale(entries, available);
    let changed = false;
    plan.forEach(({ scale, collapsed }, index) => {
      const entry = entries[index];
      // המכווצת חוזרת לשלב 0: הפופאובר שלה מציג את הקבוצה המלאה, כמו ב-Word.
      setScale(entry.el, collapsed ? 0 : scale);
      if (entry.collapsed.value === collapsed) return;
      entry.collapsed.value = collapsed;
      if (!collapsed) entry.close();
      changed = true;
    });

    // הצ'יפ הראשון נמדד רק אחרי שהוא צויר, ולכן פעימה נוספת מדייקת את התוכנית.
    // התקרה היא מה שמונע רדיפה בין שתי תוכניות שאינן מתכנסות.
    if (changed && passes < 4) {
      passes++;
      schedule();
    } else {
      passes = 0;
    }
  }

  const observer =
    typeof ResizeObserver === 'function' ? new ResizeObserver(() => schedule()) : null;
  let observed: HTMLElement | null = null;

  function sync(): void {
    if (!observer) return;
    if (observed === host.value) return;
    if (observed) observer.unobserve(observed);
    observed = host.value;
    if (observed) observer.observe(observed);
  }

  const api: RibbonOverflow = {
    register(entry) {
      const natural = entry.el.offsetWidth;
      const tracked: Tracked = {
        ...entry,
        steps: new Array<number>(SCALE_STEPS).fill(natural),
        chip: CHIP_ESTIMATE_PX,
      };
      entries.push(tracked);
      sync();
      passes = 0;
      schedule();
      return () => {
        const index = entries.indexOf(tracked);
        if (index >= 0) entries.splice(index, 1);
        passes = 0;
        schedule();
      };
    },
    remeasure() {
      passes = 0;
      schedule();
    },
  };

  provide(RIBBON_OVERFLOW, api);
  return api;
}

export function useRibbonOverflow(): RibbonOverflow | null {
  return inject(RIBBON_OVERFLOW, null);
}
