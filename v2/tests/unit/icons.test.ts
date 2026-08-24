/**
 * הבדיקות כאן נגזרות מביקורת ויזואלית שמצאה בספריית האייקונים ארבעה סוגי
 * תקלות שכולן „נכשלות בשקט” — הן לא מפילות שום דבר, הן פשוט נראות רע:
 *
 * 1. גריד לא אחיד. `word` ו-`paste` היו על viewBox 24, `launcher` על 16 וכל
 *    השאר על 20. ה-SvgIcon קובע גודל בפיקסלים, ולכן גריד שונה = עובי קווים
 *    שונה באותה שורה בסרגל.
 * 2. paths שחורגים מה-viewBox ולכן נחתכים. `dirRtl`/`dirLtr` חרגו 1.1 יחידות
 *    מעל הגבול העליון ו-`cut` 1.3 מתחת לתחתון. חריגה כזאת אינה נראית בקוד
 *    ואינה מפילה כלום — היא רק חותכת חלק מהצורה במסך.
 * 3. משקל אופטי פרוע: `reject` תפס 47%x47% מה-viewBox ליד `footnote` שתפס
 *    90%x75%.
 * 4. שם אייקון שלא קיים ב-ICONS. `SvgIcon` מחזיר `ICONS[name] || ''`, כלומר
 *    כפתור עם שם שגוי מוצג בלי אייקון ובלי שגיאה.
 *
 * לכן הבדיקות רצות על הנתונים האמיתיים: ICONS עצמו, וסריקת קובצי ה-Vue
 * שבפועל קוראים לו — ולא על כפילים.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join, relative } from 'node:path';
import { ICONS, PLANNED_ICONS } from '../../src/ui/icons/icons';

const VIEWBOX = '0 0 20 20';
const [VB_X, VB_Y, VB_W, VB_H] = VIEWBOX.split(' ').map(Number) as [
  number,
  number,
  number,
  number,
];

/** סובלנות לשגיאת דגימה של קשתות ולעיגול המספרים ב-path. */
const EPS = 0.05;

interface Point {
  x: number;
  y: number;
}

/**
 * המרת קשת מ-endpoint parameterization ל-center parameterization לפי נספח
 * F.6 של מפרט SVG, ואז דגימה של נקודות על הקשת.
 *
 * הדגימה אינה קוסמטיקה: הקצוות של קשת נמצאים בדרך כלל *בין* נקודות הקצה
 * שלה, ובדיוק שם היו החריגות של `dirRtl`/`dirLtr` — טופס חצי-עיגול שיצא מעל
 * הגבול העליון בזמן ששתי נקודות הקצה שלו בתוך ה-viewBox.
 */
function arcPoints(
  x0: number,
  y0: number,
  rxIn: number,
  ryIn: number,
  phiDeg: number,
  largeArc: number,
  sweep: number,
  x: number,
  y: number,
  steps = 48
): Point[] {
  if (rxIn === 0 || ryIn === 0) return [{ x, y }];
  const phi = (phiDeg * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx2 = (x0 - x) / 2;
  const dy2 = (y0 - y) / 2;
  const x1p = cos * dx2 + sin * dy2;
  const y1p = -sin * dx2 + cos * dy2;
  let rx = Math.abs(rxIn);
  let ry = Math.abs(ryIn);

  // רדיוס קטן מהמיתר: המפרט מחייב להגדיל אותו עד שהקשת אפשרית.
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    const s = Math.sqrt(lambda);
    rx *= s;
    ry *= s;
  }

  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let coef = den === 0 ? 0 : Math.sqrt(Math.max(0, num / den));
  if (largeArc === sweep) coef = -coef;
  const cxp = (coef * rx * y1p) / ry;
  const cyp = (-coef * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x0 + x) / 2;
  const cy = sin * cxp + cos * cyp + (y0 + y) / 2;

  const angle = (ux: number, uy: number, vx: number, vy: number): number => {
    const dot = (ux * vx + uy * vy) / (Math.hypot(ux, uy) * Math.hypot(vx, vy));
    const a = Math.acos(Math.min(1, Math.max(-1, dot)));
    return ux * vy - uy * vx < 0 ? -a : a;
  };

  const ux = (x1p - cxp) / rx;
  const uy = (y1p - cyp) / ry;
  const vx = (-x1p - cxp) / rx;
  const vy = (-y1p - cyp) / ry;
  const theta1 = angle(1, 0, ux, uy);
  let delta = angle(ux, uy, vx, vy);
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;

  const out: Point[] = [];
  for (let i = 0; i <= steps; i++) {
    const t = theta1 + (delta * i) / steps;
    out.push({
      x: cos * rx * Math.cos(t) - sin * ry * Math.sin(t) + cx,
      y: sin * rx * Math.cos(t) + cos * ry * Math.sin(t) + cy,
    });
  }
  return out;
}

const ARITY: Record<string, number> = {
  M: 2,
  L: 2,
  H: 1,
  V: 1,
  C: 6,
  S: 4,
  Q: 4,
  T: 2,
  A: 7,
  Z: 0,
};

/**
 * תת-נתיב אחד — צורה אחת בתוך ה-`d`: הטבעת החיצונית של מסגרת, החור שבתוכה,
 * שורת טקסט. ההפרדה נדרשת כי „האייקון הוא משולש” היא טענה על **צורה** ולא על
 * ה-path כולו, ומשולש האזהרה שהיה כאן היה תת-הנתיב הראשון מתוך ארבעה.
 */
interface Subpath {
  pts: Point[];
  /** רק קווים ישרים — כלומר ה-`pts` הן הקודקודים עצמם ולא דגימה או נקודות בקרה. */
  straight: boolean;
}

/**
 * נקודות המעטפת של path, מקובצות לתת-נתיבים. נקודות הבקרה של בזייה נכללות
 * בכוונה: הן חוסמות את העקומה מלמעלה, ולכן bounding box שמבוסס עליהן הוא שמרני
 * — הוא עלול לדווח חריגה שאין, אך לא יפספס חריגה שיש.
 *
 * הפירוק לתת-נתיבים נעשה כאן ולא בחיתוך המחרוזת על `M`: `m` היא moveto **יחסי**
 * לסוף התת-נתיב הקודם (כך נכתב משולש האזהרה: `...z m0 3.8...`), ולכן חיתוך
 * טקסטואלי היה מחשב את הקודקודים של כל צורה שנייה והלאה במקום הלא נכון.
 */
function walkPath(d: string): Subpath[] {
  const tokens = d.match(/[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d*\.\d+|\d+)/g) ?? [];
  const subs: Subpath[] = [];
  let current: Subpath = { pts: [], straight: true };
  const push = (point: Point): void => {
    current.pts.push(point);
  };
  const startSubpath = (): void => {
    if (current.pts.length) subs.push(current);
    current = { pts: [], straight: true };
  };
  let cmd = '';
  let cx = 0;
  let cy = 0;
  let startX = 0;
  let startY = 0;
  let i = 0;

  while (i < tokens.length) {
    const token = tokens[i]!;
    if (/[A-Za-z]/.test(token)) {
      cmd = token;
      i += 1;
      if (cmd === 'Z' || cmd === 'z') {
        cx = startX;
        cy = startY;
        continue;
      }
    }
    if (!cmd) break;

    const upper = cmd.toUpperCase();
    const rel = cmd !== upper;
    const need = ARITY[upper];
    if (need === undefined) throw new Error(`פקודת path לא מוכרת: ${cmd}`);
    if (i + need > tokens.length) break;
    const a = tokens.slice(i, i + need).map(Number);
    i += need;

    const ax = (v: number): number => (rel ? cx + v : v);
    const ay = (v: number): number => (rel ? cy + v : v);

    switch (upper) {
      case 'M':
        cx = ax(a[0]!);
        cy = ay(a[1]!);
        startX = cx;
        startY = cy;
        startSubpath();
        push({ x: cx, y: cy });
        // אחרי M נוספים, זוגות נוספים הם L (מפרט SVG §9.3.3).
        cmd = rel ? 'l' : 'L';
        break;
      case 'L':
      case 'T':
        cx = ax(a[0]!);
        cy = ay(a[1]!);
        // T היא בזייה, ולכן היא מבטלת את „ישר” גם אם הנקודה שנרשמה היא קצה.
        if (upper === 'T') current.straight = false;
        push({ x: cx, y: cy });
        break;
      case 'H':
        cx = ax(a[0]!);
        push({ x: cx, y: cy });
        break;
      case 'V':
        cy = ay(a[0]!);
        push({ x: cx, y: cy });
        break;
      case 'C':
      case 'S':
      case 'Q': {
        const pairs = upper === 'C' ? [0, 2, 4] : [0, 2];
        current.straight = false;
        for (const p of pairs) push({ x: ax(a[p]!), y: ay(a[p + 1]!) });
        cx = ax(a[pairs[pairs.length - 1]!]!);
        cy = ay(a[pairs[pairs.length - 1]! + 1]!);
        break;
      }
      case 'A': {
        const ex = ax(a[5]!);
        const ey = ay(a[6]!);
        current.straight = false;
        for (const point of arcPoints(cx, cy, a[0]!, a[1]!, a[2]!, a[3]!, a[4]!, ex, ey)) push(point);
        cx = ex;
        cy = ey;
        break;
      }
    }
  }
  if (current.pts.length) subs.push(current);
  return subs;
}

/** כל נקודות המעטפת, בלי חלוקה לצורות — לבדיקות ה-viewBox ויחס המילוי. */
function pathPoints(d: string): Point[] {
  return walkPath(d).flatMap((sub) => sub.pts);
}

interface Measured {
  viewBox: string | null;
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
}

function measure(svg: string): Measured {
  const vb = /viewBox="([^"]+)"/.exec(svg);
  const pts = [...svg.matchAll(/\sd="([^"]+)"/g)].flatMap((m) => pathPoints(m[1]!));
  if (!pts.length) throw new Error('לא נמצא אף path באייקון');
  return {
    viewBox: vb ? vb[1]! : null,
    minX: Math.min(...pts.map((p) => p.x)),
    minY: Math.min(...pts.map((p) => p.y)),
    maxX: Math.max(...pts.map((p) => p.x)),
    maxY: Math.max(...pts.map((p) => p.y)),
  };
}

const NAMES = Object.keys(ICONS);
const MEASURED = new Map(NAMES.map((name) => [name, measure(ICONS[name]!)]));

/**
 * טווח יחס המילוי: המידה הדומיננטית של ה-bounding box חייבת לתפוס 70%–85%
 * מה-viewBox, והמידה הקטנה לפחות 50%. המספרים אינם שרירותיים — הם הטווח
 * שבו הסט נמצא בפועל אחרי הנרמול, ולכן אייקון חדש שנופל מחוץ להם הוא אייקון
 * שייראה גדול או קטן מהשכנים שלו באותה שורה.
 *
 * התקרה הייתה 84% כל עוד כל הסט צויר בבית. היא עלתה ל-85% כשלשונית "קובץ"
 * עברה ל-Fluent System Icons, כי הגריד של Fluent מכוון אחרת לפי צורת הגליף:
 * `info` (עיגול) יושב 2–18 = 80%, אבל `document_add` (דף עם תג) יושב 2–19
 * לגובה = 85%, ו-`folder_open` מגיע ל-85% רוחב רק בגלל נקודות בקרה של בזייה
 * שהמעטפת סופרת — הדיו עצמו נעצר ב-x=18. שתי חריגות של יחידה אחת, ולא
 * מקום להתרחבות: מעל 85% אין אף אייקון בסט.
 */
const DOMINANT_MIN = 0.7;
const DOMINANT_MAX = 0.85;
const MINOR_MIN = 0.5;

/**
 * אייקונים שטוחים לגיטימית: הצורה שלהם *היא* קו או חץ רחב, ולכן המידה הקטנה
 * שלהם נמוכה בכוונה ואין מה לאכוף עליה. הם עדיין נבדקים על המידה הדומיננטית,
 * בטווח רחב יותר — כדי שאייקון שהתכווץ לכתם לא יעבור בשקט.
 *
 * הרשימה סגורה במכוון: כל תוספת אליה היא החלטה עיצובית שצריכה להיות מוסברת
 * כאן, ולא דרך לעקוף כשל בבדיקה.
 */
const FLAT_ICONS = new Set([
  'chevronDown',
  'chevronUp',
  'chevronLeft',
  'chevronRight',
  'undo',
  'redo',
  'strikethrough',
  'fitWidth',
]);
const FLAT_DOMINANT_MIN = 0.55;

describe('גריד האייקונים', () => {
  it('הספרייה אינה ריקה, וכל שם מפנה ל-SVG', () => {
    expect(NAMES.length).toBeGreaterThan(50);
    for (const name of NAMES) {
      expect(ICONS[name], name).toMatch(/^<svg[\s\S]*<\/svg>$/);
    }
  });

  it(`לכל אייקון viewBox="${VIEWBOX}" בדיוק`, () => {
    const wrong = NAMES.filter((n) => MEASURED.get(n)!.viewBox !== VIEWBOX).map(
      (n) => `${n}: ${MEASURED.get(n)!.viewBox}`
    );
    expect(wrong).toEqual([]);
  });

  it('כל האייקונים משתמשים ב-currentColor ולא בצבע קבוע', () => {
    const hardcoded = NAMES.filter((n) => /(?:fill|stroke)="(?!currentColor|none)/.test(ICONS[n]!));
    expect(hardcoded).toEqual([]);
  });
});

describe('גבולות ה-viewBox', () => {
  it('אף נקודה על אף path אינה חורגת מה-viewBox', () => {
    const over: string[] = [];
    for (const name of NAMES) {
      const m = MEASURED.get(name)!;
      const parts: string[] = [];
      if (m.minX < VB_X - EPS) parts.push(`שמאל ${(m.minX - VB_X).toFixed(2)}`);
      if (m.minY < VB_Y - EPS) parts.push(`למעלה ${(m.minY - VB_Y).toFixed(2)}`);
      if (m.maxX > VB_X + VB_W + EPS) parts.push(`ימין +${(m.maxX - VB_X - VB_W).toFixed(2)}`);
      if (m.maxY > VB_Y + VB_H + EPS) parts.push(`למטה +${(m.maxY - VB_Y - VB_H).toFixed(2)}`);
      if (parts.length) over.push(`${name}: ${parts.join(', ')}`);
    }
    expect(over).toEqual([]);
  });

  it('הדגימה של קשתות באמת מודדת את טופס הקשת ולא רק את נקודות הקצה', () => {
    // חצי-עיגול מ-(10,10) ל-(10,2) שטופסו יוצא שמאלה עד x=6: שתי נקודות
    // הקצה בתוך ה-viewBox, והטופס הוא מה שקובע. זו התבנית שהחביאה את
    // החריגה ב-dirRtl/dirLtr.
    const pts = pathPoints('M10 10a4 4 0 0 1 0-8z');
    expect(Math.min(...pts.map((p) => p.x))).toBeCloseTo(6, 2);
    expect(Math.min(...pts.map((p) => p.y))).toBeCloseTo(2, 2);
  });
});

/**
 * הקודקודים של תת-נתיב שבנוי רק מקווים ישרים, בלי כפילויות — למשל נקודת
 * הסגירה שחוזרת על נקודת הפתיחה. תת-נתיב עם עקומה מחזיר `null`: שם הנקודות
 * שנרשמו הן דגימה ונקודות בקרה, ולא קודקודים של פוליגון.
 */
function polygonVertices(sub: Subpath): Point[] | null {
  if (!sub.straight) return null;
  const seen = new Set<string>();
  const out: Point[] = [];
  for (const point of sub.pts) {
    const key = `${point.x.toFixed(2)},${point.y.toFixed(2)}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(point);
  }
  return out;
}

/**
 * מתחת לזה משולש הוא פרט בתוך אייקון (ראש החץ של `indentIncrease`, המשולש
 * של `replace`, הדגל של `dirRtl`); מעל זה הוא **הצורה של האייקון**. הסט בפועל
 * נמצא הרחק משני צדי הגבול: הפרטים תופסים עד 35% מהמידה, ומשולש האזהרה שהיה
 * כאן תפס 78%x75%.
 */
const WARNING_TRIANGLE_MIN = 0.5;

describe('סמנטיקה של הצורה', () => {
  it('אין אייקון שהוא משולש שממלא את ה-viewBox — זה סימן האזהרה המוסכם', () => {
    // מה שקרה: `otzaria` היה משולש חלול עם מקף אנכי ונקודה מתחתיו — כלומר
    // משולש + סימן קריאה — על כפתור „פתח ספרייה”. שום בדיקה לא התלוננה, כי
    // הגאומטריה הייתה תקינה לחלוטין: viewBox נכון, בתוך הגבולות, יחס מילוי
    // בטווח. „נראה כמו ספר” אינו ניתן לבדיקה, אבל „הצורה היא סימן אזהרה” כן:
    // משולש שהוא הצורה הראשית של אייקון פירושו אזהרה או שגיאה בכל סט אייקונים
    // מוכר, ואין לנו אף פקד שזה תפקידו.
    const warnings: string[] = [];
    for (const name of NAMES) {
      for (const match of ICONS[name]!.matchAll(/\sd="([^"]+)"/g)) {
        for (const sub of walkPath(match[1]!)) {
          const vertices = polygonVertices(sub);
          if (!vertices || vertices.length !== 3) continue;
          const w = (Math.max(...vertices.map((v) => v.x)) - Math.min(...vertices.map((v) => v.x))) / VB_W;
          const h = (Math.max(...vertices.map((v) => v.y)) - Math.min(...vertices.map((v) => v.y))) / VB_H;
          if (w >= WARNING_TRIANGLE_MIN && h >= WARNING_TRIANGLE_MIN) {
            warnings.push(`${name}: משולש ${(w * 100).toFixed(0)}%x${(h * 100).toFixed(0)}%`);
          }
        }
      }
    }
    expect(warnings).toEqual([]);
  });

  it('הגלאי אכן מזהה את משולש האזהרה שהיה כאן', () => {
    // בלי הבדיקה הזאת שער כמו זה שלמעלה יכול לעבור בירוק מפני שהוא אינו מודד
    // כלום — למשל אחרי שינוי בפירוק תת-הנתיבים.
    const [triangle] = walkPath('M10 2l-8 15h16L10 2zm0 3.8l5.5 10.2H4.5L10 5.8z');
    const vertices = polygonVertices(triangle!);
    expect(vertices).toHaveLength(3);
    const w = Math.max(...vertices!.map((v) => v.x)) - Math.min(...vertices!.map((v) => v.x));
    expect(w / VB_W).toBeGreaterThanOrEqual(WARNING_TRIANGLE_MIN);
  });
});

describe('משקל אופטי', () => {
  it('יחס המילוי של כל אייקון בטווח המוגדר', () => {
    const bad: string[] = [];
    for (const name of NAMES) {
      const m = MEASURED.get(name)!;
      const w = (m.maxX - m.minX) / VB_W;
      const h = (m.maxY - m.minY) / VB_H;
      const dominant = Math.max(w, h);
      const minor = Math.min(w, h);
      const flat = FLAT_ICONS.has(name);
      const min = flat ? FLAT_DOMINANT_MIN : DOMINANT_MIN;
      const detail = `${name}: ${(w * 100).toFixed(0)}%x${(h * 100).toFixed(0)}%`;
      if (dominant < min - 0.001 || dominant > DOMINANT_MAX + 0.001) {
        bad.push(`${detail} — מידה דומיננטית ${(dominant * 100).toFixed(0)}% מחוץ לטווח`);
      } else if (!flat && minor < MINOR_MIN - 0.001) {
        bad.push(`${detail} — מידה קטנה ${(minor * 100).toFixed(0)}% מתחת למינימום`);
      }
    }
    expect(bad).toEqual([]);
  });

  it('רשימת האייקונים השטוחים אינה מכילה שמות שאינם קיימים', () => {
    expect([...FLAT_ICONS].filter((n) => !(n in ICONS))).toEqual([]);
  });
});

// vitest רץ עם root=v2, ולכן cwd הוא שורש הפרויקט (כמו ב-engine-boundaries).
const SRC = join(process.cwd(), 'src');

function vueFiles(dir = SRC): string[] {
  return sourceFiles(dir).filter((file) => file.endsWith('.vue'));
}

/** כל קובצי המקור — גם .ts, כי שם נבחרים אייקונים בזמן ריצה. */
function sourceFiles(dir = SRC): string[] {
  const files: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...sourceFiles(full));
    else if (/\.(vue|ts)$/.test(entry.name)) files.push(full);
  }
  return files;
}

/**
 * שימושים בפועל של אייקונים בתבניות. שתי החרגות מכוונות:
 *
 * - הערות HTML מוסרות. ב-RibbonGroup.vue יש כפתור launcher בהערה (הוא הוסר
 *   מהעיצוב), ואייקון שרק הערה מזכירה אינו „בשימוש”.
 * - מאפיין שקדם לו `:` הוא binding דינמי (`:name="icon"`), כלומר השם נקבע
 *   בזמן ריצה מ-prop של הרכיב — ואת הערך האמיתי נותן אתר הקריאה, שגם הוא
 *   נסרק כאן.
 *
 * מה שסריקת המאפיינים הזאת **אינה** רואה: שם שנבחר בביטוי, כמו
 * `:name="isCollapsed ? 'chevronDown' : 'chevronUp'"` ב-Ribbon.vue, או שם
 * שמוחזר מפונקציה ב-.ts (כך נבחרים chevronLeft/chevronRight בגלריית
 * הסגנונות). בגללה `chevronUp` נראה בטעות כאייקון ללא צרכן ונכנס
 * ל-PLANNED_ICONS. ראו dynamicIconNames.
 */
function usedIcons(): Map<string, string[]> {
  const used = new Map<string, string[]>();
  for (const file of vueFiles()) {
    const text = readFileSync(file, 'utf8').replace(/<!--[\s\S]*?-->/g, '');
    for (const m of text.matchAll(/(?<!:)\b(?:icon|name)="([A-Za-z][A-Za-z0-9]*)"/g)) {
      const name = m[1]!;
      const where = used.get(name) ?? [];
      where.push(relative(process.cwd(), file));
      used.set(name, where);
    }
  }
  return used;
}

/**
 * שמות אייקונים שנבחרים בזמן ריצה: מחרוזת מצוטטת בכל קובץ מקור, שמצטלבת עם
 * שמות ה-ICONS. פחות מדויק מסריקת מאפיינים — מחרוזת שמקרית לה אותו שם תיחשב
 * שימוש — ולכן היא משמשת רק כדי **להרחיב** את קבוצת „בשימוש”, ולא כדי לאמת
 * שהשם קיים. עדיף מהחלופה: אייקון שכן מחובר לפקד ונראה כמיותם, ואז נדחף
 * ל-PLANNED_ICONS וההחרגה גדלה בשקט.
 */
function dynamicIconNames(): Set<string> {
  const known = new Set(NAMES);
  const found = new Set<string>();
  // icons.ts עצמו מוחרג: PLANNED_ICONS הוא מחרוזות מצוטטות של שמות אייקונים,
  // וסריקה שלו הייתה מדווחת שכל אייקון מתוכנן „בשימוש”.
  const definition = join(SRC, 'ui/icons/icons.ts');
  for (const file of sourceFiles()) {
    if (file === definition) continue;
    const text = readFileSync(file, 'utf8').replace(/<!--[\s\S]*?-->/g, '');
    for (const m of text.matchAll(/['"`]([A-Za-z][A-Za-z0-9]*)['"`]/g)) {
      if (known.has(m[1]!)) found.add(m[1]!);
    }
  }
  return found;
}

describe('אתרי הקריאה לאייקונים', () => {
  const used = usedIcons();
  const referenced = new Set([...used.keys(), ...dynamicIconNames()]);

  it('הסריקה אכן מצאה שימושים (הגנה מפני regex שהפסיק להתאים)', () => {
    expect(used.size).toBeGreaterThan(40);
    expect(used.has('bold')).toBe(true);
  });

  it('כל שם אייקון שמופיע בתבנית Vue קיים ב-ICONS', () => {
    // SvgIcon מחזיר `ICONS[name] || ''`, ולכן שם שגוי = כפתור בלי אייקון,
    // בלי שגיאה ובלי שאף אחד ישים לב. זו הבדיקה שמונעת את זה.
    const missing = [...used.entries()]
      .filter(([name]) => !(name in ICONS))
      .map(([name, files]) => `${name} (${[...new Set(files)].join(', ')})`);
    expect(missing).toEqual([]);
  });

  it('אין אייקון מוגדר שאינו בשימוש, למעט רשימת PLANNED_ICONS', () => {
    const planned = new Set<string>(PLANNED_ICONS);
    const orphans = NAMES.filter((n) => !referenced.has(n) && !planned.has(n));
    expect(orphans).toEqual([]);
  });

  it('כל שם ב-PLANNED_ICONS מוגדר, ואינו בשימוש בפועל', () => {
    // אייקון מתוכנן שכבר חובר לפקד צריך לצאת מהרשימה, אחרת ההחרגה מתרחבת
    // בשקט ומכסה גם אייקונים שנשכחו.
    expect(PLANNED_ICONS.filter((n) => !(n in ICONS))).toEqual([]);
    expect(PLANNED_ICONS.filter((n) => referenced.has(n))).toEqual([]);
  });

  it('פעולות הקובץ מקבלות אייקונים נפרדים ונכונים', () => {
    // ארבע הפעולות האלה חלקו קודם אייקונים שגויים: „פתח קובץ” הציג זכוכית
    // מגדלת, „מסמך חדש” הציג את לוגו האפליקציה, ו„שמור בשם” ו„ייצוא ל-Word”
    // הציגו את אותו אייקון בדיוק.
    // הנרמול הוא מה שמפעיל את הבדיקה בכלל: קובצי המקור נשמרים ב-CRLF,
    // וההשוואה למטה מצפה למפריד שורה יחיד. בלעדיו הבדיקה נכשלה תמיד —
    // כלומר החיווט של לשונית "קובץ" לא היה שמור בפועל.
    const fileTab = readFileSync(join(SRC, 'ui/ribbon/tabs/FileTab.vue'), 'utf8').replace(
      /\r\n/g,
      '\n'
    );
    for (const [icon, label] of [
      ['newDoc', 'מסמך חדש'],
      ['folder', 'פתח קובץ'],
      ['save', 'שמור'],
      ['saveAs', 'שמור בשם...'],
      ['export', 'ייצוא ל-Word'],
    ] as const) {
      expect(fileTab, label).toContain(`icon="${icon}"\n        label="${label}"`);
    }
    const icons = [...fileTab.matchAll(/icon="([A-Za-z]+)"/g)].map((m) => m[1]!);
    expect(new Set(icons).size, 'אין אייקון כפול בלשונית קובץ').toBe(icons.length);
  });

  it('trackChanges ו-info אינם אותו path', () => {
    // הם היו זהים לחלוטין, ולכן „עקוב אחר שינויים” קיבל אייקון מידע.
    expect(ICONS.trackChanges).not.toBe(ICONS.info);
  });

  it('אין שני אייקונים עם אותו path בכלל', () => {
    const byPath = new Map<string, string[]>();
    for (const name of NAMES) {
      const paths = [...ICONS[name]!.matchAll(/\sd="([^"]+)"/g)].map((m) => m[1]).join('|');
      byPath.set(paths, [...(byPath.get(paths) ?? []), name]);
    }
    const dupes = [...byPath.values()].filter((g) => g.length > 1).map((g) => g.join(' = '));
    expect(dupes).toEqual([]);
  });
});
