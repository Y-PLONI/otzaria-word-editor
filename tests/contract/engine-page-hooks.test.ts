/**
 * העיגון שהסרגל נשען עליו, מול האריזה עצמה.
 *
 * הסרגל מודד את המלבן של העמוד המצויר (engine/page-ruler.ts, ושם גם ההסבר
 * למה מדידה ולא חישוב). המדידה נשענת על שני דברים שאינם חוזה מתועד:
 *
 *   1. **`data-page-index`** — התכונה שהמנוע מסמן בה כל עמוד. שדרוג שישמיט
 *      אותה יהפוך את הסרגל לרצועה ריקה, בלי שום שגיאה: `querySelector`
 *      מחזיר `null`, והקוד מטפל בזה כ„המסמך עדיין נטען”. כלומר בלי הבדיקה
 *      הזאת התוצאה של שדרוג היא פקד שנעלם בשקט.
 *   2. **`ui.viewport.getHost()`** — כאן דווקא **יש** חוזה מוקלד, והבדיקה
 *      מאמתת שהוא עדיין שם. הוא מה שמאפשר לצמצם את החיפוש ל-host שהמנוע
 *      מצייר בתוכו, במקום לסרוק את המסמך כולו.
 *
 * מה שהבדיקה **אינה** עושה: להריץ את המנוע. היא קוראת מחרוזות מהאריזה ומן
 * הטיפוסים — אותה גישה בדיוק כמו tests/contract/engine-hf-chrome.test.ts,
 * ומאותה סיבה (הרישיון אוסר לשנות את האריזה, ולכן היא נקראת בלבד).
 */
import { describe, it, expect } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { PAGE_INDEX_ATTRIBUTE } from '../../src/engine/page-ruler';

/** vitest רץ משורש המאגר. */
const ENGINE = join(process.cwd(), 'node_modules/@superdoc/docx-engine/dist/docx-engine.es.js');
const SUPERDOC_TYPES = join(process.cwd(), 'node_modules/superdoc/dist/superdoc/src');

const bundle = readFileSync(ENGINE, 'utf8');
/** האריזה הציבורית — זו שאנחנו מייבאים, ושדרכה `activeEditor` מגיע. */
const superdocBundle = readFileSync(
  join(process.cwd(), 'node_modules/superdoc/dist/superdoc.es.js'),
  'utf8',
);

/** כל קובצי הטיפוסים של superdoc — הנתיב הפנימי אינו חוזה, ולכן סורקים. */
function typeDeclarations(dir = SUPERDOC_TYPES): string {
  let text = '';
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) text += typeDeclarations(full);
    else if (entry.name.endsWith('.d.ts')) text += readFileSync(full, 'utf8');
  }
  return text;
}

const types = typeDeclarations();

describe('העיגון של העמוד המצויר', () => {
  it('הקריאה של האריזה אכן הצליחה', () => {
    // שער שקורא קובץ ריק עובר בירוק על כל טענה שאחריו.
    expect(bundle.length).toBeGreaterThan(1_000_000);
  });

  it('המנוע עדיין מסמן כל עמוד ב-`data-page-index`', () => {
    // באריזה זה מופיע כשם ה-dataset (`pageIndex`) וכשם התכונה. די באחד מהם
    // כדי לדעת שהסימון קיים; שניהם נבדקים כדי שהכשל יצביע על השינוי המדויק.
    expect(bundle).toContain('pageIndex');
    expect(bundle.includes(PAGE_INDEX_ATTRIBUTE) || bundle.includes('data-page-index')).toBe(true);
  });

  it('המנוע עצמו מזהה עמוד לפי אותה תכונה — כלומר זה סימון ולא במקרה', () => {
    // `Number.isInteger(Number(el.dataset.pageIndex))` הוא הסינון שהמנוע
    // עושה בעצמו כשהוא מחפש עמודים; אותו מבנה קיים באריזה.
    expect(bundle).toContain("classList']['contains']('superdoc-page')");
  });

  it('`viewport.getHost` עדיין בחוזה הציבורי של superdoc/ui', () => {
    expect(types).toContain('getHost(): HTMLElement | null');
  });

  it('`viewport.observe` עדיין שם — בלעדיו הסרגל לא היה יודע על שינוי זום', () => {
    expect(types).toContain('observe(listener: () => void): () => void');
  });
});

/**
 * העיגון השני: המדידה של השוליים ה**אפקטיביים**.
 *
 * `activeEditor.pageMetrics.getSnapshot()` מחזיר `pages[0].base.marginTopPx` —
 * ושם, בשונה מ-`sections.list()`, מופיע הערך אחרי הרצפה שהמנוע כופה כשיש
 * כותרת עליונה. זה מה שמונע מהסרגל להבטיח שוליים שהטקסט לא יזוז אליהם.
 *
 * כמו `data-page-index`, גם זה אינו חוזה מוקלד: `pageMetrics` אינו מופיע
 * ב-`.d.ts` של `superdoc`. הקריאה עצמה מתגוננת ויש לה נפילה אחורה
 * (`readEffectiveMargins` ב-page-setup.ts), ולכן שדרוג שישמיט אותה **לא**
 * ישבור כלום — הסרגל פשוט יחזור לשקר בשקט. הבדיקה הזאת היא מה שהופך את
 * ההשמטה לרועשת.
 */
describe('המדידה של השוליים האפקטיביים', () => {
  it('`pageMetrics` עדיין באריזה של superdoc', () => {
    // כאן הוא נחשף על `activeEditor`, וזה הצד שאנחנו קוראים.
    expect(bundle.length).toBeGreaterThan(1_000_000);
    expect(superdocBundle).toContain('pageMetrics');
    expect(superdocBundle).toContain('getSnapshot');
  });

  it('התצלום עדיין נושא את השוליים בפיקסלים', () => {
    // `marginTopPx`/`marginBottomPx` הם השדות שהסרגל קורא. הם נבנים במנוע.
    expect(bundle).toContain('marginTopPx');
    expect(bundle).toContain('marginBottomPx');
  });

  it('הרצפה עצמה עדיין בקוד — `headerDistance + גובה הכותרת`', () => {
    // התיעוד של מנוע הפריסה מנסח את זה כ-`Math.max(top, headerDistance + h)`,
    // וזה השם שנושא את התוצאה. בלעדיו אין למי להתאים את החסם.
    expect(bundle).toContain('pendingTopMargin');
    expect(bundle).toContain('activeHeaderDistance');
  });
});
