/**
 * השלב המקדים הוא הדבר היחיד שעומד בין המשתמש לבין קיפאון שאין ממנו יציאה:
 * `<w:defaultTabStop w:val="0"/>` שולח את המנוע ללולאה על החוט הראשי, ולכן
 * `OPEN_TIMEOUT_MS` אינו יכול לירות ואין שום שגיאה להציג. מכאן שתי החובות
 * שהבדיקות כאן שומרות עליהן:
 *
 * 1. **הערך אכן מתוקן** — אחרת אין לשלב הזה טעם.
 * 2. **שום מסמך אחר אינו נפגע** — הארכיון שנכתב מחדש חייב לשמור את כל שאר
 *    החלקים בייט-בבייט, ומסמך שאין בו מה לתקן חייב לעבור הלאה כפי שהוא, ולא
 *    להיכתב מחדש „ליתר ביטחון”.
 *
 * ה-zip נבנה כאן ביד, עם `deflateRawSync` אמיתי, מפני שהקורא שנבדק קורא
 * מהספרייה המרכזית ומכותרות מקומיות — ומבנה מזויף היה בודק את עצמו.
 *
 * מאז שהתיקון השני נכנס — `<w:b/>` שמושלם לצד `<w:bCs/>`, כדי שהדגשה של כתב
 * מורכב תגיע למסך — יש כאן חובה שלישית, והיא זו שקשה: **מה שאין לגעת בו**.
 * `w:bCs` מופיע גם בתוך `w:rPrChange`, שהוא העיצוב שהיה *לפני* שינוי מסומן;
 * כתיבה שם משנה היסטוריה, ומסמך עם מעקב שינויים הוא בדיוק המסמך שאין רשות
 * לשבור. לכן הבדיקות כאן מודדות גם קינון, גם `w:val` מכובה, וגם מסמך שאין בו
 * מה לתקן ולכן אינו נכתב מחדש בכלל.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { Blob as NativeBlob } from 'node:buffer';
import { deflateRawSync } from 'node:zlib';
import {
  COMPLEX_SCRIPT_BOLD_NOTICE,
  DEFAULT_TAB_STOP_TWIPS,
  FONT_TABLE_PART,
  preflightDocx,
  preflightSource,
  readDocxPart,
  repairComplexScriptBold,
  repairSettings,
  SETTINGS_PART,
} from '../../src/engine/docx-preflight';
import { readZip, readZipComment } from '../../src/engine/docx-parts';
import { CONTENT_PARTS, crc32 as moduleCrc32 } from '../../src/engine/docx-parts';
import { NO_VBA } from '../../src/engine/vba-import';

const SETTINGS_WITH_ZERO =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:settings xmlns:w="ns"><w:zoom w:percent="100"/><w:defaultTabStop w:val="0"/>' +
  '<w:characterSpacingControl w:val="doNotCompress"/></w:settings>';

const DOCUMENT_XML = '<w:document xmlns:w="ns"><w:body><w:p/></w:body></w:document>';

interface Part {
  name: string;
  content: string;
  /** `false` מאחסן את החלק כמות שהוא — כמו ש-DOCX אמיתי עושה לחלקים קטנים. */
  deflate?: boolean;
  /** שדות ה-extra והערה, כמו שכלי אריזה אחרים כותבים אותם. */
  localExtra?: Uint8Array;
  centralExtra?: Uint8Array;
  comment?: Uint8Array;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.byteLength; i++) {
    let value = (crc ^ bytes[i]) & 0xff;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    crc = value ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** ארכיון ZIP מינימלי אך אמיתי: כותרות מקומיות, ספרייה מרכזית ו-EOCD. */
const NO_BYTES = new Uint8Array(0);

function buildZip(parts: Part[], archiveComment: Uint8Array = NO_BYTES): Uint8Array<ArrayBuffer> {
  const encoder = new TextEncoder();
  const records = parts.map((part) => {
    const raw = encoder.encode(part.content);
    const stored = part.deflate === false ? raw : new Uint8Array(deflateRawSync(raw));
    return {
      nameBytes: encoder.encode(part.name),
      raw,
      stored,
      method: part.deflate === false ? 0 : 8,
      crc: crc32(raw),
      localExtra: part.localExtra ?? NO_BYTES,
      centralExtra: part.centralExtra ?? NO_BYTES,
      comment: part.comment ?? NO_BYTES,
    };
  });

  const size =
    records.reduce(
      (total, r) => total + 30 + r.nameBytes.byteLength + r.localExtra.byteLength + r.stored.byteLength,
      0,
    ) +
    records.reduce(
      (total, r) => total + 46 + r.nameBytes.byteLength + r.centralExtra.byteLength + r.comment.byteLength,
      0,
    ) +
    22 +
    archiveComment.byteLength;

  const out = new Uint8Array(size);
  const view = new DataView(out.buffer);
  const offsets: number[] = [];
  let at = 0;

  for (const record of records) {
    offsets.push(at);
    view.setUint32(at, 0x04034b50, true);
    view.setUint16(at + 4, 20, true);
    view.setUint16(at + 8, record.method, true);
    view.setUint32(at + 14, record.crc, true);
    view.setUint32(at + 18, record.stored.byteLength, true);
    view.setUint32(at + 22, record.raw.byteLength, true);
    view.setUint16(at + 26, record.nameBytes.byteLength, true);
    view.setUint16(at + 28, record.localExtra.byteLength, true);
    at += 30;
    out.set(record.nameBytes, at);
    at += record.nameBytes.byteLength;
    out.set(record.localExtra, at);
    at += record.localExtra.byteLength;
    out.set(record.stored, at);
    at += record.stored.byteLength;
  }

  const centralOffset = at;
  records.forEach((record, index) => {
    view.setUint32(at, 0x02014b50, true);
    view.setUint16(at + 4, 20, true);
    view.setUint16(at + 6, 20, true);
    view.setUint16(at + 10, record.method, true);
    view.setUint32(at + 16, record.crc, true);
    view.setUint32(at + 20, record.stored.byteLength, true);
    view.setUint32(at + 24, record.raw.byteLength, true);
    view.setUint16(at + 28, record.nameBytes.byteLength, true);
    view.setUint16(at + 30, record.centralExtra.byteLength, true);
    view.setUint16(at + 32, record.comment.byteLength, true);
    view.setUint32(at + 42, offsets[index], true);
    at += 46;
    out.set(record.nameBytes, at);
    at += record.nameBytes.byteLength;
    out.set(record.centralExtra, at);
    at += record.centralExtra.byteLength;
    out.set(record.comment, at);
    at += record.comment.byteLength;
  });

  view.setUint32(at, 0x06054b50, true);
  view.setUint16(at + 8, records.length, true);
  view.setUint16(at + 10, records.length, true);
  view.setUint32(at + 12, at - centralOffset, true);
  view.setUint32(at + 16, centralOffset, true);
  view.setUint16(at + 20, archiveComment.byteLength, true);
  out.set(archiveComment, at + 22);
  return out;
}

function contains(haystack: Uint8Array, needle: Uint8Array): boolean {
  outer: for (let start = 0; start + needle.byteLength <= haystack.byteLength; start++) {
    for (let i = 0; i < needle.byteLength; i++) {
      if (haystack[start + i] !== needle[i]) continue outer;
    }
    return true;
  }
  return false;
}

const bytesOf = (text: string): Uint8Array<ArrayBuffer> => new TextEncoder().encode(text);

/**
 * החלק כטקסט, דרך אותו קורא שהמוצר משתמש בו. החלק המתוקן נכתב **דחוס**, ולכן
 * חיפוש טקסט גלוי בבייטים של הארכיון היה מוכיח רק שהוא *לא* דחוס.
 */
const partText = (bytes: Uint8Array<ArrayBuffer>, name: string) => readDocxPart(bytes, name);

/**
 * jsdom אינו מממש `Blob.arrayBuffer`, שקיים בכל דפדפן מ-2019. בלי ההשלמה הזאת
 * מסלול הטיוטה — היחיד שמגיע ל-`preflightSource` כ-Blob ולא כ-URL — היה נופל
 * בשקט למסלול „לא הצלחתי לקרוא, מחזיר את המקור” ולא נבדק כלל.
 */
if (typeof Blob.prototype.arrayBuffer !== 'function') {
  Object.defineProperty(Blob.prototype, 'arrayBuffer', {
    configurable: true,
    value(this: Blob) {
      return new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as ArrayBuffer);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(this);
      });
    },
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('repairSettings', () => {
  it('מחליפה אפס בברירת המחדל של Word', () => {
    const repaired = repairSettings(SETTINGS_WITH_ZERO);
    expect(repaired).toContain(`<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`);
    expect(repaired).not.toContain('w:val="0"');
    // כל השאר נשאר: התיקון אינו כותב מחדש קובץ הגדרות.
    expect(repaired).toContain('<w:zoom w:percent="100"/>');
    expect(repaired).toContain('<w:characterSpacingControl w:val="doNotCompress"/>');
  });

  it('אינה נוגעת בערך תקין', () => {
    expect(repairSettings(SETTINGS_WITH_ZERO.replace('"0"', '"720"'))).toBeNull();
    expect(repairSettings(SETTINGS_WITH_ZERO.replace('"0"', '"425"'))).toBeNull();
  });

  it('אינה נוגעת בקובץ שאין בו המאפיין', () => {
    expect(repairSettings('<w:settings xmlns:w="ns"><w:zoom w:percent="100"/></w:settings>')).toBeNull();
  });

  it('תופסת גם ערך שלילי וגם ערך שאינו מספר', () => {
    // שניהם מגיעים למנוע כאותו צעד שאינו מקדם, ולכן אינם שונים מאפס.
    expect(repairSettings(SETTINGS_WITH_ZERO.replace('"0"', '"-1"'))).toContain('"720"');
    expect(repairSettings(SETTINGS_WITH_ZERO.replace('"0"', '""'))).toContain('"720"');
  });

  it('סובלנית לרווחים ולסדר מאפיינים', () => {
    const odd = '<w:settings xmlns:w="ns"><w:defaultTabStop w:foo="x" w:val = "0" /></w:settings>';
    expect(repairSettings(odd)).toContain(`"${DEFAULT_TAB_STOP_TWIPS}"`);
  });

  /* ------- שלוש דרכים שבהן הקיפאון לא נמנע, וסקירה יריבה מצאה אותן -------
   *
   * לא „תיקון פחות מדויק”: בכל השלוש התיקון פשוט לא קרה, והמשתמש נשאר עם
   * אוצריא תקועה שאין ממנה יציאה. לכן כל אחת מהן היא בדיקה משלה.                */

  it('הערה שיש בה ערך תקין אינה מסתירה את הערך האמיתי שאחריה', () => {
    // הרגקס הקודם לקח את ההתאמה הראשונה — זו שבהערה — ראה 720, והחזיר null.
    const shadowed =
      '<w:settings xmlns:w="ns"><!-- <w:defaultTabStop w:val="720"/> -->' +
      '<w:defaultTabStop w:val="0"/></w:settings>';
    const repaired = repairSettings(shadowed);
    expect(repaired).not.toBeNull();
    // ההערה נשארת תו-בתו; מה שהשתנה הוא רק האלמנט החי.
    expect(repaired).toContain('<!-- <w:defaultTabStop w:val="720"/> -->');
    expect(repaired).toContain(`<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`);
  });

  it('הערה שיש בה הערך המקפיא אינה נערכת — אין שם עיצוב', () => {
    const inComment =
      '<w:settings xmlns:w="ns"><!-- <w:defaultTabStop w:val="0"/> --></w:settings>';
    expect(repairSettings(inComment)).toBeNull();
  });

  it('צורה שאינה סוגרת את עצמה מתוקנת, ונשארת זוג פתיחה-סגירה', () => {
    const paired =
      '<w:settings xmlns:w="ns"><w:defaultTabStop w:val="0"></w:defaultTabStop></w:settings>';
    const repaired = repairSettings(paired);
    expect(repaired).toBe(
      '<w:settings xmlns:w="ns">' +
        `<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"></w:defaultTabStop></w:settings>`,
    );
  });

  it('`>` בתוך ערך של מאפיין אחר אינו מסתיר את האלמנט', () => {
    const tricky = '<w:settings xmlns:w="ns"><w:defaultTabStop w:foo="a>b" w:val="0"/></w:settings>';
    const repaired = repairSettings(tricky);
    expect(repaired).toContain('w:foo="a>b"');
    expect(repaired).toContain(`w:val="${DEFAULT_TAB_STOP_TWIPS}"`);
  });

  /* ------------------- מה שנשמר, ולא נכתב מחדש ------------------- */

  it('קידומת שאינה `w` נשמרת', () => {
    const other = '<ns0:settings xmlns:ns0="ns"><ns0:defaultTabStop ns0:val="0"/></ns0:settings>';
    expect(repairSettings(other)).toContain(`<ns0:defaultTabStop ns0:val="${DEFAULT_TAB_STOP_TWIPS}"/>`);
  });

  it('מרכאות בודדות נקראות, ונשארות בודדות', () => {
    const single = "<w:settings xmlns:w='ns'><w:defaultTabStop w:val='0'/></w:settings>";
    expect(repairSettings(single)).toContain(`w:val='${DEFAULT_TAB_STOP_TWIPS}'`);
  });

  it('מרכאות בודדות עם ערך תקין אינן נוגעות', () => {
    // הרגקס הקודם קרא רק מרכאות כפולות, ולכן ראה „אין ערך” וכתב 720 על ערך
    // תקין לחלוטין — כתיבה מיותרת לתוך המסמך של המשתמש.
    const single = "<w:settings xmlns:w='ns'><w:defaultTabStop w:val='425'/></w:settings>";
    expect(repairSettings(single)).toBeNull();
  });

  it('אלמנט בלי `w:val` כלל מקבל אותו, בקידומת שלו', () => {
    const bare = '<w:settings xmlns:w="ns"><w:defaultTabStop/></w:settings>';
    expect(repairSettings(bare)).toBe(
      `<w:settings xmlns:w="ns"><w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/></w:settings>`,
    );
  });

  it('רווח לפני `/>` אינו מוכפל כשהמאפיין מתווסף', () => {
    expect(repairSettings('<w:settings><w:defaultTabStop /></w:settings>')).toBe(
      `<w:settings><w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}" /></w:settings>`,
    );
  });

  it('זוג פתיחה-סגירה בלי `w:val` מקבל אותו בפתיחה', () => {
    expect(repairSettings('<w:defaultTabStop></w:defaultTabStop>')).toBe(
      `<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"></w:defaultTabStop>`,
    );
  });

  it('ערך שאינו מספר, ערך שגולש, וערך עם סימן פלוס', () => {
    expect(repairSettings('<w:defaultTabStop w:val="abc"/>')).toContain(`"${DEFAULT_TAB_STOP_TWIPS}"`);
    expect(repairSettings('<w:defaultTabStop w:val="1e999"/>')).toContain(`"${DEFAULT_TAB_STOP_TWIPS}"`);
    expect(repairSettings('<w:defaultTabStop w:val="+720"/>')).toBeNull();
  });

  it('הוראת עיבוד שיש בה האלמנט אינה נערכת', () => {
    const pi = '<w:settings xmlns:w="ns"><?tool <w:defaultTabStop w:val="0"/> ?></w:settings>';
    expect(repairSettings(pi)).toBeNull();
  });

  it('הצהרת ה-XML אינה מפריעה לתיקון שאחריה', () => {
    // ההצהרה היא הוראת עיבוד, וההוראות נבלעות שלמות — צריך לוודא שהבליעה
    // אינה בולעת גם את מה שבא אחריה.
    expect(repairSettings(SETTINGS_WITH_ZERO)).toContain(`"${DEFAULT_TAB_STOP_TWIPS}"`);
  });
});

/** `rPr` שלמה, כדי שהבדיקות יקראו כמו ה-XML האמיתי ולא כמו קטעים. */
const runProps = (inner: string) => `<w:r><w:rPr>${inner}</w:rPr><w:t>שלום</w:t></w:r>`;

describe('repairComplexScriptBold', () => {
  it('משלימה `w:b` לצד `bCs` בודדת', () => {
    const repaired = repairComplexScriptBold(runProps('<w:bCs/>'));
    // מיד לפני `bCs`, מפני ש-CT_RPr היא רצף ו-`b` בא לפניה.
    expect(repaired).toContain('<w:rPr><w:b/><w:bCs/></w:rPr>');
  });

  it('שומרת על סדר האלמנטים גם כשיש שכנים לפני ואחרי', () => {
    const repaired = repairComplexScriptBold(
      runProps('<w:rFonts w:cs="David"/><w:bCs/><w:i/><w:szCs w:val="28"/>'),
    );
    expect(repaired).toContain('<w:rFonts w:cs="David"/><w:b/><w:bCs/><w:i/>');
  });

  it('מתקנת כל ה-rPr במסמך, ולא רק את הראשונה', () => {
    const repaired = repairComplexScriptBold(
      runProps('<w:bCs/>') + runProps('<w:iCs/>') + runProps('<w:bCs/>'),
    );
    expect(repaired!.match(/<w:b\/>/g)).toHaveLength(2);
  });

  it('אינה נוגעת ב-rPr שיש בה כבר `w:b`', () => {
    expect(repairComplexScriptBold(runProps('<w:b/><w:bCs/>'))).toBeNull();
    expect(repairComplexScriptBold(runProps('<w:b w:val="1"/><w:bCs/>'))).toBeNull();
  });

  it('אינה הופכת `w:b` שכובה במפורש', () => {
    // אמירה מפורשת של מי שכתב את הקובץ: „לא מודגש בלטינית, מודגש בעברית”.
    // אנחנו משלימים מה שנעדר, לא מבטלים מה שנכתב.
    expect(repairComplexScriptBold(runProps('<w:b w:val="0"/><w:bCs/>'))).toBeNull();
  });

  it('אינה נוגעת ב-`bCs` שכובה', () => {
    for (const off of ['0', 'false', 'off', 'FALSE']) {
      expect(repairComplexScriptBold(runProps(`<w:bCs w:val="${off}"/>`))).toBeNull();
    }
  });

  it('מתקנת `bCs` שדולקת במפורש', () => {
    for (const on of ['1', 'true', 'on']) {
      expect(repairComplexScriptBold(runProps(`<w:bCs w:val="${on}"/>`))).toContain('<w:b/>');
    }
  });

  it('אינה כותבת בתוך `w:rPrChange` — זו היסטוריה של שינוי מסומן', () => {
    const tracked = `<w:r><w:rPr><w:rPrChange w:id="1" w:author="x"><w:rPr><w:bCs/></w:rPr></w:rPrChange></w:rPr></w:r>`;
    expect(repairComplexScriptBold(tracked)).toBeNull();
  });

  it('מתקנת את הריצה עצמה גם כשיש לצדה `rPrChange`', () => {
    // המקרה שרגקס לא-להוט על `<w:rPr>…</w:rPr>` היה טועה בו: הוא היה קושר את
    // הפתיחה החיצונית לסגירה הפנימית, ומודד את ההיסטוריה במקום את הריצה.
    const tracked =
      `<w:r><w:rPr><w:bCs/>` +
      `<w:rPrChange w:id="1" w:author="x"><w:rPr><w:i/></w:rPr></w:rPrChange>` +
      `</w:rPr><w:t>שלום</w:t></w:r>`;
    const repaired = repairComplexScriptBold(tracked);
    expect(repaired).toContain('<w:rPr><w:b/><w:bCs/>');
    // וההיסטוריה נשארה כפי שהייתה.
    expect(repaired).toContain('<w:rPrChange w:id="1" w:author="x"><w:rPr><w:i/></w:rPr>');
  });

  it('מתקנת סגנון, שזה המקרה שדווח', () => {
    // בדיוק מה שיש בקובץ שדווח: „כותרת 2” עם `bCs` ובלי `b`.
    const styles =
      '<w:styles xmlns:w="ns"><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/>' +
      '<w:rPr><w:rFonts w:cs="FrankRuehl DP"/><w:bCs/><w:szCs w:val="28"/></w:rPr></w:style></w:styles>';
    expect(repairComplexScriptBold(styles)).toContain('<w:b/><w:bCs/>');
  });

  it('`b` ו-`bCs` בצורת זוג פתיחה-סגירה נקראות כמו סוגרות-עצמן', () => {
    expect(repairComplexScriptBold(runProps('<w:b></w:b><w:bCs/>'))).toBeNull();
    expect(repairComplexScriptBold(runProps('<w:bCs></w:bCs>'))).toContain(
      '<w:rPr><w:b/><w:bCs></w:bCs></w:rPr>',
    );
  });

  it('ילד-הרחבה בתוך `rPr` (w14) אינו מבלבל את ההכנסה', () => {
    const repaired = repairComplexScriptBold(
      runProps(
        '<w:bCs/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="FF0000"/>' +
          '</w14:solidFill></w14:textFill>',
      ),
    )!;
    expect(repaired.split('<w:b/>')).toHaveLength(2);
    expect(repaired).toContain('<w:rPr><w:b/><w:bCs/><w14:textFill>');
  });

  it('ריצה בתוך נוסחה (OMML): `m:rPr` אינה נוגעת, ו-`w:rPr` שלצדה מתוקנת', () => {
    const math =
      '<m:oMath><m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:bCs/></w:rPr>' +
      '<m:t>x</m:t></m:r></m:oMath>';
    const repaired = repairComplexScriptBold(math)!;
    expect(repaired).toContain('<m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:b/><w:bCs/></w:rPr>');
    expect(repaired.split('<w:b/>')).toHaveLength(2);
  });

  it('אינה נוגעת במסמך שאין בו `bCs` בכלל', () => {
    expect(repairComplexScriptBold(runProps('<w:b/><w:i/>'))).toBeNull();
    expect(repairComplexScriptBold(DOCUMENT_XML)).toBeNull();
  });

  it('אינה מתקנת `rPr` ריקה או סוגרת-עצמה', () => {
    expect(repairComplexScriptBold('<w:r><w:rPr/><w:t>שלום</w:t></w:r>')).toBeNull();
    expect(repairComplexScriptBold('<w:r><w:rPr />' + '<w:t>שלום</w:t></w:r>')).toBeNull();
  });

  /* ---------------- מה שהתיקון אסור לו לגעת בו ---------------- */

  it('אינה כותבת בתוך הערת XML', () => {
    // ההערה נראית לרגקס בדיוק כמו תגים, וכתיבה בתוכה היא עריכה של טקסט ולא
    // של עיצוב.
    const commented = `<w:body><!-- ${runProps('<w:bCs/>')} --></w:body>`;
    expect(repairComplexScriptBold(commented)).toBeNull();
  });

  it('אינה כותבת בתוך CDATA', () => {
    const cdata = `<w:r><w:rPr><w:i/></w:rPr><w:t><![CDATA[${runProps('<w:bCs/>')}]]></w:t></w:r>`;
    expect(repairComplexScriptBold(cdata)).toBeNull();
  });

  it('הערה או CDATA אינן מכבות תיקון אמיתי שאחריהן', () => {
    // המקרה שקבר את הגרסה הקודמת בצורה אחרת: `</w:rPr>` שיושבת בתוך הערה
    // הוציאה מהמחסנית את ה-scope החי, והתיקון האמיתי נעלם.
    const mixed = `<w:body><!-- </w:rPr> -->${runProps('<w:bCs/>')}</w:body>`;
    expect(repairComplexScriptBold(mixed)).toContain('<w:b/><w:bCs/>');
  });

  it('`>` בתוך ערך מאפיין אינו מבלבל את הסורק', () => {
    // XML חוקי לגמרי. רגקס שנעצר על ה-`>` הראשון קרא כאן תג אחר לגמרי.
    const odd =
      '<w:r><w:rPr><w:rStyle w:val="a>b"/><w:bCs/></w:rPr><w:t>שלום</w:t></w:r>';
    expect(repairComplexScriptBold(odd)).toContain('<w:b/><w:bCs/>');
  });

  it('`rPrChange` שאינו נקרא כמצופה אינו מכבה את שאר החלק', () => {
    // ההתנהגות שהייתה הגרועה מכולן: מונה `rPrChange` שנתקע על 1 והשתיק את
    // התיקון מאותו בייט ועד סוף הקובץ, בשקט. הכלל היום הוא קינון, ולכן
    // ה-`rPr` שבתוך ההיסטוריה מדולגת בזכות עומקה ולא בזכות שם התג.
    const brokenName =
      '<w:body><w:r><w:rPr><w:rPrChange w:id="1" w:author="a>b">' +
      '<w:rPr><w:bCs/></w:rPr></w:rPrChange></w:rPr></w:r>' +
      runProps('<w:bCs/>') +
      '</w:body>';
    const repaired = repairComplexScriptBold(brokenName);
    expect(repaired).toContain('<w:b/><w:bCs/>');
    // ההיסטוריה עצמה נשארה כפי שהייתה.
    expect(repaired).toContain('<w:rPrChange w:id="1" w:author="a>b"><w:rPr><w:bCs/></w:rPr>');
  });

  it('סגירת `rPr` תלושה אינה מפילה את המונה מתחת לאפס', () => {
    // וזה מה שהיה מרשה ל-`rPrChange` **האמיתי** הבא להיכתב לתוכו.
    const stray =
      '</w:rPr><w:r><w:rPr><w:rPrChange w:id="2" w:author="x">' +
      '<w:rPr><w:bCs/></w:rPr></w:rPrChange></w:rPr></w:r>';
    expect(repairComplexScriptBold(stray)).toBeNull();
  });

  it('`rPr` שאינה נסגרת בסוף החלק מאבדת רק את עצמה', () => {
    const truncated = `<w:body>${runProps('<w:bCs/>')}<w:r><w:rPr><w:bCs/>`;
    const repaired = repairComplexScriptBold(truncated);
    expect(repaired!.match(/<w:b\/>/g)).toHaveLength(1);
  });

  it('`rPr` שאינה נסגרת **באמצע** מכבה את מה שאחריה — וזה מתועד, לא מתוקן', () => {
    // סקירה יריבה מצאה את זה, וזו אותה צורת כשל שכלל העומק נבחר כדי למנוע
    // („מונה שנתקע מכבה את התיקון עד סוף החלק”) — רק שהטריגר אחר: פתיחה בלי
    // סגירה משאירה את העומק על 1 לנצח, ולכן כל `bCs` שאחריה נראית כהיסטוריה.
    //
    // לא מתוקן, וזו החלטה: כדי להתאושש צריך לדעת **איפה** ה-`rPr` הפתוחה
    // הייתה נסגרת, וזה בדיוק מה שאין ב-XML שאינו תקין. המנוע עצמו אינו פורס
    // מסמך כזה, ולכן אין כאן הדגשה שאובדת למשתמש — יש מסמך שאינו נפתח.
    // הבדיקה קיימת כדי שההתנהגות תהיה נמדדת ולא מופתעת.
    const broken = `<w:body><w:r><w:rPr><w:bCs/><w:r><w:rPr><w:bCs/></w:rPr></w:r>`;
    expect(repairComplexScriptBold(broken)).toBeNull();
  });

  /* ---------------- קידומות ---------------- */

  it('מתקנת גם כשמרחב השמות קשור לקידומת אחרת, ובאותה קידומת', () => {
    // החבילה רשאית לקשור את WordprocessingML לכל קידומת. `w:b` שנכתבת לתוך
    // מסמך כזה שייכת למרחב שמות אחר לגמרי — כלומר לא הדגשה.
    const other = '<ns0:r><ns0:rPr><ns0:bCs/></ns0:rPr><ns0:t>שלום</ns0:t></ns0:r>';
    expect(repairComplexScriptBold(other)).toContain('<ns0:b/><ns0:bCs/>');
  });

  it('אינה מוסיפה `b` שנייה כשהקיימת נושאת קידומת אחרת', () => {
    // שתי `b` באותה `rPr` הן `CT_RPr` פסולה.
    expect(repairComplexScriptBold('<w:rPr><ns0:b/><w:bCs/></w:rPr>')).toBeNull();
  });

  /* ---------------- ערכי ST_OnOff ---------------- */

  it('קוראת `w:val` גם במרכאות בודדות', () => {
    // אחרת „לא מודגש” שנכתב במפורש היה נהפך למודגש — שינוי במסמך.
    expect(repairComplexScriptBold(runProps("<w:bCs w:val='0'/>"))).toBeNull();
    expect(repairComplexScriptBold(runProps("<w:bCs w:val='1'/>"))).toContain('<w:b/>');
    expect(repairComplexScriptBold(runProps("<w:b w:val='0'/><w:bCs/>"))).toBeNull();
  });

  it('שתי `bCs` באותה rPr — האחרונה קובעת', () => {
    expect(repairComplexScriptBold(runProps('<w:bCs w:val="0"/><w:bCs/>'))).toContain('<w:b/>');
    expect(repairComplexScriptBold(runProps('<w:bCs/><w:bCs w:val="0"/>'))).toBeNull();
  });

  it('שתי `bCs` — ההוספה לפני ה**ראשונה**, מפני ה-`xsd:sequence`', () => {
    // הבדיקה שמעל מודדת איזה **ערך** מכריע; זו מודדת **איפה** ההוספה נוחתת,
    // וזה חצי אחר של אותה החלטה. `CT_RPr` היא `xsd:sequence` שבו `b` בא לפני
    // `bCs`, ולכן הוספה לפני ה-`bCs` השנייה מציבה `b` אחרי `bCs` — חלק
    // שהסכימה של Word פוסלת. קלט עם שתי `bCs` פסול מלכתחילה, אבל התוצאה אסור
    // לה להיות פסולה יותר ממנו.
    const repaired = repairComplexScriptBold(runProps('<w:bCs w:val="0"/><w:bCs/>'));
    expect(repaired).toContain('<w:rPr><w:b/><w:bCs w:val="0"/><w:bCs/></w:rPr>');
    expect(repaired, 'ולא בין השתיים').not.toContain('<w:bCs w:val="0"/><w:b/>');
  });

  /* ---------------- אינווריאנטים ---------------- */

  it('אינה משנה דבר מלבד ה-`b` שהוסיפה', () => {
    // `toContain` אינו יכול לתפוס פלט שכל השאר בו נהרס. זה כן.
    const source =
      '<w:body>' +
      runProps('<w:rFonts w:cs="David"/><w:bCs/><w:szCs w:val="28"/>') +
      runProps('<w:b/><w:bCs/>') +
      '<w:p><w:pPr><w:rPr><w:bCs/></w:rPr></w:pPr></w:p>' +
      '</w:body>';
    const repaired = repairComplexScriptBold(source)!;
    const strip = (text: string) => text.split('<w:b/>').join('');
    expect(strip(repaired)).toBe(strip(source));
  });

  it('מעבר שני אינו מוצא מה לתקן', () => {
    const once = repairComplexScriptBold(runProps('<w:bCs/>') + runProps('<w:bCs/>'))!;
    expect(repairComplexScriptBold(once)).toBeNull();
  });

  it('סורקת בזמן לינארי, גם על חלק בגודל של ספר', () => {
    // המודול הזה קיים מפני שחסימה של החוט הראשי אינה נתפסת אחר כך. סריקה
    // ריבועית כאן הייתה בדיוק אותו כשל, במסלול חדש.
    const unit = runProps('<w:bCs/>');
    const small = unit.repeat(2_000);
    const large = unit.repeat(8_000);

    const time = (xml: string): number => {
      const started = performance.now();
      expect(repairComplexScriptBold(xml)).not.toBeNull();
      return performance.now() - started;
    };
    time(small); // חימום, שלא נמדוד את ההידור הראשון
    const smallMs = Math.max(time(small), 1);
    const largeMs = time(large);

    // פי ארבעה קלט. סף רחב בכוונה — זה שער נגד ריבועיות, לא מדידת ביצועים.
    expect(largeMs / smallMs).toBeLessThan(12);
  });

  it('הפלט הוא XML תקין', () => {
    const source = `<w:document xmlns:w="ns"><w:body>${runProps('<w:bCs/>')}</w:body></w:document>`;
    const repaired = repairComplexScriptBold(source)!;
    const parsed = new DOMParser().parseFromString(repaired, 'application/xml');
    expect(parsed.querySelector('parsererror')).toBeNull();
  });
});

describe('crc32', () => {
  // `crc32` המקומי בקובץ הזה מחשב את הפולינום בכל בייט מחדש, בלי טבלה — ולכן
  // הוא מימוש **בלתי תלוי** במה שבמודול, ולא העתקה שלו. CRC שגוי הוא ארכיון
  // שבור, וזה כשל שקט.
  //
  // הבדיקות כאן נשארות גם אחרי ש-slice-by-4 הוסר מהמודול (הוא נמדד איטי
  // יותר; ראו את ההערה שם). זה בדיוק מה שהן נועדו לתפוס: מי שיכתוב שוב
  // מימוש „מהיר” ויטעה בשארית או בסדר הבייטים ייפול כאן.

  it('מסכימה עם וקטור הבדיקה המקובל', () => {
    // "123456789" → 0xCBF43926, הווקטור של CRC-32/ISO-HDLC.
    expect(moduleCrc32(bytesOf('123456789'))).toBe(0xcbf43926);
  });

  it('מסכימה עם מימוש הייחוס בכל אורך, ובפרט באלה שאינם כפולה של ארבע', () => {
    // מימוש שמתקדם יותר מבייט בשלב נשבר בשארית, ולכן כל אורך קצר נבדק.
    for (let length = 0; length <= 16; length++) {
      const bytes = new Uint8Array(length) as Uint8Array<ArrayBuffer>;
      for (let i = 0; i < length; i++) bytes[i] = (i * 31 + 7) & 0xff;
      expect(moduleCrc32(bytes)).toBe(crc32(bytes));
    }

    const big = new Uint8Array(5_003) as Uint8Array<ArrayBuffer>;
    for (let i = 0; i < big.byteLength; i++) big[i] = (i * 131 + 17) & 0xff;
    expect(moduleCrc32(big)).toBe(crc32(big));
  });

  it('מסכימה עם מימוש הייחוס על טקסט עברי', () => {
    // עברית היא שני בייטים לתו, ולכן גם הגושים וגם השארית נראים אחרת.
    const bytes = bytesOf('שבועת הדיינין, מודה במקצת');
    expect(moduleCrc32(bytes)).toBe(crc32(bytes));
  });
});

describe('CONTENT_PARTS', () => {
  it('תופס את החלקים שיש בהם תכונות ריצה', () => {
    for (const name of [
      'word/document.xml',
      'word/styles.xml',
      'word/numbering.xml',
      'word/footnotes.xml',
      'word/endnotes.xml',
      'word/comments.xml',
      'word/header1.xml',
      'word/footer3.xml',
      // שמות שאינם מה ש-Word כותב אבל חוקיים ב-OPC: הוא נפתר דרך ה-rels ואינו
      // מחייב ספרה, ושמות חלקים אינם תלויי רישיות.
      'word/header.xml',
      'word/footer.xml',
      'word/document2.xml',
      'Word/Document.xml',
      // חלק אמיתי של Word 2010, ובו גיליון סגנונות שלם.
      'word/stylesWithEffects.xml',
      // הספרה מותרת בכל שם שיכול לשאת אותה, ולא רק בארבעה מהם.
      'word/footnotes2.xml',
      'word/endnotes2.xml',
      'word/comments2.xml',
      'word/numbering2.xml',
    ]) {
      expect(CONTENT_PARTS.test(name)).toBe(true);
    }
  });

  it('אינו תופס חלקים שאין להם מה לתרום', () => {
    // `settings.xml` יש לו תיקון משלו; `glossary` אינו מרונדר; `styles` שאינו
    // תחת `word/` אינו החלק שלנו.
    for (const name of [
      SETTINGS_PART,
      'word/fontTable.xml',
      'word/glossary/document.xml',
      'word/theme/theme1.xml',
      'customXml/item1.xml',
      'styles.xml',
      // ההתאמה היא על שם החלק כולו, ולא על חלק ממנו.
      'word/document.xml.rels',
      'word/_rels/document.xml.rels',
      'word/../word/document.xml',
      'word/documentation.xml',
    ]) {
      expect(CONTENT_PARTS.test(name)).toBe(false);
    }
  });
});

describe('preflightDocx', () => {
  it('BOM של חלק שתוקן אינו נמחק', async () => {
    // ההבטחה בכותרת הקובץ היא „זהה למקור בכל מה שאינו התיקון עצמו”, ו-BOM הוא
    // חלק ממנה. ברירת המחדל של `TextDecoder` פושטת U+FEFF ו-`TextEncoder`
    // אינו מחזיר אותו, ולכן חלק שהתחיל ב-BOM היה נכתב מחדש בלעדיו.
    const repaired = await preflightDocx(
      buildZip([
        { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO },
        { name: 'word/styles.xml', content: `\ufeff<w:styles xmlns:w="ns"><w:style><w:rPr><w:bCs/></w:rPr></w:style></w:styles>` },
      ]),
    );

    expect(repaired).not.toBeNull();
    const styles = await readDocxPart(repaired!.bytes, 'word/styles.xml');
    expect(styles?.startsWith('\ufeff'), 'ה-BOM שרד את הכתיבה מחדש').toBe(true);
    expect(styles).toContain('<w:b/><w:bCs/>');
  });

  it('מתקנת את ההגדרות ומשאירה את שאר החלקים כפי שהיו', async () => {
    const original = buildZip([
      { name: '[Content_Types].xml', content: '<Types/>' },
      { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO },
      { name: 'word/document.xml', content: DOCUMENT_XML, deflate: false },
    ]);

    const repaired = await preflightDocx(original);
    expect(repaired).not.toBeNull();

    const settings = await partText(repaired!.bytes, SETTINGS_PART);
    expect(settings).toContain(`<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`);
    expect(settings).not.toContain('<w:defaultTabStop w:val="0"/>');
    // והחלק המתוקן נכתב דחוס: הטקסט אינו גלוי בבייטים של הארכיון.
    expect(contains(repaired!.bytes, bytesOf('<w:defaultTabStop'))).toBe(false);

    // החלקים האחרים עברו בייט-בבייט: גם הדחוס וגם זה שלא.
    expect(contains(repaired!.bytes, new Uint8Array(deflateRawSync(bytesOf('<Types/>'))))).toBe(true);
    expect(contains(repaired!.bytes, bytesOf(DOCUMENT_XML))).toBe(true);

    // היומן מדווח מה שונה, ולא רק שנגענו.
    expect(repaired!.notes).toEqual([expect.stringContaining('defaultTabStop')]);
  });

  it('שדות ה-extra, הערות הרשומות והערת הארכיון שורדים את הכתיבה', async () => {
    // הכותב רץ על **כל** הרשומות, גם על אלה שלא נגענו בהן, ולכן מה
    // שהוא משמיט אובד מהמסמך. Info-ZIP, 7-Zip וכלים אחרים כותבים שם
    // חותמות זמן והרשאות, וחבילות OPC מיוחדות נושאות שם מידע משלן.
    const localExtra = new Uint8Array([0x55, 0x54, 0x05, 0x00, 0x03, 0x11, 0x22, 0x33, 0x44]);
    const centralExtra = new Uint8Array([0x55, 0x54, 0x01, 0x00, 0x03]);
    const entryComment = bytesOf('הערה על הרשומה');
    const archiveComment = bytesOf('הערת הארכיון');

    const repaired = await preflightDocx(
      buildZip(
        [
          { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO, localExtra, centralExtra, comment: entryComment },
          { name: 'word/document.xml', content: DOCUMENT_XML, localExtra, centralExtra, comment: entryComment },
        ],
        archiveComment,
      ),
    );
    expect(repaired).not.toBeNull();

    const entries = readZip(repaired!.bytes as Uint8Array<ArrayBuffer>);
    expect(entries, 'הארכיון נקרא בחזרה').not.toBeNull();
    expect(entries).toHaveLength(2);
    for (const entry of entries!) {
      expect(Array.from(entry.localExtra ?? []), `extra מקומי של ${entry.name}`).toEqual(
        Array.from(localExtra),
      );
      expect(Array.from(entry.centralExtra ?? []), `extra מרכזי של ${entry.name}`).toEqual(
        Array.from(centralExtra),
      );
      expect(Array.from(entry.comment ?? []), `הערה של ${entry.name}`).toEqual(
        Array.from(entryComment),
      );
    }
    expect(
      Array.from(readZipComment(repaired!.bytes as Uint8Array<ArrayBuffer>)),
      'הערת הארכיון',
    ).toEqual(Array.from(archiveComment));

    // והתיקון עצמו עדיין קרה.
    const settings = await partText(repaired!.bytes, SETTINGS_PART);
    expect(settings).toContain(`<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`);
  });

  it('הארכיון שנכתב נקרא בחזרה, ואין בו יותר מה לתקן', async () => {
    const repaired = await preflightDocx(
      buildZip([
        { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO },
        { name: 'word/document.xml', content: DOCUMENT_XML },
      ]),
    );
    expect(repaired).not.toBeNull();
    // אותה קריאה בדיוק על התוצאה: מוכיחה גם שהמבנה תקין וגם שהתיקון הושלם.
    await expect(preflightDocx(repaired!.bytes)).resolves.toBeNull();
  });

  it('מתקנת גם סגנונות, ומדווחת על שני התיקונים בנפרד', async () => {
    const styles =
      '<w:styles xmlns:w="ns"><w:style w:styleId="2"><w:rPr><w:bCs/></w:rPr></w:style></w:styles>';
    const repaired = await preflightDocx(
      buildZip([
        { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO },
        { name: 'word/styles.xml', content: styles },
        { name: 'word/document.xml', content: DOCUMENT_XML },
      ]),
    );

    expect(repaired).not.toBeNull();
    await expect(partText(repaired!.bytes, 'word/styles.xml')).resolves.toContain(
      '<w:rPr><w:b/><w:bCs/></w:rPr>',
    );
    expect(repaired!.notes).toHaveLength(2);
    expect(repaired!.notes[1]).toContain('word/styles.xml');
    // ההודעה למשתמש היא של התיקון שנכתב לתוך המסמך, ולא נגזרת מנוסח היומן.
    expect(repaired!.notice).toBe(COMPLEX_SCRIPT_BOLD_NOTICE);
    // ושוב על התוצאה: אין יותר מה לתקן, לא בהגדרות ולא בסגנונות.
    await expect(preflightDocx(repaired!.bytes)).resolves.toBeNull();
  });

  it('חלק שאין בו `bCs` אינו נכתב מחדש בכלל', async () => {
    // חשוב לא פחות מהתיקון: מסמך עברי רגיל אינו משלם על התיקון הזה כלום. הגוף
    // עובר בייט-בבייט, כלומר גם דחוס כפי שהיה.
    const body =
      '<w:document xmlns:w="ns"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr>' +
      '<w:t>מודגש</w:t></w:r></w:p></w:body></w:document>';
    const repaired = await preflightDocx(
      buildZip([
        { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO },
        { name: 'word/document.xml', content: body },
      ]),
    );

    expect(repaired!.notes).toEqual([expect.stringContaining('defaultTabStop')]);
    expect(contains(repaired!.bytes, new Uint8Array(deflateRawSync(bytesOf(body))))).toBe(true);
  });

  it('תיקון שאינו נראה למשתמש אינו מפיק הודעה', async () => {
    const repaired = await preflightDocx(buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO }]));
    expect(repaired!.notes).toHaveLength(1);
    expect(repaired!.notice).toBeNull();
  });

  it('החלק המתוקן נכתב דחוס, וקטן מהמקור הגלוי', async () => {
    // גוף גדול ובעל חזרות, כמו מסמך אמיתי: כאן ההפרש בין `STORED` לדחוס הוא
    // ההפרש שנמדד על ספרים — פי 5 עד פי 15 בזיכרון של המנוע.
    const body =
      '<w:document xmlns:w="ns"><w:body>' +
      runProps('<w:bCs/>').repeat(2000) +
      '</w:body></w:document>';
    const original = buildZip([{ name: 'word/document.xml', content: body }]);

    const repaired = await preflightDocx(original);
    expect(repaired).not.toBeNull();
    // קטן בהרבה מהטקסט עצמו — ולא רק „לא גדול ממנו”.
    expect(repaired!.bytes.byteLength).toBeLessThan(body.length / 4);
    // ונקרא בחזרה בדיוק, דרך המסלול שהמנוע קורא בו.
    const restored = await partText(repaired!.bytes, 'word/document.xml');
    expect(restored).toBe(repairComplexScriptBold(body));
  });

  it('בלי `CompressionStream` החלק נכתב גלוי, והתיקון עדיין קורה', async () => {
    vi.stubGlobal('CompressionStream', undefined);
    const repaired = await preflightDocx(
      buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO }]),
    );
    expect(repaired).not.toBeNull();
    // `STORED`: הטקסט המתוקן גלוי בבייטים.
    expect(contains(repaired!.bytes, bytesOf(`<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`))).toBe(true);
    await expect(preflightDocx(repaired!.bytes)).resolves.toBeNull();
  });

  it('דוחס שאינו משחזר את המקור נופל ל-`STORED`, ולא למסמך', async () => {
    // דוחס שמחזיר זבל תקין-למראה: הפריסה שלו לא תיתן את המקור, וזה בדיוק
    // המקרה שהאימות קיים בשבילו.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    class BrokenCompression extends TransformStream<BufferSource, Uint8Array> {
      constructor() {
        super({
          transform(_chunk, controller) {
            controller.enqueue(new Uint8Array(deflateRawSync(bytesOf('לא המקור'))));
          },
        });
      }
    }
    vi.stubGlobal('CompressionStream', BrokenCompression);

    const repaired = await preflightDocx(
      buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO }]),
    );
    expect(repaired).not.toBeNull();
    expect(contains(repaired!.bytes, bytesOf(`<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`))).toBe(true);
    await expect(partText(repaired!.bytes, SETTINGS_PART)).resolves.toContain(
      `w:val="${DEFAULT_TAB_STOP_TWIPS}"`,
    );
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('הדחיסה לא שחזרה'));
    warn.mockRestore();
  });

  it('אינה נוגעת במסמך תקין', async () => {
    const clean = buildZip([
      { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO.replace('"0"', '"720"') },
      { name: 'word/document.xml', content: DOCUMENT_XML },
    ]);
    await expect(preflightDocx(clean)).resolves.toBeNull();
  });

  it('אינה נוגעת במסמך שאין בו קובץ הגדרות', async () => {
    const noSettings = buildZip([{ name: 'word/document.xml', content: DOCUMENT_XML }]);
    await expect(preflightDocx(noSettings)).resolves.toBeNull();
  });

  it('מוותרת בשקט על מה שאינו ארכיון', async () => {
    await expect(preflightDocx(bytesOf('לא zip בכלל'))).resolves.toBeNull();
  });
});

describe('preflightSource', () => {
  // Response.blob() comes from Node's Fetch implementation, so use its Blob
  // realm too. The browser path retains that Blob before transferring bytes.
  beforeEach(() => vi.stubGlobal('Blob', NativeBlob));
  it('מסמך ריק עובר כפי שהוא', async () => {
    await expect(preflightSource(undefined)).resolves.toEqual({
      source: undefined,
      fontTable: null,
      vba: NO_VBA,
      notice: null,
    });
  });

  it('מסמך בלי מאקרו מדווח שאין בו מאקרו', async () => {
    const clean = buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO.replace('"0"', '"720"') }]);
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(clean)),
    );

    const { vba } = await preflightSource('http://127.0.0.1:1/doc.docx');
    // `hasMacroPart` הוא מה שקובע את סיומת השמירה, ולכן חשוב שהוא יהיה שקר
    // כאן: מסמך רגיל אינו אמור להיות נשמר כ-`.docm`.
    expect(vba.hasMacroPart).toBe(false);
    expect(vba.status).toBeNull();
  });

  it('מסמך שההדגשה בו הושלמה מדווח על כך למשתמש', async () => {
    const withBoldCs = buildZip([
      { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO.replace('"0"', '"720"') },
      {
        name: 'word/styles.xml',
        content: '<w:styles xmlns:w="ns"><w:style><w:rPr><w:bCs/></w:rPr></w:style></w:styles>',
      },
    ]);
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(withBoldCs)),
    );

    const { notice } = await preflightSource('http://127.0.0.1:1/doc.docx');
    expect(notice).toBe(COMPLEX_SCRIPT_BOLD_NOTICE);
  });

  it('תיקון שאינו נראה למשתמש אינו מדווח לו', async () => {
    // `defaultTabStop` הוא ההפרש בין מסמך שנפתח למסמך שקופא, ואינו משנה דבר
    // במסמך עצמו. שורת מצב שמדווחת על כל דבר היא שורת מצב שאיש אינו קורא.
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO }]))),
    );

    const { notice } = await preflightSource('http://127.0.0.1:1/doc.docx');
    expect(notice).toBeNull();
  });

  it('קריאה שנכשלה אינה מדווחת על מאקרו שלא נבדקו', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => {
        throw new Error('אין רשת');
      }),
    );

    // „לא יודעים” אינו „יש מאקרו”: דיווח חיובי כאן היה משנה סיומת של מסמך
    // שכלל לא נקרא.
    const { vba } = await preflightSource('http://127.0.0.1:1/doc.docx');
    expect(vba).toEqual(NO_VBA);
  });

  it('ספרייה אחת מזינה גם את טבלת הגופנים וגם את התיקונים', async () => {
    // הקוראים חולקים את הרשומות שנקראו פעם אחת (ראו `preflightSource`), ולכן
    // מה שנמדד כאן הוא שאף אחד מהם לא נפגע מהשיתוף: התיקון נכתב, טבלת
    // הגופנים חוזרת, והתוצאה זהה לקריאה נפרדת של כל אחד מהם.
    const fontTable = '<w:fonts xmlns:w="ns"><w:font w:name="David"/></w:fonts>';
    const both = buildZip([
      { name: FONT_TABLE_PART, content: fontTable },
      { name: SETTINGS_PART, content: SETTINGS_WITH_ZERO },
    ]);
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(both)),
    );

    const result = await preflightSource('http://127.0.0.1:1/doc.docx');

    expect(result.fontTable).toBe(await readDocxPart(both, FONT_TABLE_PART));
    expect(result.fontTable).toBe(fontTable);
    const separately = await preflightDocx(both);
    expect(separately).not.toBeNull();
    expect(result.notice).toBe(separately?.notice ?? null);
    expect(new Uint8Array(await (result.source as Blob).arrayBuffer())).toEqual(separately!.bytes);
  });

  it('URL שאין בו מה לתקן נמסר כבייטים שנקראו — בלי קריאה שנייה מהמנוע', async () => {
    const clean = buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO.replace('"0"', '"720"') }]);
    const fetchSpy = vi.fn(async () => new Response(clean));
    vi.stubGlobal('fetch', fetchSpy);
    const { source } = await preflightSource('http://127.0.0.1:1/doc.docx');
    expect(source).toBeInstanceOf(Blob);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(new Uint8Array(await (source as Blob).arrayBuffer())).toEqual(clean);
  });

  it('טיוטה שאין בה מה לתקן חוזרת כמות שהיא, בלי העתקה', async () => {
    const draft = new Blob([buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO.replace('"0"', '"720"') }])]);
    const { source } = await preflightSource(draft);
    expect(source).toBe(draft);
  });

  it('URL שיש בו מה לתקן מוחלף בבייטים מתוקנים', async () => {
    const broken = buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO }]);
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(broken)),
    );

    const { source } = await preflightSource('http://127.0.0.1:1/doc.docx');
    expect(source).toBeInstanceOf(Blob);
    const bytes = new Uint8Array(await (source as Blob).arrayBuffer());
    await expect(partText(bytes, SETTINGS_PART)).resolves.toContain(
      `<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`,
    );
  });

  it('שלב מקדים שזורק מחזיר את המקור, ואינו מונע פתיחה', async () => {
    // הזריקה היחידה שנשארה בפנים היא הקצאה שנכשלה; כאן היא מדומה על המקודד.
    // „לתקן, ולא לחסום”: מסמך שהיה נפתח בלי השלב הזה נפתח בלעדיו.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const bytes = buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO }]);
    // המקודד נשבר רק **אחרי** שהבייטים נקראו: כך הכשל נופל בתוך `preflightDocx`
    // ולא במסלול הקריאה, שיש לו טיפול משלו.
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => ({
        blob: async () => {
          vi.stubGlobal(
            'TextEncoder',
            class {
              encode(): never {
                throw new RangeError('Array buffer allocation failed');
              }
            },
          );
          return new Blob([bytes]);
        },
      })),
    );

    const { source, notice } = await preflightSource('http://127.0.0.1:1/doc.docx');
    expect(source).toBeInstanceOf(Blob);
    expect(new Uint8Array(await (source as Blob).arrayBuffer())).toEqual(bytes);
    expect(notice).toBeNull();
    expect(warn).toHaveBeenCalledWith(
      expect.stringContaining('הבדיקה המקדימה נכשלה'),
      // לא `expect.any(RangeError)`: המחלקה ב-jsdom היא של realm אחר.
      expect.objectContaining({ message: 'Array buffer allocation failed' }),
    );
    warn.mockRestore();
  });

  it('קריאה שנכשלה מחזירה את המקור, ואינה מונעת פתיחה', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => {
        throw new Error('אין רשת');
      }),
    );
    const { source } = await preflightSource('http://127.0.0.1:1/doc.docx');
    expect(source).toBe('http://127.0.0.1:1/doc.docx');
  });

  it('טיוטה שיש בה מה לתקן מתוקנת גם היא', async () => {
    // טיוטת שחזור מגיעה כ-Blob ולא כ-URL, וגם היא עלולה לשאת את הערך.
    const draft = new Blob([buildZip([{ name: SETTINGS_PART, content: SETTINGS_WITH_ZERO }])]);
    const { source } = await preflightSource(draft);
    expect(source).not.toBe(draft);
    const bytes = new Uint8Array(await (source as Blob).arrayBuffer());
    await expect(partText(bytes, SETTINGS_PART)).resolves.toContain(
      `<w:defaultTabStop w:val="${DEFAULT_TAB_STOP_TWIPS}"/>`,
    );
  });
});

/**
 * הסורק מול פרסר XML אמיתי.
 *
 * כל הבדיקות שלמעלה נכתבו מול מקרים ש**נחשבו מראש**, וזו בדיוק החולשה שלהן:
 * הסורק כבר נשבר פעם אחת בשש נקודות שאיש לא חשב עליהן. לכן כאן מימוש ייחוס
 * עצמאי — `DOMParser`, שאינו יודע דבר על רגקסים, הערות וקידומות — ומסמכים
 * שנבנים באקראי מהחלקים שמסמך אמיתי בנוי מהם: `rPrChange` מקננת, הערות
 * ו-CDATA שיש בהם מה שנראה כמו עיצוב, `>` בתוך ערך של מאפיין, ושתי קידומות.
 *
 * הזרע קבוע, ולכן כשל כאן משוחזר.
 */
describe('repairComplexScriptBold מול מימוש ייחוס', () => {
  const OFF = new Set(['0', 'false', 'off']);

  /** כמה `w:b` צריך להשלים, לפי הכלל, כפי שפרסר אמיתי רואה אותו. */
  function referenceInserts(xml: string): number {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    expect(doc.querySelector('parsererror')).toBeNull();

    let needed = 0;
    for (const element of Array.from(doc.getElementsByTagName('*'))) {
      if (element.localName !== 'rPr') continue;
      // `rPr` בתוך `rPr` היא היסטוריה של שינוי מסומן, ואינה עיצוב חי.
      let ancestor = element.parentElement;
      while (ancestor && ancestor.localName !== 'rPr') ancestor = ancestor.parentElement;
      if (ancestor) continue;

      let hasBold = false;
      let boldCsOn: boolean | null = null;
      for (const child of Array.from(element.children)) {
        if (child.localName === 'b') hasBold = true;
        else if (child.localName === 'bCs') {
          const value = Array.from(child.attributes).find((a) => a.localName === 'val')?.value;
          boldCsOn = value === undefined || !OFF.has(value.trim().toLowerCase());
        }
      }
      if (boldCsOn && !hasBold) needed += 1;
    }
    return needed;
  }

  /** מחולל דטרמיניסטי — mulberry32. */
  function random(seed: number): () => number {
    return () => {
      seed = (seed + 0x6d2b79f5) | 0;
      let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  it('מסכים איתו על 400 מסמכים שנבנו באקראי', () => {
    const next = random(20260903);
    const pick = <T,>(options: T[]): T => options[Math.floor(next() * options.length)];

    // ההיסטוריה נכתבת עם `>` בתוך שם המחבר בכוונה: זה XML חוקי, והוא מה
    // שהפיל את הגרסה הראשונה של הסורק.
    const history = (p: string) =>
      `<${p}:rPrChange ${p}:id="7" ${p}:author="a>b"><${p}:rPr><${p}:bCs/><${p}:sz ${p}:val="28"/>` +
      `</${p}:rPr></${p}:rPrChange>`;

    let withInserts = 0;
    let inserted = 0;
    for (let document = 0; document < 400; document++) {
      const p = pick(['w', 'ns0']);
      const body: string[] = [];

      for (let block = 0; block < 1 + Math.floor(next() * 6); block++) {
        const props = [
          pick(['', `<${p}:rFonts ${p}:cs="David"/>`]),
          pick(['', '', `<${p}:b/>`, `<${p}:b ${p}:val="0"/>`]),
          pick([
            `<${p}:bCs/>`,
            `<${p}:bCs/>`,
            `<${p}:bCs ${p}:val="1"/>`,
            `<${p}:bCs ${p}:val='0'/>`,
            `<${p}:bCs ${p}:val="off"/>`,
            '',
          ]),
          pick(['', `<${p}:szCs ${p}:val="28"/>`]),
          pick(['', history(p)]),
        ].join('');

        body.push(
          pick([
            `<${p}:r><${p}:rPr>${props}</${p}:rPr><${p}:t>שלום</${p}:t></${p}:r>`,
            `<${p}:p><${p}:pPr><${p}:rPr>${props}</${p}:rPr></${p}:pPr></${p}:p>`,
            `<${p}:sdt><${p}:sdtPr><${p}:rPr>${props}</${p}:rPr></${p}:sdtPr></${p}:sdt>`,
            `<${p}:style><${p}:rPr>${props}</${p}:rPr></${p}:style>`,
            `<!-- <${p}:rPr><${p}:bCs/></${p}:rPr> -->`,
            `<${p}:r><${p}:t><![CDATA[<${p}:rPr><${p}:bCs/></${p}:rPr>]]></${p}:t></${p}:r>`,
            `<?otzaria <${p}:bCs/> ?>`,
            `<${p}:rPr/>`,
          ]),
        );
      }

      const xml = `<${p}:document xmlns:${p}="ns"><${p}:body>${body.join('')}</${p}:body></${p}:document>`;
      const expected = referenceInserts(xml);
      const repaired = repairComplexScriptBold(xml);
      const actual = repaired === null ? 0 : (repaired.length - xml.length) / `<${p}:b/>`.length;

      expect([document, actual]).toEqual([document, expected]);
      if (repaired === null) continue;

      withInserts += 1;
      inserted += actual;
      // הפלט תקין, לא נגע בשום דבר מלבד ה-b שהוסיף, ואין בו יותר מה לתקן.
      expect(referenceInserts(repaired)).toBe(0);
      const strip = (text: string) => text.split(`<${p}:b/>`).join('');
      expect(strip(repaired)).toBe(strip(xml));
    }

    // הסף הוא נגד בדיקה ריקה: מחולל שהפסיק לייצר `bCs` היה „מסכים” על הכול.
    expect(withInserts).toBeGreaterThan(50);
    expect(inserted).toBeGreaterThan(100);
  });
});

/**
 * המקומות ש-`rPr` יושבת בהם במסמך אמיתי, מלבד ריצה.
 *
 * הכלל שהסורק אוכף הוא „`rPr` בתוך `rPr` היא היסטוריה”, וכל השאר הוא עיצוב חי
 * שיש לתקן. הבדיקות כאן הן הצד השני של הכלל: לא „מה אין לגעת בו”, אלא שהצמצום
 * לא בלע גם עיצוב חי שיושב במכולה שלא חשבו עליה.
 */
describe('repairComplexScriptBold — מכולות שאינן ריצה', () => {
  const cases: [string, string][] = [
    ['סימן פסקה', '<w:p><w:pPr><w:rPr><w:bCs/></w:rPr></w:pPr></w:p>'],
    [
      'סימן פסקה עם מעקב שינויים לצדו',
      // `w:ins` ו-`w:del` הם איברים חוקיים ב-`CT_ParaRPr`, ואינם מקננים `rPr`.
      '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="a"/><w:bCs/></w:rPr></w:pPr></w:p>',
    ],
    ['רמה במספור', '<w:num><w:lvl w:ilvl="0"><w:rPr><w:bCs/></w:rPr></w:lvl></w:num>'],
    ['סגנון', '<w:style w:styleId="Heading2"><w:rPr><w:bCs/></w:rPr></w:style>'],
    [
      // ברירת המחדל של המסמך כולו — אותו כלל, בסקלה של כל ריצה לטינית במסמך.
      // ראו את ההסתייגות בכותרת המודול.
      'ברירות המחדל של המסמך',
      '<w:docDefaults><w:rPrDefault><w:rPr><w:bCs/></w:rPr></w:rPrDefault></w:docDefaults>',
    ],
    [
      'סגנון טבלה מותנה',
      '<w:style><w:tblStylePr w:type="firstRow"><w:rPr><w:bCs/></w:rPr></w:tblStylePr></w:style>',
    ],
    ['פקד תוכן', '<w:sdt><w:sdtPr><w:rPr><w:bCs/></w:rPr></w:sdtPr></w:sdt>'],
    ['ריצה שנמחקה במעקב שינויים', '<w:del><w:r><w:rPr><w:bCs/></w:rPr></w:r></w:del>'],
  ];

  it.each(cases)('%s', (_name, inner) => {
    expect(repairComplexScriptBold(`<w:document xmlns:w="ns">${inner}</w:document>`)).toContain(
      '<w:b/><w:bCs/>',
    );
  });

  it('קינון עמוק של היסטוריה אינו נספר כעיצוב חי', () => {
    const inner =
      '<w:rPr><w:bCs/><w:rPrChange w:author="a"><w:rPr><w:bCs/>' +
      '<w:rPrChange w:author="b"><w:rPr><w:bCs/></w:rPr></w:rPrChange>' +
      '</w:rPr></w:rPrChange></w:rPr>';
    const repaired = repairComplexScriptBold(`<w:document xmlns:w="ns">${inner}</w:document>`)!;
    expect(repaired.split('<w:b/>')).toHaveLength(2);
    expect(repaired).toContain('<w:rPr><w:b/><w:bCs/><w:rPrChange');
  });
});

describe('איחוד ריצות בתוך השלב המקדים', () => {
  const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const run = (t: string) => `<w:r w:rsidR="1"><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
  const documentWith = (body: string) => `<w:document xmlns:w="${NS}"><w:body><w:p>${body}</w:p></w:body></w:document>`;
  const cleanSettings = `<w:settings xmlns:w="${NS}"><w:proofState w:spelling="clean" w:grammar="clean"/></w:settings>`;

  it('סימוני איות נשמרים במקומם והגדרות clean אינן משתנות', async () => {
    const start = '<w:proofErr w:type="spellStart"/>';
    const end = '<w:proofErr w:type="spellEnd"/>';
    const repaired = await preflightDocx(
      buildZip([
        { name: SETTINGS_PART, content: cleanSettings },
        { name: 'word/document.xml', content: documentWith(start + run('א') + run('ב') + end + run('ג')) },
      ]),
    );
    expect(repaired).not.toBeNull();
    const document = await readDocxPart(repaired!.bytes, 'word/document.xml');
    expect(document).toContain(start + '<w:r w:rsidR="1"><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">אב</w:t></w:r>' + end);
    expect(await readDocxPart(repaired!.bytes, SETTINGS_PART)).toBe(cleanSettings);
  });

  it('גם בלי settings, איחוד rsid אינו מוחק גבולות איות', async () => {
    const start = '<w:proofErr w:type="spellStart"/>';
    const end = '<w:proofErr w:type="spellEnd"/>';
    const repaired = await preflightDocx(
      buildZip([{ name: 'word/document.xml', content: documentWith(run('א') + start + run('ב') + run('ג') + end) }]),
    );
    const document = await readDocxPart(repaired!.bytes, 'word/document.xml');
    expect(document).toContain(run('א') + start);
    expect(document).toContain('<w:t xml:space="preserve">בג</w:t></w:r>' + end);
    expect(await readDocxPart(repaired!.bytes, SETTINGS_PART)).toBeNull();
  });

  it('איחוד שלא הסיר סימון איות אינו נוגע בהגדרות', async () => {
    const repaired = await preflightDocx(
      buildZip([
        { name: SETTINGS_PART, content: cleanSettings },
        { name: 'word/document.xml', content: documentWith(run('א') + run('ב')) },
      ]),
    );
    expect(repaired).not.toBeNull();
    expect(await readDocxPart(repaired!.bytes, SETTINGS_PART)).toBe(cleanSettings);
  });

  it('תיקון defaultTabStop אינו משנה מצב איות או גבולות סימון', async () => {
    const settings = `<w:settings xmlns:w="${NS}"><w:defaultTabStop w:val="0"/><w:proofState w:spelling="clean"/></w:settings>`;
    const repaired = await preflightDocx(
      buildZip([
        { name: SETTINGS_PART, content: settings },
        { name: 'word/document.xml', content: documentWith(run('א') + '<w:proofErr w:type="spellEnd"/>' + run('ב')) },
      ]),
    );
    const out = await readDocxPart(repaired!.bytes, SETTINGS_PART);
    expect(out).toContain(`w:val="${DEFAULT_TAB_STOP_TWIPS}"`);
    expect(out).toContain('w:spelling="clean"');
    expect(await readDocxPart(repaired!.bytes, 'word/document.xml')).toBe(documentWith(run('א') + '<w:proofErr w:type="spellEnd"/>' + run('ב')));
  });
});
