/**
 * סימני פיסוק בדיבור: כשאומרים "פסיק" נכתב `,` ולא המילה, וכך "נקודה" ו"סימן שאלה".
 *
 * ההחלפה נעשית על הטקסט שהשירות מחזיר, לפני הכתיבה למסמך. הסימן נצמד למילה
 * שלפניו, ורווח בא רק אחריו.
 *
 * חריגים: "פסיק רישיה" ו"פסיק רישא" הם מונח הלכתי ולא פקודה, ונשארים כמו שהם.
 */

/** פקודות של שתי מילים — נבדקות לפני פקודות של מילה אחת. */
const TWO_WORDS: Record<string, string> = {
  'סימן שאלה': '?',
  'סימן קריאה': '!',
  'סימן פסיק': ',',
  'סימן נקודה': '.',
};

const ONE_WORD: Record<string, string> = {
  פסיק: ',',
  נקודה: '.',
  נקודתיים: ':',
};

/** מילים שאחרי "פסיק" הופכות אותו למונח ולא לפקודה. */
const NOT_A_COMMAND_AFTER_PSIK = new Set(['רישיה', 'רישא', 'רישה']);

const PUNCT = new Set([',', '.', '?', '!', ':']);

export function isPunctuation(token: string): boolean {
  return token.length === 1 && PUNCT.has(token);
}

/** מחליפה את פקודות הפיסוק שבטקסט בסימנים. */
export function applySpokenPunctuation(text: string): string {
  const words = text.split(/\s+/).filter((w) => w);
  const out: string[] = [];
  for (let i = 0; i < words.length; i++) {
    const word = words[i]!;
    const next = words[i + 1];
    const pair = next !== undefined ? TWO_WORDS[`${word} ${next}`] : undefined;
    if (pair) {
      out.push(pair);
      i += 1;
      continue;
    }
    const single = ONE_WORD[word];
    if (single && !(word === 'פסיק' && next !== undefined && NOT_A_COMMAND_AFTER_PSIK.has(next))) {
      out.push(single);
      continue;
    }
    out.push(word);
  }
  // סימן נצמד למילה שלפניו; סימן בתחילת הטקסט נשאר לבדו (הרווח מטופל בכתיבה).
  let joined = '';
  for (const token of out) {
    if (joined === '') joined = token;
    else if (isPunctuation(token)) joined += token;
    else joined += ' ' + token;
  }
  return joined;
}

/** פקודות ירידת שורה (כמו Enter) — נבדקות כרצף מילים. */
const LINE_COMMANDS = [['תרד', 'שורה'], ['שורה', 'חדשה']];

/**
 * מפצלת את הטקסט בפקודות ירידת שורה. "א ב תרד שורה ג" → ["א ב", "ג"];
 * פקודה בסוף נותנת מקטע אחרון ריק. בלי פקודה — מקטע אחד.
 */
export function splitLineCommands(text: string): string[] {
  const words = text.split(/\s+/).filter((w) => w);
  const pieces: string[] = [];
  let current: string[] = [];
  for (let i = 0; i < words.length; i++) {
    const hit = LINE_COMMANDS.find((c) => c.every((w, k) => words[i + k] === w));
    if (hit) {
      pieces.push(current.join(' '));
      current = [];
      i += hit.length - 1;
    } else {
      current.push(words[i]!);
    }
  }
  pieces.push(current.join(' '));
  return pieces;
}
