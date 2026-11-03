<template>
  <div class="ribbon-tab-pane insert-tab">
    <!-- קבוצה 1: עמודים -->
    <RibbonGroup
      title="עמודים"
      icon="pageBreak"
    >
      <!--
        התווית אינה „מעבר עמוד”: המימוש הוא `w:pageBreakBefore`, שמזיז את כל
        הפסקה לעמוד הבא ואינו מפצל אותה בסמן כמו Word. ההסבר המלא, כולל
        החלופות שנדחו, ב-engine/page-break.ts.
      -->
      <RibbonButton
        icon="pageBreak"
        label="התחל בעמוד חדש"
        shortcut-id="page-break"
        variant="large"
        :tooltip="pageBreakTooltip"
        :disabled="!pageBreak.available"
        :active="pageBreakOn"
        @click="onStartOnNewPage"
      />
    </RibbonGroup>

    <!-- קבוצה 2: טבלאות -->
    <RibbonGroup
      title="טבלאות"
      icon="table"
    >
      <TablePicker
        :disabled="!tableCmd.enabled.value"
        @select="onInsertTable"
      />
    </RibbonGroup>

    <!-- קבוצה 3: איורים ומדיה -->
    <RibbonGroup
      title="איורים"
      icon="image"
    >
      <RibbonButton
        icon="image"
        label="תמונות"
        variant="large"
        :tooltip="imageTooltip"
        :disabled="!imageCmd.enabled.value || imageBusy"
        @click="onInsertImage"
      />
    </RibbonGroup>

    <!-- קבוצה 4: קישורים -->
    <RibbonGroup
      title="קישורים"
      icon="link"
    >
      <RibbonButton
        icon="link"
        label="קישור"
        variant="large"
        tooltip="הוספת היפר-קישור לכתובת אינטרנט או לדואר"
        shortcut-id="link"
        :disabled="!linkCmd.enabled.value"
        @click="$emit('open-link')"
      />
      <!--
        „סימנייה” יושבת כאן מפני שזה מקומה ב-Word העברי: הפקד השני בקבוצה
        „קישורים” של לשונית „הוספה”, ליד „קישור” ו„הפניה מקושרת”. הפקד
        השלישי אינו כאן — ראו engine/cross-refs.ts.

        ולכן היא גם נשארה **גדולה** כששאר הלשוניות עברו לדפוס „ראשי + מחסנית”:
        שלושת הפקדים האלה הם שווים ב-Word, ו„ליד” בהערה שלמעלה הוא טענה על
        המקום, לא רק על הקבוצה.
      -->
      <RibbonButton
        icon="bookmark"
        label="סימנייה"
        variant="large"
        :tooltip="bookmarkTooltip"
        :disabled="!can('canManageBookmarks')"
        @click="onOpenBookmarkDialog"
      />

      <!-- גל 22: „הסר קישור" — hyperlinks.remove על הטווח המסומן.
           הוא שיורד למחסנית ולא „סימנייה”: ל-Word אין פקד „הסר קישור” ברצועה
           בכלל (הוא בתפריט ההקשר), ולכן הוא המשני מבין השלושה. -->
      <RibbonStack>
        <RibbonButton
          icon="link"
          label="הסר קישור"
          variant="small"
          tooltip="הסרת ההיפר-קישור מהטקסט המסומן (הטקסט נשמר)"
          :disabled="removeLinkInFlight"
          @click="onRemoveHyperlink"
        />
      </RibbonStack>
    </RibbonGroup>

    <!-- קבוצה 5: כותרת עליונה ותחתונה -->
    <RibbonGroup
      title="כותרת עליונה ותחתונה"
      icon="header"
    >
      <RibbonMenuButton
        icon="header"
        label="כותרת עליונה"
        :tooltip="headerTooltip"
        :disabled="!can('canEditHeaderFooter')"
        :items="headerItems"
        @select="onHeaderAction"
      />
      <RibbonMenuButton
        icon="footer"
        label="כותרת תחתונה"
        :tooltip="footerTooltip"
        :disabled="!can('canEditHeaderFooter')"
        :items="footerItems"
        @select="onFooterAction"
      />
      <!--
        „מספר עמוד” יושב כאן ולא בקבוצה משלו מפני שזה מקומו ב-Word העברי:
        הוא הפקד השלישי בקבוצה „כותרת עליונה ותחתונה”, מפני ששם מספרי העמודים
        חיים בפועל. השדה עצמו נכנס **במקום הסמן** ולא בכותרת התחתונה — אין
        API ציבורי שמזיז את הסמן ל-story אחר (ראו הערת הכותרות למטה) — וזה מה
        שה-tooltip אומר.

        („הרביעי” שנכתב כאן קודם לא היה נכון: בתבנית הוא היה החמישי, אחרי שני
        המתגים שירדו מאז למחסנית שלצד. עכשיו הוא באמת השלישי, כמו ב-Word.)
      -->
      <RibbonMenuButton
        icon="pageNumber"
        label="מספר עמוד"
        :tooltip="tip('canInsertField', 'הכנסת שדה מספר עמוד במקום הסמן')"
        :disabled="!can('canInsertField')"
        :items="pageNumberItems"
        @select="onPageNumberAction"
      />

      <!--
        שלושת המתגים, במחסנית לצד שלושת הפקדים הראשיים. ב-Word הם אינם כפתורי
        ענק אלא תיבות סימון קטנות בקבוצה „אפשרויות” של לשונית ההקשר של
        הכותרות, וזה בדיוק מה שהם: אפשרויות של הכותרת, לא פעולות הוספה.

        זה גם מה שמחזיר את הקבוצה לרוחב סביר — „שונה בעמודים זוגיים
        ואי-זוגיים” ככפתור גדול מתח משבצת על רוחב של שלוש.
      -->
      <RibbonStack>
        <RibbonButton
          icon="firstPageHeader"
          label="שונה בעמוד ראשון"
          variant="small"
          :tooltip="tip('canSetTitlePage', 'לעמוד הראשון תהיה כותרת משלו')"
          :disabled="!can('canSetTitlePage')"
          :active="headerFooter.titlePage"
          @click="onToggleTitlePage"
        />
        <RibbonButton
          icon="oddEvenPages"
          label="שונה בעמודים זוגיים ואי-זוגיים"
          variant="small"
          :tooltip="tip('canSetOddEvenHeaders', 'כותרת אחת לעמודים הזוגיים ואחרת לאי-זוגיים')"
          :disabled="!can('canSetOddEvenHeaders')"
          :active="headerFooter.oddEven"
          @click="onToggleOddEven"
        />
        <RibbonButton
          icon="link"
          label="קשר לקודם"
          variant="small"
          :tooltip="linkToPreviousTooltip"
          :disabled="!can('canLinkToPrevious') || headerFooter.sectionCount < 2"
          :active="headerFooter.linkedToPrevious"
          @click="onToggleLinkToPrevious"
        />
      </RibbonStack>
    </RibbonGroup>

    <!--
      קבוצה 6: טקסט. ב-Word העברי „תאריך ושעה” יושב בקבוצה „טקסט” של לשונית
      „הוספה”, בין „חלקים מהירים” ל„אובייקט” — ו„חלקים מהירים” הוא גם המקום
      שממנו נפתח „שדה…”. „עדכן שדות” אינו פקד ברצועה של Word אלא F9 ותפריט
      הקשר, ואין לו בית טבעי; הוא ממוקם כאן מפני שזו הקבוצה שהשדות נכנסים
      ממנה, וכי „עדכן” לצד „הוסף” הוא הצמד שהמשתמש צריך — במחסנית שלצד הפקד
      הראשי, כי הוא מתחזק ואינו מוסיף.
    -->
    <RibbonGroup
      title="טקסט"
      icon="dateTime"
    >
      <RibbonMenuButton
        split
        icon="dateTime"
        label="תאריך ושעה"
        variant="large"
        :tooltip="dateTooltip"
        :disabled="!can('canInsertField')"
        :items="dateItems"
        @open="dateNow = new Date()"
        @action="onInsertDate"
        @select="onDateAction"
      />
      <RibbonStack>
        <RibbonButton
          icon="updateFields"
          label="עדכן שדות"
          variant="small"
          :tooltip="rebuildTooltip"
          :disabled="!can('canRebuildFields')"
          @click="onRebuildFields"
        />
      </RibbonStack>
    </RibbonGroup>

    <!-- קבוצה 7: תוכן עניינים -->
    <RibbonGroup
      title="תוכן עניינים"
      icon="toc"
    >
      <RibbonButton
        icon="toc"
        label="תוכן עניינים"
        variant="large"
        tooltip="יצירת תוכן עניינים אוטומטי"
        :disabled="!tocCmd.enabled.value"
        @click="onInsertToc"
      />
    </RibbonGroup>

    <BookmarkDialog
      :is-open="bookmarkDialogOpen"
      :names="bookmarks.names"
      @close="bookmarkDialogOpen = false"
      @add="onAddBookmark"
      @remove="onRemoveBookmark"
      @rename="onRenameBookmark"
    />
  </div>
</template>

<script setup lang="ts">
/**
 * „הוספה”.
 *
 * ## התמונה, וההכרעה שקדמה לה
 *
 * `create.image` **מטמיע בייטים ואינו מפנה ל-URL** — נמדד במימוש, וההסבר המלא
 * ב-engine/payloads.ts. בורר הקבצים של אוצריא מחזיר `url` של שרת loopback
 * שהפורט שלו משתנה בכל הפעלה, ולכן הזרימה כאן היא בורר → קריאת בייטים →
 * data URI → פקודה. ה-URL עצמו אינו מגיע למנוע בשום שלב.
 *
 * למה דרך פקודת ה-registry (`image`) ולא `doc.create.image` ישירות: הפקודה היא
 * שמחשבת את `at` מהסמן, כולל המקרים שאין להם מסלול פשוט — סמן בכותרת עליונה או
 * תחתונה מקבל `in: resolveActiveHeaderFooterSlot(story)`. קריאה ישירה הייתה
 * מחייבת אותנו לשחזר את החישוב הזה, ולטעות בו בשקט.
 *
 * ## הקישור, ולמה הוא כבר לא כאן
 *
 * הדיאלוג עבר למעטפת (`composables/use-link-dialog.ts`), כדי ש-`Ctrl+K` יפתח
 * אותו מכל לשונית: לשונית שאינה פעילה אינה מורכבת, ודיאלוג שחי בתוכה אינו
 * קיים. הכפתור כאן רק משדר `open-link`, וההנמקה על תצלום הבחירה — שבלעדיו
 * הקישור נכתב על טווח שכבר אינו קיים — נמצאת שם.
 *
 * ## מעבר העמוד, ומה הוא באמת
 *
 * אין ב-2.8.0 API ציבורי למעבר עמוד בסמן (`<w:br w:type="page"/>`). מה שכן יש
 * הוא `w:pageBreakBefore` על פסקה, וזה מה שהפקד עושה — ולכן התווית היא „התחל
 * בעמוד חדש”. ההנמקה המלאה, כולל שתי החלופות שנדחו ומה שנמצא בצינור הקלט
 * הפנימי של המנוע, ב-engine/page-break.ts.
 *
 * ## הכותרת העליונה והתחתונה, ומה הפקד **אינו** עושה
 *
 * ב-Word „ערוך כותרת עליונה” מוודא שיש חלק כותרת **ומעביר אליו את הסמן**.
 * השני אינו אפשרי כאן: `doc.selection` הוא קריאה בלבד, ואין פעולה ציבורית
 * שמזיזה את המיקוד ל-story אחר. לכן הפקד יוצר כותרת ריקה, וה-tooltip אומר
 * במפורש שהכניסה אליה נעשית בלחיצה כפולה על אזור הכותרת — כמו ב-Word. ההנמקה
 * המלאה ב-engine/header-footer.ts.
 *
 * שלושת המתגים מציגים את המצב האמיתי של המסמך ולא מצב מקומי: `w:titlePg`
 * ו-`w:evenAndOddHeaders` נשמרים במסמך, ומתג שמחזיק דגל משלו היה מציג „כבוי”
 * על מסמך שנפתח והדגל בו דלוק.
 *
 * ## השדות, ולמה אין „עמוד X מתוך Y”
 *
 * „מספר עמוד” ו„תאריך ושעה” מכניסים שדה Word אמיתי (`{ PAGE }`,
 * `{ DATE \@ "dd/MM/yyyy HH:mm" }`)
 * ולא טקסט, ולכן הם מתעדכנים בהדפסה ובפתיחה מחדש. „עמוד X מתוך Y” אינו שדה
 * אלא רצף של טקסט ושני שדות, ואין דרך ציבורית לחשב את ההיסט שבין החלקים —
 * ולכן התפריט מציע את שני השדות בנפרד. ההנמקה המלאה, כולל מדידת מתגי הפורמט
 * של התאריך במנוע האמיתי, ב-engine/fields.ts.
 *
 * ## הסימנייה, ומה שאין לצידה
 *
 * „סימנייה” יושבת בקבוצה „קישורים”, כמו ב-Word העברי. השכנה שלה שם — „הפניה
 * מקושרת” — **אינה כאן**, ולא מפני שנשכחה: `crossRefs.insert` כותב קוד שדה
 * שאינו קוד Word, ושגם המנוע עצמו מחזיר עליו טקסט ריק אחרי `rebuild`. נמדד,
 * וההנמקה המלאה ב-engine/cross-refs.ts.
 *
 * שם סימנייה בעברית עובד במלואו (נמדד), אבל המנוע אינו אוכף על השם דבר מלבד
 * „אינו ריק” — הוא יכתוב למסמך שם עם רווחים שהוא פסול ב-Word. הוולידציה
 * שהדיאלוג מציג באה לכן מ-`normalizeBookmarkName` ולא מהמנוע. היא נשענת על
 * הכללים של Word ומחמירה עליהם בנקודה אחת במודע — ההנמקה ב-engine/bookmarks.ts.
 *
 * ## המצב הקודם
 *
 * שלושת הפקדים כאן לא עשו כלום: `imageCmd.run()` ו-`linkCmd.run()` נקראו **בלי
 * payload**, וכל אחת מהן נכשלת סגור במנוע (`create.image` דורש `src`,
 * `executeLinkCommand` דורש `href`), ו-`insertPageBreak` הייתה פונקציה ריקה
 * שהציגה `Ctrl+Enter` כקיצור שלה. אף אחד מהם לא היה מנוטרל ואף אחד לא דיווח
 * כשל — כלומר שלושה כפתורים שנראים עובדים.
 */
import { computed, inject, onUnmounted, ref, shallowRef, watch } from 'vue';
import type { SuperDoc } from 'superdoc';
import RibbonGroup from '../common/RibbonGroup.vue';
import RibbonStack from '../common/RibbonStack.vue';
import RibbonButton from '../common/RibbonButton.vue';
import RibbonMenuButton from '../common/RibbonMenuButton.vue';
import TablePicker from '../common/TablePicker.vue';
import BookmarkDialog from '../../panels/BookmarkDialog.vue';
import { useCommand } from '../../../composables/useCommand';
import { COMMAND_REPORTER, DOCUMENT_GENERATION, type CommandReporter } from '../../../composables/keys';
import { removeHyperlink } from '../../../engine/hyperlinks-manage';
import type { CommandOutcome } from '../../../engine/command-adapter';
import { ACTIVE_SUPERDOC } from '../../../engine/document-api';
import { collectTableNodeIds, placeCursorAfterInsertedTable } from '../../../engine/table-insert';
import { layOutTocRows } from '../../../engine/toc';
import {
  readPageBreakSupport,
  readPageBreakNodeId,
  setParagraphPageBreak,
  pageBreakTracker,
  type PageBreakSupport,
} from '../../../engine/page-break';
import {
  readDocCapabilities,
  type DocCapabilityQuestion,
  type DocCapabilityReport,
} from '../../../engine/doc-capabilities';
import {
  emptyHeaderFooterState,
  ensureHeaderFooter,
  readHeaderFooterState,
  removeHeaderFooter,
  setDifferentFirstPage,
  setDifferentOddEvenPages,
  setLinkedToPrevious,
  type HeaderFooterKind,
  type HeaderFooterState,
} from '../../../engine/header-footer';
import {
  emptyFieldsState,
  insertDate,
  insertDateOnly,
  insertPageCount,
  insertPageNumber,
  insertTimeOnly,
  readFieldsState,
  rebuildAllFields,
  type FieldsState,
} from '../../../engine/fields';
import {
  DATE_CALENDAR_SETTING,
  DEFAULT_DATE_CALENDAR,
  formatHebrewDate,
  insertHebrewDate,
  parseDateCalendar,
  type DateCalendar,
} from '../../../engine/hebrew-date';
import { loadSetting, saveSetting } from '../../../host/settings';
import {
  emptyBookmarksState,
  insertBookmark,
  readBookmarks,
  removeBookmark,
  renameBookmark,
  type BookmarksState,
} from '../../../engine/bookmarks';
import { imagePayload } from '../../../engine/payloads';
import { pickImageFile, readImageAsDataUrl } from '../../../host/files';

/**
 * „קישור” נפתח במעטפת ולא כאן. הדיאלוג עבר לשם כדי ש-Ctrl+K יוכל לפתוח אותו
 * מכל לשונית — לשונית שאינה פעילה אינה מורכבת, ודיאלוג שחי בתוכה אינו קיים.
 */
defineEmits<{
  (e: 'open-link'): void;
}>();

/** ברירת המחדל כשאין מדווח — הרכבה חלקית בבדיקות. זהה להתנהגות של `useCommand`. */
const fallbackReporter: CommandReporter = (outcome, id) => {
  if (!outcome.ok) console.warn(`[otzaria-word] ${id}: ${outcome.message}`);
};

const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null));
const report = inject(COMMAND_REPORTER, fallbackReporter);
/**
 * „מסמך אחר” לפי `sessions/editor-swap.ts`, לא לפי זהות `host` — ראו ההסבר
 * המלא ב-composables/keys.ts וב„„QA עצמאי” בהערת הפתיחה של engine/page-break.ts.
 */
const documentGeneration = inject(DOCUMENT_GENERATION, shallowRef(0));

const tableCmd = useCommand('table-insert');
const imageCmd = useCommand('image');
const linkCmd = useCommand('link');
const tocCmd = useCommand('table-of-contents-insert');

/**
 * בורר הקבצים והורדת הבייטים אינם מיידיים. בלי הדגל הזה לחיצה שנייה בזמן
 * שהדיאלוג פתוח פותחת דיאלוג שני, ושתי התמונות נכנסות לאותו מקום.
 */
const imageBusy = ref(false);

const imageTooltip = computed(() =>
  imageBusy.value ? 'התמונה נוספת למסמך…' : 'הוספת תמונה מקובץ (PNG או JPEG)'
);

/**
 * המנוע מניח את הסמן בתא הראשון של הטבלה החדשה — נמדד. מיקום הסמן אחרי
 * הטבלה נעשה מבחוץ, בהפרש טבלאות לפני/אחרי ההכנסה: ראו ההסבר המלא
 * ב-engine/table-insert.ts למה זה לא נקרא מתוצאת הפקודה עצמה.
 */
async function onInsertTable(dimensions: { rows: number; cols: number }): Promise<void> {
  const before = await collectTableNodeIds(superdoc.value);
  const outcome = await tableCmd.run({ rows: dimensions.rows, cols: dimensions.cols });
  if (outcome.ok) void placeCursorAfterInsertedTable(superdoc.value, before);
}

/**
 * אותם שני צעדים כמו בלשונית „הפניות”: פקודת ההכנסה של המנוע, ואחריה מיקום
 * מספרי העמודים — השורות שהמנוע כותב נושאות תו טאב בלי עצירה, ולכן מספר
 * העמוד יוצא צמוד לכותרת ובלי נקודות. ההנמקה ב-engine/toc.ts.
 *
 * הפקד הזה הוא אותה פקודת registry שבלשונית „הפניות”, ולכן הוא חייב להשאיר
 * את אותו מסמך: תוכן עניינים שנראה אחרת לפי הכפתור שנלחץ הוא באג.
 */
async function onInsertToc(): Promise<void> {
  const inserted = await tocCmd.run();
  if (inserted.ok) report(await layOutTocRows(superdoc.value), 'toc-align');
}

async function onInsertImage(): Promise<void> {
  if (imageBusy.value) return;
  imageBusy.value = true;
  try {
    const file = await pickImage();
    // ביטול. לא כשל, ואין להציג עליו שום דבר.
    if (!file) return;

    const bytes = await readImageAsDataUrl(file);
    if (!bytes.ok) {
      report({ ok: false, message: bytes.message, reason: bytes.reason }, 'image');
      return;
    }

    // `alt` הוא שם הקובץ: זה המידע היחיד שיש, והמנוע כותב אותו גם ל-
    // `wp:docPr/@name` וגם ל-`@descr` (נמדד). טוב יותר מ-descr ריק, ופחות
    // טוב מתיאור שהמשתמש כתב — עריכת ה-alt דורשת `images.setAltText`, פקד
    // שאינו בלשונית הזאת.
    const payload = imagePayload({ src: bytes.dataUrl, alt: file.name });
    if (!payload) {
      // הגנה על החוזה, לא מצב שאפשר להגיע אליו: `readImageAsDataUrl` בונה
      // את ה-src. אם הצורה תשתנה, עדיף שהמשתמש יקבל הודעה מאשר שהמנוע
      // ידחה בשקט.
      report(
        { ok: false, message: 'לא ניתן להטמיע את התמונה הזאת במסמך', reason: 'invalid-src' },
        'image',
      );
      return;
    }

    // `run` מדווח בעצמו על כשל של הפקודה — אין להוסיף כאן דיווח שני.
    await imageCmd.run(payload);
  } finally {
    imageBusy.value = false;
  }
}

/** `pickUserFile` זורק על שגיאת Host; הזריקה הופכת להודעה ולא למסך לבן. */
async function pickImage(): ReturnType<typeof pickImageFile> {
  try {
    return await pickImageFile();
  } catch (error) {
    report(
      {
        ok: false,
        message: error instanceof Error ? error.message : 'בחירת התמונה נכשלה',
        reason: 'threw',
      },
      'image',
    );
    return null;
  }
}

/* ------------------------------------------------------------------ */
/* קישור                                                              */
/* ------------------------------------------------------------------ */

/*
 * הדיאלוג עצמו אינו כאן: הוא עבר למעטפת (`composables/use-link-dialog.ts`)
 * כדי ש-`Ctrl+K` יפתח אותו מכל לשונית — לשונית שאינה פעילה אינה מורכבת.
 * „הסר קישור” נשאר, מפני שהוא פעולה על הבחירה ואינו דורש דיאלוג.
 */

/**
 * „הסר קישור" (גל 22) — `hyperlinks.remove` על הטווח המסומן
 * (engine/hyperlinks-manage.ts). אין לו פקודה בקטלוג, ולכן Document API.
 */
const removeLinkInFlight = shallowRef(false);

async function onRemoveHyperlink(): Promise<void> {
  if (removeLinkInFlight.value) return;
  removeLinkInFlight.value = true;
  try {
    report(await removeHyperlink(superdoc.value), 'hyperlink-remove');
  } finally {
    removeLinkInFlight.value = false;
  }
}
/* התחלה בעמוד חדש                                                     */
/* ------------------------------------------------------------------ */

/** עד שהיכולת נקראה — מנוטרל. „אולי כן” הוא בדיוק הכפתור המת. */
const pageBreak = shallowRef<PageBreakSupport>({ available: false, explanation: 'המסמך עדיין נטען' });

/**
 * מונה דורות: קריאת היכולת א-סינכרונית, ומסמך שנפתח אחרי מסמך אחר עשוי להשיב
 * לפניו. בלי המונה התשובה של המסמך הקודם הייתה דורסת את זו של הנוכחי.
 */
let pageBreakGeneration = 0;

/**
 * האם ידוע לנו שהפסקה שהסמן בה **כרגע** מסומנת — מ-`pageBreakTracker`, לא
 * מהמסמך. ההסבר המלא למה זה כך, ולא קריאה אמיתית מ-`doc.get()`, נמצא
 * בהערת הפתיחה של engine/page-break.ts תחת „„אינו מתג””: `doc.get()` אינו
 * מחזיר `pageBreakBefore` בתכונות הפסקה (נמדד), ואין API אחר לקרוא אותו.
 */
const pageBreakOn = shallowRef(false);

/** מונה דורות נפרד לקריאת המצב: היא רצה הרבה יותר (כל תזוזת סמן), לא רק בפתיחת מסמך. */
let pageBreakStateGeneration = 0;

/** מזהה הפסקה שהסמן בה מול `pageBreakTracker`. לא זורקת — `readPageBreakNodeId` עצמה אינה. */
async function refreshPageBreakOn(): Promise<void> {
  const mine = ++pageBreakStateGeneration;
  const nodeId = await readPageBreakNodeId(superdoc.value);
  if (mine !== pageBreakStateGeneration) return;
  pageBreakOn.value = pageBreakTracker.isOn(nodeId);
}

/**
 * השהיית הקריאה אחרי תזוזת סמן — אותה תבנית ואותו ערך כמו
 * `RULER_SELECTION_DEBOUNCE_MS` (engine/page-ruler.ts): תזוזה של סמן שמעדכנת
 * חיווי חצי שנייה אחריה נראית תקועה.
 */
const PAGE_BREAK_SELECTION_DEBOUNCE_MS = 150;
let pageBreakSelectionTimer: ReturnType<typeof setTimeout> | undefined;

function schedulePageBreakRefresh(): void {
  clearTimeout(pageBreakSelectionTimer);
  pageBreakSelectionTimer = setTimeout(() => void refreshPageBreakOn(), PAGE_BREAK_SELECTION_DEBOUNCE_MS);
}

/**
 * מנוי החיים על תזוזת סמן, כדי שהחיווי יתעדכן כשהסמן עובר לפסקה אחרת בלי
 * שהמסמך עצמו השתנה. `ui.selection` הוא אותו handle ש-App.vue מאזין לו
 * (`editor.ui.selection.observe`, לסרגל ולשורת המצב) — כאן הוא מחובר ישירות
 * מהלשונית ולא ממעטפת מרכזית, כי זה הצרכן היחיד שצריך אותו. `observe` יורה
 * מיד עם snapshot נוכחי ואז על כל שינוי (SnapshotSubscribable) — ולכן אין
 * צורך בקריאה נפרדת ל„עכשיו” מלבד הקריאה המפורשת למטה, שמכסה גם מנוע ישן
 * בלי `observe` בכלל.
 */
let unsubscribePageBreakSelection: (() => void) | null = null;

watch(
  superdoc,
  async (host) => {
    const mine = ++pageBreakGeneration;
    pageBreak.value = { available: false, explanation: 'המסמך עדיין נטען' };
    const support = await readPageBreakSupport(host);
    if (mine === pageBreakGeneration) pageBreak.value = support;
  },
  { immediate: true }
);

watch(
  superdoc,
  (host) => {
    // מסמך אחר: מאפסת את הידע ואת החיווי, ומחליפה מנוי. אותו מסמך בדיוק
    // (לשונית שהוחלפה וחזרה) — `syncDocument` אינה מאפסת, וההסבר המלא נמצא
    // לצידה ב-page-break.ts. המפתח הוא `documentGeneration.value` ולא `host`
    // עצמו — ראו „QA עצמאי” בהערת הפתיחה של page-break.ts למה.
    pageBreakTracker.syncDocument(documentGeneration.value);
    pageBreakOn.value = false;

    unsubscribePageBreakSelection?.();
    unsubscribePageBreakSelection = null;

    const observeSelection = (host as SuperDoc | null | undefined)?.ui?.selection?.observe;
    if (typeof observeSelection === 'function') {
      try {
        unsubscribePageBreakSelection =
          observeSelection.call((host as SuperDoc).ui.selection, schedulePageBreakRefresh) ?? null;
      } catch {
        unsubscribePageBreakSelection = null;
      }
    }
    // נפילה בטוחה למנוע בלי `observe` בכלל, וגם קריאה ראשונה שאינה תלויה
    // ב„ירה מיד” של המנוע.
    schedulePageBreakRefresh();
  },
  { immediate: true }
);

/**
 * מנוי על שינויים ב-`pageBreakTracker` עצמו — לא רק על תזוזת סמן. חובה בשביל
 * `forgetAll()` (App.vue, אחרי Undo/Redo מוצלח): היא נקראת **מחוץ** לרכיב
 * הזה, ואינה מזיזה את הסמן — כלומר `ui.selection.observe` שמעל לא יורה, וה-
 * חיווי היה נשאר תקוע על „פעיל” עד לחיצה הבאה. נמדד ב-QA בדפדפן אמיתי לפני
 * שהמנוי הזה נוסף. נרשם פעם אחת לכל חיי הרכיב, לא בתוך ה-watch שמעל.
 */
const unsubscribePageBreakTracker = pageBreakTracker.onChange(() => void refreshPageBreakOn());

onUnmounted(() => {
  unsubscribePageBreakTracker();
  unsubscribePageBreakSelection?.();
  clearTimeout(pageBreakSelectionTimer);
});

/**
 * ה-tooltip אומר בדיוק מה יקרה, ולא מבטיח את ההתנהגות של Word. פקד מנוטרל
 * מסביר **למה** הוא מנוטרל.
 */
const pageBreakTooltip = computed(() => {
  if (!pageBreak.value.available) return pageBreak.value.explanation;
  return pageBreakOn.value
    ? 'הפסקה הזאת כבר מתחילה בעמוד חדש. לחיצה תבטל זאת'
    : 'הפסקה שבה הסמן תתחיל בראש עמוד חדש';
});

/**
 * מתג אמיתי: לחיצה כשהפסקה „כבויה” שולחת `true`, וכשהיא „דלוקה” שולחת
 * `false`. `pageBreakOn` הוא הידע שלנו על הפסקה **הנוכחית** בלבד (ראו
 * ההסבר למעלה), ולכן נקרא מיד לפני השליחה ולא נשמר ממערכת נפרדת.
 */
async function onStartOnNewPage(): Promise<void> {
  const outcome = await setParagraphPageBreak(superdoc.value, !pageBreakOn.value, pageBreakTracker);
  report(outcome, 'page-break-before');
  await refreshPageBreakOn();
}

/* ------------------------------------------------------------------ */
/* כותרת עליונה ותחתונה                                                */
/* ------------------------------------------------------------------ */

const capabilities = shallowRef<DocCapabilityReport | null>(null);
const headerFooter = shallowRef<HeaderFooterState>(emptyHeaderFooterState());
/**
 * מספר השדות במסמך. נקרא באותו דור כמו מצב הכותרות ומאותה סיבה — הוא נקרא
 * מהמסמך ולא מוחזק כדגל מקומי, כדי שמסמך שנפתח וכבר יש בו שדות לא יציג „אין
 * מה לעדכן”.
 */
const fieldsState = shallowRef<FieldsState>(emptyFieldsState());

/**
 * מונה דורות משותף לשתי הקריאות: שתיהן א-סינכרוניות, ומסמך שנפתח אחרי מסמך
 * אחר עשוי להשיב לפניו. בלי המונה התשובה של המסמך הקודם הייתה דורסת את זו של
 * הנוכחי. אותה תבנית כמו `pageBreakGeneration` שמעל.
 */
let headerFooterGeneration = 0;

watch(
  superdoc,
  async (host) => {
    const mine = ++headerFooterGeneration;
    capabilities.value = null;
    headerFooter.value = emptyHeaderFooterState();
    // בלי האיפוס הזה מונה השדות של המסמך הקודם שורד עד שהקריאה חוזרת, וה-
    // tooltip של „עדכן שדות” מבטיח חישוב מחדש במסמך שאין בו שדה אחד.
    fieldsState.value = emptyFieldsState();

    const [capabilityReport, state, fields] = await Promise.all([
      readDocCapabilities(host),
      readHeaderFooterState(host),
      readFieldsState(host),
    ]);
    if (mine !== headerFooterGeneration) return;
    capabilities.value = capabilityReport;
    headerFooter.value = state;
    fieldsState.value = fields;
  },
  { immediate: true }
);

/** עד שהיכולות נקראו — הכול מנוטרל. „אולי כן” הוא בדיוק הכפתור המת. */
function can(question: DocCapabilityQuestion): boolean {
  return capabilities.value?.can(question) ?? false;
}

function tip(question: DocCapabilityQuestion, enabledText: string): string {
  const explanation = capabilities.value?.explain(question);
  return can(question) ? enabledText : explanation || 'המסמך עדיין נטען';
}

const headerItems = [
  { id: 'edit', label: 'עריכת כותרת עליונה', hint: 'יוצר כותרת ריקה אם עדיין אין' },
  { id: 'remove', label: 'הסרת כותרת עליונה', hint: 'מוחק את הכותרת מכל המסמך' },
];

const footerItems = [
  { id: 'edit', label: 'עריכת כותרת תחתונה', hint: 'יוצר כותרת ריקה אם עדיין אין' },
  { id: 'remove', label: 'הסרת כותרת תחתונה', hint: 'מוחק את הכותרת מכל המסמך' },
];

/**
 * ה-tooltip אומר מה באמת יקרה: הכותרת נוצרת, והכניסה אליה היא לחיצה כפולה.
 * הבטחה ש„הסמן יעבור לכותרת” הייתה תיאור של פעולה שאין לה API.
 */
const headerTooltip = computed(() =>
  tip(
    'canEditHeaderFooter',
    headerFooter.value.hasHeader
      ? 'למסמך יש כותרת עליונה. לחיצה כפולה על אזור הכותרת פותחת אותה לעריכה'
      : 'יצירת כותרת עליונה ריקה. לעריכתה — לחיצה כפולה על אזור הכותרת'
  )
);

const footerTooltip = computed(() =>
  tip(
    'canEditHeaderFooter',
    headerFooter.value.hasFooter
      ? 'למסמך יש כותרת תחתונה. לחיצה כפולה על אזור הכותרת פותחת אותה לעריכה'
      : 'יצירת כותרת תחתונה ריקה. לעריכתה — לחיצה כפולה על אזור הכותרת'
  )
);

/**
 * „קשר לקודם” נוגע רק במקטעים שאחרי הראשון, ולמסמך בעל מקטע יחיד אין קודם —
 * ולכן הפקד מנוטרל שם, כמו ב-Word. הוא אינו פקד מת: ה-API קיים והוא נדלק
 * ברגע שיש במסמך מקטע שני. ה-tooltip מסביר את זה במקום להשאיר את המשתמש
 * מול לחיצה שנכשלת בהכרח.
 */
const linkToPreviousTooltip = computed(() =>
  tip(
    'canLinkToPrevious',
    headerFooter.value.sectionCount > 1
      ? 'הכותרות של המקטעים הבאים יהיו זהות לאלה של המקטע שלפניהם'
      : 'אין במסמך מקטע נוסף — הקישור נוגע רק במקטעים שאחרי הראשון'
  )
);

/**
 * מריצה פעולה, מדווחת עליה, וקוראת מחדש את המצב. הקריאה מחדש נדרשת גם כשהפעולה
 * נכשלה: מוטציה שנכשלה באמצע רצף מקטעים משאירה מסמך שאינו במצב שהמתג מציג.
 */
async function runHeaderFooter(id: string, action: Promise<CommandOutcome>): Promise<void> {
  report(await action, id);
  // **קורא** את המונה ואינו מקדם אותו: קידום כאן היה מבטל את הקריאה של ה-watch
  // אם מסמך אחר נטען בזמן שהפעולה רצה, ו-`capabilities` היה נשאר null — כלומר
  // כל הקבוצה מנוטרלת לצמיתות עם „המסמך עדיין נטען”. הקריאה בלבד עדיין זורקת
  // תשובה שהגיעה אחרי שהמסמך התחלף.
  const mine = headerFooterGeneration;
  const state = await readHeaderFooterState(superdoc.value);
  if (mine === headerFooterGeneration) headerFooter.value = state;
}

function onHeaderFooterAction(kind: HeaderFooterKind, action: string): void {
  const host = superdoc.value;
  void runHeaderFooter(
    `header-footer-${kind}`,
    action === 'remove' ? removeHeaderFooter(host, kind) : ensureHeaderFooter(host, kind)
  );
}

function onHeaderAction(action: string): void {
  onHeaderFooterAction('header', action);
}

function onFooterAction(action: string): void {
  onHeaderFooterAction('footer', action);
}

function onToggleTitlePage(): void {
  void runHeaderFooter(
    'header-footer-title-page',
    setDifferentFirstPage(superdoc.value, !headerFooter.value.titlePage)
  );
}

function onToggleOddEven(): void {
  void runHeaderFooter(
    'header-footer-odd-even',
    setDifferentOddEvenPages(superdoc.value, !headerFooter.value.oddEven)
  );
}

function onToggleLinkToPrevious(): void {
  void runHeaderFooter(
    'header-footer-link-to-previous',
    setLinkedToPrevious(superdoc.value, !headerFooter.value.linkedToPrevious)
  );
}

/* ------------------------------------------------------------------ */
/* שדות                                                                */
/* ------------------------------------------------------------------ */

/**
 * שני פריטים ולא „עמוד X מתוך Y” אחד: הצירוף הזה דורש טקסט **בין** שני שדות,
 * ואין דרך ציבורית לחשב את ההיסט שאחרי השדה הראשון. ה-hint אומר את זה
 * למשתמש במקום להשאיר אותו לחפש פריט שאינו קיים.
 */
const pageNumberItems = [
  { id: 'page', label: 'מספר עמוד', hint: 'מתעדכן לפי העמוד שהשדה נמצא בו' },
  {
    id: 'count',
    label: 'מספר העמודים במסמך',
    hint: 'לצירוף „עמוד X מתוך Y” יש להקליד את המילים ולהוסיף את שני השדות',
  },
];

/**
 * הלוח של „תאריך ושעה”: עברי כברירת מחדל, לועזי בבחירה מהתפריט. העדפה של
 * התוכנה ולא של המסמך, ולכן נזכרת בין הפעלות (host/settings.ts).
 *
 * בלוח העברי התאריך נכתב כטקסט קבוע, בלועזי כשדה שמתעדכן — ראו
 * engine/hebrew-date.ts.
 */
const dateCalendar = ref<DateCalendar>(DEFAULT_DATE_CALENDAR);
void loadSetting(DATE_CALENDAR_SETTING, DEFAULT_DATE_CALENDAR, parseDateCalendar).then((saved) => {
  dateCalendar.value = saved;
});

function setDateCalendar(calendar: DateCalendar): void {
  dateCalendar.value = calendar;
  void saveSetting(DATE_CALENDAR_SETTING, calendar);
}

/** „היום” של הדוגמאות בתפריט — מתרענן בכל פתיחה, כדי שתפריט שנפתח אחרי חצות לא יציג את אתמול. */
const dateNow = ref(new Date());

function gregorianExample(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}`;
}

/**
 * ה-tooltip אומר את מה שייכתב: בלועזי — שדה שמתעדכן, ובפורמט המפורש
 * `dd/MM/yyyy HH:mm` (המתג נשלח בקוד השדה, ראו engine/fields.ts); בעברי —
 * טקסט שאינו מתעדכן, כי זה ההבדל שהמשתמש צריך לדעת.
 */
const dateTooltip = computed(() =>
  dateCalendar.value === 'hebrew'
    ? tip('canInsertField', 'הכנסת התאריך העברי והשעה במקום הסמן, כטקסט שאינו מתעדכן')
    : tip('canInsertField', 'הכנסת שדה תאריך ושעה שמתעדכן, בפורמט יום/חודש/שנה שעה:דקות')
);

/**
 * תפריט החץ של „תאריך ושעה”: שלוש ברירות תוכן, ומתחת לקו — הלוח. הלוח הוא
 * בחירה מתמשכת ולא פעולה, ולכן פריטי רדיו (`checked`) עם דוגמה של היום.
 */
const dateItems = computed(() => {
  const hebrew = dateCalendar.value === 'hebrew';
  return [
    {
      id: 'date-time',
      label: 'תאריך ושעה',
      hint: hebrew ? 'תאריך עברי ושעה, כטקסט קבוע' : 'יום/חודש/שנה ושעה',
    },
    {
      id: 'date-only',
      label: 'תאריך בלבד',
      hint: hebrew ? 'תאריך עברי, בלי שעה' : 'יום/חודש/שנה, בלי שעה',
    },
    { id: 'time-only', label: 'שעה בלבד', hint: 'שעה בלבד, בלי תאריך' },
    {
      id: 'calendar-hebrew',
      label: 'לוח עברי',
      hint: formatHebrewDate(dateNow.value),
      checked: hebrew,
      separatorBefore: true,
    },
    {
      id: 'calendar-gregorian',
      label: 'לוח לועזי',
      hint: gregorianExample(dateNow.value),
      checked: !hebrew,
    },
  ];
});

const rebuildTooltip = computed(() =>
  tip(
    'canRebuildFields',
    fieldsState.value.count > 0
      ? 'חישוב מחדש של כל השדות במסמך, כמו F9 ב-Word'
      : 'אין במסמך שדות לעדכן'
  )
);

/**
 * מריצה פעולת שדה, מדווחת עליה, וקוראת מחדש את המונה. הקריאה מחדש נדרשת גם
 * בכשל, מאותו טעם כמו ב-`runHeaderFooter`: מוטציה שנכשלה באמצע משאירה מסמך
 * שאינו במצב שה-tooltip מתאר.
 */
async function runFields(id: string, action: Promise<CommandOutcome>): Promise<void> {
  report(await action, id);
  // **קורא** את המונה ואינו מקדם אותו — ראו ההסבר ב-`runHeaderFooter`.
  const mine = headerFooterGeneration;
  const fields = await readFieldsState(superdoc.value);
  if (mine === headerFooterGeneration) fieldsState.value = fields;
}

function onPageNumberAction(action: string): void {
  const host = superdoc.value;
  void runFields(
    `field-${action === 'count' ? 'numpages' : 'page'}`,
    action === 'count' ? insertPageCount(host) : insertPageNumber(host)
  );
}

function onInsertDate(): void {
  if (dateCalendar.value === 'hebrew') {
    void insertHebrewDate(superdoc.value, true).then((outcome) => report(outcome, 'date-hebrew'));
    return;
  }
  void runFields('field-date', insertDate(superdoc.value));
}

/**
 * בחירה בתפריט החץ. „תאריך ושעה” מזוהה עם לחיצת הגוף — אותה פעולה בדיוק.
 * „שעה בלבד” אינה תלויה בלוח ונשארת שדה בשניהם.
 */
function onDateAction(action: string): void {
  if (action === 'calendar-hebrew') return setDateCalendar('hebrew');
  if (action === 'calendar-gregorian') return setDateCalendar('gregorian');
  if (action === 'date-time') return onInsertDate();
  const host = superdoc.value;
  if (action === 'date-only' && dateCalendar.value === 'hebrew') {
    void insertHebrewDate(host, false).then((outcome) => report(outcome, 'date-hebrew-only'));
    return;
  }
  void runFields(
    action === 'date-only' ? 'field-date-only' : 'field-time-only',
    action === 'date-only' ? insertDateOnly(host) : insertTimeOnly(host)
  );
}

function onRebuildFields(): void {
  void runFields('fields-rebuild', rebuildAllFields(superdoc.value));
}

/* ------------------------------------------------------------------ */
/* סימניות                                                             */
/* ------------------------------------------------------------------ */

const bookmarkDialogOpen = ref(false);

/**
 * שמות הסימניות שבמסמך. נקראים כשהדיאלוג נפתח ואחרי כל פעולה, ולא מוחזקים
 * כרשימה מקומית: הרשימה הזאת היא **המסמך**, ורשימה שנבנית מהפעולות שהמשתמש
 * עשה בסשן הזה הייתה מציגה מסמך שנפתח עם סימניות כאילו אין בו אף אחת.
 */
const bookmarks = shallowRef<BookmarksState>(emptyBookmarksState());

const bookmarkTooltip = computed(() =>
  tip('canManageBookmarks', 'סימון הפסקה שבה הסמן בשם, לניווט ולהפניות מתוך Word')
);

/**
 * קוראת מחדש את הסימניות. משתמשת במונה הדורות של הקבוצה שמעל ומאותו טעם:
 * **קוראת** אותו ואינה מקדמת אותו, כדי לא לבטל את הקריאה של ה-watch. ההסבר
 * המלא ב-`runHeaderFooter`.
 */
async function refreshBookmarks(): Promise<void> {
  const mine = headerFooterGeneration;
  const state = await readBookmarks(superdoc.value);
  if (mine === headerFooterGeneration) bookmarks.value = state;
}

/**
 * הרשימה נקראת בלחיצה ולא מוחזקת עדכנית ברקע: דיאלוג הסימניות נפתח לעיתים
 * רחוקות, ומנוי שסורק את כל הסימניות בכל הקלדה במסמך הוא מחיר על מידע שאיש
 * אינו רואה.
 */
async function onOpenBookmarkDialog(): Promise<void> {
  await refreshBookmarks();
  bookmarkDialogOpen.value = true;
}

/**
 * מריצה פעולת סימנייה, מדווחת עליה, וקוראת מחדש את הרשימה — גם בכשל. הדיאלוג
 * נשאר פתוח: הוא מציג את הרשימה המעודכנת, וזה בדיוק מה שהמשתמש צריך לראות
 * אחרי הפעולה.
 */
async function runBookmark(id: string, action: Promise<CommandOutcome>): Promise<void> {
  report(await action, id);
  await refreshBookmarks();
}

function onAddBookmark(name: string): void {
  void runBookmark('bookmark-insert', insertBookmark(superdoc.value, name));
}

function onRemoveBookmark(name: string): void {
  void runBookmark('bookmark-remove', removeBookmark(superdoc.value, name));
}

function onRenameBookmark(change: { from: string; to: string }): void {
  void runBookmark('bookmark-rename', renameBookmark(superdoc.value, change.from, change.to));
}
</script>

<style scoped>
.ribbon-tab-pane {
  display: flex;
  align-items: stretch;
  gap: 0;
  height: 100%;
}
</style>
