<template>
  <div class="ribbon-tab-pane home-tab">
    <!-- קבוצה 1: לוח -->
    <RibbonGroup
      title="לוח"
      icon="paste"
    >
      <!--
        מפוצל, כמו ב-Word: הגוף מדביק לפי ברירת המחדל של Ctrl+V, והחץ פותח את
        „אפשרויות הדבקה” ואת בחירת ברירת המחדל עצמה.
      -->
      <RibbonMenuButton
        split
        icon="paste"
        label="הדבק"
        variant="large"
        :tooltip="pasteTooltip"
        :description="pasteDescription"
        shortcut-id="paste"
        menu-tooltip="אפשרויות הדבקה"
        menu-description="שמירה על עיצוב המקור, מיזוג עיצוב או טקסט בלבד — ובחירת ההתנהגות של Ctrl+V"
        :disabled="!canPaste"
        :items="pasteMenuItems"
        @action="doPaste(pasteDefaultMode)"
        @select="onPasteMenuSelect"
      />
      <RibbonStack>
        <RibbonButton
          icon="cut"
          label="גזור"
          variant="small"
          :tooltip="cutTooltip"
          shortcut-id="cut"
          :disabled="!canCut"
          @click="doCut"
        />
        <RibbonButton
          icon="copy"
          label="העתק"
          variant="small"
          :tooltip="copyTooltip"
          shortcut-id="copy"
          :disabled="!canCopy"
          @click="doCopy"
        />
        <RibbonButton
          icon="formatPainter"
          label="מברשת עיצוב"
          variant="small"
          tooltip="העתק עיצוב ממקום אחד והחל במקום אחר"
          shortcut-id="format-painter"
          :active="formatPainterCmd.active.value"
          :disabled="!formatPainterCmd.enabled.value"
          @click="formatPainterCmd.run()"
        />
      </RibbonStack>
    </RibbonGroup>

    <!-- קבוצה 2: גופן -->
    <RibbonGroup
      title="גופן"
      icon="bold"
      :column-flow="true"
    >
      <!-- שורה עליונה: גופן, גודל, הגדל/הקטן, נקה -->
      <div class="word-group-row">
        <!--
          `:model-value` ולא `v-model`: הערך המוצג הוא computed שמקורו במנוע,
          ואין לפקד רשות לכתוב אליו. מה שהמשתמש בוחר עובר בפקודה, וחוזר משם.
        -->
        <!--
          RibbonCombo ולא RibbonSelect: מרגע שהמכונה נמנית הרשימה מונה מאות
          שמות, ובבורר נייטיב אין חיפוש. ראו common/RibbonCombo.vue.
        -->
        <RibbonCombo
          :model-value="fonts.family.value"
          :options="fonts.familyOptions.value"
          :disabled="!fonts.familyEnabled.value"
          :sample="fonts.sampleText.value"
          :sample-size="fonts.sampleSize.value"
          width="130px"
          title="גופן"
          @update:model-value="fonts.setFamily"
          @preview="fonts.hoverFamily"
          @preview-end="fonts.endHoverFamily"
        />
        <!--
          גם הגודל הוא RibbonCombo, ולא `<select>` שהיה כאן: בסולם של Word יש
          ארבעה עשר גדלים, והמסמכים שנערכים כאן משתמשים גם ב-13, ב-15 וב-10.5.
          בבורר סגור הם היו נגישים רק דרך „הגדל/הקטן” — כלומר לא נגישים.
          `normalize` הוא מה שהופך את התיבה לתיבת ערך: ראו common/RibbonCombo.vue.
        -->
        <RibbonCombo
          :model-value="fonts.size.value"
          :options="fonts.sizeOptions.value"
          :disabled="!fonts.sizeEnabled.value"
          :normalize="fonts.normalizeSize"
          empty-text="Enter מחיל את הגודל שהוקלד"
          list-min-width="52px"
          width="54px"
          title="גודל גופן"
          @update:model-value="fonts.setSize"
        />
        <RibbonButton
          icon="growFont"
          variant="icon-only"
          tooltip="הגדל גופן"
          description="מגדיל את הטקסט המסומן בדרגה אחת בכל לחיצה"
          shortcut-id="font-grow"
          :disabled="!fonts.sizeEnabled.value"
          @click="fonts.grow"
        />
        <RibbonButton
          icon="shrinkFont"
          variant="icon-only"
          tooltip="הקטן גופן"
          description="מקטין את הטקסט המסומן בדרגה אחת בכל לחיצה"
          shortcut-id="font-shrink"
          :disabled="!fonts.sizeEnabled.value"
          @click="fonts.shrink"
        />
        <RibbonButton
          icon="clearFormatting"
          variant="icon-only"
          tooltip="נקה את כל העיצוב"
          description="מחזיר את הטקסט המסומן לעיצוב הרגיל, והתוכן נשאר"
          shortcut-id="clear-formatting"
          :disabled="!clearFormatCmd.enabled.value"
          @click="clearFormatCmd.run()"
        />

        <!--
          „גופן מתקדם" — פתח הדיאלוג, כמו פתח ה„גופן" של Word בקצה הקבוצה.
          הדיאלוג עצמו נכשל סגור כשאין Document API, ולכן הכפתור נשאר לחיץ
          והפתיחה מסבירה. `@pointerdown.prevent` מונע גזילת המיקוד מהעורך —
          הבחירה חייבת לשרוד עד החלת העיצוב.
        -->
        <RibbonButton
          label="מתקדם"
          variant="small"
          description="ריווח תווים, מיקום ביחס לשורה, אפקטים וגופן מורכב"
          :disabled="fontAdvInFlight"
          @click="onOpenFontAdvanced"
        />
      </div>

      <!-- שורה תחתונה: B, I, U, S, sub, super, highlight, color -->
      <div class="word-group-row">
        <RibbonButton
          icon="bold"
          variant="icon-only"
          tooltip="מודגש"
          description="מעבה את הטקסט המסומן"
          shortcut-id="bold"
          :active="boldCmd.active.value"
          :disabled="!boldCmd.enabled.value"
          @click="boldCmd.run()"
        />
        <RibbonButton
          icon="italic"
          variant="icon-only"
          tooltip="נטוי"
          description="מטה את הטקסט המסומן"
          shortcut-id="italic"
          :active="italicCmd.active.value"
          :disabled="!italicCmd.enabled.value"
          @click="italicCmd.run()"
        />
        <RibbonButton
          icon="underline"
          variant="icon-only"
          tooltip="קו תחתון"
          description="מוסיף קו מתחת לטקסט המסומן"
          shortcut-id="underline"
          :active="underlineCmd.active.value"
          :disabled="!underlineCmd.enabled.value"
          @click="underlineCmd.run()"
        />
        <RibbonButton
          icon="strikethrough"
          variant="icon-only"
          tooltip="קו חוצה"
          description="מעביר קו באמצע הטקסט המסומן"
          shortcut-id="strikethrough"
          :active="strikeCmd.active.value"
          :disabled="!strikeCmd.enabled.value"
          @click="strikeCmd.run()"
        />
        <RibbonButton
          icon="subscript"
          variant="icon-only"
          :tooltip="vertAlignTooltip('subscript')"
          label="כתב תחתי"
          shortcut-id="subscript"
          :disabled="!canSetVertAlign"
          @click="onVertAlign('subscript')"
        />
        <RibbonButton
          icon="superscript"
          variant="icon-only"
          :tooltip="vertAlignTooltip('superscript')"
          label="כתב עליון"
          shortcut-id="superscript"
          :disabled="!canSetVertAlign"
          @click="onVertAlign('superscript')"
        />

        <div class="word-separator" />

        <ColorPickerPopover
          :model-value="highlightColor"
          icon="highlight"
          title="צבע סימון טקסט"
          default-color="#FFFF00"
          :disabled="!highlightCmd.enabled.value"
          @change="onHighlightChange"
        />
        <ColorPickerPopover
          :model-value="textColor"
          icon="fontColor"
          title="צבע גופן"
          default-color="#000000"
          :disabled="!fontColorCmd.enabled.value"
          @change="onTextColorChange"
        />
      </div>
    </RibbonGroup>

    <!-- קבוצה 3: פיסקה -->
    <RibbonGroup
      title="פיסקה"
      icon="alignRight"
      :column-flow="true"
    >
      <!-- שורה עליונה: תבליטים, מספור, הזחה, כיווניות, סימני עיצוב, תפריט פסקה -->
      <div class="word-group-row">
        <!--
          „תבליטים” ו„מספור” הם כפתורים מפוצלים: הגוף מחיל את הרשימה, והחץ
          פותח את פעולות הרשימה. עד עכשיו הן ישבו בכפתור „רשימה” נפרד לצדם —
          שלושה פקדים לדבר אחד, ומי שרצה מספור עברי היה צריך לדעת שהוא מסתתר
          מאחורי כפתור שאינו הכפתור שיצר את הרשימה.
        -->
        <RibbonMenuButton
          split
          icon="bulletList"
          label="תבליטים"
          variant="icon-only"
          tooltip="תבליטים"
          description="הופך את הפסקאות המסומנות לרשימה מסומנת בנקודות"
          menu-tooltip="פעולות תבליטים"
          :menu-description="bulletMenuHint"
          :active="bulletCmd.active.value"
          :action-disabled="!bulletCmd.enabled.value"
          :disabled="!listsAvailable"
          :items="bulletMenuItems"
          @action="bulletCmd.run()"
          @select="onListMenuSelect"
        />
        <RibbonMenuButton
          split
          icon="numberList"
          label="מספור"
          variant="icon-only"
          tooltip="מספור"
          description="הופך את הפסקאות המסומנות לרשימה ממוספרת"
          menu-tooltip="פעולות מספור"
          :menu-description="numberMenuHint"
          :active="numberedCmd.active.value"
          :action-disabled="!numberedCmd.enabled.value"
          :disabled="!listsAvailable"
          :items="numberMenuItems"
          @action="numberedCmd.run()"
          @select="onListMenuSelect"
        />
        <RibbonButton
          icon="indentDecrease"
          variant="icon-only"
          tooltip="הקטן הזחה"
          description="מקרב את הפסקה לשולי הדף"
          shortcut-id="indent-decrease"
          :disabled="!indentDecCmd.enabled.value"
          @click="indentDecCmd.run()"
        />
        <RibbonButton
          icon="indentIncrease"
          variant="icon-only"
          tooltip="הגדל הזחה"
          description="מרחיק את הפסקה משולי הדף"
          shortcut-id="indent-increase"
          :disabled="!indentIncCmd.enabled.value"
          @click="indentIncCmd.run()"
        />

        <div class="word-separator" />

        <RibbonButton
          icon="dirRtl"
          variant="icon-only"
          tooltip="כיוון פסקה מימין לשמאל"
          description="מסדר את הפסקה לקריאה בעברית: ההזחה והיישור בצד ימין"
          shortcut-id="direction-rtl"
          :active="dirRtlActive"
          :disabled="!dirRtlCmd.enabled.value"
          @click="runDirectionRtl()"
        />
        <RibbonButton
          icon="dirLtr"
          variant="icon-only"
          tooltip="כיוון פסקה משמאל לימין"
          description="מסדר את הפסקה לקריאה בלטינית: ההזחה והיישור בצד שמאל"
          shortcut-id="direction-ltr"
          :active="dirLtrActive"
          :disabled="!dirLtrCmd.enabled.value"
          @click="runDirectionLtr()"
        />
        <RibbonButton
          icon="pilcrow"
          variant="icon-only"
          tooltip="הצג/הסתר סימני עיצוב"
          description="מציג סימני פסקה, טאבים ורווחים על המסך. הם אינם מודפסים"
          shortcut-id="formatting-marks"
          :active="marksCmd.active.value"
          :disabled="!marksCmd.enabled.value"
          @click="marksCmd.run()"
        />

        <div class="word-separator" />

        <!--
          „תפריט פסקה” — פתח הדיאלוג, כמו פתח ה„פסקה” של Word בקצה הקבוצה.
          האייקון אינו `pilcrow`, ובכוונה: הוא היה זהה לזה של „הצג/הסתר
          סימני עיצוב” שלצדו באותה קבוצה, ושני כפתורים שכנים עם אותו ציור
          אינם שני כפתורים אלא אחד שנראה שבור.

          זמינותו אינה של פקודה: הדיאלוג עצמו נכשל סגור כשאין Document API,
          ולכן הכפתור נשאר לחיץ והפתיחה מסבירה. `@pointerdown.prevent` מונע
          גזילת המיקוד מהעורך — הבחירה חייבת לשרוד עד פתיחת הדיאלוג.
        -->
        <RibbonButton
          icon="paragraphOptions"
          variant="icon-only"
          tooltip="תפריט פסקה"
          description="כניסות, ריווח בין פסקאות, מרווח שורות ועצירות טאב"
          :disabled="paraInFlight"
          @click="onOpenParagraph"
        />
      </div>

      <!-- שורה תחתונה: יישור ימין, מרכז, שמאל, מלא, מרווח שורות וצעד בו -->
      <div class="word-group-row">
        <RibbonButton
          icon="alignRight"
          variant="icon-only"
          tooltip="יישור לימין"
          description="מיישר את הפסקה לשוליים הימניים"
          shortcut-id="align-right"
          :active="alignCmd.value.value === 'right'"
          :disabled="!alignCmd.enabled.value"
          @click="onAlign('right')"
        />
        <RibbonButton
          icon="alignCenter"
          variant="icon-only"
          tooltip="מרכז"
          description="ממרכז את הפסקה בין שני השוליים"
          shortcut-id="align-center"
          :active="alignCmd.value.value === 'center'"
          :disabled="!alignCmd.enabled.value"
          @click="onAlign('center')"
        />
        <RibbonButton
          icon="alignLeft"
          variant="icon-only"
          tooltip="יישור לשמאל"
          description="מיישר את הפסקה לשוליים השמאליים"
          shortcut-id="align-left"
          :active="alignCmd.value.value === 'left'"
          :disabled="!alignCmd.enabled.value"
          @click="onAlign('left')"
        />
        <RibbonButton
          icon="alignJustify"
          variant="icon-only"
          tooltip="יישור לשני הצדדים"
          description="מותח את השורות עד שני השוליים, מלבד השורה האחרונה"
          shortcut-id="align-justify"
          :active="alignCmd.value.value === 'justify'"
          :disabled="!alignCmd.enabled.value"
          @click="onAlign('justify')"
        />

        <div class="word-separator" />

        <!--
          „מרווח שורות וריווח” — התפריט שב-Word מחזיק את שני הדברים יחד: את
          המרחק בין השורות **בתוך** הפסקה, ואת הרווח שמפריד פסקה מחברתה.

          היה כאן בורר מספרים, והוא החזיק רק את הראשון; השני היה מגיע רק דרך
          דיאלוג „פסקה”, כלומר מי שחיפש „רווח בין פסקאות” לא מצא אותו. שניהם
          תחת פקד אחד הם גם פחות פקדים ברצועה, לא יותר.

          שני כפתורי הצעד שלצדו נשארים: הם מגיעים לערכים שאינם ברשימה (1.2,
          1.35) בלי לפתוח דבר.
        -->
        <RibbonMenuButton
          icon="lineSpacing"
          label="מרווח שורות וריווח"
          variant="icon-only"
          tooltip="מרווח שורות וריווח"
          :disabled="!lineSpacingCmd.enabled.value || lineSpacingInFlight"
          :items="spacingMenuItems"
          @open="onSpacingMenuOpen"
          @hover="onSpacingMenuHover"
          @select="onSpacingMenuSelect"
        >
          <template #footer>
            <SpacingPreview
              :text-width-twips="paraTextWidthTwips"
              :font-size-pt="paraFontSizePt"
              :before-twips="previewSpacing.beforeTwips"
              :after-twips="previewSpacing.afterTwips"
              :line-twips="previewSpacing.lineTwips"
              :line-rule="previewSpacing.rule"
              :rtl="paraSectionRtl"
            />
          </template>
        </RibbonMenuButton>

        <!--
          שני הכפתורים האלה הם הבורר שלצדם בצעד אחד: הרשימה מציעה שש
          אפשרויות, והם מגיעים למה שאינו בהן (1.2, 1.35) בלי דיאלוג. הם
          נכבים בקצות הסולם — ראו `MIN_LINE_HEIGHT` ב-engine/payloads.ts —
          כי כפתור שנשאר דלוק על 10.0 הוא לחיצה שאינה משנה כלום ואינה
          אומרת זאת.
        -->
        <RibbonButton
          icon="lineSpacingIncrease"
          variant="icon-only"
          tooltip="הגדל מרווח שורות"
          :description="growSpacingHint"
          :disabled="!lineSpacingCmd.enabled.value || !canGrowSpacing || lineSpacingInFlight"
          @click="onLineSpacingStep(grownLineHeight)"
        />
        <RibbonButton
          icon="lineSpacingDecrease"
          variant="icon-only"
          tooltip="הקטן מרווח שורות"
          :description="shrinkSpacingHint"
          :disabled="!lineSpacingCmd.enabled.value || !canShrinkSpacing || lineSpacingInFlight"
          @click="onLineSpacingStep(shrunkLineHeight)"
        />
      </div>
    </RibbonGroup>

    <!-- קבוצה 4: סגנונות -->
    <RibbonGroup
      title="סגנונות"
      icon="shading"
      class="styles-group"
    >
      <!-- `disabled` ולא רק `current-style`: בלי בחירה במסמך הפקודה
           `linked-style` נכשלת, וגלריה שנראית פעילה מזמינה לחיצה שלא תעשה
           כלום. הדיווח על הכשל עצמו הוא של `useCommand().run()`. -->
      <StyleGallery
        :current-style="String(styleCmd.value.value || 'Normal')"
        :disabled="!styleCmd.enabled.value"
        :manage-disabled="!styleManager.available.value || styleManager.busy.value"
        @select-style="onApplyStyle"
        @style-action="styleManager.onAction"
      />
    </RibbonGroup>

    <!-- קבוצה 5: עריכה -->
    <RibbonGroup
      title="עריכה"
      icon="search"
    >
      <RibbonStack>
        <RibbonButton
          icon="search"
          label="חפש"
          variant="small"
          tooltip="חיפוש טקסט במסמך"
          shortcut-id="find"
          @click="$emit('open-find')"
        />
        <RibbonButton
          icon="replace"
          label="החלפה"
          variant="small"
          tooltip="החלפת טקסט במסמך"
          shortcut-id="replace"
          @click="$emit('open-replace')"
        />
        <!--
          Ctrl+A אינו פקודה של ה-controller — אין במנוע `select-all` — ולכן
          הקיצור מנתב לאותו מסלול Document API שהכפתור הזה מנתב אליו
          (`selectWholeDocument`), ולא לבחירה המקורית של משטח העריכה, שאינה
          מגיעה לכל המסמך בפריסה מרובת עמודים.
        -->
        <RibbonButton
          icon="select"
          label="בחר הכל"
          shortcut-id="select-all"
          variant="small"
          :tooltip="selectAllTooltip"
          :disabled="!canSelectAll"
          @click="doSelectAll"
        />
      </RibbonStack>
    </RibbonGroup>

    <ParagraphDialog
      :is-open="paragraphOpen"
      :busy="paraInFlight"
      :tabs-enabled="tabsEnabled"
      :snapshot="paraSnapshot"
      :page-text-width-twips="paraTextWidthTwips"
      :font-size-pt="paraFontSizePt"
      :section-rtl="paraSectionRtl"
      @close="paragraphOpen = false"
      @submit="onParagraphSubmit"
      @tab-add="onParagraphTabAdd"
      @tab-remove="onParagraphTabRemove"
      @tabs-clear="onParagraphTabsClear"
    />

    <StyleDialog
      :is-open="styleManager.dialog.value !== null"
      :busy="styleManager.busy.value"
      :mode="styleManager.dialog.value?.mode ?? 'modify'"
      :style-id="styleManager.dialog.value?.styleId ?? null"
      :style-name="styleManager.dialog.value?.styleName ?? ''"
      :style-type="styleManager.dialog.value?.styleType ?? 'paragraph'"
      :styles="styleManager.dialog.value?.styles ?? []"
      :based-on="styleManager.dialog.value?.basedOn ?? null"
      :next="styleManager.dialog.value?.next ?? null"
      :format="styleManager.dialog.value?.format ?? {}"
      :error="styleManager.error.value"
      :shortcut="styleManager.dialog.value?.shortcut ?? null"
      :new-documents-option="styleManager.dialog.value?.newDocumentsOption ?? false"
      :combo-problem="(combo) => styleManager.comboProblem(styleManager.dialog.value?.styleId ?? null, combo)"
      @close="styleManager.close"
      @submit="styleManager.onSubmit"
    />

    <FontAdvancedDialog
      :is-open="fontAdvOpen"
      :busy="fontAdvInFlight"
      @close="fontAdvOpen = false"
      @submit="onFontAdvancedSubmit"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, inject, onUnmounted, ref, shallowRef, watch } from 'vue';
import type { SuperDoc } from 'superdoc';
import RibbonGroup from '../common/RibbonGroup.vue';
import RibbonStack from '../common/RibbonStack.vue';
import RibbonButton from '../common/RibbonButton.vue';
import RibbonMenuButton from '../common/RibbonMenuButton.vue';
import SpacingPreview from '../common/SpacingPreview.vue';
import type { SelectOption } from '../common/RibbonSelect.vue';
import RibbonCombo from '../common/RibbonCombo.vue';
import ColorPickerPopover from '../common/ColorPickerPopover.vue';
import StyleGallery from '../common/StyleGallery.vue';
import StyleDialog from '../../panels/StyleDialog.vue';
import { useStyleManager } from '../../../composables/use-style-manager';
import { useCommand } from '../../../composables/useCommand';
import { useFontControls } from '../../../composables/use-font-controls';
import { applyOptimistically, withCurrent } from '../../../composables/picker-value';
import { COMMAND_REPORTER, type CommandReporter } from '../../../composables/keys';
import type { CommandOutcome } from '../../../engine/command-adapter';
import { ACTIVE_SUPERDOC } from '../../../engine/document-api';
import { readDocCapabilities, type DocCapabilityReport } from '../../../engine/doc-capabilities';
import {
  copySelection,
  cutSelection,
  pasteFromClipboard,
  selectWholeDocument,
} from '../../../engine/clipboard';
import {
  PASTE_MODES,
  isPasteMode,
  type PasteMode,
} from '../../../engine/paste-formatting';
import { pasteDefaultMode, setPasteDefault } from '../../../composables/paste-default';
import {
  readVertAlignSupport,
  toggleVertAlign,
  type VertAlignKind,
  type VertAlignSupport,
} from '../../../engine/vert-align';
import ParagraphDialog from '../../panels/ParagraphDialog.vue';
import FontAdvancedDialog from '../../panels/FontAdvancedDialog.vue';
import {
  TWIPS_PER_CM,
  TWIPS_PER_PT,
  addParagraphTabStop,
  applyParagraphIndentation,
  applyParagraphContextualSpacing,
  applyParagraphKeepOptions,
  applyParagraphSpacing,
  clearAllParagraphTabStops,
  emptyParagraphFormat,
  readParagraphFormat,
  readParagraphIndents,
  readSelectionSpacing,
  removeParagraphTabStop,
  toggleSelectionSpacing,
  type LineSpacingRule,
  type ParagraphSpacingEntry,
  type TabStop,
} from '../../../engine/paragraph-format';
import {
  applyFontAdvanced,
  type FontAdvancedPatch,
} from '../../../engine/font-advanced';
import { readPageMargins } from '../../../engine/page-setup';
import { invalidateDocumentReads, observeDocumentReadInvalidation } from '../../../engine/document-read-state';
import {
  NUMBER_STYLE_LABELS,
  continuePreviousList,
  convertListToText,
  restartListAt,
  setListNumberStyle,
  setListToBullets,
} from '../../../engine/lists';
import {
  DEFAULT_LINE_HEIGHT,
  MAX_LINE_HEIGHT,
  MIN_LINE_HEIGHT,
  alignmentPayload,
  colorPayload,
  grownLineHeight,
  lineHeightPayload,
  parseColor,
  parseLineHeight,
  shrunkLineHeight,
  stylePayload,
  type ParagraphAlignment,
} from '../../../engine/payloads';

const emit = defineEmits<{
  (e: 'open-find'): void;
  (e: 'open-replace'): void;
  (e: 'toggle-list-autoformat'): void;
}>();

/**
 * המתג של „זיהוי רשימות בהקלדה”. הוא יושב בתפריט המספור ולא בדיאלוג הגדרות
 * (שאין) מסיבה אחת: מי שרוצה לכבות אותו הוא מי שהמנגנון הרגיז אותו זה עתה,
 * והוא מסתכל בדיוק על הפקד הזה. ב-Word הוא נמצא ב„אפשרויות תיקון שגיאות”,
 * שהוא שלושה מסכים משם.
 */
const props = withDefaults(defineProps<{
  listAutoformatEnabled?: boolean;
  /** Only visual paragraph reads pause; commands and shared font state stay live. */
  readoutsVisible?: boolean;
}>(), {
  listAutoformatEnabled: true,
  readoutsVisible: true,
});

const SPACING_OPTIONS: SelectOption[] = [
  { value: '1.0', label: '1.0' },
  { value: '1.15', label: '1.15' },
  { value: '1.5', label: '1.5' },
  { value: '2.0', label: '2.0' },
  { value: '2.5', label: '2.5' },
  { value: '3.0', label: '3.0' },
];

// פקודות SuperDoc
const boldCmd = useCommand('bold');
const italicCmd = useCommand('italic');
const underlineCmd = useCommand('underline');
const strikeCmd = useCommand('strikethrough');
const clearFormatCmd = useCommand('clear-formatting');
const formatPainterCmd = useCommand('copy-format');

const fontColorCmd = useCommand('text-color');
const highlightCmd = useCommand('highlight-color');

const bulletCmd = useCommand('bullet-list');
const numberedCmd = useCommand('numbered-list');
const indentIncCmd = useCommand('indent-increase');
const indentDecCmd = useCommand('indent-decrease');
const dirRtlCmd = useCommand('direction-rtl');
const dirLtrCmd = useCommand('direction-ltr');
const marksCmd = useCommand('formatting-marks');
const alignCmd = useCommand('text-align');
const lineSpacingCmd = useCommand('line-height');
const styleCmd = useCommand('linked-style');

/**
 * „שנה סגנון”, „צור סגנון” ושאר התפריט של הגלריה — composables/use-style-manager.ts.
 * הסגנון הנוכחי הוא זה שהגלריה מסמנת: בבחירה מעורבת אין כזה, ו„צור” נשען
 * אז על „רגיל”.
 */
const styleManager = useStyleManager(() => {
  const value = styleCmd.value.value;
  return typeof value === 'string' && value.trim() !== '' ? value : null;
});

/**
 * בורר הגופן ובורר הגודל. המצב שלהם אינו כאן אלא ב-use-font-controls.ts, ומאותו
 * טעם שבו האפשרויות אינן כאן: אותם שני בוררים מופיעים גם בתפריט הלחצן הימני,
 * ועותק פרטי לכל אחד מהם היה מציג בהם ערכים שונים באותו רגע — בדיוק ברגע שבו
 * המנוע אינו מדווח ערך.
 */
const fonts = useFontControls();

/* ------------------------------------------------------------------ */
/* מה שהמנוע מדווח על הבחירה                                            */
/* ------------------------------------------------------------------ */

/**
 * `CommandState.value` הוא המקור לערך שהפקד מציג — לא ref מקומי. שלושת
 * הבוררים היו פעם refs שאותחלו לערך קשיח ולעולם לא התעדכנו: לחיצה על טקסט
 * ב-20pt השאירה „12” בתיבה. בורר הגופן והגודל עברו ל-use-font-controls.ts;
 * מה שנשאר כאן הוא מרווח השורות ושני הצבעים.
 */
const engineLineHeight = computed(() => parseLineHeight(lineSpacingCmd.value.value));
const engineTextColor = computed(() => parseColor(fontColorCmd.value.value));
const engineHighlight = computed(() => parseColor(highlightCmd.value.value));

/**
 * מה שהמנוע דיווח לאחרונה.
 *
 * למה בכלל נדרש זיכרון כזה: המנוע מדווח `undefined` בשני מצבים שונים — בחירה
 * עם יותר מערך אחד, ורגע שבו הוא עוד לא פתר את הבחירה. הבורר הוא **פקד**, ולא
 * דוח; לרוקן אותו בכל תנועת סמן היה גרוע יותר מלהציג את הערך האחרון שכן ידענו.
 */
const lastLineHeight = ref(DEFAULT_LINE_HEIGHT);

/**
 * מה שהמשתמש בחר וטרם קיבל תשובה מהמנוע. `null` = אין בקשה באוויר.
 *
 * למה שכבה נפרדת ולא כתיבה לזיכרון שלמעלה, וזה **הבאג שהיה כאן**: הזיכרון
 * נכתב לפני `run()` ולא הוחזר בכשל, ולכן הבורר הציג גופן שלא הוחל על שום דבר —
 * אותה תקלה בדיוק שתוקנה בזום, שבה הסרגל הזיז את התווית ורוחב העמוד לא זז.
 * מסוכן ממנה היה „הגדל גופן” (שעבר מכאן ל-use-font-controls.ts), שמחשב מהערך
 * הזה: שלוש לחיצות על מסמך שדוחה שלחו 14, 16, 18 — הפקד התרחק מהמסמך בכל לחיצה.
 *
 * ההפרדה גם מה שמתקן את המקרה הקשה יותר, שבו למנוע **יש** ערך: כתיבה לזיכרון
 * לא שינתה שם את הערך המוצג בכלל (המנוע גובר עליו), ולכן לא היה רינדור חוזר —
 * ו-Vue מסנכרן `<select :value>` רק כשהערך הקשור משתנה. כלומר הבחירה שנדחתה
 * נשארה על המסך עד לרינדור חוזר מסיבה אחרת לגמרי. עם שכבת ה-`pending` הערך
 * המוצג משתנה בשני הכיוונים, והסנכרון נקשר לתשובה עצמה.
 */
const pendingLineHeight = ref<number | null>(null);

/**
 * הפקודה אינה מדווחת ערך, ולכן בקשות חופפות אינן רק עניין של ציור: הצלחה של
 * 1.6 וכשל של 1.7 שאחריה היו מחזירים את הבורר ל-1.5, אף שהמסמך נשאר ב-1.6.
 * נעילה קצרה עד הקבלה משאירה לכל בקשה נקודת חזרה אחת נכונה.
 */
const lineSpacingInFlight = ref(false);

watch(engineLineHeight, (value) => {
  if (value) lastLineHeight.value = value;
});

/**
 * סדר העדיפויות: מה שנבחר וטרם נענה, אחר כך מה שהמנוע מדווח, ולבסוף האחרון
 * שידענו. בקשה שנדחתה נעלמת מהשכבה הראשונה, ואז המסמך חוזר להיות מה שמוצג.
 */
const currentLineHeight = computed(
  () => pendingLineHeight.value ?? engineLineHeight.value ?? lastLineHeight.value,
);

/** „1.5” ולא „1.50”, כדי שהערך יתאים לאפשרות בבורר. */
const selectedLineSpacing = computed(() => currentLineHeight.value.toFixed(2).replace(/0$/, ''));

/**
 * צבע המסמך — מה שמסמן את המשבצת בפלטה, ולא הצבע שהכפתור יחיל. הצבע שהכפתור
 * מחיל הוא בחירה שנדבקת לפקד עצמו; ראו `activeColor` ב-ColorPickerPopover.
 */
const textColor = computed(() => engineTextColor.value ?? '');
const highlightColor = computed(() => engineHighlight.value ?? '');

/**
 * שני כפתורי הצעד נכבים בקצות הסולם. `currentLineHeight` ולא מה שהמנוע
 * מדווח: מה שהכפתור מזיז הוא מה שמוצג, כולל בחירה שטרם נענתה — אחרת שתי
 * לחיצות רצופות היו שולחות את אותו ערך פעמיים.
 */
const canGrowSpacing = computed(() => currentLineHeight.value < MAX_LINE_HEIGHT);
const canShrinkSpacing = computed(() => currentLineHeight.value > MIN_LINE_HEIGHT);

/**
 * ההסבר בכרטיס העזרה נוקב בערך שהלחיצה **תגיע** אליו ולא בצעד עצמו: „ל-1.2”
 * אומר יותר מ„ב-0.1” כשהמרווח הוא 1.1. בקצה הסולם הוא אומר זאת במפורש, מפני
 * ש**הכפתור כבוי שם** — והכרטיס הוא המקום היחיד שמסביר למה, ומפנה לדיאלוג
 * שכן יודע לרדת מתחת לסולם הזה („בדיוק” ו„לפחות”, בנקודות).
 */
const growSpacingHint = computed(() =>
  canGrowSpacing.value
    ? `מרחיב את מרווח השורות ל-${grownLineHeight(currentLineHeight.value)}`
    : `${MAX_LINE_HEIGHT} הוא המרווח הגדול ביותר כאן. מעליו — „תפריט פסקה”`,
);
const shrinkSpacingHint = computed(() =>
  canShrinkSpacing.value
    ? `מצמצם את מרווח השורות ל-${shrunkLineHeight(currentLineHeight.value)}`
    : `${MIN_LINE_HEIGHT} הוא המרווח הקטן ביותר כאן. מתחתיו — „תפריט פסקה”`,
);

/* ------------------------------------------------------------------ */
/* תפריט „מרווח שורות וריווח”                                          */
/* ------------------------------------------------------------------ */

/**
 * כמה „הוסף רווח לפני/אחרי הפסקה” מוסיף: 12 נקודות.
 *
 * זה הערך ש-Word קובע ב-Ctrl+0 מאז שהקיצור קיים, ולכן הוא גם מה שהמשתמש
 * שמגיע משם מצפה לו. הוא יושב כאן ולא ב-payloads.ts מפני שאינו חוזה מול
 * ולידטור של המנוע אלא בחירת מוצר — המנוע מקבל כל מספר לא-שלילי.
 */
const SPACE_STEP_TWIPS = 12 * TWIPS_PER_PT;

const SPACE_BEFORE_ID = 'space-before';
const SPACE_AFTER_ID = 'space-after';
const PARAGRAPH_OPTIONS_ID = 'paragraph-options';
const LINE_ITEM_PREFIX = 'line:';

/** הכפולה שפריט מרווח שורות נושא, או `null` אם אינו כזה. */
function lineItemMultiplier(id: string | null): number | null {
  if (id === null || !id.startsWith(LINE_ITEM_PREFIX)) return null;
  return parseLineHeight(id.slice(LINE_ITEM_PREFIX.length));
}

/**
 * הריווח של כל פסקה בבחירה, כפי שנקרא בפתיחת התפריט. `[]` = טרם נקרא, או
 * שאין מה לקרוא (אין סמן, המסמך נטען).
 *
 * למה נקרא בפתיחה ולא מוחזק חי: הקריאה עוברת ב-`doc.get()`, שסורק את המסמך
 * כולו. מעקב מתמיד אחרי הסמן היה משלם את המחיר הזה בכל תזוזה — זו הסיבה
 * שהסרגל משהה את הקריאה שלו ומפסיק כשהוא מוסתר (ראו `readParagraphIndents`).
 */
const spacingEntries = shallowRef<readonly ParagraphSpacingEntry[]>([]);

/** הפריט שהעכבר עומד עליו, למען פס התצוגה. `null` = אף אחד. */
const spacingHover = ref<string | null>(null);

/** הריווח של הפסקה שהתווית מדברת עליה — הראשונה בבחירה, כמו ב-Word. */
const spacingHead = computed(() => spacingEntries.value[0]?.spacing ?? null);

/**
 * האם כבר יש רווח לפני/אחרי הפסקה — כלומר אם הפריט אומר „הסר” או „הוסף”.
 *
 * ההכרעה נופלת על מה ש**מוצהר בפסקה**, ולא על מה שהיא מציגה בפועל: ריווח
 * שיורש מהסגנון אינו נקרא מ-`doc.get()` (ראו `spacingFromProps`). המשמעות
 * המעשית היא מקרה קצה אחד — פסקה שהסגנון שלה כבר קובע 12 נק' תראה „הוסף”,
 * והלחיצה תקבע במפורש את אותו מספר. לא נמדד שינוי חזותי, וזו גם ההתנהגות
 * הסבירה: מה שהמשתמש רואה נשאר.
 */
const hasSpaceBefore = computed(() => (spacingHead.value?.beforeTwips ?? 0) > 0);
const hasSpaceAfter = computed(() => (spacingHead.value?.afterTwips ?? 0) > 0);

/**
 * מה שהפס מצייר: מצב הפסקה, ועליו מה שהפריט שמרחפים עליו יעשה.
 *
 * `lineTwips` נופל ל-`currentLineHeight` כשהפסקה אינה מצהירה מרווח שורות,
 * ולא ל-240: הבורר שהיה כאן הציג את מה שהמנוע מדווח — כלומר את הערך האפקטיבי,
 * כולל זה שיורש מהסגנון — ופס שהיה מצייר „בודדת” על פסקה שנראית 1.5 היה סותר
 * את מה שרואים מאחוריו.
 */
const previewSpacing = computed<{
  beforeTwips: number;
  afterTwips: number;
  lineTwips: number;
  rule: LineSpacingRule;
}>(() => {
  const head = spacingHead.value;
  const base = {
    beforeTwips: head?.beforeTwips ?? 0,
    afterTwips: head?.afterTwips ?? 0,
    lineTwips: head?.lineTwips ?? Math.round(currentLineHeight.value * 240),
    rule: head?.rule ?? ('auto' as LineSpacingRule),
  };
  const hovered = spacingHover.value;
  if (hovered === null) return base;
  if (hovered === SPACE_BEFORE_ID) {
    return { ...base, beforeTwips: hasSpaceBefore.value ? 0 : SPACE_STEP_TWIPS };
  }
  if (hovered === SPACE_AFTER_ID) {
    return { ...base, afterTwips: hasSpaceAfter.value ? 0 : SPACE_STEP_TWIPS };
  }
  const multiplier = lineItemMultiplier(hovered);
  if (multiplier === null) return base;
  // ריחוף על ערך מרווח שורות מצייר אותו כ„אוטומטי” גם כשהפסקה ב„בדיוק”:
  // זה מה שהלחיצה תעשה — פקודת `line-height` שולחת כפולה, לא מרחק.
  return { ...base, lineTwips: Math.round(multiplier * 240), rule: 'auto' };
});

/**
 * הפריטים, בסדר של Word: הערכים, „אפשרויות…”, ואז שני הרווחים.
 *
 * `withCurrent` נשאר מהבורר שהיה כאן, ולא הוחלף ב-`checked` בלבד: כפתורי
 * הצעד שלצד התפריט מייצרים ערכים שאינם ברשימה (1.2, 1.35), ורשימה קבועה
 * הייתה משאירה אותם בלי שום סימן — כלומר לחיצה על „הגדל” שאינה נראית בשום
 * מקום. `numeric` הוא מה שמכניס אותם **במקומם** בסולם ולא בראשו.
 */
const spacingMenuItems = computed(() => [
  ...withCurrent(SPACING_OPTIONS, selectedLineSpacing.value, { numeric: true }).map((option) => ({
    id: `${LINE_ITEM_PREFIX}${option.value}`,
    label: option.label,
    checked: option.value === selectedLineSpacing.value,
  })),
  {
    id: PARAGRAPH_OPTIONS_ID,
    label: 'אפשרויות מרווח שורות…',
    separatorBefore: true,
  },
  {
    id: SPACE_BEFORE_ID,
    label: hasSpaceBefore.value ? 'הסר רווח לפני הפסקה' : 'הוסף רווח לפני הפסקה',
    separatorBefore: true,
  },
  {
    id: SPACE_AFTER_ID,
    label: hasSpaceAfter.value ? 'הסר רווח אחרי הפסקה' : 'הוסף רווח אחרי הפסקה',
  },
]);

/**
 * פתיחת התפריט: קוראים את מצב הבחירה, ואת גיאומטריית העמוד שהפס נמדד בה.
 *
 * שתי הקריאות אינן מעכבות את הפתיחה — התפריט כבר מצויר כשהן יוצאות לדרך.
 * עד שהן חוזרות התוויות אומרות „הוסף” (ברירת המחדל של פסקה שאינה מצהירה),
 * והפס מצייר את מרווח השורות שהבורר ידע ממילא. כישלון של אחת מהן אינו מדווח:
 * תפריט שנפתח אינו פעולה של המשתמש שאפשר להתלונן עליה, והלחיצה עצמה כן מדווחת.
 */
async function onSpacingMenuOpen(): Promise<void> {
  spacingHover.value = null;
  spacingEntries.value = (await readSelectionSpacing(superdoc.value)) ?? [];
  // `paraTextWidthTwips`/`paraSectionRtl` משותפים עם דיאלוג הפסקה, ובכוונה:
  // שניהם מתארים את **המקטע** ולא את הפקד, ושתי עותקים שלהם היו יכולים לומר
  // דברים שונים על אותו עמוד. מי שנפתח אחרון פשוט קורא אותם מחדש.
  const page = await readPageMargins(superdoc.value);
  paraTextWidthTwips.value = page
    ? Math.max(0, page.pageWidthTwips - page.leftTwips - page.rightTwips)
    : 0;
  paraSectionRtl.value = page?.direction === 'rtl';
}

function onSpacingMenuHover(id: string | null): void {
  spacingHover.value = id;
}

function onSpacingMenuSelect(id: string): void {
  if (id === PARAGRAPH_OPTIONS_ID) {
    void onOpenParagraph();
    return;
  }
  if (id === SPACE_BEFORE_ID || id === SPACE_AFTER_ID) {
    void applySpacePatch(id);
    return;
  }
  const multiplier = lineItemMultiplier(id);
  if (multiplier !== null) void applyLineSpacing(multiplier);
}

/**
 * „הוסף/הסר רווח” על **כל** הפסקאות המסומנות.
 *
 * הכלל — מתי זה מוסיף ומתי מסיר — יושב ב-`toggleSelectionSpacing` ולא כאן,
 * מפני שהקיצור Ctrl+0 מפעיל בדיוק את אותו דבר מהמקלדת. הוא גם קורא את הבחירה
 * מחדש ואינו נשען על מה שנקרא בפתיחת התפריט: בין פתיחה ללחיצה אפשר לבטל
 * בקיצור מקלדת, ורשימה ישנה הייתה כותבת על מצב שכבר אינו קיים.
 */
async function applySpacePatch(id: string): Promise<void> {
  if (lineSpacingInFlight.value) return;
  lineSpacingInFlight.value = true;
  try {
    const outcome = await toggleSelectionSpacing(
      superdoc.value,
      id === SPACE_BEFORE_ID ? 'before' : 'after',
      SPACE_STEP_TWIPS,
    );
    if (!outcome.ok) {
      report(outcome, 'paragraph-spacing');
      return;
    }
    // התוויות („הוסף” מול „הסר”) נגזרות מהמצב שנקרא, ולכן הן חייבות להתרענן
    // אחרי כתיבה — אחרת פתיחה שנייה של התפריט מציעה להוסיף מה שהרגע נוסף.
    spacingEntries.value = (await readSelectionSpacing(superdoc.value)) ?? [];
  } finally {
    lineSpacingInFlight.value = false;
  }
}

/* ------------------------------------------------------------------ */
/* הפעלה                                                              */
/* ------------------------------------------------------------------ */

// כל ה-payloads נבנים ב-engine/payloads.ts, ולא כליטרל כאן: מה שנשלח לפקודה
// הוא חוזה מול ולידטור בתוך המנוע, והוולידטור נכשל **סגור**. ראו את הטבלה
// שם, ואת בדיקת החוזה ב-tests/contract/command-payloads.test.ts.
function onTextColorChange(color: string | null): void {
  void fontColorCmd.run(colorPayload(color));
}

function onHighlightChange(color: string | null): void {
  void highlightCmd.run(colorPayload(color));
}

async function applyLineSpacing(multiplier: number): Promise<void> {
  if (lineSpacingInFlight.value) return;
  const payload = lineHeightPayload(multiplier);
  if (payload === null) return;
  lineSpacingInFlight.value = true;
  try {
    await applyOptimistically(pendingLineHeight, lastLineHeight, multiplier, () =>
      lineSpacingCmd.run(payload),
    );
  } finally {
    lineSpacingInFlight.value = false;
  }
}

/**
 * צעד אחד במרווח, מהערך שמוצג עכשיו. `step` הוא פונקציה ולא מספר כדי
 * שהחישוב וההגבלה יישארו ב-engine/payloads.ts, במקום שבו הם נבדקים.
 */
function onLineSpacingStep(step: (current: number) => number): void {
  const multiplier = step(currentLineHeight.value);
  void applyLineSpacing(multiplier);
}

function onAlign(alignment: ParagraphAlignment): void {
  void alignCmd.run(alignmentPayload(alignment));
}

function onApplyStyle(styleId: string): void {
  const payload = stylePayload(styleId);
  if (payload === null) return;
  void styleCmd.run(payload);
}

/* ------------------------------------------------------------------ */
/* לוח                                                                */
/* ------------------------------------------------------------------ */

/**
 * פעולות הלוח אינן פקודות ב-registry של ה-controller — אין להן מזהה בקטלוג —
 * ולכן הן עוברות ב-Document API דרך engine/clipboard.ts, ומדווחות ידנית. ראו
 * ReferencesTab: אותו דפוס בדיוק, כולל קריאת היכולות בכל החלפת מסמך.
 */
const fallbackReporter: CommandReporter = (outcome, id) => {
  if (!outcome.ok) console.warn(`[otzaria-word] ${id}: ${outcome.message}`);
};

const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null));

const report = inject(COMMAND_REPORTER, fallbackReporter);

/* ------------------------------------------------------------------ */
/* חיווי „פעיל” של RTL/LTR                                             */
/* ------------------------------------------------------------------ */

/** The registry's direction active state is unreliable; use the targeted paragraph read shared with the ruler. */
const paragraphBidi = shallowRef<boolean | null>(null);
const BIDI_SELECTION_DEBOUNCE_MS = 150;
let bidiSelectionTimer: ReturnType<typeof setTimeout> | undefined;
let bidiGeneration = 0;
let bidiDisposed = false;
let unsubscribeBidiSelection: (() => void) | null = null;
let unsubscribeBidiDocument: (() => void) | null = null;
let bidiDocument: object | null | undefined;

function bindBidiDocument(): void {
  const doc = superdoc.value?.activeEditor?.doc;
  if (bidiDocument === doc) return;
  unsubscribeBidiDocument?.();
  bidiDocument = doc;
  unsubscribeBidiDocument = observeDocumentReadInvalidation(doc, scheduleBidiRefresh);
}

async function refreshParagraphBidi(): Promise<void> {
  if (bidiDisposed || !props.readoutsVisible) return;
  clearTimeout(bidiSelectionTimer);
  bindBidiDocument();
  const mine = ++bidiGeneration;
  const reading = await readParagraphIndents(superdoc.value);
  if (mine !== bidiGeneration || bidiDisposed || !props.readoutsVisible) return;
  paragraphBidi.value = reading ? reading.indents.bidi : null;
}

function scheduleBidiRefresh(): void {
  if (bidiDisposed || !props.readoutsVisible) return;
  // Reject a previous response even while the next read is still debounced.
  bidiGeneration += 1;
  bindBidiDocument();
  clearTimeout(bidiSelectionTimer);
  bidiSelectionTimer = setTimeout(() => void refreshParagraphBidi(), BIDI_SELECTION_DEBOUNCE_MS);
}

function stopBidiReadout(): void {
  bidiGeneration += 1;
  clearTimeout(bidiSelectionTimer);
  bidiSelectionTimer = undefined;
  unsubscribeBidiSelection?.();
  unsubscribeBidiSelection = null;
  unsubscribeBidiDocument?.();
  unsubscribeBidiDocument = null;
  bidiDocument = undefined;
}

watch(
  [superdoc, () => props.readoutsVisible],
  ([host, visible], [previousHost, wasVisible]) => {
    stopBidiReadout();
    if (host !== previousHost) paragraphBidi.value = null;
    if (!visible) return;
    bindBidiDocument();
    const selection = host?.ui?.selection;
    if (typeof selection?.observe === 'function') {
      try {
        unsubscribeBidiSelection = selection.observe(scheduleBidiRefresh) ?? null;
      } catch {
        unsubscribeBidiSelection = null;
      }
    }
    // observe() may schedule its initial snapshot; the immediate reveal read
    // clears that timer, so opening the ribbon requests one current reading.
    if (wasVisible === false) void refreshParagraphBidi();
    else scheduleBidiRefresh();
  },
  { immediate: true, flush: 'sync' },
);

onUnmounted(() => {
  bidiDisposed = true;
  stopBidiReadout();
});

/**
 * מה שהכפתורים מציגים בפועל. `paragraphBidi.value === null` פירושו „אין
 * עדיין קריאה תקפה" (אין סמן, מסמך נטען) — ואז נופלים חזרה למה שהמנוע
 * מדווח, כדי לא לייצר טענה שלא נמדדה. בפועל זה כמעט תמיד `false` (הבאג
 * המתועד למעלה), אבל זו נפילה זהה לזו שהייתה קיימת לפני התיקון, לא גרועה
 * ממנה.
 */
const dirRtlActive = computed(() =>
  paragraphBidi.value !== null ? paragraphBidi.value === true : dirRtlCmd.active.value,
);
const dirLtrActive = computed(() =>
  paragraphBidi.value !== null ? paragraphBidi.value === false : dirLtrCmd.active.value,
);

/**
 * לחיצה על כפתורי הכיוון: מריצה את הפקודה (שכותבת `<w:bidi>` נכון — זה
 * עצמו לא שבור, ראו התיעוד למעלה) ומרעננת מיד את החיווי, בלי לחכות
 * ל-debounce של תזוזת סמן. הלחיצה עצמה לא מזיזה את הסמן, ולכן בלי הרענון
 * המיידי הזה `ui.selection.observe` שמעל לא היה יורה כלל.
 */
async function runDirectionRtl(): Promise<void> {
  const host = superdoc.value;
  const doc = host?.activeEditor?.doc;
  await dirRtlCmd.run();
  invalidateDocumentReads(doc);
  if (superdoc.value === host) await refreshParagraphBidi();
}
async function runDirectionLtr(): Promise<void> {
  const host = superdoc.value;
  const doc = host?.activeEditor?.doc;
  await dirLtrCmd.run();
  invalidateDocumentReads(doc);
  if (superdoc.value === host) await refreshParagraphBidi();
}

const capabilities = shallowRef<DocCapabilityReport | null>(null);

/** קריאת היכולות א-סינכרונית, ותשובה של מסמך קודם לא תדרוס את הנוכחי. */
let generation = 0;

watch(
  superdoc,
  async (host) => {
    const mine = ++generation;
    capabilities.value = null;
    const result = await readDocCapabilities(host);
    if (mine === generation) capabilities.value = result;
  },
  { immediate: true }
);

const canCopy = computed(() => capabilities.value?.can('canCopySelection') ?? false);

/** „גזור” הוא סדרוּר **ומחיקה**. מנוע שיודע רק להעתיק משאיר „העתק” פעיל. */
const canCut = computed(
  () => canCopy.value && (capabilities.value?.can('canDeleteSelection') ?? false),
);

const canPaste = computed(() => capabilities.value?.can('canPasteContent') ?? false);

/**
 * „בחר הכל” עובר ב-`ranges.resolve` (גבולות גוף המסמך) ואז ב-
 * `ui.selection.apply`. רק הראשון נשאל ביכולות — לשני אין מזהה פעולה בקטלוג,
 * והוא נכשל סגור עם `reason` משלו.
 */
const canSelectAll = computed(() => capabilities.value?.can('canResolveRange') ?? false);

type ClipboardQuestion =
  | 'canCopySelection'
  | 'canDeleteSelection'
  | 'canPasteContent'
  | 'canResolveRange';

function tooltipFor(enabled: boolean, question: ClipboardQuestion, text: string): string {
  if (enabled) return text;
  return capabilities.value?.explain(question) || 'המסמך עדיין נטען';
}

const copyTooltip = computed(() =>
  tooltipFor(canCopy.value, 'canCopySelection', 'העתקת הבחירה ללוח'),
);
const cutTooltip = computed(() =>
  tooltipFor(
    canCut.value,
    canCopy.value ? 'canDeleteSelection' : 'canCopySelection',
    'גזירת הבחירה ללוח',
  ),
);
const pasteTooltip = computed(() =>
  tooltipFor(canPaste.value, 'canPasteContent', 'הדבקת תוכן מהלוח'),
);
const selectAllTooltip = computed(() =>
  tooltipFor(canSelectAll.value, 'canResolveRange', 'בחירת כל הטקסט במסמך'),
);

async function doCopy(): Promise<void> {
  report(await copySelection(superdoc.value), 'clipboard-copy');
}

async function doCut(): Promise<void> {
  report(await cutSelection(superdoc.value), 'clipboard-cut');
}

async function doPaste(mode: PasteMode): Promise<void> {
  report(await pasteFromClipboard(superdoc.value, { mode }), 'clipboard-paste');
}

/**
 * המחרוזות של כל אפשרות, מילוליות ולא מורכבות: כל מחרוזת שהרצועה מציגה
 * עוברת ב-`menuString`, והמילון (ui/ribbon/i18n.ts) מחפש אותה כמות שהיא.
 */
const PASTE_OPTIONS: Record<PasteMode, { label: string; hint: string; asDefault: string; onClick: string }> = {
  'keep-source': {
    label: 'שמירה על עיצוב המקור',
    hint: 'הגופן, הגודל, הצבע והעיצוב כמו במקום שממנו הועתק',
    asDefault: 'Ctrl+V: שמירה על עיצוב המקור',
    onClick: 'לחיצה: שמירה על עיצוב המקור',
  },
  merge: {
    label: 'מיזוג עיצוב',
    hint: 'העיצוב של המקום שאליו מדביקים, עם המודגש, הנטוי והקו התחתון של המקור',
    asDefault: 'Ctrl+V: מיזוג עיצוב',
    onClick: 'לחיצה: מיזוג עיצוב',
  },
  'text-only': {
    label: 'שמירת טקסט בלבד',
    hint: 'הטקסט בלבד, בעיצוב של המקום שאליו מדביקים',
    asDefault: 'Ctrl+V: שמירת טקסט בלבד',
    onClick: 'לחיצה: שמירת טקסט בלבד',
  },
};

const pasteDescription = computed(() =>
  canPaste.value ? PASTE_OPTIONS[pasteDefaultMode.value].onClick : undefined,
);

/**
 * שתי קבוצות: „אפשרויות הדבקה” — הדבקה חד-פעמית באופן שנבחר — ו„הדבקה
 * רגילה” — מה ש-Ctrl+V וגוף הכפתור עושים מעכשיו (composables/paste-default.ts).
 * המזהה נושא את הקבוצה, כדי שלחיצה על אותה אפשרות בשתיהן לא תתבלבל.
 */
const pasteMenuItems = computed(() => [
  ...PASTE_MODES.map((mode) => ({
    id: `paste:${mode}`,
    label: PASTE_OPTIONS[mode].label,
    hint: PASTE_OPTIONS[mode].hint,
  })),
  ...PASTE_MODES.map((mode, index) => ({
    id: `default:${mode}`,
    label: PASTE_OPTIONS[mode].asDefault,
    checked: pasteDefaultMode.value === mode,
    ...(index === 0 ? { separatorBefore: true, hint: 'ההדבקה הרגילה — נשמרת להפעלות הבאות' } : {}),
  })),
]);

function onPasteMenuSelect(id: string): void {
  const [group, mode] = id.split(':');
  if (!isPasteMode(mode)) return;
  if (group === 'paste') void doPaste(mode);
  else if (group === 'default') setPasteDefault(mode);
}

async function doSelectAll(): Promise<void> {
  report(await selectWholeDocument(superdoc.value), 'select-all');
}

/* ------------------------------------------------------------------ */
/* כתב עליון וכתב תחתי                                                 */
/* ------------------------------------------------------------------ */

/**
 * שני הפקדים היו `:disabled="true"` קשיח עם tooltip שהאשים את המנוע („אינו
 * נתמך במנוע הנוכחי”). הטענה לא הייתה נכונה: `format.vertAlign` הוא פעולה
 * ציבורית בקטלוג, ומה שחסר היה פקודה ברצועה. הם עוברים מעכשיו ב-Document API,
 * בדיוק כמו פעולות הלוח שמעל — ההסבר המלא, כולל למה אין כאן חיווי „דלוק”,
 * ב-engine/vert-align.ts.
 *
 * זמינות נפרדת מזו של הלוח: היא נשענת גם על נוכחות `doc.format.vertAlign`
 * בפאסדה ולא רק על מפת היכולות, וזו בדיקה שהדוח הכללי אינו עושה.
 */
const vertAlign = shallowRef<VertAlignSupport>({ available: false, explanation: 'המסמך עדיין נטען' });

/** ראו למעלה: תשובה של מסמך קודם לא תדרוס את הנוכחי. */
let vertAlignGeneration = 0;

watch(
  superdoc,
  async (host) => {
    const mine = ++vertAlignGeneration;
    vertAlign.value = { available: false, explanation: 'המסמך עדיין נטען' };
    const support = await readVertAlignSupport(host);
    if (mine === vertAlignGeneration) vertAlign.value = support;
  },
  { immediate: true }
);

const canSetVertAlign = computed(() => vertAlign.value.available);

const VERT_ALIGN_TEXT: Record<VertAlignKind, string> = {
  // התווית אומרת בדיוק מה הלחיצה עושה: היא מחילה על **הטקסט המסומן**, והיא
  // מכבה בלחיצה שנייה. שני הדברים הם ההתנהגות שנמדדה, ולא הבטחה.
  superscript: 'הרמת הטקסט המסומן לכתב עליון; לחיצה נוספת מחזירה אותו לשורה',
  subscript: 'הנמכת הטקסט המסומן לכתב תחתי; לחיצה נוספת מחזירה אותו לשורה',
};

function vertAlignTooltip(kind: VertAlignKind): string {
  if (canSetVertAlign.value) return VERT_ALIGN_TEXT[kind];
  return vertAlign.value.explanation || 'המסמך עדיין נטען';
}

async function onVertAlign(kind: VertAlignKind): Promise<void> {
  report(await toggleVertAlign(superdoc.value, kind), `vert-align-${kind}`);
}

/* ------------------------------------------------------------------ */
/* תפריט פסקה                                                          */
/* ------------------------------------------------------------------ */

/**
 * דיאלוג „פסקה” — כניסות, ריווח, אפשרויות שמירה וטאבים, דרך
 * `doc.format.paragraph.*` (engine/paragraph-format.ts). אין לזה פקודה בקטלוג,
 * ולכן המסלול הוא Document API — אותו דפוס של פעולות הלוח מעל.
 *
 * `paraInFlight` הוא TOCTOU-lock ולא נוחות: הדיאלוג נפתח על **תצלום** מצב
 * הפסקה (`readParagraphFormat`), ושלושת סעיפי „אישור” כותבים מצב מלא. פעולה
 * שנייה שתיקלט בזמן שהראשונה באוויר הייתה כותבת על תצלום מיושן.
 */
const paragraphOpen = shallowRef(false);
const paraInFlight = shallowRef(false);
/** יעד הפסקה שהתצלום נלקח ממנו; כל פעולת „אישור” חוזרת אליו. */
let paraTarget: unknown = null;
const paraSnapshot = shallowRef(emptyParagraphFormat());
/**
 * רוחב עמודת הטקסט של המקטע, למכנה של פס התצוגה המקדימה בדיאלוג.
 *
 * `0` = אין גיאומטריה, ואז הדיאלוג אינו מצייר פס. נקרא בפתיחה יחד עם התצלום
 * ולא נשמר מעבר לה, מאותו טעם שהתצלום עצמו נקרא בפתיחה: המשתמש יכול לשנות
 * שוליים בלשונית „פריסה” בין פתיחה לפתיחה.
 */
const paraTextWidthTwips = shallowRef(0);
/**
 * כיוון המקטע — מה שפסקה שאינה מצהירה `<w:bidi>` יורשת, ולכן הצד שבו הפס
 * מצייר את „לפני טקסט”. נקרא מאותה קריאה של `readPageMargins`.
 */
const paraSectionRtl = shallowRef(false);
/**
 * גודל הגופן שבסמן — הסרגל של הציר האנכי בפס. אותו מספר שבבורר הגודל
 * שברצועה, כלומר בלי קריאה נוספת למנוע.
 */
const paraFontSizePt = computed(() => Number(fonts.size.value) || 0);

const tabsEnabled = computed(() => capabilities.value?.can('canManageParagraphTabs') ?? false);

async function runParagraph(action: () => Promise<void>): Promise<void> {
  if (paraInFlight.value) return;
  paraInFlight.value = true;
  try {
    await action();
  } finally {
    paraInFlight.value = false;
  }
}

/**
 * קוראת את מצב הפסקה ואז פותחת — בדיוק `openWithState` של LayoutTab:
 * דיאלוג שנפתח ריק וממלא את עצמו כעבור tick הוא דיאלוג שהמשתמש עשוי לאשר
 * לפני שהוא מלא. כשל קריאה מסביר ולא פותח דיאלוג על ערכים ריקים.
 */
async function onOpenParagraph(): Promise<void> {
  if (paraInFlight.value) return;
  paraInFlight.value = true;
  try {
    const result = await readParagraphFormat(superdoc.value);
    if (!result.ok) {
      report(result.outcome, 'paragraph-open');
      return;
    }
    paraTarget = result.target;
    paraSnapshot.value = result.snapshot;
    // קריאה בלבד של `sections.list()`, אותה קריאה שהסרגל מצייר לפיה. כשל שלה
    // אינו מונע את הדיאלוג — הוא מוריד ממנו את הפס בלבד.
    const page = await readPageMargins(superdoc.value);
    paraTextWidthTwips.value = page
      ? Math.max(0, page.pageWidthTwips - page.leftTwips - page.rightTwips)
      : 0;
    paraSectionRtl.value = page?.direction === 'rtl';
    paragraphOpen.value = true;
  } finally {
    paraInFlight.value = false;
  }
}

function onParagraphSubmit(payload: {
  leftCm: number;
  rightCm: number;
  special: 'none' | 'firstLine' | 'hanging';
  amountCm: number;
  beforePt: number;
  afterPt: number;
  lineTwips: number;
  lineRule: 'auto' | 'exact' | 'atLeast';
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  /** `null` = „ללא שינוי” בפקד התלת-מצבי, ואז לא נשלחת קריאה כלל. */
  contextualSpacing: boolean | null;
}): void {
  paragraphOpen.value = false;
  const target = paraTarget;
  // המרת היחידות כאן ולא בדיאלוג: engine/paragraph-format.ts הוא המקום היחיד
  // שמכיר את שתי המערכות (cm/pt אצל המשתמש, twips ב-API).
  const indentation = {
    leftTwips: Math.round(payload.leftCm * TWIPS_PER_CM),
    rightTwips: Math.round(payload.rightCm * TWIPS_PER_CM),
    special: payload.special,
    amountTwips: Math.round(payload.amountCm * TWIPS_PER_CM),
  };
  const spacing = {
    beforeTwips: Math.round(payload.beforePt * TWIPS_PER_PT),
    afterTwips: Math.round(payload.afterPt * TWIPS_PER_PT),
    lineTwips: payload.lineTwips,
    rule: payload.lineRule,
  };
  const keep = {
    keepNext: payload.keepNext,
    keepLines: payload.keepLines,
    widowControl: payload.widowControl,
  };
  void runParagraph(async () => {
    // שלוש פעולות על אותו pPr; כשל אחד אינו מבטל את האחרים, וכל אחת מדווחת
    // בפני עצמה — NO_OP כבר מטופל בתוך המודול.
    // „ללא שינוי” אינו קריאה: `setFlowOptions` הוא patch, ולכן דילוג עליו
    // משאיר את `w:contextualSpacing` בדיוק כפי שהיה. ראו את המדידה ב-
    // `applyParagraphContextualSpacing`.
    const outcomes = [
      await applyParagraphIndentation(superdoc.value, target, indentation),
      await applyParagraphSpacing(superdoc.value, target, spacing),
      await applyParagraphKeepOptions(superdoc.value, target, keep),
      ...(payload.contextualSpacing === null
        ? []
        : [await applyParagraphContextualSpacing(superdoc.value, target, payload.contextualSpacing)]),
    ];
    for (const [index, outcome] of outcomes.entries()) {
      if (!outcome.ok) {
        report(outcome, ['paragraph-indent', 'paragraph-spacing', 'paragraph-keep', 'paragraph-contextual'][index]);
        return;
      }
    }
    report({ ok: true }, 'paragraph-format');
  });
}

function onParagraphTabAdd(tab: {
  positionTwips: number;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar';
  leader?: string;
}): void {
  void runParagraph(async () => {
    // ה-leader מגיע מה-select של הדיאלוג עם ערכי ה-union בלבד; ההיצמדות
    // ל-union נבדקת שוב בתוך המודול.
    const outcome = await addParagraphTabStop(superdoc.value, paraTarget, tab as TabStop);
    report(outcome, 'paragraph-tab-add');
  });
}

function onParagraphTabRemove(payload: { positionTwips: number }): void {
  void runParagraph(async () => {
    const outcome = await removeParagraphTabStop(superdoc.value, paraTarget, payload.positionTwips);
    report(outcome, 'paragraph-tab-remove');
  });
}

function onParagraphTabsClear(): void {
  void runParagraph(async () => {
    const outcome = await clearAllParagraphTabStops(superdoc.value, paraTarget);
    report(outcome, 'paragraph-tabs-clear');
  });
}

/* ------------------------------------------------------------------ */
/* גופן מתקדם                                                          */
/* ------------------------------------------------------------------ */

/**
 * דיאלוג „גופן מתקדם" — קריאה אחת ל-`format.apply` (engine/font-advanced.ts).
 * אין לו פקודה ב-registry, ולכן המסלול הוא Document API — כמו פעולות הלוח
 * ותפריט הפסקה מעל.
 */
const fontAdvOpen = shallowRef(false);
const fontAdvInFlight = shallowRef(false);

function onOpenFontAdvanced(): void {
  // אין קריאת מצב לפני הפתיחה: format.apply הוא patch לפי מפתח, ואין
  // קריאה ציבורית של עיצוב ריצות בבחירה (ראו vert-align.ts). הדיאלוג
  // נפתח על „ללא שינוי" ושולח רק מה שהמשתמש מילא.
  if (fontAdvInFlight.value) return;
  fontAdvOpen.value = true;
}

/**
 * הדיאלוג נסגר על **הצלחה בלבד**.
 *
 * הוא היה נסגר קודם לפני שההחלה בכלל יצאה לדרך, ולכן כשלון — „יש לסמן טקסט
 * תחילה” הוא השכיח — השאיר הודעה בשורת המצב ולקח איתו את שבעה-עשר השדות
 * שהמשתמש מילא. זה חצי מהבאג שדווח („בחרתי, לחצתי, ולא קרה כלום”): לא רק
 * שלא קרה דבר, אלא שגם לא נשאר מה לתקן ולנסות שוב.
 *
 * הדיאלוג אינו חוסם את המסמך והוא שואל את מצב הבחירה מחדש בכל שינוי
 * (FontAdvancedDialog.vue), ולכן „נשאר פתוח” אינו מבוי סתום: המשתמש מסמן
 * טקסט, „אישור” נפתח, והוא לוחץ שוב על אותם ערכים.
 *
 * `busy` בזמן הריצה, כדי שהכפתור לא יישלח פעמיים על אותה בחירה.
 */
function onFontAdvancedSubmit(patch: FontAdvancedPatch): void {
  if (fontAdvInFlight.value) return;
  fontAdvInFlight.value = true;
  void (async () => {
    try {
      const outcome = await applyFontAdvanced(superdoc.value, patch);
      report(outcome, 'font-advanced');
      if (outcome.ok) fontAdvOpen.value = false;
    } finally {
      fontAdvInFlight.value = false;
    }
  })();
}

/* ------------------------------------------------------------------ */
/* רשימה (גל 14א)                                                      */
/* ------------------------------------------------------------------ */

/**
 * פעולות רשימה שאין להן פקודה בקטלוג: המספור העברי (hebrew1 — נמדד),
 * התחלה מחדש, המשך מספור קודם, והמרה לטקסט. `listsInFlight` הוא נעילת
 * TOCTOU: `canContinuePrevious` בוליאני ואין לסמוך עליו בין קריאות.
 */
const listsInFlight = shallowRef(false);

/** אישור דו-לחיצה ל„המר לטקסט" — בלתי-הפיך (נמדד: הסמן מועתק לטקסט). */
const convertArmed = shallowRef(false);

const listsAvailable = computed(() => capabilities.value?.can('canManageLists') ?? false);

/**
 * סדר סגנונות המספור בתפריט — עברי ראשון.
 *
 * `NUMBER_STYLE_LABELS` הוא מפה, וסדר ההכנסה שלה הוא סדר ה-`numFmt` של
 * ECMA-376: `decimal` בראש ו-`hebrew1` שישי. זה הסדר הנכון למי שקורא את
 * התקן, ולא למי שכותב כאן מסמך — ולכן הסדר לתצוגה נכתב במפורש, ואינו נגזר
 * מסדר המפה. `bullet` אינו כאן: הוא התפריט של „תבליטים”, לא של „מספור”.
 */
const NUMBER_STYLE_ORDER = [
  'hebrew1',
  'hebrew2',
  'decimal',
  'upperLetter',
  'lowerLetter',
  'upperRoman',
  'lowerRoman',
] as const;

/** „המר לטקסט” — אותו פריט בשני התפריטים, כולל מצב החימוש. */
const convertItem = computed(() => ({
  id: 'convert',
  label: convertArmed.value ? 'לחץ שוב לאישור — הפעולה בלתי-הפיכה' : 'המר לטקסט…',
}));

const numberMenuItems = computed(() => [
  ...NUMBER_STYLE_ORDER.map((id) => ({ id: `style:${id}`, label: NUMBER_STYLE_LABELS[id] })),
  { id: 'restart', label: 'התחל מחדש מ-1' },
  { id: 'continue', label: 'המשך מספור קודם' },
  convertItem.value,
  {
    id: 'autoformat',
    // הגדרה ולא פעולה על הרשימה — ולכן מופרדת מהפעולות שמעליה.
    separatorBefore: true,
    label: props.listAutoformatEnabled
      ? 'זיהוי רשימות בהקלדה: פעיל'
      : 'זיהוי רשימות בהקלדה: כבוי',
  },
]);

/**
 * התפריט של „תבליטים”. `to-bullets` אינו כפילות של הכפתור שמעליו: הכפתור
 * מחיל רשימה על פסקאות, וזה מחליף את סגנון הסמן של רשימה **קיימת** — כלומר
 * הדרך להפוך רשימה ממוספרת לתבליטים בלי לפרק אותה ולבנות מחדש. לכן הוא אינו
 * מקבל `createList` ב-`onListMenuSelect`, בשונה מסגנוני המספור: על פסקה רגילה
 * הוא מסרב ואינו נוגע במסמך.
 */
const bulletMenuItems = computed(() => [
  { id: 'to-bullets', label: 'הפוך רשימה ממוספרת לתבליטים' },
  convertItem.value,
]);

/**
 * שורת ההסבר בטולטיפ של חץ התפריט — מה יש בו, ובמצב מנוטרל למה אין. השם
 * עצמו („פעולות תבליטים” / „פעולות מספור”) קבוע ואינו מתחלף בין המצבים: הוא
 * גם השם הנגיש של החץ, וגם הידית שבה שערי ה-QA מאתרים אותו.
 */
function listsMenuHint(available: string): string {
  if (listsAvailable.value) return available;
  return capabilities.value?.explain('canManageLists') || 'המסמך עדיין נטען';
}

const bulletMenuHint = computed(() =>
  listsMenuHint('החלפת רשימה ממוספרת לתבליטים, והמרה לטקסט'),
);

const numberMenuHint = computed(() =>
  listsMenuHint('סגנון מספור (כולל עברי), התחלה מחדש והמרה לטקסט'),
);

async function runList(action: () => Promise<CommandOutcome>): Promise<void> {
  if (listsInFlight.value) return;
  listsInFlight.value = true;
  try {
    report(await action(), 'lists');
  } finally {
    listsInFlight.value = false;
  }
}

function onListMenuSelect(id: string): void {
  if (id === 'autoformat') {
    convertArmed.value = false;
    emit('toggle-list-autoformat');
    return;
  }

  // כל פעולה שאינה „המר לטקסט” מנטרלת את החימוש שלו: המשתמש עשה משהו אחר
  // בינתיים, ולחיצה הבאה על „המר” חייבת להיות שוב לחיצה ראשונה.
  if (id !== 'convert') convertArmed.value = false;

  if (id.startsWith('style:')) {
    const style = id.slice('style:'.length);
    // סגנון מספור על פסקה שאינה רשימה ממספר אותה קודם (issue #14 ג׳). ראו
    // `SetNumberStyleOptions.createList` ב-engine/lists.ts.
    void runList(() => setListNumberStyle(superdoc.value, style, { createList: numberedCmd.run }));
    return;
  }

  if (id === 'to-bullets') {
    void runList(() => setListToBullets(superdoc.value));
    return;
  }

  if (id === 'restart') {
    void runList(() => restartListAt(superdoc.value, 1));
    return;
  }

  if (id === 'continue') {
    void runList(() => continuePreviousList(superdoc.value));
    return;
  }

  if (id === 'convert') {
    // אישור דו-לחיצה: הפעולה בלתי-הפיכה (נמדד). לחיצה ראשונה חומשת,
    // שנייה מבצעת. כל בחירה אחרת בדרך מנטרלת (ראו למעלה) — אחרת „המר
    // לטקסט” היה נשאר חמוש בשני התפריטים עד סוף החיים של הלשונית.
    if (!convertArmed.value) {
      convertArmed.value = true;
      return;
    }
    convertArmed.value = false;
    void runList(() => convertListToText(superdoc.value));
  }
}
</script>

<style scoped>
.ribbon-tab-pane {
  display: flex;
  align-items: stretch;
  gap: 0;
  height: 100%;
  width: 100%;
}
</style>
