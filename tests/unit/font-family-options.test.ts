import { describe, expect, it, vi } from 'vitest';
import { composeFontOptions } from '../../src/engine/font-options';
import { familyPickerOptions } from '../../src/composables/font-family-options';
import { withCurrent } from '../../src/composables/picker-value';

describe('font picker pending state', () => {
  it('does not probe a selected family missing from the option list', () => {
    const canvas = vi.spyOn(HTMLCanvasElement.prototype, 'getContext');
    try {
      const families = composeFontOptions(null).families;
      const options = familyPickerOptions(families, 'New document font');
      expect(canvas).not.toHaveBeenCalled();
      expect(options[0]).toMatchObject({ value: 'New document font', pending: true, measured: false });
      expect(options[0].preview).toBeUndefined();
      expect(options[0].unavailable).not.toBe(true);
      expect(options.find((option) => option.value === 'Arial')?.pending).toBe(true);
    } finally { canvas.mockRestore(); }
  });

  it('distinguishes unknown, available and missing for free-text selections', () => {
    expect(withCurrent([], 'Unknown', { preview: true, available: () => undefined })[0])
      .toMatchObject({ pending: true, measured: false });
    expect(withCurrent([], 'Available', { preview: true, available: () => true })[0])
      .toMatchObject({ preview: 'Available', measured: true });
    expect(withCurrent([], 'Missing', { preview: true, available: () => false })[0])
      .toMatchObject({ unavailable: true, measured: true });
    expect(withCurrent([], '13', { numeric: true })[0]).toEqual({ value: '13', label: '13' });
  });
});
