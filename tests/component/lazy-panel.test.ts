import { afterEach, describe, expect, it, vi } from 'vitest';
import { defineComponent, h, onUnmounted, watch, type Component } from 'vue';
import { flushPromises, mount } from '@vue/test-utils';
import { lazyPanel } from '../../src/ui/panels/lazy-panel';
import OpenDocumentDialog from '../../src/ui/panels/OpenDocumentDialog.vue';

const cleanups: Array<() => void> = [];
afterEach(() => {
  cleanups.splice(0).reverse().forEach((cleanup) => cleanup());
  vi.restoreAllMocks();
});

function outsideInput(): HTMLInputElement {
  const input = document.createElement('input');
  document.body.appendChild(input);
  cleanups.push(() => input.remove());
  return input;
}

describe('lazy panel', () => {
  it('does no work while closed and preserves first-open, close and reopen transitions', async () => {
    const transitions: boolean[] = [];
    const unmount = vi.fn();
    const panel = defineComponent({
      props: { isOpen: Boolean, label: String }, emits: ['close', 'save'],
      setup(props, { emit }) {
        watch(() => props.isOpen, (open) => transitions.push(open));
        onUnmounted(unmount);
        return () => props.isOpen ? h('button', { onClick: () => emit('save', props.label) }, props.label) : null;
      },
    });
    const loader = vi.fn(async () => ({ default: panel }));
    const save = vi.fn();
    const wrapper = mount(lazyPanel(loader), { props: { isOpen: false, label: 'first', onSave: save } });
    expect(loader).not.toHaveBeenCalled();
    await wrapper.setProps({ isOpen: true });
    await flushPromises();
    expect(transitions).toEqual([true]);
    await wrapper.get('button').trigger('click');
    expect(save).toHaveBeenCalledWith('first');
    await wrapper.setProps({ isOpen: false });
    expect(transitions).toEqual([true, false]);
    expect(unmount).not.toHaveBeenCalled();
    await wrapper.setProps({ isOpen: true, label: 'second' });
    expect(wrapper.text()).toBe('second');
    expect(transitions).toEqual([true, false, true]);
    expect(loader).toHaveBeenCalledTimes(1);
    wrapper.unmount();
    expect(unmount).toHaveBeenCalledTimes(1);
  });

  it('shows a recoverable error and retries the same requested panel', async () => {
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    const panel = defineComponent({ render: () => h('span', 'panel ready') });
    const loader = vi.fn().mockRejectedValueOnce(new Error('missing chunk')).mockResolvedValue({ default: panel });
    const wrapper = mount(lazyPanel(loader), { props: { isOpen: true } });
    await flushPromises();
    expect(wrapper.get('[role="alert"]').text()).toContain('לא ניתן לפתוח');
    await wrapper.get('button').trigger('click');
    await flushPromises();
    expect(wrapper.text()).toBe('panel ready');
    expect(loader).toHaveBeenCalledTimes(2);
    wrapper.unmount();
  });

  it('forwards cancel while pending and never opens a panel after it was closed', async () => {
    let resolve!: (module: { default: Component }) => void;
    const loader = vi.fn(() => new Promise<{ default: Component }>((done) => { resolve = done; }));
    const visible: boolean[] = [];
    const panel = defineComponent({ props: { isOpen: Boolean }, setup(props) {
      watch(() => props.isOpen, (open) => visible.push(open), { immediate: true });
      return () => props.isOpen ? h('span', 'open') : null;
    } });
    const wrapper = mount(lazyPanel(loader), { props: { isOpen: true } });
    expect(wrapper.get('[role="status"]').text()).toContain('טוען');
    await wrapper.get('button').trigger('click');
    expect(wrapper.emitted('close')).toHaveLength(1);
    await wrapper.setProps({ isOpen: false });
    resolve({ default: panel });
    await flushPromises();
    expect(visible).toEqual([false]);
    expect(wrapper.text()).toBe('');
    wrapper.unmount();
  });

  it('does not mount a completed request after disposal', async () => {
    let resolve!: (module: { default: Component }) => void;
    const loader = () => new Promise<{ default: Component }>((done) => { resolve = done; });
    const setup = vi.fn(() => () => h('span', 'late'));
    const wrapper = mount(lazyPanel(loader), { props: { isOpen: true } });
    wrapper.unmount();
    resolve({ default: defineComponent({ setup }) });
    await flushPromises();
    expect(setup).not.toHaveBeenCalled();
  });

  it.each([false, true])('returns focus after a real Open dialog load failure and retry (moved outside: %s)', async (movedOutside) => {
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    const origin = outsideInput();
    const other = outsideInput();
    origin.focus();
    const loader = vi.fn().mockRejectedValueOnce(new Error('missing chunk')).mockResolvedValue({ default: OpenDocumentDialog });
    const wrapper = mount(lazyPanel(loader), { attachTo: document.body, props: { isOpen: true, intent: 'open' } });
    cleanups.push(() => wrapper.unmount());
    await flushPromises();
    if (movedOutside) other.focus();
    const retry = wrapper.get<HTMLButtonElement>('[role="alert"] button');
    retry.element.focus();
    expect(document.activeElement).toBe(retry.element);
    await retry.trigger('click');
    await flushPromises();
    expect(wrapper.get('.open-dialog').element.contains(document.activeElement)).toBe(true);
    await wrapper.setProps({ isOpen: false });
    await flushPromises();
    expect(document.activeElement).toBe(movedOutside ? other : origin);
  });

  it('restores focus from pending Cancel and stays closed after the chunk arrives', async () => {
    const origin = outsideInput();
    origin.focus();
    let resolve!: (module: { default: Component }) => void;
    const wrapper = mount(lazyPanel(() => new Promise((done) => { resolve = done; })), {
      attachTo: document.body, props: { isOpen: true, intent: 'open' },
    });
    cleanups.push(() => wrapper.unmount());
    const cancel = wrapper.get<HTMLButtonElement>('button');
    cancel.element.focus();
    await cancel.trigger('click');
    expect(wrapper.emitted('close')).toHaveLength(1);
    expect(document.activeElement).toBe(origin);
    await wrapper.setProps({ isOpen: false });
    resolve({ default: OpenDocumentDialog });
    await flushPromises();
    expect(wrapper.find('.open-dialog').exists()).toBe(false);
    expect(document.activeElement).toBe(origin);
  });

  it('does not steal focus when the shell dismisses a pending panel after focus moved elsewhere', async () => {
    const origin = outsideInput();
    const other = outsideInput();
    origin.focus();
    let resolve!: (module: { default: Component }) => void;
    const wrapper = mount(lazyPanel(() => new Promise((done) => { resolve = done; })), {
      attachTo: document.body, props: { isOpen: true },
    });
    cleanups.push(() => wrapper.unmount());
    other.focus();
    await wrapper.setProps({ isOpen: false });
    expect(document.activeElement).toBe(other);
    resolve({ default: OpenDocumentDialog });
    await flushPromises();
    expect(wrapper.find('.open-dialog').exists()).toBe(false);
    expect(document.activeElement).toBe(other);
  });
});
