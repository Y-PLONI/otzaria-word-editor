import { defineComponent, h, nextTick, onUnmounted, shallowRef, watch, type Component } from 'vue';
import PanelLoadState from './PanelLoadState.vue';

/**
 * Load a panel only on its first open. Keep its instance afterwards: close
 * watchers, focus restoration and any draft state keep their existing behavior.
 * Commands/shortcuts remain in the shell; only the panel's view is deferred.
 */
export function lazyPanel(loader: () => Promise<{ default: Component }>): Component {
  let resolved: Component | undefined;
  let pending: Promise<Component> | undefined;

  function load(): Promise<Component> {
    if (resolved) return Promise.resolve(resolved);
    pending ??= loader().then((module) => (resolved = module.default)).finally(() => { pending = undefined; });
    return pending;
  }

  return defineComponent({
    name: 'LazyPanel',
    inheritAttrs: false,
    props: { isOpen: Boolean },
    emits: ['close'],
    setup(props, { attrs, emit }) {
      const component = shallowRef<Component>();
      const ready = shallowRef(false);
      const error = shallowRef(false);
      const fallback = shallowRef<{ contains(node: Node | null): boolean } | null>(null);
      let focusOutside: HTMLElement | null = null;
      let trackingFocus = false;
      let loading = false;
      let disposed = false;

      function rememberOutsideFocus(event: FocusEvent): void {
        const target = event.target;
        if (target instanceof HTMLElement && !fallback.value?.contains(target)) {
          focusOutside = target === document.body ? null : target;
        }
      }
      function startTrackingFocus(): void {
        if (trackingFocus) return;
        const active = document.activeElement;
        focusOutside = active instanceof HTMLElement && active !== document.body ? active : null;
        document.addEventListener('focusin', rememberOutsideFocus);
        trackingFocus = true;
      }
      function stopTrackingFocus(): void {
        document.removeEventListener('focusin', rememberOutsideFocus);
        trackingFocus = false;
        focusOutside = null;
      }
      function restoreFallbackFocus(): void {
        // Retry/Cancel are temporary controls. Preserve the most recent focus
        // outside this view before removing them; never steal focus from a user
        // who has already moved elsewhere while the chunk was loading.
        if (fallback.value?.contains(document.activeElement) && focusOutside?.isConnected) focusOutside.focus();
      }
      onUnmounted(() => { disposed = true; stopTrackingFocus(); });

      async function open(): Promise<void> {
        if (loading || component.value || disposed) return;
        loading = true;
        error.value = false;
        try {
          const panel = await load();
          if (disposed) return;
          restoreFallbackFocus();
          stopTrackingFocus();
          component.value = panel;
          // Existing panels watch a false → true transition to initialize their
          // form/focus. Mount them closed once, then open on the following tick.
          await nextTick();
          if (!disposed) ready.value = true;
        } catch (cause) {
          if (!disposed) {
            error.value = true;
            console.warn('[otzaria-word] Panel load failed:', cause);
          }
        } finally {
          loading = false;
        }
      }
      watch(() => props.isOpen, (visible) => {
        if (visible && !component.value) {
          startTrackingFocus();
          void open();
        } else if (!visible) {
          restoreFallbackFocus();
          stopTrackingFocus();
        }
      }, { immediate: true });
      const close = () => { restoreFallbackFocus(); emit('close'); };
      return () => component.value
        ? h(component.value, { ...attrs, isOpen: ready.value && props.isOpen, onClose: close })
        : props.isOpen
          ? h(PanelLoadState, { ref: fallback, failed: error.value, onRetry: open, onClose: close })
          : null;
    },
  });
}
