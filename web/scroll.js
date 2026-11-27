// px/sec for each speed level (0=off, 1=slow, 2=medium, 3=fast)
const AUTO_SCROLL_SPEEDS = [0, 20, 50, 120];
const AUTO_SCROLL_LABELS = ["Auto-Scroll", "Slow Scroll", "Medium Scroll", "Fast Scroll"];

export function createScrollController(state, ui, { getFrameIframes, syncAdvancedTools }) {
  let _autoScrollLast = null;

  function applyAutoScrollIframes() {
    const containerH = ui.frame.clientHeight || 400;
    getFrameIframes().forEach(iframe => {
      iframe.style.height = `${containerH * 3}px`;
    });
  }

  function resetAutoScrollIframes() {
    getFrameIframes().forEach(iframe => {
      iframe.style.height = "";
      iframe.style.transform = "";
      iframe.style.transformOrigin = "";
      iframe.style.opacity = "";
      iframe.style.transition = "";
    });
  }

  function stopAutoScrollAnimation() {
    if (state.autoScrollRafId) {
      cancelAnimationFrame(state.autoScrollRafId);
      state.autoScrollRafId = null;
    }
    _autoScrollLast = null;
  }

  function autoScrollTick(timestamp) {
    if (!_autoScrollLast) _autoScrollLast = timestamp;
    const delta = (timestamp - _autoScrollLast) / 1000;
    _autoScrollLast = timestamp;

    const speed = AUTO_SCROLL_SPEEDS[state.autoScrollSpeed] || 0;
    if (speed > 0) {
      const containerH = ui.frame.clientHeight || 400;
      const maxScroll = containerH * 2;
      const newPos = (state.autoScrollPos + speed * delta) % maxScroll;
      const wrapped = newPos < state.autoScrollPos;
      state.autoScrollPos = newPos;

      const iframes = getFrameIframes();
      iframes.forEach(iframe => {
        iframe.style.transformOrigin = "top left";
        iframe.style.transform = `translateY(-${state.autoScrollPos}px)`;
      });

      // On wrap, fade out instantly then fade back in so the jump reads as an
      // intentional reset rather than a visual glitch.
      if (wrapped) {
        iframes.forEach(f => {
          f.style.transition = "none";
          f.style.opacity = "0";
          requestAnimationFrame(() => {
            f.style.transition = "opacity 0.5s ease";
            f.style.opacity = "1";
          });
        });
      }
    }

    if (state.autoScrollSpeed > 0) {
      state.autoScrollRafId = requestAnimationFrame(autoScrollTick);
    } else {
      state.autoScrollRafId = null;
    }
  }

  function startAutoScrollAnimation() {
    if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    stopAutoScrollAnimation();
    _autoScrollLast = null;
    state.autoScrollRafId = requestAnimationFrame(autoScrollTick);
  }

  function toggleAutoScroll() {
    const span = ui.autoScrollBtn.querySelector("span");

    if (state.autoScrollSpeed > 0) {
      state.autoScrollSpeed = 0;
      if (span) span.textContent = AUTO_SCROLL_LABELS[0];
      stopAutoScrollAnimation();
      state.autoScrollPos = 0;
      resetAutoScrollIframes();
    } else {
      state.autoScrollSpeed = 1;
      if (span) span.textContent = AUTO_SCROLL_LABELS[1];
      applyAutoScrollIframes();
      startAutoScrollAnimation();
    }

    syncAdvancedTools();
  }

  // Convenience: stop + zero state + reset label + reset iframes in one call.
  // Used by toggleDesktopFit and toggleMobileView for mutual-exclusivity.
  function stopAndReset() {
    stopAutoScrollAnimation();
    state.autoScrollSpeed = 0;
    state.autoScrollPos = 0;
    const span = ui.autoScrollBtn.querySelector("span");
    if (span) span.textContent = AUTO_SCROLL_LABELS[0];
    resetAutoScrollIframes();
  }

  return {
    toggleAutoScroll,
    applyAutoScrollIframes,
    resetAutoScrollIframes,
    startAutoScrollAnimation,
    stopAutoScrollAnimation,
    stopAndReset,
  };
}
