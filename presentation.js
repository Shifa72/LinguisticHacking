(() => {
  const slides = [...document.querySelectorAll(".slide")];
  const bar = document.querySelector(".progress span");
  const counter = document.querySelector("[data-counter]");
  let i = 0;
  let decryptTimers = [];
  const glyphs = "*#АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ";

  const clearDecrypt = () => {
    decryptTimers.forEach((id) => clearTimeout(id));
    decryptTimers = [];
  };

  const runDecrypt = (slide) => {
    clearDecrypt();
    slide.querySelectorAll(".code-word").forEach((el, w) => {
      const word = el.dataset.word || "";
      el.classList.add("is-decoding");
      el.innerHTML = [...word].map(() => '<span class="ch star">*</span>').join("");
      [...el.children].forEach((span, i) => {
        const start = 180 + w * 480 + i * 95;
        for (let k = 0; k < 6; k++) {
          decryptTimers.push(setTimeout(() => {
            span.textContent = glyphs[Math.floor(Math.random() * glyphs.length)];
          }, start + k * 42));
        }
        decryptTimers.push(setTimeout(() => {
          span.textContent = word[i];
          span.classList.remove("star");
          span.classList.add("lock");
        }, start + 290));
      });
      decryptTimers.push(setTimeout(() => {
        el.classList.remove("is-decoding");
      }, 180 + w * 480 + word.length * 95 + 360));
    });
  };

  const hashIndex = () => {
    const n = Number(location.hash.replace("#", ""));
    return Number.isInteger(n) && n >= 1 && n <= slides.length ? n - 1 : 0;
  };

  let voyageStep = 0;
  let readingStep = 0;

  const resetVoyage = () => {
    voyageStep = 0;
    slides.forEach((slide) => slide.classList.remove("pack-left", "pack-right"));
  };

  const resetReading = () => {
    readingStep = 0;
    slides.forEach((slide) => slide.classList.remove("reveal-extra"));
  };

  const go = (n, push = true) => {
    i = Math.max(0, Math.min(slides.length - 1, n));
    resetVoyage();
    resetReading();
    slides.forEach((slide, idx) => slide.classList.toggle("active", idx === i));
    bar.style.width = `${((i + 1) / slides.length) * 100}%`;
    counter.textContent = `${i + 1} / ${slides.length}`;
    const dark = slides[i].classList.contains("slide-title")
      || slides[i].classList.contains("slide-finale")
      || slides[i].classList.contains("slide-dark")
      || slides[i].classList.contains("slide-decode");
    const stage = document.querySelector(".stage");
    stage.classList.toggle("is-dark", dark);
    stage.classList.toggle("is-first", i === 0);
    slides.forEach((slide) => slide.classList.remove("play"));
    void slides[i].offsetWidth;
    slides[i].classList.add("play");
    if (slides[i].classList.contains("slide-decode")) runDecrypt(slides[i]);
    else clearDecrypt();
    if (push) history.replaceState(null, "", `#${i + 1}`);
  };

  const next = () => {
    const slide = slides[i];
    if (slide.classList.contains("slide-voyage")) {
      if (voyageStep === 0) {
        slide.classList.add("pack-left");
        voyageStep = 1;
        return;
      }
      if (voyageStep === 1) {
        slide.classList.add("pack-right");
        voyageStep = 2;
        return;
      }
    }
    if (slide.classList.contains("slide-reading")) {
      if (readingStep === 0) {
        slide.classList.add("reveal-extra");
        readingStep = 1;
        return;
      }
    }
    go(i + 1);
  };

  const prev = () => {
    const slide = slides[i];
    if (slide.classList.contains("slide-voyage")) {
      if (voyageStep === 2) {
        slide.classList.remove("pack-right");
        voyageStep = 1;
        return;
      }
      if (voyageStep === 1) {
        slide.classList.remove("pack-left");
        voyageStep = 0;
        return;
      }
    }
    if (slide.classList.contains("slide-reading")) {
      if (readingStep === 1) {
        slide.classList.remove("reveal-extra");
        readingStep = 0;
        return;
      }
    }
    go(i - 1);
  };

  document.addEventListener("keydown", (e) => {
    const nextKeys = ["ArrowRight", "ArrowDown", "PageDown", " ", "Enter"];
    const prevKeys = ["ArrowLeft", "ArrowUp", "PageUp", "Backspace"];
    if (nextKeys.includes(e.key)) {
      e.preventDefault();
      next();
    } else if (prevKeys.includes(e.key)) {
      e.preventDefault();
      prev();
    } else if (e.key === "Home") go(0);
    else if (e.key === "End") go(slides.length - 1);
    else if (e.key === "f" || e.key === "F") {
      if (!document.fullscreenElement) document.documentElement.requestFullscreen();
      else document.exitFullscreen();
    }
  });

  let touchX = null;
  document.addEventListener("touchstart", (e) => {
    touchX = e.changedTouches[0].screenX;
  }, { passive: true });
  document.addEventListener("touchend", (e) => {
    if (touchX == null) return;
    const dx = e.changedTouches[0].screenX - touchX;
    if (Math.abs(dx) > 50) dx < 0 ? next() : prev();
    touchX = null;
  }, { passive: true });

  document.addEventListener("click", (e) => {
    if (e.target.closest("a, button")) return;
    if (window.matchMedia("(pointer: coarse)").matches) return;
    const mid = window.innerWidth / 2;
    if (e.clientX >= mid) next();
    else prev();
  });

  window.addEventListener("hashchange", () => go(hashIndex(), false));
  go(hashIndex(), false);
})();
