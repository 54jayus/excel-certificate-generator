(function () {
  function setModeLabel() {
    var modeTag = document.getElementById('modeTag');
    var toggleBtn = document.getElementById('toggleModeBtn');
    var isStatic = document.body.classList.contains('static-mode');
    if (modeTag) {
      modeTag.textContent = isStatic ? '当前：截图模式' : '当前：动画模式';
    }
    if (toggleBtn) {
      toggleBtn.textContent = isStatic ? '切换动画模式' : '切换截图模式';
    }
  }

  function applyQueryMode() {
    var query = window.location.search || '';
    var isStatic = false;
    if (typeof URLSearchParams === 'function') {
      isStatic = new URLSearchParams(query).get('mode') === 'static';
    } else {
      isStatic = /(?:^|[?&])mode=static(?:&|$)/.test(query);
    }
    if (isStatic) {
      document.body.classList.add('static-mode');
    }
    setModeLabel();
  }

  function toggleMode() {
    document.body.classList.toggle('static-mode');
    setModeLabel();
  }

  function bindToggle() {
    var toggleBtn = document.getElementById('toggleModeBtn');
    if (!toggleBtn) return;
    toggleBtn.addEventListener('click', toggleMode);
  }

  function markImageState() {
    var images = document.querySelectorAll('[data-poster-image]');
    var i;
    var img;
    for (i = 0; i < images.length; i += 1) {
      img = images[i];
      img.addEventListener('error', function () {
        var fallback = this.nextElementSibling;
        this.hidden = true;
        this.setAttribute('alt', '图片加载失败，请替换为你的效果图');
        if (fallback && fallback.classList.contains('result-fallback')) {
          fallback.hidden = false;
        }
      });
    }
  }

  function bindCoverReveal() {
    var revealBtn = document.getElementById('coverRevealBtn');
    var revealScreenshot = document.getElementById('coverScreenshot');
    if (!revealBtn || !revealScreenshot) return;

    var revealed = false;

    revealBtn.addEventListener('click', function () {
      if (revealed) return;
      revealed = true;

      // 隐藏按钮
      revealBtn.classList.add('hidden');

      // 显示截图区域
      revealScreenshot.classList.add('visible');
    });
  }

  function bindImagePreview() {
    var trigger = document.querySelector('[data-preview-trigger]');
    var lightbox = document.getElementById('posterImageLightbox');
    var lightboxImg = document.getElementById('posterImageLightboxImg');
    var closeButtons = document.querySelectorAll('[data-preview-close]');
    if (!trigger || !lightbox || !lightboxImg) return;

    var sourceImg = trigger.querySelector('img');
    if (!sourceImg) return;

    function closePreview() {
      lightbox.hidden = true;
      document.body.style.overflow = '';
    }

    trigger.addEventListener('click', function () {
      if (!sourceImg.getAttribute('src') || sourceImg.hidden) return;
      lightboxImg.src = sourceImg.currentSrc || sourceImg.src;
      lightboxImg.alt = sourceImg.alt;
      lightbox.hidden = false;
      document.body.style.overflow = 'hidden';
    });

    for (var i = 0; i < closeButtons.length; i += 1) {
      closeButtons[i].addEventListener('click', closePreview);
    }

    document.addEventListener('keydown', function (event) {
      if (event.key === 'Escape' && !lightbox.hidden) {
        closePreview();
      }
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    applyQueryMode();
    bindToggle();
    markImageState();
    bindCoverReveal();
    bindImagePreview();
  });
})();
