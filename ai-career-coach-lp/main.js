/**
 * AIキャリアコーチ LP - スクロール表示アニメーション
 */
(function () {
  'use strict';

  // スクロールで .reveal 要素を表示
  function initReveal() {
    var reveals = document.querySelectorAll('.reveal');
    if (!reveals.length) return;

    function checkReveal() {
      var windowHeight = window.innerHeight;
      var revealPoint = 120;

      reveals.forEach(function (el) {
        var top = el.getBoundingClientRect().top;
        if (top < windowHeight - revealPoint) {
          el.classList.add('is-visible');
        }
      });
    }

    window.addEventListener('scroll', checkReveal);
    window.addEventListener('resize', checkReveal);
    checkReveal(); // 初回実行
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initReveal);
  } else {
    initReveal();
  }
})();
