// /static/wa-kit.js
// 목적: Web Awesome를 "필요한 페이지만" 로드해서 기존 style.css와 공존시키기
// - 기본: CDN(jsDelivr) 사용
// - fallback: /static/vendor/webawesome (전산망 환경 대비)
// - native.css(전역 태그 스타일링)는 기본 OFF (충돌 방지)

(function () {
  "use strict";

  const VERSION = "3.2.1";
  const CDN_BASE = `https://cdn.jsdelivr.net/npm/@awesome.me/webawesome@${VERSION}/dist-cdn`;
  const LOCAL_BASE = "/static/vendor/webawesome"; // (선택) 없으면 자동으로 CDN로 fallback

  const DEFAULTS = {
    theme: "default",
    loadUtilities: true,
    loadNative: false, // ✅ 충돌 방지: 기본 태그를 건드릴 수 있는 native.css는 OFF
    bridgeCss: "/static/wa-bridge.css?v=1",
    detectSelector: [
      "wa-button",
      "wa-dialog",
      "wa-input",
      "wa-textarea",
      "wa-select",
      "wa-checkbox",
      "wa-switch",
      "wa-radio",
      "wa-radio-group",
      "wa-dropdown",
      "wa-popup",
      "wa-tooltip",
      "wa-details",
      "wa-card",
      "wa-alert",
      "wa-badge",
      "wa-tag"
    ].join(",")
  };

  function injectStylesheet(id, href) {
    return new Promise((resolve, reject) => {
      if (document.getElementById(id)) return resolve(true);

      const link = document.createElement("link");
      link.id = id;
      link.rel = "stylesheet";
      link.href = href;

      link.onload = () => resolve(true);
      link.onerror = () => reject(new Error(`CSS 로드 실패: ${href}`));

      (document.head || document.documentElement).appendChild(link);
    });
  }

  function injectModuleScript(id, src) {
    return new Promise((resolve, reject) => {
      if (document.getElementById(id)) return resolve(true);

      const s = document.createElement("script");
      s.id = id;
      s.type = "module";
      s.src = src;

      s.onload = () => resolve(true);
      s.onerror = () => reject(new Error(`JS 로드 실패: ${src}`));

      (document.head || document.documentElement).appendChild(s);
    });
  }

  async function withFallback(loadFn, local, cdn) {
    // window.WA_ASSET_MODE = "local" | "cdn" | "auto"
    const mode = String(window.WA_ASSET_MODE || "auto").toLowerCase();

    if (mode === "local") return loadFn(local);
    if (mode === "cdn") return loadFn(cdn);

    // auto: local 먼저 시도 → 실패 시 CDN
    try {
      return await loadFn(local);
    } catch (_) {
      return await loadFn(cdn);
    }
  }

  function once(fn) {
    let p = null;
    return function (...args) {
      if (!p) p = Promise.resolve().then(() => fn(...args));
      return p;
    };
  }

  const load = once(async function (opts) {
    const options = Object.assign({}, DEFAULTS, opts || {});
    const themeName = options.theme || "default";

    // 1) Theme (필수)
    await withFallback(
      href => injectStylesheet("wa-theme-css", href),
      `${LOCAL_BASE}/styles/themes/${themeName}.css`,
      `${CDN_BASE}/styles/themes/${themeName}.css`
    );

    // 2) Utilities (선택)
    if (options.loadUtilities) {
      await withFallback(
        href => injectStylesheet("wa-utilities-css", href),
        `${LOCAL_BASE}/styles/utilities.css`,
        `${CDN_BASE}/styles/utilities.css`
      );
    }

    // 3) Native (선택 / 기본 OFF)
    if (options.loadNative) {
      await withFallback(
        href => injectStylesheet("wa-native-css", href),
        `${LOCAL_BASE}/styles/native.css`,
        `${CDN_BASE}/styles/native.css`
      );
    }

    // 4) Bridge (우리 사이트 톤에 맞추는 토큰 오버라이드)
    if (options.bridgeCss) {
      await injectStylesheet("wa-bridge-css", options.bridgeCss);
    }

    // 5) Loader (컴포넌트 자동 로드)
    await withFallback(
      src => injectModuleScript("wa-loader-js", src),
      `${LOCAL_BASE}/webawesome.loader.js`,
      `${CDN_BASE}/webawesome.loader.js`
    );

    return true;
  });

  function needs() {
    try {
      if (document.documentElement?.dataset?.uiKit === "wa") return true;
      if (document.body?.dataset?.uiKit === "wa") return true;
      if (document.querySelector(DEFAULTS.detectSelector)) return true;
    } catch (_) {}
    return false;
  }

  function auto() {
    return needs() ? load() : Promise.resolve(false);
  }

  // 전역 공개 (정적 페이지에서 쓰기 편하게)
  window.WAKit = {
    VERSION,
    load,
    auto,
    needs
  };
})();
