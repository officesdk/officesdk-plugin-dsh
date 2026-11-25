import zh from '../../locale/zh.json' with { type: 'json' };

const [intro, capabilities, example] = zh.meta.description.split('\n\n');
const [heading, details] = capabilities.split('\n');
const parts = details.split(/(\d+ 项功能)/);

/** The host renders package metadata as plain text. Use the detail slot to format the same copy in place. */
export function registerPackageDescription(ctx, React, createPortal) {
  if (!createPortal) return;
  const { createElement: h, useLayoutEffect, useRef, useState } = React;
  function Description({ subject, t }) {
    const marker = useRef(null);
    const [header, setHeader] = useState(null);
    const visible = subject?.kind === 'bundle' && ['@officesdk/officesdk-plugin-dsh', 'officesdk'].includes(subject.pkg.name) && t?.('lang') === 'zh';
    useLayoutEffect(() => {
      const detail = marker.current?.closest('[data-plugin-detail]');
      const name = detail?.querySelector('code[data-plugin-name]')?.parentElement;
      setHeader(name?.nextElementSibling?.textContent === zh.meta.description ? name.parentElement : null);
    }, [visible, subject?.pkg?.name]);
    if (!visible) return null;
    return h(React.Fragment, null, h('span', { hidden: true, ref: marker }),
      header && createPortal(h('div', { className: 'officesdk-package-description' },
        h('p', null, intro), h('h4', null, heading),
        h('p', null, parts.map((part, index) => index % 2 ? h('strong', { key: index }, part) : part)),
        h('p', null, example)), header));
  }
  ctx.effect(() => ctx.slots.inject('plugins.detail.badge', () => ctx.slots.register({
    name: 'plugins.detail.badge', id: 'officesdk-package-description', locale: 'shimoPreview',
  }, Description)));
}
