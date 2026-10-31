/**
 * `a:normAutofit` with no `fontScale` asks the renderer to shrink the text until it fits its
 * box ("shrink text on overflow"). The file states no percentage, so the page measures: the
 * body is drawn at the sizes the file states, and only a body that overflows is scaled down
 * until its content fits. Bodies that already fit keep their stated sizes.
 *
 * Every element that carries its own inline size is scaled; runs that inherit their size
 * follow the body. The search is a bisection on one factor applied to all of them, so runs
 * keep their size ratios — the same shape a uniform shrink has in WPS.
 */
export function fitShrunkText(root: ParentNode): void {
  for (const body of root.querySelectorAll<HTMLElement>('.shape-text[data-shrink="1"]')) {
    const sized = [body, ...body.querySelectorAll<HTMLElement>("*")].filter(
      (el) => el.style.fontSize !== "" || el === body,
    );
    for (const el of sized) {
      // The stated size is the first inline value seen; later fits (a re-render, a resize)
      // must scale that one, never the already-scaled result of the pass before.
      el.dataset.fitBase ??= el.style.fontSize || getComputedStyle(el).fontSize;
      el.style.fontSize = el.dataset.fitBase;
    }
    if (!overflows(body)) continue;
    let lo = 0.1;
    let hi = 1;
    for (let i = 0; i < 6; i++) {
      const mid = (lo + hi) / 2;
      scale(sized, mid);
      if (overflows(body)) hi = mid;
      else lo = mid;
    }
    scale(sized, lo);
  }
}

function scale(sized: HTMLElement[], factor: number): void {
  for (const el of sized) {
    const base = el.dataset.fitBase ?? "";
    const match = /^([\d.]+)([a-z%]*)$/.exec(base);
    if (!match) continue;
    const value = Math.round(Number(match[1]) * factor * 100) / 100;
    el.style.fontSize = `${value}${match[2] || "px"}`;
  }
}

function overflows(body: HTMLElement): boolean {
  return body.scrollHeight > body.clientHeight + 0.5;
}
