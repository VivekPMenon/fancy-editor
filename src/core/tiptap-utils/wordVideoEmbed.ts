// Word's "Insert Online Video" doesn't export a real embed — just a linked
// thumbnail: `<a href="https://www.youtube.com/embed/{id}?feature=oembed">`
// wrapping a base64 poster `<img>`. There's no semantic marker for "this is
// a video" anywhere else in the export, so the embed href's `/embed/{id}`
// shape is the only signal we have that it's a video and not just a linked
// picture — Word generates this deterministically (it fetches the poster via
// YouTube's oEmbed API), but it's still a pattern-match on export shape, not
// a documented contract.
const YOUTUBE_EMBED_HREF = /(?:youtube(?:-nocookie)?\.com)\/embed\/([\w-]+)/i;

/** Rewrites Word's linked-thumbnail video markup into the `<div data-youtube-video><iframe>` shape the Youtube extension's parseHTML expects, so it becomes a real playable node instead of a static image. */
export function resolveWordVideoEmbeds(html: string): string {
  if (!html.includes('/embed/')) {
    return html;
  }

  const container = document.createElement('div');
  container.innerHTML = html;

  container.querySelectorAll('a[href*="/embed/"]').forEach((anchor) => {
    const href = anchor.getAttribute('href') ?? '';
    const match = YOUTUBE_EMBED_HREF.exec(href);
    if (!match) {
      return;
    }
    const videoId = match[1];
    const wrapper = document.createElement('div');
    wrapper.setAttribute('data-youtube-video', '');
    const iframe = document.createElement('iframe');
    iframe.setAttribute('src', `https://www.youtube.com/embed/${videoId}`);
    iframe.setAttribute('width', '480');
    iframe.setAttribute('height', '270');
    wrapper.appendChild(iframe);
    anchor.replaceWith(wrapper);
  });

  return container.innerHTML;
}
