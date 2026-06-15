// Cloudflare Pages Advanced-Mode worker.
// Routes each state domain (Host header) to its advertorial in /states/.
// Add a state = drop states/<code>.html and add one line to STATE_SITES.

const STATE_SITES = {
  "texasplancheck.org":         "tx.html",
  "northcarolinaplancheck.org": "nc.html",
  "floridaplancheck.org":       "fl.html",
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    let host = url.hostname.toLowerCase();
    if (host.startsWith("www.")) host = host.slice(4);

    // Static assets (hero image, css, etc.) — serve directly from the output dir.
    if (/\.(jpg|jpeg|png|gif|svg|webp|ico|css|js|woff2?)$/i.test(url.pathname)) {
      return env.ASSETS.fetch(request);
    }

    // Page request → serve the matching state's advertorial.
    const file = STATE_SITES[host];
    if (file) {
      return env.ASSETS.fetch(new URL(`/states/${file}`, url.origin));
    }

    // Unknown host → fall through to default asset handling (404 if none).
    return env.ASSETS.fetch(request);
  },
};
