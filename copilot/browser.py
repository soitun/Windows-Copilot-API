"""Browser-backed Copilot driver.

A Playwright fallback for the pure-HTTP :class:`copilot.client.Copilot`: it runs
the *exact same protocol* inside a real browser that already holds Cloudflare
clearance and (optionally) a signed-in Microsoft session. Useful if Microsoft
ever escalates the challenge to a Cloudflare Turnstile CAPTCHA, which needs a
browser-solved token.

``BrowserCopilot`` launches a **persistent** Playwright Chromium profile so that
Cloudflare clearance and any sign-in survive restarts. The chat protocol
(``POST /c/api/conversations`` then a ``wss://.../c/api/chat`` WebSocket speaking
``send`` -> ``appendText``* -> ``done``) is executed *in the page* via
``page.evaluate`` so the browser's own ``fetch``/``WebSocket`` carry the cookies,
Cloudflare token, and auth headers.

It exposes the same ``create_completion(prompt, stream=...)`` generator API as
:class:`copilot.client.Copilot`, so it is a drop-in replacement.

PROTOCOL ASSUMPTIONS (verify at runtime against a live session):
  * Conversation create:  POST /c/api/conversations  -> {"id": "..."}
  * Chat socket:          wss://copilot.microsoft.com/c/api/chat?api-version=2
                          &clientSessionId=<uuid> (with &accessToken=<token> when
                          signed in)
  * Preamble:             {"event":"setOptions",...} then
                          {"event":"reportLocalConsents","grantedConsents":[]}
  * Send frame:           {"event":"send","conversationId":...,
                           "content":[{"type":"text","text":...}],"mode":"smart",
                           "context":{}}
  * Stream frames:        an empty {"event":"challenge","method":null} (ignored),
                          then {"event":"appendText","text":...} and {"event":"done"}
These mirror the protocol in ``driver.py`` (re-capture with ``python -m copilot
capture`` if Microsoft changes it, and adjust the JS templates below).
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Dict, Generator, Optional

from playwright.sync_api import sync_playwright, Error as PlaywrightError

from .auth import DEFAULT_AUTH_FILE, DEFAULT_PROFILE_DIR, SESSION_DIR

COPILOT_URL = "https://copilot.microsoft.com/"

# --- in-page JavaScript -----------------------------------------------------

# Create a conversation. Runs in the page so cookies/Cloudflare apply.
_CREATE_CONVERSATION_JS = """
async () => {
  const res = await fetch('/c/api/conversations', {
    method: 'POST',
    credentials: 'include',
    headers: {'content-type': 'application/json'},
  });
  const text = await res.text();
  if (!res.ok) return {ok: false, status: res.status, text: text};
  let data = {};
  try { data = JSON.parse(text); } catch (e) {}
  return {ok: true, id: data.id || data.conversationId || null, raw: text};
}
"""

# Discover the Copilot chat MSAL access token from localStorage. The cache holds
# several tokens for different scopes; the chat WebSocket only accepts the one
# scoped 'ChatAI.ReadWrite' — a wrong-audience token (e.g. the Graph
# User.Read/Files.Read token) makes the WS upgrade 401. We therefore PREFER the
# ChatAI token and only fall back to the first token found if none matches.
# Returns null for anonymous sessions (anonymous chat may still work via cookies).
_FIND_TOKEN_JS = """
() => {
  try {
    let fallback = null;
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      const v = localStorage.getItem(k);
      if (v && v.indexOf('"credentialType":"AccessToken"') !== -1) {
        try {
          const o = JSON.parse(v);
          if (o && o.secret) {
            // Match the chat scope (e.g. '<resource>/ChatAI.ReadWrite'); take the
            // first non-matching token only as a last-resort fallback.
            if (o.target && o.target.indexOf('ChatAI') !== -1) return o.secret;
            if (!fallback) fallback = o.secret;
          }
        } catch (e) {}
      }
    }
    return fallback;
  } catch (e) {}
  return null;
}
"""

# Open the chat WebSocket and wire handlers that push into a window-scoped
# buffer. Returns immediately; messages accumulate while Python polls.
_START_STREAM_JS = """
([conversationId, accessToken, prompt]) => {
  const state = {queue: [], done: false, error: null, started: false};
  window.__copilot = state;
  // Mirror the live web client's handshake (captured via `python -m copilot
  // capture`): a per-session clientSessionId, a setOptions/reportLocalConsents
  // preamble, then the send frame with mode 'smart'.
  const clientSessionId = (crypto.randomUUID ? crypto.randomUUID() :
    'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
      const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    }));
  let url = 'wss://copilot.microsoft.com/c/api/chat?api-version=2&clientSessionId=' + clientSessionId;
  if (accessToken) url += '&accessToken=' + encodeURIComponent(accessToken);
  let ws;
  try { ws = new WebSocket(url); } catch (e) { state.error = 'ws-init: ' + e; state.done = true; return false; }
  window.__copilotWs = ws;
  ws.onopen = () => {
    ws.send(JSON.stringify({
      event: 'setOptions',
      supportedFeatures: ['partial-generated-images', 'composer-prefill-conversation-action',
        'composer-send-conversation-action-v2', 'side-by-side-comparison',
        'session-duration-nudge', 'compose-email-html'],
      supportedCards: [],
      supportedActions: []
    }));
    ws.send(JSON.stringify({event: 'reportLocalConsents', grantedConsents: []}));
    ws.send(JSON.stringify({
      event: 'send',
      conversationId: conversationId,
      content: [{type: 'text', text: prompt}],
      mode: 'smart',
      context: {}
    }));
  };
  ws.onmessage = (ev) => {
    let msg;
    try { msg = JSON.parse(ev.data); } catch (e) { return; }
    const e = msg.event;
    // An empty challenge (method/parameter null) is informational and ignored;
    // only appendText/done/error matter here.
    if (e === 'appendText') { state.started = true; if (msg.text) state.queue.push(msg.text); }
    else if (e === 'done') { state.done = true; try { ws.close(); } catch (x) {} }
    else if (e === 'error') { state.error = JSON.stringify(msg); state.done = true; try { ws.close(); } catch (x) {} }
  };
  ws.onerror = () => { state.error = state.error || 'websocket error'; state.done = true; };
  ws.onclose = () => { state.done = true; };
  return true;
}
"""

# Drain the buffer and report status in one round-trip.
_POLL_JS = """
() => {
  const s = window.__copilot || {queue: [], done: true, error: 'not started', started: false};
  const q = s.queue;
  s.queue = [];
  return {q: q, done: s.done, error: s.error, started: s.started};
}
"""


class BrowserCopilot:
    """Drives Microsoft Copilot through a real Playwright browser.

    Parameters
    ----------
    profile_dir:
        Directory for the persistent Chromium profile (cookies, Cloudflare
        clearance, sign-in). Reused across runs.
    headless:
        Run without a visible window. Use ``False`` (or :meth:`login`) for the
        first interactive sign-in, then ``True`` afterwards.
    """

    label = "Microsoft Copilot (browser)"
    default_model = "Copilot"

    def __init__(
        self,
        profile_dir: str = DEFAULT_PROFILE_DIR,
        headless: bool = True,
        nav_timeout: int = 60,
        proxy: Optional[str] = None,
    ):
        self.profile_dir = str(Path(profile_dir).resolve())
        self.headless = headless
        self.nav_timeout = nav_timeout
        # Copilot consumer chat is geo-restricted. If you are outside a supported
        # region, route the browser through a proxy/VPN in a supported region,
        # e.g. proxy="http://user:pass@host:port" or "socks5://host:port".
        self.proxy = proxy

        self._pw = None
        self._context = None
        self._page = None

    # -- lifecycle ----------------------------------------------------------

    def start(self, headless: Optional[bool] = None) -> "BrowserCopilot":
        """Launch the persistent browser context and open Copilot."""
        if self._context is not None:
            return self
        if headless is not None:
            self.headless = headless
        try:
            self._pw = sync_playwright().start()
            launch_kwargs = dict(
                headless=self.headless,
                args=["--disable-blink-features=AutomationControlled"],
            )
            if self.proxy:
                launch_kwargs["proxy"] = self._parse_proxy(self.proxy)
            self._context = self._pw.chromium.launch_persistent_context(
                self.profile_dir,
                **launch_kwargs,
            )
            self._page = self._context.pages[0] if self._context.pages else self._context.new_page()
            self._page.set_default_timeout(self.nav_timeout * 1000)
            self._page.goto(COPILOT_URL, wait_until="domcontentloaded")
            # Give Cloudflare a moment to clear on first paint.
            self._page.wait_for_load_state("networkidle", timeout=self.nav_timeout * 1000)
        except PlaywrightError as exc:
            self.close()
            raise ConnectionError(f"Failed to start browser: {exc}") from exc
        return self

    @staticmethod
    def _parse_proxy(proxy: str) -> dict:
        """Turn a ``scheme://user:pass@host:port`` string into Playwright form."""
        from urllib.parse import urlparse

        u = urlparse(proxy)
        server = f"{u.scheme}://{u.hostname}:{u.port}" if u.port else f"{u.scheme}://{u.hostname}"
        cfg = {"server": server}
        if u.username:
            cfg["username"] = u.username
        if u.password:
            cfg["password"] = u.password
        return cfg

    def region_blocked(self) -> bool:
        """True if Copilot is showing the 'Not available in your region' notice."""
        if self._page is None:
            return False
        try:
            text = self._page.evaluate("() => document.body ? document.body.innerText : ''")
        except PlaywrightError:
            return False
        return "available in your region" in (text or "").lower()

    def close(self) -> None:
        for attr, closer in (("_context", lambda c: c.close()), ("_pw", lambda p: p.stop())):
            obj = getattr(self, attr, None)
            if obj is not None:
                try:
                    closer(obj)
                except Exception:
                    pass
                setattr(self, attr, None)
        self._page = None

    def __enter__(self) -> "BrowserCopilot":
        return self.start()

    def __exit__(self, *exc) -> None:
        self.close()

    # -- auth ---------------------------------------------------------------

    def login(self, path: str = DEFAULT_AUTH_FILE) -> dict:
        """Open a visible window for interactive Microsoft sign-in.

        Blocks until you press Enter in the console. The session is persisted in
        ``profile_dir`` (and snapshotted to ``path``), so subsequent headless
        runs reuse it. Returns the captured auth dict.
        """
        self.close()
        self.start(headless=False)
        print(
            "\nA browser window is open at copilot.microsoft.com.\n"
            "Sign in (or just solve any Cloudflare check for anonymous use),\n"
            "then return here and press Enter to save the session..."
        )
        try:
            input()
        except EOFError:
            pass
        # Snapshot fresh auth so the headless curl_cffi path works immediately.
        auth: dict = {}
        try:
            auth = self.export_auth(path=path, stamp=time.time())
            print(f"Auth snapshot saved to {path}")
        except Exception as exc:
            print(f"(could not snapshot auth: {exc})")
        self.close()
        print(f"Session saved to {self.profile_dir}")
        return auth

    def access_token(self) -> Optional[str]:
        """Return the page's MSAL access token, or ``None`` if anonymous."""
        self._ensure_started()
        try:
            return self._page.evaluate(_FIND_TOKEN_JS)
        except PlaywrightError:
            return None

    def cookies(self) -> Dict[str, str]:
        """Return the signed-in Microsoft cookies as a name->value dict."""
        self._ensure_started()
        try:
            raw = self._context.cookies()
        except PlaywrightError:
            return {}
        return {c["name"]: c["value"] for c in raw if "microsoft.com" in c.get("domain", "")}

    def export_auth(self, path: str = DEFAULT_AUTH_FILE, stamp: Optional[float] = None) -> dict:
        """Snapshot the signed-in cookies + access token to ``path`` as JSON.

        ``stamp`` is the epoch seconds to record as ``saved_at`` (pass
        ``time.time()`` from the caller). Returns the auth dict.
        """
        auth = {
            "cookies": self.cookies(),
            "access_token": self.access_token(),
            "saved_at": stamp if stamp is not None else 0,
        }
        dest = Path(path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(json.dumps(auth, indent=2), encoding="utf-8")
        return auth

    # -- chat ---------------------------------------------------------------

    def create_completion(
        self,
        prompt: str,
        stream: bool = False,
        timeout: int = 900,
        **kwargs,
    ) -> Generator[str, None, None]:
        """Stream a Copilot reply to ``prompt``. Mirrors ``Copilot.create_completion``.

        Yields text chunks as they arrive. ``stream`` is accepted for API
        compatibility; chunks are always produced incrementally.
        """
        self._ensure_started()

        if self.region_blocked():
            raise RuntimeError(
                "Microsoft Copilot is not available in your region. "
                "Route the browser through a proxy/VPN in a supported region, e.g.:\n"
                "    BrowserCopilot(proxy='http://user:pass@host:port')\n"
                "or 'socks5://host:port'. See README for details."
            )

        conv = self._page.evaluate(_CREATE_CONVERSATION_JS)
        if not conv.get("ok"):
            status = conv.get("status")
            body = (conv.get("text") or "")[:500]
            if status in (401, 403):
                raise RuntimeError(
                    f"Conversation create returned HTTP {status}. "
                    f"Run login() / `python -m copilot login` to sign in. Body: {body}"
                )
            raise RuntimeError(f"Conversation create failed (HTTP {status}): {body}")

        conversation_id = conv.get("id")
        if not conversation_id:
            raise RuntimeError(f"No conversation id in response: {conv.get('raw')!r}")

        token = self._page.evaluate(_FIND_TOKEN_JS)

        started_ok = self._page.evaluate(_START_STREAM_JS, [conversation_id, token, prompt])
        if started_ok is False:
            state = self._page.evaluate(_POLL_JS)
            raise ConnectionError(f"WebSocket failed to start: {state.get('error')}")

        yield from self._pump(timeout)

    # -- protocol capture ---------------------------------------------------

    def capture_protocol(
        self,
        out_path: str = f"{SESSION_DIR}/protocol_capture.json",
        wait: int = 180,
    ) -> str:
        """Record the *real* Copilot site's chat WebSocket traffic to ``out_path``.

        Opens copilot.microsoft.com in a visible window and taps every WebSocket
        via Playwright's ``framesent``/``framereceived`` events. Send one message
        in the genuine UI and let the reply finish; the capture auto-saves once a
        chat socket reports ``done`` (or after ``wait`` seconds), writing each
        socket's URL and ordered frames to ``out_path`` as JSON.

        This captures the authoritative current protocol (query params, the
        ``setOptions``/``reportLocalConsents`` preamble, the ``send`` frame's
        ``mode``, and the stream events) so the headless
        :class:`~copilot.driver.Copilot` can be kept in sync. Sign in first
        (``python -m copilot login``) for a realistic signed-in capture, and on a
        brand-new account send at least one message manually beforehand to clear
        Copilot's onboarding/consent (until then the backend withholds replies).
        Any ``accessToken``/``access_token`` query param is redacted, but the file
        may still contain conversation text — it lives under the git-ignored
        ``session/`` folder.

        Note: we poll ``page.wait_for_timeout`` rather than blocking on ``input()``
        because Playwright's sync event loop only dispatches the frame callbacks
        while the main thread is inside a Playwright call — a bare ``input()``
        would starve it and capture nothing.
        """
        self._ensure_started()
        sockets: list = []

        def on_ws(ws) -> None:
            record = {"url": self._scrub_token(ws.url), "frames": []}
            sockets.append(record)
            ws.on("framesent",
                  lambda payload: record["frames"].append({"dir": "sent", "payload": self._frame_text(payload)}))
            ws.on("framereceived",
                  lambda payload: record["frames"].append({"dir": "recv", "payload": self._frame_text(payload)}))

        self._page.on("websocket", on_ws)
        print(
            "\nA browser window is open at copilot.microsoft.com.\n"
            "Type a short message in the real Copilot UI and send it, then wait —\n"
            f"the capture saves automatically when the reply finishes (or after {wait}s)."
        )

        def chat_done() -> bool:
            return any(
                '"event":"done"' in f["payload"]
                for s in sockets if "/c/api/chat" in s["url"]
                for f in s["frames"] if f["dir"] == "recv"
            )

        deadline = time.time() + wait
        while time.time() < deadline:
            self._page.wait_for_timeout(500)  # pumps Playwright events
            if chat_done():
                self._page.wait_for_timeout(800)  # flush trailing frames
                break

        dest = Path(out_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(json.dumps(sockets, indent=2), encoding="utf-8")
        total = sum(len(s["frames"]) for s in sockets)
        print(f"Captured {total} frame(s) from {len(sockets)} socket(s) to {out_path}")
        self.close()
        return out_path

    @staticmethod
    def _scrub_token(url: str) -> str:
        """Redact the ``accessToken`` query value so captures aren't secret-bearing."""
        import re

        return re.sub(r"(accessToken=)[^&]*", r"\1REDACTED", url or "")

    @staticmethod
    def _frame_text(payload) -> str:
        """Normalise a Playwright WS frame payload (str or bytes) to text."""
        if isinstance(payload, (bytes, bytearray)):
            return payload.decode("utf-8", errors="replace")
        return payload

    # -- internals ----------------------------------------------------------

    def _pump(self, timeout: int) -> Generator[str, None, None]:
        deadline = time.time() + timeout
        any_text = False
        while True:
            state = self._page.evaluate(_POLL_JS)
            for chunk in state.get("q") or []:
                if chunk:
                    any_text = True
                    yield chunk
            if state.get("error"):
                raise RuntimeError(f"Copilot error: {state['error']}")
            if state.get("done") and not state.get("q"):
                break
            if time.time() > deadline:
                raise TimeoutError(f"No 'done' within {timeout}s")
            time.sleep(0.08)

        if not any_text and not state.get("started"):
            raise RuntimeError("Invalid response: stream produced no text")

    def _ensure_started(self) -> None:
        if self._context is None or self._page is None:
            self.start()
