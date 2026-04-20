"""Office 365 MCP server — exposes calendar, contacts & scheduling tools via in-process MCP."""
from __future__ import annotations

import logging
import re
import unicodedata
from collections import Counter
import json
from datetime import datetime, timedelta, timezone
from difflib import SequenceMatcher
from typing import Any, Dict, List, Tuple, TYPE_CHECKING

from office_con.msgraph.directory_handler import DirectoryUser
from .mcp_base import MsGraphMCPServer

if TYPE_CHECKING:
    from office_con.msgraph.ms_graph_handler import MsGraphInstance

logger = logging.getLogger(__name__)


class Office365MCPServer(MsGraphMCPServer):
    """In-process MCP server that exposes Office 365 tools.

    Uses the authenticated user's MS Graph instance to query their
    Outlook calendar, contacts, and scheduling via the existing handlers.
    """

    def __init__(self, graph: "MsGraphInstance", *, photo_url_prefix: str = "/api/photo", company_dir: object = None) -> None:
        super().__init__(graph)
        self._photo_url_prefix = photo_url_prefix
        self._company_dir = company_dir

    async def get_prompt_hints(self) -> list[str]:
        return [
            "## Contact Card Rendering (MANDATORY)\n"
            "When presenting people found via People Search, you MUST use fenced code blocks "
            "with language `contact`. NEVER list contact details as plain text or bullet points.\n\n"
            "Format — one block per person:\n"
            "````\n"
            "```contact\n"
            "name: Jane Doe\n"
            "email: jane.doe@example.com\n"
            "department: Engineering\n"
            "title: Senior Engineer\n"
            "phone: +49 123 456 789\n"
            "location: Building A\n"
            "manager: John Smith\n"
            "company: Acme Corp\n"
            "city: Stuttgart\n"
            "country: Germany\n"
            "```\n"
            "````\n"
            "Rules:\n"
            "- Use EXACT values from the People Search tool results.\n"
            "- NEVER include a `photo` field — the UI loads photos automatically from the email.\n"
            "- Only `name` and `email` are required. Include ALL other non-empty fields from the search results.\n"
            "- Supported fields: name, email, department, title, phone, location, manager, "
            "company, building, room, street, zip, city, country, joined, birthday.\n"
            "- The `manager` field is hidden in the card but available for answering follow-up questions.\n"
            "- The UI renders these as interactive contact cards — plain text contact info will look broken.\n"
            "- Place your ```contact blocks inline in your response text, wherever you mention the person.",

            "## Org Chart Tool\n"
            "Use `o365_org_chart` when the user asks about reporting lines, org charts, team structure, "
            "direct reports, or who manages someone. The tool returns the person's manager chain (upward) "
            "and direct reports (downward) with full details.\n\n"
            "You can visualise the results as:\n"
            "- A mermaid `graph TD` diagram showing the hierarchy\n"
            "- A text-based tree with indentation\n"
            "- Contact cards for key people\n\n"
            "First use People Search to find the person's email, then call Org Chart with that email.",
        ]

    async def get_client_renderers(self) -> list[dict[str, str]]:
        config = self._contact_card_config()
        renderers = [
            {
                "lang": "contact",
                "css": self._contact_renderer_css(config),
                "js": self._contact_renderer_js(config),
            },
        ]
        # Add inline email enhancer only if domains are configured
        if config.get("emailEnhancerDomains"):
            renderers.append({
                "type": "inline",
                "css": self._email_inline_renderer_css(),
                "js": self._email_inline_renderer_js(config),
            })
        return renderers

    # ── Contact card renderer (CSS + JS) ────────────────────────
    # Override ``_contact_card_config()`` in subclasses to add custom
    # actions (Outlook, Teams, intranet links, …) without touching
    # the base renderer logic.

    # ── SVG icons for standard O365 contact card actions ─────────

    _MAIL_SVG = (
        '<svg viewBox="0 0 24 24" fill="none" stroke="#0078D4" '
        'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
        '<rect x="2" y="4" width="20" height="16" rx="2"/>'
        '<path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>'
        '</svg>'
    )

    _TEAMS_SVG = (
        '<svg viewBox="0 0 24 24" fill="none" stroke="#6264A7" '
        'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>'
        '<path d="M8 10h.01"/><path d="M12 10h.01"/><path d="M16 10h.01"/>'
        '</svg>'
    )

    _ORG_SVG = (
        '<svg viewBox="0 0 24 24" fill="none" stroke="#059669" '
        'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
        '<rect x="9" y="2" width="6" height="4" rx="1"/>'
        '<rect x="2" y="18" width="6" height="4" rx="1"/>'
        '<rect x="16" y="18" width="6" height="4" rx="1"/>'
        '<path d="M12 6v6m0 0H5m7 0h7M5 12v6m14-6v6"/>'
        '</svg>'
    )

    def _contact_card_config(self) -> dict:
        """Return configuration for the contact card renderer.

        Override in subclasses to customise email behaviour and add actions.

        Returns a dict with:
            emailHref: str — URL template for email click.
                Use ``{email}`` placeholder.
            actions: list[dict] — action buttons shown next to the email.
                Each dict has:
                    - url:   str — URL template (``{email}``, ``{name}`` placeholders)
                    - label: str — tooltip text
                    - svg:   str — (optional) inline SVG icon HTML
                    - text:  str — (optional) text label instead of icon
                    - css:   str — (optional) extra inline CSS on the ``<a>`` tag
                    - domain: str — (optional) only show for emails ending with this
            extraCss: str — (optional) additional CSS rules appended to the base styles
            emailEnhancerDomains: list[str] — email domains to enhance inline (empty = disabled)
            contactApiPrefix: str — URL prefix for contact lookup API
        """
        return {
            "emailHref": "https://outlook.office.com/mail/deeplink/compose?to={email}",
            "actions": [
                {
                    "label": "E-Mail (Outlook)",
                    "svg": self._MAIL_SVG,
                    "url": "https://outlook.office.com/mail/deeplink/compose?to={email}",
                },
                {
                    "label": "Chat (Teams)",
                    "svg": self._TEAMS_SVG,
                    "url": "https://teams.microsoft.com/l/chat/0/0?users={email}",
                },
            ],
            "emailEnhancerDomains": [],
            "contactApiPrefix": "/api/contact",
            "photoApiPrefix": self._photo_url_prefix,
            "orgSvg": self._ORG_SVG,
        }

    @staticmethod
    def _contact_renderer_css(config: dict | None = None) -> str:
        base = (
            # Container override — no wrapper chrome
            ".cv2-doc-plugin-block[data-lang='contact'] {"
            "  background: none; border: none; margin: 8px 0;"
            "  display: inline-block; overflow: visible;"
            "}"
            # Card chip
            ".cv2-contact-chip {"
            "  display: inline-flex; align-items: center; gap: 10px;"
            "  background: var(--chat-surface); border: 1px solid var(--chat-border);"
            "  border-radius: var(--chat-radius, 8px); padding: 10px 14px;"
            "  max-width: 380px;"
            "  box-shadow: 0 1px 3px rgba(0,0,0,0.08);"
            "}"
            # Avatar image
            ".cv2-contact-chip-avatar {"
            "  width: 44px; height: 44px; min-width: 44px;"
            "  border-radius: 50%; object-fit: cover;"
            "}"
            # Initials fallback
            ".cv2-contact-chip-initials {"
            "  width: 44px; height: 44px; min-width: 44px;"
            "  border-radius: 50%; background: var(--chat-accent, #6366f1);"
            "  color: #fff; display: flex; align-items: center;"
            "  justify-content: center; font-size: 15px; font-weight: 700;"
            "}"
            # Info column
            ".cv2-contact-chip-info {"
            "  display: flex; flex-direction: column; gap: 1px;"
            "  overflow: hidden;"
            "}"
            ".cv2-contact-chip-name { font-weight: 700; font-size: 14px; color: var(--chat-text, #111827); }"
            ".cv2-contact-chip-title { font-size: 12.5px; font-weight: 500; color: var(--chat-text-muted, #4b5563); }"
            ".cv2-contact-chip-dept { font-size: 12.5px; color: var(--chat-text-muted, #4b5563); opacity: 0.85; }"
            # Email link
            ".cv2-contact-chip a.cv2-contact-chip-email {"
            "  font-size: 12.5px; font-weight: 500; color: var(--chat-accent, #4f46e5);"
            "  text-decoration: none;"
            "}"
            ".cv2-contact-chip a.cv2-contact-chip-email:hover { text-decoration: underline; }"
            # Links row (email + actions)
            ".cv2-contact-chip-links {"
            "  display: flex; align-items: center; gap: 6px; margin-top: 2px;"
            "  flex-wrap: wrap;"
            "}"
            # Action icons row — always on its own line
            ".cv2-contact-chip-actions {"
            "  display: flex; align-items: center; gap: 8px;"
            "  flex-basis: 100%;"
            "}"
            # Action icon button (shared base)
            ".cv2-contact-chip a.cv2-contact-chip-action {"
            "  display: inline-flex; align-items: center; justify-content: center;"
            "  width: 26px; height: 26px; border-radius: 6px;"
            "  text-decoration: none; opacity: 0.7; transition: opacity 0.15s;"
            "}"
            ".cv2-contact-chip a.cv2-contact-chip-action:hover { opacity: 1; }"
            ".cv2-contact-chip-action svg { width: 20px; height: 20px; }"
            # Details row (phone, location, company)
            ".cv2-contact-chip-details {"
            "  display: flex; flex-wrap: wrap; gap: 2px 10px;"
            "  font-size: 11.5px; color: var(--chat-text-muted, #6b7280);"
            "  margin-top: 3px;"
            "}"
            ".cv2-contact-chip-details span {"
            "  display: inline-flex; align-items: center; gap: 3px;"
            "}"
            ".cv2-contact-chip-details .material-icons {"
            "  font-size: 13px; opacity: 0.7;"
            "}"
            # Org trigger icon
            ".cv2-org-trigger { cursor: pointer; }"
            # Org popup
            ".cv2-org-popup {"
            "  flex-direction: column; gap: 4px;"
            "  padding: 8px 12px; width: fit-content; max-width: 90vw;"
            "}"
            ".cv2-org-content {"
            "  display: flex; flex-direction: column; gap: 6px;"
            "}"
            ".cv2-org-row {"
            "  display: flex; align-items: center; gap: 8px;"
            "  font-size: 12.5px; color: var(--chat-text, #111827);"
            "}"
            ".cv2-org-row .material-icons {"
            "  font-size: 16px; color: var(--chat-text-muted, #6b7280); opacity: 0.7; flex-shrink: 0;"
            "}"
            ".cv2-org-row img { border-radius: 50%; object-fit: cover; flex-shrink: 0; }"
            ".cv2-org-row small { color: var(--chat-text-muted, #6b7280); font-size: 11px; }"
            ".cv2-org-grid {"
            "  display: flex; flex-direction: column; flex-wrap: wrap;"
            "  gap: 3px 16px; padding-left: 24px;"
            "}"
            ".cv2-org-report {"
            "  display: flex; align-items: center; gap: 6px;"
            "  font-size: 12px; color: var(--chat-text, #111827);"
            "  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;"
            "}"
            ".cv2-org-report img, .cv2-org-report .cv2-contact-chip-initials {"
            "  border-radius: 50%; object-fit: cover; flex-shrink: 0;"
            "}"
            ".cv2-org-report small { color: var(--chat-text-muted, #6b7280); font-size: 11px; }"
        )
        extra = (config or {}).get("extraCss", "")
        return base + extra if extra else base

    @staticmethod
    def _contact_renderer_js(config: dict | None = None) -> str:
        import json as _json
        cfg_json = _json.dumps(config or {"emailHref": "mailto:{email}", "actions": []})
        return """
(function(registry, lang) {
  var CFG = """ + cfg_json + """;
  function escHtml(s) {
    var d = document.createElement('div');
    d.textContent = s || '';
    return d.innerHTML;
  }
  function escAttr(s) {
    return (s || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
  }
  function parseContact(raw) {
    var data = {};
    raw.split('\\n').forEach(function(line) {
      var idx = line.indexOf(':');
      if (idx < 1) return;
      var key = line.substring(0, idx).trim().toLowerCase();
      var val = line.substring(idx + 1).trim();
      if (val) data[key] = val;
    });
    return data;
  }
  function tpl(template, email, name) {
    return template.replace(/\\{email\\}/g, encodeURIComponent(email))
                   .replace(/\\{name\\}/g, encodeURIComponent(name));
  }
  var _orgPopup = null;
  var _orgHideTimer = null;
  var _orgCache = {};

  function getOrgPopup() {
    if (_orgPopup) return _orgPopup;
    _orgPopup = document.createElement('div');
    _orgPopup.className = 'cv2-contact-popup cv2-org-popup';
    document.body.appendChild(_orgPopup);
    _orgPopup.addEventListener('mouseenter', function() {
      if (_orgHideTimer) { clearTimeout(_orgHideTimer); _orgHideTimer = null; }
    });
    _orgPopup.addEventListener('mouseleave', function() {
      _orgHideTimer = setTimeout(function() { _orgPopup.style.display = 'none'; }, 200);
    });
    return _orgPopup;
  }

  async function fetchOrg(email) {
    if (_orgCache[email] !== undefined) return _orgCache[email];
    var prefix = CFG.contactApiPrefix || '/api/contact';
    try {
      var resp = await fetch(prefix + '/by-email/' + encodeURIComponent(email));
      if (!resp.ok) { _orgCache[email] = null; return null; }
      var data = await resp.json();
      _orgCache[email] = data;
      return data;
    } catch (e) { _orgCache[email] = null; return null; }
  }

  function hasOrgData(d) {
    return d && (d.manager || (d.direct_reports && d.direct_reports.length));
  }

  function orgAvatar(person, size) {
    var ini = (person.name || '').split(/[\\s,]+/).map(function(w){return (w[0]||'')}).join('').substring(0,2).toUpperCase();
    var photoUrl = person.email ? (CFG.photoApiPrefix || '/api/photo') + '/by-email/' + encodeURIComponent(person.email) : '';
    var iniHtml = '<div class="cv2-contact-chip-initials cv2-org-av-slot" style="width:'+size+'px;height:'+size+'px;min-width:'+size+'px;font-size:'+(size*0.4)+'px"' +
      (photoUrl ? ' data-photo="' + escAttr(photoUrl) + '" data-sz="' + size + '"' : '') +
      '>' + ini + '</div>';
    return iniHtml;
  }

  function buildOrgHtml(orgData) {
    var lines = [];
    if (orgData.manager) {
      var m = orgData.manager;
      lines.push('<div class="cv2-org-row"><span class="material-icons">supervisor_account</span> ' +
        orgAvatar(m, 28) + '<span><strong>' + escHtml(m.name) + '</strong>' +
        (m.title ? '<br><small>' + escHtml(m.title) + '</small>' : '') +
        '</span></div>');
    }
    if (orgData.direct_reports && orgData.direct_reports.length) {
      var n = orgData.direct_reports.length;
      lines.push('<div class="cv2-org-row cv2-org-header"><span class="material-icons">groups</span> <span><strong>' +
        n + ' direct reports</strong></span></div>');
      var cols = Math.min(3, Math.ceil(n / 8));
      var rows = Math.ceil(n / cols);
      var mh = rows * 28;
      var grid = '<div class="cv2-org-grid" style="max-height:' + mh + 'px">';
      orgData.direct_reports.forEach(function(r) {
        grid += '<div class="cv2-org-report">' +
          orgAvatar(r, 22) + '<span>' + escHtml(r.name) +
          (r.title ? ' <small>— ' + escHtml(r.title) + '</small>' : '') +
          '</span></div>';
      });
      grid += '</div>';
      lines.push(grid);
    }
    return lines.join('');
  }

  function fixOrgAvatars(root) {
    root.querySelectorAll('.cv2-org-av-slot[data-photo]').forEach(function(slot) {
      var url = slot.dataset.photo;
      var sz = slot.dataset.sz || 22;
      if (!url) return;
      fetch(url).then(function(r) {
        if (!r.ok) return;
        return r.blob();
      }).then(function(blob) {
        if (!blob) return;
        var objUrl = URL.createObjectURL(blob);
        var img = document.createElement('img');
        img.className = 'cv2-contact-chip-avatar';
        img.style.cssText = 'width:'+sz+'px;height:'+sz+'px;min-width:'+sz+'px';
        img.src = objUrl;
        slot.replaceWith(img);
      }).catch(function(){});
    });
  }

  function showOrgPopup(el, orgData) {
    var popup = getOrgPopup();
    popup.innerHTML = '<div class="cv2-org-content">' + buildOrgHtml(orgData) + '</div>';
    fixOrgAvatars(popup);
    var rect = el.getBoundingClientRect();
    popup.style.display = 'flex';
    var popupRect = popup.getBoundingClientRect();
    var left = rect.left;
    var top = rect.bottom + 6;
    if (left + popupRect.width > window.innerWidth - 8) left = window.innerWidth - popupRect.width - 8;
    if (top + popupRect.height > window.innerHeight - 8) top = rect.top - popupRect.height - 6;
    popup.style.left = Math.max(8, left) + 'px';
    popup.style.top = Math.max(8, top) + 'px';
  }

  registry.register(lang, {
    inline: true,
    sidebar: false,
    render: function(container, rawData) {
      var data = parseContact(rawData);
      if (!data.name) { container.textContent = rawData; return; }
      var initials = data.name.split(/[\\s,]+/).map(function(w){return (w[0]||'')}).join('').substring(0,2).toUpperCase();

      // Avatar — always fetch photo by email, server returns initials fallback if needed
      var avatar;
      var avatarPhotoUrl = data.email ? (CFG.photoApiPrefix || '/api/photo') + '/by-email/' + encodeURIComponent(data.email) : null;
      avatar = '<div class="cv2-contact-chip-initials">' + initials + '</div>';

      // Info lines
      var titleHtml = data.title ? '<span class="cv2-contact-chip-title">' + escHtml(data.title) + '</span>' : '';
      var deptHtml = data.department ? '<span class="cv2-contact-chip-dept">' + escHtml(data.department) + '</span>' : '';

      // Detail row: phone, location, company (small icons)
      var detailParts = [];
      if (data.phone) detailParts.push('<span><span class="material-icons">phone</span>' + escHtml(data.phone) + '</span>');
      var loc = data.building || data.location || '';
      if (loc) detailParts.push('<span><span class="material-icons">location_on</span>' + escHtml(loc) + '</span>');
      else if (data.city) detailParts.push('<span><span class="material-icons">location_on</span>' + escHtml(data.city + (data.country ? ', ' + data.country : '')) + '</span>');
      if (data.company) detailParts.push('<span><span class="material-icons">business</span>' + escHtml(data.company) + '</span>');
      var detailsHtml = detailParts.length ? '<div class="cv2-contact-chip-details">' + detailParts.join('') + '</div>' : '';

      // Links row: email + configurable actions
      var linksHtml = '';
      if (data.email) {
        var emailHref = tpl(CFG.emailHref || 'mailto:{email}', data.email, data.name);
        var emailLink = '<a class="cv2-contact-chip-email" href="' + escAttr(emailHref) + '" target="_blank">' + escHtml(data.email) + '</a>';

        // Build action icons from config
        var actions = '';
        var emailLower = data.email.toLowerCase();
        (CFG.actions || []).forEach(function(a) {
          if (a.domain && !emailLower.endsWith(a.domain.toLowerCase())) return;
          var url = tpl(a.url, data.email, data.name);
          var content = a.svg || escHtml(a.text || '');
          var style = a.css ? ' style="' + escAttr(a.css) + '"' : '';
          actions += '<a class="cv2-contact-chip-action" href="' + escAttr(url) + '" target="_blank" title="' + escAttr(a.label || '') + '"' + style + '>' + content + '</a>';
        });

        var actionsHtml = actions ? '<span class="cv2-contact-chip-actions">' + actions + '</span>' : '';
        linksHtml = '<div class="cv2-contact-chip-links">' + emailLink + actionsHtml + '</div>';
      }

      container.innerHTML =
        '<div class="cv2-contact-chip">' +
          avatar +
          '<div class="cv2-contact-chip-info">' +
            '<span class="cv2-contact-chip-name">' + escHtml(data.name) + '</span>' +
            titleHtml + deptHtml + detailsHtml + linksHtml +
          '</div>' +
        '</div>';

      // Async photo load — only insert <img> if fetch succeeds (avoids 404 console noise)
      if (avatarPhotoUrl) {
        (function(url, chip) {
          fetch(url).then(function(r) {
            if (!r.ok) return;
            return r.blob();
          }).then(function(blob) {
            if (!blob) return;
            var objUrl = URL.createObjectURL(blob);
            var img = document.createElement('img');
            img.className = 'cv2-contact-chip-avatar';
            img.src = objUrl;
            var old = chip.querySelector('.cv2-contact-chip-initials');
            if (old) old.replaceWith(img);
          }).catch(function(){});
        })(avatarPhotoUrl, container.querySelector('.cv2-contact-chip'));
      }

      // Async: fetch org data, only show org icon if manager or direct reports exist
      if (data.email && CFG.orgSvg) {
        fetchOrg(data.email).then(function(orgData) {
          if (!hasOrgData(orgData)) return;
          var actionsRow = container.querySelector('.cv2-contact-chip-actions');
          if (!actionsRow) return;
          var orgEl = document.createElement('span');
          orgEl.className = 'cv2-contact-chip-action cv2-org-trigger';
          orgEl.title = 'Org Chart';
          orgEl.innerHTML = CFG.orgSvg;
          actionsRow.appendChild(orgEl);
          orgEl.addEventListener('mouseenter', function() {
            if (_orgHideTimer) { clearTimeout(_orgHideTimer); _orgHideTimer = null; }
            showOrgPopup(orgEl, orgData);
          });
          orgEl.addEventListener('mouseleave', function() {
            _orgHideTimer = setTimeout(function() {
              if (_orgPopup) _orgPopup.style.display = 'none';
            }, 200);
          });
        });
      }
    }
  });
})(registry, lang);
"""

    # ── Inline email enhancer (CSS + JS) ──────────────────────

    @staticmethod
    def _email_inline_renderer_css() -> str:
        return (
            ".cv2-email-enhanced {"
            "  border-bottom: 1px dashed var(--chat-accent, #4f46e5);"
            "  cursor: pointer; transition: border-color 0.15s, color 0.15s;"
            "}"
            ".cv2-email-enhanced:hover {"
            "  border-bottom-style: solid;"
            "  color: var(--chat-accent, #4f46e5);"
            "}"
            ".cv2-contact-popup {"
            "  position: fixed; z-index: 9999;"
            "  display: none; align-items: flex-start; gap: 10px;"
            "  background: var(--chat-surface, #fff); border: 1px solid var(--chat-border, #e5e7eb);"
            "  border-radius: var(--chat-radius, 8px); padding: 10px 14px;"
            "  box-shadow: 0 4px 16px rgba(0,0,0,0.15); max-width: 380px;"
            "}"
            # Extra details row (phone, location, company)
            ".cv2-contact-chip-details {"
            "  display: flex; flex-wrap: wrap; gap: 2px 10px;"
            "  font-size: 11.5px; color: var(--chat-text-muted, #6b7280);"
            "  margin-top: 3px;"
            "}"
            ".cv2-contact-chip-details span {"
            "  display: inline-flex; align-items: center; gap: 3px;"
            "}"
            ".cv2-contact-chip-details .material-icons {"
            "  font-size: 13px; opacity: 0.7;"
            "}"
            # Org trigger icon (in actions row)
            ".cv2-org-trigger, .cv2-org-trigger-inline {"
            "  cursor: pointer;"
            "}"
            # Org popup
            ".cv2-org-popup {"
            "  flex-direction: column; gap: 4px;"
            "  padding: 8px 12px; width: fit-content; max-width: 90vw;"
            "}"
            ".cv2-org-content {"
            "  display: flex; flex-direction: column; gap: 6px;"
            "}"
            ".cv2-org-row {"
            "  display: flex; align-items: center; gap: 8px;"
            "  font-size: 12.5px; color: var(--chat-text, #111827);"
            "}"
            ".cv2-org-row .material-icons {"
            "  font-size: 16px; color: var(--chat-text-muted, #6b7280); opacity: 0.7; flex-shrink: 0;"
            "}"
            ".cv2-org-row img {"
            "  border-radius: 50%; object-fit: cover; flex-shrink: 0;"
            "}"
            ".cv2-org-row small {"
            "  color: var(--chat-text-muted, #6b7280); font-size: 11px;"
            "}"
            ".cv2-org-grid {"
            "  display: flex; flex-direction: column; flex-wrap: wrap;"
            "  gap: 3px 16px; padding-left: 24px;"
            "}"
            ".cv2-org-report {"
            "  display: flex; align-items: center; gap: 6px;"
            "  font-size: 12px; color: var(--chat-text, #111827);"
            "  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;"
            "}"
            ".cv2-org-report img, .cv2-org-report .cv2-contact-chip-initials {"
            "  border-radius: 50%; object-fit: cover; flex-shrink: 0;"
            "}"
            ".cv2-org-report small { color: var(--chat-text-muted, #6b7280); font-size: 11px; }"
        )

    @staticmethod
    def _email_inline_renderer_js(config: dict) -> str:
        import json as _json
        cfg_json = _json.dumps(config)
        # NOTE: This JS is evaluated via new Function('registry', js).
        # Backslashes in the Python triple-quoted string are interpreted once
        # by Python, producing the JS source text.  So Python '\\n' → JS '\n',
        # Python '\\\\' → JS '\\', Python '\\{' → JS '\{'.
        return (
            "(function(registry) {\n"
            "var CFG = " + cfg_json + ";\n"
            r"""
var _cache = {};
var _popup = null;
var _hideTimer = null;

function escHtml(s) {
  var d = document.createElement('div');
  d.textContent = s || '';
  return d.innerHTML;
}
function escAttr(s) {
  return (s || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
}
function tpl(template, email, name) {
  return template.replace(/\{email\}/g, encodeURIComponent(email))
                 .replace(/\{name\}/g, encodeURIComponent(name));
}

function getPopup() {
  if (_popup) return _popup;
  _popup = document.createElement('div');
  _popup.className = 'cv2-contact-popup';
  document.body.appendChild(_popup);
  _popup.addEventListener('mouseenter', function() {
    if (_hideTimer) { clearTimeout(_hideTimer); _hideTimer = null; }
  });
  _popup.addEventListener('mouseleave', function() {
    _hideTimer = setTimeout(function() { _popup.style.display = 'none'; }, 200);
  });
  return _popup;
}

function matchesDomain(email) {
  var domains = CFG.emailEnhancerDomains || [];
  var lower = email.toLowerCase();
  for (var i = 0; i < domains.length; i++) {
    if (lower.endsWith('@' + domains[i].toLowerCase())) return true;
  }
  return false;
}

var _orgPopup2 = null;
var _orgHideTimer2 = null;

function getOrgPopup2() {
  if (_orgPopup2) return _orgPopup2;
  _orgPopup2 = document.createElement('div');
  _orgPopup2.className = 'cv2-contact-popup cv2-org-popup';
  document.body.appendChild(_orgPopup2);
  _orgPopup2.addEventListener('mouseenter', function() {
    if (_orgHideTimer2) { clearTimeout(_orgHideTimer2); _orgHideTimer2 = null; }
  });
  _orgPopup2.addEventListener('mouseleave', function() {
    _orgHideTimer2 = setTimeout(function() { _orgPopup2.style.display = 'none'; }, 200);
  });
  return _orgPopup2;
}

function hasOrgData2(d) {
  return d && (d.manager || (d.direct_reports && d.direct_reports.length));
}

function orgAvatar2(person, size) {
  var ini = (person.name || '').split(/[\s,]+/).map(function(w){return (w[0]||'')}).join('').substring(0,2).toUpperCase();
  var photoUrl = person.email ? (CFG.photoApiPrefix || '/api/photo') + '/by-email/' + encodeURIComponent(person.email) : '';
  var iniHtml = '<div class="cv2-contact-chip-initials cv2-org-av-slot" style="width:'+size+'px;height:'+size+'px;min-width:'+size+'px;font-size:'+(size*0.4)+'px"' +
    (photoUrl ? ' data-photo="' + escAttr(photoUrl) + '" data-sz="' + size + '"' : '') +
    '>' + ini + '</div>';
  return iniHtml;
}

function buildOrgHtml2(data) {
  var lines = [];
  if (data.manager) {
    var m = data.manager;
    lines.push('<div class="cv2-org-row"><span class="material-icons">supervisor_account</span> ' +
      orgAvatar2(m, 28) + '<span><strong>' + escHtml(m.name) + '</strong>' +
      (m.title ? '<br><small>' + escHtml(m.title) + '</small>' : '') +
      '</span></div>');
  }
  if (data.direct_reports && data.direct_reports.length) {
    var n = data.direct_reports.length;
    lines.push('<div class="cv2-org-row cv2-org-header"><span class="material-icons">groups</span> <span><strong>' +
      n + ' direct reports</strong></span></div>');
    var cols = Math.min(3, Math.ceil(n / 8));
    var rows = Math.ceil(n / cols);
    var mh = rows * 28;
    var grid = '<div class="cv2-org-grid" style="max-height:' + mh + 'px">';
    data.direct_reports.forEach(function(r) {
      grid += '<div class="cv2-org-report">' +
        orgAvatar2(r, 22) + '<span>' + escHtml(r.name) +
        (r.title ? ' <small>— ' + escHtml(r.title) + '</small>' : '') +
        '</span></div>';
    });
    grid += '</div>';
    lines.push(grid);
  }
  return lines.join('');
}

function fixOrgAvatars2(root) {
  root.querySelectorAll('.cv2-org-av-slot[data-photo]').forEach(function(slot) {
    var url = slot.dataset.photo;
    var sz = slot.dataset.sz || 22;
    if (!url) return;
    fetch(url).then(function(r) {
      if (!r.ok) return;
      return r.blob();
    }).then(function(blob) {
      if (!blob) return;
      var objUrl = URL.createObjectURL(blob);
      var img = document.createElement('img');
      img.className = 'cv2-contact-chip-avatar';
      img.style.cssText = 'width:'+sz+'px;height:'+sz+'px;min-width:'+sz+'px';
      img.src = objUrl;
      slot.replaceWith(img);
    }).catch(function(){});
  });
}

function showOrgPopup2(el, data) {
  var popup = getOrgPopup2();
  popup.innerHTML = '<div class="cv2-org-content">' + buildOrgHtml2(data) + '</div>';
  fixOrgAvatars2(popup);
  var rect = el.getBoundingClientRect();
  popup.style.display = 'flex';
  var popupRect = popup.getBoundingClientRect();
  var left = rect.left;
  var top = rect.bottom + 6;
  if (left + popupRect.width > window.innerWidth - 8) left = window.innerWidth - popupRect.width - 8;
  if (top + popupRect.height > window.innerHeight - 8) top = rect.top - popupRect.height - 6;
  popup.style.left = Math.max(8, left) + 'px';
  popup.style.top = Math.max(8, top) + 'px';
}

function showPopup(el, data) {
  var popup = getPopup();
  var initials = data.name.split(/[\s,]+/).map(function(w){return (w[0]||'')}).join('').substring(0,2).toUpperCase();

  var avatar = '<div class="cv2-contact-chip-initials">' + initials + '</div>';
  var avatarPhotoUrl2 = data.email ? (CFG.photoApiPrefix || '/api/photo') + '/by-email/' + encodeURIComponent(data.email) : null;

  var titleHtml = data.title ? '<span class="cv2-contact-chip-title">' + escHtml(data.title) + '</span>' : '';
  var deptHtml = data.department ? '<span class="cv2-contact-chip-dept">' + escHtml(data.department) + '</span>' : '';

  // Detail row: phone, location, company
  var detailParts = [];
  if (data.phone) detailParts.push('<span><span class="material-icons">phone</span>' + escHtml(data.phone) + '</span>');
  var loc = data.building || data.location || '';
  if (loc) detailParts.push('<span><span class="material-icons">location_on</span>' + escHtml(loc) + '</span>');
  else if (data.city) detailParts.push('<span><span class="material-icons">location_on</span>' + escHtml(data.city + (data.country ? ', ' + data.country : '')) + '</span>');
  if (data.company) detailParts.push('<span><span class="material-icons">business</span>' + escHtml(data.company) + '</span>');
  var detailsHtml = detailParts.length ? '<div class="cv2-contact-chip-details">' + detailParts.join('') + '</div>' : '';

  var linksHtml = '';
  if (data.email) {
    var emailHref = tpl(CFG.emailHref || 'mailto:{email}', data.email, data.name);
    var emailLink = '<a class="cv2-contact-chip-email" href="' + escAttr(emailHref) + '" target="_blank">' + escHtml(data.email) + '</a>';
    var actions = '';
    var emailLower = data.email.toLowerCase();
    (CFG.actions || []).forEach(function(a) {
      if (a.domain && !emailLower.endsWith(a.domain.toLowerCase())) return;
      var url = tpl(a.url, data.email, data.name);
      var content = a.svg || escHtml(a.text || '');
      var style = a.css ? ' style="' + escAttr(a.css) + '"' : '';
      actions += '<a class="cv2-contact-chip-action" href="' + escAttr(url) + '" target="_blank" title="' + escAttr(a.label || '') + '"' + style + '>' + content + '</a>';
    });

    // Org chart icon — shows manager + reports on hover (only if data exists)
    if (CFG.orgSvg && hasOrgData2(data)) {
      actions += '<span class="cv2-contact-chip-action cv2-org-trigger-inline" title="Org Chart">' + CFG.orgSvg + '</span>';
    }

    var actionsHtml = actions ? '<span class="cv2-contact-chip-actions">' + actions + '</span>' : '';
    linksHtml = '<div class="cv2-contact-chip-links">' + emailLink + actionsHtml + '</div>';
  }

  popup.innerHTML =
    '<div class="cv2-contact-chip" style="box-shadow:none;border:none;padding:0">' +
      avatar +
      '<div class="cv2-contact-chip-info">' +
        '<span class="cv2-contact-chip-name">' + escHtml(data.name) + '</span>' +
        titleHtml + deptHtml + detailsHtml + linksHtml +
      '</div>' +
    '</div>';

  // Async photo upgrade — only insert <img> if fetch succeeds
  if (avatarPhotoUrl2) {
    (function(url, chip) {
      fetch(url).then(function(r) {
        if (!r.ok) return;
        return r.blob();
      }).then(function(blob) {
        if (!blob) return;
        var objUrl = URL.createObjectURL(blob);
        var img = document.createElement('img');
        img.className = 'cv2-contact-chip-avatar';
        img.src = objUrl;
        var old = chip.querySelector('.cv2-contact-chip-initials');
        if (old) old.replaceWith(img);
      }).catch(function(){});
    })(avatarPhotoUrl2, popup.querySelector('.cv2-contact-chip'));
  }

  // Image error fallback -> initials
  var img = popup.querySelector('.cv2-contact-chip-avatar');
  if (img) {
    img.onerror = function() {
      var fb = document.createElement('div');
      fb.className = 'cv2-contact-chip-initials';
      fb.textContent = initials;
      img.replaceWith(fb);
    };
  }

  // Org icon hover handler within popup
  var orgTrig = popup.querySelector('.cv2-org-trigger-inline');
  if (orgTrig) {
    orgTrig.addEventListener('mouseenter', function() {
      if (_orgHideTimer2) { clearTimeout(_orgHideTimer2); _orgHideTimer2 = null; }
      showOrgPopup2(orgTrig, data);
    });
    orgTrig.addEventListener('mouseleave', function() {
      _orgHideTimer2 = setTimeout(function() {
        if (_orgPopup2) _orgPopup2.style.display = 'none';
      }, 200);
    });
  }

  // Position near the element
  var rect = el.getBoundingClientRect();
  popup.style.display = 'flex';
  var popupRect = popup.getBoundingClientRect();
  var left = rect.left;
  var top = rect.bottom + 6;
  // Keep within viewport
  if (left + popupRect.width > window.innerWidth - 8) left = window.innerWidth - popupRect.width - 8;
  if (top + popupRect.height > window.innerHeight - 8) top = rect.top - popupRect.height - 6;
  popup.style.left = Math.max(8, left) + 'px';
  popup.style.top = Math.max(8, top) + 'px';
}

async function fetchContact(email) {
  if (_cache[email] !== undefined) return _cache[email];
  var prefix = CFG.contactApiPrefix || '/api/contact';
  try {
    var resp = await fetch(prefix + '/by-email/' + encodeURIComponent(email));
    if (!resp.ok) { _cache[email] = null; return null; }
    var data = await resp.json();
    _cache[email] = data;
    return data;
  } catch (e) {
    _cache[email] = null;
    return null;
  }
}

function attachHandlers(el, email) {
  el.addEventListener('mouseenter', async function() {
    if (_hideTimer) { clearTimeout(_hideTimer); _hideTimer = null; }
    var data = await fetchContact(email);
    if (data) showPopup(el, data);
  });
  el.addEventListener('mouseleave', function() {
    _hideTimer = setTimeout(function() {
      if (_popup) _popup.style.display = 'none';
    }, 200);
  });
  el.addEventListener('click', async function(e) {
    e.preventDefault();
    e.stopPropagation();
    if (_hideTimer) { clearTimeout(_hideTimer); _hideTimer = null; }
    var data = await fetchContact(email);
    if (data) showPopup(el, data);
  });
}

registry.registerInline(async function(container) {
  // 1. Enhance <a href="mailto:..."> links whose email matches configured domains
  var mailtoLinks = container.querySelectorAll('a[href^="mailto:"]');
  for (var j = 0; j < mailtoLinks.length; j++) {
    var a = mailtoLinks[j];
    if (a.closest('.cv2-contact-chip, .cv2-contact-popup, .cv2-email-enhanced')) continue;
    var email = a.href.replace(/^mailto:/i, '').split('?')[0];
    if (!matchesDomain(email)) continue;
    a.classList.add('cv2-email-enhanced');
    a.dataset.email = email;
    attachHandlers(a, email);
  }

  // 2. Enhance plain-text emails not inside links or already enhanced elements
  var walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, {
    acceptNode: function(node) {
      var p = node.parentElement;
      if (!p) return NodeFilter.FILTER_REJECT;
      if (p.closest('.cv2-contact-chip, .cv2-contact-popup, .cv2-email-enhanced, a, code, pre'))
        return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    }
  });

  var textNodes = [];
  while (walker.nextNode()) textNodes.push(walker.currentNode);

  var emailPattern = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
  for (var i = 0; i < textNodes.length; i++) {
    var node = textNodes[i];
    var text = node.textContent;
    emailPattern.lastIndex = 0;
    if (!emailPattern.test(text)) continue;

    emailPattern.lastIndex = 0;
    var frag = document.createDocumentFragment();
    var lastIdx = 0;
    var match;
    while ((match = emailPattern.exec(text)) !== null) {
      if (!matchesDomain(match[0])) {
        // Not a matching domain — leave as text
        continue;
      }
      if (match.index > lastIdx) {
        frag.appendChild(document.createTextNode(text.substring(lastIdx, match.index)));
      }
      var span = document.createElement('span');
      span.className = 'cv2-email-enhanced';
      span.dataset.email = match[0];
      span.textContent = match[0];
      attachHandlers(span, match[0]);
      frag.appendChild(span);
      lastIdx = emailPattern.lastIndex;
    }
    if (frag.childNodes.length > 0) {
      if (lastIdx < text.length) {
        frag.appendChild(document.createTextNode(text.substring(lastIdx)));
      }
      node.parentNode.replaceChild(frag, node);
    }
  }
});

"""
            "})(registry);\n"
        )

    async def list_tools(self) -> List[Dict[str, Any]]:
        return [
            {
                "name": "o365_calendar_get_events",
                "displayName": "Calendar Events",
                "displayDescription": "Search or list calendar events",
                "icon": "event",
                "description": (
                    "Search or list calendar events for the logged-in user. "
                    "Use 'query' to find specific events by name (searches subject). "
                    "Without query, returns upcoming events. "
                    "Returns subject, time, location, and attendees."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "Search term to filter events by subject (case-insensitive). Use this to find specific events like 'DevCon' or 'standup'.",
                        },
                        "start_date": {
                            "type": "string",
                            "description": "Start date in YYYY-MM-DD format. Defaults to today.",
                        },
                        "end_date": {
                            "type": "string",
                            "description": "End date in YYYY-MM-DD format. Defaults to 14 days from start, or 6 months when query is set.",
                        },
                        "limit": {
                            "type": "integer",
                            "description": "Max events to return. Default 50.",
                            "default": 50,
                        },
                    },
                },
            },
            {
                "name": "o365_resolve_contact",
                "displayName": "People Search",
                "displayDescription": "Find colleagues in the company directory",
                "icon": "person_search",
                "description": (
                    "Look up one or more people in the company directory by name. "
                    "Supports fuzzy matching — typos, umlaut variants (Sauter/Sautter, Mueller/Müller) "
                    "are handled automatically. Pass a single name or an array of names. "
                    "Returns full name, email, department, job title, photo, and match score. "
                    "Results are grouped by search term when multiple names are given. "
                    "Use this whenever the user asks about colleagues, coworkers, or anyone who might "
                    "be in the organization — e.g. 'who is ...', 'find ...', 'look up ...', "
                    "or when you need someone's email before scheduling."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": (
                                "Name(s) to search for. For multiple people, separate with commas "
                                "(e.g. 'Daniel Sagmeister, Jörg Sauter'). Fuzzy matching handles "
                                "typos and umlaut variants automatically."
                            ),
                        },
                    },
                    "required": ["query"],
                },
            },
            {
                "name": "o365_calendar_find_free_slots",
                "displayName": "Free Slot Finder",
                "displayDescription": "Find common free time for meetings",
                "icon": "schedule",
                "description": (
                    "Find common free time slots between the logged-in user and one "
                    "or more attendees. Returns available windows that fit the "
                    "requested meeting duration within working hours."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "attendee_emails": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Email addresses of attendees to check availability for.",
                        },
                        "duration_minutes": {
                            "type": "integer",
                            "description": "Desired meeting length in minutes. Default 60.",
                            "default": 60,
                        },
                        "start_date": {
                            "type": "string",
                            "description": "Start date in YYYY-MM-DD format. Defaults to tomorrow.",
                        },
                        "end_date": {
                            "type": "string",
                            "description": "End date in YYYY-MM-DD format. Defaults to 5 business days from start_date.",
                        },
                        "start_hour": {
                            "type": "integer",
                            "description": "Earliest hour (0-23) for suggested slots. Default 8.",
                            "default": 8,
                        },
                        "end_hour": {
                            "type": "integer",
                            "description": "Latest hour (0-23) for suggested slots. Default 18.",
                            "default": 18,
                        },
                    },
                    "required": ["attendee_emails"],
                },
            },
            {
                "name": "o365_org_chart",
                "displayName": "Org Chart",
                "displayDescription": "Show reporting structure and direct reports",
                "icon": "account_tree",
                "description": (
                    "Show the organizational structure for a person: their manager chain "
                    "(upward) and direct reports (downward). Use this when the user asks "
                    "about reporting lines, org charts, team structure, who reports to whom, "
                    "or who manages someone. Can also be used to draw mermaid org chart diagrams. "
                    "Returns structured data with manager chain and direct reports including "
                    "name, email, title, department, phone, and location."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "email": {
                            "type": "string",
                            "description": "Email address of the person to show the org chart for.",
                        },
                        "depth_up": {
                            "type": "integer",
                            "description": "How many levels of managers to show upward. Default 3.",
                            "default": 3,
                        },
                        "depth_down": {
                            "type": "integer",
                            "description": "How many levels of direct reports to show downward. Default 1.",
                            "default": 1,
                        },
                    },
                    "required": ["email"],
                },
            },
            # ── Rooms ────────────────────────────────────────────
            {
                "name": "o365_list_rooms",
                "displayName": "Meeting Rooms",
                "icon": "meeting_room",
                "description": (
                    "List available meeting rooms with capacity and location.\n"
                    "Use the room name with o365_get_room_availability to check bookings."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "filter": {"type": "string", "description": "Filter rooms by name (optional)"},
                    },
                    "required": [],
                },
            },
            {
                "name": "o365_get_room_availability",
                "displayName": "Room Availability",
                "icon": "event_available",
                "description": (
                    "Check when meeting rooms are free or busy today (or a given date).\n"
                    "Pass room names (or substrings) from o365_list_rooms.\n"
                    "Returns time slots with free/busy status in the user's timezone."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "rooms": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Room names or name substrings to check",
                        },
                        "date": {"type": "string", "description": "Date YYYY-MM-DD (default: today)"},
                    },
                    "required": ["rooms"],
                },
            },
        ]

    async def call_tool(self, name: str, arguments: Dict[str, Any]) -> str:
        if name == "o365_calendar_get_events":
            return await self._call_get_events(arguments)
        elif name == "o365_resolve_contact":
            return await self._call_resolve_contact(arguments)
        elif name == "o365_calendar_find_free_slots":
            return await self._call_find_free_slots(arguments)
        elif name == "o365_org_chart":
            return await self._call_org_chart(arguments)
        elif name == "o365_list_rooms":
            return await self._call_list_rooms(arguments)
        elif name == "o365_get_room_availability":
            return await self._call_room_availability(arguments)
        else:
            return f"Unknown tool: {name}"

    # ------------------------------------------------------------------
    # o365_list_rooms
    # ------------------------------------------------------------------

    def _filter_rooms(self, rooms: list) -> list:
        """Apply domain and exclude filters from instance config."""
        if self.room_domain_filter:
            rooms = [r for r in rooms
                     if r.get("emailAddress", "").lower().endswith(self.room_domain_filter.lower())]
        if self.room_exclude_patterns:
            excl = [p.lower() for p in self.room_exclude_patterns]
            rooms = [r for r in rooms
                     if not any(ex in r.get("displayName", "").lower() for ex in excl)]
        return rooms

    async def _call_list_rooms(self, arguments: Dict[str, Any]) -> str:
        from office_con.msgraph.places_handler import PlacesHandler
        ph = PlacesHandler(self.graph)
        rooms = self._filter_rooms(await ph.get_rooms_async())
        name_filter = arguments.get("filter", "").lower()
        if name_filter:
            rooms = [r for r in rooms if name_filter in r.get("displayName", "").lower()]
        result = [
            {
                "name": r.get("displayName", ""),
                "capacity": r.get("capacity"),
                "building": r.get("building"),
                "floor": r.get("floorNumber"),
            }
            for r in rooms
        ]
        return json.dumps(result, default=str, indent=2)

    # ------------------------------------------------------------------
    # o365_get_room_availability
    # ------------------------------------------------------------------

    async def _call_room_availability(self, arguments: Dict[str, Any]) -> str:
        from office_con.msgraph.places_handler import PlacesHandler
        from office_con.msgraph.mailbox_settings_handler import MailboxSettingsHandler
        from zoneinfo import ZoneInfo

        _WIN_TZ = {
            "W. Europe Standard Time": "Europe/Berlin",
            "Central European Standard Time": "Europe/Berlin",
            "Romance Standard Time": "Europe/Paris",
            "GMT Standard Time": "Europe/London",
            "Eastern Standard Time": "America/New_York",
            "Central Standard Time": "America/Chicago",
            "Pacific Standard Time": "America/Los_Angeles",
            "China Standard Time": "Asia/Shanghai",
            "India Standard Time": "Asia/Kolkata",
        }

        # User timezone
        mbs = MailboxSettingsHandler(self.graph)
        settings = await mbs.get_mailbox_settings_async()
        win_tz = settings.get("timeZone", "W. Europe Standard Time")
        iana_tz = _WIN_TZ.get(win_tz, win_tz)
        local_tz = ZoneInfo(iana_tz)

        # Date
        date_str = arguments.get("date", "")
        if date_str:
            target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        else:
            target_date = datetime.now(local_tz).date()

        # Match rooms (apply domain/exclude filters)
        ph = PlacesHandler(self.graph)
        all_rooms = self._filter_rooms(await ph.get_rooms_async())
        queries = arguments.get("rooms", [])
        matched = []
        for q in queries:
            q_lower = q.lower()
            for r in all_rooms:
                if q_lower in r.get("displayName", "").lower() and r not in matched:
                    matched.append(r)

        if not matched:
            return json.dumps({"error": "No matching rooms found"})

        # Schedule query
        emails = [r["emailAddress"] for r in matched]
        body = {
            "schedules": emails,
            "startTime": {"dateTime": f"{target_date.isoformat()}T07:00:00", "timeZone": win_tz},
            "endTime": {"dateTime": f"{target_date.isoformat()}T20:00:00", "timeZone": win_tz},
            "availabilityViewInterval": 30,
        }
        token = await self.graph.get_access_token_async()
        resp = await self.graph.run_async(
            url=self.graph.msg_endpoint + "me/calendar/getSchedule",
            method="POST", json=body, token=token,
        )
        schedules = resp.json().get("value", [])
        email_to_name = {r["emailAddress"].lower(): r["displayName"] for r in matched}

        result = []
        for sched in schedules:
            email = sched.get("scheduleId", "").lower()
            room_name = email_to_name.get(email, email)
            bookings = []
            for item in sched.get("scheduleItems", []):
                s = item.get("start", {}).get("dateTime", "")
                e = item.get("end", {}).get("dateTime", "")
                if s and e:
                    st = datetime.fromisoformat(s.rstrip("Z")).replace(
                        tzinfo=timezone.utc).astimezone(local_tz)
                    en = datetime.fromisoformat(e.rstrip("Z")).replace(
                        tzinfo=timezone.utc).astimezone(local_tz)
                    booking = {
                        "start": st.strftime("%H:%M"),
                        "end": en.strftime("%H:%M"),
                        "status": item.get("status", "busy"),
                    }
                    if self.show_room_booking_names:
                        booking["subject"] = item.get("subject", "")
                    bookings.append(booking)

            free = self._compute_free(bookings)
            result.append({
                "room": room_name,
                "date": target_date.isoformat(),
                "timezone": iana_tz,
                "bookings": bookings,
                "free_slots": free,
            })
        return json.dumps(result, default=str, indent=2)

    @staticmethod
    def _compute_free(bookings: List[Dict]) -> List[Dict]:
        """Compute merged free slots between 07:00-20:00."""
        busy = set()
        for b in bookings:
            sh, sm = map(int, b["start"].split(":"))
            eh, em = map(int, b["end"].split(":"))
            t = sh * 60 + sm
            end = eh * 60 + em
            while t < end:
                busy.add(t)
                t += 30
        free = []
        t = 7 * 60
        while t < 20 * 60:
            if t not in busy:
                h, m = divmod(t, 60)
                nh, nm = divmod(t + 30, 60)
                free.append({"start": f"{h:02d}:{m:02d}", "end": f"{nh:02d}:{nm:02d}"})
            t += 30
        if not free:
            return []
        merged = [free[0].copy()]
        for slot in free[1:]:
            if merged[-1]["end"] == slot["start"]:
                merged[-1]["end"] = slot["end"]
            else:
                merged.append(slot.copy())
        return merged

    # ------------------------------------------------------------------
    # o365_calendar_get_events (existing logic, extracted)
    # ------------------------------------------------------------------

    async def _call_get_events(self, arguments: Dict[str, Any]) -> str:
        try:
            query = arguments.get("query", "").strip()
            start_str = arguments.get("start_date")
            end_str = arguments.get("end_date")
            limit = arguments.get("limit", 50)

            if start_str:
                start_date = datetime.strptime(start_str, "%Y-%m-%d")
            else:
                start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

            if end_str:
                end_date = datetime.strptime(end_str, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
            else:
                # Wider range when searching, shorter for listing
                days = 180 if query else 14
                end_date = (start_date + timedelta(days=days)).replace(hour=23, minute=59, second=59)

            from office_con.msgraph.calendar_handler import CalendarHandler

            handler = CalendarHandler(self.graph)

            # When searching, fetch more from the API then filter locally
            fetch_limit = 500 if query else limit
            event_list = await handler.get_events_async(
                start_date=start_date,
                end_date=end_date,
                limit=fetch_limit,
            )

            events = event_list.events or []

            # Filter by query (case-insensitive substring on subject)
            if query:
                pattern = re.compile(re.escape(query), re.IGNORECASE)
                events = [ev for ev in events if pattern.search(ev.subject or "")]

            if not events:
                msg = f"No events found between {start_date.strftime('%Y-%m-%d')} and {end_date.strftime('%Y-%m-%d')}"
                if query:
                    msg += f" matching '{query}'"
                return msg + "."

            # Trim to requested limit
            events = events[:limit]

            # Format
            header = f"Found {len(events)} event(s)"
            if query:
                header += f" matching '{query}'"
            header += f" ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}):\n"
            lines = [header]

            for ev in events:
                start_fmt = ev.start_time.strftime("%a %b %d, %H:%M")
                end_fmt = ev.end_time.strftime("%H:%M")
                line = f"- {ev.subject}  ({start_fmt} – {end_fmt})"
                if ev.location:
                    line += f"  | Location: {ev.location}"
                if ev.attendees:
                    names = [a.name or a.email or "" for a in ev.attendees]
                    line += f"  | Attendees ({len(names)}): {', '.join(names)}"
                if ev.is_online_meeting and ev.online_meeting_url:
                    line += f"  | Teams link: {ev.online_meeting_url}"
                lines.append(line)

            return "\n".join(lines)

        except Exception as e:
            logger.error(f"[Office365MCP] Error fetching events: {e}")
            return f"Error fetching calendar events: {e}"

    # ------------------------------------------------------------------
    # o365_resolve_contact
    # ------------------------------------------------------------------

    # ── Fuzzy name helpers ──────────────────────────────────────────

    _UMLAUT_MAP = str.maketrans({
        "ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss",
        "Ä": "Ae", "Ö": "Oe", "Ü": "Ue",
    })

    @staticmethod
    def _normalize(text: str) -> str:
        """Lowercase, expand umlauts, strip accents & non-alpha."""
        text = text.translate(Office365MCPServer._UMLAUT_MAP)
        # strip remaining accents (é → e, etc.)
        text = unicodedata.normalize("NFD", text)
        text = "".join(c for c in text if unicodedata.category(c) != "Mn")
        return re.sub(r"[^a-z0-9 ]", "", text.lower()).strip()

    @classmethod
    def _fuzzy_score(cls, query: str, candidate: str) -> float:
        """Score 0..1 for how well *query* matches *candidate*.

        Uses token-level matching: each query token is scored against each
        candidate token individually, so single-surname queries like "sauter"
        correctly match "Jörg Sautter" (0.92) even though the full strings
        differ in length.  Name order doesn't matter.
        """
        nq = cls._normalize(query)
        nc = cls._normalize(candidate)
        if not nq or not nc:
            return 0.0

        # exact substring → 1.0
        if nq in nc:
            return 1.0

        q_tokens = nq.split()
        c_tokens = nc.split()

        # 1. Token-level: for each query token, find the best-matching candidate token
        if q_tokens and c_tokens:
            token_scores = []
            for qt in q_tokens:
                best = 0.0
                for ct in c_tokens:
                    if qt in ct:
                        # query token is substring of candidate token (e.g. "saut" in "sautter")
                        s = 1.0
                    elif len(qt) >= 3 and len(ct) >= 3:
                        s = SequenceMatcher(None, qt, ct).ratio()
                        # Penalize large length differences (e.g. "gusenbauer" vs "bauer")
                        len_ratio = min(len(qt), len(ct)) / max(len(qt), len(ct))
                        if len_ratio < 0.6:
                            s *= len_ratio
                    else:
                        s = 0.0
                    best = max(best, s)
                token_scores.append(best)
            token_avg = sum(token_scores) / len(token_scores)
        else:
            token_avg = 0.0

        # 2. Sorted-token-set ratio (catches full-name vs full-name)
        set_score = SequenceMatcher(
            None, " ".join(sorted(q_tokens)), " ".join(sorted(c_tokens))
        ).ratio()

        return max(token_avg, set_score)

    @staticmethod
    def _format_user_details(u: "DirectoryUser", by_id: Dict[str, Any]) -> List[Tuple[str, str]]:
        """Return list of (label, value) pairs for all non-None user fields.

        Excludes: gender, id, account_enabled, has_image, object_type, external,
        and fields already shown in the main line (display_name, email).
        Works with both DirectoryUser and CompanyUser instances.
        """
        from office_con.db.company_dir import CompanyUser

        details: List[Tuple[str, str]] = []
        # DirectoryUser fields
        if u.job_title:
            details.append(("Title", u.job_title))
        if u.department:
            details.append(("Dept", u.department))
        if u.mobile_phone:
            details.append(("Phone", u.mobile_phone))
        if u.office_location:
            details.append(("Office", u.office_location))
        # Manager via ID lookup
        if u.manager_id:
            mgr = by_id.get(u.manager_id)
            if mgr:
                details.append(("Manager", mgr.display_name))
        # CompanyUser-specific fields
        if isinstance(u, CompanyUser):
            if u.company:
                details.append(("Company", u.company))
            if u.building:
                details.append(("Building", u.building))
            if u.room_name:
                details.append(("Room", u.room_name))
            if u.street:
                details.append(("Street", u.street))
            if u.zip:
                details.append(("Zip", u.zip))
            if u.city:
                details.append(("City", u.city))
            if u.country:
                details.append(("Country", u.country))
            if u.manager_email:
                details.append(("Manager Email", u.manager_email))
            if u.join_date:
                details.append(("Joined", u.join_date))
            if u.birth_date:
                details.append(("Birthday", u.birth_date))
            if u.termination_date:
                details.append(("Termination", u.termination_date))
            if u.guessed_fields:
                guessed_names = [k for k, v in u.guessed_fields.items() if v]
                if guessed_names:
                    details.append(("Guessed", ", ".join(guessed_names)))
        return details

    # ── Main resolve logic ───────────────────────────────────────

    async def _call_resolve_contact(self, arguments: Dict[str, Any]) -> str:
        try:
            raw_query = arguments.get("query", "")
            # Accept string or list
            if isinstance(raw_query, str):
                queries = [q.strip() for q in raw_query.split(",") if q.strip()] if "," in raw_query else [raw_query.strip()]
            elif isinstance(raw_query, list):
                queries = [str(q).strip() for q in raw_query if str(q).strip()]
            else:
                queries = [str(raw_query).strip()]

            if not queries:
                return "Please provide a name to search for."

            min_score = 0.75

            from office_con.msgraph.calendar_handler import CalendarHandler

            calendar = CalendarHandler(self.graph)

            # 1. Load all users once (prefer pre-loaded company dir, fall back to Graph API)
            from office_con.msgraph.directory_handler import DirectoryHandler
            cd_users = self._company_dir.users if self._company_dir else []
            logger.info(
                "[RESOLVE] company_dir=%s, company_dir.users count=%d, is_live=%s, is_populated=%s",
                bool(self._company_dir),
                len(cd_users),
                getattr(self._company_dir, '_live', '?') if self._company_dir else 'N/A',
                getattr(getattr(self._company_dir, 'data', None), 'is_populated', '?') if self._company_dir else 'N/A',
            )
            if cd_users:
                class _Wrap:
                    def __init__(self, users: list) -> None: self.users = users
                user_list: Any = _Wrap(cd_users)
                logger.info("[RESOLVE] Using company_dir (%d users)", len(cd_users))
            else:
                directory = DirectoryHandler(self.graph)
                user_list = await directory.get_all_users_async()
                logger.info("[RESOLVE] Fallback to DirectoryHandler (%d users)", len(user_list.users))

            # 2. Match each query against all users with fuzzy scoring
            all_matched_users = set()  # track unique users for photo fetching
            query_results: List[Tuple[str, List[Tuple[float, Any]]]] = []

            for q in queries:
                scored = []
                top_near_misses = []  # track best non-matching scores for debugging
                for u in user_list.users:
                    searchable = " ".join(
                        s for s in [u.display_name, u.given_name, u.surname] if s
                    )
                    score = self._fuzzy_score(q, searchable)
                    if score >= min_score:
                        scored.append((score, u))
                    elif score >= 0.4:
                        top_near_misses.append((score, u.display_name, searchable))
                scored.sort(key=lambda x: x[0], reverse=True)
                # Cap at 5 per query
                scored = scored[:5]
                query_results.append((q, scored))
                logger.info(
                    "[RESOLVE] query=%r → %d matches (total users searched: %d)",
                    q, len(scored), len(user_list.users),
                )
                if scored:
                    for sc, u in scored[:3]:
                        logger.info("[RESOLVE]   match: %.2f %s <%s>", sc, u.display_name, u.email)
                elif top_near_misses:
                    top_near_misses.sort(reverse=True)
                    for sc, dn, searchable in top_near_misses[:3]:
                        logger.info("[RESOLVE]   near-miss: %.2f %s (searchable=%r)", sc, dn, searchable)
                for _, u in scored:
                    if u.email:
                        all_matched_users.add(u.email.lower())

            # 3. Count meeting frequency over last 90 days
            now = datetime.now()
            start_90 = now - timedelta(days=90)
            event_list = await calendar.get_events_async(
                start_date=start_90,
                end_date=now,
                limit=500,
            )
            email_counter: Counter = Counter()
            for ev in event_list.events or []:
                for att in ev.attendees:
                    if att.email:
                        email_counter[att.email.lower()] += 1

            # Build id-based lookup for manager resolution
            by_id: Dict[str, Any] = {}
            for u in user_list.users:
                if u.id:
                    by_id[u.id] = u

            # 5. Format hierarchical output
            sections: List[str] = []
            total_matches = sum(len(scored) for _, scored in query_results)

            for q, scored in query_results:
                if not scored:
                    sections.append(f"### {q}\nNo matches found.")
                    continue

                lines = [f"### {q}"]
                for score, u in scored:
                    email_lower = (u.email or "").lower()
                    freq = email_counter.get(email_lower, 0)
                    match_label = "exact" if score >= 0.99 else f"fuzzy {score:.0%}"

                    line = f"- **{u.display_name}**"
                    if u.email:
                        line += f" <{u.email}>"

                    # Emit all non-None fields as pipe-separated details
                    detail_fields = self._format_user_details(u, by_id)
                    for label, val in detail_fields:
                        line += f"  | {label}: {val}"

                    line += f"  | {freq} meetings (90d)"
                    line += f"  | _{match_label}_"
                    lines.append(line)
                sections.append("\n".join(lines))

            header = f"Found {total_matches} contact(s) for {len(queries)} search term(s):\n"
            return header + "\n\n".join(sections)

        except Exception as e:
            logger.error(f"[Office365MCP] Error resolving contact: {e}")
            return f"Error resolving contact: {e}"

    # ------------------------------------------------------------------
    # o365_calendar_find_free_slots
    # ------------------------------------------------------------------

    async def _call_find_free_slots(self, arguments: Dict[str, Any]) -> str:
        try:
            attendee_emails = arguments.get("attendee_emails", [])
            if not attendee_emails:
                return "Please provide at least one attendee email."

            duration_minutes = arguments.get("duration_minutes", 60)
            start_hour = arguments.get("start_hour", 8)
            end_hour = arguments.get("end_hour", 18)

            # Parse date range
            now = datetime.now()
            if arguments.get("start_date"):
                start_date = datetime.strptime(arguments["start_date"], "%Y-%m-%d")
            else:
                start_date = (now + timedelta(days=1)).replace(
                    hour=0, minute=0, second=0, microsecond=0
                )

            if arguments.get("end_date"):
                end_date = datetime.strptime(arguments["end_date"], "%Y-%m-%d")
            else:
                # Default: 5 business days out
                end_date = start_date
                bdays = 0
                while bdays < 5:
                    end_date += timedelta(days=1)
                    if end_date.weekday() < 5:  # Mon-Fri
                        bdays += 1

            # Set end to end-of-day
            end_dt = end_date.replace(hour=23, minute=59, second=59)
            start_dt = start_date.replace(hour=0, minute=0, second=0, microsecond=0)

            # Include the logged-in user in the schedule query
            user_email = getattr(self.graph, "email", None)
            all_emails = list(attendee_emails)
            if user_email and user_email.lower() not in [e.lower() for e in all_emails]:
                all_emails.insert(0, user_email)

            from office_con.msgraph.calendar_handler import CalendarHandler

            handler = CalendarHandler(self.graph)
            interval = 15  # 15-minute granularity

            # Get the user's timezone so slots are in local time
            user_tz = await handler.get_user_timezone_async()

            schedules = await handler.get_schedule_async(
                emails=all_emails,
                start=start_dt,
                end=end_dt,
                interval=interval,
                timezone=user_tz,
            )

            if not schedules:
                return "Could not retrieve availability. Check that the email addresses are valid."

            # Parse availability views — each char is one interval slot
            # 0=free, 1=tentative, 2=busy, 3=oof, 4=workingElsewhere
            views = []
            for sched in schedules:
                av = sched.get("availabilityView", "")
                views.append(av)

            if not views:
                return "No availability data returned."

            # All views should be the same length; use the shortest as reference
            n_slots = min(len(v) for v in views)

            # Find slots where ALL participants are free (0) or tentative (1)
            free_mask = []
            for i in range(n_slots):
                all_free = all(v[i] in ("0", "1") for v in views)
                free_mask.append(all_free)

            # Convert slot indices to datetimes and filter by working hours
            slots: list[tuple[datetime, datetime]] = []
            run_start = None

            for i in range(n_slots):
                slot_time = start_dt + timedelta(minutes=i * interval)
                in_working_hours = (
                    slot_time.weekday() < 5
                    and start_hour <= slot_time.hour < end_hour
                )

                if free_mask[i] and in_working_hours:
                    if run_start is None:
                        run_start = slot_time
                else:
                    if run_start is not None:
                        run_end = start_dt + timedelta(minutes=i * interval)
                        slots.append((run_start, run_end))
                        run_start = None

            # Close any trailing run
            if run_start is not None:
                run_end = start_dt + timedelta(minutes=n_slots * interval)
                slots.append((run_start, run_end))

            # Filter by minimum duration
            min_td = timedelta(minutes=duration_minutes)
            slots = [(s, e) for s, e in slots if (e - s) >= min_td]

            if not slots:
                return (
                    f"No free slots of {duration_minutes} minutes found between "
                    f"{start_date.strftime('%Y-%m-%d')} and {end_date.strftime('%Y-%m-%d')} "
                    f"({start_hour:02d}:00–{end_hour:02d}:00)."
                )

            # Format output
            attendee_names = ", ".join(attendee_emails)
            lines = [
                f"Free slots ({duration_minutes}+ min) with {attendee_names} "
                f"({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}, "
                f"{start_hour:02d}:00–{end_hour:02d}:00, {user_tz}):\n"
            ]

            for s, e in slots:
                dur = int((e - s).total_seconds() // 60)
                if s.date() == e.date():
                    lines.append(
                        f"- {s.strftime('%a %b %d')}: {s.strftime('%H:%M')}–{e.strftime('%H:%M')} ({dur} min)"
                    )
                else:
                    lines.append(
                        f"- {s.strftime('%a %b %d %H:%M')} – {e.strftime('%a %b %d %H:%M')} ({dur} min)"
                    )

            return "\n".join(lines)

        except Exception as e:
            logger.error(f"[Office365MCP] Error finding free slots: {e}")
            return f"Error finding free slots: {e}"

    # ------------------------------------------------------------------
    # o365_org_chart
    # ------------------------------------------------------------------

    async def _call_org_chart(self, arguments: Dict[str, Any]) -> str:
        try:
            email = arguments.get("email", "").strip()
            if not email:
                return "Please provide an email address."

            depth_up = arguments.get("depth_up", 3)
            depth_down = arguments.get("depth_down", 1)

            # Use pre-loaded company directory if available (faster, richer data,
            # app-level permissions). Fall back to direct MS Graph fetch when
            # company dir is empty (e.g. async population not yet complete).
            cd_users = self._company_dir.users if self._company_dir else []
            if cd_users:
                users = cd_users
            else:
                from office_con.msgraph.directory_handler import DirectoryHandler
                directory = DirectoryHandler(self.graph)
                user_list = await directory.get_all_users_async()
                users = user_list.users or []

            # Build lookup indexes
            by_email: Dict[str, Any] = {}
            by_id: Dict[str, Any] = {}
            children: Dict[str, List[Any]] = {}  # manager_id -> list of direct reports

            for u in users:
                if u.email:
                    by_email[u.email.lower()] = u
                if u.id:
                    by_id[u.id] = u
                    if u.manager_id:
                        children.setdefault(u.manager_id, []).append(u)

            target = by_email.get(email.lower())
            if not target:
                return f"No user found with email: {email}"

            # 1. Walk manager chain upward
            manager_chain: List[Any] = []
            current = target
            for _ in range(depth_up):
                mgr_id = current.manager_id
                if not mgr_id:
                    break
                mgr = by_id.get(mgr_id)
                if not mgr:
                    break
                manager_chain.append(mgr)
                current = mgr

            # 2. Collect direct reports recursively
            def _collect_reports(user_id: str, depth: int) -> List[tuple]:
                if depth <= 0:
                    return []
                reports = children.get(user_id, [])
                reports.sort(key=lambda u: u.display_name or "")
                result = []
                for r in reports:
                    sub = _collect_reports(r.id, depth - 1) if r.id else []
                    result.append((r, sub))
                return result

            direct_reports = _collect_reports(target.id, depth_down) if target.id else []

            # 3. Build output — use contact code blocks for rich rendering
            def _contact_block(u) -> str:
                """Build a ``` contact ``` fenced block for a user."""
                block_lines = [
                    f"name: {u.display_name}",
                    f"email: {u.email}" if u.email else None,
                    f"title: {u.job_title}" if u.job_title else None,
                    f"department: {u.department}" if u.department else None,
                    f"phone: {u.mobile_phone}" if u.mobile_phone else None,
                    f"location: {u.office_location}" if u.office_location else None,
                ]
                body = "\n".join(line for line in block_lines if line)
                return f"```contact\n{body}\n```"

            lines = [f"## Org Chart for {target.display_name}\n"]

            # Manager chain (top-down)
            if manager_chain:
                lines.append("### Reporting Line (upward)\n")
                for mgr in reversed(manager_chain):
                    lines.append(_contact_block(mgr))
                    lines.append("↓")
                lines.append(_contact_block(target) + " ← _(target)_")
                lines.append("")

            # Target details (if no manager chain shown above)
            if not manager_chain:
                lines.append("### Person\n")
                lines.append(_contact_block(target))
                lines.append("")

            # Direct reports
            if direct_reports:
                lines.append(f"### Direct Reports ({len(direct_reports)})\n")
                for r, subs in direct_reports:
                    lines.append(_contact_block(r))
                    for sub, _ in subs:
                        lines.append(f"  - **{sub.display_name}** — {sub.job_title or 'N/A'} <{sub.email or 'N/A'}>")
            else:
                lines.append("### Direct Reports\nNo direct reports found in the directory.")

            return "\n".join(lines)

        except Exception as e:
            logger.error(f"[Office365MCP] Error getting org chart: {e}")
            return f"Error getting org chart: {e}"
