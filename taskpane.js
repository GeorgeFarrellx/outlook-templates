/*
 * Quick Templates – Outlook add-in task pane.
 *
 * Inserts templates at the cursor with body.setSelectedDataAsync, so they land
 * where you're typing instead of at the top of the email.
 * Templates are saved, compressed, in the add-in's roaming settings, which live
 * in your mailbox and have a 32 KB limit.
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Constants and state
  // ---------------------------------------------------------------------------
  var STORE_KEY = 'qt.templates.v1';
  var SETTINGS_KEY = 'qt.settings.v1';
  var LOCAL_BACKUP_KEY = 'qt.localBackup.v1';
  var COLLAPSED_KEY = 'qt.collapsed.v1';
  var STORAGE_LIMIT = 30000; // Outlook's hard limit is 32 KB; keep some headroom.
  var PLACEHOLDER_SOURCE = '\\{([A-Za-z0-9][A-Za-z0-9 _-]{0,40})\\}';
  var DEFAULT_SETTINGS = { fontFamily: 'Aptos', fontSize: '12', replaceSubject: true };
  var COMMON_PLACEHOLDERS = ['FirstName', 'Amount', 'Today'];

  var state = {
    templates: [],
    settings: clone(DEFAULT_SETTINGS),
    view: 'list',
    filling: null,          // { template, names, confirmBlanks }
    editing: null,          // { id, snapshot, confirmDiscard, confirmDelete }
    savedRange: null,
    confirmReplaceAll: false,
    collapsed: readLocal(COLLAPSED_KEY, {})
  };

  // ---------------------------------------------------------------------------
  // Small helpers
  // ---------------------------------------------------------------------------
  function $(id) { return document.getElementById(id); }

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        if (k === 'text') node.textContent = attrs[k];
        else if (k === 'className') node.className = attrs[k];
        else if (k.indexOf('on') === 0 && typeof attrs[k] === 'function') node.addEventListener(k.slice(2), attrs[k]);
        else node.setAttribute(k, attrs[k]);
      });
    }
    (children || []).forEach(function (c) {
      if (c == null) return;
      node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return node;
  }

  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function escapeRegExp(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  function newId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

  function naturalCompare(a, b) {
    return String(a).localeCompare(String(b), 'en-GB', { numeric: true, sensitivity: 'base' });
  }

  function readLocal(key, fallback) {
    try {
      var v = window.localStorage.getItem(key);
      return v == null ? fallback : JSON.parse(v);
    } catch (e) { return fallback; }
  }

  function writeLocal(key, value) {
    try { window.localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* ignore */ }
  }

  var toastTimer = null;
  function toast(message, opts) {
    opts = opts || {};
    var box = $('toast');
    $('toast-text').textContent = message;
    box.className = 'toast' + (opts.error ? ' error' : '');
    var action = $('toast-action');
    if (opts.actionLabel && opts.onAction) {
      action.textContent = opts.actionLabel;
      action.hidden = false;
      action.onclick = function () { hideToast(); opts.onAction(); };
    } else {
      action.hidden = true;
      action.onclick = null;
    }
    box.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(hideToast, opts.duration || (opts.actionLabel ? 9000 : 4000));
  }
  function hideToast() { $('toast').hidden = true; }

  // ---------------------------------------------------------------------------
  // Office.js promise wrappers
  // ---------------------------------------------------------------------------
  function officeCall(fn) {
    return new Promise(function (resolve, reject) {
      fn(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(result.error || new Error('Outlook reported an error.'));
      });
    });
  }

  function currentItem() {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.body || typeof item.body.setSelectedDataAsync !== 'function') return null;
    return item;
  }

  // ---------------------------------------------------------------------------
  // Storage (roaming settings, compressed)
  // ---------------------------------------------------------------------------
  function encodeTemplates(templates) {
    var slim = templates.map(function (t) {
      return { id: t.id, name: t.name, category: t.category || '', subject: t.subject || '', body: t.body || '' };
    });
    return LZString.compressToBase64(JSON.stringify({ v: 1, templates: slim }));
  }

  function decodeTemplates(payload) {
    if (typeof payload !== 'string' || !payload) return [];
    var json = LZString.decompressFromBase64(payload);
    if (!json) throw new Error('Saved templates could not be read.');
    var data = JSON.parse(json);
    return Array.isArray(data.templates) ? data.templates : [];
  }

  function storageUsed(templates, settings) {
    return encodeTemplates(templates).length + JSON.stringify(settings).length + 40;
  }

  function loadFromMailbox() {
    var rs = Office.context.roamingSettings;
    state.settings = Object.assign(clone(DEFAULT_SETTINGS), rs.get(SETTINGS_KEY) || {});
    state.templates = decodeTemplates(rs.get(STORE_KEY));
  }

  /** Saves templates (and settings) to the mailbox. Resolves once Outlook confirms. */
  function persist(templates, settings) {
    settings = settings || state.settings;
    var rs = Office.context.roamingSettings;
    var payload = encodeTemplates(templates);
    var used = payload.length + JSON.stringify(settings).length + 40;
    if (used > STORAGE_LIMIT) {
      var err = new Error('Template storage is full (' + Math.round(used / STORAGE_LIMIT * 100) +
        '% needed). Shorten or delete some templates, or remove heavy formatting.');
      err.code = 'full';
      return Promise.reject(err);
    }
    var previousPayload = rs.get(STORE_KEY);
    var previousSettings = rs.get(SETTINGS_KEY);
    rs.set(STORE_KEY, payload);
    rs.set(SETTINGS_KEY, settings);
    return officeCall(function (cb) { rs.saveAsync(cb); }).then(function () {
      state.templates = templates;
      state.settings = settings;
      writeLocal(LOCAL_BACKUP_KEY, { savedAt: new Date().toISOString(), payload: payload });
    }, function (error) {
      // Put the old values back so memory matches the mailbox.
      if (previousPayload === undefined || previousPayload === null) rs.remove(STORE_KEY); else rs.set(STORE_KEY, previousPayload);
      if (previousSettings === undefined || previousSettings === null) rs.remove(SETTINGS_KEY); else rs.set(SETTINGS_KEY, previousSettings);
      var err = new Error('Outlook couldn\'t save: ' + (error && error.message ? error.message : 'unknown error'));
      throw err;
    });
  }

  // ---------------------------------------------------------------------------
  // HTML cleaning (Word/Outlook HTML -> small, tidy HTML)
  // ---------------------------------------------------------------------------
  var DROP_TAGS = ['script', 'style', 'meta', 'link', 'title', 'head', 'xml', 'iframe', 'object', 'embed',
    'form', 'input', 'button', 'select', 'textarea', 'svg', 'noscript', 'template', 'canvas', 'video', 'audio'];
  var ALLOWED_TAGS = ['div', 'p', 'br', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'sup', 'sub', 'a', 'ul', 'ol',
    'li', 'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'img', 'span', 'blockquote', 'hr',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
  var BLOCK_TAGS = ['div', 'p', 'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th',
    'blockquote', 'hr', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
  var ALLOWED_ATTRS = {
    a: ['href', 'title'],
    img: ['src', 'alt', 'width', 'height'],
    table: ['border', 'cellpadding', 'cellspacing', 'width'],
    td: ['colspan', 'rowspan', 'width', 'valign', 'align'],
    th: ['colspan', 'rowspan', 'width', 'valign', 'align'],
    div: ['align'],
    p: ['align']
  };
  var COMMON_STYLES = ['font-weight', 'font-style', 'text-decoration', 'text-decoration-line', 'color',
    'background-color', 'text-align', 'vertical-align'];
  var BOX_STYLES = ['border', 'border-top', 'border-right', 'border-bottom', 'border-left', 'border-collapse',
    'padding', 'width', 'background'];
  var DEFAULT_COLOURS = /^(windowtext|black|#000|#000000|auto|inherit|initial|rgb\(0,\s*0,\s*0\))$/i;
  var DEFAULT_BACKGROUNDS = /^(white|#fff|#ffffff|transparent|none|auto|inherit|initial|rgb\(255,\s*255,\s*255\))$/i;

  function isBlock(node) {
    return node && node.nodeType === 1 && BLOCK_TAGS.indexOf(node.tagName.toLowerCase()) !== -1;
  }

  function unwrap(node) {
    var parent = node.parentNode;
    if (!parent) return;
    while (node.firstChild) parent.insertBefore(node.firstChild, node);
    parent.removeChild(node);
  }

  function filterStyle(tag, styleText) {
    if (!styleText) return '';
    var allowed = COMMON_STYLES.concat(['table', 'td', 'th'].indexOf(tag) !== -1 ? BOX_STYLES : []);
    var kept = [];
    styleText.split(';').forEach(function (decl) {
      var i = decl.indexOf(':');
      if (i < 1) return;
      var prop = decl.slice(0, i).trim().toLowerCase();
      var value = decl.slice(i + 1).trim();
      if (!value || allowed.indexOf(prop) === -1) return;
      if (prop === 'color' && DEFAULT_COLOURS.test(value)) return;
      if ((prop === 'background-color' || prop === 'background') && DEFAULT_BACKGROUNDS.test(value)) return;
      if (prop === 'font-weight' && /^(normal|400)$/i.test(value)) return;
      if (prop === 'font-style' && /^normal$/i.test(value)) return;
      if ((prop === 'text-decoration' || prop === 'text-decoration-line') && /^none$/i.test(value)) return;
      if (prop === 'text-align' && /^(left|start)$/i.test(value)) return;
      if (prop === 'vertical-align' && /^baseline$/i.test(value)) return;
      kept.push(prop + ':' + value);
    });
    return kept.join(';');
  }

  function hasVisibleContent(node) {
    if (node.querySelector && node.querySelector('img,table,hr')) return true;
    return (node.textContent || '').replace(/[\s ​]/g, '') !== '';
  }

  function normaliseBraces(html) {
    return html.replace(/\{([^{}<>]{1,80})\}/g, function (m) { return m.replace(/&nbsp;| /g, ' '); });
  }

  /**
   * Cleans pasted/imported HTML. Returns { html, imagesRemoved }.
   */
  function cleanHtml(input) {
    var doc = new DOMParser().parseFromString('<!DOCTYPE html><html><body>' + (input || '') + '</body></html>', 'text/html');
    // If the input was a whole HTML document, DOMParser still gives us its body.
    var root = doc.body;
    var imagesRemoved = 0;

    // Comments (including Word's conditional comments)
    var walker = doc.createTreeWalker(root, NodeFilter.SHOW_COMMENT, null);
    var comments = [];
    while (walker.nextNode()) comments.push(walker.currentNode);
    comments.forEach(function (c) { c.parentNode.removeChild(c); });

    // Elements, deepest first
    var elements = Array.prototype.slice.call(root.querySelectorAll('*')).reverse();
    elements.forEach(function (node) {
      if (!node.parentNode) return;
      var tag = node.tagName.toLowerCase();
      if (DROP_TAGS.indexOf(tag) !== -1) { node.parentNode.removeChild(node); return; }
      if (tag === 'img') {
        var src = node.getAttribute('src') || '';
        if (!/^https?:\/\//i.test(src)) { imagesRemoved++; node.parentNode.removeChild(node); return; }
      }
      if (ALLOWED_TAGS.indexOf(tag) === -1) { unwrap(node); return; }

      var keepAttrs = ALLOWED_ATTRS[tag] || [];
      var style = filterStyle(tag, node.getAttribute('style'));
      Array.prototype.slice.call(node.attributes).forEach(function (attr) {
        var name = attr.name.toLowerCase();
        if (keepAttrs.indexOf(name) === -1) node.removeAttribute(attr.name);
      });
      if (tag === 'a') {
        var href = node.getAttribute('href') || '';
        if (href && !/^(https?:|mailto:|tel:)/i.test(href)) node.removeAttribute('href');
      }
      if (style) node.setAttribute('style', style);

      if (tag === 'span' && node.attributes.length === 0) { unwrap(node); return; }
      if (tag === 'strong') { node = renameElement(node, 'b'); }
      if (tag === 'em') { node = renameElement(node, 'i'); }
      if (tag === 'p') { renameElement(node, 'div'); }
    });

    // Word breaks long lines inside a paragraph with CR/LF: treat as spaces.
    var textWalker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    var texts = [];
    while (textWalker.nextNode()) texts.push(textWalker.currentNode);
    texts.forEach(function (t) {
      if (/[\r\n\t]/.test(t.nodeValue)) t.nodeValue = t.nodeValue.replace(/[\r\n\t]+/g, ' ');
      if (/^[ ]*$/.test(t.nodeValue)) {
        var parent = t.parentNode;
        var ptag = parent.tagName ? parent.tagName.toLowerCase() : '';
        var prev = t.previousSibling, next = t.nextSibling;
        var structural = ['table', 'thead', 'tbody', 'tfoot', 'tr', 'ul', 'ol'].indexOf(ptag) !== -1;
        var besideBlock = isBlock(prev) || isBlock(next) ||
          ((!prev || !next) && (parent === root || isBlock(parent)));
        if (structural || besideBlock || t.nodeValue === '') parent.removeChild(t);
      }
    });

    // Unwrap plain wrapper divs that only contain other blocks (e.g. WordSection1).
    var changed = true;
    while (changed) {
      changed = false;
      Array.prototype.slice.call(root.querySelectorAll('div')).reverse().forEach(function (div) {
        if (div.attributes.length) return;
        var kids = Array.prototype.slice.call(div.childNodes);
        if (!kids.length) return;
        var onlyBlocks = kids.every(function (k) { return isBlock(k); });
        if (onlyBlocks) { unwrap(div); changed = true; }
      });
    }

    // Empty paragraphs become a single line break so they show as blank lines.
    root.querySelectorAll('div,h1,h2,h3,h4,h5,h6').forEach(function (block) {
      if (!hasVisibleContent(block) && !block.querySelector('div,table,ul,ol')) block.innerHTML = '<br>';
    });

    // Remove empty formatting tags left behind.
    Array.prototype.slice.call(root.querySelectorAll('b,i,u,s,strike,sup,sub,span,a')).reverse().forEach(function (n) {
      if (!n.childNodes.length) n.parentNode.removeChild(n);
    });

    // Trim blank lines at the start and end.
    trimBlankEdges(root);

    return { html: normaliseBraces(root.innerHTML), imagesRemoved: imagesRemoved };
  }

  function renameElement(node, newTag) {
    var replacement = node.ownerDocument.createElement(newTag);
    Array.prototype.slice.call(node.attributes).forEach(function (a) { replacement.setAttribute(a.name, a.value); });
    while (node.firstChild) replacement.appendChild(node.firstChild);
    node.parentNode.replaceChild(replacement, node);
    return replacement;
  }

  function isBlankLine(node) {
    if (!node) return false;
    if (node.nodeType === 3) return node.nodeValue.replace(/[\s ]/g, '') === '';
    if (node.nodeType !== 1) return true;
    var tag = node.tagName.toLowerCase();
    if (tag === 'br') return true;
    if (tag === 'div') return !hasVisibleContent(node);
    return false;
  }

  function trimBlankEdges(root) {
    while (root.firstChild && isBlankLine(root.firstChild)) root.removeChild(root.firstChild);
    while (root.lastChild && isBlankLine(root.lastChild)) root.removeChild(root.lastChild);
  }

  // ---------------------------------------------------------------------------
  // Preparing HTML/text for Outlook
  // ---------------------------------------------------------------------------
  function prepareForOutlook(html, settings) {
    var box = document.createElement('div');
    box.innerHTML = html;

    box.querySelectorAll('div,p,td,th,li').forEach(function (block) {
      if (!hasVisibleContent(block) && !block.querySelector('div,table,ul,ol')) block.innerHTML = '&nbsp;';
    });
    box.querySelectorAll('a').forEach(function (a) { a.setAttribute('id', 'LPNoLP'); });

    if (settings.fontFamily) {
      var size = parseFloat(settings.fontSize) || 12;
      var fontStyle = "font-family:'" + settings.fontFamily.replace(/['"<>;]/g, '') + "',sans-serif;font-size:" + size + 'pt';
      box.querySelectorAll('div,p,td,th,li').forEach(function (block) {
        var existing = block.getAttribute('style');
        block.setAttribute('style', existing ? existing + ';' + fontStyle : fontStyle);
      });
      // Wrap top-level inline content (text, bold, links…) in a styled span.
      var run = [];
      var flush = function () {
        if (!run.length) return;
        var hasText = run.some(function (n) { return (n.textContent || '').trim() !== '' || (n.nodeType === 1 && n.tagName === 'IMG'); });
        if (hasText) {
          var span = document.createElement('span');
          span.setAttribute('style', fontStyle);
          run[0].parentNode.insertBefore(span, run[0]);
          run.forEach(function (n) { span.appendChild(n); });
        }
        run = [];
      };
      Array.prototype.slice.call(box.childNodes).forEach(function (n) {
        if (isBlock(n)) flush(); else run.push(n);
      });
      flush();
    }
    return box.innerHTML;
  }

  function htmlToText(html) {
    var box = document.createElement('div');
    box.style.cssText = 'position:fixed;left:-10000px;top:0;width:600px;white-space:normal;';
    box.innerHTML = html;
    document.body.appendChild(box);
    var text = box.innerText;
    document.body.removeChild(box);
    return text.replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  function plainText(html) {
    var box = document.createElement('div');
    box.innerHTML = html || '';
    return (box.textContent || '').replace(/\s+/g, ' ').trim();
  }

  // ---------------------------------------------------------------------------
  // Placeholders
  // ---------------------------------------------------------------------------
  function findPlaceholders() {
    var names = [];
    Array.prototype.slice.call(arguments).forEach(function (s) {
      normaliseBraces(s || '').replace(new RegExp(PLACEHOLDER_SOURCE, 'g'), function (m, name) {
        name = name.trim().replace(/\s+/g, ' ');
        if (names.indexOf(name) === -1) names.push(name);
        return m;
      });
    });
    return names;
  }

  function fillPlaceholders(text, values, asHtml) {
    var out = normaliseBraces(text || '');
    Object.keys(values).forEach(function (name) {
      var value = (values[name] || '').trim();
      var pattern = escapeRegExp(name).replace(/ /g, '\\s+');
      if (!value) {
        // "Hi {FirstName}," with no name becomes "Hi," rather than "Hi ,"
        out = out.replace(new RegExp('(?: |&nbsp;|\\u00a0)\\{' + pattern + ' *\\}(?=[,.;:!?])', 'g'), '');
      }
      var replacement = asHtml ? escapeHtml(value) : value;
      out = out.replace(new RegExp('\\{' + pattern + ' *\\}', 'g'), function () { return replacement; });
    });
    return out;
  }

  function placeholderKey(name) { return name.toLowerCase().replace(/[^a-z0-9]/g, ''); }

  function firstNameFromDisplay(displayName) {
    var dn = (displayName || '').trim();
    if (!dn || dn.indexOf('@') !== -1) return '';
    if (dn.indexOf(',') !== -1) {
      var parts = dn.split(',');
      dn = (parts[1] || '').trim() || parts[0].trim();
    }
    dn = dn.replace(/^(mr|mrs|ms|miss|mx|dr|prof)\.?\s+/i, '');
    var first = dn.split(/\s+/)[0] || '';
    if (first === first.toUpperCase() && first.length > 1) first = first.charAt(0) + first.slice(1).toLowerCase();
    return first;
  }

  function autoValues(names) {
    var keys = names.map(placeholderKey);
    var needsRecipients = keys.some(function (k) { return ['firstname', 'clientfirstname', 'forename', 'fullname', 'recipientname'].indexOf(k) !== -1; });
    var item = currentItem();
    var recipientsPromise = (needsRecipients && item && item.to && item.to.getAsync)
      ? officeCall(function (cb) { item.to.getAsync(cb); }).catch(function () { return []; })
      : Promise.resolve([]);

    return recipientsPromise.then(function (recipients) {
      recipients = recipients || [];
      var firstNames = recipients.map(function (r) { return firstNameFromDisplay(r.displayName); }).filter(Boolean);
      var fullNames = recipients.map(function (r) { return (r.displayName || '').indexOf('@') === -1 ? (r.displayName || '').trim() : ''; }).filter(Boolean);
      var profileName = (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.displayName) || '';
      var today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

      var result = {};
      names.forEach(function (name) {
        var k = placeholderKey(name);
        if (['firstname', 'clientfirstname', 'forename'].indexOf(k) !== -1 && firstNames.length) {
          result[name] = { value: firstNames.length === 2 ? firstNames[0] + ' and ' + firstNames[1] : firstNames[0], source: 'from the To line' };
        } else if (['fullname', 'recipientname'].indexOf(k) !== -1 && fullNames.length) {
          result[name] = { value: fullNames[0], source: 'from the To line' };
        } else if (['today', 'date', 'todaysdate'].indexOf(k) !== -1) {
          result[name] = { value: today, source: "today's date" };
        } else if (['myname', 'yourname', 'sendername'].indexOf(k) !== -1 && profileName) {
          result[name] = { value: profileName, source: 'your name' };
        } else if (k === 'myfirstname' && profileName) {
          result[name] = { value: firstNameFromDisplay(profileName), source: 'your name' };
        }
      });
      return result;
    });
  }

  // ---------------------------------------------------------------------------
  // Views
  // ---------------------------------------------------------------------------
  function show(view) {
    ['list', 'fill', 'edit', 'settings'].forEach(function (v) { $('view-' + v).hidden = v !== view; });
    state.view = view;
    window.scrollTo(0, 0);
  }

  function categories() {
    var cats = [];
    state.templates.forEach(function (t) {
      var c = (t.category || '').trim();
      if (cats.indexOf(c) === -1) cats.push(c);
    });
    return cats.sort(function (a, b) {
      if (!a) return 1;
      if (!b) return -1;
      return naturalCompare(a, b);
    });
  }

  function renderCategoryFilter() {
    var select = $('category-filter');
    var current = select.value;
    select.innerHTML = '';
    select.appendChild(el('option', { value: '__all', text: 'All categories' }));
    categories().forEach(function (c) {
      select.appendChild(el('option', { value: c, text: c || 'Uncategorised' }));
    });
    select.value = Array.prototype.some.call(select.options, function (o) { return o.value === current; }) ? current : '__all';
    $('category-filter').parentNode.hidden = categories().length < 2;

    var list = $('category-options');
    list.innerHTML = '';
    categories().filter(Boolean).forEach(function (c) { list.appendChild(el('option', { value: c })); });
  }

  function renderList() {
    renderCategoryFilter();
    var query = $('search').value.trim().toLowerCase();
    var cat = $('category-filter').value;
    var container = $('template-list');
    container.innerHTML = '';

    var matches = state.templates.filter(function (t) {
      if (cat !== '__all' && (t.category || '') !== cat) return false;
      if (!query) return true;
      var hay = [t.name, t.subject, t.category, plainText(t.body)].join(' ').toLowerCase();
      return query.split(/\s+/).every(function (word) { return hay.indexOf(word) !== -1; });
    });

    var empty = $('empty-state');
    if (!state.templates.length) {
      empty.hidden = false;
      empty.textContent = 'No templates yet. Click "+ New" to create one, or import a backup from Settings (⚙).';
    } else if (!matches.length) {
      empty.hidden = false;
      empty.textContent = 'No templates match your search.';
    } else {
      empty.hidden = true;
    }

    var groups = {};
    matches.forEach(function (t) {
      var c = (t.category || '').trim();
      (groups[c] = groups[c] || []).push(t);
    });
    var groupNames = Object.keys(groups).sort(function (a, b) {
      if (!a) return 1;
      if (!b) return -1;
      return naturalCompare(a, b);
    });
    var searching = query !== '';

    groupNames.forEach(function (g) {
      var items = groups[g].sort(function (a, b) { return naturalCompare(a.name, b.name); });
      var collapsed = !searching && !!state.collapsed[g];
      var group = el('div', { className: 'group' + (collapsed ? ' collapsed' : '') });
      if (groupNames.length > 1 || g) {
        var header = el('button', { className: 'group-header', type: 'button', 'aria-expanded': String(!collapsed) }, [
          el('span', { className: 'chev', text: '▾' }),
          el('span', { text: g || 'Uncategorised' }),
          el('span', { className: 'group-count', text: '(' + items.length + ')' })
        ]);
        header.addEventListener('click', function () {
          state.collapsed[g] = !state.collapsed[g];
          writeLocal(COLLAPSED_KEY, state.collapsed);
          renderList();
        });
        group.appendChild(header);
      }
      var wrap = el('div', { className: 'group-items', role: 'list' });
      items.forEach(function (t) {
        var main = el('button', { className: 'item-main', type: 'button', title: 'Insert "' + t.name + '"' }, [
          el('span', { className: 'item-name', text: t.name }),
          t.subject ? el('span', { className: 'item-subject', text: t.subject }) : null
        ]);
        main.addEventListener('click', function () { startInsert(t); });
        var edit = el('button', { className: 'item-edit', type: 'button', title: 'Edit', 'aria-label': 'Edit ' + t.name, text: '✎' });
        edit.addEventListener('click', function () { openEditor(t); });
        wrap.appendChild(el('div', { className: 'item', role: 'listitem' }, [main, edit]));
      });
      group.appendChild(wrap);
      container.appendChild(group);
    });
  }

  // ---------------------------------------------------------------------------
  // Inserting
  // ---------------------------------------------------------------------------
  function startInsert(template) {
    if (!currentItem()) {
      toast('Open an email you are writing first, then click a template.', { error: true });
      return;
    }
    var names = findPlaceholders(template.subject, template.body);
    if (!names.length) {
      insertTemplate(template, {}, $('replace-subject').checked).then(function (ok) {
        if (ok) resetSubjectToggle();
      });
      return;
    }
    autoValues(names).then(function (auto) {
      state.filling = { template: template, names: names, confirmBlanks: false };
      $('fill-title').textContent = template.name;
      var hasSubject = !!(template.subject || '').trim();
      $('fill-subject-row').hidden = !hasSubject;
      $('fill-subject-text').textContent = template.subject || '';
      $('fill-replace-subject').checked = $('replace-subject').checked;
      var fields = $('fill-fields');
      fields.innerHTML = '';
      names.forEach(function (name, i) {
        var id = 'fill-' + i;
        var input = el('input', { type: 'text', id: id, 'data-name': name });
        if (auto[name]) input.value = auto[name].value;
        input.addEventListener('input', function () {
          state.filling.confirmBlanks = false;
          $('fill-warning').hidden = true;
          $('btn-fill-insert').textContent = 'Insert';
          input.parentNode.classList.remove('blank');
        });
        fields.appendChild(el('div', { className: 'fill-field' }, [
          el('label', { for: id }, [name, auto[name] ? el('span', { className: 'auto', text: ' – ' + auto[name].source }) : null]),
          input
        ]));
      });
      $('fill-warning').hidden = true;
      $('btn-fill-insert').textContent = 'Insert';
      show('fill');
      var inputs = Array.prototype.slice.call(fields.querySelectorAll('input'));
      var target = inputs.filter(function (i) { return !i.value; })[0] || inputs[0];
      if (target) target.focus();
    });
  }

  function onFillSubmit(e) {
    e.preventDefault();
    if (!state.filling) return;
    var values = {};
    var blanks = [];
    $('fill-fields').querySelectorAll('input').forEach(function (input) {
      var name = input.getAttribute('data-name');
      values[name] = input.value;
      if (!input.value.trim()) { blanks.push(name); input.parentNode.classList.add('blank'); }
    });
    if (blanks.length && !state.filling.confirmBlanks) {
      state.filling.confirmBlanks = true;
      var w = $('fill-warning');
      w.textContent = (blanks.length === 1 ? blanks[0] + ' is' : blanks.join(', ') + ' are') +
        ' blank. Fill ' + (blanks.length === 1 ? 'it' : 'them') + ' in, or click "Insert anyway".';
      w.hidden = false;
      $('btn-fill-insert').textContent = 'Insert anyway';
      return;
    }
    var template = state.filling.template;
    var replace = $('fill-subject-row').hidden ? $('replace-subject').checked : $('fill-replace-subject').checked;
    insertTemplate(template, values, replace).then(function (ok) {
      if (ok) { state.filling = null; resetSubjectToggle(); show('list'); }
    });
  }

  /** Puts the "Use the template's subject" tick box back to the default from Settings. */
  function resetSubjectToggle() {
    $('replace-subject').checked = state.settings.replaceSubject !== false;
  }

  function insertTemplate(template, values, replaceSubject) {
    var item = currentItem();
    if (!item) {
      toast('Open an email you are writing first, then click a template.', { error: true });
      return Promise.resolve(false);
    }
    var bodyHtml = fillPlaceholders(template.body, values, true);
    var subject = fillPlaceholders(template.subject || '', values, false).trim();

    return officeCall(function (cb) { item.body.getTypeAsync(cb); })
      .then(function (bodyType) {
        var isHtml = bodyType === Office.CoercionType.Html;
        var data = isHtml ? prepareForOutlook(bodyHtml, state.settings) : htmlToText(bodyHtml);
        return officeCall(function (cb) {
          item.body.setSelectedDataAsync(data, { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text }, cb);
        });
      })
      .then(function () {
        if (!subject || !item.subject || !item.subject.getAsync) {
          toast('Inserted "' + template.name + '".');
          return true;
        }
        if (!replaceSubject) {
          toast('Inserted "' + template.name + '". Subject left as it was.', {
            actionLabel: 'Use template subject',
            onAction: function () {
              officeCall(function (cb) { item.subject.setAsync(subject, cb); })
                .then(function () { toast('Subject updated.'); })
                .catch(function (err) { toast('Couldn\'t change the subject: ' + err.message, { error: true }); });
            }
          });
          return true;
        }
        return officeCall(function (cb) { item.subject.getAsync(cb); }).then(function (previous) {
          previous = previous || '';
          if (previous.trim() === subject) {
            toast('Inserted "' + template.name + '".');
            return true;
          }
          return officeCall(function (cb) { item.subject.setAsync(subject, cb); }).then(function () {
            if (!previous.trim()) {
              toast('Inserted "' + template.name + '" and set the subject.');
            } else {
              toast('Inserted "' + template.name + '" and replaced the subject.', {
                actionLabel: 'Undo subject',
                onAction: function () {
                  officeCall(function (cb) { item.subject.setAsync(previous, cb); })
                    .then(function () { toast('Subject put back.'); })
                    .catch(function (err) { toast('Couldn\'t change the subject: ' + err.message, { error: true }); });
                }
              });
            }
            return true;
          });
        });
      })
      .catch(function (err) {
        toast('Couldn\'t insert the template: ' + (err && err.message ? err.message : 'unknown error'), { error: true });
        return false;
      });
  }

  // ---------------------------------------------------------------------------
  // Editor
  // ---------------------------------------------------------------------------
  function editorSnapshot() {
    return JSON.stringify([$('edit-name').value, $('edit-category').value, $('edit-subject').value, $('edit-body').innerHTML]);
  }

  function openEditor(template) {
    var t = template || { id: null, name: '', category: $('category-filter').value !== '__all' ? $('category-filter').value : '', subject: '', body: '' };
    $('edit-title').textContent = template ? 'Edit template' : 'New template';
    $('edit-name').value = t.name;
    $('edit-category').value = t.category || '';
    $('edit-subject').value = t.subject || '';
    $('edit-body').innerHTML = t.body || '';
    $('btn-delete').hidden = !template;
    $('btn-delete').textContent = 'Delete';
    $('btn-edit-cancel').textContent = 'Cancel';
    $('edit-error').hidden = true;
    $('custom-placeholder').hidden = true;
    $('link-form').hidden = true;
    var canCapture = Office.context.requirements && Office.context.requirements.isSetSupported &&
      Office.context.requirements.isSetSupported('Mailbox', '1.2') && !!currentItem();
    $('btn-capture').hidden = !canCapture;
    state.editing = { id: t.id, snapshot: null, confirmDiscard: false, confirmDelete: false };
    renderCategoryFilter();
    show('edit');
    state.editing.snapshot = editorSnapshot();
    renderChips();
    if (!template) $('edit-name').focus();
  }

  function renderChips() {
    var box = $('placeholder-chips');
    box.innerHTML = '';
    var names = COMMON_PLACEHOLDERS.slice();
    findPlaceholders($('edit-subject').value, $('edit-body').innerHTML).forEach(function (n) {
      if (names.indexOf(n) === -1) names.push(n);
    });
    names.forEach(function (name) {
      var chip = el('button', { type: 'button', className: 'chip', text: '{' + name + '}' });
      chip.addEventListener('mousedown', function (e) { e.preventDefault(); });
      chip.addEventListener('click', function () { insertIntoEditor('{' + name + '}'); });
      box.appendChild(chip);
    });
    var add = el('button', { type: 'button', className: 'chip add', text: '+ Custom…' });
    add.addEventListener('mousedown', function (e) { e.preventDefault(); saveEditorRange(); });
    add.addEventListener('click', function () {
      $('link-form').hidden = true;
      $('custom-placeholder').hidden = false;
      $('custom-placeholder-name').value = '';
      $('custom-placeholder-name').focus();
    });
    box.appendChild(add);
  }

  function selectionInEditor() {
    var sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    var range = sel.getRangeAt(0);
    return $('edit-body').contains(range.commonAncestorContainer) ? range : null;
  }

  function saveEditorRange() {
    var r = selectionInEditor();
    state.savedRange = r ? r.cloneRange() : null;
  }

  function restoreEditorRange() {
    var editor = $('edit-body');
    editor.focus();
    var sel = window.getSelection();
    sel.removeAllRanges();
    if (state.savedRange) {
      sel.addRange(state.savedRange);
    } else {
      var r = document.createRange();
      r.selectNodeContents(editor);
      r.collapse(false);
      sel.addRange(r);
    }
  }

  function insertIntoEditor(text) {
    // Subject field focused? Put the placeholder there instead.
    var subjectInput = $('edit-subject');
    if (document.activeElement === subjectInput) {
      var s = subjectInput.selectionStart, en = subjectInput.selectionEnd;
      subjectInput.value = subjectInput.value.slice(0, s) + text + subjectInput.value.slice(en);
      subjectInput.selectionStart = subjectInput.selectionEnd = s + text.length;
      renderChips();
      return;
    }
    focusEditorKeepingSelection();
    document.execCommand('insertText', false, text);
    renderChips();
  }

  /** Focuses the editor, keeping the caret where it was (or putting it at the end). */
  function focusEditorKeepingSelection() {
    saveEditorRange();
    restoreEditorRange();
  }

  function insertHtmlIntoEditor(html) {
    focusEditorKeepingSelection();
    document.execCommand('insertHTML', false, html);
  }

  function onToolbarClick(e) {
    var btn = e.target.closest('button[data-cmd]');
    if (!btn) return;
    var cmd = btn.getAttribute('data-cmd');
    if (cmd === 'createLink') {
      saveEditorRange();
      $('custom-placeholder').hidden = true;
      $('link-form').hidden = false;
      $('link-url').value = 'https://';
      $('link-url').focus();
      return;
    }
    focusEditorKeepingSelection();
    document.execCommand(cmd, false, null);
  }

  function onEditorPaste(e) {
    var cd = e.clipboardData || window.clipboardData;
    if (!cd) return;
    e.preventDefault();
    var html = cd.getData('text/html');
    if (html) {
      var cleaned = cleanHtml(html);
      document.execCommand('insertHTML', false, cleaned.html);
      if (cleaned.imagesRemoved) toast('Pictures were left out. Templates can only store text and formatting.');
    } else {
      var text = cd.getData('text/plain') || '';
      var lines = text.replace(/\r\n?/g, '\n').split('\n');
      var out = lines.map(function (line, i) { return (i ? '<br>' : '') + escapeHtml(line); }).join('');
      document.execCommand('insertHTML', false, out);
    }
    renderChips();
  }

  function onCapture() {
    var item = currentItem();
    if (!item || typeof item.getSelectedDataAsync !== 'function') {
      toast('This version of Outlook can\'t read the selection.', { error: true });
      return;
    }
    officeCall(function (cb) { item.getSelectedDataAsync(Office.CoercionType.Html, cb); })
      .catch(function () { return officeCall(function (cb) { item.getSelectedDataAsync(Office.CoercionType.Text, cb); }); })
      .then(function (result) {
        var data = result && result.data ? result.data : '';
        if (!data || !data.replace(/<[^>]*>|&nbsp;|\s/g, '')) {
          toast('Select some text in the email first, then click this button again.', { error: true });
          return;
        }
        if (result.sourceProperty === 'subject') {
          $('edit-subject').value = plainText(data);
          toast('Copied into the subject.');
          return;
        }
        var html = /<[a-z][\s\S]*>/i.test(data) ? data : escapeHtml(data).replace(/\r?\n/g, '<br>');
        var cleaned = cleanHtml(html);
        var editor = $('edit-body');
        if (!plainText(editor.innerHTML)) editor.innerHTML = cleaned.html;
        else insertHtmlIntoEditor(cleaned.html);
        renderChips();
        toast(cleaned.imagesRemoved ? 'Copied the text. Pictures were left out.' : 'Copied the selected text.');
      })
      .catch(function (err) { toast('Couldn\'t read the selection: ' + err.message, { error: true }); });
  }

  function editError(message) {
    var p = $('edit-error');
    p.textContent = message;
    p.hidden = !message;
  }

  function onSave() {
    var name = $('edit-name').value.trim();
    if (!name) { editError('Give the template a name.'); $('edit-name').focus(); return; }
    var cleaned = cleanHtml($('edit-body').innerHTML);
    if (!plainText(cleaned.html) && !/<(table|img)/i.test(cleaned.html)) {
      editError('The template has no text yet.');
      return;
    }
    var record = {
      id: state.editing.id || newId(),
      name: name,
      category: $('edit-category').value.trim(),
      subject: $('edit-subject').value.trim(),
      body: cleaned.html
    };
    var next = state.templates.filter(function (t) { return t.id !== record.id; });
    var existingIndex = state.templates.findIndex(function (t) { return t.id === record.id; });
    if (existingIndex === -1) next.push(record); else next.splice(existingIndex, 0, record);

    $('btn-save').disabled = true;
    persist(next).then(function () {
      state.editing = null;
      show('list');
      renderList();
      toast('Saved "' + record.name + '".');
    }).catch(function (err) {
      editError(err.message);
    }).then(function () { $('btn-save').disabled = false; });
  }

  function onEditCancel() {
    if (!state.editing) { show('list'); return; }
    if (editorSnapshot() !== state.editing.snapshot && !state.editing.confirmDiscard) {
      state.editing.confirmDiscard = true;
      editError('You have unsaved changes. Click Cancel again to discard them.');
      $('btn-edit-cancel').textContent = 'Discard changes';
      return;
    }
    state.editing = null;
    show('list');
  }

  function onDelete() {
    if (!state.editing || !state.editing.id) return;
    if (!state.editing.confirmDelete) {
      state.editing.confirmDelete = true;
      $('btn-delete').textContent = 'Click again to delete';
      setTimeout(function () {
        if (state.editing) { state.editing.confirmDelete = false; $('btn-delete').textContent = 'Delete'; }
      }, 4000);
      return;
    }
    var id = state.editing.id;
    var removed = state.templates.filter(function (t) { return t.id === id; })[0];
    var next = state.templates.filter(function (t) { return t.id !== id; });
    persist(next).then(function () {
      state.editing = null;
      show('list');
      renderList();
      toast('Deleted "' + (removed ? removed.name : 'template') + '".', {
        actionLabel: 'Undo',
        onAction: function () {
          if (!removed) return;
          persist(state.templates.concat([removed])).then(function () { renderList(); toast('Restored.'); })
            .catch(function (err) { toast(err.message, { error: true }); });
        }
      });
    }).catch(function (err) { editError(err.message); });
  }

  // ---------------------------------------------------------------------------
  // Settings, backup and import
  // ---------------------------------------------------------------------------
  function openSettings() {
    $('set-font').value = state.settings.fontFamily || '';
    $('set-size').value = String(state.settings.fontSize || '12');
    $('set-size').disabled = !state.settings.fontFamily;
    $('set-replace-subject').checked = state.settings.replaceSubject !== false;
    $('export-area').hidden = true;
    $('import-message').hidden = true;
    $('import-text').value = '';
    $('import-file').value = '';
    state.confirmReplaceAll = false;
    $('btn-import-replace').textContent = 'Replace all';
    renderStorage();
    show('settings');
  }

  function renderStorage() {
    var used = storageUsed(state.templates, state.settings);
    var pct = Math.min(100, Math.round(used / STORAGE_LIMIT * 100));
    var bar = $('storage-bar');
    bar.style.width = pct + '%';
    bar.className = 'meter-fill' + (pct >= 95 ? ' full' : pct >= 75 ? ' high' : '');
    $('storage-text').textContent = state.templates.length + ' template' + (state.templates.length === 1 ? '' : 's') +
      ' using about ' + pct + '% of the available space.';
  }

  function onSettingsSave() {
    var settings = Object.assign({}, state.settings, {
      fontFamily: $('set-font').value,
      fontSize: $('set-size').value,
      replaceSubject: $('set-replace-subject').checked
    });
    persist(state.templates, settings).then(function () {
      renderStorage();
      resetSubjectToggle();
      toast('Settings saved.');
    }).catch(function (err) { toast(err.message, { error: true }); });
  }

  function exportText() {
    return JSON.stringify({
      format: 'quick-templates',
      version: 1,
      exportedAt: new Date().toISOString(),
      templates: state.templates.map(function (t) {
        return { name: t.name, category: t.category || '', subject: t.subject || '', body: t.body || '' };
      })
    }, null, 1);
  }

  function onExport() {
    $('export-text').value = exportText();
    $('export-area').hidden = false;
  }

  function onCopyExport() {
    var ta = $('export-text');
    var done = function () { toast('Backup copied to the clipboard.'); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(ta.value).then(done, fallback);
    } else fallback();
    function fallback() {
      ta.focus();
      ta.select();
      try { document.execCommand('copy') ? done() : toast('Press Ctrl+C to copy the selected text.'); }
      catch (e) { toast('Press Ctrl+C to copy the selected text.'); }
    }
  }

  function parseImport(text) {
    var data = JSON.parse(text);
    var list = Array.isArray(data) ? data : (data && Array.isArray(data.templates) ? data.templates : null);
    if (!list) throw new Error('That doesn\'t look like a templates backup.');
    var imagesRemoved = 0;
    var templates = list.filter(function (t) { return t && typeof t.name === 'string' && t.name.trim(); }).map(function (t) {
      var body = typeof t.body === 'string' ? t.body : (t.body && typeof t.body.content === 'string' ? t.body.content : '');
      var cleaned = cleanHtml(body);
      imagesRemoved += cleaned.imagesRemoved;
      return {
        id: newId(),
        name: t.name.trim().slice(0, 80),
        category: typeof t.category === 'string' ? t.category.trim().slice(0, 40) : '',
        subject: typeof t.subject === 'string' ? t.subject.trim().slice(0, 200) : '',
        body: cleaned.html
      };
    });
    if (!templates.length) throw new Error('No templates were found in the backup.');
    return { templates: templates, imagesRemoved: imagesRemoved };
  }

  function importMessage(text, isError) {
    var p = $('import-message');
    p.textContent = text;
    p.className = 'small' + (isError ? ' error' : '');
    p.hidden = false;
  }

  function readImportSource() {
    var file = $('import-file').files && $('import-file').files[0];
    if (file) {
      return new Promise(function (resolve, reject) {
        var reader = new FileReader();
        reader.onload = function () { resolve(String(reader.result || '')); };
        reader.onerror = function () { reject(new Error('Couldn\'t read that file.')); };
        reader.readAsText(file);
      });
    }
    var text = $('import-text').value.trim();
    if (!text) return Promise.reject(new Error('Choose a backup file or paste the backup text first.'));
    return Promise.resolve(text);
  }

  function onImport(replaceAll) {
    if (replaceAll && !state.confirmReplaceAll) {
      state.confirmReplaceAll = true;
      $('btn-import-replace').textContent = 'Click again to replace all';
      importMessage('This will delete your ' + state.templates.length + ' current template(s) and use the backup instead.', true);
      return;
    }
    readImportSource().then(function (text) {
      var parsed = parseImport(text);
      var next, added = 0, updated = 0;
      if (replaceAll) {
        next = parsed.templates;
        added = next.length;
      } else {
        next = state.templates.slice();
        parsed.templates.forEach(function (t) {
          var i = next.findIndex(function (x) {
            return x.name.toLowerCase() === t.name.toLowerCase() && (x.category || '').toLowerCase() === t.category.toLowerCase();
          });
          if (i === -1) { next.push(t); added++; } else { t.id = next[i].id; next[i] = t; updated++; }
        });
      }
      return persist(next).then(function () {
        state.confirmReplaceAll = false;
        $('btn-import-replace').textContent = 'Replace all';
        renderStorage();
        renderList();
        var msg = replaceAll ? 'Replaced with ' + added + ' template(s).' :
          added + ' added, ' + updated + ' updated.';
        if (parsed.imagesRemoved) msg += ' ' + parsed.imagesRemoved + ' picture(s) were left out.';
        importMessage(msg, false);
        toast('Import complete.');
      });
    }).catch(function (err) {
      importMessage(err instanceof SyntaxError ? 'That isn\'t valid backup text.' : err.message, true);
    });
  }

  function checkLocalBackup() {
    var banner = $('restore-banner');
    banner.hidden = true;
    if (state.templates.length) return;
    var backup = readLocal(LOCAL_BACKUP_KEY, null);
    if (!backup || !backup.payload) return;
    var templates;
    try { templates = decodeTemplates(backup.payload); } catch (e) { return; }
    if (!templates.length) return;
    banner.innerHTML = '';
    banner.appendChild(document.createTextNode('Found ' + templates.length + ' template(s) saved on this computer' +
      (backup.savedAt ? ' (' + new Date(backup.savedAt).toLocaleDateString('en-GB') + ')' : '') + '.'));
    var btn = el('button', { type: 'button', className: 'btn', text: 'Restore them' });
    btn.addEventListener('click', function () {
      persist(templates).then(function () {
        banner.hidden = true;
        renderList();
        toast('Templates restored.');
      }).catch(function (err) { toast(err.message, { error: true }); });
    });
    banner.appendChild(el('div', null, [btn]));
    banner.hidden = false;
  }

  // ---------------------------------------------------------------------------
  // Start-up
  // ---------------------------------------------------------------------------
  function fatal(message) {
    $('loading').hidden = true;
    var f = $('fatal');
    f.textContent = message;
    f.hidden = false;
  }

  function wireEvents() {
    $('search').addEventListener('input', renderList);
    $('category-filter').addEventListener('change', renderList);
    $('btn-new').addEventListener('click', function () { openEditor(null); });
    $('btn-settings').addEventListener('click', openSettings);
    document.querySelectorAll('[data-action="back"]').forEach(function (b) {
      b.addEventListener('click', function () { state.filling = null; show('list'); renderList(); });
    });

    $('fill-form').addEventListener('submit', onFillSubmit);

    $('btn-edit-back').addEventListener('click', onEditCancel);
    $('btn-edit-cancel').addEventListener('click', onEditCancel);
    $('btn-save').addEventListener('click', onSave);
    $('btn-delete').addEventListener('click', onDelete);
    $('btn-capture').addEventListener('click', onCapture);
    var toolbar = document.querySelector('.toolbar');
    toolbar.addEventListener('mousedown', function (e) { if (e.target.closest('button')) e.preventDefault(); });
    toolbar.addEventListener('click', onToolbarClick);
    $('edit-body').addEventListener('paste', onEditorPaste);
    $('edit-body').addEventListener('blur', function () { renderChips(); });
    $('edit-subject').addEventListener('change', renderChips);
    ['edit-name', 'edit-category', 'edit-subject'].forEach(function (id) {
      $(id).addEventListener('input', function () { editError(''); if (state.editing) { state.editing.confirmDiscard = false; $('btn-edit-cancel').textContent = 'Cancel'; } });
    });
    $('edit-body').addEventListener('input', function () { editError(''); if (state.editing) { state.editing.confirmDiscard = false; $('btn-edit-cancel').textContent = 'Cancel'; } });

    $('btn-custom-placeholder-add').addEventListener('click', function () {
      var name = $('custom-placeholder-name').value.trim().replace(/\s+/g, ' ');
      if (!new RegExp('^' + PLACEHOLDER_SOURCE.replace('\\{', '').replace('\\}', '') + '$').test(name)) {
        toast('Use letters, numbers and spaces only (up to 40 characters).', { error: true });
        return;
      }
      $('custom-placeholder').hidden = true;
      restoreEditorRange();
      document.execCommand('insertText', false, '{' + name + '}');
      renderChips();
    });
    $('custom-placeholder-name').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); $('btn-custom-placeholder-add').click(); }
      if (e.key === 'Escape') { $('custom-placeholder').hidden = true; }
    });
    $('btn-link-add').addEventListener('click', function () {
      var url = $('link-url').value.trim();
      if (!/^(https?:\/\/|mailto:)\S+/i.test(url)) { toast('Enter a full web address starting with https://', { error: true }); return; }
      $('link-form').hidden = true;
      restoreEditorRange();
      var sel = window.getSelection();
      if (sel && sel.isCollapsed) document.execCommand('insertHTML', false, '<a href="' + escapeHtml(url) + '">' + escapeHtml(url) + '</a>');
      else document.execCommand('createLink', false, url);
    });
    $('link-url').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); $('btn-link-add').click(); }
      if (e.key === 'Escape') { $('link-form').hidden = true; }
    });

    $('set-font').addEventListener('change', function () { $('set-size').disabled = !$('set-font').value; });
    $('btn-settings-save').addEventListener('click', onSettingsSave);
    $('btn-export').addEventListener('click', onExport);
    $('btn-copy-export').addEventListener('click', onCopyExport);
    $('btn-import-add').addEventListener('click', function () { onImport(false); });
    $('btn-import-replace').addEventListener('click', function () { onImport(true); });
    $('import-file').addEventListener('change', function () { $('import-message').hidden = true; });
  }

  function start() {
    try {
      document.execCommand('defaultParagraphSeparator', false, 'div');
    } catch (e) { /* ignore */ }
    wireEvents();
    var loadError = null;
    try {
      loadFromMailbox();
    } catch (e) {
      loadError = e;
      state.templates = [];
    }
    $('loading').hidden = true;
    resetSubjectToggle();
    show('list');
    renderList();
    if (loadError) {
      toast('Your saved templates couldn\'t be read. You can restore a backup from Settings.', { error: true, duration: 15000 });
    }
    checkLocalBackup();

    if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
        resetSubjectToggle();
        if (state.view === 'fill') { state.filling = null; show('list'); renderList(); }
        if (state.view === 'edit') { $('btn-capture').hidden = !currentItem(); }
      });
    }
  }

  // Test hooks (harmless in production).
  window.QuickTemplates = {
    cleanHtml: cleanHtml,
    prepareForOutlook: prepareForOutlook,
    fillPlaceholders: fillPlaceholders,
    findPlaceholders: findPlaceholders,
    firstNameFromDisplay: firstNameFromDisplay,
    htmlToText: htmlToText,
    encodeTemplates: encodeTemplates,
    state: state
  };

  if (typeof Office === 'undefined' || !Office.onReady) {
    fatal('Office.js didn\'t load. This page only works inside Outlook.');
    return;
  }
  Office.onReady(function (info) {
    if (info && info.host && info.host !== Office.HostType.Outlook) {
      fatal('Quick Templates only works in Outlook.');
      return;
    }
    if (!Office.context || !Office.context.roamingSettings) {
      fatal('Open Quick Templates from an email you are writing in Outlook.');
      return;
    }
    start();
  });
})();
