/* Stillsnote front end. Talks to the Google script set in /config.js. */
(function () {
  'use strict';

  var CONFIG = window.STILLSNOTE_CONFIG || {};
  var ENDPOINT = String(CONFIG.endpoint || '');
  var CONNECTED = /^https:\/\/script\.google\.com\/(a\/macros\/[^\/]+|macros)\/s\/[^\/]+\/exec$/.test(ENDPOINT.trim());
  var DEMO = /[?&]demo=1\b/.test(location.search);
  var MAX_TAGS = 5;
  var MAX_SIDE = 1568;

  var state = {
    token: null,
    member: null,
    profile: null,
    image: null,          // { data, mediaType, thumb, url }
    post: null,           // last AI result
    capKey: 'story',
    captions: { story: '', short: '', question: '' },
    tags: [],             // [{ tag, tier }]
    readTags: [],
    palette: [],
    drafts: null,
    trends: null,
    region: '',
    niche: 'all',
    lastRoute: ''
  };

  /* ------------------------------------------------------------ */
  /* Small helpers                                                 */
  /* ------------------------------------------------------------ */

  var store = {
    get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
    del: function (k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }
  };
  state.token = store.get('sn_token');

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v == null || v === false) return;
        if (k === 'text') node.textContent = v;
        else if (k === 'class') node.className = v;
        else if (k.indexOf('on') === 0) node.addEventListener(k.slice(2), v);
        else node.setAttribute(k, v === true ? '' : v);
      });
    }
    [].concat(children || []).forEach(function (c) {
      if (c == null || c === false) return;
      node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return node;
  }

  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); return node; }

  function msg(node, text, kind) {
    if (!node) return;
    node.textContent = text || '';
    node.className = 'form-msg' + (kind ? ' ' + kind : '');
  }

  var toastTimer;
  function toast(text) {
    var t = $('#toast');
    t.textContent = text;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, 2600);
  }

  function busy(btn, on, label) {
    if (!btn) return;
    if (on) {
      btn.dataset.label = btn.dataset.label || btn.textContent.trim();
      btn.disabled = true;
      btn.classList.add('loading');
      if (label) setBtnText(btn, label);
    } else {
      btn.disabled = false;
      btn.classList.remove('loading');
      if (btn.dataset.label && label !== false) setBtnText(btn, btn.dataset.label);
    }
  }
  function setBtnText(btn, text) {
    var nodes = Array.prototype.slice.call(btn.childNodes);
    var textNode = nodes.filter(function (n) { return n.nodeType === 3 && n.textContent.trim(); })[0];
    if (textNode) textNode.textContent = text; else btn.appendChild(document.createTextNode(text));
  }

  function today(offsetDays) {
    var d = new Date(Date.now() + (offsetDays || 0) * 86400000);
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return d.getFullYear() + '-' + m + '-' + day;
  }

  function prettyDate(iso) {
    if (!iso) return 'No date';
    var p = iso.split('-');
    var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    return d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' });
  }

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise(function (resolve, reject) {
      var ta = el('textarea', { style: 'position:fixed;left:-9999px' });
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy') ? resolve() : reject(); } catch (e) { reject(e); }
      document.body.removeChild(ta);
    });
  }

  /* ------------------------------------------------------------ */
  /* API                                                           */
  /* ------------------------------------------------------------ */

  function api(action, data) {
    var body = Object.assign({ action: action, token: state.token }, data || {});
    if (DEMO) return mockApi(action, body);
    if (!CONNECTED) return Promise.reject(new Error('Stillsnote is not connected yet. Add the Web app URL to config.js.'));
    return fetch(ENDPOINT.trim(), { method: 'POST', body: JSON.stringify(body) })
      .then(function (res) { return res.json(); })
      .catch(function () { throw new Error('Could not reach Stillsnote. Check your connection and try again.'); })
      .then(function (j) {
        if (!j.ok) {
          if (j.code === 'signin') signOut(true);
          var err = new Error(j.error || 'Something went wrong.');
          err.code = j.code;
          throw err;
        }
        if (j.member) setMember(j.member);
        return j;
      });
  }

  /* ------------------------------------------------------------ */
  /* Routing                                                       */
  /* ------------------------------------------------------------ */

  var ROUTES = ['studio', 'planner', 'trends', 'creative', 'style'];

  function route() {
    var name = (location.hash.replace(/^#\/?/, '') || 'studio').split('?')[0];
    if (ROUTES.indexOf(name) < 0) name = 'studio';
    if (!state.token) { showGate(); return; }
    $('#view-gate').hidden = true;
    $('#app').hidden = false;
    ROUTES.forEach(function (r) { $('#view-' + r).hidden = r !== name; });
    $all('.nav a').forEach(function (a) {
      if (a.dataset.route === name) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    if (state.lastRoute && state.lastRoute !== name) window.scrollTo(0, 0);
    state.lastRoute = name;
    document.title = (name === 'studio' ? 'Studio' : name.charAt(0).toUpperCase() + name.slice(1)) + ' · Stillsnote';
    if (name === 'planner') loadPlanner();
    if (name === 'trends') loadTrends();
    if (name === 'style') loadStyle();
  }

  function showGate() {
    $('#app').hidden = true;
    $('#view-gate').hidden = false;
    document.title = 'Stillsnote';
  }

  function setMember(m) {
    state.member = m;
    var u = $('#usage');
    if (!m) { u.textContent = ''; return; }
    u.textContent = m.limit > 0 ? (m.uses + ' / ' + m.limit + ' THIS MONTH') : (m.plan === 'Paid' ? 'PAID' : '');
  }

  function signOut(expired) {
    state.token = null;
    store.del('sn_token');
    state.member = null; state.profile = null; state.drafts = null; state.trends = null;
    setMember(null);
    showGate();
    if (expired) msg($('#form-signin .form-msg'), 'Please sign in again.', 'error');
  }

  /* ------------------------------------------------------------ */
  /* Gate                                                          */
  /* ------------------------------------------------------------ */

  function initGate() {
    var tabs = [$('#tab-signin'), $('#tab-request')];
    tabs.forEach(function (t) {
      t.addEventListener('click', function () {
        tabs.forEach(function (x) {
          var on = x === t;
          x.setAttribute('aria-selected', on ? 'true' : 'false');
          $('#' + x.getAttribute('aria-controls')).hidden = !on;
        });
      });
    });

    $('#form-signin').addEventListener('submit', function (e) {
      e.preventDefault();
      var f = e.target, out = $('.form-msg', f), btn = $('button[type=submit]', f);
      var email = f.email.value.trim(), code = f.code.value.trim();
      if (!email || !code) { msg(out, 'Enter your email and access code.', 'error'); return; }
      busy(btn, true); msg(out, '');
      api('login', { email: email, code: code }).then(function (j) {
        state.token = j.token;
        store.set('sn_token', j.token);
        f.reset();
        if (!location.hash) location.hash = '#/studio';
        route();
        loadProfile();
      }).catch(function (err) { msg(out, err.message, 'error'); })
        .then(function () { busy(btn, false); });
    });

    $('#form-request').addEventListener('submit', function (e) {
      e.preventDefault();
      var f = e.target, out = $('.form-msg', f), btn = $('button[type=submit]', f);
      if (!f.name.value.trim() || !f.email.value.trim()) { msg(out, 'Enter your name and email.', 'error'); return; }
      busy(btn, true); msg(out, '');
      api('signup', {
        name: f.name.value, email: f.email.value, instagram: f.instagram.value,
        optin: f.optin.checked, website: f.website.value
      }).then(function (j) {
        var text = {
          pending: 'Request sent. You will get your access code by email once it is approved.',
          approved: 'You are in. Check your email for your access code, then sign in.',
          resent: 'You already have an account. We just emailed your code again.'
        }[j.status] || 'Request sent.';
        msg(out, text, 'ok');
        f.reset();
      }).catch(function (err) { msg(out, err.message, 'error'); })
        .then(function () { busy(btn, false); });
    });
  }

  /* ------------------------------------------------------------ */
  /* Profile                                                       */
  /* ------------------------------------------------------------ */

  function loadProfile() {
    return api('getProfile').then(function (j) {
      state.profile = j.profile || {};
      fillPillars();
      return state.profile;
    }).catch(function () { state.profile = state.profile || {}; return state.profile; });
  }

  function fillPillars() {
    var sel = $('#save-pillar');
    var current = sel.value;
    clear(sel).appendChild(el('option', { value: '', text: 'None' }));
    ((state.profile && state.profile.pillars) || []).forEach(function (p) {
      sel.appendChild(el('option', { value: p, text: p }));
    });
    sel.value = current;
  }

  function loadStyle() {
    var f = $('#style-form');
    var p = state.profile;
    var fill = function (p) {
      ['niche', 'audience', 'tone', 'location', 'samples', 'avoid'].forEach(function (k) { f[k].value = p[k] || ''; });
      f.language.value = p.language || 'English';
      f.pillars.value = (p.pillars || []).join(', ');
    };
    if (p) fill(p); else loadProfile().then(fill);
  }

  function initStyle() {
    $('#style-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var f = e.target, out = $('.form-msg', f), btn = $('button[type=submit]', f);
      var profile = {
        niche: f.niche.value, audience: f.audience.value, tone: f.tone.value, language: f.language.value,
        location: f.location.value, samples: f.samples.value, avoid: f.avoid.value,
        pillars: f.pillars.value.split(',').map(function (s) { return s.trim(); }).filter(Boolean).slice(0, 6)
      };
      busy(btn, true); msg(out, '');
      api('saveProfile', { profile: profile }).then(function (j) {
        state.profile = j.profile;
        fillPillars();
        msg(out, 'Saved. New captions will use this.', 'ok');
      }).catch(function (err) { msg(out, err.message, 'error'); })
        .then(function () { busy(btn, false); });
    });
  }

  /* ------------------------------------------------------------ */
  /* Studio                                                        */
  /* ------------------------------------------------------------ */

  function initStudio() {
    var zone = $('#dropzone'), file = $('#file');
    file.addEventListener('change', function () { if (file.files[0]) takeFile(file.files[0]); file.value = ''; });
    ['dragenter', 'dragover'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.add('drag'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.remove('drag'); });
    });
    zone.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) takeFile(f);
    });
    document.addEventListener('paste', function (e) {
      if ($('#view-studio').hidden) return;
      var item = Array.prototype.slice.call((e.clipboardData && e.clipboardData.items) || [])
        .filter(function (i) { return i.type.indexOf('image/') === 0; })[0];
      if (item) takeFile(item.getAsFile());
    });
    $('#swap').addEventListener('click', function () { file.click(); });

    $('#generate').addEventListener('click', generate);
    $('#retry').addEventListener('click', generate);

    $all('#cap-tabs button').forEach(function (b) {
      b.addEventListener('click', function () { switchCaption(b.dataset.cap); });
    });
    $('#caption').addEventListener('input', function () {
      state.captions[state.capKey] = $('#caption').value;
      updateCount();
    });

    $('#tag-add').addEventListener('submit', function (e) {
      e.preventDefault();
      var input = $('#tag-input');
      var raw = input.value.trim().replace(/^#+/, '').replace(/\s+/g, '');
      if (!raw) return;
      if (state.tags.length >= MAX_TAGS) { toast('Instagram allows 5 hashtags. Remove one first.'); return; }
      var tag = '#' + raw.toLowerCase();
      if (state.tags.some(function (t) { return t.tag === tag; })) { input.value = ''; return; }
      state.tags.push({ tag: tag, tier: 'yours' });
      input.value = '';
      renderTags();
    });

    $('#copy').addEventListener('click', function () {
      var text = composePost();
      if (!text.trim()) { toast('Nothing to copy yet.'); return; }
      copyText(text).then(function () { toast('Copied. Paste it into Instagram.'); })
        .catch(function () { toast('Could not copy. Select the caption and copy it by hand.'); });
    });

    $('#save-date').value = today(1);
    $('#save-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var out = $('#save-msg'), btn = $('#save-form button[type=submit]');
      var caption = (state.captions[state.capKey] || '').trim();
      if (!caption && !state.image) { msg(out, 'Write a caption or add a photo first.', 'error'); return; }
      busy(btn, true); msg(out, '');
      api('saveDraft', { draft: {
        date: $('#save-date').value, status: 'Draft', pillar: $('#save-pillar').value,
        caption: caption, hashtags: state.tags.map(function (t) { return t.tag; }),
        story: $('#story').value, thumb: state.image ? state.image.thumb : '',
        readTags: state.readTags, palette: state.palette
      } }).then(function () {
        state.drafts = null;
        msg(out, 'Saved to your planner.', 'ok');
      }).catch(function (err) { msg(out, err.message, 'error'); })
        .then(function () { busy(btn, false); });
    });

    renderTags();
    updateCount();
  }

  function takeFile(file) {
    var out = $('#studio-msg');
    msg(out, '');
    if (!file || !/^image\//.test(file.type || 'image/')) { msg(out, 'That file is not a photo.', 'error'); return; }
    var url = URL.createObjectURL(file);
    var img = new Image();
    img.onload = function () {
      try {
        state.image = prepareImage(img);
        state.image.url = url;
        var pv = $('#preview');
        pv.src = url; pv.alt = 'Your photo'; pv.hidden = false;
        $('#drop-empty').hidden = true;
      } catch (e) {
        msg(out, 'Could not prepare that photo. Try another one.', 'error');
      }
    };
    img.onerror = function () {
      URL.revokeObjectURL(url);
      msg(out, 'Could not open that photo. If it is an iPhone HEIC file, export it as JPG first, or set Camera > Formats to Most Compatible.', 'error');
    };
    img.src = url;
  }

  function prepareImage(img) {
    var w = img.naturalWidth, h = img.naturalHeight;
    var scale = Math.min(1, MAX_SIDE / Math.max(w, h));
    var c = document.createElement('canvas');
    c.width = Math.round(w * scale); c.height = Math.round(h * scale);
    var ctx = c.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    ctx.drawImage(img, 0, 0, c.width, c.height);
    var data = c.toDataURL('image/jpeg', 0.85).split(',')[1];

    // 4:5 center-cropped thumbnail for the planner grid
    var tw = 320, th = 400;
    var t = document.createElement('canvas');
    t.width = tw; t.height = th;
    var tctx = t.getContext('2d');
    var srcRatio = w / h, dstRatio = tw / th;
    var sw = w, sh = h, sx = 0, sy = 0;
    if (srcRatio > dstRatio) { sw = h * dstRatio; sx = (w - sw) / 2; } else { sh = w / dstRatio; sy = (h - sh) / 2; }
    tctx.drawImage(img, sx, sy, sw, sh, 0, 0, tw, th);
    var thumb = t.toDataURL('image/jpeg', 0.7);
    if (thumb.length > 44000) thumb = t.toDataURL('image/jpeg', 0.5);
    if (thumb.length > 44000) thumb = '';
    return { data: data, mediaType: 'image/jpeg', thumb: thumb };
  }

  function generate() {
    var out = $('#studio-msg'), btn = $('#generate');
    if (!state.image) { msg(out, 'Add a photo first.', 'error'); return; }
    var vibe = ($('input[name=vibe]:checked') || {}).value || 'Chill';
    busy(btn, true, 'Writing'); busy($('#retry'), true);
    msg(out, 'Reading your photo and writing. This takes about 15 seconds.');
    api('generate', { image: state.image.data, mediaType: state.image.mediaType, story: $('#story').value, vibe: vibe })
      .then(function (j) {
        state.post = j.post;
        state.captions = Object.assign({ story: '', short: '', question: '' }, j.post.captions || {});
        state.tags = (j.post.hashtags || []).slice(0, MAX_TAGS).map(function (h) { return { tag: h.tag, tier: h.tier, why: h.why }; });
        state.readTags = (j.post.read && j.post.read.tags) || [];
        state.palette = (j.post.read && j.post.read.palette) || [];
        renderRead(j.post.read || {});
        switchCaption('story');
        renderTags();
        msg(out, '');
      })
      .catch(function (err) { msg(out, err.message, 'error'); })
      .then(function () { busy(btn, false); busy($('#retry'), false); });
  }

  function renderRead(read) {
    $('#read-summary').textContent = read.summary || '';
    var tags = clear($('#read-tags'));
    (read.tags || []).slice(0, 6).forEach(function (t) { tags.appendChild(el('span', { text: t })); });
    var grade = clear($('#read-grade'));
    var colors = (read.palette || []).filter(function (c) { return /^#[0-9a-f]{6}$/i.test(c); }).slice(0, 5);
    colors.forEach(function (c) { grade.appendChild(el('span', { style: 'background:' + c, title: c })); });
    $('#read-grade-wrap').hidden = !colors.length;
  }

  function switchCaption(key) {
    state.capKey = key;
    $all('#cap-tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.cap === key ? 'true' : 'false'); });
    $('#caption').value = state.captions[key] || '';
    updateCount();
  }

  function updateCount() {
    var n = ($('#caption').value || '').length;
    $('#cap-count').textContent = n.toLocaleString() + ' / 2,200';
  }

  function renderTags() {
    var box = clear($('#tags'));
    if (!state.tags.length) box.appendChild(el('p', { class: 'empty-note', text: 'Your five hashtags will appear here.' }));
    state.tags.forEach(function (t, i) {
      box.appendChild(el('span', { class: 'tag', title: t.why || '' }, [
        t.tag,
        t.tier ? el('span', { class: 'tier', text: t.tier }) : null,
        el('button', { type: 'button', 'aria-label': 'Remove ' + t.tag, text: '×', onclick: function () { state.tags.splice(i, 1); renderTags(); } })
      ]));
    });
    $('#tag-count').textContent = state.tags.length + ' / ' + MAX_TAGS;
  }

  function composePost() {
    var caption = (state.captions[state.capKey] || '').trim();
    var tags = state.tags.map(function (t) { return t.tag; }).join(' ');
    return caption + (tags ? '\n\n' + tags : '');
  }

  /* ------------------------------------------------------------ */
  /* Planner                                                       */
  /* ------------------------------------------------------------ */

  function loadPlanner(force) {
    if (state.drafts && !force) { renderPlanner(); return; }
    clear($('#draft-list')).appendChild(el('p', { class: 'empty-note', text: 'Loading your drafts…' }));
    api('listDrafts').then(function (j) {
      state.drafts = j.drafts || [];
      renderPlanner();
    }).catch(function (err) {
      clear($('#draft-list')).appendChild(el('p', { class: 'empty-note', text: err.message }));
    });
  }

  function renderPlanner() {
    var drafts = state.drafts || [];
    var grid = clear($('#feed-grid'));
    var cells = drafts.slice(0, 12);
    cells.forEach(function (d) {
      grid.appendChild(el('button', { type: 'button', class: 'feed-cell', title: d.caption || '', onclick: function () { focusDraft(d.id); } }, [
        d.thumb ? el('img', { src: d.thumb, alt: d.caption ? d.caption.slice(0, 80) : 'Planned post' }) : null,
        el('span', { class: 'date', text: d.date ? prettyDate(d.date) : 'No date' })
      ]));
    });
    for (var i = cells.length; i < 9; i++) grid.appendChild(el('span', { class: 'feed-cell empty', text: i === cells.length ? 'NEXT' : '' }));

    var list = clear($('#draft-list'));
    if (!drafts.length) {
      list.appendChild(el('p', { class: 'empty-note' }, ['Nothing planned yet. Write a caption in the ', el('a', { href: '#/studio', text: 'Studio' }), ' and press Save to planner.']));
      return;
    }
    drafts.forEach(function (d) {
      var dateInput = el('input', { type: 'date', value: d.date || '', 'aria-label': 'Planned date' });
      var statusSel = el('select', { 'aria-label': 'Status' }, ['Draft', 'Ready', 'Posted'].map(function (s) {
        return el('option', { value: s, text: s, selected: s === d.status });
      }));
      var save = function () {
        d.date = dateInput.value; d.status = statusSel.value;
        api('saveDraft', { draft: { id: d.id, date: d.date, status: d.status, pillar: d.pillar, caption: d.caption, hashtags: d.hashtags, story: d.story, readTags: d.readTags, palette: d.palette } })
          .then(function () { toast('Updated.'); renderPlanner(); })
          .catch(function (err) { toast(err.message); });
      };
      dateInput.addEventListener('change', save);
      statusSel.addEventListener('change', save);
      list.appendChild(el('article', { class: 'draft', id: 'draft-' + d.id }, [
        d.thumb ? el('img', { src: d.thumb, alt: '' }) : el('span', { class: 'noimg', 'aria-hidden': 'true' }),
        el('div', {}, [
          el('div', { class: 'meta' }, [
            el('span', { class: 'status ' + d.status, text: d.status }),
            el('span', { class: 'hint', style: 'margin:0', text: prettyDate(d.date) + (d.pillar ? ' · ' + d.pillar : '') })
          ]),
          el('p', { class: 'cap', text: d.caption || '(no caption)' }),
          el('div', { class: 'controls' }, [
            dateInput, statusSel,
            el('button', { type: 'button', class: 'btn-text', text: 'Copy', onclick: function () {
              copyText((d.caption || '') + (d.hashtags && d.hashtags.length ? '\n\n' + d.hashtags.join(' ') : ''))
                .then(function () { toast('Copied.'); });
            } }),
            el('button', { type: 'button', class: 'btn-text', text: 'Delete', onclick: function () {
              if (!window.confirm('Delete this draft?')) return;
              api('deleteDraft', { id: d.id }).then(function () {
                state.drafts = state.drafts.filter(function (x) { return x.id !== d.id; });
                renderPlanner();
              }).catch(function (err) { toast(err.message); });
            } })
          ])
        ])
      ]));
    });
  }

  function focusDraft(id) {
    var node = document.getElementById('draft-' + id);
    if (node) { node.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
  }

  function initPlanner() {
    $('#direction-btn').addEventListener('click', function () {
      var btn = $('#direction-btn'), box = $('#direction');
      busy(btn, true, 'Reading');
      api('direction').then(function (j) {
        var d = j.direction || {};
        clear(box);
        box.appendChild(el('p', { class: 'trend-summary', style: 'margin:0;font-size:21px', text: d.summary || '' }));
        if ((d.pillars || []).length) {
          box.appendChild(el('h3', { text: 'Pillar balance' }));
          box.appendChild(el('div', { class: 'list-ruled' }, d.pillars.map(function (p) {
            return el('div', { class: 'item' }, [
              el('div', { class: 'between' }, [el('strong', { text: p.pillar }), el('span', { class: 'hint', style: 'margin:0', text: p.count + (p.count === 1 ? ' post' : ' posts') })]),
              el('div', { class: 'item-sub', text: p.advice })
            ]);
          })));
        }
        if ((d.nextPosts || []).length) {
          box.appendChild(el('h3', { text: 'Next posts' }));
          box.appendChild(el('div', { class: 'list-ruled' }, d.nextPosts.map(function (n) {
            return el('div', { class: 'item' }, [
              el('div', { class: 'item-title', text: n.idea }),
              el('div', { class: 'item-sub', text: [n.when, n.pillar].filter(Boolean).join(' · ') }),
              el('div', { class: 'item-tip', text: n.why })
            ]);
          })));
        }
      }).catch(function (err) { clear(box).appendChild(el('p', { class: 'form-msg error', text: err.message })); })
        .then(function () { busy(btn, false); });
    });
  }

  /* ------------------------------------------------------------ */
  /* Trends                                                        */
  /* ------------------------------------------------------------ */

  function loadTrends() {
    if (state.trends) { renderTrends(); return; }
    $('#trend-summary').textContent = 'Loading trends…';
    api('trends').then(function (j) {
      state.trends = j.trends || { regions: {} };
      var regions = Object.keys(state.trends.regions || {});
      var sel = clear($('#trend-region'));
      regions.forEach(function (r) { sel.appendChild(el('option', { value: r, text: r })); });
      state.region = regions.indexOf('Korea') >= 0 ? 'Korea' : regions[0] || '';
      sel.value = state.region;
      renderNiches();
      renderTrends();
    }).catch(function (err) { $('#trend-summary').textContent = err.message; });
  }

  function renderNiches() {
    var box = clear($('#niche-filter'));
    var niches = ['all'].concat((state.trends && state.trends.niches) || []);
    niches.forEach(function (n) {
      box.appendChild(el('button', { type: 'button', 'aria-pressed': n === state.niche ? 'true' : 'false', text: n === 'all' ? 'All niches' : n, onclick: function () {
        state.niche = n; renderNiches(); renderTrends();
      } }));
    });
  }

  function matchesNiche(x) {
    if (state.niche === 'all') return true;
    var list = [].concat(x.niches || x.niche || []).map(function (s) { return String(s).toLowerCase(); });
    return !list.length || list.indexOf(state.niche.toLowerCase()) >= 0;
  }

  function renderTrends() {
    var t = state.trends || { regions: {} };
    var data = (t.regions || {})[state.region];
    var formats = clear($('#trend-formats')), topics = clear($('#trend-topics')), tags = clear($('#trend-tags')), sources = clear($('#trend-sources'));
    if (!data) {
      $('#trend-summary').textContent = 'No trend data yet. The first refresh runs at 6:00 tomorrow morning (Seoul time).';
      $('#trends-updated').textContent = 'Refreshed every morning.';
      return;
    }
    $('#trends-updated').textContent = 'Last updated ' + (data.date || t.updated || '') + '. Refreshed every morning.';
    $('#trend-summary').textContent = data.summary || '';
    (data.formats || []).forEach(function (f) {
      formats.appendChild(el('div', { class: 'item' }, [
        el('div', { class: 'item-title', text: f.name }),
        el('div', { class: 'item-sub', text: f.signal }),
        f.tip ? el('div', { class: 'item-tip', text: f.tip }) : null
      ]));
    });
    var tlist = (data.topics || []).filter(matchesNiche);
    if (!tlist.length) topics.appendChild(el('p', { class: 'empty-note', text: 'Nothing specific for this niche today.' }));
    tlist.forEach(function (x) {
      topics.appendChild(el('div', { class: 'item' }, [
        el('div', { class: 'item-title', text: x.title }),
        el('div', { class: 'item-sub', text: [x.timing, [].concat(x.niches || []).join(', ')].filter(Boolean).join(' · ') }),
        x.why ? el('div', { class: 'item-tip', text: x.why }) : null
      ]));
    });
    var hlist = (data.hashtags || []).filter(matchesNiche);
    if (!hlist.length) tags.appendChild(el('p', { class: 'empty-note', text: 'No rising tags for this niche today.' }));
    hlist.forEach(function (h) {
      tags.appendChild(el('span', { class: 'tag static', title: h.note || '' }, [
        h.tag, el('span', { class: 'tier', text: h.trend === 'rising' ? 'rising' : (h.niche || '') })
      ]));
    });
    (data.sources || []).forEach(function (s) {
      sources.appendChild(el('li', {}, [el('a', { href: s.url, target: '_blank', rel: 'noopener', text: s.title || s.url })]));
    });
    if (!(data.sources || []).length) sources.appendChild(el('li', { text: 'No sources listed for this refresh.' }));
  }

  function initTrends() {
    $('#trend-region').addEventListener('change', function (e) { state.region = e.target.value; renderTrends(); });
  }

  /* ------------------------------------------------------------ */
  /* Creative                                                      */
  /* ------------------------------------------------------------ */

  function initCreative() {
    $('#ideas-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var f = e.target, out = $('.form-msg', f), btn = $('button[type=submit]', f), box = $('#ideas-out');
      var kind = ($('input[name=kind]:checked', f) || {}).value || 'single';
      busy(btn, true, 'Thinking'); msg(out, '');
      api('ideas', { kind: kind, brief: f.brief.value }).then(function (j) {
        renderIdeas(j.ideas || {}, j.kind);
      }).catch(function (err) { msg(out, err.message, 'error'); })
        .then(function () { busy(btn, false); });
    });
  }

  function copyRow(text) {
    return el('div', { class: 'copyable' }, [
      el('span', { text: text }),
      el('button', { type: 'button', class: 'btn-text', text: 'Copy', onclick: function () { copyText(text).then(function () { toast('Copied.'); }); } })
    ]);
  }

  function renderIdeas(d, kind) {
    var box = clear($('#ideas-out'));
    var wrap = el('div', { class: 'ideas' });
    wrap.appendChild(el('span', { class: 'eyebrow', text: 'Concepts' }));
    (d.concepts || []).forEach(function (c) {
      wrap.appendChild(el('div', { class: 'concept' }, [
        el('div', { class: 'item-title', text: c.title }),
        el('div', { class: 'item-tip', text: c.shot }),
        el('div', { class: 'item-sub', text: c.why })
      ]));
    });
    var seqTitle = kind === 'carousel' ? 'Slide order' : kind === 'reel' ? 'Beats' : 'Frames to try';
    if ((d.sequence || []).length) {
      wrap.appendChild(el('h3', { text: seqTitle }));
      wrap.appendChild(el('ol', {}, d.sequence.map(function (s) {
        return el('li', {}, [el('strong', { text: s.label }), ' — ', s.purpose]);
      })));
    }
    if ((d.hooks || []).length) { wrap.appendChild(el('h3', { text: 'Hooks' })); d.hooks.forEach(function (h) { wrap.appendChild(copyRow(h)); }); }
    if ((d.coverText || []).length) { wrap.appendChild(el('h3', { text: 'Cover text' })); d.coverText.forEach(function (h) { wrap.appendChild(copyRow(h)); }); }
    if (d.gradeTip) { wrap.appendChild(el('h3', { text: 'Color and edit' })); wrap.appendChild(el('p', { style: 'margin:0', text: d.gradeTip })); }
    box.appendChild(wrap);
  }

  /* ------------------------------------------------------------ */
  /* Demo data (open the site with ?demo=1)                        */
  /* ------------------------------------------------------------ */

  var demoDrafts = [];
  function mockApi(action, b) {
    var member = { name: 'Demo', email: 'demo@stillsnote.com', plan: 'Free', uses: 4, limit: 30 };
    var res = { ok: true, member: member };
    if (action === 'login') { res.token = 'demo'; }
    if (action === 'signup') { res.status = 'pending'; }
    if (action === 'getProfile' || action === 'saveProfile') { res.profile = b.profile || { niche: 'hiking, cafes, Seoul', tone: 'dry, warm, lowercase', language: 'English', pillars: ['hikes', 'cafes', 'food', 'city at night'] }; }
    if (action === 'generate') {
      res.post = {
        read: { tags: ['summit shot', 'city skyline', 'golden light', 'two people', 'cold morning'], palette: ['#D9843F', '#F3D3AE', '#8A847D', '#4A443F', '#F1ECE4'], summary: 'Two hikers on a rocky summit at sunrise, the city hazy and gold below.' },
        captions: {
          story: 'made my brother climb bukhansan at 5am. he complained the entire way up, then took like 200 photos at the top.\n\nseoul looked unreal this morning. worth every frozen finger.\n\nwho would you drag up a mountain at sunrise?',
          short: 'bukhansan, 6:12am. brother still mad about it.',
          question: 'sunrise hike before work: genius or insane? asking for my brother, who has opinions.'
        },
        hashtags: [{ tag: '#hiking', tier: 'big', why: 'Broad context' }, { tag: '#seoulhiking', tier: 'medium', why: 'Active local hiking community' }, { tag: '#sunrisehike', tier: 'medium', why: 'Matches the moment' }, { tag: '#bukhansan', tier: 'local', why: 'The exact place' }, { tag: '#북한산', tier: 'local', why: 'Korean searches for the mountain' }],
        altText: 'Two hikers on a granite summit at sunrise above Seoul.'
      };
    }
    if (action === 'saveDraft') {
      var d = b.draft; if (!d.id) { d.id = 'd' + Date.now(); demoDrafts.unshift(d); } else { demoDrafts = demoDrafts.map(function (x) { return x.id === d.id ? Object.assign(x, d) : x; }); }
      res.id = d.id;
    }
    if (action === 'listDrafts') { res.drafts = demoDrafts.slice(); }
    if (action === 'deleteDraft') { demoDrafts = demoDrafts.filter(function (x) { return x.id !== b.id; }); }
    if (action === 'trends') {
      res.trends = { updated: today(), niches: ['travel', 'food', 'cafe', 'fitness', 'outdoors'], regions: { Korea: {
        date: today(), summary: 'Autumn foliage and first-snow content is peaking, and slow, quiet carousels are out-performing fast edits for travel and cafe accounts.',
        formats: [{ name: 'Photo carousels, 6 to 10 slides', signal: 'Carousels are getting extra reach from re-showing to people who skipped slide one.', tip: 'Lead with your strongest frame and end on a detail shot.' }, { name: 'Quiet Reels with original sound', signal: 'Ambient, unhurried clips are holding watch time.', tip: 'Keep cuts slow and let the place speak.' }],
        topics: [{ title: 'Late foliage hikes', why: 'Peak colour is moving south.', timing: 'next 2 weeks', niches: ['outdoors', 'travel'] }, { title: 'Winter cafe menus', why: 'Seasonal drinks launch across Seoul.', timing: 'mid-November', niches: ['cafe', 'food'] }],
        hashtags: [{ tag: '#단풍', niche: 'outdoors', trend: 'rising', note: 'Foliage season' }, { tag: '#seoulcafe', niche: 'cafe', trend: 'steady', note: 'Always searched' }, { tag: '#등산스타그램', niche: 'outdoors', trend: 'rising', note: 'Hiking community' }],
        sources: [{ title: 'Example source', url: 'https://example.com' }]
      } } };
    }
    if (action === 'ideas') {
      res.kind = b.kind;
      res.ideas = {
        concepts: [{ title: 'The empty cafe hour', shot: 'Shoot at opening, window light, your cup in the foreground and the street soft behind it.', why: 'Quiet scenes match what your followers save.' }, { title: 'Sea, then coffee', shot: 'Pair a wide beach frame with a close-up of foam art, same warm grade.', why: 'Contrast keeps people swiping.' }],
        sequence: [{ label: 'Slide 1', purpose: 'Widest, most atmospheric frame' }, { label: 'Slide 2', purpose: 'You in the scene, small' }, { label: 'Slide 3', purpose: 'A texture or detail' }, { label: 'Last slide', purpose: 'A quiet closer that invites a save' }],
        hooks: ['slow weekend, no plans, perfect', 'gangneung in 6 frames'], coverText: ['slow weekend', 'east coast, quiet'],
        gradeTip: 'Lift the shadows a little and warm the whites so it sits with your amber-toned feed.'
      };
    }
    if (action === 'direction') {
      res.direction = { summary: 'You have leaned on hikes lately. Balance with a cafe post this week and a night city post on the weekend.', pillars: [{ pillar: 'hikes', count: 3, advice: 'Strong, but give it a week off.' }, { pillar: 'cafes', count: 0, advice: 'Add one this week.' }], nextPosts: [{ idea: 'Window seat at your favorite cafe at opening', pillar: 'cafes', when: 'Wednesday morning', why: 'Fills the gap and fits the winter menu trend.' }] };
    }
    return new Promise(function (resolve) { setTimeout(function () { if (res.member) setMember(res.member); resolve(res); }, 500); });
  }

  /* ------------------------------------------------------------ */
  /* Start                                                         */
  /* ------------------------------------------------------------ */

  function start() {
    if (DEMO) { $('#demo-banner').hidden = false; }
    initGate(); initStudio(); initPlanner(); initTrends(); initCreative(); initStyle();
    $('#signout').addEventListener('click', function () { signOut(false); });
    window.addEventListener('hashchange', route);
    route();
    if (state.token) {
      api('me').then(function () { loadProfile(); }).catch(function () { /* handled by api() */ });
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
