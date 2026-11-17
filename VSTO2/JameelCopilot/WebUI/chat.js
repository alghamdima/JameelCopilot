/* ============================================================
   JameelCopilot Chat - WebView2 surface controller (Phase 2)
   Public API (all on window.jameelcopilot):
     appendUserMessage(text)
     appendAssistantMessage(messageId, initialText)
     appendTextDelta(messageId, delta)
     finalizeAssistantMessage(messageId, opts)
      appendToolCallCard(callId, name, argsJson)
      updateToolCallCard(callId, ok, summary, resultJson)
      onFileSaved(messageId, fileInfo)
      onExportError(messageId, error)
      appendAuditRow(text)
     showError(message)
     setComposerEnabled(enabled, isStopVisible)
     clear()
     applyTheme(themeName)              // "light" | "dark" | "high-contrast"
     setContextStrip({subject, recipients, thread})
     setStyle(style)                    // "precise" | "balanced" | "creative"; null hides the picker
     setQuickActions(groups)            // titled groups of chips
     loadConversation(messages)         // show a saved chat
     setChats(list, currentId)          // the Chats view
     setMemory(notes)                   // the Memory view
     beginDraft(id, title) / finalizeDraft(id, buttons, opts) / markDraft(id, status, done)
     setComposeMode(on)                 // the Draft button, and Insert on answers
   Host -> JS:
     C# calls these via WebView2.ExecuteScriptAsync("jameelcopilot.X(...)").
   JS -> Host:
     window.chrome.webview.postMessage(JSON.stringify({type:..., payload:...}))
      Message types: 'send', 'stop', 'copy', 'toolCardClicked',
                     'open_file', 'reveal_in_explorer', 'export_pdf',
                     'set_style', 'new_chat', 'show_chats', 'open_chat',
                     'delete_chat', 'clear_chats', 'show_memory',
                     'memory_add', 'memory_update', 'memory_remove',
                     'chip_action', 'draft', 'draft_action', 'insert_text', 'open_message',
                     'priority_reply', 'review_apply', 'meeting_open', 'tidy_apply',
                     'tidy_undo', 'followup', 'task_add', 'rule_create',
                     'show_prompts', 'prompt_add', 'prompt_update', 'prompt_remove'
   ============================================================ */

(function() {
  'use strict';

  // -- DOM refs ----------------------------------------------------
  var $messages = document.getElementById('messages');
  var $input = document.getElementById('composerInput');
  var $btnSend = document.getElementById('btnSend');
  var $btnStop = document.getElementById('btnStop');
  var $btnDraft = document.getElementById('btnDraft');
  var $btnCopy = document.getElementById('btnCopy');
  var $btnNewChat = document.getElementById('btnNewChat');
  var $btnChats = document.getElementById('btnChats');
  var $btnMemory = document.getElementById('btnMemory');
  var $viewChat = document.getElementById('viewChat');
  var $viewChats = document.getElementById('viewChats');
  var $viewMemory = document.getElementById('viewMemory');
  var $chatList = document.getElementById('chatList');
  var $memoryList = document.getElementById('memoryList');
  var $memoryInput = document.getElementById('memoryInput');
  var $viewPrompts = document.getElementById('viewPrompts');
  var $promptList = document.getElementById('promptList');
  var $promptInput = document.getElementById('promptInput');
  var $ctxSubject = document.getElementById('ctxSubject');
  var $ctxRecipients = document.getElementById('ctxRecipients');
  var $ctxThread = document.getElementById('ctxThread');
  var $quickActions = document.getElementById('quickActions');
  var $stylePicker = document.getElementById('stylePicker');
  var $styleSelect = document.getElementById('styleSelect');

  // -- Bridge to host ----------------------------------------------
  function postToHost(obj) {
    try {
      if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {
        window.chrome.webview.postMessage(JSON.stringify(obj));
      } else {
        // Dev fallback: log to console.
        console.log('[hostPost]', obj);
      }
    } catch (err) {
      console.error('postToHost failed', err);
    }
  }

  function fallbackEscapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeAttr(s) {
    return fallbackEscapeHtml(s);
  }

  function cssEscape(s) {
    if (window.CSS && typeof window.CSS.escape === 'function') {
      return window.CSS.escape(String(s));
    }
    return String(s).replace(/(["\\\]\[])/g, '\\$1');
  }

  function renderMarkdown(src) {
    if (window.markdown && typeof window.markdown.render === 'function') {
      return window.markdown.render(src);
    }
    if (!src) return '';
    return '<p>' + fallbackEscapeHtml(src).replace(/\n/g, '<br>') + '</p>';
  }

  // -- DOM-builder helpers ----------------------------------------
  function empty(node) {
    while (node.firstChild) node.removeChild(node.firstChild);
  }

  function button(cls, text) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = cls;
    btn.textContent = text;
    return btn;
  }

  var composeMode = false;
  var hasSingleSelection = false;
  var currentAnswerId = null;

  // The emails the tools found or read during this turn: listed under the
  // answer, each one opens in Outlook. Built from the tool results, never
  // from the model's text.
  var pendingSources = [];
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function collectSources(resultJson) {
    var obj;
    try { obj = (typeof resultJson === 'string') ? JSON.parse(resultJson) : resultJson; } catch (e) { return; }
    if (!obj || typeof obj !== 'object') return;
    var list = Array.isArray(obj.messages) ? obj.messages : (typeof obj.id === 'string' && 'subject' in obj ? [obj] : []);
    list.forEach(function(m) {
      if (!m || typeof m.id !== 'string' || !m.id || pendingSources.length >= 50) return;
      for (var i = 0; i < pendingSources.length; i++) if (pendingSources[i].id === m.id) return;
      pendingSources.push({
        id: m.id,
        subject: typeof m.subject === 'string' ? m.subject : '',
        from: typeof m.from === 'string' ? m.from : '',
        when: typeof m.received_at === 'string' ? m.received_at : ''
      });
    });
  }

  function shortDate(iso) {
    var d = new Date(iso);
    return isNaN(d.getTime()) ? '' : d.getDate() + ' ' + MONTHS[d.getMonth()];
  }

  // One email as a row that opens it in Outlook.
  function sourceRow(s) {
    var row = button('source-row', '');
    var from = elt('span', 'source-from', s.from || '(no sender)');
    var subject = elt('span', 'source-subject', s.subject || '(no subject)');
    var when = elt('span', 'source-when', shortDate(s.when));
    from.dir = 'auto';
    subject.dir = 'auto';
    row.appendChild(from);
    row.appendChild(subject);
    row.appendChild(when);
    row.title = 'Open this email';
    row.addEventListener('click', function() {
      postToHost({ type: 'open_message', payload: { id: s.id } });
    });
    return row;
  }

  function renderSources(sources) {
    var box = elt('div', 'msg-sources');
    box.appendChild(elt('div', 'sources-title', 'Emails used (' + sources.length + ') - click one to open it'));
    var shown = 5;
    var rows = sources.map(function(s, i) {
      var row = sourceRow(s);
      row.hidden = i >= shown;
      box.appendChild(row);
      return row;
    });
    if (rows.length > shown) {
      var more = button('sources-more', 'Show ' + (rows.length - shown) + ' more');
      more.addEventListener('click', function() {
        rows.forEach(function(r) { r.hidden = false; });
        box.removeChild(more);
      });
      box.appendChild(more);
    }
    return box;
  }

  // Deleting cannot be undone and the page has no dialogs: the first click
  // asks on the button itself, the second one (within a few seconds) does it.
  function confirmThen(btn, question, action) {
    var label = btn.textContent;
    var timer = 0;
    function disarm() {
      clearTimeout(timer);
      btn.dataset.armed = '';
      btn.textContent = label;
      btn.classList.remove('confirming');
    }
    btn.addEventListener('click', function() {
      if (btn.dataset.armed === '1') {
        disarm();
        action();
        return;
      }
      btn.dataset.armed = '1';
      btn.textContent = question;
      btn.classList.add('confirming');
      timer = setTimeout(disarm, 4000);
    });
  }

  // One view at a time: the chat, the saved chats, or the memory.
  var currentView = 'chat';
  function showView(name) {
    currentView = name;
    $viewChat.hidden = name !== 'chat';
    $viewChats.hidden = name !== 'chats';
    $viewMemory.hidden = name !== 'memory';
    $viewPrompts.hidden = name !== 'prompts';
    $btnChats.classList.toggle('on', name === 'chats');
    $btnMemory.classList.toggle('on', name === 'memory');
  }

  function elt(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }

  function formatBytes(bytes) {
    var n = Number(bytes);
    if (!isFinite(n) || n < 0) return '';
    if (n < 1024) return Math.round(n) + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    if (n < 1024 * 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + ' MB';
    return (n / 1024 / 1024 / 1024).toFixed(1) + ' GB';
  }

  function formatLabel(format) {
    switch (String(format || '').toLowerCase()) {
      case 'xlsx': return 'Excel Workbook';
      case 'pdf': return 'PDF Document';
      default: return format ? String(format).toUpperCase() : 'File';
    }
  }

  function filenameFromPath(path) {
    var value = String(path || '');
    var slash = Math.max(value.lastIndexOf('\\'), value.lastIndexOf('/'));
    return slash >= 0 ? value.substring(slash + 1) : value;
  }

  function postFileAction(type, path) {
    if (!path) return;
    postToHost({
      type: type,
      payload: { path: path }
    });
  }

  function findAssistantMessage(messageId) {
    if (messageId !== undefined && messageId !== null && messageId !== '') {
      var selector = '[data-message-id="' + cssEscape(messageId) + '"]';
      var matched = null;
      try { matched = document.querySelector(selector); } catch (e) { matched = null; }
      if (matched) return matched;
    }

    var messages = document.querySelectorAll('.msg-assistant');
    return messages.length ? messages[messages.length - 1] : null;
  }

  function appendFileCardToMessage(messageId, fileInfo) {
    if (!fileInfo || !fileInfo.path) return;

    var msgEl = findAssistantMessage(messageId);
    if (!msgEl) return;

    var attach = msgEl.querySelector('.msg-attachments');
    if (!attach) {
      attach = elt('div', 'msg-attachments');
      msgEl.appendChild(attach);
    }

    var filePath = String(fileInfo.path);
    var filename = fileInfo.filename || filenameFromPath(filePath) || 'Saved file';
    var format = String(fileInfo.format || '').toLowerCase();
    var bytes = formatBytes(fileInfo.bytes);
    var label = formatLabel(fileInfo.format);
    var sub = bytes ? (bytes + ' \u00B7 ' + label) : label;

    var card = elt('div', 'file-card');
    card.setAttribute('data-format', escapeAttr(format));

    card.appendChild(elt('div', 'file-card-icon'));

    var meta = elt('div', 'file-card-meta');
    var name = elt('div', 'file-card-name', filename);
    name.title = filename;
    meta.appendChild(name);
    meta.appendChild(elt('div', 'file-card-sub', sub));
    card.appendChild(meta);

    var actions = elt('div', 'file-card-actions');
    var open = elt('button', 'file-card-btn', 'Open');
    open.type = 'button';
    open.addEventListener('click', function() {
      postFileAction('open_file', filePath);
    });

    var reveal = elt('button', 'file-card-btn', 'Show in folder');
    reveal.type = 'button';
    reveal.addEventListener('click', function() {
      postFileAction('reveal_in_explorer', filePath);
    });

    actions.appendChild(open);
    actions.appendChild(reveal);
    card.appendChild(actions);

    attach.appendChild(card);
    scrollToBottom();
  }

  var retryableExportErrorCodes = [
    'file_locked',
    'webview2_missing',
    'path_timeout',
    'pdf_render_timeout'
  ];

  function exportErrorCode(err) {
    try {
      if (!err || typeof err === 'string') return '';
      if (err.error && typeof err.error === 'object') return String(err.error.code || '');
      return String(err.error || '');
    } catch (e) {
      return '';
    }
  }

  function isRetryableExportError(err) {
    return retryableExportErrorCodes.indexOf(exportErrorCode(err)) >= 0;
  }

  function renderInlineErrorCard(messageId, err) {
    var msgEl = findAssistantMessage(messageId);
    if (!msgEl) return false;

    var retryMessageId = messageId;
    if (msgEl.dataset && msgEl.dataset.messageId) {
      retryMessageId = msgEl.dataset.messageId;
    }

    var attach = msgEl.querySelector('.msg-attachments');
    if (!attach) {
      attach = elt('div', 'msg-attachments');
      msgEl.appendChild(attach);
    }

    var card = elt('div', 'error-card');
    card.appendChild(elt('div', 'error-card-icon'));

    var body = elt('div', 'error-card-body');
    body.appendChild(elt('div', 'error-card-title', 'Export failed'));
    var detail = elt('div', 'error-card-detail');
    detail.textContent = exportErrorMessage(err);
    body.appendChild(detail);

    var code = exportErrorCode(err);
    var actions = elt('div', 'error-card-actions');
    var hasActions = false;

    if (isRetryableExportError(err) && retryMessageId !== undefined && retryMessageId !== null && retryMessageId !== '') {
      var retryEntry = assistantMessages[retryMessageId];
      if (retryEntry && retryEntry.exportButton) {
        var retry = elt('button', 'error-card-btn', 'Retry');
        retry.type = 'button';
        retry.addEventListener('click', function() {
          handleExportPdf(retryMessageId);
        });
        actions.appendChild(retry);
        hasActions = true;
      }
    }

    if (code === 'webview2_missing') {
      var install = elt('span', 'error-card-link', 'Ask IT to install the WebView2 runtime.');
      actions.appendChild(install);
      hasActions = true;
    }

    if (hasActions) body.appendChild(actions);
    card.appendChild(body);
    attach.appendChild(card);

    resetExportButton(retryMessageId);
    scrollToBottom();
    return true;
  }

  function renderToolResultIfHandled(messageId, resultJson) {
    if (!resultJson) return false;
    try {
      var obj = (typeof resultJson === 'string') ? JSON.parse(resultJson) : resultJson;
      if (obj && obj.result_type === 'file_saved') {
        appendFileCardToMessage(obj.message_id || obj.messageId || messageId, obj);
        return true;
      }
      if (obj && obj.result_type === 'meeting_proposal' && Array.isArray(obj.slots) && obj.slots.length) {
        placeCard(buildMeetingProposal(messageId, obj));
        return true;
      }
      if (obj && obj.result_type === 'memory_proposal' && typeof obj.note === 'string' && obj.note) {
        placeCard(buildMemoryProposal(obj.note));
        return true;
      }
      if (obj && obj.result_type === 'rule_proposal' && typeof obj.name === 'string' && obj.name) {
        placeCard(buildRuleProposal(messageId, obj));
        return true;
      }
    } catch (e) { /* fall through to existing path */ }
    return false;
  }

  var DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  function hm(d) {
    return isNaN(d.getTime()) ? '' : ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
  }

  function slotLabel(slot) {
    var s = new Date(slot.start), e = new Date(slot.end);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) return String(slot.start || '');
    return DAYS[s.getDay()] + ' ' + s.getDate() + ' ' + MONTHS[s.getMonth()] + '  ' + hm(s) + ' - ' + hm(e);
  }

  // A meeting Jameel Copilot proposes: the user picks a location/room and a time, and Outlook's meeting window opens.
  function buildMeetingProposal(callId, p) {
    var node = elt('div', 'msg msg-assistant msg-draft msg-meeting');
    var head = elt('div', 'draft-head');
    head.appendChild(elt('span', 'draft-title', 'Meeting proposal'));
    var status = elt('span', 'draft-status', 'nothing sent');
    head.appendChild(status);
    var body = elt('div', 'msg-content');
    var subject = elt('div', 'meeting-subject', p.subject || '');
    subject.dir = 'auto';
    body.appendChild(subject);
    if (Array.isArray(p.attendees) && p.attendees.length) {
      var who = elt('div', 'meeting-who', 'With: ' + p.attendees.join(', '));
      who.dir = 'auto';
      body.appendChild(who);
    }
    if (Array.isArray(p.optional_attendees) && p.optional_attendees.length) {
      var optWho = elt('div', 'meeting-who', 'Optional (CC): ' + p.optional_attendees.join(', '));
      optWho.dir = 'auto';
      body.appendChild(optWho);
    }
    var roomSelect = null;
    var roomLabel = (typeof p.room_display === 'string' && p.room_display) ? p.room_display : (typeof p.room === 'string' ? p.room : '');
    if (Array.isArray(p.room_options) && p.room_options.length) {
      var roomRow = elt('div', 'meeting-who');
      roomRow.appendChild(document.createTextNode('Location: '));
      roomSelect = document.createElement('select');
      roomSelect.className = 'meeting-room-select';
      p.room_options.forEach(function(opt, idx) {
        var o = document.createElement('option');
        o.value = String(idx);
        o.textContent = opt.label || opt.value || '';
        if (p.room && (opt.value === p.room || opt.label === p.room)) o.selected = true;
        roomSelect.appendChild(o);
      });
      roomRow.appendChild(roomSelect);
      body.appendChild(roomRow);
    } else if (roomLabel) {
      var room = elt('div', 'meeting-who', 'Room: ' + roomLabel);
      room.dir = 'auto';
      body.appendChild(room);
    }
    if (typeof p.agenda === 'string' && p.agenda) {
      var agenda = elt('div', 'meeting-agenda', p.agenda);
      agenda.dir = 'auto';
      body.appendChild(agenda);
    }
    body.appendChild(elt('div', 'meeting-pick', 'Pick a time to open the meeting request:'));
    var foot = elt('div', 'draft-foot');
    p.slots.forEach(function(slot, i) {
      var b = button('btn btn-small' + (i === 0 ? ' btn-primary' : ''), slotLabel(slot));
      b.addEventListener('click', function() {
        var payload = { call: callId, slot: i };
        if (roomSelect) payload.room_index = parseInt(roomSelect.value, 10);
        postToHost({ type: 'meeting_open', payload: payload });
        status.textContent = 'opened in Outlook - send it from there';
      });
      foot.appendChild(b);
    });
    node.appendChild(head);
    node.appendChild(body);
    node.appendChild(foot);
    return node;
  }

  // An Outlook rule Jameel Copilot proposes. Nothing exists until the user presses Create rule;
  // the host creates the rule it kept for this card, not what the page shows.
  var ruleCards = {};
  function buildRuleProposal(callId, p) {
    var node = elt('div', 'msg msg-assistant msg-draft msg-rule');
    var head = elt('div', 'draft-head');
    head.appendChild(elt('span', 'draft-title', 'Rule proposal'));
    var status = elt('span', 'draft-status', 'not created yet');
    head.appendChild(status);
    var body = elt('div', 'msg-content');
    var name = elt('div', 'meeting-subject', p.name || '');
    name.dir = 'auto';
    body.appendChild(name);
    function line(label, value) {
      if (!value) return;
      var row = elt('div', 'rule-line');
      row.appendChild(elt('span', 'rule-label', label));
      var text = elt('span', 'rule-value', value);
      text.dir = 'auto';
      row.appendChild(text);
      body.appendChild(row);
    }
    var from = (Array.isArray(p.from) ? p.from : []).join(', ');
    var words = (Array.isArray(p.subject_contains) ? p.subject_contains : []).join(', ');
    var folder = typeof p.folder === 'string' ? p.folder : '';
    line('When new mail is from', from);
    line(from ? 'and its subject contains' : 'When the subject of new mail contains', words);
    line('move it to the folder', folder);
    line(folder ? 'and give it the category' : 'give it the category', typeof p.category === 'string' ? p.category : '');
    var foot = elt('div', 'draft-foot');
    var create = button('btn btn-small btn-primary', 'Create rule');
    var no = button('btn btn-small btn-ghost', 'No');
    create.addEventListener('click', function() {
      create.disabled = true;
      postToHost({ type: 'rule_create', payload: { call: callId } });
    });
    no.addEventListener('click', function() {
      if (node.parentNode) node.parentNode.removeChild(node);
      delete ruleCards[callId];
    });
    foot.appendChild(create);
    foot.appendChild(no);
    node.appendChild(head);
    node.appendChild(body);
    node.appendChild(foot);
    ruleCards[callId] = { node: node, status: status, foot: foot };
    return node;
  }

  // A to-do row: what to do, by when, the email it comes from (a click opens it),
  // and "Add to tasks". The host makes the task from the row it kept: the page
  // only names the card and the row.
  function todoRow(cardId, index, item, rows) {
    var row = elt('div', 'todo-row');
    var main = elt('div', 'todo-main');
    var text = elt('div', 'todo-text', item.text || '');
    text.dir = 'auto';
    main.appendChild(text);
    var source = button('todo-source', '');
    if (item.due) source.appendChild(elt('span', 'todo-due', 'Due ' + shortDate(item.due + 'T00:00')));
    var where = elt('span', 'todo-from', (item.from || '') + (item.subject ? ' - ' + item.subject : ''));
    where.dir = 'auto';
    source.appendChild(where);
    source.title = 'Open the email';
    source.addEventListener('click', function() {
      postToHost({ type: 'open_message', payload: { id: item.id } });
    });
    main.appendChild(source);
    var add = button('btn btn-small btn-ghost todo-add', 'Add to tasks');
    add.title = 'Add this to your Outlook tasks';
    add.addEventListener('click', function() {
      add.disabled = true;
      postToHost({ type: 'task_add', payload: { id: cardId, index: index } });
    });
    row.appendChild(main);
    row.appendChild(add);
    rows[index] = { row: row, add: add };
    return row;
  }

  // A list of the user's one-line notes with Edit and Delete: the memory
  // ('memory') and their own prompts ('prompt').
  function renderNotes($list, notes, kind, emptyText) {
    empty($list);
    if (!notes || !notes.length) {
      $list.appendChild(elt('div', 'panel-empty', emptyText));
      return;
    }
    notes.forEach(function(note) {
      var row = elt('div', 'mem-row');
      var text = elt('div', 'mem-text', note.text);
      text.dir = 'auto';
      var actions = elt('div', 'mem-actions');
      var edit = button('btn btn-small btn-ghost', 'Edit');
      var del = button('btn btn-small btn-ghost', 'Delete');
      edit.addEventListener('click', function() {
        var box = document.createElement('textarea');
        box.className = 'composer-input mem-edit';
        box.dir = 'auto';
        box.maxLength = 300;
        box.value = note.text;
        var save = button('btn btn-small btn-primary', 'Save');
        var cancel = button('btn btn-small btn-ghost', 'Cancel');
        save.addEventListener('click', function() {
          postToHost({ type: kind + '_update', payload: { id: note.id, text: box.value } });
        });
        cancel.addEventListener('click', function() {
          postToHost({ type: kind === 'memory' ? 'show_memory' : 'show_prompts' });
        });
        row.replaceChild(box, text);
        empty(actions);
        actions.appendChild(save);
        actions.appendChild(cancel);
        box.focus();
      });
      confirmThen(del, 'Delete?', function() {
        postToHost({ type: kind + '_remove', payload: { id: note.id } });
      });
      actions.appendChild(edit);
      actions.appendChild(del);
      row.appendChild(text);
      row.appendChild(actions);
      $list.appendChild(row);
    });
  }

  // A card a tool asked for goes inside the answer being written, after its
  // text, so it is where the user reads; with no answer open, at the end.
  var pendingCards = [];
  function placeCard(node) {
    var open = currentAnswerId && assistantMessages[currentAnswerId];
    if (open && !open.complete) {
      pendingCards.push(node);
    } else {
      $messages.appendChild(node);
      scrollToBottom();
    }
  }

  // Jameel Copilot may only propose a note. It is kept when the user presses Save.
  function buildMemoryProposal(note) {
    var node = elt('div', 'msg msg-assistant msg-draft');
    var head = elt('div', 'draft-head');
    head.appendChild(elt('span', 'draft-title', 'Remember this?'));
    var status = elt('span', 'draft-status', 'not saved yet');
    head.appendChild(status);
    var content = elt('div', 'msg-content', note);
    content.dir = 'auto';
    var foot = elt('div', 'draft-foot');
    var save = button('btn btn-small btn-primary', 'Save to memory');
    var no = button('btn btn-small btn-ghost', 'No');
    save.addEventListener('click', function() {
      postToHost({ type: 'memory_add', payload: { text: note } });
      status.textContent = 'saved - see Memory';
      node.classList.add('draft-done');
      empty(foot);
    });
    no.addEventListener('click', function() {
      if (node.parentNode) node.parentNode.removeChild(node);
    });
    foot.appendChild(save);
    foot.appendChild(no);
    node.appendChild(head);
    node.appendChild(content);
    node.appendChild(foot);
    return node;
  }

  function normalizePdfExportToolError(resultJson) {
    if (!resultJson) return null;
    try {
      var obj = (typeof resultJson === 'string') ? JSON.parse(resultJson) : resultJson;
      if (!obj || !obj.error || typeof obj.error !== 'object') return null;
      return {
        error: obj.error.code || 'pdf_render_failed',
        detail: obj.error.message || obj.error.detail || 'Unknown export error.'
      };
    } catch (e) {
      return null;
    }
  }

  function exportErrorMessage(error) {
    var fallback = 'Unknown export error.';
    try {
      if (!error) return fallback;
      if (typeof error === 'string') return error || fallback;
      if (error.error && typeof error.error === 'object') {
        return String(error.error.message || error.error.detail || error.error.code || fallback);
      }
      var detail = error.detail || error.message || error.error;
      if (detail) return String(detail);
      return fallback;
    } catch (e) {
      return fallback;
    }
  }

  function scrollToBottom() {
    $messages.scrollTop = $messages.scrollHeight;
  }

  function deriveFilenameHint(markdown) {
    var text = String(markdown || '');
    var lines = text.split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var m = /^#{1,3}\s+(.+?)\s*$/.exec(lines[i]);
      if (m && m[1]) {
        var heading = m[1].trim();
        heading = heading.replace(/\s+#+\s*$/, '').trim();
        if (heading) return heading.substring(0, 60);
      }
    }
    return 'Jameel Copilot Report';
  }

  function createExportPdfButton(messageId) {
    var btn = elt('button', 'msg-action msg-action-pdf', 'PDF');
    btn.type = 'button';
    btn.title = 'Save message as PDF';
    btn.setAttribute('aria-label', 'Save message as PDF');
    btn.disabled = true;
    btn.dataset.defaultText = btn.textContent;
    btn.addEventListener('click', function() {
      handleExportPdf(messageId);
    });
    return btn;
  }

  function setExportButtonPending(messageId, pending) {
    var entry = assistantMessages[messageId];
    var btn = entry && entry.exportButton;
    if (!btn) return;
    btn.disabled = true;
    btn.dataset.exportPending = pending ? '1' : '0';
    if (pending) btn.textContent = 'Saving...';
  }

  function resetExportButton(messageId) {
    var entry = assistantMessages[messageId];
    var btn = entry && entry.exportButton;
    if (!btn) return;
    btn.dataset.exportPending = '0';
    btn.textContent = btn.dataset.defaultText || 'PDF';
    btn.disabled = !(entry && entry.complete);
  }

  function handleExportPdf(messageId) {
    var entry = assistantMessages[messageId];
    var msgEl = entry && entry.container;
    if (!msgEl && messageId !== undefined && messageId !== null && messageId !== '') {
      try { msgEl = document.querySelector('[data-message-id="' + cssEscape(messageId) + '"]'); } catch (e) { msgEl = null; }
    }
    if (!entry || !msgEl || !entry.complete) return;
    if (entry.exportButton && entry.exportButton.dataset.exportPending === '1') return;

    var markdown = entry.raw;
    if (markdown === undefined || markdown === null || markdown === '') {
      markdown = entry.content ? entry.content.innerText : '';
    }
    markdown = String(markdown || '');

    setExportButtonPending(messageId, true);
    postToHost({
      type: 'export_pdf',
      payload: {
        message_id: messageId,
        filename_hint: deriveFilenameHint(markdown),
        content_markdown: markdown
      }
    });
  }

  var assistantMessages = {}; // id -> { container, content, raw }
  var toolCards = {};         // callId -> element

  // -- Public API --------------------------------------------------
  var api = {
    appendUserMessage: function(text) {
      pendingSources = [];
      var node = elt('div', 'msg msg-user');
      var content = elt('div', 'msg-content');
      content.innerHTML = renderMarkdown(text);
      node.appendChild(content);
      $messages.appendChild(node);
      scrollToBottom();
    },

    appendAssistantMessage: function(id, initialText) {
      currentAnswerId = id;
      var node = elt('div', 'msg msg-assistant');
      node.dataset.messageId = id;
      node.dataset.state = 'streaming';
      node.classList.add('is-streaming');
      var exportButton = createExportPdfButton(id);
      node.appendChild(exportButton);
      var content = elt('div', 'msg-content');
      content.innerHTML = renderMarkdown(initialText || '');
      node.appendChild(content);
      $messages.appendChild(node);
      assistantMessages[id] = {
        container: node,
        content: content,
        raw: initialText || '',
        complete: false,
        exportButton: exportButton
      };
      scrollToBottom();
    },

    appendTextDelta: function(id, delta) {
      var entry = assistantMessages[id];
      if (!entry) {
        api.appendAssistantMessage(id, delta);
        return;
      }
      entry.raw += delta;
      // For streaming, render the running text as markdown. This is
      // cheap enough for typical email-length replies.
      entry.content.innerHTML = renderMarkdown(entry.raw);
      scrollToBottom();
    },

    finalizeAssistantMessage: function(id, opts) {
      var entry = assistantMessages[id];
      if (!entry) return;
      opts = opts || {};
      entry.complete = true;
      entry.container.dataset.state = 'complete';
      entry.container.classList.remove('is-streaming');
      if (opts.stopped) entry.container.classList.add('msg-stopped');
      if (opts.error) entry.container.classList.add('msg-error');
      resetExportButton(id);
      pendingCards.forEach(function(card) {
        card.classList.add('msg-card-inline');
        entry.container.appendChild(card);
      });
      pendingCards = [];
      if (pendingSources.length && !opts.error) {
        entry.container.appendChild(renderSources(pendingSources));
        pendingSources = [];
      }
      if (composeMode && !opts.error && entry.raw && !entry.insertButton) {
        var foot = elt('div', 'msg-foot');
        entry.insertButton = button('btn btn-small btn-ghost', 'Insert into email');
        entry.insertButton.addEventListener('click', function() {
          postToHost({ type: 'insert_text', payload: { text: entry.raw } });
        });
        foot.appendChild(entry.insertButton);
        entry.container.appendChild(foot);
      } else if (!composeMode && hasSingleSelection && !opts.error && entry.raw && !entry.replyFoot) {
        var rfoot = elt('div', 'msg-foot');
        var openReply = button('btn btn-small btn-primary', 'Open reply');
        var openReplyAll = button('btn btn-small btn-ghost', 'Reply all');
        var copyReply = button('btn btn-small btn-ghost', 'Copy');
        openReply.addEventListener('click', function() {
          postToHost({ type: 'draft_action', payload: { id: id, action: 'reply' } });
        });
        openReplyAll.addEventListener('click', function() {
          postToHost({ type: 'draft_action', payload: { id: id, action: 'reply_all' } });
        });
        copyReply.addEventListener('click', function() {
          postToHost({ type: 'copy', payload: { text: entry.raw } });
          copyReply.textContent = 'Copied';
          setTimeout(function() { copyReply.textContent = 'Copy'; }, 1500);
        });
        rfoot.appendChild(openReply);
        rfoot.appendChild(openReplyAll);
        rfoot.appendChild(copyReply);
        entry.replyFoot = rfoot;
        entry.status = elt('span', 'draft-note', '');
        rfoot.appendChild(entry.status);
        entry.foot = rfoot;
        entry.container.appendChild(rfoot);
      }
      scrollToBottom();
    },

    // Compact "working on it" status line. Replaces the previous verbose
    // expandable tool cards (the chat got too cluttered - user feedback).
    // Each tool call shows as a single italic line like
    //   ... Searching messages
    //   ... Reading message
    //   ... Listing folders
    // When the tool completes, the line is REMOVED entirely (we don't
    // need a permanent record - the model uses the result to produce
    // the actual assistant text, which is what the user reads).
    //
    // Friendly verbs per tool name. Anything not in the map falls back
    // to "Working on it..." so unknown future tools still render fine.
    appendToolCallCard: function(callId, name, argsJson) {
      var verb = ({
        outlook_get_current_compose_state: 'Reading compose context',
        outlook_get_current_selection:     'Reading current selection',
        outlook_list_folders:              'Listing folders',
        outlook_search_messages:           'Searching messages',
        outlook_read_message:              'Reading message',
        outlook_count_messages:            'Counting messages',
        outlook_list_recent_threads_with:  'Listing recent threads',
        outlook_create_draft:              'Creating draft',
        outlook_mark_as_read:              'Marking as read',
        outlook_flag_message:              'Flagging message',
        outlook_set_category:              'Setting category',
        jc_remember:                       'Preparing a note',
        outlook_list_calendar_events:      'Reading your calendar',
        outlook_find_meeting_times:        'Looking for free times',
        jc_propose_meeting:                'Preparing a meeting',
        outlook_resolve_people:            'Looking people up',
        outlook_find_rooms:                'Looking for a room',
        outlook_get_org_context:           'Reading org profile',
        jc_propose_rule:                   'Preparing a rule',
      })[name] || 'Working on it';

      var row = elt('div', 'tool-status');
      row.dataset.callId = callId;
      row.dataset.toolName = String(name || '');
      row.dataset.verb = verb;
      row.dataset.startedAt = String(Date.now());
      row.textContent = '\u2026 ' + verb;
      $messages.appendChild(row);

      // Live time counter. Updates the status line every second with the
      // elapsed seconds so the user can see the tool is making progress
      // rather than guessing whether Outlook has frozen. Counter math
      // uses Date.now() - startedAt so the displayed value is always
      // wall-clock-correct even if setInterval is throttled when the
      // WebView2 loses focus.
      var tick = function() {
        if (!row.parentNode) { return; }
        var ms = Date.now() - parseInt(row.dataset.startedAt, 10);
        var secs = Math.max(0, Math.floor(ms / 1000));
        if (secs > 0) {
          row.textContent = '\u2026 ' + verb + ' (' + secs + 's)';
        }
      };
      var tickId = setInterval(function() {
        if (!row.parentNode) { clearInterval(tickId); return; }
        tick();
      }, 1000);
      row.dataset.tickId = String(tickId);

      // Force an immediate refresh when the WebView2 regains focus.
      // WebView2 throttles setInterval while hidden; without this, the
      // counter appears to lag (or to "reset" perceptually) until the
      // next tick after focus returns.
      var visHandler = function() {
        if (document.visibilityState === 'visible') tick();
      };
      document.addEventListener('visibilitychange', visHandler);
      window.addEventListener('focus', visHandler);
      row.dataset.hasVisHandler = '1';
      row._oai_visHandler = visHandler;

      toolCards[callId] = row;
      scrollToBottom();
    },

    updateToolCallCard: function(callId, ok, summary, resultJson) {
      var row = toolCards[callId];
      if (!row) return;
      // Stop the live time counter.
      var tickId = parseInt(row.dataset.tickId || '0', 10);
      if (tickId) clearInterval(tickId);
      // Drop the visibility-change handler (one was attached per row).
      if (row._oai_visHandler) {
        try {
          document.removeEventListener('visibilitychange', row._oai_visHandler);
          window.removeEventListener('focus', row._oai_visHandler);
        } catch (e) { /* best-effort */ }
        row._oai_visHandler = null;
      }
      // On completion, drop the status line. Errors stick around as a
      // muted single-line error so the user sees that something failed
      // without the full JSON dump.
      if (ok) {
        collectSources(resultJson);
        renderToolResultIfHandled(callId, resultJson);
        if (row.parentNode) row.parentNode.removeChild(row);
      } else {
        var pdfExportError = row.dataset.toolName === 'outlook_export_pdf' ? normalizePdfExportToolError(resultJson) : null;
        if (pdfExportError && renderInlineErrorCard(null, pdfExportError)) {
          if (row.parentNode) row.parentNode.removeChild(row);
        } else {
          // The tool's own error text is for the model, not the user.
          row.classList.add('tool-status-err');
          row.textContent = '\u26A0 ' + (row.dataset.verb || 'A step') + ' did not work; Jameel Copilot carried on without it.';
        }
      }
      delete toolCards[callId];
      scrollToBottom();
    },

    appendAuditRow: function(text) {
      var row = elt('div', 'audit-row');
      var time = elt('span', 'audit-time', new Date().toLocaleTimeString());
      row.appendChild(time);
      row.appendChild(document.createTextNode(text));
      $messages.appendChild(row);
      scrollToBottom();
    },

    onFileSaved: function(messageId, fileInfo) {
      appendFileCardToMessage(messageId, fileInfo);
      resetExportButton(messageId);
    },

    onExportError: function(messageId, error) {
      if (!renderInlineErrorCard(messageId, error)) {
        resetExportButton(messageId);
        api.showError(exportErrorMessage(error));
      }
    },

    handleExportPdf: handleExportPdf,

    showError: function(message) {
      var node = elt('div', 'msg msg-assistant msg-error');
      var content = elt('div', 'msg-content');
      content.textContent = message || 'An error occurred.';
      node.appendChild(content);
      $messages.appendChild(node);
      scrollToBottom();
    },

    setComposerEnabled: function(enabled, isStopVisible) {
      $input.disabled = !enabled;
      $btnSend.disabled = !enabled;
      $btnSend.hidden = !!isStopVisible;
      $btnStop.hidden = !isStopVisible;
      $btnDraft.disabled = !enabled;
      $btnCopy.disabled = false; // always allow copy
      if (enabled) $input.focus();
    },

    clear: function() {
      // Stop any live tool-status time counters and unbind any
      // visibility-change handlers before tossing their DOM.
      for (var cid in toolCards) {
        try {
          var row = toolCards[cid];
          var tickId = row && parseInt(row.dataset && row.dataset.tickId || '0', 10);
          if (tickId) clearInterval(tickId);
          if (row && row._oai_visHandler) {
            try {
              document.removeEventListener('visibilitychange', row._oai_visHandler);
              window.removeEventListener('focus', row._oai_visHandler);
            } catch (e) { /* best-effort */ }
            row._oai_visHandler = null;
          }
        } catch (e) { /* best-effort */ }
      }
      $messages.innerHTML = '';
      assistantMessages = {};
      toolCards = {};
    },

    applyTheme: function(themeName) {
      var allowed = ['light', 'dark', 'high-contrast'];
      var theme = allowed.indexOf(themeName) >= 0 ? themeName : 'light';
      document.body.classList.remove('theme-light', 'theme-dark', 'theme-high-contrast');
      document.body.classList.add('theme-' + theme);
    },

    setContextStrip: function(ctx) {
      ctx = ctx || {};
      // Phase 3a: support two shapes of context.
      //   Inbox shape:    { folder, unread_count, total_count, mailbox_only?, selection? }
      //   Compose shape:  { subject, recipients, thread }
      // Disambiguate by checking for ctx.folder.
      if (ctx.folder !== undefined) {
        empty($ctxSubject);
        var folderLabel = ctx.folder || 'Mailbox';
        $ctxSubject.appendChild(document.createTextNode('In: ' + folderLabel + ' '));
        if (ctx.unread_count != null && ctx.unread_count > 0) {
          var unreadBtn = button('ctx-pill ctx-unread-btn', ctx.unread_count + ' unread');
          unreadBtn.title = 'Click to triage & mark unread emails as read / flagged';
          unreadBtn.addEventListener('click', function() {
            if ($input.disabled) return;
            postToHost({ type: 'chip_action', payload: { id: 'tidy' } });
          });
          $ctxSubject.appendChild(unreadBtn);
        } else if (ctx.unread_count === 0) {
          $ctxSubject.appendChild(elt('span', 'ctx-faint', '(0 unread)'));
        }

        var mailboxOnly = !!ctx.mailbox_only;
        var hasSel = !!(ctx.selection && ctx.selection.count > 0);
        hasSingleSelection = !mailboxOnly && !!(ctx.selection && ctx.selection.count === 1);

        if (hasSel) {
          var toggleBtn = button('ctx-pill ctx-scope-btn', mailboxOnly ? 'Back to email' : 'All mail / + Meeting');
          toggleBtn.title = mailboxOnly
            ? 'Switch back to the selected email'
            : 'Switch to whole-mailbox actions & schedule a new meeting';
          toggleBtn.addEventListener('click', function() {
            if ($input.disabled) return;
            postToHost({ type: 'set_scope', payload: { scope: mailboxOnly ? 'selection' : 'mailbox' } });
          });
          $ctxSubject.appendChild(toggleBtn);
        }

        empty($ctxRecipients);
        if (hasSel && !mailboxOnly) {
          var selText = ctx.selection.count === 1
            ? ('Selected: ' + (ctx.selection.subject || '') + (ctx.selection.from ? (' \u2014 ' + ctx.selection.from) : ''))
            : ('Selected: ' + ctx.selection.count + ' messages');
          $ctxRecipients.textContent = selText;
          if ($input) $input.placeholder = 'Ask about this email, or click "All mail / + Meeting" above...';
        } else if (hasSel && mailboxOnly) {
          $ctxRecipients.textContent = 'Scope: All mailbox & calendar (selected email ignored)';
          if ($input) $input.placeholder = 'Search all emails, or schedule a new meeting with anyone...';
        } else {
          $ctxRecipients.textContent = '';
          if ($input) $input.placeholder = 'Ask about your mailbox or schedule a meeting...';
        }
        $ctxThread.textContent = '';
        return;
      }
      // Compose shape (Phase 2 behaviour unchanged).
      $ctxSubject.textContent = ctx.subject || 'New email';
      var recipients = (ctx.recipients || []).join(', ');
      $ctxRecipients.textContent = recipients ? ('To: ' + recipients) : '';
      $ctxThread.textContent = ctx.thread || '';
    },

    /**
     * The chips above the input, in titled groups:
     *   [{ title, chips: [{ label, prompt, mode, id, primary }] }]
     * mode "send" (the default) sends the prompt as the user's message,
     * "fill" puts it in the input for the user to complete (a prompt with
     * a [placeholder]), "action" asks the host to run the chip (id).
     * An empty list leaves no row.
     */
    setQuickActions: function(groups) {
      if (!$quickActions) return;
      empty($quickActions);
      // A plain list of chips is one group without a title.
      if (groups && groups.length && !groups[0].chips) groups = [{ chips: groups }];
      (groups || []).forEach(function(group) {
        if (!group || !group.chips || !group.chips.length) return;
        var box = elt('div', 'qa-group');
        if (group.title) box.appendChild(elt('div', 'qa-title', group.title));
        var row = elt('div', 'qa-row');
        group.chips.forEach(function(chip) {
          var btn = button('qa-chip' + (chip.primary ? ' qa-primary' : ''), chip.label);
          if (chip.prompt) btn.title = chip.prompt;
          btn.addEventListener('click', function() {
            if ($input.disabled) return;
            if (chip.mode === 'action') {
              postToHost({ type: 'chip_action', payload: { id: chip.id } });
              return;
            }
            if (chip.mode === 'prompts') {
              postToHost({ type: 'show_prompts' });
              showView('prompts');
              return;
            }
            $input.value = chip.prompt || '';
            if (chip.mode === 'fill') {
              try {
                $input.focus();
                $input.setSelectionRange($input.value.length, $input.value.length);
              } catch (e) { /* best-effort */ }
            } else {
              sendInput();
            }
          });
          row.appendChild(btn);
        });
        box.appendChild(row);
        $quickActions.appendChild(box);
      });
      var first = $quickActions.querySelector('.qa-row');
      if (first && $quickActions.querySelectorAll('.qa-group').length > 1) {
        var more = button('qa-chip qa-more', 'More');
        more.addEventListener('click', function() {
          var open = $quickActions.classList.toggle('expanded');
          more.textContent = open ? 'Less' : 'More';
        });
        first.appendChild(more);
      }
      $quickActions.classList.add('compact');
      $quickActions.classList.remove('expanded');
    },

    // ---- One pane: saved chats, memory, drafts ------------------------

    // Replace what is shown with a saved conversation: [{role, text}].
    loadConversation: function(messages) {
      api.clear();
      (messages || []).forEach(function(m, i) {
        if (!m || typeof m.text !== 'string') return;
        if (m.role === 'user') {
          api.appendUserMessage(m.text);
        } else {
          var id = 'saved_' + i;
          api.appendAssistantMessage(id, m.text);
          api.finalizeAssistantMessage(id, {});
        }
      });
      showView('chat');
    },

    // The list behind the Chats button: [{id, title, thread, when}].
    setChats: function(list, currentId) {
      empty($chatList);
      if (!list || !list.length) {
        $chatList.appendChild(elt('div', 'panel-empty', 'No saved chats yet. Ask Jameel Copilot something and the chat is kept here.'));
        return;
      }
      list.forEach(function(chat) {
        var row = elt('div', 'chat-row' + (chat.id === currentId ? ' current' : ''));
        var open = button('chat-open', '');
        var title = elt('span', 'chat-title', chat.title || 'Chat');
        title.dir = 'auto';
        open.appendChild(title);
        open.appendChild(elt('span', 'chat-sub', (chat.thread ? 'Email thread' : 'Mailbox') + ' - ' + (chat.when || '')));
        open.addEventListener('click', function() {
          postToHost({ type: 'open_chat', payload: { id: chat.id } });
        });
        var del = button('chat-del', 'Delete');
        del.title = 'Delete this chat';
        confirmThen(del, 'Delete?', function() {
          postToHost({ type: 'delete_chat', payload: { id: chat.id } });
        });
        row.appendChild(open);
        row.appendChild(del);
        $chatList.appendChild(row);
      });
    },

    // The notes behind the Memory button: [{id, text}].
    setMemory: function(notes) {
      renderNotes($memoryList, notes, 'memory', 'Nothing yet. Add what Jameel Copilot should know about you, or how you like your replies.');
    },

    // The user's own prompts, behind the last chip: [{id, text}].
    setPrompts: function(notes) {
      renderNotes($promptList, notes, 'prompt', 'No prompts of your own yet. Write a request you use often and it becomes a button.');
    },

    // After "Create rule": what happened to the rule of that card.
    markRule: function(callId, status, done) {
      var card = ruleCards[callId];
      if (!card) return;
      card.status.textContent = status || '';
      empty(card.foot);
      if (done) card.node.classList.add('draft-done');
    },

    // "Action items": what one email asks of the user.
    // items: [{text, due, id, from, subject}]
    showTodos: function(id, items) {
      var entry = assistantMessages[id];
      if (!entry) return;
      empty(entry.content);
      entry.raw = '';
      entry.container.classList.add('msg-day');
      entry.todoRows = [];
      if (!items || !items.length) {
        entry.content.appendChild(elt('div', 'review-ok', 'This email asks nothing of you.'));
      } else {
        entry.content.appendChild(elt('div', 'day-section', 'Action items (' + items.length + ')'));
        items.forEach(function(item, i) {
          entry.content.appendChild(todoRow(id, i, item, entry.todoRows));
        });
      }
      api.finalizeAssistantMessage(id, {});
    },

    // After "Add to tasks": the row says where it went, or lets the user try again.
    markTodo: function(id, index, status, done) {
      var entry = assistantMessages[id];
      var row = entry && entry.todoRows && entry.todoRows[index];
      if (!row) return;
      if (done) {
        row.row.replaceChild(elt('span', 'todo-done', status || ''), row.add);
      } else {
        row.add.disabled = false;
        row.add.textContent = 'Try again';
        row.add.title = status || '';
      }
    },

    // "Start my day": what to do, what to know, and today's meetings.
    // data: {date, todos: [as showTodos], topics: [{title, summary, emails: [{id, from, subject, when}]}],
    //        meetings: [{subject, start, end, all_day, location}], or null when the calendar could not be read}
    showDay: function(id, data) {
      var entry = assistantMessages[id];
      if (!entry) return;
      data = data || {};
      empty(entry.content);
      entry.raw = '';
      entry.container.classList.add('msg-day');
      entry.todoRows = [];
      var day = new Date((data.date || '') + 'T00:00');
      entry.content.appendChild(elt('div', 'day-title',
        'Your day' + (isNaN(day.getTime()) ? '' : ' - ' + DAYS[day.getDay()] + ' ' + day.getDate() + ' ' + MONTHS[day.getMonth()])));

      var todos = data.todos || [];
      entry.content.appendChild(elt('div', 'day-section', 'To do (' + todos.length + ')'));
      if (!todos.length) entry.content.appendChild(elt('div', 'day-empty', 'Nothing in your recent mail asks something of you.'));
      todos.forEach(function(item, i) {
        entry.content.appendChild(todoRow(id, i, item, entry.todoRows));
      });

      var topics = data.topics || [];
      if (topics.length) {
        entry.content.appendChild(elt('div', 'day-section', 'Catch up (' + topics.length + ')'));
        topics.forEach(function(topic) {
          var box = elt('div', 'topic');
          var title = elt('div', 'topic-title', topic.title || '');
          title.dir = 'auto';
          box.appendChild(title);
          if (topic.summary) {
            var summary = elt('div', 'topic-summary', topic.summary);
            summary.dir = 'auto';
            box.appendChild(summary);
          }
          (topic.emails || []).forEach(function(m) { box.appendChild(sourceRow(m)); });
          entry.content.appendChild(box);
        });
      }

      var meetings = data.meetings;
      entry.content.appendChild(elt('div', 'day-section', "Today's meetings" + (meetings ? ' (' + meetings.length + ')' : '')));
      if (!meetings) {
        entry.content.appendChild(elt('div', 'day-empty', 'Your calendar could not be read.'));
      } else if (!meetings.length) {
        entry.content.appendChild(elt('div', 'day-empty', 'No meetings today.'));
      } else {
        meetings.forEach(function(m) {
          var row = elt('div', 'meet-row');
          row.appendChild(elt('span', 'meet-time', m.all_day ? 'All day' : hm(new Date(m.start)) + ' - ' + hm(new Date(m.end))));
          var what = elt('span', 'meet-what', m.subject || '(no subject)');
          what.dir = 'auto';
          row.appendChild(what);
          if (m.location) {
            var where = elt('span', 'meet-where', m.location);
            where.dir = 'auto';
            row.appendChild(where);
          }
          entry.content.appendChild(row);
        });
      }
      api.finalizeAssistantMessage(id, {});
    },

    // A draft card: text the user decides about. It streams like an
    // answer (appendTextDelta) and ends with the buttons the host names.
    beginDraft: function(id, title) {
      var node = elt('div', 'msg msg-assistant msg-draft is-streaming');
      node.dataset.messageId = id;
      var head = elt('div', 'draft-head');
      head.appendChild(elt('span', 'draft-title', title || 'Draft'));
      var status = elt('span', 'draft-status', 'writing...');
      head.appendChild(status);
      var content = elt('div', 'msg-content');
      content.dir = 'auto';
      var foot = elt('div', 'draft-foot');
      node.appendChild(head);
      node.appendChild(content);
      node.appendChild(foot);
      $messages.appendChild(node);
      assistantMessages[id] = {
        container: node, content: content, raw: '', complete: false,
        exportButton: null, status: status, foot: foot
      };
      scrollToBottom();
    },

    // buttons: [{label, action, primary}]; opts: {status, note, error}.
    finalizeDraft: function(id, buttons, opts) {
      var entry = assistantMessages[id];
      if (!entry || !entry.foot) return;
      opts = opts || {};
      entry.complete = true;
      entry.container.classList.remove('is-streaming');
      entry.status.textContent = opts.status || '';
      if (opts.error) entry.container.classList.add('msg-error');
      empty(entry.foot);
      (buttons || []).forEach(function(b) {
        var btn = button('btn btn-small' + (b.primary ? ' btn-primary' : (b.action === 'discard' ? ' btn-ghost' : '')), b.label);
        btn.dataset.action = b.action;
        btn.addEventListener('click', function() {
          if (b.action === 'discard') {
            if (entry.container.parentNode) entry.container.parentNode.removeChild(entry.container);
            delete assistantMessages[id];
          }
          postToHost({ type: 'draft_action', payload: { id: id, action: b.action } });
        });
        entry.foot.appendChild(btn);
      });
      if (opts.note) entry.container.appendChild(elt('div', 'draft-note', opts.note));
      scrollToBottom();
    },

    // After a button did its work: say where the draft went.
    markDraft: function(id, status, done) {
      var entry = assistantMessages[id];
      if (!entry || !entry.status) return;
      entry.status.textContent = status || '';
      if (!done) return;
      entry.container.classList.add('draft-done');
      // A used card cannot put its text in a second time: only a way to put the card away stays.
      [].slice.call(entry.foot.children).forEach(function(b) {
        if (b.dataset.action === 'discard') {
          b.textContent = 'Dismiss';
        } else {
          entry.foot.removeChild(b);
        }
      });
    },

    // Replaces what an answer shows with plain text (a status, or why nothing came).
    replaceAssistantText: function(id, text) {
      var entry = assistantMessages[id];
      if (!entry) return;
      entry.raw = text || '';
      entry.content.textContent = entry.raw;
    },

    // "Prioritize my inbox": the answer becomes the ranked list.
    // groups: [{level, label, items: [{id, from, subject, when, reason}]}]
    showPriorities: function(id, groups) {
      var entry = assistantMessages[id];
      if (!entry) return;
      empty(entry.content);
      entry.raw = '';
      entry.container.classList.add('msg-priorities');
      (groups || []).forEach(function(group) {
        if (!group || !group.items || !group.items.length) return;
        var section = elt('div', 'prio-group prio-' + group.level);
        var head = button('prio-head', (group.label || group.level) + ' (' + group.items.length + ')');
        var list = elt('div', 'prio-list');
        if (group.level === 'low') list.hidden = true;
        head.addEventListener('click', function() { list.hidden = !list.hidden; });
        section.appendChild(head);
        group.items.forEach(function(item) {
          var row = elt('div', 'prio-row');
          var open = button('prio-open', '');
          var top = elt('span', 'prio-top');
          var from = elt('span', 'source-from', item.from || '(no sender)');
          var subject = elt('span', 'source-subject', item.subject || '(no subject)');
          var when = elt('span', 'source-when', shortDate(item.when));
          from.dir = 'auto';
          subject.dir = 'auto';
          top.appendChild(from);
          top.appendChild(subject);
          top.appendChild(when);
          open.appendChild(top);
          if (item.reason) {
            var reason = elt('span', 'prio-reason', item.reason);
            reason.dir = 'auto';
            open.appendChild(reason);
          }
          open.title = 'Open this email';
          open.addEventListener('click', function() {
            postToHost({ type: 'open_message', payload: { id: item.id } });
          });
          row.appendChild(open);
          if (group.level !== 'low') {
            // "Waiting for a reply" asks for a follow-up to the user's own email instead of a reply.
            var followUp = !!(group.button && group.button.type === 'followup');
            var reply = button('btn btn-small btn-ghost prio-reply', followUp ? 'Follow up' : 'Reply');
            reply.title = followUp ? 'Write a follow-up to this email' : 'Write a reply to this email';
            reply.addEventListener('click', function() {
              if ($input.disabled) return;
              postToHost({ type: followUp ? 'followup' : 'priority_reply', payload: { id: item.id } });
            });
            row.appendChild(reply);
          }
          list.appendChild(row);
        });
        section.appendChild(list);
        entry.content.appendChild(section);
      });
      api.finalizeAssistantMessage(id, {});
    },

    // "Review": the summary, then points the user ticks; Apply writes a draft with those changes.
    showReview: function(id, summary, points) {
      var entry = assistantMessages[id];
      if (!entry) return;
      empty(entry.content);
      entry.raw = '';
      entry.container.classList.add('msg-review');
      var head = elt('div', 'review-summary', summary || '');
      head.dir = 'auto';
      entry.content.appendChild(head);
      if (!points || !points.length) {
        entry.content.appendChild(elt('div', 'review-ok', 'Nothing to change. It is ready to send.'));
        api.finalizeAssistantMessage(id, {});
        return;
      }
      var boxes = [];
      points.forEach(function(p) {
        var row = elt('label', 'review-point');
        var box = document.createElement('input');
        box.type = 'checkbox';
        box.checked = true;
        box.value = p.id;
        boxes.push(box);
        var text = elt('span', 'review-text');
        text.dir = 'auto';
        if (p.issue) text.appendChild(elt('span', 'review-issue', p.issue));
        text.appendChild(elt('span', 'review-suggestion', p.suggestion || ''));
        row.appendChild(box);
        row.appendChild(text);
        entry.content.appendChild(row);
      });
      var foot = elt('div', 'draft-foot');
      var apply = button('btn btn-small btn-primary', 'Apply selected');
      var dismiss = button('btn btn-small btn-ghost', 'Dismiss');
      apply.addEventListener('click', function() {
        if ($input.disabled) return;
        var picks = boxes.filter(function(b) { return b.checked; }).map(function(b) { return b.value; });
        if (!picks.length) return;
        postToHost({ type: 'review_apply', payload: { id: id, picks: picks } });
      });
      dismiss.addEventListener('click', function() {
        if (entry.container.parentNode) entry.container.parentNode.removeChild(entry.container);
        delete assistantMessages[id];
      });
      foot.appendChild(apply);
      foot.appendChild(dismiss);
      entry.content.appendChild(foot);
      api.finalizeAssistantMessage(id, {});
    },

    // "Tidy my inbox": proposed changes the user ticks; nothing changes until Apply.
    // groups: [{action, label, items: [{id, from, subject, when, reason}]}]
    showTidy: function(id, groups) {
      var entry = assistantMessages[id];
      if (!entry) return;
      empty(entry.content);
      entry.raw = '';
      entry.container.classList.add('msg-review', 'msg-tidy');
      var boxes = [];
      (groups || []).forEach(function(group) {
        if (!group || !group.items || !group.items.length) return;
        entry.content.appendChild(elt('div', 'prio-head', (group.label || group.action) + ' (' + group.items.length + ')'));
        group.items.forEach(function(item) {
          var row = elt('div', 'review-point');
          var box = document.createElement('input');
          box.type = 'checkbox';
          box.checked = true;
          box.value = item.id;
          boxes.push(box);
          var text = elt('span', 'review-text');
          var top = elt('span', 'prio-top');
          var from = elt('span', 'source-from', item.from || '(no sender)');
          var subject = elt('span', 'source-subject', item.subject || '(no subject)');
          from.dir = 'auto';
          subject.dir = 'auto';
          top.appendChild(from);
          top.appendChild(subject);
          text.appendChild(top);
          if (item.reason) {
            var reason = elt('span', 'review-issue', item.reason);
            reason.dir = 'auto';
            text.appendChild(reason);
          }
          var open = button('btn btn-small btn-ghost', 'Open');
          open.addEventListener('click', function() {
            postToHost({ type: 'open_message', payload: { id: item.id } });
          });
          row.appendChild(box);
          row.appendChild(text);
          row.appendChild(open);
          entry.content.appendChild(row);
        });
      });
      if (!boxes.length) {
        entry.content.appendChild(elt('div', 'review-ok', 'Nothing to tidy.'));
        api.finalizeAssistantMessage(id, {});
        return;
      }
      var foot = elt('div', 'draft-foot');
      var status = elt('span', 'draft-note', 'Nothing has changed yet.');
      var apply = button('btn btn-small btn-primary', 'Apply selected');
      apply.addEventListener('click', function() {
        if ($input.disabled) return;
        var picks = boxes.filter(function(b) { return b.checked; }).map(function(b) { return b.value; });
        if (!picks.length) return;
        boxes.forEach(function(b) { b.disabled = true; });
        apply.disabled = true;
        postToHost({ type: 'tidy_apply', payload: { id: id, picks: picks } });
      });
      foot.appendChild(apply);
      foot.appendChild(status);
      entry.content.appendChild(foot);
      entry.tidyFoot = foot;
      entry.tidyStatus = status;
      entry.tidyApply = apply;
      api.finalizeAssistantMessage(id, {});
    },

    // After Apply (undoable) or Undo (done): the status, and Undo while it can still undo.
    markTidy: function(id, status, undoable) {
      var entry = assistantMessages[id];
      if (!entry || !entry.tidyFoot) return;
      entry.tidyStatus.textContent = status || '';
      if (entry.tidyApply.parentNode) entry.tidyFoot.removeChild(entry.tidyApply);
      if (entry.tidyUndo && entry.tidyUndo.parentNode) entry.tidyFoot.removeChild(entry.tidyUndo);
      entry.tidyUndo = null;
      if (undoable) {
        var undo = button('btn btn-small', 'Undo');
        undo.addEventListener('click', function() {
          undo.disabled = true;
          postToHost({ type: 'tidy_undo', payload: { id: id } });
        });
        entry.tidyFoot.insertBefore(undo, entry.tidyStatus);
        entry.tidyUndo = undo;
      }
    },

    // In a compose window the input can also write the email (Draft).
    setComposeMode: function(on) {
      composeMode = !!on;
      $btnDraft.hidden = !on;
      // Writing is what a compose window is for: Draft leads, the chat is "Ask".
      $btnSend.textContent = on ? 'Ask' : 'Send';
      $btnSend.classList.toggle('btn-primary', !on);
      $input.placeholder = on
        ? 'Tell Jameel Copilot what to write, or ask about this thread...'
        : 'Ask about this email or your mailbox...';
    },

    // The response style in use. The host hides the picker (null) when
    // the administrator has switched temperature off.
    setStyle: function(style) {
      if (!style) {
        $stylePicker.hidden = true;
        return;
      }
      $styleSelect.value = style;
      $stylePicker.hidden = false;
    },

    // Used by C# during dev to verify the bridge is live before
    // pushing real messages. Returns a string from JS that C# can
    // assert on with await ExecuteScriptAsync("jameelcopilot.ping()").
    ping: function() { return 'pong'; }
  };

  window.jameelcopilot = window.jameelcopilot || {};
  for (var apiName in api) {
    if (Object.prototype.hasOwnProperty.call(api, apiName)) {
      window.jameelcopilot[apiName] = api[apiName];
    }
  }

  // -- Composer event wiring --------------------------------------
  function sendInput() {
    // Guard against double-send when a previous turn is still running.
    // The host disables $input + $btnSend when a turn is in flight, but
    // a fast Enter or a race between WebMessageReceived and the C# side
    // re-enabling the composer can still let a second submission through
    // and produce two concurrent tool-status rows.
    if ($input.disabled || $btnSend.disabled) return;
    var text = $input.value.trim();
    if (!text) return;
    $input.value = '';
    postToHost({
      type: 'send',
      payload: {
        text: text
      }
    });
  }

  $btnSend.addEventListener('click', sendInput);
  $btnStop.addEventListener('click', function() {
    postToHost({ type: 'stop' });
  });
  $btnDraft.addEventListener('click', function() {
    if ($input.disabled || $btnDraft.disabled) return;
    var text = $input.value.trim();
    if (!text) {
      $input.focus();
      return;
    }
    $input.value = '';
    postToHost({ type: 'draft', payload: { text: text } });
  });
  $btnNewChat.addEventListener('click', function() {
    if ($input.disabled) return;   // a turn is running
    postToHost({ type: 'new_chat' });
    showView('chat');
  });
  $btnChats.addEventListener('click', function() {
    if (currentView === 'chats') {
      showView('chat');
      return;
    }
    postToHost({ type: 'show_chats' });
    showView('chats');
  });
  $btnMemory.addEventListener('click', function() {
    if (currentView === 'memory') {
      showView('chat');
      return;
    }
    postToHost({ type: 'show_memory' });
    showView('memory');
  });
  document.getElementById('btnChatsBack').addEventListener('click', function() { showView('chat'); });
  document.getElementById('btnMemoryBack').addEventListener('click', function() { showView('chat'); });
  document.getElementById('btnPromptsBack').addEventListener('click', function() { showView('chat'); });
  document.getElementById('btnPromptAdd').addEventListener('click', function() {
    var text = $promptInput.value.trim();
    if (!text) {
      $promptInput.focus();
      return;
    }
    $promptInput.value = '';
    postToHost({ type: 'prompt_add', payload: { text: text } });
  });
  confirmThen(document.getElementById('btnClearChats'), 'Delete all chats?', function() {
    postToHost({ type: 'clear_chats' });
  });
  document.getElementById('btnMemoryAdd').addEventListener('click', function() {
    var text = $memoryInput.value.trim();
    if (!text) {
      $memoryInput.focus();
      return;
    }
    $memoryInput.value = '';
    postToHost({ type: 'memory_add', payload: { text: text } });
  });
  $btnCopy.addEventListener('click', function() {
    postToHost({ type: 'copy' });
  });
  $styleSelect.addEventListener('change', function() {
    postToHost({ type: 'set_style', payload: { style: $styleSelect.value } });
  });

  // Enter sends, Shift+Enter inserts a newline (standard chat UX).
  $input.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendInput();
    }
  });

  // Collapse secondary chip groups whenever a message is sent so the chat stays clean.
  if (window.MutationObserver) {
    new MutationObserver(function() {
      $quickActions.classList.add('compact');
      if ($messages.childElementCount > 0) $quickActions.classList.remove('expanded');
    }).observe($messages, { childList: true });
  }

  // Tell the host we're ready so it can push the initial context strip
  // + apply the theme before the first user message.
  postToHost({ type: 'ready' });
})();
