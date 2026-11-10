'use strict';

const $ = id => document.getElementById(id);
const token = document.querySelector('meta[name=local-token]').content;
const editor = $('body');
let data = {Projects: [], Sections: [], Events: []};
let projectId, sectionId, eventId, draft, nameRequest, selection;
let activeTool, busy = false, month = new Date(), noticeTimer;
const restartMessage = '程式已更新，請先在原 PowerShell 視窗按 Ctrl+C，重新執行 LocalDataCenter.ps1，再開啟新網址。目前服務仍是舊版，暫停寫入以免狀態或時程漏存。';

function eventStatus(event) { return event?.status === 'completed' ? 'completed' : 'active'; }
function eventSchedules(event) {
    if (Array.isArray(event?.schedules)) return event.schedules;
    return event?.start && event.end ? [{name:'原有時程', start:event.start, end:event.end}] : [];
}

function element(tag, content, className) {
    const item = document.createElement(tag);
    if (content !== undefined) item.textContent = content;
    if (className) item.className = className;
    return item;
}

async function api(path, options = {}) {
    if (options.method === 'POST' && (data.apiVersion !== 3 || !data.capabilities?.includes('attachment-cleanup'))) throw Error(restartMessage);
    const result = await fetch(path, options);
    const value = await result.json();
    if (!result.ok) throw Error(value.error);
    if (value.warnings?.length) alert(value.warnings.join('\n'));
    return value;
}

function save(payload) {
    return api('/api/save', {method:'POST', headers:{'X-Local-Token':token, 'Content-Type':'application/json'}, body:JSON.stringify(payload)});
}

function notice(message) {
    $('notice').textContent = message;
    clearTimeout(noticeTimer);
    noticeTimer = setTimeout(() => $('notice').textContent = '', 6500);
}

const safeImage = /^\/attachment\/[a-f0-9]{32}\.(png|jpg|gif|webp)$/;
const safeFile = /^[a-f0-9]{32}\.(png|jpg|gif|webp|bin)$/;
const formats = new Set('B STRONG I EM U S BR P DIV SPAN FONT BLOCKQUOTE UL OL LI TABLE THEAD TBODY TFOOT TR TH TD IMG'.split(' '));

function cleanHTML(html) {
    const source = document.createElement('template');
    source.innerHTML = html;
    const destination = document.createElement('div');
    function copy(node, parent) {
        if (node.nodeType === Node.TEXT_NODE) {
            parent.append(document.createTextNode(node.textContent));
            return;
        }
        if (node.nodeType !== Node.ELEMENT_NODE || /^(SCRIPT|STYLE|IFRAME|SVG|MATH|OBJECT|TEMPLATE)$/.test(node.tagName)) return;
        if (!formats.has(node.tagName)) {
            for (const child of node.childNodes) copy(child, parent);
            return;
        }
        const next = document.createElement(node.tagName);
        if (node.tagName === 'IMG') {
            const src = node.getAttribute('src');
            if (!safeImage.test(src)) return;
            next.src = src;
            next.alt = node.alt;
        }
        if (node.tagName === 'FONT') {
            if (/^[1-7]$/.test(node.size)) next.size = node.size;
            if (/^#[a-f0-9]{6}$/i.test(node.color)) next.color = node.color;
        }
        if (node.style.color && /^(#[\da-f]{3,8}|rgba?\([\d., %]+\)|[a-z]+)$/i.test(node.style.color)) next.style.color = node.style.color;
        if (/^(TD|TH)$/.test(node.tagName)) {
            for (const attr of ['rowspan','colspan']) {
                const span = Number(node.getAttribute(attr));
                if (Number.isInteger(span) && span > 0 && span <= 1000) next.setAttribute(attr, String(span));
            }
        }
        if (/^(#[\da-f]{3,8}|rgba?\([\d., %]+\)|[a-z]+)$/i.test(node.style.backgroundColor)) next.style.backgroundColor = node.style.backgroundColor;
        if (/^(left|right|center|justify)$/.test(node.style.textAlign || node.getAttribute('align'))) next.style.textAlign = node.style.textAlign || node.getAttribute('align');
        if (/^(top|middle|bottom)$/.test(node.style.verticalAlign)) next.style.verticalAlign = node.style.verticalAlign;
        if (/^(\d{1,2}px|xx-small|x-small|small|medium|large|x-large|xx-large)$/.test(node.style.fontSize)) next.style.fontSize = node.style.fontSize;
        if (/^(bold|[1-9]00)$/.test(node.style.fontWeight)) next.style.fontWeight = node.style.fontWeight;
        for (const child of node.childNodes) copy(child, next);
        parent.append(next);
    }
    for (const node of source.content.childNodes) copy(node, destination);
    return destination.innerHTML;
}

function drawFiles(target, files, editable = false, onRemove) {
    target.className = 'files';
    target.replaceChildren();
    for (const file of files || []) {
        if (!safeFile.test(file.id)) continue;
        const item = element('span', undefined, 'file');
        const link = element('button', undefined, 'file-preview');
        link.type = 'button';
        const icon = element('span', /\.pdf$/i.test(file.name) ? '▤' : /\.(png|jpe?g|gif|webp)$/i.test(file.name) ? '▧' : '▱', 'file-icon');
        icon.setAttribute('aria-hidden', 'true');
        link.append(icon, document.createTextNode(file.name));
        link.title = '預覽 ' + file.name;
        link.onclick = () => openPreview(file);
        item.append(link);
        if (editable) {
            const remove = element('button', '移除');
            remove.type = 'button';
            remove.onclick = () => {
                if (onRemove) { onRemove(file); return; }
                for (const image of editor.querySelectorAll('img')) {
                    if (image.getAttribute('src') === '/attachment/' + file.id) image.remove();
                }
                draft.files = draft.files.filter(entry => entry.id !== file.id);
                drawFiles(target, draft.files, true);
            };
            item.append(' ', remove);
        }
        target.append(item);
    }
}

let previewSession;

function disposePreview() {
    const session = previewSession;
    previewSession = undefined;
    if (!session) return;
    session.controller.abort();
    session.observer?.disconnect();
    session.resizeObserver?.disconnect();
    for (const task of session.tasks || []) task.cancel();
    for (const layer of session.layers || []) layer.cancel();
    if (session.renderTask) session.renderTask.cancel();
    if (session.loadingTask) session.loadingTask.destroy().catch(() => {});
    if (session.url) URL.revokeObjectURL(session.url);
    for (const image of $('preview-content').querySelectorAll('img')) image.removeAttribute('src');
    for (const canvas of $('preview-content').querySelectorAll('canvas')) canvas.width = canvas.height = 0;
    $('preview-content').replaceChildren();
    $('preview-pages').hidden = true;
    $('preview-status').textContent = '';
    $('preview-download').removeAttribute('href');
}

async function renderPreviewPage(session, number = session.page) {
    if (previewSession !== session || session.rendered.has(number)) return;
    session.rendered.add(number);
    const wrapper = session.pages[number - 1];
    try {
        const page = await session.pdf.getPage(number);
        if (previewSession !== session) return;
        const original = page.getViewport({scale:1});
        const viewport = page.getViewport({scale:Math.min(2, 1800 / Math.max(original.width, original.height))});
        const canvas = element('canvas');
        canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
        canvas.setAttribute('aria-label', 'PDF 第 ' + number + ' 頁');
        wrapper.style.aspectRatio = viewport.width + ' / ' + viewport.height;
        wrapper.style.setProperty('--scale-factor', viewport.scale);
        wrapper.style.width = viewport.width + 'px';
        wrapper.replaceChildren(canvas);
        const task = page.render({canvasContext:canvas.getContext('2d'), viewport});
        session.tasks.add(task);
        await task.promise;
        session.tasks.delete(task);
        if (previewSession !== session) return;
        const text = element('div', undefined, 'textLayer');
        text.style.width = viewport.width + 'px'; text.style.height = viewport.height + 'px';
        wrapper.append(text);
        wrapper.dataset.renderWidth = viewport.width;
        text.style.transform = 'scale(' + (wrapper.clientWidth / viewport.width) + ')';
        const layer = new session.pdfjs.TextLayer({textContentSource:await page.getTextContent(), container:text, viewport});
        session.layers.add(layer);
        await layer.render();
        session.layers.delete(layer);
        $('preview-status').textContent = '';
        page.cleanup();
    } catch (error) {
        session.rendered.delete(number);
        if (previewSession === session) $('preview-status').textContent = '第 ' + number + ' 頁預覽失敗，可下載原檔查看。';
    }
}

async function openPreview(file) {
    disposePreview();
    const session = {controller:new AbortController(), page:1, tasks:new Set(), layers:new Set(), rendered:new Set()};
    previewSession = session;
    const url = file.url || '/attachment/' + file.id;
    $('preview-title').textContent = file.name;
    $('preview-download').href = url;
    $('preview-download').download = file.name;
    $('preview-status').textContent = '正在載入預覽…';
    if (!$('preview').open) $('preview').showModal();
    const imageType = {png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',webp:'image/webp'};
    const ext = file.name.split('.').pop().toLowerCase();
    const type = Object.prototype.hasOwnProperty.call(imageType, ext) ? imageType[ext] : '';
    const textTypes = new Set(['txt','csv','log','md','json','xml','yaml','yml','ini']);
    if (!type && !['pdf','xls','xlsx'].includes(ext) && !textTypes.has(ext)) {
        $('preview-status').textContent = '此格式不支援預覽，請下載原檔查看。';
        return;
    }
    try {
        const response = await fetch(url, {signal:session.controller.signal});
        if (!response.ok) throw Error('無法讀取附件');
        const bytes = await response.arrayBuffer();
        if (previewSession !== session) return;
        if (type) {
            session.url = URL.createObjectURL(new Blob([bytes], {type}));
            const image = element('img'); image.alt = file.name;
            image.onload = () => { if (previewSession === session) $('preview-status').textContent = ''; };
            image.onerror = () => { if (previewSession === session) $('preview-status').textContent = '圖片無法預覽，請下載原檔查看。'; };
            image.src = session.url;
            $('preview-content').append(image);
        } else if (ext === 'pdf') {
            const pdfjs = session.pdfjs = await import('/vendor/pdf.mjs');
            if (previewSession !== session) return;
            pdfjs.GlobalWorkerOptions.workerSrc = '/vendor/pdf.worker.mjs';
            session.loadingTask = pdfjs.getDocument({data:bytes, isEvalSupported:false,
                cMapUrl:'/vendor/cmaps/', cMapPacked:true, standardFontDataUrl:'/vendor/standard_fonts/'});
            session.pdf = await session.loadingTask.promise;
            if (previewSession !== session) return;
            $('preview-pages').hidden = false;
            session.pages = [];
            session.resizeObserver = new ResizeObserver(entries => {
                for (const entry of entries) {
                    const layer = entry.target.querySelector('.textLayer');
                    if (layer) layer.style.transform = 'scale(' + (entry.target.clientWidth / Number(entry.target.dataset.renderWidth)) + ')';
                }
            });
            const first = await session.pdf.getPage(1);
            const size = first.getViewport({scale:1});
            for (let n = 1; n <= session.pdf.numPages; n++) {
                const wrapper = element('div', undefined, 'pdf-page');
                wrapper.dataset.page = n;
                wrapper.style.aspectRatio = size.width + ' / ' + size.height;
                wrapper.setAttribute('aria-label', '第 ' + n + ' 頁');
                session.pages.push(wrapper);
                $('preview-content').append(wrapper);
                session.resizeObserver.observe(wrapper);
            }
            $('preview-content').scrollTop = 0;
            updatePreviewPage(session, 1);
            session.observer = new IntersectionObserver(entries => {
                for (const entry of entries) if (entry.isIntersecting) renderPreviewPage(session, Number(entry.target.dataset.page));
            }, {root:$('preview-content'), rootMargin:'400px'});
            session.pages.forEach(page => session.observer.observe(page));
            await renderPreviewPage(session);
        } else if (ext === 'xls' || ext === 'xlsx') {
            const signature = new Uint8Array(bytes);
            if (ext === 'xlsx' ? !(signature[0] === 0x50 && signature[1] === 0x4b) : !((signature[0] === 0xd0 && signature[1] === 0xcf) || signature[0] === 0x09)) throw Error('檔案內容不是有效的 Excel 格式。');
            renderWorkbookPreview(bytes);
            $('preview-status').textContent = '';
        } else {
            const text = new TextDecoder('utf-8', {fatal:true}).decode(bytes);
            $('preview-content').append(element('pre', text));
            $('preview-status').textContent = '';
        }
    } catch (error) {
        if (previewSession === session) $('preview-status').textContent = '無法預覽此檔案：' + (error.message || String(error)) + '。可下載原檔查看（純文字需為 UTF-8 編碼）。';
    }
}
$('preview').addEventListener('close', () => { if (!$('preview').open) disposePreview(); });
let previewPointerOutside = false;
function outsidePreview(event) {
    const rect = $('preview').getBoundingClientRect();
    return event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom;
}
$('preview').addEventListener('pointerdown', event => { previewPointerOutside = event.target === $('preview') && outsidePreview(event); });
$('preview').addEventListener('click', event => {
    if (previewPointerOutside && event.target === $('preview') && outsidePreview(event)) $('preview').close();
    previewPointerOutside = false;
});
for (const [id, offset] of [['preview-prev', -1], ['preview-next', 1]]) $(id).onclick = () => {
    const session = previewSession;
    if (!session?.pdf) return;
    session.page = Math.max(1, Math.min(session.pdf.numPages, session.page + offset));
    session.pages[session.page - 1].scrollIntoView({block:'start'});
    updatePreviewPage(session, session.page);
    renderPreviewPage(session);
};

// Calendar-day indices avoid DST changing the number of days between dates.
function dayNumber(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return NaN;
    const date = new Date(value + 'T00:00:00Z');
    return Number.isFinite(+date) && date.toISOString().slice(0, 10) === value ? +date / 86400000 : NaN;
}

function scheduleState(schedule, now = new Date()) {
    const start = dayNumber(schedule.start), end = dayNumber(schedule.end);
    const today = dayNumber(localDate(now));
    const midnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    const currentDay = today + (now - midnight) / (tomorrow - midnight);
    const remaining = end - today;
    const completed = schedule.completed === true && Number.isFinite(dayNumber(schedule.completedDate));
    const completedParts = (schedule.completedDate || '').split('-');
    return {
        start, end, remaining, completed,
        progress: completed ? 1 : Math.max(0, Math.min(1, (currentDay - start) / (end + 1 - start))),
        tone: completed ? 'normal' : remaining < 5 ? 'danger' : remaining < 10 ? 'warning' : 'normal',
        label: completed ? Number(completedParts[1]) + '月' + Number(completedParts[2]) + '日已完成' : remaining < 0 ? '已逾期 ' + -remaining + ' 日' : remaining === 0 ? '今日到期' :
            (today < start ? '尚未開始・' : '') + '剩餘 ' + remaining + ' 日'
    };
}

let scheduleDraft;

function drawScheduleImages(target, files) {
    target.replaceChildren();
    for (const file of files) {
        const url = '/attachment/' + file.id;
        if (!safeImage.test(url)) continue;
        const button = element('button', undefined, 'schedule-image-preview'); button.type = 'button';
        button.title = '預覽 ' + file.name;
        const image = element('img'); image.src = url; image.alt = file.name; image.loading = 'lazy';
        button.append(image); button.onclick = () => openPreview(file); target.append(button);
    }
}

function drawScheduleDraftFiles() {
    drawScheduleImages($('schedule-edit-images'), scheduleDraft.files);
    drawFiles($('schedule-edit-files'), scheduleDraft.files, true, file => {
        scheduleDraft.files = scheduleDraft.files.filter(entry => entry.id !== file.id);
        drawScheduleDraftFiles();
    });
}

function openScheduleEditor(event, schedule, focusDate = false) {
    if (data.apiVersion !== 3) { notice(restartMessage); return; }
    scheduleDraft = {...schedule, eventId:event.id, files:[...(schedule.files || [])]};
    $('schedule-name').value = schedule.name;
    $('schedule-start').value = schedule.start;
    $('schedule-end').value = schedule.end;
    $('schedule-notes').value = schedule.notes || '';
    $('schedule-completed').checked = schedule.completed === true;
    $('schedule-completed-date').value = schedule.completedDate || '';
    $('schedule-completed-date').required = schedule.completed === true;
    $('schedule-completed-date').setCustomValidity('');
    $('schedule-hint').textContent = '請先手動選擇完成日期，再切換為已完成；切回未完成會清空完成日期。';
    drawScheduleDraftFiles();
    if (!$('schedule-edit').open) $('schedule-edit').showModal();
    if (focusDate) $('schedule-completed-date').focus();
}

$('schedule-completed-date').oninput = () => $('schedule-completed-date').setCustomValidity('');
$('schedule-completed').onchange = () => {
    const date = $('schedule-completed-date');
    date.setCustomValidity('');
    if ($('schedule-completed').checked && !Number.isFinite(dayNumber(date.value))) {
        $('schedule-completed').checked = false;
        date.setCustomValidity('請先手動選擇完成日期，再切換為已完成。');
        date.reportValidity(); return;
    }
    date.required = $('schedule-completed').checked;
    if (!$('schedule-completed').checked) date.value = '';
};

function saveSchedule(eventId, schedule) {
    return api('/api/schedule',{method:'POST', headers:{'X-Local-Token':token,'Content-Type':'application/json'},body:JSON.stringify({eventId,schedule})});
}

$('schedule-form').onsubmit = async event => {
    event.preventDefault();
    const button = event.submitter || $('schedule-form').querySelector('[type=submit]');
    button.disabled = true;
    try {
        const schedule = {...scheduleDraft, name:$('schedule-name').value.trim(), start:$('schedule-start').value,
            end:$('schedule-end').value, notes:$('schedule-notes').value, completed:$('schedule-completed').checked,
            completedDate:$('schedule-completed').checked ? $('schedule-completed-date').value : ''};
        if (!schedule.name || !Number.isFinite(dayNumber(schedule.start)) || !Number.isFinite(dayNumber(schedule.end)) || schedule.end < schedule.start) throw Error('請填寫時程名稱與有效日期，結束日期不可早於開始日期。');
        if (schedule.completed && !Number.isFinite(dayNumber(schedule.completedDate))) throw Error('請先手動選擇完成日期。');
        await saveSchedule(scheduleDraft.eventId,schedule);
        $('schedule-edit').close();
        await refresh(); notice('時程已儲存');
    } catch (error) { alert(error.message); }
    finally { button.disabled = false; }
};

async function uploadSchedule(input, image) {
    const file = input.files[0]; input.value = '';
    if (!file) return;
    if (file.size > 25000000) return alert('單一檔案最多 25 MB。');
    if (image && !/^image\/(png|jpeg|gif|webp)$/.test(file.type)) return alert('請選 PNG、JPEG、GIF 或 WebP 圖片。');
    if (scheduleDraft.files.length >= 100) return alert('每條時程最多 100 個檔案。');
    const original = scheduleDraft;
    const controls = [$('schedule-add-image'),$('schedule-add-file'),$('schedule-form').querySelector('[type=submit]')];
    controls.forEach(button => button.disabled = true);
    try {
        const result = await api('/api/upload',{method:'POST',headers:{'X-Local-Token':token,'X-Image-Type':image ? file.type : ''},body:file});
        if (scheduleDraft !== original || !$('schedule-edit').open) return;
        scheduleDraft.files.push({id:result.id,name:file.name}); drawScheduleDraftFiles();
    } catch (error) { alert(error.message); }
    finally { controls.forEach(button => button.disabled = false); }
}
$('schedule-add-image').onclick = () => $('schedule-image-file').click();
$('schedule-add-file').onclick = () => $('schedule-attach-file').click();
$('schedule-image-file').onchange = event => uploadSchedule(event.target,true);
$('schedule-attach-file').onchange = event => uploadSchedule(event.target,false);

function renderTimeline() {
    const slot = $('timeline');
    slot.replaceChildren();
    const current = data.Events.find(event => event.id === eventId);
    const schedules = eventSchedules(current);
    if (!schedules.length) {
        slot.append(element('p', current ? '此事件尚未設定時程。' : '請展開事件，查看案件時程。', 'empty'));
        return;
    }
    const rows = schedules.map(schedule => ({...scheduleState(schedule), schedule, name:schedule.name, dates:schedule.start + ' 至 ' + schedule.end}));
    for (const row of rows) {
        const item = element('div', undefined, 'timeline-row ' + row.tone);
        const heading = element('div', undefined, 'bar');
        heading.append(element('strong', row.name), element('small', row.dates));
        const lane = element('button', undefined, 'timeline-lane');
        lane.type = 'button'; lane.title = '編輯這條時程';
        lane.setAttribute('aria-label', '編輯時程：' + row.name);
        lane.disabled = data.apiVersion !== 3;
        lane.onclick = () => openScheduleEditor(current, row.schedule);
        const track = element('span', undefined, 'schedule-track');
        track.style.width = '100%';
        track.setAttribute('role', 'progressbar');
        track.setAttribute('aria-label', row.name + ' 時間進度');
        track.setAttribute('aria-valuemin', '0');
        track.setAttribute('aria-valuemax', '100');
        track.setAttribute('aria-valuenow', String(Math.round(row.progress * 100)));
        track.setAttribute('aria-valuetext', row.label);
        const fill = element('span', undefined, 'schedule-fill');
        fill.style.width = row.progress * 100 + '%';
        track.append(fill); lane.append(track);
        item.append(heading, lane);
        if (row.schedule.notes) item.append(element('div', row.schedule.notes, 'schedule-notes'));
        const images = element('div', undefined, 'schedule-images');
        drawScheduleImages(images, row.schedule.files || []);
        if (images.childNodes.length) item.append(images);
        const footer = element('div', undefined, 'schedule-footer');
        const files = element('div'); drawFiles(files, row.schedule.files || []);
        const controls = element('div', undefined, 'schedule-controls');
        const toggle = element('button', row.completed ? '已完成' : '未完成', 'schedule-toggle');
        toggle.type = 'button'; toggle.setAttribute('role','switch');
        toggle.setAttribute('aria-checked',String(row.completed));
        toggle.setAttribute('aria-label',row.name + '：已完成');
        toggle.disabled = data.apiVersion !== 3;
        toggle.onclick = async () => {
            if (!row.completed) { openScheduleEditor(current,row.schedule,true); return; }
            toggle.disabled = true;
            try {
                await saveSchedule(current.id,{...row.schedule,completed:false,completedDate:''});
                await refresh();
            } catch (error) { alert(error.message); }
            finally { toggle.disabled = data.apiVersion !== 3; }
        };
        controls.append(toggle,element('div',row.label,'timeline-remaining'));
        footer.append(files,controls); item.append(footer);
        slot.append(item);
    }
    const openRows = rows.filter(row => !row.completed);
    const pending = openRows.filter(row => row.remaining >= 0).sort((a,b) => a.end - b.end);
    const overdue = openRows.length - pending.length;
    const summary = !openRows.length ? '全部時程已完成，共 ' + rows.length + ' 筆' : pending.length ? pending[0].name + '：' + pending[0].label + (overdue ? '；已逾期 ' + overdue + ' 筆' : '') : '尚未完成的時程已逾期，共 ' + overdue + ' 筆';
    slot.append(element('div', summary, 'timeline-summary'));
}

function addScheduleRow(schedule = {}) {
    const row = element('div', undefined, 'schedule-edit-row');
    row.schedule = {...schedule, id:schedule.id || crypto.randomUUID().replaceAll('-',''), files:[...(schedule.files || [])]};
    for (const [key, title] of [['name','名稱'], ['start','開始日期'], ['end','結束日期']]) {
        const label = element('label', title);
        const input = element('input');
        input.type = key === 'name' ? 'text' : 'date';
        input.dataset.field = key; input.required = true;
        if (key === 'name') input.maxLength = 80;
        input.value = schedule[key] || '';
        input.oninput = () => input.setCustomValidity('');
        label.append(input); row.append(label);
    }
    const remove = element('button', '移除'); remove.type = 'button';
    remove.onclick = () => row.remove(); row.append(remove);
    $('schedule-rows').append(row);
}

function readSchedules() {
    return [...$('schedule-rows').children].map(row => {
        const inputs = [...row.querySelectorAll('input')];
        const schedule = Object.fromEntries(inputs.map(input => [input.dataset.field, input.value.trim()]));
        if (!schedule.name || schedule.name.length > 80 || !Number.isFinite(dayNumber(schedule.start)) || !Number.isFinite(dayNumber(schedule.end)) || schedule.end < schedule.start) {
            throw Error('請填寫時程名稱與有效日期，結束日期不可早於開始日期。');
        }
        return {...row.schedule,...schedule};
    });
}
$('add-schedule').onclick = () => addScheduleRow();
$('event-filter').onchange = () => render();

async function toggleEventStatus(event, button) {
    button.disabled = true;
    try {
        await api('/api/status', {method:'POST', headers:{'X-Local-Token':token, 'Content-Type':'application/json'},
            body:JSON.stringify({id:event.id, status:event.status === 'completed' ? 'active' : 'completed'})});
        await refresh();
        notice(event.status === 'completed' ? '事件已改為進行中' : '事件已設為已結束');
    } catch (error) { alert(error.message); }
    finally { button.disabled = false; }
}

function switches(target, items, activeId, onSelect) {
    target.replaceChildren();
    for (const item of items) {
        const button = element('button', item.name);
        button.setAttribute('aria-pressed', String(item.id === activeId));
        button.onclick = () => onSelect(item.id);
        target.append(button);
    }
}

function render() {
    if (!data.Projects.some(project => project.id === projectId)) projectId = data.Projects[0]?.id;
    const sections = data.Sections.filter(section => section.projectId === projectId);
    if (!sections.some(section => section.id === sectionId)) sectionId = sections[0]?.id;
    switches($('projects'), data.Projects, projectId, id => {
        projectId = id;
        sectionId = eventId = undefined;
        render();
    });
    switches($('sections'), sections, sectionId, id => {
        sectionId = id;
        eventId = undefined;
        render();
    });
    $('rename-project').textContent = $('rename-section').textContent = '修改';
    $('rename-project').disabled = $('new-section').disabled = !projectId;
    $('rename-section').disabled = $('new-event').disabled = !sectionId;
    const writable = data.apiVersion === 3 && data.capabilities?.includes('attachment-cleanup');
    $('new-project').disabled = !writable;
    $('backend-warning').hidden = writable;
    $('backend-warning').textContent = writable ? '' : restartMessage;
    if (!writable) {
        for (const id of ['new-project','rename-project','new-section','rename-section','new-event']) $(id).disabled = true;
    }
    $('ip').textContent = (data.ip || []).join(' / ') || '未偵測到 IPv4';
    const filter = $('event-filter').value;
    const sectionEvents = data.Events.filter(event => event.sectionId === sectionId);
    const query = $('event-search').value.trim().toLocaleLowerCase();
    const events = sectionEvents.filter(event => (filter === 'all' || eventStatus(event) === filter) && (!query || eventSearchText(event).includes(query))).sort((a,b) => b.date.localeCompare(a.date));
    if (!events.some(event => event.id === eventId)) eventId = undefined;
    $('events').replaceChildren();
    if (!events.length) {
        const hint = !projectId ? '先新增一個專案。' : !sectionId ? '在這個專案新增一個分頁。' : !sectionEvents.length ? '這個分頁還沒有事件。' : filter === 'completed' ? '這個分頁沒有已結束的事件。' : '這個分頁沒有進行中的事件。';
        $('events').append(element('p', query ? '沒有符合搜尋條件的事件。' : hint, 'empty'));
    }
    for (const event of events) {
        const completed = eventStatus(event) === 'completed';
        const row = element('details', undefined, 'event ' + (completed ? 'event-completed' : 'event-active'));
        const summary = element('summary');
        summary.append(element('time', event.date), element('strong', event.title));
        addEventDragHandle(summary, event, writable);
        const toggle = element('button', completed ? '已結束' : '進行中', 'event-status');
        toggle.type = 'button';
        toggle.setAttribute('role', 'switch');
        toggle.setAttribute('aria-checked', String(completed));
        toggle.setAttribute('aria-label', event.title + '：已結束');
        toggle.title = completed ? '切換為進行中' : '設為已結束';
        toggle.disabled = !writable;
        toggle.onclick = click => {
            click.preventDefault(); click.stopPropagation();
            toggleEventStatus(event, toggle);
        };
        summary.append(toggle);
        const article = element('article');
        const content = element('div');
        content.innerHTML = cleanHTML(event.html);
        const files = element('div');
        drawFiles(files, event.files);
        const edit = element('button', '編輯事件');
        edit.disabled = !writable;
        edit.onclick = () => openEditor(event);
        article.append(content, files, edit);
        row.open = event.id === eventId;
        summary.onclick = click => {
            click.preventDefault();
            eventId = row.open ? undefined : event.id;
            render();
            $('events').querySelectorAll('summary')[events.indexOf(event)]?.focus();
        };
        row.append(summary, article);
        $('events').append(row);
    }
    renderTimeline();
}

async function refresh() {
    data = await api('/api/data');
    data.Events = data.Events.map(event => ({...event, status:eventStatus(event), schedules:eventSchedules(event)}));
    render();
}

function showName(kind, item) {
    nameRequest = {kind, id:item?.id, projectId};
    $('name-title').textContent = (item ? '修改' : '新增') + (kind === 'project' ? '專案' : '分頁');
    $('delete-container').hidden = !item;
    $('delete-container').disabled = !data.capabilities?.includes('container-delete');
    $('name-value').value = item?.name || '';
    $('names').showModal();
}
$('new-project').onclick = () => showName('project');
$('new-section').onclick = () => showName('section');
$('rename-project').onclick = () => showName('project', data.Projects.find(p => p.id === projectId));
$('rename-section').onclick = () => showName('section', data.Sections.find(s => s.id === sectionId));
$('name-form').onsubmit = async event => {
    event.preventDefault();
    const button = event.submitter || $('name-form').querySelector('[type=submit]');
    button.disabled = true;
    try {
        const result = await save({...nameRequest, name:$('name-value').value});
        if (nameRequest.kind === 'project') projectId = result.id;
        else sectionId = result.id;
        $('names').close();
        await refresh();
    } catch (error) { alert(error.message); }
    finally { button.disabled = false; }
};

function localDate(date) {
    return [date.getFullYear(), date.getMonth()+1, date.getDate()].map((n,i) => String(n).padStart(i ? 2 : 4, '0')).join('-');
}

function openEditor(event) {
    cellStart = cellEnd = undefined;
    draft = {id:event?.id, sectionId, status:eventStatus(event), files:[...(event?.files || [])]};
    selection = undefined;
    $('event-title').value = event?.title || '';
    $('event-date').value = event?.date || localDate(new Date());
    $('schedule-rows').replaceChildren();
    for (const schedule of eventSchedules(event)) addScheduleRow(schedule);
    editor.innerHTML = cleanHTML(event?.html || '');
    drawFiles($('edit-files'), draft.files, true);
    $('delete-event').textContent = event?.id ? '移除事件' : '取消新增';
    $('delete-event').disabled = !!event?.id && !data.capabilities?.includes('event-delete');
    $('delete-event').title = $('delete-event').disabled ? restartMessage : '';
    $('edit').showModal();
}
$('new-event').onclick = () => openEditor();
$('delete-event').onclick = async () => {
    if (!draft?.id) { $('edit').close(); return; }
    if (!data.capabilities?.includes('event-delete')) { notice(restartMessage); return; }
    const current = draft;
    const saved = data.Events.find(event => event.id === current.id);
    if (!confirm('確定移除「' + (saved?.title || $('event-title').value) + '」？此事件、所有時程及無其他引用的附件原檔將一併移除。')) return;
    const controls = [$('delete-event'),$('add-image'),$('add-file'),$('edit-form').querySelector('[type=submit]')];
    controls.forEach(button => button.disabled = true);
    try {
        await api('/api/delete-event',{method:'POST',headers:{'X-Local-Token':token,'Content-Type':'application/json'},body:JSON.stringify({id:current.id})});
        if (draft === current && $('edit').open) $('edit').close();
        if (eventId === current.id) eventId = undefined;
        await refresh(); notice('事件已移除');
    } catch (error) { alert(error.message); }
    finally { controls.forEach(button => button.disabled = false); }
};
$('edit-form').onsubmit = async event => {
    event.preventDefault();
    const button = event.submitter || $('edit-form').querySelector('[type=submit]');
    button.disabled = true;
    $('delete-event').disabled = true;
    try {
        const result = await save({...draft, kind:'event', title:$('event-title').value, date:$('event-date').value,
            schedules:readSchedules(), html:cleanHTML(editor.innerHTML)});
        eventId = result.id;
        if ($('event-filter').value !== 'all' && $('event-filter').value !== draft.status) $('event-filter').value = 'all';
        $('edit').close();
        await refresh();
        notice('事件已儲存');
    } catch (error) { alert(error.message); }
    finally { button.disabled = false; $('delete-event').disabled = !!draft?.id && !data.capabilities?.includes('event-delete'); }
};

document.addEventListener('selectionchange', () => {
    const current = window.getSelection();
    if (current.rangeCount && editor.contains(current.getRangeAt(0).commonAncestorContainer)) selection = current.getRangeAt(0).cloneRange();
});

function command(action, value) {
    editor.focus();
    const active = window.getSelection();
    const range = selection && editor.contains(selection.commonAncestorContainer) ? selection : document.createRange();
    if (range !== selection) {
        range.selectNodeContents(editor);
        range.collapse(false);
    }
    active.removeAllRanges();
    active.addRange(range);
    document.execCommand(action, false, value);
    if (active.rangeCount) selection = active.getRangeAt(0).cloneRange();
}

for (const button of document.querySelectorAll('[data-command]')) button.onclick = () => command(button.dataset.command);
$('font').onchange = event => { command('fontSize', event.target.value); event.target.value = ''; };
$('color').oninput = event => command('foreColor', event.target.value);
$('table').onclick = () => {
    const rows = prompt('表格有幾個 row？（1–20）', '3');
    if (rows === null) return;
    const columns = prompt('表格有幾個 column？（1–12）', '3');
    if (columns === null) return;
    if (!/^([1-9]|1[0-9]|20)$/.test(rows) || !/^([1-9]|1[0-2])$/.test(columns)) {
        alert('請輸入允許範圍內的整數。');
        return;
    }
    command('insertHTML', '<table><tbody>' + ('<tr>' + '<td>&nbsp;</td>'.repeat(+columns) + '</tr>').repeat(+rows) + '</tbody></table><p><br></p>');
};

editor.addEventListener('paste', event => {
    event.preventDefault();
    if (pasteTable(event.clipboardData)) return;
    command('insertText', event.clipboardData.getData('text/plain'));
});
editor.addEventListener('drop', event => event.preventDefault());
$('add-image').onclick = () => $('image-file').click();
$('add-file').onclick = () => $('attach-file').click();

async function upload(input, inline) {
    const file = input.files[0];
    input.value = '';
    if (!file) return;
    if (file.size > 25000000) return alert('單一檔案最多 25 MB。');
    if (inline && !/^image\/(png|jpeg|gif|webp)$/.test(file.type)) return alert('請選 PNG、JPEG、GIF 或 WebP 圖片。');
    const originalDraft = draft;
    const controls = [$('add-image'), $('add-file'), $('delete-event'), $('edit-form').querySelector('[type=submit]')];
    controls.forEach(button => button.disabled = true);
    try {
        const result = await api('/api/upload', {method:'POST', headers:{'X-Local-Token':token, 'X-Image-Type':inline ? file.type : ''}, body:file});
        if (draft !== originalDraft || !$('edit').open) return;
        draft.files.push({id:result.id, name:file.name});
        if (inline) command('insertImage', '/attachment/' + result.id);
        drawFiles($('edit-files'), draft.files, true);
    } catch (error) { alert(error.message); }
    finally { controls.forEach(button => button.disabled = false); }
}
$('image-file').onchange = event => upload(event.target, true);
$('attach-file').onchange = event => upload(event.target, false);

function setToolStatus(text) { $('tool-status').textContent = text; }
function drawCalendar() {
    const first = new Date(month.getFullYear(), month.getMonth(), 1);
    const days = new Date(first.getFullYear(), first.getMonth()+1, 0).getDate();
    $('month').textContent = first.getFullYear() + ' 年 ' + (first.getMonth()+1) + ' 月';
    const grid = $('month-grid');
    grid.replaceChildren();
    for (const label of '日一二三四五六') grid.append(element('strong', label));
    for (let i=0; i<first.getDay(); i++) grid.append(element('span', ''));
    for (let day=1; day<=days; day++) {
        const date = new Date(first.getFullYear(),first.getMonth(),day);
        grid.append(element('span', day, localDate(date) === localDate(new Date()) ? 'today' : ''));
    }
}
function stepMonth(offset) {
    month = new Date(month.getFullYear(), month.getMonth()+offset, 1);
    drawCalendar();
}
$('prev-month').onclick = () => stepMonth(-1);
$('next-month').onclick = () => stepMonth(1);

for (const button of document.querySelectorAll('[data-tool]')) button.onclick = () => {
    activeTool = button.dataset.tool;
    for (const name of ['pdf','ocr']) $(name).hidden = name !== activeTool;
    $('tool-title').textContent = {pdf:'PDF 切割',ocr:'OCR 辨識'}[activeTool];
    setOCRContext(button.closest('dialog') === $('edit'));
    setToolStatus('');
    $('tool').showModal();
};
for (const button of document.querySelectorAll('[data-close]')) button.onclick = () => {
    if (button.closest('dialog') !== $('tool') || !busy) button.closest('dialog').close();
};
$('tool').addEventListener('cancel', event => { if (busy) event.preventDefault(); });

async function toolTask(button, task) {
    if (busy) return;
    busy = button.disabled = true;
    setToolStatus('處理中…');
    try { await task(); }
    catch (error) { setToolStatus(error.message); }
    finally { busy = button.disabled = false; }
}

function selectPages(text, total) {
    const pages = [];
    for (const part of text.trim().split(/[,，]/)) {
        const match = part.trim().match(/^(\d+)\s*(?:-\s*(\d+))?$/);
        if (!match) throw Error('頁碼格式應為 2-5,8。');
        const first = Number(match[1]), last = Number(match[2] || match[1]);
        if (first < 1 || last < first || last > total) throw Error('頁碼須介於 1 至 ' + total + '，且範圍必須由小到大。');
        for (let page=first; page<=last; page++) pages.push(page-1);
    }
    return [...new Set(pages)];
}

$('split').onclick = () => toolTask($('split'), async () => {
    const file = $('pdf-file').files[0];
    if (!file) throw Error('請先選擇 PDF。');
    const original = await PDFLib.PDFDocument.load(await file.arrayBuffer());
    const output = await PDFLib.PDFDocument.create();
    const pages = await output.copyPages(original, selectPages($('pdf-pages').value, original.getPageCount()));
    for (const page of pages) output.addPage(page);
    const blob = new Blob([await output.save()], {type:'application/pdf'});
    const url = URL.createObjectURL(blob);
    const link = element('a');
    link.href = url;
    link.download = file.name.replace(/\.pdf$/i, '') + '-抽頁.pdf';
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    setToolStatus('已產生 ' + pages.length + ' 頁。');
});

$('recognize').onclick = () => toolTask($('recognize'), runOCR);

$('insert-ocr').onclick = () => {
    if (busy || !$('edit').open || !$('ocr-result').value.trim()) return;
    $('tool').close();
    command('insertText', $('ocr-result').value);
};

function updateClock() { renderTimeline(); drawCalendar(); }
setInterval(updateClock, 60000);
document.addEventListener('visibilitychange', () => { if (!document.hidden) updateClock(); });
drawCalendar();
