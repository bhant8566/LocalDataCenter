'use strict';

let calendarProjectId, calendarMonth, calendarDay;

function calendarDate(day) { return new Date(day * 86400000).toISOString().slice(0, 10); }

function calendarSchedules() {
    const sections = new Map(data.Sections.filter(s => s.projectId === calendarProjectId).map(s => [s.id, s]));
    return data.Events.filter(e => sections.has(e.sectionId)).flatMap(event => eventSchedules(event).map((schedule, index) => ({
        event, schedule, section:sections.get(event.sectionId), ...scheduleState(schedule),
        key:event.id + ':' + (schedule.id || String(index).padStart(8, '0'))
    }))).filter(row => Number.isFinite(row.start) && Number.isFinite(row.end) && row.end >= row.start)
        .sort((a, b) => a.start - b.start || a.end - b.end || a.key.localeCompare(b.key));
}

// Assign each weekly segment once, so its label and bar stay in the same lane all week.
function calendarWeekLayout(rows, start, end) {
    const lanes = [-Infinity, -Infinity, -Infinity], visible = [], hidden = [];
    for (const row of rows) {
        if (row.end < start || row.start > end) continue;
        const from = Math.max(start, row.start), to = Math.min(end, row.end);
        const lane = lanes.findIndex(last => last < from);
        if (lane < 0) hidden.push(row);
        else { lanes[lane] = to; visible.push({row, from, to, lane}); }
    }
    return {visible, hidden};
}

function openScheduleCalendar() {
    calendarProjectId = projectId;
    calendarMonth = new Date(month.getFullYear(), month.getMonth(), 1);
    calendarDay = undefined;
    renderScheduleCalendar();
    if (!$('schedule-calendar').open) $('schedule-calendar').showModal();
}

function openCalendarDay(date) {
    calendarDay = date;
    renderScheduleCalendar();
    $('calendar-day-close').focus();
}

function openCalendarEvent(id) {
    const event = data.Events.find(e => e.id === id);
    const section = data.Sections.find(s => s.id === event?.sectionId);
    if (!event || !section) { notice('此事件已不存在，請重新選擇。'); renderScheduleCalendar(); return; }
    const previous = calendarEventView?.previous || {projectId, sectionId, eventId,
        query:$('event-search').value, filter:$('event-filter').value, scrollY:window.scrollY};
    calendarEventView = {targetId:id, previous};
    projectId = section.projectId; sectionId = section.id; eventId = id;
    $('schedule-calendar').close();
    render();
    const summary = $('events').querySelector('summary');
    summary?.focus({preventScroll:true});
    summary?.scrollIntoView({block:'center'});
}

$('return-event-list').onclick = () => {
    if (!calendarEventView) return;
    const previous = calendarEventView.previous;
    calendarEventView = undefined;
    projectId = previous.projectId; sectionId = previous.sectionId; eventId = previous.eventId;
    $('event-search').value = previous.query; $('event-filter').value = previous.filter;
    render();
    $('event-search').focus({preventScroll:true});
    window.scrollTo(0, previous.scrollY);
};

function renderScheduleCalendar() {
    if (!data.Projects.some(p => p.id === calendarProjectId)) calendarProjectId = data.Projects[0]?.id;
    switches($('calendar-projects'), data.Projects, calendarProjectId, id => {
        calendarProjectId = id;
        renderScheduleCalendar();
        $('calendar-projects').querySelector('[aria-pressed=true]')?.focus();
    });
    const year = calendarMonth.getFullYear(), m = calendarMonth.getMonth();
    $('calendar-month').textContent = year + ' 年 ' + (m + 1) + ' 月';
    const first = dayNumber(localDate(new Date(year, m, 1)));
    const last = dayNumber(localDate(new Date(year, m + 1, 0)));
    const gridStart = first - new Date(year, m, 1).getDay();
    const today = localDate(new Date());
    const rows = calendarSchedules();
    $('calendar-empty').hidden = rows.some(row => row.start <= last && row.end >= first);
    $('calendar-empty').textContent = data.Projects.length ? '此專案本月沒有時程。' : '尚未建立專案。';
    $('calendar-weekdays').replaceChildren(...[...'日一二三四五六'].map(day => element('strong', day)));
    const weeks = $('calendar-weeks'); weeks.replaceChildren();
    for (let start = gridStart; start <= last; start += 7) {
        const week = element('div', undefined, 'calendar-week');
        const weekStart = Math.max(start, first), weekEnd = Math.min(start + 6, last);
        const layout = calendarWeekLayout(rows, weekStart, weekEnd);
        for (let offset = 0; offset < 7; offset++) {
            const day = start + offset, date = calendarDate(day);
            const cell = element('div', undefined, 'calendar-cell');
            cell.style.gridColumn = String(offset + 1); cell.style.gridRow = '1 / 6';
            if (day < first || day > last) cell.classList.add('calendar-outside');
            else {
                const button = element('button', String(Number(date.slice(-2))), 'calendar-date' + (date === today ? ' today' : ''));
                button.type = 'button'; button.setAttribute('aria-label', date + ' 的時程');
                button.setAttribute('aria-pressed', String(calendarDay === date));
                if (date === today) button.setAttribute('aria-current', 'date');
                button.onclick = () => openCalendarDay(date); cell.append(button);
                cell.onclick = e => { if (e.target === cell) openCalendarDay(date); };
                const count = layout.hidden.filter(row => row.start <= day && row.end >= day).length;
                if (count) {
                    const more = element('button', '另有 ' + count + ' 條時程', 'calendar-more');
                    more.type = 'button'; more.onclick = () => openCalendarDay(date);
                    more.style.gridColumn = String(offset + 1); more.style.gridRow = '5';
                    more.setAttribute('aria-label', date + '，另有 ' + count + ' 條時程，查看當日全部時程');
                    week.append(more);
                }
            }
            week.append(cell);
        }
        for (const {row, from, to, lane} of layout.visible) {
            const button = element('button', undefined, 'calendar-segment timeline-row ' + row.tone);
            button.type = 'button';
            button.style.gridColumn = (from - start + 1) + ' / ' + (to - start + 2);
            button.style.gridRow = String(lane + 2);
            button.dataset.scheduleKey = row.key;
            const before = row.start < from, after = row.end > to;
            const name = element('span', (before ? '‹ ' : '') + row.schedule.name + (row.completed ? '（已完成）' : '') + (after ? ' ›' : ''), 'calendar-schedule-name');
            const track = element('span', undefined, 'calendar-schedule-bar');
            track.textContent = (before ? '‹' : '') + (after ? ' ›' : '');
            button.append(name, track);
            // Status is computed by the same function as the existing event timeline.
            button.setAttribute('aria-label', row.schedule.name + '，' + row.schedule.start + ' 至 ' + row.schedule.end + '，' + row.event.title + '，' + row.label);
            button.onclick = () => openCalendarEvent(row.event.id);
            week.append(button);
        }
        weeks.append(week);
    }
    $('calendar-day-panel').hidden = !calendarDay;
    $('calendar-body').classList.toggle('with-day', !!calendarDay);
    const list = $('calendar-day-list'); list.replaceChildren();
    if (!calendarDay) return;
    const day = dayNumber(calendarDay), selected = rows.filter(row => row.start <= day && row.end >= day);
    $('calendar-day-title').textContent = calendarDay + ' · ' + selected.length + ' 條時程';
    if (!selected.length) list.append(element('p', '當天沒有時程。', 'empty'));
    for (const row of selected) {
        const button = element('button', undefined, 'calendar-day-item timeline-row ' + row.tone);
        button.type = 'button';
        button.append(element('strong', row.schedule.name), element('span', row.schedule.start + ' 至 ' + row.schedule.end),
            element('span', 'Event：' + row.event.title), element('span', '分頁：' + row.section.name),
            element('span', (row.completed ? '已完成 · ' : '未完成 · ') + row.label));
        button.onclick = () => openCalendarEvent(row.event.id);
        list.append(button);
    }
}

$('open-calendar').onclick = openScheduleCalendar;
$('month-grid').onclick = openScheduleCalendar;
$('month-grid').onkeydown = event => {
    if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); openScheduleCalendar(); }
};
$('calendar-day-close').onclick = () => {
    const date = calendarDay;
    calendarDay = undefined; renderScheduleCalendar();
    $('calendar-weeks').querySelector('[aria-label="' + date + ' 的時程"]')?.focus();
};
for (const [id, offset] of [['calendar-prev', -1], ['calendar-next', 1]]) $(id).onclick = () => {
    calendarMonth = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth() + offset, 1);
    calendarDay = undefined; renderScheduleCalendar();
};
$('calendar-today').onclick = () => {
    const now = new Date(); calendarMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    calendarDay = undefined; renderScheduleCalendar();
};
