(async () => {
    const checks = [];
    function assert(value, message) { if (!value) throw Error(message); checks.push(message); }
    const original = {data, projectId, sectionId, eventId, month, calendarEventView,
        query:$('event-search').value, filter:$('event-filter').value};
    try {
        const schedules = Array.from({length:12}, (_, i) => ({id:'s' + i, name:i === 4 ? '完整長名稱'.repeat(25) : '重複時程',
            start:'2026-11-07', end:'2026-11-10', ...(i === 0 ? {completed:true, completedDate:'2026-11-08'} : {})}));
        data = {...data, Projects:[{id:'p1',name:'月曆專案一'},{id:'p2',name:'月曆專案二'},{id:'p3',name:'空專案'}],
            Sections:[{id:'a',projectId:'p1',name:'原分頁'},{id:'b',projectId:'p1',name:'其他分頁'},{id:'c',projectId:'p2',name:'其他專案分頁'}],
            Events:[{id:'e1',sectionId:'a',title:'同名 Event',date:'2026-11-01',html:'',schedules:[]},
                {id:'e2',sectionId:'b',title:'同名 Event',date:'2026-11-02',html:'',status:'completed',schedules},
                {id:'e3',sectionId:'c',title:'跨年 Event',date:'2026-12-01',html:'',schedules:[{id:'year',name:'跨年',start:'2026-12-30',end:'2027-01-03'}]}]};
        projectId = 'p1'; sectionId = 'a'; eventId = 'e1'; calendarEventView = undefined;
        month = new Date(2026,10,1); $('event-search').value = '無相符內容'; $('event-filter').value = 'active'; render();
        $('month-grid').click();
        assert($('schedule-calendar').open && calendarProjectId === 'p1' && calendarMonth.getMonth() === 10, 'Calendar opens on current project and mini-calendar month');
        assert($('calendar-projects').children.length === 3, 'Modal includes project switch buttons');
        assert(calendarSchedules().length === 12, 'Calendar includes all sections and completed events independently of main filters');
        const first = dayNumber('2026-11-01'), rows = calendarSchedules();
        const firstLayout = calendarWeekLayout(rows,first,first+6);
        assert(firstLayout.visible.length === 3 && firstLayout.hidden.get(first+6).length === 9, 'Twelve overlapping schedules produce three visible lanes and nine hidden schedules');
        assert(new Set(firstLayout.visible.map(v => v.lane)).size === 3, 'Overlapping segments occupy distinct lanes');
        assert(JSON.stringify(firstLayout.visible.map(v => v.row.key)) === JSON.stringify(calendarWeekLayout(calendarSchedules(),first,first+6).visible.map(v => v.row.key)), 'Repeated layouts preserve schedule order');
        assert($('calendar-weeks').querySelector('[aria-label^="2026-11-08，另有 9 條"]'), 'Hidden counts are calculated per date across week boundaries');
        assert(document.querySelectorAll('.calendar-segment').length === 6, 'Cross-week schedules continue as separate weekly segments');
        assert(![...document.querySelectorAll('.calendar-schedule-name')].some(n => n.textContent.includes('已完成')), 'Unfinished schedules take priority over completed schedules');
        assert([...document.querySelectorAll('.calendar-schedule-bar')].some(n => n.textContent.includes('‹')) && [...document.querySelectorAll('.calendar-schedule-bar')].some(n => n.textContent.includes('›')), 'Continuations are marked at both week boundaries');
        $('calendar-weeks').querySelector('[aria-label^="2026-11-08，另有"]').click();
        assert(!$('calendar-day-panel').hidden && $('calendar-day-list').children.length === 12, 'Overflow opens the complete inclusive daily list');
        assert($('calendar-day-list').lastChild.textContent.includes('已完成'), 'Daily list places completed schedules last');
        assert($('calendar-day-list').textContent.includes('完整長名稱'.repeat(25)) && $('calendar-day-list').textContent.includes('其他分頁'), 'Daily list preserves full names and source section');
        openCalendarDay('2026-11-07'); assert($('calendar-day-list').children.length === 12, 'Start date is inclusive');
        openCalendarDay('2026-11-10'); assert($('calendar-day-list').children.length === 12, 'End date is inclusive');
        openCalendarDay('2026-11-11'); assert($('calendar-day-list').textContent.includes('當天沒有時程'), 'Empty date has an explicit message');
        $('calendar-day-close').click(); assert($('calendar-day-panel').hidden, 'Daily list can be dismissed');
        $('calendar-projects').children[1].click();
        assert(projectId === 'p1' && sectionId === 'a' && calendarMonth.getMonth() === 10, 'Modal project switching preserves main selection and displayed month');
        $('calendar-next').click();
        assert(calendarMonth.getMonth() === 11 && document.querySelector('.calendar-schedule-name').textContent.includes('›'), 'December segment marks continuation into next year');
        $('calendar-next').click();
        assert(calendarMonth.getFullYear() === 2027 && document.querySelector('.calendar-schedule-name').textContent.includes('‹'), 'January segment marks continuation from previous year');
        $('calendar-projects').children[2].click();
        assert(!$('calendar-empty').hidden && !document.querySelector('.calendar-segment'), 'Empty project displays a clear empty state');
        $('calendar-today').click();
        assert(calendarMonth.getFullYear() === new Date().getFullYear() && calendarMonth.getMonth() === new Date().getMonth(), 'Today resets month and year');
        assert(!$('open-calendar') && !$('calendar').querySelector('h2') && !$('calendar-title') && !$('schedule-calendar').querySelector('[data-close]'), 'Calendar heading rows and open/close buttons are removed');
        assert($('schedule-calendar').getAttribute('aria-label'), 'Calendar dialog retains an accessible name');
        const dialog = $('schedule-calendar');
        const pointer = (type, x, y) => dialog.dispatchEvent(new PointerEvent(type, {bubbles:true,clientX:x,clientY:y}));
        const rect = dialog.getBoundingClientRect();
        pointer('pointerdown', rect.left+2, rect.top+2); pointer('click', rect.left+2, rect.top+2);
        assert(dialog.open, 'Clicking dialog padding does not close the calendar');
        pointer('pointerdown', rect.left+2, rect.top+2); pointer('click', 0, 0);
        assert(dialog.open, 'Dragging from inside to the backdrop does not close the calendar');
        pointer('pointerdown', 0, 0); pointer('click', 0, 0);
        assert(!dialog.open, 'Clicking backdrop closes the calendar');
        for (const key of ['Enter', ' ']) {
            $('month-grid').dispatchEvent(new KeyboardEvent('keydown', {key,bubbles:true,cancelable:true}));
            assert(dialog.open, 'Keyboard opens the calendar with '+JSON.stringify(key));
            dialog.close();
        }
        openScheduleCalendar();
        assert(calendarProjectId === projectId && calendarMonth.getMonth() === month.getMonth() && !calendarDay, 'Reopening resets modal scope and uses mini-calendar month');
        document.querySelector('.calendar-segment').click();
        assert(!$('schedule-calendar').open && projectId === 'p1' && sectionId === 'b' && eventId === 'e2', 'Segment click navigates to its exact event and section');
        assert(document.querySelectorAll('#events .event').length === 1 && document.querySelector('#events .event').open && document.querySelectorAll('#timeline .timeline-row').length === 12, 'Exact event view bypasses search/status filters and expands event and timeline');
        assert(!$('calendar-event-return').hidden, 'Focused view provides return to original list');
        $('return-event-list').click();
        assert(projectId === 'p1' && sectionId === 'a' && eventId === undefined && $('event-search').value === '無相符內容' && $('event-filter').value === 'active', 'Return restores original selection and filters (filtered-out event remains collapsed)');
        assert(!$('event-search').disabled && !$('event-filter').disabled && $('calendar-event-return').hidden, 'Return restores normal list controls');
        openScheduleCalendar(); openCalendarDay('2026-11-08'); $('calendar-day-list').firstChild.click();
        assert(eventId === 'e2' && document.querySelectorAll('#events .event').length === 1, 'Daily list navigates using event identity, even with duplicate titles');
        $('return-event-list').click();
        $('event-search').value = ''; eventId = 'e1'; render();
        openScheduleCalendar(); $('calendar-projects').children[1].click();
        $('calendar-next').click(); document.querySelector('.calendar-segment').click();
        assert(projectId === 'p2' && sectionId === 'c' && eventId === 'e3', 'Calendar can navigate into an event in another project');
        $('return-event-list').click();
        assert(projectId === 'p1' && sectionId === 'a' && eventId === 'e1' && document.querySelector('#events .event').open, 'Return restores an originally expanded event across projects');
        const oldProjects = data.Projects; data = {...data,Projects:[]};
        openScheduleCalendar();
        assert($('calendar-empty').textContent === '尚未建立專案。' && !document.querySelector('.calendar-segment'), 'Calendar opens safely without any projects');
        data = {...data,Projects:oldProjects}; $('schedule-calendar').close();
        const sameDay = {start:dayNumber('2026-11-08'),end:dayNumber('2026-11-08')};
        assert(calendarWeekLayout([sameDay],sameDay.start,sameDay.end).visible.length === 1, 'Single-day schedules render');
        const disjoint = [0,1,2,3].map(n => ({start:first+n,end:first+n}));
        assert(calendarWeekLayout(disjoint,first,first+6).visible.every(v => v.lane === 0), 'Nonoverlapping schedules reuse lanes without unnecessary hiding');
        const mixed = [true,true,false,false].map((completed,i) => ({key:String(i),completed,start:first,end:first+6}));
        const visibleOn = (layout, day) => layout.visible.filter(v => v.from <= day && v.to >= day).map(v => v.row);
        for (const count of [2,3,4]) {
            const layout = calendarWeekLayout(mixed.slice(0,count),first,first+6);
            assert(visibleOn(layout,first).length === Math.min(3,count) && layout.hidden.get(first).length === Math.max(0,count-3), 'Mixed schedules fill three slots with correct overflow for '+count+' schedules');
        }
        const mixedLayout = calendarWeekLayout(mixed,first,first+6);
        assert(visibleOn(mixedLayout,first).filter(r => !r.completed).length === 2, 'Two unfinished schedules and one completed schedule fill the calendar');
        for (const completed of [false,true]) {
            const rows = mixed.map(r => ({...r,completed}));
            const layout = calendarWeekLayout(rows,first,first+6);
            assert(visibleOn(layout,first).every((r,i) => r === rows[i]) && layout.hidden.get(first).length === 1, 'Same-status schedules preserve order: '+completed);
        }
        const spanning = [0,1,2].map(i => ({key:'long'+i,completed:true,start:first,end:first+6}));
        const urgent = {key:'urgent',completed:false,start:first+2,end:first+3};
        const splitLayout = calendarWeekLayout([...spanning,urgent],first,first+6);
        assert(visibleOn(splitLayout,first+2).includes(urgent) && splitLayout.hidden.get(first+2)[0] === spanning[2], 'Midweek unfinished schedule displaces a completed continuation');
        assert(visibleOn(splitLayout,first+4).includes(spanning[2]) && splitLayout.hidden.get(first+4).length === 0, 'Completed continuation returns when daily capacity is available');
        spanning[0].completed = false;
        const changed = calendarWeekLayout([...spanning,urgent],first,first+6);
        assert(visibleOn(changed,first+2).includes(spanning[0]) && visibleOn(changed,first+2).includes(urgent), 'Completion changes recompute daily priority');
        for (const term of ['1106','11/6','11/06','11-6','11-06']) {
            assert(eventDateMatches('2026-11-06',term) && eventDateMatches('2025-11-06',term), 'Month/day search spans years: '+term);
        }
        for (const term of ['2026-11-06','2026/11/6','20261106']) {
            assert(eventDateMatches('2026-11-06',term) && !eventDateMatches('2025-11-06',term), 'Full date search respects year: '+term);
        }
        for (const term of ['1131','13/6','2026-02-30','110','not-a-date']) assert(!eventDateMatches('2026-11-06',term), 'Invalid or nonmatching date does not match: '+term);
        data.Events[0].date = '2026-11-06';
        projectId = 'p1'; sectionId = 'a'; calendarEventView = undefined;
        $('event-filter').value = 'all'; $('event-search').value = '1106'; render();
        assert(document.querySelectorAll('#events .event').length === 1, 'Date query filters actual event list');
        $('event-filter').value = 'completed'; render();
        assert(!document.querySelector('#events .event'), 'Date search combines with status filter');
        $('event-filter').value = 'all'; $('event-search').value = '同名'; render();
        assert(document.querySelectorAll('#events .event').length === 1, 'Text search remains available in current section');
        const red = node => getComputedStyle(node).color === 'rgb(187, 48, 48)';
        month = new Date(2026,3,1); drawCalendar();
        const mini = $('month-grid');
        assert(red(mini.children[0]) && red(mini.children[6]) && !red(mini.children[1]), 'Mini calendar colors only weekend headings red');
        assert(red(mini.querySelector('[data-date="2026-04-03"]')) && red(mini.querySelector('[data-date="2026-04-04"]')) && !red(mini.querySelector('[data-date="2026-04-07"]')), 'Mini calendar marks weekday holidays, weekends and ordinary days correctly');
        openScheduleCalendar();
        assert(red($('calendar-weekdays').children[0]) && red($('calendar-weekdays').children[6]) && !red($('calendar-weekdays').children[2]), 'Modal colors weekend headings red');
        const holidayButton = $('calendar-weeks').querySelector('[data-date="2026-04-04"]');
        assert(red(holidayButton) && holidayButton.parentNode.textContent.includes('兒童節'), 'Modal labels named weekend holiday in red');
        assert(red($('calendar-weeks').querySelector('[data-date="2026-04-03"]')) && $('calendar-weeks').textContent.includes('補假'), 'Modal marks substitute days off');
        holidayButton.click();
        assert($('calendar-day-title').textContent.includes('兒童節'), 'Daily panel preserves full holiday name');
        $('calendar-day-close').click();
        assert(document.activeElement.dataset.date === '2026-04-04', 'Closing daily panel restores focus to holiday date');
        calendarMonth = new Date(2027,5,1); renderScheduleCalendar();
        assert($('calendar-weeks').querySelector('[data-date="2027-06-09"]').parentNode.textContent.includes('端午節'), 'Next year lunar holiday uses bundled official date');
        assert($('calendar-holiday-note').hidden && calendarHolidays['2026-06-19'] === '端午節' && calendarHolidays['2027-06-19'] === undefined, 'Holiday dates are year-specific');
        calendarMonth = new Date(2028,0,1); renderScheduleCalendar();
        assert(!$('calendar-holiday-note').hidden && !$('calendar-weeks').querySelector('.calendar-holiday') && red($('calendar-weeks').querySelector('[data-date="2028-01-01"]')), 'Unsupported year shows weekend colors and explicit coverage notice without invented holidays');
        return {passed:checks.length,checks};
    } finally {
        $('schedule-calendar').close();
        data = original.data; projectId = original.projectId; sectionId = original.sectionId; eventId = original.eventId;
        month = original.month; calendarEventView = original.calendarEventView;
        $('event-search').value = original.query; $('event-filter').value = original.filter;
        render(); drawCalendar();
    }
})()
