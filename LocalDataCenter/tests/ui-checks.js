(async () => {
    const checks = [];
    function assert(value, message) { if (!value) throw Error(message); checks.push(message); }
    async function until(test) {
        for (let i=0; i<600; i++) { if (test()) return; await new Promise(resolve => setTimeout(resolve,100)); }
        throw Error('Timed out waiting for UI');
    }
    await until(() => typeof data !== 'undefined' && data.apiVersion === 3);
    assert($('event-filter').value === 'all', 'Default event filter is all');
    assert($('calendar').closest('aside.right') && !$('calendar').hidden, 'Calendar is visible below tools');
    assert(!document.querySelector('[data-tool="calendar"]'), 'Calendar tool button removed');
    const previousMonth = $('month').textContent;
    $('next-month').click(); assert($('month').textContent !== previousMonth, 'Calendar next month works');
    $('prev-month').click(); assert($('month').textContent === previousMonth, 'Calendar previous month works');

    const noon = new Date(2026,10,8,12);
    const oneDay = scheduleState({start:'2026-11-08',end:'2026-11-08'},noon);
    assert(oneDay.progress === 0.5 && oneDay.label === '今日到期', 'Same-day schedule is half full at noon');
    assert(scheduleState({start:'2026-11-08',end:'2026-11-08'},new Date(2026,10,9)).label === '已逾期 1 日', 'Day after deadline is overdue');
    assert(scheduleState({start:'2026-11-10',end:'2026-11-12'},noon).progress === 0, 'Future schedule is empty');
    assert(scheduleState({start:'2026-11-10',end:'2026-11-12'},noon).label === '尚未開始', 'Future schedule displays only not-started text without remaining days');
    assert(scheduleState({start:'2026-11-08',end:'2026-11-12'},noon).label === '剩餘 4 日', 'Countdown begins on schedule start date');
    assert(scheduleState({start:'2026-11-01',end:'2026-11-07'},noon).progress === 1, 'Expired schedule is full');
    assert(dayNumber('2027-01-01') - dayNumber('2026-12-30') === 2, 'Calendar day math crosses years');
    assert(Number.isNaN(dayNumber('2026-02-30')), 'Impossible dates rejected');
    for (const [end,tone] of [['2026-11-07','danger'],['2026-11-08','danger'],['2026-11-12','danger'],['2026-11-13','warning'],['2026-11-17','warning'],['2026-11-18','normal']]) {
        assert(scheduleState({start:'2026-11-01',end},noon).tone === tone, 'Deadline color boundary: ' + end);
    }

    projectId = (await save({kind:'project',name:'UI Test'})).id;
    sectionId = (await save({kind:'section',projectId,name:'Tests'})).id;
    const schedules = [{name:'Long',start:'2026-11-01',end:'2026-11-10'}, {name:'Overlap',start:'2026-11-06',end:'2026-11-10'}, {name:'Single day',start:'2026-11-08',end:'2026-11-08'}];
    eventId = (await save({kind:'event',sectionId,title:'Multiple schedules',date:'2026-11-08',schedules,files:[],html:'<p>Editor content</p>'})).id;
    await refresh();
    assert(document.querySelectorAll('.timeline-row').length === 3, 'Selected event displays all schedules');
    assert(!document.querySelector('.timeline-axis'), 'Timeline date axis is removed');
    const tracks = document.querySelectorAll('.schedule-track');
    assert([...tracks].every(track => track.style.width === '100%' && !track.style.marginLeft), 'Each schedule uses its own full-width percentage bar');
    assert([...tracks].every(track => !track.textContent && track.querySelector('.schedule-fill').style.width.endsWith('%')), 'Percentage is shown by fill width without numeric text');
    assert(document.querySelector('.timeline-row .bar small').textContent === schedules[0].start + ' 至 ' + schedules[0].end, 'Schedule date text keeps its original format');
    assert([...document.querySelectorAll('.timeline-row')].every(row => row.querySelector('.timeline-remaining').textContent === row.querySelector('[role=progressbar]').getAttribute('aria-valuetext')), 'Remaining-days text is preserved');
    assert(!!document.querySelector('.timeline-summary'), 'Deadline summary rendered');
    const event = data.Events.find(entry => entry.id === eventId);
    const snapshot = data;
    const legacy = {...event}; delete legacy.status; delete legacy.schedules;
    data = {...data, apiVersion:undefined, Events:[legacy]};
    $('event-filter').value = 'active'; render();
    assert(document.querySelector('.event-active'), 'Legacy events without status remain visible in active filter');
    assert(document.querySelectorAll('.timeline-row').length === 1, 'Legacy start/end dates render a timeline without schedules');
    assert(!$('backend-warning').hidden && $('new-event').disabled && document.querySelector('.event-status').disabled, 'Old backend shows restart warning and disables writes');
    let oldWriteRejected = false;
    try { await save({...event,kind:'event'}); } catch (error) { oldWriteRejected = error.message.includes('PowerShell'); }
    assert(oldWriteRejected, 'Writes to old backend are blocked before data can be lost');
    assert(eventSchedules({...legacy,schedules:[]}).length === 0, 'Explicit empty schedules do not resurrect legacy dates');
    data = snapshot; $('event-filter').value = 'all'; render();
    assert($('backend-warning').hidden && !$('new-project').disabled, 'Compatible backend clears restart warning and restores editing');
    openEditor(event);
    assert($('schedule-rows').children.length === 3, 'Editor restores schedule rows');
    addScheduleRow({name:'Added',start:'2026-11-09',end:'2026-11-09'});
    assert(readSchedules().length === 4, 'Editor adds schedules');
    $('schedule-rows').lastChild.querySelector('button').click();
    assert(readSchedules().length === 3, 'Editor removes schedules');
    $('schedule-rows').firstChild.querySelector('input').value = '  ';
    let rejected = false; try { readSchedules(); } catch { rejected = true; }
    assert(rejected, 'Frontend rejects blank schedule names');
    $('schedule-rows').firstChild.querySelector('input').value = 'Long';
    $('schedule-rows').replaceChildren(); addScheduleRow(schedules[0]);
    $('edit-form').requestSubmit();
    await until(() => !$('edit').open && data.Events.find(entry=>entry.id===event.id).schedules.length === 1 && document.querySelectorAll('.timeline-row').length === 1);
    assert($('event-filter').value === 'all', 'Saving a single schedule preserves all filter and displays its bar');
    openEditor(data.Events.find(entry=>entry.id===event.id));
    addScheduleRow(schedules[1]); addScheduleRow(schedules[2]);
    $('edit-form').requestSubmit();
    await until(() => !$('edit').open && data.Events.find(entry=>entry.id===event.id).schedules.length === 3 && document.querySelectorAll('.timeline-row').length === 3);
    assert(true, 'Submitting the editor saves and immediately displays multiple schedule bars');
    $('edit').close();
    const scheduleId = data.Events.find(entry=>entry.id===event.id).schedules[0].id;
    document.querySelector('.timeline-lane').click();
    assert($('schedule-edit').open && scheduleDraft.id === scheduleId, 'Clicking a time bar opens its independent editor');
    $('schedule-completed').click();
    assert(!$('schedule-completed').checked && !$('schedule-completed-date').value, 'Completion requires a manually selected date');
    let missingDateRejected = false;
    try { await saveSchedule(event.id,{...scheduleDraft,completed:true,completedDate:''}); } catch { missingDateRejected = true; }
    assert(missingDateRejected, 'Backend rejects completion without a date');
    const fixturePng = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII='),char=>char.charCodeAt(0));
    await uploadSchedule({files:[new File([fixturePng],'schedule.png',{type:'image/png'})]},true);
    await uploadSchedule({files:[new File(['中文文件已繳齊'],'schedule.txt',{type:'text/plain'})]},false);
    assert(scheduleDraft.files.length === 2 && $('schedule-edit-images').querySelector('img'), 'Schedule editor accepts its own images and files');
    $('schedule-notes').value = '文件已繳齊 <img src=x onerror=alert(1)>';
    await openPreview(scheduleDraft.files[0]);
    await until(() => $('preview-content').querySelector('img')?.complete);
    const bounds = $('preview').getBoundingClientRect();
    $('preview').dispatchEvent(new PointerEvent('pointerdown',{clientX:bounds.left+10,clientY:bounds.top+10,bubbles:true}));
    $('preview').dispatchEvent(new MouseEvent('click',{clientX:bounds.left+10,clientY:bounds.top+10,bubbles:true}));
    assert($('preview').open, 'Clicking inside the preview keeps it open');
    $('preview').dispatchEvent(new MouseEvent('click',{clientX:0,clientY:0,bubbles:true}));
    assert($('preview').open, 'Dragging from inside to outside does not discard the preview');
    $('preview').dispatchEvent(new PointerEvent('pointerdown',{clientX:0,clientY:0,bubbles:true}));
    $('preview').dispatchEvent(new MouseEvent('click',{clientX:0,clientY:0,bubbles:true}));
    await until(() => !$('preview').open && !previewSession);
    assert($('schedule-edit').open && $('schedule-notes').value.startsWith('文件已繳齊'), 'Backdrop click closes preview and preserves schedule draft');
    await openPreview(scheduleDraft.files[1]);
    assert($('preview-content').querySelector('pre').textContent === '中文文件已繳齊', 'Schedule attachment uses existing UTF-8 preview');
    $('preview').close(); await new Promise(resolve=>setTimeout(resolve,20));
    $('schedule-completed-date').value = '2026-11-08';
    $('schedule-completed-date').dispatchEvent(new Event('input'));
    $('schedule-completed').click();
    $('schedule-form').requestSubmit();
    await until(() => !$('schedule-edit').open && data.Events.find(entry=>entry.id===event.id).schedules[0].completed);
    const completedRow = document.querySelector('.timeline-row');
    assert(completedRow.classList.contains('normal') && completedRow.querySelector('.schedule-fill').style.width === '100%', 'Completed schedule is full and green regardless of deadline');
    assert(completedRow.querySelector('.timeline-remaining').textContent === '11月8日已完成', 'Completion date replaces remaining days');
    assert(completedRow.querySelector('.schedule-notes').textContent.includes('<img') && !completedRow.querySelector('.schedule-notes img'), 'Schedule notes display safely below the bar');
    assert(completedRow.querySelectorAll('.file-icon').length === 2 && completedRow.querySelector('.schedule-images img'), 'Schedule files and image appear below the bar');
    assert(completedRow.querySelector('.files').compareDocumentPosition(completedRow.querySelector('.schedule-controls')) & Node.DOCUMENT_POSITION_FOLLOWING, 'Attachment icons precede the completion controls');
    assert(data.Events.find(entry=>entry.id===event.id).status === 'active', 'Completing a schedule leaves event status independent');
    openEditor(data.Events.find(entry=>entry.id===event.id));
    $('edit-form').requestSubmit();
    await until(() => !$('edit').open && data.Events.find(entry=>entry.id===event.id).schedules[0].completed);
    const preserved = data.Events.find(entry=>entry.id===event.id).schedules[0];
    assert(preserved.files.length === 2 && preserved.notes.startsWith('文件已繳齊') && preserved.completedDate === '2026-11-08', 'Editing the event preserves schedule attachments, notes and completion');
    document.querySelector('.schedule-toggle').click();
    await until(() => !data.Events.find(entry=>entry.id===event.id).schedules[0].completed);
    const reopened = data.Events.find(entry=>entry.id===event.id).schedules[0];
    assert(!reopened.completedDate && reopened.files.length === 2 && !document.querySelector('.timeline-remaining').textContent.includes('已完成'), 'Reopening clears completion date and restores countdown without losing attachments');
    await saveSchedule(event.id,{...reopened,completed:true,completedDate:'2026-11-08'}); await refresh();
    openEditor(data.Events.find(entry=>entry.id===event.id));

    if (!window.statusOnly) {
    async function attachment(name, bytes, type='') {
        const result = await api('/api/upload',{method:'POST',headers:{'X-Local-Token':token,'X-Image-Type':type},body:bytes});
        return {id:result.id,name};
    }
    const text = '<img src=x onerror=alert(1)> 中文預覽';
    const textFile = await attachment('notes.txt',new TextEncoder().encode(text));
    await openPreview(textFile);
    assert($('preview-content').querySelector('pre').textContent === text && !$('preview-content').querySelector('img'), 'UTF-8 text is rendered without executing HTML');
    assert($('edit').open && editor.textContent === 'Editor content', 'Preview preserves open editor');
    assert($('preview-download').download === 'notes.txt', 'Download retains original filename');
    $('preview').close(); await new Promise(resolve => setTimeout(resolve,20));
    assert(!previewSession && !$('preview-content').children.length, 'Closing preview releases state');

    const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII='),char=>char.charCodeAt(0));
    const image = await attachment('image.png',png,'image/png');
    const stored = new Uint8Array(await (await fetch('/attachment/'+image.id)).arrayBuffer());
    assert(stored.length === png.length && stored.every((value,index)=>value===png[index]), 'Image bytes remain unchanged');
    await openPreview(image); await until(() => $('preview-content').querySelector('img')?.complete);
    assert($('preview-content').querySelector('img').naturalWidth === 1, 'Image preview loads');
    const imageUrl = previewSession.url;
    await openPreview(await attachment('document.docx',new Uint8Array([1,2,3])));
    assert($('preview-status').textContent.includes('不支援'), 'Unsupported documents offer download fallback');
    let revoked = false; try { await fetch(imageUrl); } catch { revoked = true; }
    assert(revoked, 'Image object URL is revoked');

    const pdf = await PDFLib.PDFDocument.create(); pdf.addPage([200,200]); pdf.addPage([300,200]);
    const pdfFile = await attachment('two-pages.pdf',await pdf.save());
    await openPreview(pdfFile);
    await until(() => $('preview-content').querySelector('.pdf-page canvas'));
    assert($('preview-content').querySelectorAll('.pdf-page').length === 2 && /[12] \/ 2/.test($('preview-page').textContent), 'PDF pages use continuous scroll locally: ' + $('preview-page').textContent + ' ' + $('preview-status').textContent);
    $('preview-next').click(); await until(() => $('preview-page').textContent === '2 / 2');
    assert($('preview-next').disabled && !$('preview-prev').disabled, 'PDF navigation updates boundaries');
    await openPreview(await attachment('broken.pdf',new TextEncoder().encode('invalid PDF')));
    assert($('preview-status').textContent.includes('無法預覽') && $('preview-download').getAttribute('href'), 'Damaged PDF retains download fallback');
    assert(!$('preview-status').textContent.includes('UTF-8'), 'PDF errors do not suggest changing text encoding');
    await openPreview({id:'00000000000000000000000000000000.bin',name:'missing.pdf'});
    assert($('preview-status').textContent.includes('找不到附件原檔') && $('preview-status').textContent.includes('Attachments') && $('preview-status').textContent.includes('HTTP 404'), 'Missing restored PDF reports actual attachment location and HTTP status');
    assert(!$('preview-status').textContent.includes('UTF-8'), 'Missing PDF is not reported as a text encoding problem');
    const loading = openPreview(pdfFile);
    $('preview').close(); disposePreview(); await loading;
    assert(!previewSession, 'Closing during loading does not restore preview');
    await new Promise(resolve => setTimeout(resolve,20));

    draft.files = [textFile,image,pdfFile]; drawFiles($('edit-files'),draft.files,true);
    assert($('edit-files').querySelectorAll('.file-icon').length === 3, 'Attachments show icons');
    $('edit-files').querySelector('.file button:last-child').click();
    assert(draft.files.length === 2, 'Attachment removal still works');
    }
    $('edit').close();
    const selectedProject = projectId;
    projectId = (await save({kind:'project',name:'Empty project'})).id;
    sectionId = eventId = undefined; await refresh();
    assert(!document.querySelector('.timeline-row'), 'Changing project clears timeline');
    projectId = selectedProject; sectionId = event.sectionId; eventId = event.id; render();
    updateClock(); assert(document.querySelectorAll('.timeline-row').length === 3, 'Clock refresh preserves selected event');
    document.dispatchEvent(new Event('visibilitychange'));
    assert($('calendar').querySelector('.today'), 'Visible calendar marks today');
    assert(event.status === 'active' && document.querySelector('.event-active'), 'New events default to active with green rows');
    $('event-filter').value = 'active'; $('event-filter').dispatchEvent(new Event('change'));
    const statusButton = document.querySelector('.event-status');
    statusButton.click();
    await until(() => data.Events.find(entry=>entry.id===event.id).status === 'completed' && !document.querySelector('.event'));
    assert(!document.querySelector('.event') && !document.querySelector('.timeline-row'), 'Completed event leaves active filter and clears selected timeline');
    $('event-filter').value = 'completed'; $('event-filter').dispatchEvent(new Event('change'));
    assert(document.querySelector('.event-completed .event-status').getAttribute('aria-checked') === 'true', 'Completed filter shows gray row and enabled completed state');
    const completed = data.Events.find(entry => entry.id === event.id);
    openEditor(completed);
    assert(draft.status === 'completed', 'Editing retains completed status');
    $('edit').close();
    const {status, ...withoutStatus} = completed;
    await save({...withoutStatus, kind:'event'}); await refresh();
    assert(data.Events.find(entry=>entry.id===event.id).status === 'completed', 'Saving without status preserves current state');
    await toggleEventStatus(data.Events.find(entry=>entry.id===event.id),document.querySelector('.event-status'));
    assert(!document.querySelector('.event'), 'Reopening event removes it from completed filter');
    $('event-filter').value = 'active'; $('event-filter').dispatchEvent(new Event('change'));
    assert(document.querySelector('.event-active'), 'Reopened event returns to active filter');
    const active = data.Events.find(entry=>entry.id===event.id);
    await toggleEventStatus(active,document.querySelector('.event-status'));
    let invalidStatus = false;
    try { await api('/api/status',{method:'POST',headers:{'X-Local-Token':token,'Content-Type':'application/json'},body:JSON.stringify({id:event.id,status:'invalid'})}); } catch { invalidStatus = true; }
    assert(invalidStatus, 'Backend rejects invalid event status');
    await save({kind:'event',sectionId:event.sectionId,title:'Active event',date:'2026-11-08',schedules:[],files:[],html:''});
    await refresh();
    $('event-filter').value = 'all'; $('event-filter').dispatchEvent(new Event('change'));
    assert(document.querySelector('.event-active') && document.querySelector('.event-completed'), 'All filter shows active and completed rows together');
    const originalCount = data.Events.length;
    openEditor();
    assert($('delete-event').textContent === '取消新增', 'Unsaved events offer cancel instead of server deletion');
    $('delete-event').click();
    assert(!$('edit').open && data.Events.length === originalCount, 'Canceling an unsaved event leaves saved data intact');
    const sharedFiles = data.Events.find(entry=>entry.id===event.id).schedules[0].files;
    const deleteId = (await save({kind:'event',sectionId:event.sectionId,title:'Delete fixture',date:'2026-11-08',
        schedules:[{name:'Delete schedule',start:'2026-11-08',end:'2026-11-09'}],files:sharedFiles,html:'Delete body'})).id;
    eventId = deleteId; await refresh();
    openEditor(data.Events.find(entry=>entry.id===deleteId));
    assert($('delete-event').textContent === '移除事件' && !$('delete-event').disabled, 'Saved events expose an enabled remove button');
    const originalConfirm = window.confirm;
    try {
        window.confirm = () => false;
        $('delete-event').click();
        assert($('edit').open && data.Events.some(entry=>entry.id===deleteId), 'Canceling deletion keeps event and editor intact');
        window.confirm = () => true;
        $('delete-event').click();
        await until(() => !$('edit').open && !data.Events.some(entry=>entry.id===deleteId));
    } finally { window.confirm = originalConfirm; }
    assert(!document.querySelector('.timeline-row'), 'Deleting selected event clears its time bars');
    assert(data.Events.length === originalCount && data.Events.some(entry=>entry.id===event.id), 'Deleting an event preserves other events');
    assert((await fetch('/attachment/'+sharedFiles[0].id)).ok, 'Event deletion preserves shared attachment files');
    let missingDeleteRejected = false;
    try { await api('/api/delete-event',{method:'POST',headers:{'X-Local-Token':token,'Content-Type':'application/json'},body:JSON.stringify({id:deleteId})}); } catch { missingDeleteRejected = true; }
    assert(missingDeleteRejected, 'Backend rejects deleting an already removed event');
    eventId = event.id; render();
    return {passed:checks.length,checks};
})()
