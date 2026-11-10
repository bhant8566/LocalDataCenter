(async () => {
    const checks=[];
    const assert=(v,m)=>{if(!v) throw Error(m); checks.push(m);};
    async function until(test) { for(let i=0;i<600;i++){if(test())return;await new Promise(r=>setTimeout(r,100));}throw Error('Feature test timeout'); }
    const original={projectId,sectionId,eventId};
    const originalConfirm=window.confirm, originalAlert=window.alert;
    const alerts=[]; window.alert=m=>alerts.push(m);
    const uploadFile=async(name,bytes)=>({id:(await api('/api/upload',{method:'POST',headers:{'X-Local-Token':token},body:bytes})).id,name});
    try {
        const project=(await save({kind:'project',name:'Feature project'})).id;
        const a=(await save({kind:'section',projectId:project,name:'Source'})).id;
        const b=(await save({kind:'section',projectId:project,name:'Target'})).id;
        const foreign=(await save({kind:'project',name:'Foreign project'})).id;
        const foreignSection=(await save({kind:'section',projectId:foreign,name:'Other'})).id;
        const file=await uploadFile('Unique.txt',new TextEncoder().encode('unique'));
        const shared=await uploadFile('Shared.txt',new TextEncoder().encode('shared'));
        const scheduleFile=await uploadFile('Schedule.txt',new TextEncoder().encode('schedule'));
        const id=(await save({kind:'event',sectionId:a,title:'Search NEEDLE',date:'2026-11-09',html:'<p>body-token</p>',files:[file,shared],schedules:[{name:'Schedule title',start:'2026-11-01',end:'2026-11-30',notes:'schedule-token',files:[scheduleFile]}]})).id;
        await save({kind:'event',sectionId:b,title:'Older',date:'2026-11-01',html:'',files:[shared],schedules:[]});
        projectId=project;sectionId=a;eventId=id;await refresh();
        for(const term of ['needle','BODY-token','schedule-token','schedule.txt']) {
            $('event-search').value=term;$('event-search').dispatchEvent(new Event('input'));
            assert(document.querySelectorAll('.event').length===1,'Search matches '+term);
        }
        $('event-search').value='not-present';render();assert(!document.querySelector('.event'),'Search hides nonmatches');
        $('event-search').value='needle';$('event-filter').value='completed';render();assert(!document.querySelector('.event'),'Search combines with status');
        $('event-search').value='';$('event-filter').value='all';render();

        openEditor(data.Events.find(e=>e.id===id));
        $('edit-files').querySelector('.file button:last-child').click();$('edit').close();
        assert((await fetch('/attachment/'+file.id)).ok,'Canceled attachment removal retains physical file');
        const event=data.Events.find(e=>e.id===id);
        await save({...event,kind:'event',files:[shared]});
        assert(!(await fetch('/attachment/'+file.id)).ok,'Saving attachment removal deletes physical file');
        await saveSchedule(id,{...event.schedules[0],files:[]});
        assert(!(await fetch('/attachment/'+scheduleFile.id)).ok,'Saving schedule removal deletes its physical file');
        await refresh();

        openEditor(data.Events.find(e=>e.id===id));
        editor.innerHTML='<table><tbody><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table>';
        const table=editor.querySelector('table');cellStart=table.rows[0].cells[0];cellEnd=table.rows[1].cells[1];
        $('merge-cells').click();assert(table.rows[0].cells[0].rowSpan===2 && table.rows[0].cells[0].colSpan===2,'Rectangular merge spans rows and columns');
        assert(table.textContent==='ABCD','Merge preserves reading order');
        $('split-cells').click();assert(table.rows[0].cells.length===2 && table.rows[1].cells.length===2,'Split restores merged grid');
        assert(table.rows[0].cells[0].textContent==='ABCD' && !table.rows[1].cells[1].textContent,'Split keeps content in top left');
        cellStart=cellEnd=table.rows[0].cells[0];
        const range=document.createRange();range.selectNodeContents(cellStart);range.collapse(true);window.getSelection().removeAllRanges();window.getSelection().addRange(range);
        const clip={getData:type=>type==='text/html'?'<style>.xl1{background-color:#ffff00;color:#ff0000;text-align:center}</style><table><tr><td class="xl1" colspan="2" onclick="alert(1)">Excel</td></tr><tr><td>2</td><td>3</td></tr></table>':''};
        window.confirm=()=>false;const before=table.innerHTML;pasteTable(clip);assert(table.innerHTML===before,'Cancel paste leaves table unchanged');
        window.confirm=()=>true;pasteTable(clip);
        assert(table.rows[0].cells[0].colSpan===2 && table.rows[0].cells[0].style.backgroundColor==='rgb(255, 255, 0)','Excel class styles and merged cells survive paste');
        assert(!table.querySelector('[onclick]'),'Paste removes executable attributes');
        const savedHTML=cleanHTML(editor.innerHTML);
        await save({...data.Events.find(e=>e.id===id),kind:'event',html:savedHTML});$('edit').close();await refresh();
        openEditor(data.Events.find(e=>e.id===id));assert(editor.querySelector('td').colSpan===2 && editor.querySelector('td').style.textAlign==='center','Reload preserves table spans and formatting');
        editor.innerHTML='';cellStart=cellEnd=undefined;selection=undefined;window.getSelection().removeAllRanges();
        pasteTable({getData:type=>type==='text/plain'?'one\ttwo\nthree\tfour':''});
        assert(editor.querySelectorAll('tr').length===2,'TSV paste creates a table');
        let targetTable=editor.querySelector('table');cellStart=cellEnd=targetTable.rows[1].cells[1];
        range.selectNodeContents(cellStart);range.collapse(true);getSelection().removeAllRanges();getSelection().addRange(range);
        pasteTable({getData:type=>type==='text/plain'?'A\tB\nC\tD':''});
        assert(targetTable.rows.length===3 && targetTable.rows[2].cells.length===3 && targetTable.rows[2].cells[2].textContent==='D','Paste expands existing rows and columns');
        editor.innerHTML='<table><tr><td>Start</td><td rowspan="2">Crossing</td></tr><tr><td>Keep</td></tr></table>';
        targetTable=editor.querySelector('table');cellStart=cellEnd=targetTable.rows[0].cells[0];
        range.selectNodeContents(cellStart);range.collapse(true);getSelection().removeAllRanges();getSelection().addRange(range);
        pasteTable({getData:type=>type==='text/plain'?'Left\tRight':''});
        assert(targetTable.rows[0].cells[1].rowSpan===1 && targetTable.rows[1].cells.length===2 && targetTable.rows[1].cells[0].textContent==='Keep','Confirmed paste splits crossing merge and preserves outside cells');
        $('edit').close();

        const workbook=XLSX.utils.book_new();
        const sheet=XLSX.utils.aoa_to_sheet([['Merged',null],[123.5,new Date(2026,10,10)]]);sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:1}}];sheet.A2.z='0.00';
        XLSX.utils.book_append_sheet(workbook,sheet,'First');XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([['Second tab']]),'Second');
        for(const ext of ['xlsx','xls']) {
            const bytes=XLSX.write(workbook,{bookType:ext==='xls'?'biff8':'xlsx',type:'array'});
            await openPreview(await uploadFile('Workbook.'+ext,bytes));
            assert($('preview-content').querySelector('td[colspan="2"]') && $('preview-content').textContent.includes('123.50'),'Excel '+ext+' previews merged and formatted cells');
            const select=$('preview-content').querySelector('select');select.value='Second';select.dispatchEvent(new Event('change'));
            assert($('preview-content').textContent.includes('Second tab'),'Excel '+ext+' switches worksheets');$('preview').close();disposePreview();
        }
        await openPreview(await uploadFile('broken.xlsx',new Uint8Array([0,1,2,3])));
        // Some spreadsheet parsers accept arbitrary bytes as text: invalid workbooks must never crash UI.
        assert($('preview-download').getAttribute('href'),'Invalid Excel retains original download');$('preview').close();disposePreview();

        let rejected=false;try{await postAction('/api/move-event',{id,sectionId:foreignSection});}catch{rejected=true;}
        assert(rejected,'Backend rejects cross-project movement');
        await refresh();
        const drag=new DataTransfer();const handle=document.querySelector('.event-drag');handle.dispatchEvent(new DragEvent('dragstart',{bubbles:true,dataTransfer:drag}));
        $('sections').children[1].dispatchEvent(new DragEvent('drop',{bubbles:true,dataTransfer:drag}));
        await until(()=>sectionId===b && data.Events.find(e=>e.id===id).sectionId===b);
        assert([...document.querySelectorAll('.event summary strong')][0].textContent==='Search NEEDLE','Drop moves event and sorts target by newest date');
        const moved=data.Events.find(e=>e.id===id);assert(moved.html===savedHTML && moved.schedules[0].notes==='schedule-token','Move retains content and schedules');
        await postAction('/api/delete-event',{id});
        assert((await fetch('/attachment/'+shared.id)).ok,'Shared attachment survives deletion of one owner');
        await refresh();
        showName('section',data.Sections.find(s=>s.id===b));window.confirm=()=>false;$('delete-container').click();assert($('names').open,'Cancel container deletion keeps dialog open');
        window.confirm=()=>true;$('delete-container').click();await until(()=>!$('names').open && !data.Sections.some(s=>s.id===b));
        assert(!(await fetch('/attachment/'+shared.id)).ok,'Deleting last owning section removes shared file');
        assert(!data.Events.some(e=>e.sectionId===b),'Section deletion removes contained events');
        await postAction('/api/delete-project',{id:project});await postAction('/api/delete-project',{id:foreign});await refresh();
        assert(!data.Sections.some(s=>s.projectId===project),'Project deletion removes all contained sections');
        return {passed:checks.length,checks};
    } finally {
        window.confirm=originalConfirm;window.alert=originalAlert;
        for(const dialog of document.querySelectorAll('dialog[open]')) dialog.close();disposePreview();
        projectId=original.projectId;sectionId=original.sectionId;eventId=original.eventId;$('event-search').value='';$('event-filter').value='all';await refresh();
    }
})()
