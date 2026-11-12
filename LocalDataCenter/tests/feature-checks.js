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

        let closePrompts=0;
        window.confirm=()=>{closePrompts++;return false;};
        openEditor(data.Events.find(e=>e.id===id));
        $('edit').querySelector('[data-close]').click();
        assert(!$('edit').open && closePrompts===0,'Unchanged editor closes without warning');
        openEditor(data.Events.find(e=>e.id===id));
        $('event-title').value+=' changed';
        $('edit').querySelector('[data-close]').click();
        assert($('edit').open && closePrompts===1,'Cancel discard warning keeps title edits open');
        $('edit').dispatchEvent(new Event('cancel',{cancelable:true}));
        assert($('edit').open && closePrompts===2,'Escape uses the same unsaved-change confirmation');
        window.confirm=()=>true;$('edit').querySelector('[data-close]').click();
        assert(!$('edit').open,'Confirm discard closes editor');
        for(const change of [()=>{editor.innerHTML+='<p>new content</p>';},()=>{draft.files.push({id:'draft',name:'draft'});},()=>addScheduleRow()]) {
            openEditor(data.Events.find(e=>e.id===id));change();
            window.confirm=()=>false;$('edit').querySelector('[data-close]').click();
            assert($('edit').open,'Unsaved body, attachment or schedule changes prevent closing when canceled');
            $('edit').close();
        }
        window.confirm=()=>true;
        openEditor(data.Events.find(e=>e.id===id));
        editor.innerHTML='<table><tbody><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table>';
        const table=editor.querySelector('table');table.scrollIntoView({block:'center'});
        const center=cell=>{const rect=cell.getBoundingClientRect();return {clientX:rect.left+rect.width/2,clientY:rect.top+rect.height/2};};
        const mouse=(target,type,options={})=>target.dispatchEvent(new PointerEvent(type,{bubbles:true,cancelable:true,pointerId:1,pointerType:'mouse',...options}));
        const rightClick=cell=>cell.dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,cancelable:true,button:2,...center(cell)}));
        mouse(table.rows[0].cells[0],'pointerdown',{button:0,buttons:1,...center(table.rows[0].cells[0])});
        assert(!selectingCells && !editor.querySelector('.cell-selected'),'Normal left click edits text without selecting table cells');
        mouse(table.rows[0].cells[0],'pointerdown',{button:0,buttons:1,shiftKey:true,...center(table.rows[0].cells[0])});
        mouse(document,'pointermove',{buttons:1,shiftKey:true,...center(table.rows[1].cells[1])});
        mouse(document,'pointerup',{button:0,buttons:0});
        assert(editor.querySelectorAll('.cell-selected').length===4 && !selectingCells,'Shift drag from cell center selects a full rectangle and ends on release');
        rightClick(table.rows[1].cells[1]);
        assert(!$('table-menu').hidden && editor.querySelectorAll('.cell-selected').length===4 && !$('merge-cells').disabled,'Right click retains selected rectangle and offers merge');
        assert(!document.querySelector('.toolbar #merge-cells') && !document.querySelector('.toolbar #split-cells'),'Merge and split are only in the context menu');
        $('merge-cells').click();assert(table.rows[0].cells[0].rowSpan===2 && table.rows[0].cells[0].colSpan===2,'Rectangular merge spans rows and columns');
        assert($('table-menu').hidden,'Selecting context action closes the menu');
        assert(table.textContent==='ABCD','Merge preserves reading order');
        rightClick(table.rows[0].cells[0]);
        assert($('merge-cells').disabled && !$('split-cells').disabled,'Merged cell offers split and disables invalid merge');
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
        editor.innerHTML='<p>Color selection</p>';
        const textRange=document.createRange();textRange.selectNodeContents(editor.querySelector('p'));
        getSelection().removeAllRanges();getSelection().addRange(textRange);selection=textRange.cloneRange();clearCellSelection();
        $('color').click();assert(!$('color-palette').hidden && $('color-palette').children.length===30,'Fixed palette provides standard, pastel and grayscale colors');
        $('color-palette').querySelector('[data-color="#ed1c24"]').click();
        assert($('color-palette').hidden && /rgb\(237, 28, 36\)|#ed1c24/i.test(editor.innerHTML),'Palette colors selected text and closes without losing the selection');
        editor.innerHTML='<table><tr><td style="color:blue"><span style="color:red">One</span></td><td>Two</td></tr></table>';
        cellStart=editor.querySelector('td');cellEnd=editor.querySelectorAll('td')[1];highlightCells();
        $('color').click();$('color-palette').querySelector('[data-color="#22b14c"]').click();
        assert([...editor.querySelectorAll('td')].every(cell=>cell.style.color==='rgb(34, 177, 76)') && !editor.querySelector('span').style.color,'Palette applies one color to all selected cells');
        editor.innerHTML='<table><tr><td>A</td><td rowspan="2">B</td></tr><tr><td>C</td></tr></table>';
        let insertion=editor.querySelector('table');cellStart=cellEnd=insertion.rows[0].cells[0];highlightCells();
        $('insert-row').click();
        assert(insertion.rows.length===3 && insertion.rows[0].cells[1].rowSpan===3 && insertion.rows[2].cells[0].textContent==='C','Row inserts below selected row and extends crossing rowspan');
        editor.innerHTML='<table><tr><td>A</td><td>B</td></tr><tr><td colspan="2">C</td></tr></table>';
        insertion=editor.querySelector('table');cellStart=cellEnd=insertion.rows[0].cells[0];highlightCells();
        $('insert-column').click();
        assert(insertion.rows[0].cells.length===3 && insertion.rows[0].cells[2].textContent==='B' && insertion.rows[1].cells[0].colSpan===3,'Column inserts to right and extends crossing colspan');
        cellStart=insertion.rows[0].cells[0];cellEnd=insertion.rows[0].cells[1];highlightCells();
        $('cell-background').value='#fff2cc';$('cell-background').dispatchEvent(new Event('input'));
        assert(insertion.rows[0].cells[0].style.backgroundColor==='rgb(255, 242, 204)' && insertion.rows[0].cells[1].style.backgroundColor==='rgb(255, 242, 204)' && !insertion.rows[0].cells[2].style.backgroundColor,'Cell background applies only to selected cells');
        assert(cleanHTML(editor.innerHTML).includes('background-color'),'Cell background survives save sanitization');
        $('clear-cell-background').click();assert(!insertion.rows[0].cells[0].style.backgroundColor,'Clear background removes selected cell fill');
        cellStart=cellEnd=insertion.rows[1].cells[0];highlightCells();$('insert-row').click();
        assert(insertion.rows.length===3,'Row insertion works at end of merged table');
        const wideHTML='<table><tr>'+Array.from({length:18},(_,i)=>'<td>Long column value '+i+'</td>').join('')+'</tr></table>';
        editor.innerHTML=wideHTML;clearCellSelection();
        assert(editor.scrollWidth>editor.clientWidth && getComputedStyle(editor.querySelector('td')).whiteSpace==='pre','Editor scrolls wide tables without forced wrapping');
        await save({...data.Events.find(e=>e.id===id),kind:'event',html:wideHTML});
        $('edit').close();
        await refresh();eventId=id;render();
        const content=document.querySelector('.event[open] .event-content');
        assert(content.scrollWidth>content.clientWidth && document.documentElement.scrollWidth<=innerWidth,'Event article scrolls its own wide table without expanding page: '+JSON.stringify({content:content.clientWidth,table:content.scrollWidth,page:document.documentElement.scrollWidth,viewport:innerWidth}));
        const pan=(root,scroller)=>{
            scroller.scrollLeft=0;
            const rect=root.getBoundingClientRect(),x=rect.left+rect.width/2,y=rect.top+rect.height/2;
            mouse(root,'pointerdown',{button:0,buttons:1,clientX:x,clientY:y});
            root.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,cancelable:true,button:2,buttons:3,clientX:x,clientY:y}));
            mouse(document,'pointermove',{buttons:3,clientX:x-150,clientY:y});
            const moved=scroller.scrollLeft>0;
            const context=new MouseEvent('contextmenu',{bubbles:true,cancelable:true,button:2});root.dispatchEvent(context);
            document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true,button:2,buttons:1}));
            const released=!root.classList.contains('horizontal-panning');
            mouse(document,'pointermove',{buttons:1,clientX:x-250,clientY:y});
            return moved && context.defaultPrevented && released && scroller.scrollLeft===150;
        };
        assert(pan(content,content),'Both-button drag pans event content, suppresses context menu, and stops when either button is released');
        await save({...data.Events.find(e=>e.id===id),kind:'event',html:savedHTML});await refresh();

        const workbook=XLSX.utils.book_new();
        const sheet=XLSX.utils.aoa_to_sheet([['Merged',null],[123.5,new Date(2026,10,10)]]);sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:1}}];sheet.A2.z='0.00';
        XLSX.utils.book_append_sheet(workbook,sheet,'First');XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([['Second tab',...Array.from({length:18},(_,i)=>'Wide column '+i)]]),'Second');
        XLSX.utils.book_append_sheet(workbook,{},'Empty');
        for(const ext of ['xlsx','xls']) {
            const bytes=XLSX.write(workbook,{bookType:ext==='xls'?'biff8':'xlsx',type:'array'});
            await openPreview(await uploadFile('Workbook.'+ext,bytes));
            assert($('preview-content').querySelector('td[colspan="2"]') && $('preview-content').textContent.includes('123.50'),'Excel '+ext+' previews merged and formatted cells');
            const tabs=$('preview-content').querySelectorAll('[role=tab]');assert(tabs.length===3,'Excel '+ext+' lists all sheet names as tabs');tabs[1].click();
            assert($('preview-content').textContent.includes('Second tab') && tabs[1].getAttribute('aria-selected')==='true','Excel '+ext+' switches worksheets');
            const scroller=$('preview-content').querySelector('.workbook-sheet');
            assert(scroller.scrollWidth>scroller.clientWidth && getComputedStyle(scroller.querySelector('td')).whiteSpace==='pre','Excel '+ext+' wide sheet keeps column content unwrapped');
            assert(pan($('preview'),scroller),'Excel '+ext+' can pan from preview area with both mouse buttons');
            mouse(tabs[2],'pointerdown',{button:0,buttons:1,...center(tabs[2])});
            tabs[2].click();assert(scroller.textContent.includes('空白工作表'),'Empty worksheet can be selected');
            tabs[0].click();assert(scroller.querySelector('[colspan="2"]') && scroller.scrollLeft===0,'Switching back restores merged sheet and resets horizontal scroll');
            $('preview').close();disposePreview();
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
