'use strict';

function eventSearchText(event) {
    const body = document.createElement('template');
    body.innerHTML = cleanHTML(event.html || '');
    return [event.title, body.content.textContent, ...(event.files || []).map(f => f.name),
        ...eventSchedules(event).flatMap(s => [s.name, s.notes, ...(s.files || []).map(f => f.name)])].join('\n').toLocaleLowerCase();
}
$('event-search').oninput = () => render();

function postAction(path, payload) {
    return api(path, {method:'POST', headers:{'X-Local-Token':token,'Content-Type':'application/json'}, body:JSON.stringify(payload)});
}

let draggedEvent;
function addEventDragHandle(summary, event, writable) {
    const handle = element('button', '⠿', 'event-drag');
    handle.type = 'button'; handle.draggable = writable && data.capabilities?.includes('event-move');
    handle.disabled = !handle.draggable;
    handle.title = '拖曳到同專案的其他分頁';
    handle.setAttribute('aria-label', '搬移 ' + event.title);
    handle.onclick = e => { e.preventDefault(); e.stopPropagation(); };
    handle.ondragstart = e => {
        draggedEvent = event.id;
        e.dataTransfer.setData('application/x-local-event', event.id);
        e.dataTransfer.effectAllowed = 'move';
        e.stopPropagation();
    };
    handle.ondragend = () => { draggedEvent = undefined; document.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target')); };
    summary.prepend(handle);
}
$('sections').addEventListener('dragover', e => {
    const button = e.target.closest('button');
    if (!button || !draggedEvent) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'move'; button.classList.add('drop-target');
});
$('sections').addEventListener('dragleave', e => e.target.closest('button')?.classList.remove('drop-target'));
$('sections').addEventListener('drop', async e => {
    const button = e.target.closest('button');
    if (!button || !draggedEvent) return;
    e.preventDefault(); button.classList.remove('drop-target');
    const sections = data.Sections.filter(s => s.projectId === projectId);
    const target = sections[[...$('sections').children].indexOf(button)];
    const event = data.Events.find(v => v.id === draggedEvent);
    draggedEvent = undefined;
    if (!target || !event || event.sectionId === target.id) return;
    button.disabled = true;
    try {
        await postAction('/api/move-event', {id:event.id, sectionId:target.id});
        sectionId = target.id; eventId = event.id;
        $('event-filter').value = 'all'; $('event-search').value = '';
        await refresh(); notice('事件已搬移至「' + target.name + '」');
    } catch (error) { alert(error.message); }
    finally { button.disabled = false; }
});

function deletionImpact(kind, id) {
    const sections = kind === 'project' ? data.Sections.filter(s => s.projectId === id) : data.Sections.filter(s => s.id === id);
    const events = data.Events.filter(e => sections.some(s => s.id === e.sectionId));
    const ids = new Set();
    for (const event of events) {
        for (const file of [...(event.files || []), ...eventSchedules(event).flatMap(s => s.files || [])]) ids.add(file.id);
        for (const match of (event.html || '').matchAll(/\/attachment\/([a-f0-9]{32}\.(?:png|jpg|gif|webp|bin))/g)) ids.add(match[1]);
    }
    return `${sections.length} 個分頁、${events.length} 筆事件、${events.reduce((n,e) => n + eventSchedules(e).length,0)} 條時程、${ids.size} 個附件（共用原檔保留）`;
}
$('delete-container').onclick = async () => {
    const request = {...nameRequest};
    if (!request.id || !data.capabilities?.includes('container-delete')) return;
    const item = data[request.kind === 'project' ? 'Projects' : 'Sections'].find(i => i.id === request.id);
    if (!confirm('確定移除' + (request.kind === 'project' ? '專案' : '分頁') + '「' + item.name + '」？\n影響：' + deletionImpact(request.kind,request.id) + '\n資料及無其他引用的附件原檔將一併刪除，無法透過 Excel 備份還原附件。')) return;
    const controls = [...$('name-form').querySelectorAll('button')];
    controls.forEach(b => b.disabled = true);
    try {
        await postAction('/api/delete-' + request.kind, {id:request.id});
        $('names').close(); eventId = undefined;
        await refresh(); notice('已移除');
    } catch (error) { alert(error.message); }
    finally { controls.forEach(b => b.disabled = false); }
};

function renderWorkbookPreview(bytes) {
    const workbook = XLSX.read(bytes, {type:'array', cellText:true, cellHTML:false, bookVBA:false});
    const tabs = element('div',undefined,'workbook-tabs');
    tabs.setAttribute('role','tablist'); tabs.setAttribute('aria-label','工作表');
    const container = element('div',undefined,'workbook-sheet');
    container.id='workbook-panel'; container.setAttribute('role','tabpanel'); container.tabIndex=0;
    container.setAttribute('aria-label','工作表內容，可水平捲動');
    let activeSheet=workbook.SheetNames[0];
    workbook.SheetNames.forEach((name,index)=>{
        const button=element('button',name); button.type='button'; button.id='workbook-tab-'+index;
        button.dataset.sheet=name; button.setAttribute('role','tab'); button.setAttribute('aria-controls',container.id);
        button.onclick=()=>{activeSheet=name;draw();};
        button.onkeydown=e=>{
            const offsets={ArrowRight:1,ArrowLeft:-1,Home:-index,End:workbook.SheetNames.length-1-index};
            if (!(e.key in offsets)) return;
            e.preventDefault();
            const next=tabs.children[(index+offsets[e.key]+workbook.SheetNames.length)%workbook.SheetNames.length]; next.click(); next.focus();
        };
        tabs.append(button);
    });
    const draw = () => {
        for (const button of tabs.children) {
            const selected=button.dataset.sheet===activeSheet;
            button.setAttribute('aria-selected',String(selected)); button.tabIndex=selected?0:-1;
            if(selected) container.setAttribute('aria-labelledby',button.id);
        }
        const sheet = workbook.Sheets[activeSheet];
        container.replaceChildren();
        container.scrollLeft=container.scrollTop=0;
        if (!sheet?.['!ref'] || !Object.keys(sheet).some(key=>/^[A-Z]+[1-9]\d*$/.test(key) && sheet[key].t!=='z')) { container.append(element('p','空白工作表')); return; }
        const range = XLSX.utils.decode_range(sheet['!ref']);
        if ((range.e.r + 1) * (range.e.c + 1) > 200000) { container.append(element('p','工作表過大，請下載原檔查看。')); return; }
        const table = element('table');
        const header = element('tr'); header.append(element('th',''));
        for (let c = 0; c <= range.e.c; c++) header.append(element('th',XLSX.utils.encode_col(c)));
        table.append(header);
        const merged = new Map(), covered = new Set();
        for (const merge of sheet['!merges'] || []) {
            if ((merge.e.r-merge.s.r+1)*(merge.e.c-merge.s.c+1) > 200000 || merge.e.r > range.e.r || merge.e.c > range.e.c) continue;
            merged.set(merge.s.r + ':' + merge.s.c, merge);
            for (let r=merge.s.r;r<=merge.e.r;r++) for (let c=merge.s.c;c<=merge.e.c;c++) if (r!==merge.s.r || c!==merge.s.c) covered.add(r+':'+c);
        }
        for (let r=0;r<=range.e.r;r++) {
            const row = element('tr'); row.append(element('th',String(r+1)));
            for (let c=0;c<=range.e.c;c++) {
                const key = r+':'+c; if (covered.has(key)) continue;
                const cell = sheet[XLSX.utils.encode_cell({r,c})];
                const td = element('td', cell ? (cell.w ?? XLSX.utils.format_cell(cell)) : '');
                const merge = merged.get(key);
                if (merge) { td.rowSpan=merge.e.r-merge.s.r+1; td.colSpan=merge.e.c-merge.s.c+1; }
                row.append(td);
            }
            table.append(row);
        }
        container.append(table);
    };
    const hint=element('p','同時按住滑鼠左右鍵，再左右移動可拖移表格；亦可使用下方水平捲動條。','table-pan-hint');
    $('preview-content').append(tabs,hint,container); draw();
}

// Listen for both pointer and mouse events: pressing the second button does not
// fire another pointerdown, and releasing just one button is not a pointerup.
function enableHorizontalPan(root, getScroller) {
    let pan, suppressMenu=false, suppressClick=false;
    const end=()=>{
        if(!pan)return;
        root.classList.remove('horizontal-panning'); pan.scroller.classList.remove('horizontal-panning'); pan=undefined;
        document.removeEventListener('pointermove',move);
        document.removeEventListener('mouseup',release);
        document.removeEventListener('pointercancel',end);
        window.removeEventListener('blur',end);
    };
    const start=e=>{
        if((e.buttons&3)!==3) { if(!pan && e.type==='pointerdown') suppressMenu=suppressClick=false; return; }
        if(root===$('preview') && outsidePreview(e)) return;
        const scroller=getScroller();
        if(!scroller || scroller.scrollWidth<=scroller.clientWidth) return;
        e.preventDefault(); suppressMenu=suppressClick=true;
        if(pan)return;
        pan={scroller,x:e.clientX,left:scroller.scrollLeft};
        root.classList.add('horizontal-panning'); scroller.classList.add('horizontal-panning');
        document.addEventListener('pointermove',move,{passive:false});
        document.addEventListener('mouseup',release);
        document.addEventListener('pointercancel',end);
        window.addEventListener('blur',end);
    };
    const move=e=>{
        if(!pan)return;
        if((e.buttons&3)!==3 || !root.isConnected) {end();return;}
        e.preventDefault(); pan.scroller.scrollLeft=pan.left+pan.x-e.clientX;
    };
    const release=e=>{if((e.buttons&3)!==3)end();};
    root.addEventListener('pointerdown',start);
    root.addEventListener('mousedown',start);
    root.addEventListener('pointermove',e=>{if(!pan && (e.buttons&3)===3)start(e);});
    root.addEventListener('contextmenu',e=>{if(pan || suppressMenu){e.preventDefault();e.stopPropagation();suppressMenu=false;}},true);
    for(const type of ['click','auxclick']) root.addEventListener(type,e=>{if(suppressClick){e.preventDefault();e.stopPropagation();}},true);
    root.addEventListener('keydown',()=>{suppressMenu=suppressClick=false;});
    root.closest('dialog')?.addEventListener('close',end);
}
enableHorizontalPan($('preview'),()=>$('preview-content').querySelector('.workbook-sheet'));
enableHorizontalPan(editor,()=>editor);

function updatePreviewPage(session, number) {
    if (previewSession !== session) return;
    session.page = number;
    $('preview-page').textContent = number + ' / ' + session.pdf.numPages;
    $('preview-prev').disabled = number <= 1;
    $('preview-next').disabled = number >= session.pdf.numPages;
}
$('preview-content').addEventListener('scroll', () => {
    const session = previewSession;
    if (!session?.pages) return;
    const viewport = $('preview-content').getBoundingClientRect();
    const center = viewport.top + viewport.height / 2;
    const page = session.pages.find(p => p.getBoundingClientRect().bottom > center);
    if (page) updatePreviewPage(session,Number(page.dataset.page));
});

// Logical table coordinates include cells covered by rowspan/colspan.
function tableModel(table) {
    const records = [], grid = []; let width = 0;
    for (const [r,row] of [...table.rows].entries()) {
        grid[r] ||= []; let c = 0;
        for (const td of row.cells) {
            while (grid[r][c]) c++;
            const rs = td.rowSpan || 1, cs = td.colSpan || 1;
            if (r+rs > 1000 || c+cs > 1000 || (r+rs)*(c+cs) > 50000) throw Error('表格過大，請縮小範圍。');
            const record = {td,r,c,rs,cs}; records.push(record);
            for (let y=r;y<r+rs;y++) for (let x=c;x<c+cs;x++) {
                grid[y] ||= [];
                if (grid[y][x]) throw Error('表格合併範圍重疊，請先分割。');
                grid[y][x] = record;
            }
            c += cs; width = Math.max(width,c);
        }
    }
    return {records,grid,width,height:Math.max(table.rows.length,grid.length)};
}
function blankCell(r,c) { const td=element('td'); td.append(element('br')); return {td,r,c,rs:1,cs:1}; }
function drawTableModel(table, model) {
    if (model.width * model.height > 50000) throw Error('表格過大，請縮小範圍。');
    const body = element('tbody'), occupied = new Set();
    for (const cell of model.records) for (let r=cell.r;r<cell.r+cell.rs;r++) for (let c=cell.c;c<cell.c+cell.cs;c++) occupied.add(r+':'+c);
    for (let r=0;r<model.height;r++) {
        const row = element('tr');
        for (let c=0;c<model.width;c++) if (!occupied.has(r+':'+c)) model.records.push(blankCell(r,c));
        for (const cell of model.records.filter(cell => cell.r===r).sort((a,b)=>a.c-b.c)) {
            cell.td.rowSpan=cell.rs; cell.td.colSpan=cell.cs; cell.td.classList.remove('cell-selected'); row.append(cell.td);
        }
        body.append(row);
    }
    table.replaceChildren(body);
}
function splitRecord(model, cell) {
    model.records = model.records.filter(v => v !== cell);
    for (let r=cell.r;r<cell.r+cell.rs;r++) for (let c=cell.c;c<cell.c+cell.cs;c++) {
        const part = blankCell(r,c);
        if (r===cell.r && c===cell.c) part.td=cell.td;
        else part.td.style.cssText=cell.td.style.cssText;
        model.records.push(part);
    }
}
let cellStart, cellEnd, selectingCells = false;
let cellPointer, cellFrame, cellPosition;
function cellAt(node) { return (node?.nodeType===1 ? node : node?.parentElement)?.closest('td,th'); }
function clearCellSelection() { editor.querySelectorAll('.cell-selected').forEach(c=>c.classList.remove('cell-selected')); }
function selectedRectangle() {
    if (!cellStart?.isConnected || !cellEnd?.isConnected || cellStart.closest('table')!==cellEnd.closest('table')) throw Error('請先選取同一表格的儲存格。');
    const table=cellStart.closest('table'), model=tableModel(table);
    const a=model.records.find(v=>v.td===cellStart), b=model.records.find(v=>v.td===cellEnd);
    const bounds={r:Math.min(a.r,b.r),c:Math.min(a.c,b.c),bottom:Math.max(a.r+a.rs,b.r+b.rs),right:Math.max(a.c+a.cs,b.c+b.cs)};
    let cells, expanded;
    do {
        expanded=false;
        cells=model.records.filter(v=>v.r<bounds.bottom && v.r+v.rs>bounds.r && v.c<bounds.right && v.c+v.cs>bounds.c);
        for(const cell of cells) {
            if(cell.r<bounds.r || cell.c<bounds.c || cell.r+cell.rs>bounds.bottom || cell.c+cell.cs>bounds.right) expanded=true;
            bounds.r=Math.min(bounds.r,cell.r); bounds.c=Math.min(bounds.c,cell.c);
            bounds.bottom=Math.max(bounds.bottom,cell.r+cell.rs); bounds.right=Math.max(bounds.right,cell.c+cell.cs);
        }
    } while(expanded);
    return {table,model,bounds,cells};
}
function highlightCells() {
    clearCellSelection();
    try { selectedRectangle().cells.forEach(v=>v.td.classList.add('cell-selected')); } catch {}
}
editor.addEventListener('pointerdown', e => {
    if(e.button!==0 || (e.buttons&3)===3) return;
    closeTableMenu();
    const cell=cellAt(e.target);
    if (!cell || !editor.contains(cell)) { cellStart=cellEnd=undefined; clearCellSelection(); return; }
    cellStart=cellEnd=cell; clearCellSelection();
    if (!e.shiftKey) return;
    e.preventDefault(); selection=undefined; getSelection().removeAllRanges();
    selectingCells=true; cellPointer=e.pointerId; cellPosition={x:e.clientX,y:e.clientY};
    editor.classList.add('selecting-cells'); highlightCells();
    if(e.isTrusted) editor.setPointerCapture(e.pointerId);
    cellFrame=requestAnimationFrame(trackCellSelection);
});
function trackCellSelection() {
    if(!selectingCells || !cellStart?.isConnected) {endCellSelection();return;}
    const rect=editor.getBoundingClientRect();
    const dx=cellPosition.x<rect.left+24?-14:cellPosition.x>rect.right-24?14:0;
    const dy=cellPosition.y<rect.top+20?-10:cellPosition.y>rect.bottom-20?10:0;
    editor.scrollLeft+=dx; editor.scrollTop+=dy;
    const hit=document.elementFromPoint(Math.max(rect.left+4,Math.min(rect.right-4,cellPosition.x)),Math.max(rect.top+4,Math.min(rect.bottom-4,cellPosition.y)));
    const cell=cellAt(hit);
    if(cell && cell.closest('table')===cellStart.closest('table') && cell!==cellEnd) {cellEnd=cell;highlightCells();}
    cellFrame=requestAnimationFrame(trackCellSelection);
}
function endCellSelection() {
    selectingCells=false; cancelAnimationFrame(cellFrame); editor.classList.remove('selecting-cells');
    if(cellPointer!==undefined && editor.hasPointerCapture(cellPointer)) editor.releasePointerCapture(cellPointer);
    cellPointer=undefined;
}
document.addEventListener('pointermove', e => {
    if(!selectingCells)return;
    if(!(e.buttons&1) || (e.buttons&3)===3) {endCellSelection();return;}
    e.preventDefault(); cellPosition={x:e.clientX,y:e.clientY};
    const cell=cellAt(document.elementFromPoint(e.clientX,e.clientY));
    if(cell && cell.closest('table')===cellStart?.closest('table') && cell!==cellEnd) {cellEnd=cell;highlightCells();}
},{passive:false});
document.addEventListener('pointerup',endCellSelection);
document.addEventListener('pointercancel',endCellSelection);
editor.addEventListener('lostpointercapture',endCellSelection);
window.addEventListener('blur',endCellSelection);
$('edit').addEventListener('close',()=>{endCellSelection();closeTableMenu();closeColorPalette();});

function closeTableMenu() { $('table-menu').hidden=true; }
editor.addEventListener('contextmenu',e=>{
    if(e.defaultPrevented || (e.buttons&3)===3)return;
    const cell=cellAt(e.target);
    if(!cell || !editor.contains(cell)) {closeTableMenu();return;}
    e.preventDefault(); endCellSelection();
    if(!cell.classList.contains('cell-selected')) {cellStart=cellEnd=cell;highlightCells();}
    const {cells,bounds:b}=selectedRectangle();
    $('merge-cells').disabled=cells.length<2 || cells.reduce((sum,v)=>sum+v.rs*v.cs,0)!==(b.bottom-b.r)*(b.right-b.c);
    $('split-cells').disabled=!cells.some(v=>v.rs>1 || v.cs>1);
    const menu=$('table-menu');menu.hidden=false;
    const rect=cell.getBoundingClientRect();
    menu.style.left=Math.max(8,Math.min(e.clientX || rect.left,innerWidth-menu.offsetWidth-8))+'px';
    menu.style.top=Math.max(8,Math.min(e.clientY || rect.bottom,innerHeight-menu.offsetHeight-8))+'px';
    menu.querySelector('button:not(:disabled)')?.focus({preventScroll:true});
});
document.addEventListener('pointerdown',e=>{if(!$('table-menu').contains(e.target))closeTableMenu();},true);
for(const target of [editor,$('edit'),window]) target.addEventListener('scroll',closeTableMenu,{passive:true});
$('table-menu').addEventListener('keydown',e=>{
    if(e.key==='Escape') {e.preventDefault();e.stopPropagation();closeTableMenu();editor.focus({preventScroll:true});return;}
    if(!['ArrowDown','ArrowUp','Home','End'].includes(e.key))return;
    e.preventDefault();const buttons=[...$('table-menu').querySelectorAll('button:not(:disabled)')];
    const i=buttons.indexOf(document.activeElement);
    const next=e.key==='Home'?0:e.key==='End'?buttons.length-1:(i+(e.key==='ArrowDown'?1:-1)+buttons.length)%buttons.length;
    buttons[next]?.focus();
});
$('merge-cells').onclick = () => {
    closeTableMenu();
    try {
        const {table,model,bounds:b,cells}=selectedRectangle();
        if (cells.length<2) throw Error('請選取兩個以上儲存格。');
        if (cells.some(v=>v.r<b.r || v.c<b.c || v.r+v.rs>b.bottom || v.c+v.cs>b.right) || cells.reduce((n,v)=>n+v.rs*v.cs,0)!==(b.bottom-b.r)*(b.right-b.c)) throw Error('請選取完整矩形，並完整包含既有合併格。');
        cells.sort((a,b)=>a.r-b.r || a.c-b.c);
        const td=cells[0].td, fragments=cells.map(v=>v.td.innerHTML);
        td.innerHTML=fragments.join('<br>');
        model.records=model.records.filter(v=>!cells.includes(v));
        model.records.push({td,r:b.r,c:b.c,rs:b.bottom-b.r,cs:b.right-b.c});
        drawTableModel(table,model); cellStart=cellEnd=td; highlightCells();
    } catch (error) { notice(error.message); }
};
$('split-cells').onclick = () => {
    closeTableMenu();
    try {
        const {table,model,cells}=selectedRectangle();
        cells.filter(v=>v.rs>1 || v.cs>1).forEach(v=>splitRecord(model,v));
        drawTableModel(table,model); clearCellSelection();
    } catch (error) { notice(error.message); }
};

function insertTableAxis(axis) {
    closeTableMenu();
    try {
        const {table,model,bounds}=selectedRectangle();
        const row=axis==='row', coordinate=row?'r':'c', span=row?'rs':'cs';
        const boundary=row?bounds.bottom:bounds.right;
        const height=model.height+(row?1:0), width=model.width+(row?0:1);
        if(height>1000 || width>1000 || height*width>50000)throw Error('表格過大，無法繼續新增。');
        for(const cell of model.records) {
            if(cell[coordinate]>=boundary)cell[coordinate]++;
            else if(cell[coordinate]+cell[span]>boundary)cell[span]++;
        }
        model.height=height;model.width=width;
        drawTableModel(table,model);
        cellStart=cellEnd=undefined;selection=undefined;clearCellSelection();
    } catch(error) {notice(error.message);}
}
$('insert-row').onclick=()=>insertTableAxis('row');
$('insert-column').onclick=()=>insertTableAxis('column');
function applyCellBackground(color) {
    try {
        if(!editor.querySelector('.cell-selected')) {
            const cell=cellAt(selection?.startContainer);
            if(cell && editor.contains(cell))cellStart=cellEnd=cell;
        }
        for(const {td} of selectedRectangle().cells) {
            if(color)td.style.backgroundColor=color;
            else td.style.removeProperty('background-color');
        }
    } catch {notice('請先點選或框選要變更背景的儲存格。');}
}
$('cell-background').oninput=event=>applyCellBackground(event.target.value);
$('clear-cell-background').onpointerdown=event=>event.preventDefault();
$('clear-cell-background').onclick=()=>applyCellBackground('');

const textColors=[
    ['#000000','黑色'],['#808080','灰色'],['#99001a','深紅'],['#ed1c24','紅色'],['#ff7f27','橘色'],['#fff200','黃色'],['#22b14c','綠色'],['#00a2e8','藍綠'],['#3f48cc','藍色'],['#a349a4','紫色'],
    ['#ffffff','白色'],['#c0c0c0','淺灰'],['#b97a57','棕色'],['#ffaec9','粉紅'],['#ffc90e','金黃'],['#efe4b0','米黃'],['#b5e61d','黃綠'],['#99d9ea','淺藍綠'],['#7092be','灰藍'],['#c8bfe7','淡紫'],
    ['#f5f5f5','灰階 1'],['#e0e0e0','灰階 2'],['#cccccc','灰階 3'],['#b3b3b3','灰階 4'],['#999999','灰階 5'],['#777777','灰階 6'],['#555555','灰階 7'],['#333333','灰階 8'],['#19312d','深綠'],['#005a9c','深藍']
];
function closeColorPalette() { $('color-palette').hidden=true; $('color').setAttribute('aria-expanded','false'); }
for(const [color,name] of textColors) {
    const button=element('button');button.type='button';button.dataset.color=color;
    button.title=name+' '+color;button.setAttribute('aria-label',button.title);button.setAttribute('aria-pressed',String(color==='#000000'));
    button.style.backgroundColor=color;
    button.onpointerdown=e=>e.preventDefault();
    button.onclick=()=>{
        if(editor.querySelector('.cell-selected')) {
            for(const {td} of selectedRectangle().cells) {
                td.style.color=color;
                for(const child of td.querySelectorAll('*')) {child.style.removeProperty('color');if(child.tagName==='FONT')child.removeAttribute('color');}
            }
        } else command('foreColor',color);
        $('color').style.setProperty('--selected-color',color);
        for(const swatch of $('color-palette').children)swatch.setAttribute('aria-pressed',String(swatch===button));
        closeColorPalette();
    };
    $('color-palette').append(button);
}
$('color').onpointerdown=e=>e.preventDefault();
$('color').onclick=()=>{
    const open=$('color-palette').hidden;
    $('color-palette').hidden=!open; $('color').setAttribute('aria-expanded',String(open));
    if(open) {
        const rect=$('color').getBoundingClientRect(), palette=$('color-palette');
        palette.style.left=Math.max(8,Math.min(rect.left,innerWidth-palette.offsetWidth-8))+'px';
        palette.style.top=Math.max(8,Math.min(rect.bottom+5,innerHeight-palette.offsetHeight-8))+'px';
    }
};
document.addEventListener('pointerdown',e=>{if(!e.target.closest('.color-picker'))closeColorPalette();},true);
$('color-palette').addEventListener('keydown',e=>{
    if(e.key==='Escape') {e.preventDefault();e.stopPropagation();closeColorPalette();$('color').focus();return;}
    const offset={ArrowLeft:-1,ArrowRight:1,ArrowUp:-10,ArrowDown:10}[e.key];
    if(!offset)return;
    e.preventDefault();const buttons=[...$('color-palette').children],i=buttons.indexOf(document.activeElement);
    buttons[(i+offset+buttons.length)%buttons.length].focus();
});

function pasteTable(clipboard) {
    const html=clipboard.getData('text/html'), text=clipboard.getData('text/plain');
    const raw=document.createElement('template'); raw.innerHTML=html;
    // Desktop Excel often places formatting in .xlNN classes instead of inline styles.
    for (const style of raw.content.querySelectorAll('style')) {
        try {
            const sheet=new CSSStyleSheet(); sheet.replaceSync(style.textContent);
            for (const rule of sheet.cssRules) {
                if (!rule.selectorText || !/^[\w.\s,#>:-]+$/.test(rule.selectorText)) continue;
                for (const cell of raw.content.querySelectorAll(rule.selectorText)) {
                    for (const property of ['color','background-color','font-weight','font-size','text-align','vertical-align']) {
                        if (!cell.style.getPropertyValue(property) && rule.style.getPropertyValue(property)) cell.style.setProperty(property,rule.style.getPropertyValue(property));
                    }
                }
            }
        } catch { /* Only supported CSS declarations enter the HTML allowlist below. */ }
    }
    for (const cell of raw.content.querySelectorAll('td,th')) {
        if (!cell.style.backgroundColor && cell.getAttribute('bgcolor')) cell.style.backgroundColor=cell.getAttribute('bgcolor');
        if (/^\d+(?:\.\d+)?pt$/.test(cell.style.fontSize)) cell.style.fontSize=Math.min(99,Math.round(parseFloat(cell.style.fontSize)*4/3))+'px';
    }
    const source=element('div'); source.innerHTML=cleanHTML(raw.innerHTML);
    let incoming=source.querySelector('table');
    if (!incoming && text.includes('\t')) {
        incoming=element('table');
        for (const line of text.replace(/\r\n?/g,'\n').replace(/\n$/,'').split('\n')) {
            const row=element('tr'); for (const value of line.split('\t')) row.append(element('td',value)); incoming.append(row);
        }
    }
    if (!incoming) return false;
    try {
        if (incoming.querySelector('table')) throw Error('不支援巢狀表格，請分別貼上。');
        const from=tableModel(incoming);
        const active=window.getSelection();
        const target=cellAt(active?.anchorNode) || (cellStart?.isConnected ? cellStart : undefined);
        if (!target || !editor.contains(target)) { command('insertHTML',cleanHTML(incoming.outerHTML)); return true; }
        const table=target.closest('table'), original=tableModel(table), origin=original.records.find(v=>v.td===target);
        const clone=table.cloneNode(true), model=tableModel(clone);
        const bottom=origin.r+from.height, right=origin.c+from.width;
        if (Math.max(model.height,bottom)*Math.max(model.width,right)>50000) throw Error('表格過大，請縮小貼上範圍。');
        const affected=model.records.filter(v=>v.r<bottom && v.r+v.rs>origin.r && v.c<right && v.c+v.cs>origin.c);
        const hasContent=affected.some(v=>v.td.textContent.trim() || v.td.querySelector('img'));
        const crosses=affected.some(v=>v.r<origin.r || v.c<origin.c || v.r+v.rs>bottom || v.c+v.cs>right);
        if ((hasContent || crosses) && !confirm('貼上將覆蓋既有內容，並分割跨越貼上邊界的合併格。確定繼續？')) return true;
        for (const v of affected) if (v.rs>1 || v.cs>1) splitRecord(model,v);
        model.records=model.records.filter(v=>v.r<origin.r || v.r>=bottom || v.c<origin.c || v.c>=right);
        model.records.push(...from.records.map(v=>({...v,r:v.r+origin.r,c:v.c+origin.c})));
        model.height=Math.max(model.height,bottom); model.width=Math.max(model.width,right);
        drawTableModel(clone,model); table.replaceChildren(...clone.childNodes);
        cellStart=cellEnd=undefined; selection=undefined; clearCellSelection();
    } catch (error) { notice(error.message); }
    return true;
}

let ocrFiles = [], ocrJob, ocrOutput, ocrContext;
function clearOCRLayout() {
    $('ocr-layout').replaceChildren();$('ocr-layout').hidden=true;
    $('ocr-layout').contentEditable='false';$('ocr-text-label').hidden=false;
    $('ocr-format-label').hidden=true;$('ocr-result-format').value='layout';
    $('ocr-layout-note').hidden=true;$('ocr-layout-note').textContent='';
}
function showOCRFormat() {
    const layout=!$('ocr-format-label').hidden && $('ocr-result-format').value==='layout';
    $('ocr-layout').hidden=!layout;$('ocr-text-label').hidden=layout;
}
$('ocr-result-format').onchange=showOCRFormat;
$('ocr-layout').addEventListener('paste',event=>{
    event.preventDefault();
    if($('ocr-layout').contentEditable==='true')document.execCommand('insertText',false,event.clipboardData.getData('text/plain'));
});
$('ocr-layout').addEventListener('drop',event=>event.preventDefault());
// Group scan-line runs into physical rulings, then join intersecting rulings.
function ruledTables(runs) {
    if(runs.length>12000)return [];
    const lines=[];
    for(const run of runs.sort((a,b)=>a.axis.localeCompare(b.axis)||a.pos-b.pos)) {
        const previous=lines.findLast(line=>line.axis===run.axis && run.pos-line.max<=4 && run.start<=line.end+4 && run.end>=line.start-4);
        if(previous) {
            previous.max=run.pos;previous.pos=(previous.min+previous.max)/2;
            previous.start=Math.min(previous.start,run.start);previous.end=Math.max(previous.end,run.end);
        } else lines.push({...run,min:run.pos,max:run.pos});
    }
    if(lines.length>600)return [];
    const intersects=(a,b)=>a.axis!==b.axis && (a.axis==='h'?b.pos>=a.start-5 && b.pos<=a.end+5 && a.pos>=b.start-5 && a.pos<=b.end+5:intersects(b,a));
    const remaining=new Set(lines),tables=[];
    while(remaining.size) {
        const group=[remaining.values().next().value];remaining.delete(group[0]);
        for(let i=0;i<group.length;i++)for(const line of remaining)if(intersects(group[i],line)){group.push(line);remaining.delete(line);}
        const positions=axis=>{
            const result=[];
            for(const value of group.filter(line=>line.axis===axis).map(line=>line.pos).sort((a,b)=>a-b)) {
                if(!result.length || value-result[result.length-1]>6)result.push(value);
            }
            return result;
        };
        const xs=positions('v'),ys=positions('h'),cols=xs.length-1,rows=ys.length-1;
        if(cols<1 || rows<1 || cols>100 || rows>1000 || cols*rows>10000)continue;
        if(xs.some((x,i)=>i && x-xs[i-1]<8)||ys.some((y,i)=>i && y-ys[i-1]<8))continue;
        const border=(axis,pos,start,end)=>{
            const segments=group.filter(line=>line.axis===axis && Math.abs(line.pos-pos)<=5)
                .map(line=>[Math.max(start,line.start),Math.min(end,line.end)]).filter(([a,b])=>b>a).sort((a,b)=>a[0]-b[0]);
            let covered=0,until=start;
            for(const [a,b] of segments){covered+=Math.max(0,b-Math.max(a,until));until=Math.max(until,b);}
            return covered>=(end-start)*0.85;
        };
        if(!border('h',ys[0],xs[0],xs[cols])||!border('h',ys[rows],xs[0],xs[cols])||
            !border('v',xs[0],ys[0],ys[rows])||!border('v',xs[cols],ys[0],ys[rows]))continue;
        const parent=Array.from({length:rows*cols},(_,i)=>i);
        const root=index=>{while(parent[index]!==index){parent[index]=parent[parent[index]];index=parent[index];}return index;};
        const join=(a,b)=>{parent[root(a)]=root(b);};
        for(let r=0;r<rows;r++)for(let c=0;c<cols;c++) {
            if(c && !border('v',xs[c],ys[r],ys[r+1]))join(r*cols+c,r*cols+c-1);
            if(r && !border('h',ys[r],xs[c],xs[c+1]))join(r*cols+c,(r-1)*cols+c);
        }
        const regions=new Map();
        for(let r=0;r<rows;r++)for(let c=0;c<cols;c++) {
            const id=root(r*cols+c);
            if(!regions.has(id))regions.set(id,[]);regions.get(id).push({r,c});
        }
        const cells=[];let valid=true;
        for(const slots of regions.values()) {
            const r=Math.min(...slots.map(s=>s.r)),c=Math.min(...slots.map(s=>s.c));
            const bottom=Math.max(...slots.map(s=>s.r))+1,right=Math.max(...slots.map(s=>s.c))+1;
            if((bottom-r)*(right-c)!==slots.length){valid=false;break;}
            cells.push({r,c,rs:bottom-r,cs:right-c,x0:xs[c],x1:xs[right],y0:ys[r],y1:ys[bottom]});
        }
        if(valid && cells.length>1)tables.push({rows,cols,x0:xs[0],x1:xs[cols],y0:ys[0],y1:ys[rows],cells});
    }
    return tables.sort((a,b)=>a.y0-b.y0||a.x0-b.x0);
}
function ocrWords(tsv) {
    return String(tsv||'').split(/\r?\n/).slice(1).flatMap(line=>{
        const fields=line.split('\t'),text=fields.slice(11).join('\t').trim();
        if(fields[0]!=='5'||!text)return [];
        const [x,y,w,h]=fields.slice(6,10).map(Number);
        if(![x,y,w,h].every(Number.isFinite)||w<=0||h<=0)return [];
        return [{text,x,y,w,h,line:fields.slice(1,5).join(':')}];
    });
}
function ocrLineText(words) {
    const lines=new Map();
    for(const word of words){if(!lines.has(word.line))lines.set(word.line,[]);lines.get(word.line).push(word);}
    return [...lines.values()].sort((a,b)=>a[0].y-b[0].y||a[0].x-b[0].x)
        .map(line=>line.sort((a,b)=>a.x-b.x).map(word=>word.text).join(' ')).join('\n');
}
function ocrPageLayout(tables,tsv,text,pageNumber) {
    const section=element('section'),heading=element('p','〔第 '+pageNumber+' 頁〕');section.append(heading);
    const words=ocrWords(tsv),assigned=new Set(),pieces=[];
    for(const table of tables) {
        const node=element('table'),body=element('tbody');node.append(body);
        for(let r=0;r<table.rows;r++) {
            const row=element('tr');
            for(const cell of table.cells.filter(cell=>cell.r===r).sort((a,b)=>a.c-b.c)) {
                const inside=words.filter(word=>!assigned.has(word) && word.x+word.w/2>=cell.x0 && word.x+word.w/2<cell.x1 && word.y+word.h/2>=cell.y0 && word.y+word.h/2<cell.y1);
                inside.forEach(word=>assigned.add(word));
                const td=element('td',ocrLineText(inside));td.rowSpan=cell.rs;td.colSpan=cell.cs;
                if(!td.textContent)td.append(element('br'));row.append(td);
            }
            body.append(row);
        }
        pieces.push({x:table.x0,y:table.y0,node});
    }
    const outside=new Map();
    for(const word of words.filter(word=>!assigned.has(word))) {
        if(!outside.has(word.line))outside.set(word.line,[]);outside.get(word.line).push(word);
    }
    for(const line of outside.values())pieces.push({x:Math.min(...line.map(w=>w.x)),y:Math.min(...line.map(w=>w.y)),node:element('p',ocrLineText(line))});
    if(!words.length)pieces.push({x:0,y:0,node:element('p',text)});
    for(const piece of pieces.sort((a,b)=>a.y-b.y||a.x-b.x))section.append(piece.node);
    return section;
}
function appendOCRPage(page,pageNumber) {
    $('ocr-layout').append(ocrPageLayout(page.tables||[],page.tsv,page.text,pageNumber));
    if(ocrJob.pages.some(page=>page.tables?.length)) {
        $('ocr-format-label').hidden=false;showOCRFormat();
    }
    if(ocrJob.tableMode) {
        const missing=ocrJob.pages.filter(page=>!page.tables?.length).length;
        $('ocr-layout-note').hidden=false;
        $('ocr-layout-note').textContent='表格版與純文字可分別校對，插入事件使用目前選取的格式。'+
            (missing?'有 '+missing+' 頁未偵測到完整格線，該頁保留文字。':'請確認列欄、合併格與文字位置後再插入事件。');
    }
}
function resetOCR() {
    clearOCROutput();
    clearOCRLayout();
    ocrFiles=[]; ocrJob=undefined; ocrContext=undefined;
    $('ocr-file').value=''; $('ocr-result').value=''; $('ocr-inputs').replaceChildren();
    $('ocr-target').textContent=''; $('recognize').textContent='開始辨識';
    $('recognize').hidden=false;
    $('ocr-language').value='chi_tra+eng';$('ocr-table-mode').checked=false;
    $('insert-ocr').disabled=true;
    setToolStatus('');
}
$('tool').addEventListener('close',()=>{if(!$('tool').open)resetOCR();});
function setOCRContext(fromEditor) {
    ocrContext = fromEditor ? draft : undefined;
    $('insert-ocr').disabled = !ocrContext;
    $('insert-ocr').hidden = !ocrContext;
    $('ocr-attach').hidden = !!ocrContext;
    $('ocr-target').textContent = ocrContext ? '目標事件：' + (ocrContext.id ? '' : '新增事件：') + ($('event-title').value || '未命名') : '辨識後可選擇新增事件或加入既有事件。';
    $('ocr-attach').textContent = ocrContext ? '加入此事件附件' : '加入事件…';
}
function clearOCROutput() {
    if (ocrOutput?.url) URL.revokeObjectURL(ocrOutput.url);
    ocrOutput = undefined;
    for (const id of ['ocr-preview','ocr-download','ocr-attach']) $(id).disabled = true;
}
function drawOCRInputs() {
    $('ocr-inputs').replaceChildren();
    ocrFiles.forEach((file,index) => {
        const row=element('li'); row.append(document.createTextNode(file.name + ' '));
        for (const [label,offset] of [['上移',-1],['下移',1]]) {
            const button=element('button',label); button.type='button';
            button.disabled=busy || index+offset<0 || index+offset>=ocrFiles.length;
            button.onclick=()=>{
                [ocrFiles[index],ocrFiles[index+offset]]=[ocrFiles[index+offset],ocrFiles[index]];
                ocrJob=undefined; clearOCROutput(); $('ocr-result').value=''; drawOCRInputs();
                clearOCRLayout();
                $('recognize').textContent='開始辨識';$('recognize').hidden=false;
            }; row.append(button);
        }
        $('ocr-inputs').append(row);
    });
}
$('ocr-file').onchange = () => {
    ocrFiles=[...$('ocr-file').files]; ocrJob=undefined; clearOCROutput(); $('ocr-result').value='';
    clearOCRLayout();
    $('recognize').textContent='開始辨識';$('recognize').hidden=false; drawOCRInputs();
};
for(const control of [$('ocr-language'),$('ocr-table-mode')])control.onchange=()=>{
    ocrJob=undefined;clearOCROutput();$('ocr-result').value='';
    clearOCRLayout();
    $('recognize').textContent='開始辨識';$('recognize').hidden=false;setToolStatus('設定已變更，請重新開始辨識。');
};
// Work on a same-size copy so the original image and PDF coordinates are preserved.
function tableOCRCanvas(source) {
    const canvas=element('canvas');canvas.width=source.width;canvas.height=source.height;
    try {
        const context=canvas.getContext('2d',{willReadFrequently:true});
        context.drawImage(source,0,0);
        const image=context.getImageData(0,0,canvas.width,canvas.height),pixels=image.data;
        const w=canvas.width,h=canvas.height,mask=new Uint8Array(w*h),runs=[];
        const dark=index=>pixels[index*4+3]>128 && (pixels[index*4]+pixels[index*4+1]+pixels[index*4+2])<360;
        const horizontal=Math.max(100,Math.round(w*0.08)),vertical=Math.max(60,Math.round(h*0.025));
        for(let y=0;y<h;y++) {
            let start=0;
            for(let x=0;x<=w;x++) {
                if(x<w && dark(y*w+x))continue;
                if(x-start>=horizontal) {
                    runs.push({axis:'h',pos:y,start,end:x});
                    for(let k=start;k<x;k++)mask[y*w+k]=1;
                }
                start=x+1;
            }
        }
        for(let x=0;x<w;x++) {
            let start=0;
            for(let y=0;y<=h;y++) {
                if(y<h && dark(y*w+x))continue;
                if(y-start>=vertical) {
                    runs.push({axis:'v',pos:x,start,end:y});
                    for(let k=start;k<y;k++)mask[k*w+x]=1;
                }
                start=y+1;
            }
        }
        for(let index=0;index<mask.length;index++)if(mask[index]) {
            pixels[index*4]=pixels[index*4+1]=pixels[index*4+2]=pixels[index*4+3]=255;
        }
        canvas.ocrTables=ruledTables(runs);
        context.putImageData(image,0,0);return canvas;
    } catch(error) {canvas.width=canvas.height=0;throw error;}
}
async function imageCanvas(file) {
    const bitmap=await createImageBitmap(file);
    const scale=Math.min(1,3500/Math.max(bitmap.width,bitmap.height));
    const canvas=element('canvas'); canvas.width=Math.max(1,Math.round(bitmap.width*scale)); canvas.height=Math.max(1,Math.round(bitmap.height*scale));
    const context=canvas.getContext('2d'); context.fillStyle='#fff'; context.fillRect(0,0,canvas.width,canvas.height); context.drawImage(bitmap,0,0,canvas.width,canvas.height);
    bitmap.close(); return canvas;
}
async function assembleOCR(job) {
    const output=job.isPdf ? await PDFLib.PDFDocument.load(job.original) : await PDFLib.PDFDocument.create();
    for (let i=0;i<job.pages.length;i++) {
        const result=job.pages[i];
        let page;
        if (job.isPdf) page=output.getPage(i);
        else {
            const image=await output.embedPng(result.image);
            page=output.addPage([result.width*72/150,result.height*72/150]);
            page.drawImage(image,{x:0,y:0,width:page.getWidth(),height:page.getHeight()});
        }
        const [text]=await output.embedPdf(result.pdf);
        const box=page.getCropBox(), rotation=((page.getRotation().angle%360)+360)%360;
        const swapped=rotation===90 || rotation===270;
        let x=box.x,y=box.y;
        if (rotation===90 || rotation===180) x+=box.width;
        if (rotation===180 || rotation===270) y+=box.height;
        page.drawPage(text,{x,y,width:swapped?box.height:box.width,height:swapped?box.width:box.height,rotate:PDFLib.degrees(rotation)});
    }
    const bytes=await output.save();
    clearOCROutput();
    const blob=new Blob([bytes],{type:'application/pdf'});
    ocrOutput={blob,name:job.files[0].name.replace(/\.[^.]+$/,'').slice(0,240)+'-可搜尋.pdf',url:URL.createObjectURL(blob)};
    for (const id of ['ocr-preview','ocr-download','ocr-attach']) $(id).disabled=false;
}
async function renderOCRPage(page) {
    // PDF coordinates use 72 points per inch; do not cap the pixel dimensions.
    const view=page.getViewport({scale:300/72});
    const canvas=element('canvas');
    try {
        const width=Math.ceil(view.width), height=Math.ceil(view.height);
        if(!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width<1 || height<1) throw Error('頁面尺寸無效');
        canvas.width=width; canvas.height=height;
        const context=canvas.getContext('2d');
        if(canvas.width!==width || canvas.height!==height || !context) throw Error('頁面過大或瀏覽器無法配置畫布記憶體');
        await page.render({canvasContext:context,viewport:view}).promise;
        return canvas;
    } catch(error) {
        canvas.width=canvas.height=0;
        throw Error('無法以 300 DPI 建立辨識影像（未降低解析度）：'+(error.message || String(error)));
    } finally { page.cleanup(); }
}

async function runOCR() {
    if (!ocrFiles.length) throw Error('請先選擇一份 PDF 或多張圖片。');
    const isPdf=file=>file.type==='application/pdf' || /\.pdf$/i.test(file.name);
    if (ocrFiles.some(isPdf) && (ocrFiles.length!==1 || !isPdf(ocrFiles[0]))) throw Error('一次可處理一份 PDF，或多張圖片；請勿混合選取。');
    if (ocrFiles.some(file=>!isPdf(file) && !file.type.startsWith('image/'))) throw Error('請選擇 PDF 或圖片。');
    const language=$('ocr-language').value,tableMode=$('ocr-table-mode').checked;
    if(!['chi_tra+eng','chi_tra','eng'].includes(language))throw Error('請選擇支援的辨識語言。');
    if(ocrJob && (ocrJob.language!==language || ocrJob.tableMode!==tableMode))ocrJob=undefined;
    if (!ocrJob) {
        clearOCRLayout();
        ocrJob={files:[...ocrFiles],isPdf:isPdf(ocrFiles[0]),pages:[],original:undefined,language,tableMode};
    }
    const job=ocrJob; let pdf,worker,deadline;
    $('ocr-file').disabled=true; $('ocr-language').disabled=$('ocr-table-mode').disabled=true;drawOCRInputs();
    $('ocr-layout').contentEditable='false';$('ocr-result').readOnly=true;$('ocr-result-format').disabled=true;
    for (const id of ['ocr-preview','ocr-download','ocr-attach']) $(id).disabled=true;
    try {
        if (job.isPdf) {
            job.original ||= new Uint8Array(await job.files[0].arrayBuffer());
            const pdfjs=await import('/vendor/pdf.mjs'); pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdf.worker.mjs';
            pdf=await pdfjs.getDocument({data:job.original.slice(),isEvalSupported:false,cMapUrl:'/vendor/cmaps/',cMapPacked:true,standardFontDataUrl:'/vendor/standard_fonts/'}).promise;
        }
        const total=pdf ? pdf.numPages : job.files.length;
        let current=job.pages.length+1;
        if (job.pages.length < total) {
            const engine=Tesseract.createWorker(job.language.split('+'),1,{
                workerPath:'/vendor/worker.min.js',corePath:'/vendor/core',langPath:'/vendor/lang',workerBlobURL:true,cacheMethod:'none',
                errorHandler:error=>setToolStatus('OCR 引擎錯誤：'+String(error)),
                logger:info=>setToolStatus('第 '+current+' / '+total+' 頁 · '+(info.status==='recognizing text' ? '辨識中 '+Math.round(info.progress*100)+'%' : '準備辨識中…'))
            });
            try {
                worker=await Promise.race([engine,new Promise((_,reject)=>{deadline=setTimeout(()=>reject(Error('OCR 載入逾時，請重試。')),45000);})]);
            } catch (error) { engine.then(late=>late.terminate()).catch(()=>{}); throw error; }
            finally { clearTimeout(deadline); }
            await worker.setParameters({tessedit_pageseg_mode:job.tableMode?'11':'3',preserve_interword_spaces:'1'});
            for (let index=job.pages.length;index<total;index++) {
                current=index+1; let canvas,prepared;
                try {
                    setToolStatus('第 '+current+' / '+total+' 頁 · 準備辨識…');
                    if (pdf) {
                        canvas=await renderOCRPage(await pdf.getPage(current));
                    } else canvas=await imageCanvas(job.files[index]);
                    // Tesseract embeds Unicode mappings in its invisible PDF text layer.
                    if(job.tableMode) {
                        setToolStatus('第 '+current+' / '+total+' 頁 · 正在處理表格線…');
                        prepared=tableOCRCanvas(canvas);
                    }
                    const recognized=await worker.recognize(prepared || canvas,{pdfTextOnly:true},{text:true,pdf:true,tsv:job.tableMode});
                    if (!recognized.data.pdf?.length) throw Error('辨識引擎未產生 PDF 文字層。');
                    let image;
                    if (!pdf) image=new Uint8Array(await (await new Promise(resolve=>canvas.toBlob(resolve,'image/png'))).arrayBuffer());
                    const page={pdf:new Uint8Array(recognized.data.pdf),text:recognized.data.text.trim(),image,width:canvas.width,height:canvas.height,
                        tables:prepared?.ocrTables||[],tsv:recognized.data.tsv||''};
                    job.pages.push(page);appendOCRPage(page,current);
                    $('ocr-result').value=job.pages.map((p,i)=>'〔第 '+(i+1)+' 頁〕\n'+p.text).join('\n\n');
                } catch (error) { throw Error('第 '+current+' / '+total+' 頁失敗：'+error.message+' 已保留前面結果，按「重試未完成頁面」繼續。'); }
                finally { if (prepared) prepared.width=prepared.height=0;if (canvas) canvas.width=canvas.height=0; }
            }
        }
        setToolStatus('正在建立可搜尋 PDF…'); await assembleOCR(job);
        $('recognize').hidden=true;
        setToolStatus('已完成全部 '+total+' 頁，可預覽、下載或加入事件。插入事件使用目前選取的表格版或純文字；校對不會修改 PDF 文字層。');
    } catch (error) { $('recognize').textContent='重試未完成頁面';$('recognize').hidden=false; throw Error(error?.message || String(error)); }
    finally {
        if (worker) await worker.terminate();
        if (pdf) await pdf.destroy();
        $('ocr-file').disabled=false;
        $('ocr-language').disabled=$('ocr-table-mode').disabled=false;
        $('ocr-layout').contentEditable='true';$('ocr-result').readOnly=false;$('ocr-result-format').disabled=false;
        // toolTask resets busy after this finally.
        setTimeout(drawOCRInputs,0);
    }
}
$('ocr-preview').onclick=()=>{if(ocrOutput) openPreview({name:ocrOutput.name,url:ocrOutput.url});};
$('ocr-download').onclick=()=>{
    if (!ocrOutput) return;
    const link=element('a'); link.href=ocrOutput.url; link.download=ocrOutput.name; link.click();
};
async function attachOCRToDraft(output=ocrOutput) {
    if (!output || !$('edit').open || !draft) throw Error('請先開啟目標事件。');
    if (output.blob.size > 25000000) throw Error('辨識後 PDF 超過單檔 25 MB，請下載保存或分割後再加入附件。');
    if (draft.files.length >= 100) throw Error('每個事件最多 100 個附件。');
    const original=draft;
    const result=await api('/api/upload',{method:'POST',headers:{'X-Local-Token':token},body:output.blob});
    if (draft!==original || !$('edit').open) throw Error('目標事件已關閉，請重新選擇。');
    draft.files.push({id:result.id,name:output.name}); drawFiles($('edit-files'),draft.files,true);
    $('tool').close(); notice('PDF 已加入「'+($('event-title').value || '未命名事件')+'」，請確認並儲存事件。');
}
function destinationOptions(select, items, value) {
    select.replaceChildren();
    for (const item of items) { const option=element('option',item.name); option.value=item.id; select.append(option); }
    if (items.some(item=>item.id===value)) select.value=value;
}
function drawDestinationSections() {
    destinationOptions($('ocr-section'),data.Sections.filter(s=>s.projectId===$('ocr-project').value),sectionId);
    drawDestinationEvents();
}
function drawDestinationEvents() {
    destinationOptions($('ocr-event'),[{id:'',name:'新增事件'},...data.Events.filter(e=>e.sectionId===$('ocr-section').value).sort((a,b)=>b.date.localeCompare(a.date)).map(e=>({id:e.id,name:e.date+' '+e.title}))],'');
}
$('ocr-project').onchange=drawDestinationSections;
$('ocr-section').onchange=drawDestinationEvents;
$('ocr-attach').onclick=async()=>{
    if (busy || !ocrOutput) return;
    if (ocrContext) {
        if (draft!==ocrContext || !$('edit').open) { notice('原事件已關閉，請重新從目標事件開啟 OCR。'); return; }
        return toolTask($('ocr-attach'),attachOCRToDraft);
    }
    destinationOptions($('ocr-project'),data.Projects,projectId); drawDestinationSections();
    $('ocr-destination').showModal();
};
$('ocr-destination-form').onsubmit=async e=>{
    e.preventDefault();
    if (busy) return;
    const targetSection=data.Sections.find(s=>s.id===$('ocr-section').value && s.projectId===$('ocr-project').value);
    if (!targetSection) { notice('請先建立專案與分頁，再選擇目的地。'); return; }
    const existing=$('ocr-event').value ? data.Events.find(v=>v.id===$('ocr-event').value && v.sectionId===targetSection.id) : undefined;
    const button=e.submitter || $('ocr-destination-form').querySelector('[type=submit]'); button.disabled=true;
    try {
        if (existing && !data.Events.some(v=>v.id===existing.id)) throw Error('找不到目標事件。');
        projectId=targetSection.projectId; sectionId=targetSection.id; eventId=existing?.id;
        const output=ocrOutput;
        $('ocr-destination').close(); $('tool').close(); render(); openEditor(existing);
        await toolTask($('ocr-attach'),()=>attachOCRToDraft(output));
    } catch (error) { alert(error.message); }
    finally { button.disabled=false; }
};

refresh().catch(error=>notice(error.message));
