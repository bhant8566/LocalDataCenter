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
    const select = element('select'); select.setAttribute('aria-label','工作表');
    const container = element('div',undefined,'workbook-sheet');
    for (const name of workbook.SheetNames) { const option = element('option',name); option.value = name; select.append(option); }
    const draw = () => {
        const sheet = workbook.Sheets[select.value];
        container.replaceChildren();
        if (!sheet?.['!ref']) { container.append(element('p','空白工作表')); return; }
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
    select.onchange = draw;
    $('preview-content').append(select,container); draw();
}

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
function cellAt(node) { return (node?.nodeType===1 ? node : node?.parentElement)?.closest('td,th'); }
function clearCellSelection() { editor.querySelectorAll('.cell-selected').forEach(c=>c.classList.remove('cell-selected')); }
function selectedRectangle() {
    if (!cellStart?.isConnected || !cellEnd?.isConnected || cellStart.closest('table')!==cellEnd.closest('table')) throw Error('請先選取同一表格的儲存格。');
    const table=cellStart.closest('table'), model=tableModel(table);
    const a=model.records.find(v=>v.td===cellStart), b=model.records.find(v=>v.td===cellEnd);
    const bounds={r:Math.min(a.r,b.r),c:Math.min(a.c,b.c),bottom:Math.max(a.r+a.rs,b.r+b.rs),right:Math.max(a.c+a.cs,b.c+b.cs)};
    const cells=model.records.filter(v=>v.r<bounds.bottom && v.r+v.rs>bounds.r && v.c<bounds.right && v.c+v.cs>bounds.c);
    return {table,model,bounds,cells};
}
function highlightCells() {
    clearCellSelection();
    try { selectedRectangle().cells.forEach(v=>v.td.classList.add('cell-selected')); } catch {}
}
editor.addEventListener('pointerdown', e => {
    const cell=cellAt(e.target);
    if (!cell || !editor.contains(cell)) { cellStart=cellEnd=undefined; clearCellSelection(); return; }
    if (e.shiftKey && cellStart?.isConnected && cellStart.closest('table')===cell.closest('table')) {
        e.preventDefault(); cellEnd=cell; highlightCells(); return;
    }
    cellStart=cellEnd=cell; clearCellSelection();
    if (e.clientX-cell.getBoundingClientRect().left<=10) { e.preventDefault(); selectingCells=true; highlightCells(); }
});
editor.addEventListener('pointerover', e => {
    const cell=cellAt(e.target);
    if (selectingCells && cell && cell.closest('table')===cellStart.closest('table')) { cellEnd=cell; highlightCells(); }
});
document.addEventListener('pointerup',()=>selectingCells=false);
$('merge-cells').onclick = () => {
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
    try {
        const {table,model,cells}=selectedRectangle();
        cells.filter(v=>v.rs>1 || v.cs>1).forEach(v=>splitRecord(model,v));
        drawTableModel(table,model); clearCellSelection();
    } catch (error) { notice(error.message); }
};

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
function setOCRContext(fromEditor) {
    ocrContext = fromEditor ? draft : undefined;
    $('insert-ocr').disabled = !ocrContext;
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
            }; row.append(button);
        }
        $('ocr-inputs').append(row);
    });
}
$('ocr-file').onchange = () => {
    ocrFiles=[...$('ocr-file').files]; ocrJob=undefined; clearOCROutput(); $('ocr-result').value='';
    $('recognize').textContent='開始辨識'; drawOCRInputs();
};
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
async function runOCR() {
    if (!ocrFiles.length) throw Error('請先選擇一份 PDF 或多張圖片。');
    const isPdf=file=>file.type==='application/pdf' || /\.pdf$/i.test(file.name);
    if (ocrFiles.some(isPdf) && (ocrFiles.length!==1 || !isPdf(ocrFiles[0]))) throw Error('一次可處理一份 PDF，或多張圖片；請勿混合選取。');
    if (ocrFiles.some(file=>!isPdf(file) && !file.type.startsWith('image/'))) throw Error('請選擇 PDF 或圖片。');
    if (!ocrJob) ocrJob={files:[...ocrFiles],isPdf:isPdf(ocrFiles[0]),pages:[],original:undefined};
    const job=ocrJob; let pdf,worker,deadline;
    $('ocr-file').disabled=true; drawOCRInputs();
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
            const engine=Tesseract.createWorker(['chi_tra','eng'],1,{
                workerPath:'/vendor/worker.min.js',corePath:'/vendor/core',langPath:'/vendor/lang',workerBlobURL:true,cacheMethod:'none',
                errorHandler:error=>setToolStatus('OCR 引擎錯誤：'+String(error)),
                logger:info=>setToolStatus('第 '+current+' / '+total+' 頁 · '+(info.status==='recognizing text' ? '辨識中 '+Math.round(info.progress*100)+'%' : '準備辨識中…'))
            });
            try {
                worker=await Promise.race([engine,new Promise((_,reject)=>{deadline=setTimeout(()=>reject(Error('OCR 載入逾時，請重試。')),45000);})]);
            } catch (error) { engine.then(late=>late.terminate()).catch(()=>{}); throw error; }
            finally { clearTimeout(deadline); }
            for (let index=job.pages.length;index<total;index++) {
                current=index+1; let canvas;
                try {
                    setToolStatus('第 '+current+' / '+total+' 頁 · 準備辨識…');
                    if (pdf) {
                        const page=await pdf.getPage(current), base=page.getViewport({scale:1});
                        const view=page.getViewport({scale:Math.min(2.5,3500/Math.max(base.width,base.height))});
                        canvas=element('canvas'); canvas.width=Math.ceil(view.width); canvas.height=Math.ceil(view.height);
                        await page.render({canvasContext:canvas.getContext('2d'),viewport:view}).promise; page.cleanup();
                    } else canvas=await imageCanvas(job.files[index]);
                    // Tesseract embeds Unicode mappings in its invisible PDF text layer.
                    const recognized=await worker.recognize(canvas,{pdfTextOnly:true},{text:true,pdf:true});
                    if (!recognized.data.pdf?.length) throw Error('辨識引擎未產生 PDF 文字層。');
                    let image;
                    if (!pdf) image=new Uint8Array(await (await new Promise(resolve=>canvas.toBlob(resolve,'image/png'))).arrayBuffer());
                    job.pages.push({pdf:new Uint8Array(recognized.data.pdf),text:recognized.data.text.trim(),image,width:canvas.width,height:canvas.height});
                    $('ocr-result').value=job.pages.map((p,i)=>'〔第 '+(i+1)+' 頁〕\n'+p.text).join('\n\n');
                } catch (error) { throw Error('第 '+current+' / '+total+' 頁失敗：'+error.message+' 已保留前面結果，按「重試未完成頁面」繼續。'); }
                finally { if (canvas) canvas.width=canvas.height=0; }
            }
        }
        setToolStatus('正在建立可搜尋 PDF…'); await assembleOCR(job);
        $('recognize').textContent='重新產生 PDF';
        setToolStatus('已完成全部 '+total+' 頁，可預覽、下載或加入事件。校對文字僅影響純文字插入，PDF 文字層使用原辨識結果。');
    } catch (error) { $('recognize').textContent='重試未完成頁面'; throw Error(error?.message || String(error)); }
    finally {
        if (worker) await worker.terminate();
        if (pdf) await pdf.destroy();
        $('ocr-file').disabled=false;
        // toolTask resets busy after this finally.
        setTimeout(drawOCRInputs,0);
    }
}
$('ocr-preview').onclick=()=>{if(ocrOutput) openPreview({name:ocrOutput.name,url:ocrOutput.url});};
$('ocr-download').onclick=()=>{
    if (!ocrOutput) return;
    const link=element('a'); link.href=ocrOutput.url; link.download=ocrOutput.name; link.click();
};
async function attachOCRToDraft() {
    if (!ocrOutput || !$('edit').open || !draft) throw Error('請先開啟目標事件。');
    if (ocrOutput.blob.size > 25000000) throw Error('辨識後 PDF 超過單檔 25 MB，請下載保存或分割後再加入附件。');
    if (draft.files.length >= 100) throw Error('每個事件最多 100 個附件。');
    const original=draft;
    const result=await api('/api/upload',{method:'POST',headers:{'X-Local-Token':token},body:ocrOutput.blob});
    if (draft!==original || !$('edit').open) throw Error('目標事件已關閉，請重新選擇。');
    draft.files.push({id:result.id,name:ocrOutput.name}); drawFiles($('edit-files'),draft.files,true);
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
        $('ocr-destination').close(); $('tool').close(); render(); openEditor(existing);
        await toolTask($('ocr-attach'),attachOCRToDraft);
    } catch (error) { alert(error.message); }
    finally { button.disabled=false; }
};

refresh().catch(error=>notice(error.message));
