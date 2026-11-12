(async()=>{
    const checks=[];
    const assert=(v,m)=>{if(!v)throw Error(m);checks.push(m);};
    const until=async test=>{for(let i=0;i<600;i++){if(test())return;await new Promise(r=>setTimeout(r,100));}throw Error('OCR UI timeout');};
    const original={projectId,sectionId,eventId};
    const nativeWorker=Tesseract.createWorker;
    try {
        const canvas=document.createElement('canvas');canvas.width=1000;canvas.height=360;
        const ctx=canvas.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,1000,360);ctx.fillStyle='black';ctx.font='52px "Microsoft JhengHei"';
        ctx.fillText('繁體中文 測試文件',40,100);ctx.fillText('LOCAL DATA CENTER 12345',40,210);
        const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));
        const one=new File([blob],'first.png',{type:'image/png'}),two=new File([blob],'second.png',{type:'image/png'});
        ocrFiles=[one,two];ocrJob=undefined;drawOCRInputs();
        $('ocr-inputs').children[1].querySelector('button').click();assert(ocrFiles[0]===two,'OCR images reorder visibly');
        // Inject a one-shot failure on page two, but run the real engine on both pages.
        let calls=0,inject=true,lastLanguages;
        Tesseract.createWorker=async(...args)=>{
            lastLanguages=args[0];
            const worker=await nativeWorker(...args),recognize=worker.recognize.bind(worker);
            worker.recognize=async(...parameters)=>{calls++;if(inject && calls===2){inject=false;throw Error('Test page failure');}return recognize(...parameters);};
            return worker;
        };
        let failed=false;try{await runOCR();}catch(error){failed=error.message.includes('2 / 2');if(!failed)throw error;}
        assert(failed && ocrJob.pages.length===1 && !ocrOutput,'Failed OCR preserves completed pages without exposing incomplete PDF');
        await runOCR();
        console.log('TEST: image recognition and failed-page retry completed');
        assert(ocrJob.pages.length===2 && calls===3 && ocrOutput,'Retry resumes at failed page');
        assert($('recognize').hidden,'Completed OCR hides recognition action instead of offering PDF regeneration');
        setOCRContext(false);
        assert($('insert-ocr').hidden && !$('ocr-attach').hidden && !$('ocr-preview').hidden && !$('ocr-download').hidden,'Sidebar OCR shows preview, download and attach without insert action');
        const pdfjs=await import('/vendor/pdf.mjs');
        pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdf.worker.mjs';
        let pdfDocument=await pdfjs.getDocument({data:new Uint8Array(await ocrOutput.blob.arrayBuffer()),isEvalSupported:false}).promise;
        assert(pdfDocument.numPages===2,'Image OCR produces all PDF pages');
        for(let n=1;n<=2;n++) {
            const text=(await (await pdfDocument.getPage(n)).getTextContent()).items.map(i=>i.str).join('');
            assert(/LOCAL|CENTER|12345/.test(text),'OCR page '+n+' PDF has extractable English text');
            assert(/[\u4e00-\u9fff]/.test(text),'OCR page '+n+' PDF has Unicode Chinese text');
        }
        await pdfDocument.destroy();
        await openPreview({name:ocrOutput.name,url:ocrOutput.url});
        await until(()=>$('preview-content').querySelector('.textLayer span'));
        assert($('preview-content').querySelectorAll('.pdf-page').length===2,'OCR preview scrolls through all pages with selectable text layer');
        const span=$('preview-content').querySelector('.textLayer span');const range=window.document.createRange();range.selectNodeContents(span);getSelection().removeAllRanges();getSelection().addRange(range);
        assert(getSelection().toString().length>0,'PDF preview text can be selected');$('preview').close();disposePreview();

        const source=await PDFLib.PDFDocument.create();
        for(let i=0;i<2;i++){const image=await source.embedPng(new Uint8Array(await blob.arrayBuffer()));const page=source.addPage([595.28,841.89]);page.drawImage(image,{x:30,y:620,width:500,height:180});if(i===1)page.setRotation(PDFLib.degrees(90));}
        ocrFiles=[new File([await source.save()],'scan.pdf',{type:'application/pdf'})];ocrJob=undefined;
        await runOCR();
        console.log('TEST: PDF recognition completed');
        assert(ocrJob.pages[0].width===2481 && ocrJob.pages[0].height===3508,'A4 PDF is recognized at 300 DPI without a 3500-pixel cap');
        assert(ocrJob.pages[1].width===3508 && ocrJob.pages[1].height===2481,'Rotated A4 PDF is recognized at 300 DPI with swapped dimensions');
        pdfDocument=await pdfjs.getDocument({data:new Uint8Array(await ocrOutput.blob.arrayBuffer()),isEvalSupported:false}).promise;
        assert(pdfDocument.numPages===2 && (await pdfDocument.getPage(2)).rotate===90,'PDF OCR preserves page count and rotation');
        for(let n=1;n<=2;n++)assert((await (await pdfDocument.getPage(n)).getTextContent()).items.some(i=>i.str.trim()),'PDF OCR preserves text layer on page '+n);
        const outputPage=await pdfDocument.getPage(1);
        assert(Math.abs(outputPage.view[2]-outputPage.view[0]-595.28)<0.01 && Math.abs(outputPage.view[3]-outputPage.view[1]-841.89)<0.01,'300 DPI recognition preserves original PDF page dimensions');
        await pdfDocument.destroy();

        let renderAttempts=0, cleaned=false, failedCanvas, requestedScale;
        const failingPage={getViewport:({scale})=>{requestedScale=scale;return {width:100,height:100};},render:({canvasContext})=>{renderAttempts++;failedCanvas=canvasContext.canvas;return {promise:Promise.reject(Error('Browser canvas limit'))};},cleanup:()=>{cleaned=true;}};
        let renderFailure;
        try {await renderOCRPage(failingPage);} catch(error) {renderFailure=error.message;}
        assert(renderFailure?.includes('300 DPI') && renderFailure.includes('Browser canvas limit') && requestedScale===300/72 && renderAttempts===1,'Canvas-limit failure reports 300 DPI error without retrying at lower resolution');
        assert(cleaned && failedCanvas.width===0 && failedCanvas.height===0,'Failed PDF rendering releases canvas memory and cleans up page');

        const grid=document.createElement('canvas');grid.width=1200;grid.height=500;
        const gridContext=grid.getContext('2d');gridContext.fillStyle='white';gridContext.fillRect(0,0,1200,500);
        gridContext.strokeStyle='black';gridContext.lineWidth=2;
        for(const y of [20,240,480]){gridContext.beginPath();gridContext.moveTo(20,y);gridContext.lineTo(1180,y);gridContext.stroke();}
        for(const x of [20,600,1180]){gridContext.beginPath();gridContext.moveTo(x,20);gridContext.lineTo(x,480);gridContext.stroke();}
        gridContext.fillStyle='black';gridContext.font='42px Arial';gridContext.fillText('TABLE ALPHA 12345',50,140);gridContext.fillText('BETA 67890',650,140);gridContext.fillText('GAMMA 24680',50,370);gridContext.fillText('DELTA 13579',650,370);
        const prepared=tableOCRCanvas(grid);
        assert(prepared.width===grid.width && prepared.height===grid.height,'Table preprocessing keeps coordinates and page dimensions');
        assert(gridContext.getImageData(300,20,1,1).data[0]===0 && prepared.getContext('2d').getImageData(300,20,1,1).data[0]===255,'Table line removed from OCR copy while original remains unchanged');
        assert(prepared.ocrTables.length===1 && prepared.ocrTables[0].rows===2 && prepared.ocrTables[0].cols===2 && prepared.ocrTables[0].cells.length===4,'Original ruling geometry reconstructs the two-by-two grid');
        const mergedRuns=[
            {axis:'h',pos:20,start:20,end:1180},{axis:'h',pos:240,start:20,end:1180},{axis:'h',pos:480,start:20,end:1180},
            {axis:'v',pos:20,start:20,end:480},{axis:'v',pos:600,start:240,end:480},{axis:'v',pos:1180,start:20,end:480}
        ];
        const merged=ruledTables(mergedRuns);
        assert(merged.length===1 && merged[0].cells.some(cell=>cell.r===0 && cell.cs===2),'Missing internal vertical ruling creates a merged column cell');
        const rowspan=ruledTables(mergedRuns.map(run=>({...run})).filter(run=>!(run.axis==='h'&&run.pos===240)).map(run=>run.axis==='v'&&run.pos===600?{...run,start:20}:run).concat([{axis:'h',pos:240,start:600,end:1180}]));
        assert(rowspan[0]?.cells.some(cell=>cell.c===0&&cell.rs===2),'Missing internal horizontal ruling creates a merged row cell');
        assert(ruledTables(mergedRuns.filter(run=>!(run.axis==='h'&&run.pos===480))).length===0,'Incomplete outer border falls back rather than inventing table cells');
        const second=mergedRuns.map(run=>({...run,pos:run.pos+(run.axis==='h'?600:0),start:run.start+(run.axis==='v'?600:0),end:run.end+(run.axis==='v'?600:0)}));
        assert(ruledTables([...mergedRuns.map(run=>({...run})),...second]).length===2,'Separate ruled tables retain independent grids');
        const tsv='level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n'+
            '5\t1\t1\t1\t1\t1\t40\t80\t150\t30\t95\tMerged text\n'+
            '5\t1\t1\t1\t2\t1\t40\t300\t150\t30\t95\t<img src=x>\n'+
            '5\t1\t2\t1\t1\t1\t40\t2\t150\t10\t95\tOutside caption';
        const mergedLayout=ocrPageLayout(merged,tsv,'fallback',1);
        assert(mergedLayout.querySelector('td').colSpan===2 && mergedLayout.querySelector('td').textContent==='Merged text','Word positions fill merged cells');
        assert(mergedLayout.textContent.includes('Outside caption')&&!mergedLayout.querySelector('img'),'Non-table captions preserved and OCR text cannot create executable HTML');
        prepared.width=prepared.height=0;
        const gridBlob=await new Promise(resolve=>grid.toBlob(resolve,'image/png'));
        $('ocr-language').value='eng';$('ocr-table-mode').checked=true;
        $('ocr-language').dispatchEvent(new Event('change'));
        assert(!ocrJob && !ocrOutput && !$('ocr-result').value,'Changing OCR settings invalidates previous results and retry state');
        assert(!$('recognize').hidden,'Changing settings restores start-recognition action');
        ocrFiles=[new File([gridBlob],'table.png',{type:'image/png'})];
        await runOCR();
        assert(lastLanguages.join('+')==='eng' && ocrJob.language==='eng' && ocrJob.tableMode,'English-only and table-mode options reach real OCR engine');
        assert(/ALPHA/.test(ocrJob.pages[0].text) && /BETA/.test(ocrJob.pages[0].text) && /24680/.test(ocrJob.pages[0].text),'Real OCR recognizes text inside multiple ruled table cells');
        const resultCells=[...$('ocr-layout').querySelectorAll('td')];
        assert(!$('ocr-layout').hidden && resultCells.length===4 && /ALPHA/.test(resultCells[0].textContent) && /BETA/.test(resultCells[1].textContent) && /24680/.test(resultCells[2].textContent),'Real OCR words are restored to their original row and column');
        resultCells[0].textContent='校對後 ALPHA';
        $('ocr-result-format').value='text';$('ocr-result-format').dispatchEvent(new Event('change'));
        assert($('ocr-layout').hidden&&!$('ocr-text-label').hidden,'User can select plain-text result');
        $('ocr-result-format').value='layout';$('ocr-result-format').dispatchEvent(new Event('change'));
        assert($('ocr-layout').querySelector('td').textContent==='校對後 ALPHA','Switching formats preserves table edits');
        pdfDocument=await pdfjs.getDocument({data:new Uint8Array(await ocrOutput.blob.arrayBuffer()),isEvalSupported:false}).promise;
        assert((await (await pdfDocument.getPage(1)).getTextContent()).items.some(item=>item.str.includes('ALPHA')),'Table-mode PDF has searchable text');await pdfDocument.destroy();
        await refresh();openEditor(data.Events.find(event=>event.id===original.eventId));setOCRContext(true);
        $('tool').showModal();$('insert-ocr').click();
        assert(editor.querySelector('table') && editor.querySelector('td').textContent==='校對後 ALPHA' && !$('tool').open,'Insert into event preserves edited table HTML');
        await until(()=>!ocrJob && !ocrOutput);
        assert(!$('ocr-layout').children.length&&$('ocr-format-label').hidden,'Closing OCR releases table results as well as text and files');
        await save({...data.Events.find(event=>event.id===original.eventId),kind:'event',html:cleanHTML(editor.innerHTML)});
        $('edit').close();await refresh();openEditor(data.Events.find(event=>event.id===original.eventId));
        assert(editor.querySelector('table').rows.length===2 && editor.querySelector('td').textContent==='校對後 ALPHA','Saved event reloads the reconstructed table and corrected text');$('edit').close();
        $('ocr-language').value='chi_tra';$('ocr-table-mode').checked=false;$('ocr-language').dispatchEvent(new Event('change'));
        ocrFiles=[one];await runOCR();
        assert(lastLanguages.join('+')==='chi_tra' && /[\u4e00-\u9fff]/.test(ocrJob.pages[0].text),'Traditional-Chinese-only mode recognizes Chinese text');
        assert($('ocr-layout').hidden&&$('ocr-format-label').hidden,'Ordinary OCR retains the plain-text result');

        await refresh();const existing=data.Events.find(e=>e.id===original.eventId);openEditor(existing);setOCRContext(true);
        assert($('ocr-target').textContent.includes(existing.title),'Editor OCR shows target event name');
        assert(!$('insert-ocr').hidden && $('ocr-attach').hidden && !$('ocr-preview').hidden && !$('ocr-download').hidden,'Editor OCR keeps insert, preview and download but hides attachment action');
        const reusableOutput=ocrOutput;
        const count=draft.files.length;await attachOCRToDraft();assert(draft.files.length===count+1 && $('edit').open,'Editor OCR attaches into original draft');
        const attached=draft.files.at(-1);$('edit').close();await refresh();assert(!data.Events.find(e=>e.id===existing.id).files.some(f=>f.id===attached.id),'OCR attachment remains draft until saved');

        await new Promise(r=>setTimeout(r,30));
        ocrOutput={...reusableOutput,url:URL.createObjectURL(reusableOutput.blob)};
        setOCRContext(false);$('ocr-attach').click();assert($('ocr-destination').open,'Tool OCR requests destination');
        $('ocr-destination').close();assert(ocrOutput,'Cancel destination preserves OCR result');
        $('ocr-attach').click();$('ocr-project').value=original.projectId;drawDestinationSections();$('ocr-section').value=original.sectionId;drawDestinationEvents();$('ocr-event').value=existing.id;
        const output=ocrOutput;
        $('ocr-destination-form').requestSubmit();await until(()=>$('edit').open && draft.id===existing.id && draft.files.some(f=>f.name===output.name) && !busy);
        assert(draft.id===existing.id,'Tool OCR can open existing target event');$('edit').close();
        await new Promise(r=>setTimeout(r,30));
        if(!ocrOutput)ocrOutput={...reusableOutput,url:URL.createObjectURL(reusableOutput.blob)};
        setOCRContext(false);$('ocr-attach').click();$('ocr-project').value=original.projectId;drawDestinationSections();$('ocr-section').value=original.sectionId;drawDestinationEvents();$('ocr-event').value='';
        $('ocr-destination-form').requestSubmit();await until(()=>$('edit').open && !draft.id && draft.files.length===1 && !busy);
        assert(!draft.id && draft.sectionId===original.sectionId,'Tool OCR can open a new event draft in selected section');$('edit').close();
        $('tool').showModal(); $('tool').querySelector('[data-close]').click();
        await until(()=>!ocrJob && !ocrOutput);
        assert(ocrFiles.length===0 && !$('ocr-file').value && !$('ocr-result').value && !$('ocr-inputs').children.length,'Closing OCR clears inputs, results, job and generated PDF');
        document.querySelector('[data-tool="ocr"]').click();
        assert(!$('ocr-result').value && $('ocr-preview').disabled && $('recognize').textContent==='開始辨識','Reopening OCR starts with a fresh session');
        $('tool').close();
        return {passed:checks.length,checks};
    } finally {
        Tesseract.createWorker=nativeWorker;
        for(const dialog of window.document.querySelectorAll('dialog[open]'))dialog.close();disposePreview();
        projectId=original.projectId;sectionId=original.sectionId;eventId=original.eventId;await refresh();
    }
})()
