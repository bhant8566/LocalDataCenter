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
        let calls=0,inject=true;
        Tesseract.createWorker=async(...args)=>{
            const worker=await nativeWorker(...args),recognize=worker.recognize.bind(worker);
            worker.recognize=async(...parameters)=>{calls++;if(inject && calls===2){inject=false;throw Error('Test page failure');}return recognize(...parameters);};
            return worker;
        };
        let failed=false;try{await runOCR();}catch(error){failed=error.message.includes('2 / 2');if(!failed)throw error;}
        assert(failed && ocrJob.pages.length===1 && !ocrOutput,'Failed OCR preserves completed pages without exposing incomplete PDF');
        await runOCR();
        console.log('TEST: image recognition and failed-page retry completed');
        assert(ocrJob.pages.length===2 && calls===3 && ocrOutput,'Retry resumes at failed page');
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

        await refresh();const existing=data.Events.find(e=>e.id===original.eventId);openEditor(existing);setOCRContext(true);
        assert($('ocr-target').textContent.includes(existing.title),'Editor OCR shows target event name');
        const count=draft.files.length;await attachOCRToDraft();assert(draft.files.length===count+1 && $('edit').open,'Editor OCR attaches into original draft');
        const attached=draft.files.at(-1);$('edit').close();await refresh();assert(!data.Events.find(e=>e.id===existing.id).files.some(f=>f.id===attached.id),'OCR attachment remains draft until saved');

        setOCRContext(false);$('ocr-attach').click();assert($('ocr-destination').open,'Tool OCR requests destination');
        $('ocr-destination').close();assert(ocrOutput,'Cancel destination preserves OCR result');
        $('ocr-attach').click();$('ocr-project').value=original.projectId;drawDestinationSections();$('ocr-section').value=original.sectionId;drawDestinationEvents();$('ocr-event').value=existing.id;
        $('ocr-destination-form').requestSubmit();await until(()=>$('edit').open && draft.id===existing.id && draft.files.some(f=>f.name===ocrOutput.name) && !busy);
        assert(draft.id===existing.id,'Tool OCR can open existing target event');$('edit').close();
        setOCRContext(false);$('ocr-attach').click();$('ocr-project').value=original.projectId;drawDestinationSections();$('ocr-section').value=original.sectionId;drawDestinationEvents();$('ocr-event').value='';
        $('ocr-destination-form').requestSubmit();await until(()=>$('edit').open && !draft.id && draft.files.length===1 && !busy);
        assert(!draft.id && draft.sectionId===original.sectionId,'Tool OCR can open a new event draft in selected section');$('edit').close();
        return {passed:checks.length,checks};
    } finally {
        Tesseract.createWorker=nativeWorker;
        for(const dialog of window.document.querySelectorAll('dialog[open]'))dialog.close();disposePreview();
        projectId=original.projectId;sectionId=original.sectionId;eventId=original.eventId;await refresh();
    }
})()
