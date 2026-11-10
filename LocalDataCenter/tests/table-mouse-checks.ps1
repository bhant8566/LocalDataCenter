# Native CDP mouse input exercises pointer capture and mouse-button chords.
function Invoke-TableUI([string]$expression) {
    $reply = Cdp 'Runtime.evaluate' @{expression=$expression;awaitPromise=$true;returnByValue=$true}
    if ($reply.exceptionDetails) { throw ($reply.exceptionDetails | ConvertTo-Json -Depth 10) }
    return $reply.result.value
}
function Send-TableMouse($type, $point, $button = 'none', $buttons = 0, $modifiers = 0) {
    $null = Cdp 'Input.dispatchMouseEvent' @{type=$type;x=$point.x;y=$point.y;button=$button;buttons=$buttons;modifiers=$modifiers;clickCount=1}
}
function Click-TableUI($point, $button = 'left', $mask = 1) {
    Send-TableMouse 'mousePressed' $point $button $mask
    Send-TableMouse 'mouseReleased' $point $button 0
}
function Save-TableScreenshot($name) {
    $shot = Cdp 'Page.captureScreenshot' @{format='png'}
    [IO.File]::WriteAllBytes((Join-Path $testRoot $name),[Convert]::FromBase64String($shot.data))
}
try {
    $points = Invoke-TableUI '(()=>{openEditor();editor.innerHTML="<table><tr><td><b>Alpha</b></td><td>Bravo</td><td>Charlie</td></tr><tr><td>Delta</td><td>Echo</td><td>Foxtrot</td></tr></table>";editor.scrollIntoView({block:"center"});const cells=editor.querySelectorAll("td");const center=el=>{const r=el.getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};};return {a:center(cells[0]),b:center(cells[5])};})()'
    Click-TableUI $points.a
    if (!(Invoke-TableUI '!selectingCells && !editor.querySelector(".cell-selected")')) { throw 'Native click unexpectedly selected cells' }
    $null = Cdp 'Input.dispatchKeyEvent' @{type='keyDown';key='Shift';code='ShiftLeft';windowsVirtualKeyCode=16;modifiers=8}
    Send-TableMouse 'mousePressed' $points.a 'left' 1 8
    Send-TableMouse 'mouseMoved' $points.b 'none' 1 8
    Send-TableMouse 'mouseReleased' $points.b 'left' 0 8
    $null = Cdp 'Input.dispatchKeyEvent' @{type='keyUp';key='Shift';code='ShiftLeft';windowsVirtualKeyCode=16;modifiers=0}
    if (!(Invoke-TableUI 'editor.querySelectorAll(".cell-selected").length===6 && !selectingCells')) { throw 'Native Shift drag failed to select six cells' }
    Save-TableScreenshot 'table-selection.png'
    Click-TableUI $points.b 'right' 2
    if (!(Invoke-TableUI '!$("table-menu").hidden && editor.querySelectorAll(".cell-selected").length===6')) { throw 'Native right click lost table selection' }
    Save-TableScreenshot 'table-menu.png'
    $point = Invoke-TableUI '(()=>{const r=$("merge-cells").getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()'
    Click-TableUI $point
    if (!(Invoke-TableUI 'editor.querySelector("td").rowSpan===2 && editor.querySelector("td").colSpan===3 && $("table-menu").hidden')) { throw 'Native merge menu action failed' }
    $point = Invoke-TableUI '(()=>{const r=editor.querySelector("td").getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()'
    Click-TableUI $point 'right' 2
    $point = Invoke-TableUI '(()=>{const r=$("split-cells").getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()'
    Click-TableUI $point
    if (!(Invoke-TableUI 'editor.querySelectorAll("td").length===6')) { throw 'Native split menu action failed' }
    $points = Invoke-TableUI '(()=>{editor.innerHTML="<table><tr>"+Array.from({length:25},(_,i)=>"<td>Wide value "+i+"</td>").join("")+"</tr></table>";editor.scrollLeft=0;const a=editor.querySelector("td").getBoundingClientRect(),r=editor.getBoundingClientRect();return {a:{x:a.left+a.width/2,y:a.top+a.height/2},b:{x:r.right-2,y:a.top+a.height/2}};})()'
    Send-TableMouse 'mousePressed' $points.a 'left' 1 8
    Send-TableMouse 'mouseMoved' $points.b 'none' 1 8
    $null = Invoke-TableUI 'new Promise(resolve=>setTimeout(resolve,180))'
    Send-TableMouse 'mouseReleased' $points.b 'left' 0 8
    if (!(Invoke-TableUI 'editor.scrollLeft>0 && editor.querySelectorAll(".cell-selected").length>2 && !selectingCells')) { throw 'Native selection did not auto-scroll at editor edge' }
    $null = Invoke-TableUI '$("color").click()'
    Save-TableScreenshot 'color-palette.png'
    $null = Invoke-TableUI '$("edit").close();(async()=>{const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([Array.from({length:25},(_,i)=>"Long preview column "+i)]),"First sheet");XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([["Another sheet"]]),"Second sheet");window.mouseTestURL=URL.createObjectURL(new Blob([XLSX.write(book,{bookType:"xlsx",type:"array"})]));await openPreview({name:"mouse.xlsx",url:window.mouseTestURL});})()'
    $points = Invoke-TableUI '(()=>{const r=$("preview-title").getBoundingClientRect();return {a:{x:r.left+160,y:r.top+r.height/2},b:{x:r.left+40,y:r.top+r.height/2}};})()'
    Send-TableMouse 'mousePressed' $points.a 'left' 1
    Send-TableMouse 'mousePressed' $points.a 'right' 3
    Send-TableMouse 'mouseMoved' $points.b 'none' 3
    $offset = Invoke-TableUI 'document.querySelector(".workbook-sheet").scrollLeft'
    if ($offset -lt 100) { throw 'Native both-button chord did not pan from preview header' }
    Send-TableMouse 'mouseReleased' $points.b 'right' 1
    Send-TableMouse 'mouseMoved' $points.a 'none' 1
    if ((Invoke-TableUI 'document.querySelector(".workbook-sheet").scrollLeft') -ne $offset) { throw 'Panning continued after one button was released' }
    Send-TableMouse 'mouseReleased' $points.a 'left' 0
    Save-TableScreenshot 'workbook-tabs.png'
    $point = Invoke-TableUI '(()=>{const r=document.querySelectorAll(".workbook-tabs button")[1].getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()'
    Click-TableUI $point
    if (!(Invoke-TableUI 'document.querySelector(".workbook-sheet").textContent.includes("Another sheet")')) { throw 'Normal sheet click was blocked after panning' }
    Write-Output 'PASS: native mouse Shift selection, pointer capture, right-click merge/split, edge scrolling, both-button pan from preview header, release and subsequent sheet switch.'
} finally {
    $null = Invoke-TableUI 'for(const dialog of document.querySelectorAll("dialog[open]"))dialog.close();disposePreview();if(window.mouseTestURL){URL.revokeObjectURL(window.mouseTestURL);delete window.mouseTestURL;}'
}
