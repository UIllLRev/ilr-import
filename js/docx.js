//----------------------------------------------------------
// Copyright (C) Microsoft Corporation. All rights reserved.
// Released under the Microsoft Office Extensible File License
// https://raw.github.com/stephen-hardy/docx.js/master/LICENSE.txt
//----------------------------------------------------------

// Made to actually work and substantially improved by Matt Loar.

function convertContent(input) { 'use strict'; // Convert HTML to WordprocessingML, and vice versa
    function newXMLnode(name, text) {
        var el = doc.createElement('w:' + name);
        if (text) { el.appendChild(doc.createTextNode(text)); }
        return el;
    }
    function newHTMLnode(name, html) {
        var el = document.createElement(name);
        el.innerHTML = html || '';
        return el;
    }
    function color(str) { // Return hex or named color
        if (str.charAt(0) === '#') { return str.substr(1); }
        if (str.indexOf('rgb') < 0) { return str; }
        var values = /rgb\((\d+), (\d+), (\d+)\)/.exec(str), red = +values[1], green = +values[2], blue = +values[3];
        return (blue | (green << 8) | (red << 16)).toString(16);
    }
    function processRunStyle(node, val) {
        var inNode, i, styleAttrNode;
        if (node.getElementsByTagName('smallCaps').length) { val = '<span style="font-variant: small-caps">' + val + '</span>'; }
        if (node.getElementsByTagName('b').length) { val = '<b>' + val + '</b>'; }
        if (node.getElementsByTagName('i').length) { val = '<i>' + val + '</i>'; }
        if (node.getElementsByTagName('u').length) { val = '<u>' + val + '</u>'; }
        if (node.getElementsByTagName('strike').length) { val = '<s>' + val + '</s>'; }
        if (styleAttrNode = node.getElementsByTagName('vertAlign')[0]) {
            if (styleAttrNode.getAttribute('w:val') === 'subscript') { val = '<sub>' + val + '</sub>'; }
            if (styleAttrNode.getAttribute('w:val') === 'superscript') { val = '<sup>' + val + '</sup>'; }
        }
        if (styleAttrNode = node.getElementsByTagName('sz')[0]) { val = '<span style="font-size:' + (styleAttrNode.getAttribute('w:val') / 2) + 'pt">' + val + '</span>'; }
        if (styleAttrNode = node.getElementsByTagName('highlight')[0]) { val = '<span style="background-color: ' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>'; }
        if (styleAttrNode = node.getElementsByTagName('color')[0]) { val = '<span style="color: #' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>'; }
        if (styleAttrNode = node.getElementsByTagName('blip')[0]) {
            id = styleAttrNode.getAttribute('r:embed');
            tempNode = toXML(input.files['word/_rels/document.xml.rels'].data);
            k = tempNode.childNodes.length;
            while (k--) {
                if (tempNode.childNodes[k].getAttribute('Id') === id) {
                    val = '<img src="data:image/png;base64,' + JSZipBase64.encode(input.files['word/' + tempNode.childNodes[k].getAttribute('Target')].data) + '">';
                    break;
                }
            }
        }
        return val;
    }
    function processRun(node, state) {
        var val = '', inNode, i, fnId;
        for (i = 0; inNode = node.childNodes[i]; i++) {
            if (inNode.tagName == 't') { val += inNode.textContent; }
            if (inNode.tagName == 'tab') { val += ' '; }
            if (inNode.tagName == 'footnoteRef') {
                // In this case, footnoteId is just a scalar value
                val += '<span class="fn-ref">' + state.footnoteId + '</span>';
            }
            if (inNode.tagName == 'footnoteReference') {
                fnId = inNode.getAttribute('w:id');
                if (inNode.getAttribute('w:customMarkFollows') == 1) {
                    fnId = inNode.getAttribute('w:id');
                    val += '<sup><a class="fn-reference" href="#note-' + fnId + '">';
                    inNode = node.childNodes[++i];
                    if (inNode.tagName == 't') {
                        val += inNode.textContent;
                    } else {
                        console.warn('customMarkFollows not followed by t');
                        val += '*';
                    }
                    val += '</a></sup>';
                } else {
                    // Here footnoteId is an object reference so we can sequentially number
                    // the non-customMark footnotes.
                    val += '<sup><a class="fn-reference" href="#note-' + fnId + '">'
                        + state.footnoteId.value + '</a></sup>';
                    state.footnoteId.value += 1;
                }
            }
        }
        if (inNode = node.getElementsByTagName('rPr')[0]) {
            val = processRunStyle(inNode, val);
        }
        return val;
    }
    function processEtc(inNode, state, outNode) {
        var styleAttrNode;
        if (inNode.nodeName === 'pPr') {
            if (styleAttrNode = inNode.getElementsByTagName('jc')[0]) { outNode.style.textAlign = styleAttrNode.getAttribute('w:val'); }
            if (styleAttrNode = inNode.getElementsByTagName('pStyle')[0]) { outNode.className = 'pt-' + styleAttrNode.getAttribute('w:val'); }
            if (styleAttrNode = inNode.getElementsByTagName('color')[0]) { outNode.style.color = '#' + styleAttrNode.getAttribute('w:val'); }
        }
        if (inNode.nodeName == 'hyperlink') {
            var id = inNode.getAttribute('r:id');
            var anchor = inNode.getAttribute('w:anchor');
            outNode = outNode.appendChild(newHTMLnode('A'));
            if (id) {
                outNode.href = state.hyperlinks[id];
            } else if (anchor) {
                outNode.href = "#" + anchor;
            }
            for (var i = 0; i < inNode.childNodes.length; i++)
                processEtc(inNode.childNodes[i], state, outNode);
        }
        if (inNode.nodeName === 'bookmarkStart') {
            var name = inNode.getAttribute('w:name');
            outNode = outNode.appendChild(newHTMLnode('A'));
            outNode.name = name;
        }
        if (inNode.nodeName === 'r') {
            var styleNode;
            if (styleNode = inNode.getElementsByTagName('rStyle')[0]) {
                outNode = outNode.appendChild(newHTMLnode('SPAN'));
                outNode.className = 'pt-' + styleNode.getAttribute('w:val');
            }
            outNode.innerHTML += processRun(inNode, state);
        }
    }
    function processPara(node, state, output) {
        var numNode, outNode = output, inNode, j, styleAttrNode, lvl, num;
        if (numNode = node.getElementsByTagName('numPr')[0]) {
            lvl = numNode.getElementsByTagName('ilvl')[0].getAttribute('w:val');
            num = numNode.getElementsByTagName('numId')[0].getAttribute('w:val');
        } else if (inNode = node.getElementsByTagName('pStyle')[0]) {
            lvl = "0";
            num = state.styleNums[inNode.getAttribute('w:val')];
        }
        if (lvl && num) {
            var list = output.querySelector('#list-' + num + '-' + lvl);
            if (!list) {
                if (state.numbering.abstracts[state.numbering.nums[num].abstractId][lvl].numFmt == 'disc') {
                    list = output.appendChild(newHTMLnode('UL'));
                    list.style['list-style-type'] = 'disc';
                } else {
                    list = output.appendChild(newHTMLnode('OL'));
                    list.style['list-style-type'] = state.numbering.abstracts[state.numbering.nums[num].abstractId][lvl].numFmt;
                }
                list.id = 'list-' + num + '-' + lvl;
            } else {
                while (inNode = list.nextSibling) {
                    inNode.parentNode.removeChild(inNode);
                    list.appendChild(inNode);
                }
            }
            outNode = list.appendChild(newHTMLnode('LI'));
        }
        outNode = outNode.appendChild(newHTMLnode('P'));;
        for (j = 0; inNode = node.childNodes[j]; j++) {
            processEtc(inNode, state, outNode);
        }
    }
    function processCell(node, state, output) {
        var inNode, i, fnId;
        var outNode = output.appendChild(newHTMLnode('TD'));
        for (i = 0; inNode = node.childNodes[i]; i++) {
            if (inNode.tagName == 'p') {
                processPara(inNode, state, outNode);
            }
        }
    }
    function processRow(node, state, output) {
        var inNode, i, fnId;
        var outNode = output.appendChild(newHTMLnode('TR'));
        for (i = 0; inNode = node.childNodes[i]; i++) {
            if (inNode.tagName == 'tc')
                processCell(inNode, state, outNode);
        }
    }
    function processTable(node, state, output) {
        var inNode, i, fnId;
        for (i = 0; inNode = node.childNodes[i]; i++) {
            if (inNode.tagName == 'tr')
                processRow(inNode, state, output);
        }
    }

    function toXML(str) {
      return new DOMParser().parseFromString(
                                             str.replace(/<[a-zA-Z]*?:/g, '<').replace(/<\/[a-zA-Z]*?:/g, '</'),
                                             'text/xml'
      ).firstChild;
    }

    if (input.files) { // input is file object
        var promises = [];
        promises.push(input.files['word/_rels/document.xml.rels'].async("string").then(function (data) {
            var output = {"hyperlinks": {}}, inputDoc, i, inNode;
            inputDoc = toXML(data);
            for (i = 0; inNode = inputDoc.childNodes[i]; i++) {
                if (inNode.getAttribute('Type') === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink") {
                    output.hyperlinks[inNode.getAttribute('Id')] = inNode.getAttribute('Target');
                }
            }

            return {"part": "docRels", "content": output};
        }));
        promises.push(input.files['word/styles.xml'].async("string").then(function (data) {
            var output, inputDoc, i, j, k, id, doc, inNode, inNodeChild, outNode, outNodeChild, styleAttrNode, footnoteNode, pCount = 0, tempNode, val, styleNums = {};
            inputDoc = toXML(data);
            output = newHTMLnode('STYLE');
            for (i = 0; inNode = inputDoc.childNodes[i]; i++) {
                if (inNode.nodeName === 'style' /*&& inNode.getAttribute('w:customStyle') == "1"*/) {
                    output.appendChild(document.createTextNode(".pt-" + inNode.getAttribute('w:styleId') + '{'));
                    for (j = 0; inNodeChild = inNode.childNodes[j]; j++) {
                        if (inNodeChild.nodeName === 'pPr') {
                            if (styleAttrNode = inNodeChild.getElementsByTagName('jc')[0]) {
                                output.appendChild(document.createTextNode('text-align: ' + styleAttrNode.getAttribute('w:val') + ';'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('sz')[0]) { 
                                output.appendChild(document.createTextNode('font-size: ' + (styleAttrNode.getAttribute('w:val') / 2) + 'pt;'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('fonts')[0]) { 
                                output.appendChild(document.createTextNode('font-family: ' + styleAttrNode.getAttribute('w:ascii') + ';'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('u')[0]) { 
                                output.appendChild(document.createTextNode('text-decoration: ' + (styleAttrNode.getAttribute('w:val') === 'none' ? 'none' : 'underline') + ';'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('b')[0]) {
                                val = styleAttrNode.getAttribute('w:val');
                                output.appendChild(document.createTextNode('font-weight: ' + (val === null ? 'bold' : (!!val ? 'bold' : 'normal')) + ';'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('numId')[0]) {
                                styleNums[inNode.getAttribute('w:styleId')] = styleAttrNode.getAttribute('w:val');
                            }
                        }
                        if (inNodeChild.nodeName === 'rPr') {
                            if (inNodeChild.getElementsByTagName('smallCaps').length) {
                                output.appendChild(document.createTextNode("font-variant: small-caps;"));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('sz')[0]) { 
                                output.appendChild(document.createTextNode('font-size: ' + (styleAttrNode.getAttribute('w:val') / 2) + 'pt;'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('fonts')[0]) { 
                                output.appendChild(document.createTextNode('font-family: ' + styleAttrNode.getAttribute('w:ascii') + ';'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('u')[0]) { 
                                output.appendChild(document.createTextNode('text-decoration: ' + (styleAttrNode.getAttribute('w:val') === 'none' ? 'none' : 'underline') + ';'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('b')[0]) {
                                val = styleAttrNode.getAttribute('w:val');
                                output.appendChild(document.createTextNode('font-weight: ' + (val === null ? 'bold' : (!!val ? 'bold' : 'normal')) + ';'));
                            }
                            if (styleAttrNode = inNodeChild.getElementsByTagName('color')[0]) {
                                output.appendChild(document.createTextNode('color: #' + styleAttrNode.getAttribute('w:val') + ';'));
                            }
                        }
                    }
                    output.appendChild(document.createTextNode("}\r\n"));
                }
            }
            output = {"stylesheet" : output, "styleNums": styleNums};
            return {"part": "styles", "content": output};
        }));
        if ('word/numbering.xml' in input.files) {
            promises.push(input.files['word/numbering.xml'].async("string").then(function (data) {
                function translateFormat(fmt) {
                    switch (fmt) {
                        case 'bullet':
                            return 'disc';
                        case 'upperLetter':
                            return 'upper-latin';
                        case 'lowerLetter':
                            return 'lower-latin';
                        case 'decimal':
                            return 'decimal';
                        case 'upperRoman':
                            return 'upper-roman';
                        case 'lowerRoman':
                            return 'lower-roman';
                    }
                }
                var output = {"abstracts": {}, "nums": {}}, inputDoc, h, i, j, k, id, doc, fnNode, inNode, inNodeChild, outNode, outNodeChild, styleAttrNode, footnoteNode, pCount = 0, tempStr, tempNode, val;
                inputDoc = toXML(data);
                for (h = 0; fnNode = inputDoc.childNodes[h]; h++) {
                    if (fnNode.nodeName === 'abstractNum') {
                        var abstractId = fnNode.getAttribute('w:abstractNumId');
                        output.abstracts[abstractId] = {};
                        for (i = 0; inNode = fnNode.childNodes[i]; i++) {
                            if (inNode.nodeName === 'lvl') {
                                var lvl = inNode.getAttribute('w:ilvl');
                                output.abstracts[abstractId][lvl] = {};
                                for (j = 0; inNodeChild = inNode.childNodes[j]; j++) {
                                    if (inNodeChild.nodeName === 'numFmt') {
                                        output.abstracts[abstractId][lvl].numFmt = translateFormat(inNodeChild.getAttribute('w:val'));
                                    }
                                }
                            }
                        }
                    } else if (fnNode.nodeName === 'num') {
                        inNode = fnNode.childNodes[0];
                        var numId = fnNode.getAttribute('w:numId');
                        var abstractId = inNode.getAttribute('w:val');
                        output.nums[numId] = {"abstractId": abstractId};
                    }
                }
                return {"part": "numbering", "content": output};
            }));
        }
        if ('word/footnotes.xml' in input.files) {
            promises.push(input.files['word/footnotes.xml'].async("string").then(function (data) {
                var output, inputDoc, h, i, j, k, id, doc, fnNode, inNode, inNodeChild, outNode, outNodeChild, styleAttrNode, footnoteNode, pCount = 0, tempStr, tempNode, val;
                inputDoc = toXML(data);
                output = newHTMLnode('DIV');
                for (h = 0; fnNode = inputDoc.childNodes[h]; h++) {
                    if (!fnNode.getAttribute('w:type')) {
                        for (i = 0; inNode = fnNode.childNodes[i]; i++) {
                            j = inNode.childNodes.length;
                            outNode = output.appendChild(newHTMLnode('P'));
                            tempStr = '';
                            for (j = 0; inNodeChild = inNode.childNodes[j]; j++) {
                                if (inNodeChild.nodeName === 'pPr') {
                                    if (styleAttrNode = inNodeChild.getElementsByTagName('jc')[0]) { outNode.style.textAlign = styleAttrNode.getAttribute('w:val'); }
                                    if (styleAttrNode = inNodeChild.getElementsByTagName('pStyle')[0]) { outNode.className = 'pt-' + styleAttrNode.getAttribute('w:val'); }
                                }
                                if (inNodeChild.nodeName === 'r') {
                                    tempStr += processRun(inNodeChild, {"footnoteId": fnNode.getAttribute('w:id')});
                                }
                                outNode.innerHTML = tempStr;
                            }
                            outNode.id = 'note-' + fnNode.getAttribute('w:id');
                        }
                    }
                }
                return {"part": "footnotes", "content": output};
            }));
        }
        return Promise.all(promises).then(function (results) {
            var ret = {};
            results.forEach(function (d) {
                ret[d.part] = d.content;
            });
            return input.files['word/document.xml'].async("string").then(function (data) {
                var output, inputDoc, i, j, k, id, doc, inNode, inNodeChild, outNode, outNodeChild, styleAttrNode, footnoteNode, pCount = 0, tempNode, val, state = {"footnoteId": {"value": 1}, "numbering": ret.numbering, "styleNums": ret.styles.styleNums, "hyperlinks": ret.docRels.hyperlinks};
                inputDoc = toXML(data).getElementsByTagName('body')[0];
                output = newHTMLnode('DIV');
                for (i = 0; inNode = inputDoc.childNodes[i]; i++) {
                    if (inNode.nodeName == 'p') {
                        processPara(inNode, state, output);
                    } else if (inNode.nodeName == 'tbl') {
                        outNode = output.appendChild(newHTMLnode('TABLE'));
                        processTable(inNode, state, outNode);
                    }
                }
                return output;
            }).then(function (output) {
                ret.mainDocument = output;

                // Fixup footnotes
                var references = ret.mainDocument.getElementsByClassName("fn-reference");
                var refNode, footnote;
                for (var i = 0; refNode = references[i]; i++) {
                    for (var j = i; footnote = ret.footnotes.childNodes[j]; j++) {
                        var noteId = refNode.getAttribute('href').substr(1);
                        if (noteId == footnote.id) {
                            var ref;
                            if (ref = footnote.getElementsByClassName('fn-ref')[0]) {
                                ref.textContent = refNode.textContent;
                            }
                            break;
                        }
                    }
                }
                return ret;
            });
        });
    }
}

function docx(file) { 'use strict'; // v1.0.1
    var zip = new JSZip();

    return zip.loadAsync(file).then(function (zip) {
        return convertContent(zip);
    });
}
