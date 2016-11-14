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
        for (i = 0; inNode = node.childNodes[i]; i++) {
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
            if (styleAttrNode = node.getElementsByTagName('highlight')[0]) { val = '<span style="background-color:' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>'; }
            if (styleAttrNode = node.getElementsByTagName('color')[0]) { val = '<span style="color:#' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>'; }
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
        }
        return val;
    }
    function processRun(node, footnoteId) {
        var val = '', inNode, i, fnId;
        for (i = 0; inNode = node.childNodes[i]; i++) {
            if (inNode.tagName == 't') { val += inNode.textContent; }
            if (inNode.tagName == 'tab') { val += ' '; }
            if (inNode.tagName == 'footnoteRef') {
                // In this case, footnoteId is just a scalar value
                val += '<span class="fn-ref">' + footnoteId + '</span>';
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
                        + footnoteId.value + '</a></sup>';
                    footnoteId.value += 1;
                }
            }
        }
        if (inNode = node.getElementsByTagName('rPr')[0]) {
            val = processRunStyle(inNode, val);
        }
        return val;
    }
    function toXML(str) { return new DOMParser().parseFromString(str.replace(/<[a-zA-Z]*?:/g, '<').replace(/<\/[a-zA-Z]*?:/g, '</'), 'text/xml').firstChild; }
    if (input.files) { // input is file object
        var styles = input.files['word/styles.xml'].async("string").then(function (data) {
            var output, inputDoc, i, j, k, id, doc, inNode, inNodeChild, outNode, outNodeChild, styleAttrNode, footnoteNode, pCount = 0, tempStr, tempNode, val;
            inputDoc = toXML(data);
            output = newHTMLnode('STYLE');
            for (i = 0; inNode = inputDoc.childNodes[i]; i++) {
                if (inNode.getAttribute('w:customStyle') == "1") {
                    output.appendChild(document.createTextNode("." + inNode.getAttribute('w:styleId') + '{'));
                    j = inNode.childNodes.length;
                    tempStr = '';
                    for (j = 0; inNodeChild = inNode.childNodes[j]; j++) {
                        if (inNodeChild.nodeName === 'pPr') {
                            if (styleAttrNode = inNodeChild.getElementsByTagName('jc')[0]) { 
                                output.appendChild(document.createTextNode('text-align: ' + styleAttrNode.getAttribute('w:val') + ';'));
                            }
                        }
                        if (inNodeChild.nodeName === 'rPr') {
                            if (inNodeChild.getElementsByTagName('smallCaps').length) {
                                output.appendChild(document.createTextNode("font-variant: small-caps;"));
                            }
                        }
                    }
                    output.appendChild(document.createTextNode("}\r\n"));
                }
            }
            return output;
        });
        var mainDocument = input.files['word/document.xml'].async("string").then(function (data) {
            var output, inputDoc, i, j, k, id, doc, inNode, inNodeChild, outNode, outNodeChild, styleAttrNode, footnoteNode, pCount = 0, tempStr, tempNode, val, footnoteId = {value : 1};
            inputDoc = toXML(data).getElementsByTagName('body')[0];
            output = newHTMLnode('DIV');
            for (i = 0; inNode = inputDoc.childNodes[i]; i++) {
                j = inNode.childNodes.length;
                outNode = output.appendChild(newHTMLnode('P'));
                tempStr = '';
                for (j = 0; inNodeChild = inNode.childNodes[j]; j++) {
                    if (inNodeChild.nodeName === 'pPr') {
                        if (styleAttrNode = inNodeChild.getElementsByTagName('jc')[0]) { outNode.style.textAlign = styleAttrNode.getAttribute('w:val'); }
                        if (styleAttrNode = inNodeChild.getElementsByTagName('pStyle')[0]) { outNode.className = 'pt-' + styleAttrNode.getAttribute('w:val'); }
                    }
                    if (inNodeChild.nodeName === 'r') {
                        tempStr += processRun(inNodeChild, footnoteId);
                    }
                    outNode.innerHTML = tempStr;
                }
            }
            return output;
        });
        var footnotes = input.files['word/footnotes.xml'].async("string").then(function (data) {
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
                                tempStr += processRun(inNodeChild, fnNode.getAttribute('w:id'));
                            }
                            outNode.innerHTML = tempStr;
                        }
                        outNode.id = 'note-' + fnNode.getAttribute('w:id');
                    }
                }
            }
            return output;
        });
        return Promise.all([styles, mainDocument, footnotes]).then(function (results) {
            var output = {
               styles: results[0],
               mainDocument: results[1],
               footnotes: results[2]
            };

            // Fixup footnotes
            var references = output.mainDocument.getElementsByClassName("fn-reference");
            var refNode, footnote;
            for (var i = 0; refNode = references[i]; i++) {
                for (var j = i; footnote = output.footnotes.childNodes[j]; j++) {
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

            return output;
        });
    }
}

function docx(file) { 'use strict'; // v1.0.1
    var zip = new JSZip();

    return zip.loadAsync(file).then(function (zip) {
        return convertContent(zip);
    });
}
