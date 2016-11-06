//----------------------------------------------------------
// Copyright (C) Microsoft Corporation. All rights reserved.
// Released under the Microsoft Office Extensible File License
// https://raw.github.com/stephen-hardy/docx.js/master/LICENSE.txt
//----------------------------------------------------------

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
            var output, inputDoc, i, j, k, id, doc, inNode, inNodeChild, outNode, outNodeChild, styleAttrNode, footnoteNode, pCount = 0, tempStr, tempNode, val;
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
                        val = inNodeChild.textContent;
                        if (footnoteNode = inNodeChild.getElementsByTagName('footnoteReference')[0]) { val = '<sup><a href="#note-' + footnoteNode.getAttribute('w:id') + '">' + footnoteNode.getAttribute('w:id') + '</a></sup>'; }
                        if (inNodeChild.getElementsByTagName('smallCaps').length) { val = '<span style="font-variant: small-caps">' + val + '</span>'; }
                        if (inNodeChild.getElementsByTagName('b').length) { val = '<b>' + val + '</b>'; }
                        if (inNodeChild.getElementsByTagName('i').length) { val = '<i>' + val + '</i>'; }
                        if (inNodeChild.getElementsByTagName('u').length) { val = '<u>' + val + '</u>'; }
                        if (inNodeChild.getElementsByTagName('strike').length) { val = '<s>' + val + '</s>'; }
                        if (styleAttrNode = inNodeChild.getElementsByTagName('vertAlign')[0]) {
                            if (styleAttrNode.getAttribute('w:val') === 'subscript') { val = '<sub>' + val + '</sub>'; }
                            if (styleAttrNode.getAttribute('w:val') === 'superscript') { val = '<sup>' + val + '</sup>'; }
                        }
                        if (styleAttrNode = inNodeChild.getElementsByTagName('sz')[0]) { val = '<span style="font-size:' + (styleAttrNode.getAttribute('w:val') / 2) + 'pt">' + val + '</span>'; }
                        if (styleAttrNode = inNodeChild.getElementsByTagName('highlight')[0]) { val = '<span style="background-color:' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>'; }
                        if (styleAttrNode = inNodeChild.getElementsByTagName('color')[0]) { val = '<span style="color:#' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>'; }
                        if (styleAttrNode = inNodeChild.getElementsByTagName('blip')[0]) {
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
                        tempStr += val;
                    }
                    outNode.innerHTML = tempStr;
                }
            }
            output = output.childNodes;
            return output;
        });
        var footnotes = null; //input.files['word/footnotes.xml'].async("string").then(function (data) {
        return Promise.all([styles, mainDocument, footnotes]).then(function (results) {
            var output = {
               styles: results[0],
               mainDocument: results[1],
               footnotes: results[2]
            };
            return output;
        });
    }
}

function docx(file) { 'use strict'; // v1.0.1
    var result, zip = new JSZip(), zipTime, processTime, docProps, word, content;

    zipTime = Date.now();
    return zip.loadAsync(file).then(function (zip) {
        return convertContent(zip);
    });
}
