//=============================================================================
// File: google_docs_highlighting.js
//-----------------------------------------------------------------------------
// Description: This script is intended to be loaded in google docs to add
//              basic syntax matching for text files.
//-----------------------------------------------------------------------------
// Author: Danny Sarraf
//-----------------------------------------------------------------------------
// URL: https://github.com/dddansar/scripts
//-----------------------------------------------------------------------------
// Copyright: MIT License
//
// Copyright (c) 2026 Danny Sarraf
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to
// deal in the Software without restriction, including without limitation the
// rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
// copies of the Software, and to permit persons to whom the Software is
// sell furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
// IN THE SOFTWARE.
//=============================================================================


//=============================================================================
// Color Constants
//=============================================================================
const COLORS = {
  background:   '#FFFFFF',
  defaultText:  '#000000',
  AllSymbols:   '#FF8800',
  AllNumbers:   '#00CCCC',
  AllComments:  '#b0b0b0',
  AllCaps:      '#4488FF',
  AllHLTodo:    '#7c7c00',
  AllHLNote:    '#00AA00',
  AllTitles:    '#f40084',
};

//=============================================================================
// Rules
// Returns the list of regex rules, each mapped to a color group.
//=============================================================================
function createRules() {
  return [
    { group: 'AllSymbols',    pattern: /[&|*+\-^~?!%@#$``'"\/\\<>=!:;,.()\[\]{}]/g },
    { group: 'AllNumbers',    pattern: /\d+/g },
    { group: 'AllComments',   pattern: /^\s*(\/\/|" ).*/gm },
    { group: 'AllCaps',       pattern: /\b[A-Z][A-Z0-9_]+s?\b/g },
    { group: 'AllHLTodo',     pattern: /\bTODO\b/g },
    { group: 'AllHLNote',     pattern: /\bNOTE\b/g },
    { group: 'AllTitles',     pattern: /^\s*#+ .*/gm },
  ];
}

//=============================================================================
// Build a compact run list from char arrays
// Collapses the per-character color array into contiguous {start, end, color} runs.
// Skips null entries (default color), reducing the number of API calls needed.
//=============================================================================
function mergeColorRuns(charColor, textLength) {
  const runs = [];
  let i = 0;
  while (i < textLength) {
    const color = charColor[i];
    if (color === null) { i++; continue; }

    let j = i + 1;
    while (j < textLength && charColor[j] === color) j++;
    runs.push({ start: i, end: j - 1, color });
    i = j;
  }
  return runs;
}

//=============================================================================
// Core: compute highlight runs for one paragraph text
// Applies each rule's regex to text, painting matched characters with
// the rule's color. Later rules overwrite earlier ones. Returns color runs.
//=============================================================================
function applyRulesToText(text, rules) {
  const textLength = text.length;
  const charColor  = new Array(textLength).fill(null);

  for (const rule of rules) {
    const color = COLORS[rule.group];
    rule.pattern.lastIndex = 0;

    let match;
    while ((match = rule.pattern.exec(text)) !== null) {
      const start = match.index;
      const end   = match.index + match[0].length;
      for (let i = start; i < end; i++) {
        charColor[i] = color;
      }
      if (match[0].length === 0) rule.pattern.lastIndex++;
    }
  }
  return mergeColorRuns(charColor, textLength);
}

//=============================================================================
// Apply highlighting to a Text element
// Writes each color run to the Google Docs Text element via the API.
//=============================================================================
function applyRunsToElement(textElement, textLength, runs) {

  for (const run of runs) {
    if (run.color !== null) {
      textElement.setForegroundColor(run.start, run.end, run.color);
    }
  }
}

//=============================================================================
// Main, applies syntax highlighting to the entire document.
// Resets the document colors then re-highlights every paragraph in the body.
//=============================================================================
function highlightFullDocument() {
  const doc   = DocumentApp.getActiveDocument();
  const body  = doc.getBody();
  const rules = createRules();
  const childCount = body.getNumChildren();

  body.setBackgroundColor(COLORS.background);
  body.setForegroundColor(COLORS.defaultText);
  //body.editAsText().setBold(false)
  //body.editAsText().setUnderline(false)
  //body.editAsText().setItalic(false)
  //body.editAsText().setStrikethrough(false);

  for (let p = 0; p < childCount; p++) {
    const block = body.getChild(p);
    if (block.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph  = block.asParagraph();
    const text       = paragraph.getText();
    const textLength = text.length;
    if (textLength === 0) continue;

    const runs = applyRulesToText(text, rules);

    // Skip paragraphs that need no changes (pure default text)
    // This alone eliminates most API calls for sparse documents
    if (runs.length === 0) continue;

    applyRunsToElement(paragraph.editAsText(), textLength, runs);
  }
}

//=============================================================================
// Highlight the next `count` paragraphs from the cursor position
// Finds the cursor's paragraph, then highlights the next `count` paragraphs.
// Resets each paragraph's colors before reapplying, so stale colors are cleared.
//=============================================================================
function highlightAtCursor(count) {
  const doc   = DocumentApp.getActiveDocument();
  const body  = doc.getBody();
  const rules = createRules();
  const childCount = body.getNumChildren();

  const cursor = doc.getCursor();
  if (!cursor) {
    DocumentApp.getUi().alert('Place your cursor in a paragraph first.');
    return;
  }

  // Walk up to a direct block of body, then get its index
  let element = cursor.getElement();
  while (element.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    element = element.getParent();
  }
  const startIndex = body.getChildIndex(element);

  const endIndex = Math.min(startIndex + count, childCount);

  for (let i = startIndex; i < endIndex; i++) {
    const block = body.getChild(i);
    if (block.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph  = block.asParagraph();
    const text       = paragraph.getText();
    const textLength = text.length;

    paragraph.editAsText().setBackgroundColor(COLORS.background);
    paragraph.editAsText().setForegroundColor(COLORS.defaultText);

    if (textLength === 0) continue;

    const runs = applyRulesToText(text, rules);
    if (runs.length === 0) continue;

    applyRunsToElement(paragraph.editAsText(), textLength, runs);
  }
}

//=============================================================================
// Individual entry points for each line-count option
//=============================================================================
function highlight1()   { highlightAtCursor(1);   }
function highlight5()   { highlightAtCursor(5);   }
function highlight10()  { highlightAtCursor(10);  }
function highlight25()  { highlightAtCursor(25);  }
function highlight50()  { highlightAtCursor(50);  }
function highlight100() { highlightAtCursor(100); }


//=============================================================================
// Add menu
//=============================================================================
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Syntax')
    .addItem('Apply Highlighting (whole doc)', 'highlightFullDocument')
    .addSeparator()
    .addItem('Highlight Next  1 Paragraph',  'highlight1')
    .addItem('Highlight Next  5 Paragraphs', 'highlight5')
    .addItem('Highlight Next 10 Paragraphs', 'highlight10')
    .addItem('Highlight Next 25 Paragraphs', 'highlight25')
    .addItem('Highlight Next 50 Paragraphs', 'highlight50')
    .addItem('Highlight Next 100 Paragraphs','highlight100')
    .addToUi();
}

