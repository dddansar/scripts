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


// ============================================================
// Color Constants
// ============================================================
const COLORS = {
  background:   '#FFFFFF',
  defaultText:  '#000000',
  AllSymbols:   '#FF8800',
  AllNumbers:   '#00CCCC',
  AllCaps:      '#4488FF',
  AllWebLinks:  '#C3B148',
  AllHLTodo:    '#7c7c00',
  AllHLNote:    '#00AA00',
  AllHLWarning: '#FF8800',
  AllComments:  '#b0b0b0',
  TxtTitles1:   '#7c7c00',
  TxtTitles2:   '#FF8C00',
  TxtTitles3:   '#00AA00',
  TxtTitles4:   '#4488FF',
};

// ============================================================
// Rules
// ============================================================
function getRules() {
  return [
    { group: 'AllSymbols',    pattern: /[&|*+\-^~?!%@#$``\/\\<>=!:;,.()\[\]{}]/g, bold: true },
    { group: 'AllNumbers',    pattern: /\d+/g, bold: true },
    { group: 'AllWebLinks',   pattern: /\bwww\.[a-zA-Z0-9.?!\-_=\/~@()]+/g, bold: true },
    { group: 'AllComments',   pattern: /^(\/\/|").*/gm },
    { group: 'AllCaps',       pattern: /\b[A-Z][A-Z0-9_]+s?\b/g, bold: true },
    { group: 'AllHLTodo',     pattern: /\b(TODO|VITODO)\b/g, bold: true },
    { group: 'AllHLNote',     pattern: /\b(NOTE|VINOTE)\b/g, bold: true },
    { group: 'AllHLWarning',  pattern: /\bWARNING\b/g, bold: true },
    { group: 'TxtTitles4',    pattern: /^\s*#### .*/gm, bold: true },
    { group: 'TxtTitles3',    pattern: /^\s*### .*/gm, bold: true },
    { group: 'TxtTitles2',    pattern: /^\s*## .*/gm, bold: true },
    { group: 'TxtTitles1',    pattern: /^\s*# .*/gm,    bold: true },
  ];
}

// ============================================================
// Build a compact run list from char arrays
// [ {start, end, color, bold}, ... ]  (only styled spans)
// ============================================================
function buildRuns(charColor, charBold, paraLen) {
  const runs = [];
  let i = 0;
  while (i < paraLen) {
    const color = charColor[i];
    const bold  = charBold[i];
    if (color === null && !bold) { i++; continue; }

    let j = i + 1;
    while (j < paraLen && charColor[j] === color && charBold[j] === bold) j++;
    runs.push({ start: i, end: j - 1, color, bold });
    i = j;
  }
  return runs;
}

// ============================================================
// Core: compute highlight runs for one paragraph text
// ============================================================
function computeRuns(paraText, rules) {
  const paraLen   = paraText.length;
  const charColor = new Array(paraLen).fill(null);
  const charBold  = new Array(paraLen).fill(false);

  for (const rule of rules) {
    const color = COLORS[rule.group];
    const bold  = rule.bold || false;
    rule.pattern.lastIndex = 0;

    let match;
    while ((match = rule.pattern.exec(paraText)) !== null) {
      const start = match.index;
      const end   = match.index + match[0].length;
      for (let i = start; i < end; i++) {
        charColor[i] = color;
        if (bold) charBold[i] = true;
      }
      if (match[0].length === 0) rule.pattern.lastIndex++;
    }
  }
  return buildRuns(charColor, charBold, paraLen);
}

// ============================================================
// Apply runs to a Text element with MINIMAL api calls
// Key insight: reset whole paragraph in ONE call, then only
// make calls for non-default styled spans.
// ============================================================
function applyRuns(textEl, paraLen, runs) {
  // Single call to reset entire paragraph
  textEl.setForegroundColor(0, paraLen - 1, COLORS.defaultText);
  textEl.setBold(0, paraLen - 1, false);

  // One call per styled run
  for (const run of runs) {
    if (run.color !== null) {
      textEl.setForegroundColor(run.start, run.end, run.color);
    }
    if (run.bold) {
      textEl.setBold(run.start, run.end, true);
    }
  }
}

// ============================================================
// Main
// ============================================================
function applySyntaxHighlighting() {
  const doc   = DocumentApp.getActiveDocument();
  const body  = doc.getBody();
  const rules = getRules();
  const n     = body.getNumChildren();

  body.setBackgroundColor(COLORS.background);

  for (let p = 0; p < n; p++) {
    const child = body.getChild(p);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const para     = child.asParagraph();
    const paraText = para.getText();
    const paraLen  = paraText.length;
    if (paraLen === 0) continue;

    const runs = computeRuns(paraText, rules);

    // Skip paragraphs that need no changes (pure default text)
    // This alone eliminates most API calls for sparse documents
    if (runs.length === 0) continue;

    applyRuns(para.editAsText(), paraLen, runs);
  }

  DocumentApp.getUi().alert('Done!');
}

// ============================================================
// Find the start index of the paragraph under the cursor
// ============================================================
function getCursorParagraphIndex(doc, body) {
  const cursor = doc.getCursor();
  if (!cursor) {
    DocumentApp.getUi().alert('Place your cursor in a paragraph first.');
    return -1;
  }

  const n = body.getNumChildren();

  // Walk up until we find a direct child of body
  let element = cursor.getElement();
  while (
    element &&
    element.getParent() &&
    element.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION
  ) {
    element = element.getParent();
  }
  if (!element) return 0;

  // Find the matching child index
  for (let i = 0; i < n; i++) {
    const child = body.getChild(i);
    if (child.getText && element.getText && child.getText() === element.getText()) {
      return i;
    }
  }

  return 0; // fallback
}

// ============================================================
// Highlight the next `count` paragraphs from the cursor position
// ============================================================
function highlightFromCursor(count) {
  const doc   = DocumentApp.getActiveDocument();
  const body  = doc.getBody();
  const rules = getRules();
  const n     = body.getNumChildren();

  const startIndex = getCursorParagraphIndex(doc, body);
  if (startIndex === -1) return;

  const endIndex = Math.min(startIndex + count, n);

  for (let i = startIndex; i < endIndex; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const para     = child.asParagraph();
    const paraText = para.getText();
    const paraLen  = paraText.length;
    if (paraLen === 0) continue;

    const runs = computeRuns(paraText, rules);
    if (runs.length === 0) continue;

    applyRuns(para.editAsText(), paraLen, runs);
  }
}

// ============================================================
// Individual entry points for each line-count option
// ============================================================
function highlight1()   { highlightFromCursor(1);   }
function highlight5()   { highlightFromCursor(5);   }
function highlight10()  { highlightFromCursor(10);  }
function highlight25()  { highlightFromCursor(25);  }
function highlight50()  { highlightFromCursor(50);  }
function highlight100() { highlightFromCursor(100); }

// ============================================================
// Add menu
// ============================================================
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Syntax')
    .addItem('Apply Highlighting (whole doc)', 'applySyntaxHighlighting')
    .addSeparator()
    .addItem('Highlight Next  1 Paragraph',  'highlight1')
    .addItem('Highlight Next  5 Paragraphs', 'highlight5')
    .addItem('Highlight Next 10 Paragraphs', 'highlight10')
    .addItem('Highlight Next 25 Paragraphs', 'highlight25')
    .addItem('Highlight Next 50 Paragraphs', 'highlight50')
    .addItem('Highlight Next 100 Paragraphs','highlight100')
    .addToUi();
}
