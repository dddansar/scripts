//=============================================================================
// File: google_docs_highlighting.js
//-----------------------------------------------------------------------------
// Description: This script is intended to be loaded in google docs to add
//              basic syntax matching for text files.
//              NOTE: To add/update scripts in google docs click on:
//                    Extensions → Apps Script → add changes → save → run →
//                    reload google docs page → you should see a new/updated
//                    menu called Syntax where you can select your options.
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
      textElement.setForegroundColor(run.start, run.end, run.color); // Uses API call
    }
  }
}

//=============================================================================
// Main, applies syntax highlighting to the entire document.
// Resets the document colors then re-highlights every paragraph in the body.
//=============================================================================
function highlightFullDocument() {
  const doc   = DocumentApp.getActiveDocument(); // Uses API call
  const body  = doc.getBody(); // Uses API call
  const rules = createRules();
  const childCount = body.getNumChildren(); // Uses API call

  body.setBackgroundColor(COLORS.background); // Uses API call
  body.setForegroundColor(COLORS.defaultText); // Uses API call
  //body.editAsText().setBold(false) // Uses API call
  //body.editAsText().setUnderline(false) // Uses API call
  //body.editAsText().setItalic(false) // Uses API call
  //body.editAsText().setStrikethrough(false); // Uses API call

  for (let i = 0; i < childCount; i++) {
    // const block = body.getChild(i); // Uses API call
    // if (block.getType() !== DocumentApp.ElementType.PARAGRAPH) continue; // Uses API call

    const paragraph  = body.getChild(i).asParagraph(); // Uses API call
    const text       = paragraph.getText(); // Uses API call
    const textLength = text.length;
    if (textLength === 0) continue;

    const runs = applyRulesToText(text, rules);

    // Skip paragraphs that need no changes (pure default text)
    // This alone eliminates most API calls for sparse documents
    if (runs.length === 0) continue;

    applyRunsToElement(paragraph.editAsText(), textLength, runs); // Uses API call
  }
}

//=============================================================================
// Highlight the next `count` paragraphs from the cursor or selection position
// Falls back to the start of the selection if no cursor (i.e. text is selected).
// Resets each paragraph's colors before reapplying, so stale colors are cleared.
//=============================================================================
function highlightAtCursor(count) {
  const doc   = DocumentApp.getActiveDocument(); // Uses API call
  const body  = doc.getBody(); // Uses API call
  const rules = createRules();
  const childCount = body.getNumChildren(); // Uses API call

  // getCursor() returns null when text is selected — fall back to selection start
  let startIndex;
  const cursor = doc.getCursor(); // Uses API call
  if (cursor) {
    let element = cursor.getElement(); // Uses API call
    while (element.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) { // Uses API call
      element = element.getParent(); // Uses API call
    }
    startIndex = body.getChildIndex(element); // Uses API call
  } else {
    const selRange = getSelectionParagraphRange(doc, body);
    if (!selRange) {
      DocumentApp.getUi().alert('Place your cursor anywhere or select valid text first.'); // Uses API call
      return;
    }
    startIndex = selRange.startIndex;
  }

  highlightParagraphRange(body, rules, startIndex, startIndex + count);
}

//=============================================================================
// Get selected paragraph index range from the current selection
// Returns {startIndex, endIndex} of body children covered by the selection,
// or null if there is no selection.
//=============================================================================
function getSelectionParagraphRange(doc, body) {
  const selection = doc.getSelection(); // Uses API call
  if (!selection) return null;

  const rangeElements = selection.getRangeElements(); // Uses API call
  if (!rangeElements || rangeElements.length === 0) return null;

  // Walk each range element up to a direct child of body
  let minIndex = Infinity;
  let maxIndex = -Infinity;

  for (const rangeEl of rangeElements) {
    let element = rangeEl.getElement(); // Uses API call
    while (element.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) { // Uses API call
      element = element.getParent(); // Uses API call
    }
    const idx = body.getChildIndex(element); // Uses API call
    if (idx < minIndex) minIndex = idx;
    if (idx > maxIndex) maxIndex = idx;
  }

  return { startIndex: minIndex, endIndex: maxIndex + 1 };
}

//=============================================================================
// Highlight a range of paragraphs by index [startIndex, endIndex)
// Applies syntax highlighting rules to each paragraph in range.
//=============================================================================
function highlightParagraphRange(body, rules, startIndex, endIndex) {
  const childCount = body.getNumChildren(); // Uses API call
  const safeEnd = Math.min(endIndex, childCount);

  for (let i = startIndex; i < safeEnd; i++) {
    const paragraph  = body.getChild(i).asParagraph(); // Uses API call
    const text       = paragraph.getText(); // Uses API call
    const textLength = text.length;
    if (textLength === 0) continue;

    // Reset to default color before reapplying so stale colors are cleared
    paragraph.editAsText().setForegroundColor(0, textLength - 1, COLORS.defaultText); // Uses API call

    const runs = applyRulesToText(text, rules);
    if (runs.length === 0) continue;

    applyRunsToElement(paragraph.editAsText(), textLength, runs); // Uses API call
  }
}

//=============================================================================
// Comment selected lines
// Prepends "// " to each selected paragraph, then re-highlights.
//=============================================================================
function commentSelection() {
  const C_COMMENT_PREFIX = '// ';
  const doc   = DocumentApp.getActiveDocument(); // Uses API call
  const body  = doc.getBody(); // Uses API call
  const rules = createRules();

  const range = getSelectionParagraphRange(doc, body);
  if (!range) {
    DocumentApp.getUi().alert('Select text first.'); // Uses API call
    return;
  }

  const childCount = body.getNumChildren(); // Uses API call
  const safeEnd = Math.min(range.endIndex, childCount);

  for (let i = range.startIndex; i < safeEnd; i++) {
    const paragraph = body.getChild(i).asParagraph(); // Uses API call
    const text = paragraph.getText(); // Uses API call
    if (text.length === 0) continue;

    // Only prepend if not already commented
    if (!text.startsWith(C_COMMENT_PREFIX)) {
      paragraph.editAsText().insertText(0, C_COMMENT_PREFIX); // Uses API call
    }
  }

  highlightParagraphRange(body, rules, range.startIndex, range.endIndex);
}

//=============================================================================
// Uncomment selected lines
// Removes leading "// " from each selected paragraph, then re-highlights.
//=============================================================================
function uncommentSelection() {
  const C_COMMENT_PREFIX = '// ';
  const C_COMMENT_LEN = C_COMMENT_PREFIX.length;
  const doc   = DocumentApp.getActiveDocument(); // Uses API call
  const body  = doc.getBody(); // Uses API call
  const rules = createRules();

  const range = getSelectionParagraphRange(doc, body);
  if (!range) {
    DocumentApp.getUi().alert('Select text first.'); // Uses API call
    return;
  }

  const childCount = body.getNumChildren(); // Uses API call
  const safeEnd = Math.min(range.endIndex, childCount);

  for (let i = range.startIndex; i < safeEnd; i++) {
    const paragraph = body.getChild(i).asParagraph(); // Uses API call
    const text = paragraph.getText(); // Uses API call
    if (!text.startsWith(C_COMMENT_PREFIX)) continue;

    paragraph.editAsText().deleteText(0, C_COMMENT_LEN - 1); // Uses API call
  }

  highlightParagraphRange(body, rules, range.startIndex, range.endIndex);
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
  DocumentApp.getUi() // Uses API call
    .createMenu('Syntax')
    .addItem('Highlight Entire Document', 'highlightFullDocument')
    .addSeparator()
    .addItem('Highlight Next  1 Paragraph',  'highlight1')
    .addItem('Highlight Next  5 Paragraphs', 'highlight5')
    .addItem('Highlight Next 10 Paragraphs', 'highlight10')
    .addItem('Highlight Next 25 Paragraphs', 'highlight25')
    .addItem('Highlight Next 50 Paragraphs', 'highlight50')
    .addItem('Highlight Next 100 Paragraphs','highlight100')
    .addSeparator()
    .addItem('Comment Selection',   'commentSelection')
    .addItem('Uncomment Selection', 'uncommentSelection')
    .addToUi();
}

