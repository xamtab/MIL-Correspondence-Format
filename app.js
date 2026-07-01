/**
 * MIL-Correspondence-Format - Main JS Logic
 * Developed by: xam, 72 FCC
 * Standard: JSSDM 2022
 */

// Global State Variables
let currentActiveTab = 'converter';
let letterReferences = [];
let letterParagraphs = [];
let selectedChangeSpan = null; // Clicked interactive span element
let parsedInputText = ""; // Holds raw output text before highlights

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initTheme();
    initConverter();
    initLetterBuilder();
    initServicePaperBuilder();
    initDictionary();
    
    // Load standard JSSDM sample by default
    loadTemplateSample();
});

/* ==========================================================================
   Tab Navigation Handler
   ========================================================================== */
function initTabs() {
    const navButtons = document.querySelectorAll('.nav-btn');
    const tabPanels = document.querySelectorAll('.tab-panel');
    
    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const tabName = btn.getAttribute('data-tab');
            
            // Toggle buttons
            navButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Toggle panels
            tabPanels.forEach(p => p.classList.remove('active'));
            document.getElementById(`tab-${tabName}`).classList.add('active');
            
            currentActiveTab = tabName;
        });
    });
}

/* ==========================================================================
   Theme Handler (Dark / Light Mode Toggle)
   ========================================================================== */
function initTheme() {
    const themeBtn = document.getElementById('theme-toggle');
    
    // Default theme check
    if (localStorage.getItem('theme') === 'light') {
        document.body.classList.remove('dark-mode');
        document.body.classList.add('light-mode');
        themeBtn.innerHTML = '<i class="fa-solid fa-moon"></i> <span>Dark Mode</span>';
    }
    
    themeBtn.addEventListener('click', () => {
        if (document.body.classList.contains('dark-mode')) {
            document.body.classList.remove('dark-mode');
            document.body.classList.add('light-mode');
            themeBtn.innerHTML = '<i class="fa-solid fa-moon"></i> <span>Dark Mode</span>';
            localStorage.setItem('theme', 'light');
        } else {
            document.body.classList.remove('light-mode');
            document.body.classList.add('dark-mode');
            themeBtn.innerHTML = '<i class="fa-solid fa-sun"></i> <span>Light Mode</span>';
            localStorage.setItem('theme', 'dark');
        }
    });
}

/* ==========================================================================
   Abbreviation & Conversion Engine (Tab 1)
   ========================================================================== */
function initConverter() {
    const inputArea = document.getElementById('converter-input');
    const searchInput = document.getElementById('converter-search');
    
    // Auto-update word count as you type
    inputArea.addEventListener('input', () => {
        updateWordCount();
        runAISuggestions();
    });
    
    // Action Buttons
    document.getElementById('btn-abbreviate').addEventListener('click', () => {
        convertText(true);
    });
    
    document.getElementById('btn-deabbreviate').addEventListener('click', () => {
        convertText(false);
    });
    
    document.getElementById('btn-clear').addEventListener('click', () => {
        inputArea.value = '';
        document.getElementById('converter-output').innerHTML = '<span class="placeholder-text">Result will appear here...</span>';
        updateWordCount(0);
        document.getElementById('change-count').innerText = '0';
        runAISuggestions();
    });
    
    // Quick search bar
    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.trim().toLowerCase();
        const resultsBox = document.getElementById('search-quick-result');
        resultsBox.innerHTML = '';
        
        if (query.length < 2) return;
        
        let matches = 0;
        
        // Search in abbreviations
        for (const [meaning, abbr] of Object.entries(MIL_ABBREVIATIONS)) {
            if (meaning.toLowerCase().includes(query) || abbr.toLowerCase().includes(query)) {
                const item = document.createElement('div');
                item.className = 'quick-result-item';
                
                const meanings = MIL_DEABBREVIATIONS[abbr.toLowerCase()] || [meaning];
                const meaningsDisplay = meanings.join(' / ');
                
                item.innerHTML = `
                    <span class="result-abbr">${abbr}</span>
                    <span class="result-meaning" title="${meaningsDisplay}">${meaningsDisplay}</span>
                `;
                
                // Add click behavior to append to input
                item.style.cursor = 'pointer';
                item.addEventListener('click', () => {
                    inputArea.value += (inputArea.value ? ' ' : '') + abbr;
                    searchInput.value = '';
                    resultsBox.innerHTML = '';
                    updateWordCount();
                    runAISuggestions();
                });
                
                resultsBox.appendChild(item);
                matches++;
                if (matches >= 6) break;
            }
        }
    });

    // Copy result
    document.getElementById('btn-copy-result').addEventListener('click', () => {
        const outputText = document.getElementById('converter-output').innerText;
        if (!outputText || outputText.includes('Result will appear here...')) {
            alert('No result to copy.');
            return;
        }
        navigator.clipboard.writeText(outputText).then(() => {
            alert('Converted text copied to clipboard!');
        });
    });
    
    // Download TXT
    document.getElementById('btn-download-txt').addEventListener('click', () => {
        const outputText = document.getElementById('converter-output').innerText;
        if (!outputText || outputText.includes('Result will appear here...')) return;
        downloadBlob(outputText, 'converted-military-text.txt', 'text/plain');
    });

    // Download Word (HTML wrap)
    document.getElementById('btn-download-word').addEventListener('click', () => {
        const outputText = document.getElementById('converter-output').innerText;
        if (!outputText || outputText.includes('Result will appear here...')) return;
        const paragraphs = outputText.split('\n').map(line => 
            `<p style="font-family:'Times New Roman',serif; font-size:12pt; margin:0 0 6pt 0; text-align:justify;">${line || '&nbsp;'}</p>`
        ).join('\n');
        const htmlContent = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>MIL-Correspondence-Format</title>
<!--[if gte mso 9]>
<xml>
<w:WordDocument>
<w:View>Print</w:View>
<w:Zoom>100</w:Zoom>
<w:DoNotOptimizeForBrowser/>
</w:WordDocument>
</xml>
<![endif]-->
<style>
    @page { size: 21cm 29.7cm; margin: 2.5cm; }
    body { font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.5; color: #000000; }
    p { font-family: 'Times New Roman', serif; text-align: justify; }
</style>
</head>
<body>
${paragraphs}
</body>
</html>`;
        downloadBlob(htmlContent, 'converted-military-text.doc', 'application/msword');
    });

    // Document File Upload
    initFileUploader();
}

// File Upload Handler (PDF, DOCX, TXT)
function initFileUploader() {
    const dropZone = document.getElementById('file-drop-zone');
    const fileInput = document.getElementById('file-input');
    
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = 'var(--active-accent)';
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.style.borderColor = 'var(--border-color)';
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = 'var(--border-color)';
        if (e.dataTransfer.files.length > 0) {
            handleFileUpload(e.dataTransfer.files[0]);
        }
    });
    
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            handleFileUpload(fileInput.files[0]);
        }
    });
}

function handleFileUpload(file) {
    const reader = new FileReader();
    const inputArea = document.getElementById('converter-input');
    
    if (file.name.endsWith('.txt')) {
        reader.onload = (e) => {
            inputArea.value = e.target.result;
            updateWordCount();
            runAISuggestions();
        };
        reader.readAsText(file);
    } 
    else if (file.name.endsWith('.docx')) {
        reader.onload = (e) => {
            const arrayBuffer = e.target.result;
            mammoth.extractRawText({ arrayBuffer: arrayBuffer })
                .then((result) => {
                    inputArea.value = result.value;
                    updateWordCount();
                    runAISuggestions();
                })
                .catch((err) => {
                    alert('Error reading DOCX file: ' + err.message);
                });
        };
        reader.readAsArrayBuffer(file);
    } 
    else if (file.name.endsWith('.pdf')) {
        reader.onload = (e) => {
            const typedarray = new Uint8Array(e.target.result);
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
            pdfjsLib.getDocument(typedarray).promise.then((pdf) => {
                let maxPages = pdf.numPages;
                let countPromises = [];
                
                for (let j = 1; j <= maxPages; j++) {
                    let page = pdf.getPage(j);
                    countPromises.push(page.then((page) => {
                        return page.getTextContent().then((text) => {
                            return text.items.map((s) => s.str).join(' ');
                        });
                    }));
                }
                
                Promise.all(countPromises).then((texts) => {
                    inputArea.value = texts.join('\n\n');
                    updateWordCount();
                    runAISuggestions();
                });
            }).catch(err => {
                alert('Error parsing PDF file: ' + err.message);
            });
        };
        reader.readAsArrayBuffer(file);
    }
}

function updateWordCount(value = null) {
    const input = document.getElementById('converter-input').value;
    const count = value !== null ? value : (input.trim() ? input.trim().split(/\s+/).length : 0);
    document.getElementById('word-count').innerText = count;
}

// Convert core text logic (to Abbr / from Abbr)
function convertText(toAbbreviate = true) {
    const rawText = document.getElementById('converter-input').value;
    const outputElement = document.getElementById('converter-output');
    
    if (!rawText.trim()) {
        outputElement.innerHTML = '<span class="placeholder-text">Please input military text to convert...</span>';
        return;
    }
    
    let processedText = rawText;
    let changeCount = 0;
    
    if (toAbbreviate) {
        // --- 1. ABBREVIATE CONVERSION (Full forms to Abbreviations) ---
        // Sort keys by length descending to match longer phrases first
        const sortedPhrases = Object.keys(MIL_ABBREVIATIONS).sort((a, b) => b.length - a.length);
        
        const replacements = [];
        
        sortedPhrases.forEach(phrase => {
            // Build regex with word boundary, handling special chars (like & and /)
            const escapedPhrase = phrase.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
            const regex = new RegExp(`\\b${escapedPhrase}\\b`, 'gi');
            
            processedText = processedText.replace(regex, (match) => {
                const replacementCode = MIL_ABBREVIATIONS[phrase];
                changeCount++;
                
                // Match original casing if possible
                let casedReplacement = replacementCode;
                if (match === match.toUpperCase()) {
                    casedReplacement = replacementCode.toUpperCase();
                } else if (match[0] === match[0].toUpperCase()) {
                    // Title Case
                    casedReplacement = replacementCode.charAt(0).toUpperCase() + replacementCode.slice(1);
                }
                
                // Create unique token to prevent nested replacements
                const token = `__TOK_${replacements.length}__`;
                replacements.push({
                    token: token,
                    html: `<span class="changed-term" data-original="${match}" data-term="${casedReplacement}" onclick="resolveAmbiguity(this, event)">${casedReplacement}</span>`
                });
                return token;
            });
        });
        
        // Restore all unique html badges
        replacements.forEach(rep => {
            processedText = processedText.replace(rep.token, rep.html);
        });
        
    } else {
        // --- 2. DEABBREVIATE CONVERSION (Abbreviations to Full forms) ---
        // Sort keys of abbreviations by length descending
        const sortedAbbrs = Object.keys(MIL_DEABBREVIATIONS).sort((a, b) => b.length - a.length);
        
        const replacements = [];
        
        sortedAbbrs.forEach(abbr => {
            const regex = new RegExp(`\\b${abbr}\\b`, 'gi');
            
            processedText = processedText.replace(regex, (match) => {
                const meanings = MIL_DEABBREVIATIONS[abbr.toLowerCase()];
                changeCount++;
                
                if (meanings.length === 1) {
                    // Single meaning - direct replacement
                    let textRepl = meanings[0];
                    // Match original casing
                    if (match === match.toUpperCase()) {
                        textRepl = textRepl.toUpperCase();
                    } else if (match[0] === match[0].toUpperCase()) {
                        textRepl = textRepl.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
                    }
                    
                    const token = `__TOK_${replacements.length}__`;
                    replacements.push({
                        token: token,
                        html: `<span class="changed-term" data-original="${match}" data-term="${textRepl}" onclick="resolveAmbiguity(this, event)">${textRepl}</span>`
                    });
                    return token;
                } else {
                    // Multiple meanings - mark as ambiguous
                    const token = `__TOK_${replacements.length}__`;
                    replacements.push({
                        token: token,
                        html: `<span class="changed-term ambiguous" data-original="${match}" data-term="${match}" data-code="${abbr}" onclick="resolveAmbiguity(this, event)">${match}</span>`
                    });
                    return token;
                }
            });
        });
        
        // Restore unique html tokens
        replacements.forEach(rep => {
            processedText = processedText.replace(rep.token, rep.html);
        });
    }
    
    // Set HTML result
    outputElement.innerHTML = processedText;
    document.getElementById('change-count').innerText = changeCount;
}

// Ambiguity Popover Selector
function resolveAmbiguity(spanElement, event) {
    event.stopPropagation();
    selectedChangeSpan = spanElement;
    
    const popup = document.getElementById('ambiguous-popup-menu');
    const termCodeEl = document.getElementById('popup-term-code');
    const optionsList = document.getElementById('popup-options-list');
    const customInput = document.getElementById('popup-custom-text');
    
    customInput.value = '';
    
    // Determine whether this was an abbreviation de-abbr or original translation
    const code = spanElement.getAttribute('data-code') || spanElement.innerText;
    const cleanCode = code.toLowerCase().trim().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"");
    
    const meanings = MIL_DEABBREVIATIONS[cleanCode] || [];
    
    termCodeEl.innerText = code;
    optionsList.innerHTML = '';
    
    if (meanings.length > 0) {
        meanings.forEach(meaning => {
            const btn = document.createElement('button');
            btn.className = 'popup-option-btn';
            btn.innerText = meaning;
            btn.addEventListener('click', () => {
                applyTermSelection(meaning);
            });
            optionsList.appendChild(btn);
        });
    } else {
        const item = document.createElement('p');
        item.style.fontSize = '12px';
        item.style.color = 'var(--text-muted)';
        item.innerText = 'No multiple meanings found for code.';
        optionsList.appendChild(item);
    }
    
    // Positioning popover cleanly relative to span
    const rect = spanElement.getBoundingClientRect();
    popup.style.top = `${rect.bottom + window.scrollY + 6}px`;
    popup.style.left = `${Math.min(rect.left + window.scrollX, window.innerWidth - 340)}px`;
    popup.style.display = 'flex';
    
    // Click outside listener
    document.addEventListener('click', closeAmbiguityPopupOutside);
}

function applyTermSelection(term) {
    if (!selectedChangeSpan) return;
    
    selectedChangeSpan.innerText = term;
    selectedChangeSpan.classList.remove('ambiguous');
    selectedChangeSpan.removeAttribute('data-code');
    
    // Increment changes if needed
    closeAmbiguityPopup();
}

// Custom text resolver apply button
document.getElementById('btn-popup-apply-custom').addEventListener('click', () => {
    const textVal = document.getElementById('popup-custom-text').value.trim();
    if (textVal) {
        applyTermSelection(textVal);
    }
});

function closeAmbiguityPopup() {
    const popup = document.getElementById('ambiguous-popup-menu');
    popup.style.display = 'none';
    document.removeEventListener('click', closeAmbiguityPopupOutside);
    selectedChangeSpan = null;
}

function closeAmbiguityPopupOutside(e) {
    const popup = document.getElementById('ambiguous-popup-menu');
    if (!popup.contains(e.target) && !e.target.classList.contains('changed-term')) {
        closeAmbiguityPopup();
    }
}

/* ==========================================================================
   AI Correspondence Assistant Logic
   ========================================================================== */
const AI_RULES = [
    {
        pattern: /was (\w+ed) by/i,
        replace: "Active Voice Construction",
        type: "passive",
        reason: "JSSDM Section 1 Para 0107a(2) orders: 'Use the active rather than the passive construction of the verb.'"
    },
    {
        pattern: /were (\w+ed) by/i,
        replace: "Active Voice Construction",
        type: "passive",
        reason: "JSSDM Section 1 Para 0107a(2) orders: 'Use the active rather than the passive construction of the verb.'"
    },
    {
        pattern: /at this point in time/i,
        replace: "now",
        type: "flowery",
        reason: "JSSDM Section 1 Para 0107a(8) warns: 'Avoid jargon and officialese. Replace flowery expressions with simple terms.'"
    },
    {
        pattern: /prior to/i,
        replace: "before",
        type: "flowery",
        reason: "Replace officialese 'Prior to' with 'Before' (JSSDM Section 1 Para 0107a(8))."
    },
    {
        pattern: /subsequent to/i,
        replace: "after",
        type: "flowery",
        reason: "Replace officialese 'Subsequent to' with 'After' (JSSDM Section 1 Para 0107a(8))."
    },
    {
        pattern: /please refer to/i,
        replace: "refer to",
        type: "tone",
        reason: "Avoid redundant prepositions. Use direct action words in official military correspondence."
    },
    {
        pattern: /endeavour to commence/i,
        replace: "try to begin",
        type: "flowery",
        reason: "Replace 'Endeavour to commence' with 'Try to begin' for clear communication (JSSDM Section 1 Para 0107a(8))."
    },
    {
        pattern: /this is a matter of very considerable urgency/i,
        replace: "this is urgent",
        type: "flowery",
        reason: "Short, direct phrasing is mandatory in staff duties (JSSDM Section 1 Para 0107a(8))."
    },
    {
        pattern: /having regard to the fact/i,
        replace: "as",
        type: "flowery",
        reason: "Shorten wordy officialese phrases (JSSDM Section 1 Para 0107a(8))."
    }
];

function runAISuggestions() {
    const text = document.getElementById('converter-input').value;
    const noSugg = document.getElementById('ai-no-suggestions');
    const listSugg = document.getElementById('ai-suggestions-list');
    
    listSugg.innerHTML = '';
    
    if (!text.trim()) {
        noSugg.style.display = 'block';
        listSugg.style.display = 'none';
        return;
    }
    
    let foundCount = 0;
    
    AI_RULES.forEach((rule, idx) => {
        let match;
        // Run match to find issues in text
        if (rule.pattern.test(text)) {
            foundCount++;
            
            // Try to extract matching text for displaying
            const matches = text.match(rule.pattern);
            const originalMatch = matches ? matches[0] : rule.pattern.toString();
            
            const card = document.createElement('div');
            card.className = 'suggestion-card';
            
            let typeLabel = "Grammar Rule";
            if (rule.type === 'passive') typeLabel = "Passive Voice";
            if (rule.type === 'flowery') typeLabel = "Officialese Jargon";
            if (rule.type === 'tone') typeLabel = "Correspondence Tone";
            
            card.innerHTML = `
                <div class="suggestion-content">
                    <span class="suggestion-type ${rule.type}">${typeLabel}</span>
                    <div class="suggestion-diff">
                        <span class="original-phrase">${originalMatch}</span>
                        <span class="arrow-separator">&rarr;</span>
                        <span class="correction-phrase">${rule.replace}</span>
                    </div>
                    <div class="suggestion-reason">${rule.reason}</div>
                </div>
                <button class="btn btn-secondary btn-sm" onclick="applyAISuggestion(${idx}, '${originalMatch.replace(/'/g, "\\'")}', '${rule.replace}')">Apply</button>
            `;
            
            listSugg.appendChild(card);
        }
    });
    
    if (foundCount > 0) {
        noSugg.style.display = 'none';
        listSugg.style.display = 'flex';
    } else {
        noSugg.style.display = 'block';
        noSugg.innerHTML = '<i class="fa-solid fa-circle-check" style="color:var(--success-color)"></i> Excellent! No JSSDM syntax violations or passive voice detected.';
        listSugg.style.display = 'none';
    }
}

function applyAISuggestion(ruleIdx, originalMatch, replacementText) {
    const inputArea = document.getElementById('converter-input');
    let text = inputArea.value;
    
    const rule = AI_RULES[ruleIdx];
    
    // Swap text
    if (rule.type === 'passive') {
        // Special case parsing passive voice "A was B-ed by C" to "C B-ed A"
        // Regex: (.*) was (.*) by (.*)
        const regexPassive = new RegExp(`(.*)\\b(was|were)\\b\\s+(\\w+ed)\\s+by\\s+([^\\n\\.,]*)`, 'i');
        text = text.replace(regexPassive, (m, g1, g2, g3, g4) => {
            return `${g4.trim()} ${g3.trim()} ${g1.trim()}`;
        });
    } else {
        // Regular direct string replacements
        const regex = new RegExp(originalMatch, 'i');
        text = text.replace(regex, replacementText);
    }
    
    inputArea.value = text;
    updateWordCount();
    runAISuggestions();
    
    // Automatically re-run Abbreviation conversion to reflect changes
    convertText(true);
}

/* ==========================================================================
   Routine Letter Template Builder (Tab 2)
   ========================================================================== */
function initLetterBuilder() {
    // Bind change listeners to input tags to update the live preview panel
    const inputIds = [
        'let-security', 'let-precedence', 'let-copy', 'let-pages',
        'let-sender', 'let-ref', 'let-date', 'let-subject',
        'let-sig-name', 'let-sig-rank', 'let-sig-appt',
        'let-distr-action', 'let-distr-info', 'let-distr-int-action', 'let-distr-int-info'
    ];
    
    inputIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', updateLetterPreview);
            el.addEventListener('change', updateLetterPreview);
        }
    });
    
    // Print triggers
    document.getElementById('btn-print-letter').addEventListener('click', () => {
        window.print();
    });
    
    // Copy letter text
    document.getElementById('btn-copy-letter-text').addEventListener('click', () => {
        const text = compileRawLetterText();
        navigator.clipboard.writeText(text).then(() => {
            alert('Compiled routine letter text copied!');
        });
    });
    
    // Download Word doc
    document.getElementById('btn-download-word-letter').addEventListener('click', () => {
        const html = buildWordDocHTML();
        downloadBlob(html, 'military-routine-letter.doc', 'application/msword');
    });
}

function toggleEditorSection(element) {
    const sections = document.querySelectorAll('.editor-section .section-content');
    const titles = document.querySelectorAll('.editor-section .section-title');
    const content = element.nextElementSibling;
    
    const isShowing = content.classList.contains('show');
    
    // Collapse all
    sections.forEach(s => s.classList.remove('show'));
    titles.forEach(t => {
        t.classList.remove('active');
        t.querySelector('i').className = 'fa-solid fa-angle-down';
    });
    
    // Toggle clicked
    if (!isShowing) {
        content.classList.add('show');
        element.classList.add('active');
        element.querySelector('i').className = 'fa-solid fa-angle-up';
    }
}

// Letter References management
function addLetterReference(val = '') {
    const container = document.getElementById('refs-container');
    const index = letterReferences.length;
    const letter = String.fromCharCode(65 + index); // A, B, C, D...
    
    const row = document.createElement('div');
    row.className = 'dynamic-item-row';
    row.id = `let-ref-row-${index}`;
    row.innerHTML = `
        <span style="font-weight:700; width:15px;">${letter}.</span>
        <input type="text" value="${val}" placeholder="Reference document details..." oninput="updateRefValue(${index}, this.value)">
        <button class="btn-remove" onclick="removeLetterReference(${index})"><i class="fa-solid fa-trash-can"></i></button>
    `;
    
    container.appendChild(row);
    letterReferences.push(val);
    updateLetterPreview();
}

function updateRefValue(index, val) {
    letterReferences[index] = val;
    updateLetterPreview();
}

function removeLetterReference(index) {
    const row = document.getElementById(`let-ref-row-${index}`);
    if (row) row.remove();
    
    letterReferences.splice(index, 1);
    
    // Re-render editor list to correct lettering sequence
    rebuildRefEditor();
}

function rebuildRefEditor() {
    const container = document.getElementById('refs-container');
    const tempValues = [...letterReferences];
    container.innerHTML = '';
    letterReferences = [];
    
    tempValues.forEach(val => {
        addLetterReference(val);
    });
}

// Letter Paragraph builder management
function addLetterParagraph(textVal = '', subParas = []) {
    const index = letterParagraphs.length;
    const paraNum = index + 1;
    
    const container = document.getElementById('paras-container');
    
    const card = document.createElement('div');
    card.className = 'para-block';
    card.id = `para-block-${index}`;
    
    card.innerHTML = `
        <div class="para-header">
            <span class="para-number-label">Paragraph ${paraNum}</span>
            <button class="btn-remove" onclick="removeLetterParagraph(${index})"><i class="fa-solid fa-trash-can"></i></button>
        </div>
        <div class="form-group">
            <input type="text" class="para-main-text" value="${textVal}" placeholder="Paragraph body text..." oninput="updateParaText(${index}, this.value)">
        </div>
        <div class="sub-paras-container" id="sub-paras-container-${index}">
            <!-- Sub-paragraphs loaded dynamically -->
        </div>
        <button class="btn btn-secondary btn-sm mt-1" onclick="addSubParagraph(${index})"><i class="fa-solid fa-plus"></i> Add Sub-Para</button>
    `;
    
    container.appendChild(card);
    letterParagraphs.push({
        text: textVal,
        subParas: subParas
    });
    
    // Rebuild sub paragraphs if loaded from template
    if (subParas.length > 0) {
        rebuildSubParaEditor(index);
    }
    
    updateLetterPreview();
}

function updateParaText(index, val) {
    letterParagraphs[index].text = val;
    updateLetterPreview();
}

function removeLetterParagraph(index) {
    letterParagraphs.splice(index, 1);
    rebuildParagraphEditor();
}

function rebuildParagraphEditor() {
    const container = document.getElementById('paras-container');
    container.innerHTML = '';
    const temp = [...letterParagraphs];
    letterParagraphs = [];
    
    temp.forEach(p => {
        addLetterParagraph(p.text, p.subParas);
    });
}

// Sub-paragraphs management
function addSubParagraph(paraIndex, subText = '') {
    const para = letterParagraphs[paraIndex];
    const subIndex = para.subParas.length;
    const subLetter = String.fromCharCode(97 + subIndex); // a, b, c... (JSSDM standard)
    
    const container = document.getElementById(`sub-paras-container-${paraIndex}`);
    
    const row = document.createElement('div');
    row.className = 'sub-para-row';
    row.id = `sub-para-row-${paraIndex}-${subIndex}`;
    
    row.innerHTML = `
        <span class="sub-para-label">${subLetter}.</span>
        <input type="text" value="${subText}" placeholder="Sub-paragraph text..." style="flex-grow:1; font-size:12px; padding:6px 10px;" oninput="updateSubParaText(${paraIndex}, ${subIndex}, this.value)">
        <button class="btn-remove" onclick="removeSubParagraph(${paraIndex}, ${subIndex})"><i class="fa-solid fa-xmark" style="font-size:12px;"></i></button>
    `;
    
    container.appendChild(row);
    
    if (subText === '') {
        para.subParas.push('');
    }
    updateLetterPreview();
}

function updateSubParaText(paraIndex, subIndex, val) {
    letterParagraphs[paraIndex].subParas[subIndex] = val;
    updateLetterPreview();
}

function removeSubParagraph(paraIndex, subIndex) {
    letterParagraphs[paraIndex].subParas.splice(subIndex, 1);
    rebuildSubParaEditor(paraIndex);
}

function rebuildSubParaEditor(paraIndex) {
    const container = document.getElementById(`sub-paras-container-${paraIndex}`);
    container.innerHTML = '';
    const temp = [...letterParagraphs[paraIndex].subParas];
    letterParagraphs[paraIndex].subParas = [];
    
    temp.forEach(subText => {
        addSubParagraph(paraIndex, subText);
    });
    updateLetterPreview();
}

// Letter Live compiler mapping variables to Live paper page preview
function updateLetterPreview() {
    // 1. Security & Precedence
    const sec = document.getElementById('let-security').value;
    document.getElementById('p-top-security').innerText = sec;
    document.getElementById('p-bottom-security').innerText = sec;
    
    const prec = document.getElementById('let-precedence').value;
    const precEl = document.getElementById('p-precedence');
    if (prec) {
        precEl.innerText = prec;
        precEl.style.display = 'block';
    } else {
        precEl.style.display = 'none';
    }
    
    // Copy & Pages
    const copyVal = document.getElementById('let-copy').value;
    document.getElementById('p-copy').innerText = copyVal ? `Copy No ${copyVal}` : '';
    
    const pagesVal = document.getElementById('let-pages').value;
    document.getElementById('p-pages').innerText = pagesVal ? `Total Pages ${pagesVal}` : '';
    
    // Sender Address
    const sender = document.getElementById('let-sender').value;
    document.getElementById('p-sender').innerHTML = sender.replace(/\n/g, '<br>');
    
    // Reference & Date
    document.getElementById('p-ref').innerText = document.getElementById('let-ref').value;
    document.getElementById('p-date').innerText = document.getElementById('let-date').value;
    
    // Subject Heading
    const subject = document.getElementById('let-subject').value.trim();
    const subjEl = document.getElementById('p-subject');
    if (subject) {
        subjEl.innerText = subject;
        subjEl.style.display = 'block';
    } else {
        subjEl.style.display = 'none';
    }
    
    // References
    const refContainer = document.getElementById('p-refs');
    if (letterReferences.length > 0) {
        refContainer.style.display = 'block';
        // Use singular "Ref:" for one reference, "Refs:" for multiple
        const refLabel = refContainer.querySelector('.ref-label');
        refLabel.innerText = letterReferences.length === 1 ? 'Ref:' : 'Refs:';
        const listEl = refContainer.querySelector('.ref-list');
        listEl.innerHTML = '';
        letterReferences.forEach((ref, idx) => {
            const letter = String.fromCharCode(65 + idx);
            const item = document.createElement('div');
            item.className = 'ref-item';
            item.innerHTML = `${letter}.&nbsp;&nbsp;&nbsp;&nbsp;${ref}`;
            listEl.appendChild(item);
        });
    } else {
        refContainer.style.display = 'none';
    }
    
    // Paragraphs (Body)
    const bodyEl = document.getElementById('p-body');
    bodyEl.innerHTML = '';
    
    letterParagraphs.forEach((para, idx) => {
        const paraEl = document.createElement('div');
        paraEl.className = 'preview-para-container';
        
        // Main Paragraph text
        const mainRow = document.createElement('div');
        mainRow.className = 'preview-para';
        let formattedMain = para.text.replace(/^([^\.]{2,50}\.)\s/, '$1&nbsp;&nbsp;&nbsp;');
        mainRow.innerHTML = `
            <span class="preview-para-num">${idx + 1}.&nbsp;&nbsp;&nbsp;</span><span class="preview-para-content">${formattedMain}</span>
        `;
        paraEl.appendChild(mainRow);
        
        // Nested sub-paragraphs
        if (para.subParas.length > 0) {
            para.subParas.forEach((subText, subIdx) => {
                const subRow = document.createElement('div');
                subRow.className = 'preview-sub-para';
                const subLetter = String.fromCharCode(97 + subIdx);
                let formattedSub = subText.replace(/^([^\.]{2,50}\.)\s/, '$1&nbsp;&nbsp;&nbsp;');
                subRow.innerHTML = `
                    <span class="preview-sub-para-num">${subLetter}.&nbsp;&nbsp;&nbsp;</span><span class="preview-para-content">${formattedSub}</span>
                `;
                paraEl.appendChild(subRow);
            });
        }
        
        bodyEl.appendChild(paraEl);
    });
    
    // Signature block
    document.getElementById('p-sig-name').innerText = document.getElementById('let-sig-name').value.toUpperCase();
    document.getElementById('p-sig-rank').innerText = document.getElementById('let-sig-rank').value;
    document.getElementById('p-sig-appt').innerText = document.getElementById('let-sig-appt').value;
    
    // Distribution Block
    const distEl = document.getElementById('p-distr');
    distEl.innerHTML = '';
    
    const act = document.getElementById('let-distr-action').value.trim();
    const info = document.getElementById('let-distr-info').value.trim();
    const intAct = document.getElementById('let-distr-int-action').value.trim();
    const intInfo = document.getElementById('let-distr-int-info').value.trim();
    
    if (act || info || intAct || intInfo) {
        distEl.style.display = 'block';
        distEl.innerHTML += '<div class="dist-section-header">Distr:</div>';
        
        // Build 2-column or list view
        const grid = document.createElement('div');
        grid.className = 'dist-col-grid';
        
        // Column 1: External Action / Info
        const col1 = document.createElement('div');
        if (act || info) {
            col1.innerHTML += '<div class="dist-sub-section">Extl:</div>';
            if (act) {
                const actDiv = document.createElement('div');
                actDiv.className = 'dist-list';
                actDiv.innerHTML = '<span>Act:</span>';
                act.split('\n').forEach(line => {
                    actDiv.innerHTML += `<div class="dist-item">${line}</div>`;
                });
                col1.appendChild(actDiv);
            }
            if (info) {
                const infoDiv = document.createElement('div');
                infoDiv.className = 'dist-list';
                infoDiv.innerHTML = '<span>Info:</span>';
                info.split('\n').forEach(line => {
                    infoDiv.innerHTML += `<div class="dist-item">${line}</div>`;
                });
                col1.appendChild(infoDiv);
            }
        }
        grid.appendChild(col1);
        
        // Column 2: Internal Action / Info
        const col2 = document.createElement('div');
        if (intAct || intInfo) {
            col2.innerHTML += '<div class="dist-sub-section">Intl:</div>';
            if (intAct) {
                const intActDiv = document.createElement('div');
                intActDiv.className = 'dist-list';
                intActDiv.innerHTML = '<span>Act:</span>';
                intAct.split('\n').forEach(line => {
                    intActDiv.innerHTML += `<div class="dist-item">${line}</div>`;
                });
                col2.appendChild(intActDiv);
            }
            if (intInfo) {
                const intInfoDiv = document.createElement('div');
                intInfoDiv.className = 'dist-list';
                intInfoDiv.innerHTML = '<span>Info:</span>';
                intInfo.split('\n').forEach(line => {
                    intInfoDiv.innerHTML += `<div class="dist-item">${line}</div>`;
                });
                col2.appendChild(intInfoDiv);
            }
        }
        grid.appendChild(col2);
        
        distEl.appendChild(grid);
    } else {
        distEl.style.display = 'none';
    }
}

// Resets letter forms
function resetLetterBuilder() {
    if (confirm('Are you sure you want to clear the Routine Letter Builder?')) {
        document.getElementById('let-copy').value = '';
        document.getElementById('let-pages').value = '';
        document.getElementById('let-subject').value = '';
        document.getElementById('let-sig-name').value = '';
        document.getElementById('let-sig-rank').value = '';
        document.getElementById('let-sig-appt').value = '';
        document.getElementById('let-distr-action').value = '';
        document.getElementById('let-distr-info').value = '';
        document.getElementById('let-distr-int-action').value = '';
        document.getElementById('let-distr-int-info').value = '';
        
        letterReferences = [];
        document.getElementById('refs-container').innerHTML = '';
        
        letterParagraphs = [];
        document.getElementById('paras-container').innerHTML = '';
        
        updateLetterPreview();
    }
}

// Compile a plain text string representation of the letter layout
function compileRawLetterText() {
    const sec = document.getElementById('let-security').value;
    const prec = document.getElementById('let-precedence').value;
    const copyVal = document.getElementById('let-copy').value;
    const pagesVal = document.getElementById('let-pages').value;
    const sender = document.getElementById('let-sender').value;
    const ref = document.getElementById('let-ref').value;
    const date = document.getElementById('let-date').value;
    const subject = document.getElementById('let-subject').value;
    
    let text = `\t\t\t\t\t\t${sec}\n\n`;
    
    if (prec) text += `\t\t\t\t\t\t\t${prec}\n`;
    if (copyVal) text += `\t\t\t\t\t\t\tCopy No ${copyVal}\n`;
    if (pagesVal) text += `\t\t\t\t\t\t\tTotal Pages ${pagesVal}\n`;
    
    text += `${sender.split('\n').map(l => '\t\t\t\t\t\t\t' + l).join('\n')}\n\n`;
    text += `${ref}\t\t\t\t\t\t${date}\n\n`;
    text += `${subject.toUpperCase()}\n\n\n`;
    
    if (letterReferences.length > 0) {
        text += `Refs:\n`;
        letterReferences.forEach((r, i) => {
            const letter = String.fromCharCode(65 + i);
            text += `${letter}.\t${r}\n`;
        });
        text += `\n\n`;
    }
    
    letterParagraphs.forEach((p, idx) => {
        let formattedMain = p.text.replace(/^([^\.]{2,50}\.)\s/, '$1\t');
        text += `${idx + 1}.\t${formattedMain}\n\n`;
        p.subParas.forEach((sub, subIdx) => {
            const letter = String.fromCharCode(97 + subIdx);
            let formattedSub = sub.replace(/^([^\.]{2,50}\.)\s/, '$1\t');
            text += `\t${letter}.\t${formattedSub}\n`;
        });
        text += `\n`;
    });
    
    const sigName = document.getElementById('let-sig-name').value;
    const sigRank = document.getElementById('let-sig-rank').value;
    const sigAppt = document.getElementById('let-sig-appt').value;
    
    text += `\n\t\t\t\t\t\t\t${sigName.toUpperCase()}\n`;
    text += `\t\t\t\t\t\t\t${sigRank}\n`;
    text += `\t\t\t\t\t\t\t${sigAppt}\n\n`;
    
    text += `Distribution:\n`;
    text += `Action:\n${document.getElementById('let-distr-action').value}\n`;
    text += `Information:\n${document.getElementById('let-distr-info').value}\n`;
    
    text += `\n\t\t\t\t${sec}\n`;
    return text;
}

// Populates builder fields with sample memo standard (from JSP-001)
function loadTemplateSample() {
    document.getElementById('let-security').value = 'RESTRICTED';
    document.getElementById('let-precedence').value = 'IMMEDIATE';
    document.getElementById('let-copy').value = '1 of 27';
    document.getElementById('let-pages').value = '2';
    
    document.getElementById('let-sender').value = 'AHQ\nGS Br\nMT Dte\nDhaka Cantt\nTel: 9110341 ext: 5122';
    document.getElementById('let-ref').value = '06.02.2626.121.55.003.22';
    document.getElementById('let-date').value = 'Sep 22';
    
    document.getElementById('let-subject').value = 'DEMO-BDE HQ AND BAA IN THE FD';
    
    // Reset lists
    letterReferences = [];
    document.getElementById('refs-container').innerHTML = '';
    addLetterReference('DSCSC ltr no 06.02.2626.123.65.013.22 dt 30 Jul 22.');
    addLetterReference('Air HQ ltr no 06.02.2626.121.54.007.22 dt 04 Aug 22.');
    
    letterParagraphs = [];
    document.getElementById('paras-container').innerHTML = '';
    
    addLetterParagraph('For the 25th Army Staff Course, 19th Navy Staff Course and 21st Air Staff Course, DSCSC has req for a demo of a mob and static bde HQ along with the layout of a BAA at Savar Cantt. Tentatively, the demo is scheduled for 23 Sep from 0930-1400 hrs with dress rehearsal on 22 Sep.', []);
    
    addLetterParagraph('The broad reqrs are:', [
        'Bde HQ. As detailed in Annex A.',
        'BAA. Layout per Annex B.',
        'Army/Air Coop. Air HQ has detailed a BASO with FFR vehicle.'
    ]);
    
    addLetterParagraph('For greater details on the demo see Ref A. Air HQ has detailed a BASO with an FFR veh and a Bell-212 hel from BAF BSR for the demo.', []);
    addLetterParagraph('DSCSC will liaise with your HQ for further coord.', []);
    
    document.getElementById('let-sig-name').value = 'MD HUMAYUN KABIR';
    document.getElementById('let-sig-rank').value = 'Col';
    document.getElementById('let-sig-appt').value = 'For CGS';
    
    document.getElementById('let-distr-action').value = 'HQ 9 Inf Div\nCommander BN Fleet\nAOC BAF BBD';
    document.getElementById('let-distr-info').value = 'HQ 81 Inf Bde\nAir HQ (Ops & Trg Br)\nDSCSC';
    
    document.getElementById('let-distr-int-action').value = 'AHQ, GS Br (SD Dte)';
    document.getElementById('let-distr-int-info').value = 'MO Dte\nMI Dte';
    
    updateLetterPreview();
}

/* ==========================================================================
   Reference Dictionary Browser (Tab 3)
   ========================================================================== */
function initDictionary() {
    const searchInput = document.getElementById('dict-search-input');
    const filterSelect = document.getElementById('dict-filter');
    const tableBody = document.getElementById('dict-table-body');
    const alphaContainer = document.querySelector('.alphabet-filter');
    
    let activeLetterFilter = 'all';
    
    // Generate A-Z filter buttons
    for (let i = 65; i <= 90; i++) {
        const char = String.fromCharCode(i);
        const btn = document.createElement('button');
        btn.className = 'letter-btn';
        btn.innerText = char;
        btn.setAttribute('data-letter', char.toLowerCase());
        
        btn.addEventListener('click', () => {
            document.querySelectorAll('.letter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            activeLetterFilter = char.toLowerCase();
            renderDictionary();
        });
        
        alphaContainer.appendChild(btn);
    }
    
    // Bind all buttons to 'All' filter reset
    document.querySelector('.letter-btn[data-letter="all"]').addEventListener('click', (e) => {
        document.querySelectorAll('.letter-btn').forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        activeLetterFilter = 'all';
        renderDictionary();
    });
    
    searchInput.addEventListener('input', renderDictionary);
    filterSelect.addEventListener('change', renderDictionary);
    
    // Render initially
    renderDictionary();
}

function renderDictionary() {
    const tableBody = document.getElementById('dict-table-body');
    const query = document.getElementById('dict-search-input').value.trim().toLowerCase();
    const filter = document.getElementById('dict-filter').value;
    const alphabetFilter = document.querySelector('.alphabet-filter .active').getAttribute('data-letter');
    
    tableBody.innerHTML = '';
    
    let count = 0;
    
    // Sort keys alphabetically
    const sortedMeanings = Object.keys(MIL_ABBREVIATIONS).sort();
    
    for (const meaning of sortedMeanings) {
        const abbr = MIL_ABBREVIATIONS[meaning];
        const lowerMeaning = meaning.toLowerCase();
        const lowerAbbr = abbr.toLowerCase();
        
        // 1. Query search filter
        if (query && !lowerMeaning.includes(query) && !lowerAbbr.includes(query)) {
            continue;
        }
        
        // 2. Alphabet filter
        if (alphabetFilter !== 'all' && !lowerMeaning.startsWith(alphabetFilter) && !lowerAbbr.startsWith(alphabetFilter)) {
            continue;
        }
        
        // 3. Ambiguous meanings filter
        const meanings = MIL_DEABBREVIATIONS[lowerAbbr] || [];
        const isMultiple = meanings.length > 1;
        
        if (filter === 'ambiguous' && !isMultiple) {
            continue;
        }
        
        const row = document.createElement('tr');
        
        const badgeClass = isMultiple ? 'multiple' : 'general';
        const badgeText = isMultiple ? `${meanings.length} Meanings` : 'JSSDM Std';
        
        row.innerHTML = `
            <td>
                <div style="font-weight: 500;">${meaning.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ')}</div>
                <div style="font-size: 11px; color: var(--text-muted); margin-top: 4px;">${meanings.join(' / ')}</div>
            </td>
            <td><strong>${abbr}</strong></td>
            <td><span class="dict-badge ${badgeClass}">${badgeText}</span></td>
        `;
        
        tableBody.appendChild(row);
        count++;
        if (count >= 150) { // Limit results to prevent DOM slowdown
            const capRow = document.createElement('tr');
            capRow.innerHTML = `<td colspan="3" style="text-align:center; color:var(--text-muted); font-style:italic;">Showing first 150 JSSDM database matches... Refine search for more.</td>`;
            tableBody.appendChild(capRow);
            break;
        }
    }
    
    if (count === 0) {
        tableBody.innerHTML = `<tr><td colspan="3" style="text-align:center; color:var(--text-muted); padding: 40px 0;"><i class="fa-solid fa-face-frown" style="font-size:24px; display:block; margin-bottom:10px;"></i>No standard JSSDM abbreviations match your filters.</td></tr>`;
    }
}

/* ==========================================================================
   Word Document Export Builder
   Generates Word-compatible HTML with inline styles and tables
   (Word does not support flexbox/grid, so we use tables for layout)
   ========================================================================== */
function buildWordDocHTML() {
    const sec = document.getElementById('let-security').value;
    const prec = document.getElementById('let-precedence').value;
    const copyVal = document.getElementById('let-copy').value;
    const pagesVal = document.getElementById('let-pages').value;
    const sender = document.getElementById('let-sender').value;
    const ref = document.getElementById('let-ref').value;
    const date = document.getElementById('let-date').value;
    const subject = document.getElementById('let-subject').value;
    const sigName = document.getElementById('let-sig-name').value;
    const sigRank = document.getElementById('let-sig-rank').value;
    const sigAppt = document.getElementById('let-sig-appt').value;
    const act = document.getElementById('let-distr-action').value.trim();
    const info = document.getElementById('let-distr-info').value.trim();
    const intAct = document.getElementById('let-distr-int-action').value.trim();
    const intInfo = document.getElementById('let-distr-int-info').value.trim();

    let body = '';

    // --- Precedence, Copy, Pages ---
    if (prec) body += `<p style="margin: 0 0 4pt 325pt; font-family:Arial,sans-serif; font-weight:bold; font-size:12pt; color:#cc0000;">${prec}</p>\n`;
    if (copyVal) body += `<p style="margin: 0 0 2pt 325pt; font-family:Arial,sans-serif; font-size:12pt;">Copy No ${copyVal}</p>\n`;
    if (pagesVal) body += `<p style="margin: 0 0 2pt 325pt; font-family:Arial,sans-serif; font-size:12pt;">Total Pages ${pagesVal}</p>\n`;

    // --- Sender Address ---
    const senderLines = sender.split('\n').map(l => l.trim()).filter(l => l);
    if (senderLines.length > 0) {
        body += `<p style="margin: 8pt 0 16pt 325pt; font-family:Arial,sans-serif; font-size:12pt; line-height:1.4;">${senderLines.join('<br>')}</p>\n`;
    }

    // --- Reference Number and Date (Ref left, Date aligned with sender via tab) ---
    body += `<p style="margin: 0 0 16pt 0; text-indent: 0; mso-tab-count: 1; tab-stops: 325pt; font-family:Arial,sans-serif; font-size:12pt;">`;
    body += `<span style="font-weight:500;">${ref}</span><span style="mso-tab-count:1">\t</span>${date}`;
    body += `</p>\n`;

    // --- Subject Heading ---
    if (subject.trim()) {
        body += `<p style="font-family:Arial,sans-serif; font-weight:bold; font-size:12pt; text-decoration:underline; text-transform:uppercase; margin: 0 0 12pt 0;">${subject.toUpperCase()}</p>\n`;
    }

    // --- References ---
    if (letterReferences.length > 0) {
        const refLabelText = letterReferences.length === 1 ? 'Ref:' : 'Refs:';
        body += `<p style="font-family:Arial,sans-serif; font-weight:bold; text-decoration:none; font-size:12pt; margin:0 0 4pt 0;">${refLabelText}</p>\n`;
        letterReferences.forEach((refText, idx) => {
            const letter = String.fromCharCode(65 + idx);
            body += `<p style="font-family:Arial,sans-serif; font-size:12pt; margin:0 0 4pt 0; text-indent: -24pt; margin-left: 24pt; tab-stops: 24pt;">${letter}.<span style="mso-tab-count:1">\t</span>${refText}</p>\n`;
        });
        body += `<p style="margin:0 0 12pt 0;">&nbsp;</p>\n`;
    }

    // --- Paragraphs (Body) ---
    letterParagraphs.forEach((para, idx) => {
        let formattedMain = para.text.replace(/^([^\.]{2,50}\.)\s/, '$1<span style="mso-tab-count:1">\t</span>');
        body += `<p style="margin: 0 0 12pt 0; text-align: justify; font-family:Arial,sans-serif; font-size:12pt;">`;
        body += `<span style="font-weight:500;">${idx + 1}.</span><span style="mso-tab-count:1">\t</span>${formattedMain}`;
        body += `</p>\n`;

        if (para.subParas.length > 0) {
            para.subParas.forEach((subText, subIdx) => {
                const subLetter = String.fromCharCode(97 + subIdx);
                let formattedSub = subText.replace(/^([^\.]{2,50}\.)\s/, '$1<span style="mso-tab-count:1">\t</span>');
                body += `<p style="margin: 0 0 12pt 28pt; text-align: justify; font-family:Arial,sans-serif; font-size:12pt;">`;
                body += `${subLetter}.<span style="mso-tab-count:1">\t</span>${formattedSub}`;
                body += `</p>\n`;
            });
        }
    });

    // --- Signature Block ---
    body += `<p style="margin: 40pt 0 2pt 325pt; font-family:Arial,sans-serif; font-size:12pt; font-weight:bold; text-transform:uppercase;">${sigName.toUpperCase()}</p>\n`;
    body += `<p style="margin: 0 0 2pt 325pt; font-family:Arial,sans-serif; font-size:12pt;">${sigRank}</p>\n`;
    body += `<p style="margin: 0 0 24pt 325pt; font-family:Arial,sans-serif; font-size:12pt;">${sigAppt}</p>\n`;

    // --- Distribution Block ---
    if (act || info || intAct || intInfo) {
        body += `<p style="margin: 12pt 0 8pt 0; font-family:Arial,sans-serif; font-size:12pt;">Distr:</p>\n`;
        
        let extLines = [];
        if (act || info) {
            extLines.push("Extl:");
            if (act) {
                extLines.push("Act:");
                extLines = extLines.concat(act.split('\n').map(l => l.trim()).filter(l => l));
            }
            if (info) {
                extLines.push("Info:");
                extLines = extLines.concat(info.split('\n').map(l => l.trim()).filter(l => l));
            }
        }
        
        let intLines = [];
        if (intAct || intInfo) {
            intLines.push("Intl:");
            if (intAct) {
                intLines.push("Act:");
                intLines = intLines.concat(intAct.split('\n').map(l => l.trim()).filter(l => l));
            }
            if (intInfo) {
                intLines.push("Info:");
                intLines = intLines.concat(intInfo.split('\n').map(l => l.trim()).filter(l => l));
            }
        }
        
        const maxLines = Math.max(extLines.length, intLines.length);
        for (let i = 0; i < maxLines; i++) {
            const eLine = extLines[i] || "";
            const iLine = intLines[i] || "";
            if (iLine) {
                body += `<p style="margin: 0 0 1pt 0; mso-tab-count: 1; tab-stops: 250pt; font-family:Arial,sans-serif; font-size:12pt;">`;
                body += `${eLine}<span style="mso-tab-count:1">\t</span>${iLine}</p>\n`;
            } else {
                body += `<p style="margin: 0 0 1pt 0; font-family:Arial,sans-serif; font-size:12pt;">${eLine}</p>\n`;
            }
        }
    }

    // --- Build the full Word-compatible HTML document ---
    const fullHTML = `<html xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns:m="http://schemas-microsoft.com/office/2004/12/omml"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>MIL-Correspondence-Format - JSSDM Letter</title>
<style>
    @page Section1 {
        size: 21cm 29.7cm;
        margin: 25mm 25mm 25mm 25mm;
        mso-header-margin: 15mm;
        mso-footer-margin: 15mm;
        mso-header: h1;
        mso-footer: f1;
        mso-paper-source: 0;
    }
    div.Section1 {
        page: Section1;
    }
    body {
        font-family: Arial, sans-serif;
        font-size: 12pt;
        line-height: 1.5;
        color: #000000;
    }
    p {
        font-family: Arial, sans-serif;
    }
</style>
</head>
<body>

<div class="Section1">
    ${body}

    <!-- MS Word Header & Footer Definitions (inside Section1 to prevent duplication in body) -->
    <div style="mso-element:header" id="h1">
        <p style="text-align:center; font-family:Arial,sans-serif; font-size:12pt; letter-spacing:1px; margin:0; color:#000000;">${sec}</p>
    </div>
    <div style="mso-element:footer" id="f1">
        <p style="text-align:center; font-family:Arial,sans-serif; font-size:12pt; letter-spacing:1px; margin:0; color:#000000;">${sec}</p>
    </div>
</div>

</body>
</html>`;

    return fullHTML;
}

/* ==========================================================================
   Utility Functions
   ========================================================================== */
function downloadBlob(content, filename, contentType) {
    const blob = new Blob([content], { type: contentType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

/* ==========================================================================
   Service Paper Builder (Tab 3)
   ========================================================================== */
let spParagraphs = [];

function initServicePaperBuilder() {
    // Buttons
    document.getElementById('btn-sp-analyze').addEventListener('click', () => {
        alert('Please upload a document to analyze. The button is locked until a document is parsed.');
    });

    const spDropZone = document.getElementById('sp-file-drop-zone');
    const spFileInput = document.getElementById('sp-file-input');

    spDropZone.addEventListener('click', () => spFileInput.click());

    spDropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        spDropZone.classList.add('dragover');
    });

    spDropZone.addEventListener('dragleave', () => {
        spDropZone.classList.remove('dragover');
    });

    spDropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        spDropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleSpFileUpload(e.dataTransfer.files[0]);
        }
    });

    spFileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleSpFileUpload(e.target.files[0]);
        }
    });

    // Editor bindings
    const inputs = ['sp-ref', 'sp-copy', 'sp-pages', 'sp-title', 'sp-sig-left', 'sp-sig-right', 'sp-distribution'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('input', updateSpPreview);
    });

    document.getElementById('btn-download-word-sp').addEventListener('click', downloadSpWordDoc);
    document.getElementById('btn-copy-sp-text').addEventListener('click', copySpText);

    // Initial render
    renderSpParagraphsEditor();
    updateSpPreview();
}

function handleSpFileUpload(file) {
    const reader = new FileReader();
    const btnAnalyze = document.getElementById('btn-sp-analyze');
    
    // Change button to show processing
    btnAnalyze.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processing...';
    btnAnalyze.disabled = true;

    if (file.name.endsWith('.docx')) {
        if (typeof mammoth === 'undefined') {
            alert("CRITICAL ERROR: mammoth.js did not load from the CDN. Please check your internet connection.");
            btnAnalyze.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> Correct Paper via AI Agent';
            btnAnalyze.disabled = false;
            return;
        }
        reader.onload = function(e) {
            const arrayBuffer = e.target.result;
            mammoth.extractRawText({arrayBuffer: arrayBuffer})
                .then(function(result) {
                    try {
                        const text = result.value;
                        analyzeServicePaperAI(text);
                        btnAnalyze.innerHTML = '<i class="fa-solid fa-check"></i> Paper Corrected via AI';
                        btnAnalyze.disabled = false;
                    } catch (parserError) {
                        console.error("AI Parser Error:", parserError);
                        alert("AI Parser crashed: " + parserError.message);
                        btnAnalyze.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> Correct Paper via AI Agent';
                        btnAnalyze.disabled = false;
                    }
                })
                .catch(function(err) {
                    console.error("Mammoth Error:", err);
                    alert("Error parsing DOCX file: " + (err.message || err) + "\n\nNote: Make sure the file is a real Microsoft Word .docx file (not a renamed .doc or PDF).");
                    btnAnalyze.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> Correct Paper via AI Agent';
                    btnAnalyze.disabled = false;
                });
        };
        reader.readAsArrayBuffer(file);
    } else if (file.name.endsWith('.txt')) {
        reader.onload = function(e) {
            const text = e.target.result;
            analyzeServicePaperAI(text);
            btnAnalyze.innerHTML = '<i class="fa-solid fa-check"></i> Paper Corrected via AI';
            btnAnalyze.disabled = false;
        };
        reader.readAsText(file);
    } else if (file.name.endsWith('.doc')) {
        // Read as text (since our .doc exports are actually HTML)
        reader.onload = function(e) {
            const htmlText = e.target.result;
            // Create a temporary element to strip HTML tags
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlText;
            const plainText = tempDiv.innerText || tempDiv.textContent;
            
            analyzeServicePaperAI(plainText);
            btnAnalyze.innerHTML = '<i class="fa-solid fa-check"></i> Paper Corrected via AI';
            btnAnalyze.disabled = false;
        };
        reader.readAsText(file);
    } else {
        alert("Only .txt, .doc, and .docx files are supported.");
        btnAnalyze.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> Correct Paper via AI Agent';
        btnAnalyze.disabled = false;
    }
}

function analyzeServicePaperAI(text) {
    // Rule-based parsing engine with post-parse signature block cleanup
    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    
    spParagraphs = [];
    let currentPara = null;

    const titleEl = document.getElementById('sp-title');
    const distEl = document.getElementById('sp-distribution');
    const sigLeftEl = document.getElementById('sp-sig-left');
    const sigRightEl = document.getElementById('sp-sig-right');

    // Known main heading keywords for service papers
    const MAIN_HEADING_KEYWORDS = [
        'INTRODUCTION', 'AIM', 'RECOMMENDATIONS', 'RECOMMENDATION',
        'CONCLUSION', 'CONCLUSIONS', 'SUMMARY'
    ];

    // Signature / end-of-body detection patterns
    const SIGNATURE_PATTERNS = [
        /^(squadron leader|wing commander|group captain|flight lieutenant|air commodore|air vice marshal|air marshal)/i,
        /^(colonel|brigadier|major general|lieutenant general|general|lt col|lieutenant colonel|major|captain)/i,
        /^(commander|commodore|rear admiral|vice admiral|admiral|lt cdr|lieutenant commander)/i,
        /^index no/i,
        /^group (member|leader)/i,
        /^course participant/i,
        /^directing staff/i,
        /^(legends?:|TS\s*=|EI\s*=|Ex\s*=|TH\s*=)/i,
        /^total word count/i,
        /^word count/i
    ];

    function isSignatureLine(line) {
        return SIGNATURE_PATTERNS.some(p => p.test(line));
    }

    function isAllCapsName(line) {
        // ALL CAPS, 2-5 words, no numbers, not a known heading
        return line === line.toUpperCase() && 
               line.length > 5 && line.length < 80 &&
               /^[A-Z\s]+$/.test(line) &&
               line.split(/\s+/).length >= 2 && line.split(/\s+/).length <= 6 &&
               !MAIN_HEADING_KEYWORDS.includes(line.trim()) &&
               !line.includes('SECURITY') && !line.includes('RESTRICTED') &&
               !line.includes('CONFIDENTIAL') && !line.includes('SECRET');
    }

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];

        // Skip security classification lines
        if (/^(RESTRICTED|CONFIDENTIAL|SECRET|TOP SECRET|UNCLASSIFIED)$/i.test(line)) {
            continue;
        }

        // Skip page numbers (e.g. "1", "2 of 5", "1-3")
        if (/^\d+(\s+of\s+\d+)?$/.test(line) || /^\d+-\d+$/.test(line)) {
            continue;
        }

        // Identifying reference detection (e.g. PF/419/DSCSC/2026)
        if (/^[A-Z]{2,}\/\d+\/[A-Z]+/i.test(line) && spParagraphs.length === 0) {
            const refEl = document.getElementById('sp-ref');
            if (refEl && !refEl.value) refEl.value = line;
            continue;
        }

        // Title heuristics (All Caps, longer line, before any paragraphs)
        if (line === line.toUpperCase() && line.length < 200 && !line.match(/^[0-9]+\./) && spParagraphs.length === 0 && (!titleEl || !titleEl.value)) {
            if (MAIN_HEADING_KEYWORDS.includes(line.trim())) {
                spParagraphs.push({ type: 'main_heading', text: line, heading: '', subParas: [] });
            } else if (titleEl && line.length > 10) {
                titleEl.value = line;
            }
            continue;
        }

        // Distribution detection - extract and stop
        if (/^distribution:?$/i.test(line) || line.toLowerCase().startsWith('distribution:')) {
            let dist = "";
            let j = i + 1;
            while (j < lines.length) {
                let dLine = lines[j];
                if (/^(total )?word count/i.test(dLine) || /^legends?:/i.test(dLine) || /^TS\s*=/i.test(dLine)) break;
                if (/^(RESTRICTED|CONFIDENTIAL|SECRET|TOP SECRET|UNCLASSIFIED)$/i.test(dLine)) break;
                if (/^\d+(\s+of\s+\d+)?$/.test(dLine)) break;
                dist += dLine + '\n';
                j++;
            }
            if (distEl) distEl.value = dist.trim();
            break; // Stop parsing after distribution
        }

        // Stop on bibliography
        if (/^bibliography:?$/i.test(line) || line === 'BIBLIOGRAPHY') {
            break;
        }

        // Stop on word count / legends
        if (/^(total )?word count/i.test(line) || /^legends?:/i.test(line)) {
            break;
        }

        // Stop on signature-pattern lines (rank, index no, group member etc.)
        if (isSignatureLine(line)) {
            break;
        }

        // Stop if we hit an ALL CAPS name that looks like a signatory
        // (only if we've already parsed some content and it's near the end)
        if (spParagraphs.length > 3 && isAllCapsName(line)) {
            // Look ahead: if the next line is a rank or "Index No", this is a signature block
            let nextLine = (i + 1 < lines.length) ? lines[i + 1] : '';
            if (isSignatureLine(nextLine) || isAllCapsName(nextLine)) {
                break;
            }
        }

        // Main heading heuristic (All Caps, known keywords or short caps line)
        if (line === line.toUpperCase() && line.length < 80 && !line.match(/^[0-9]+\./) && /[A-Z]/.test(line)) {
            // Don't treat very short lines or lines near the end as headings
            if (MAIN_HEADING_KEYWORDS.some(kw => line.includes(kw)) || (line.length > 3 && line.length < 60)) {
                spParagraphs.push({ type: 'main_heading', text: line, heading: '', subParas: [] });
                continue;
            }
        }

        // Paragraph parsing (numbered: "1.", "2.", etc.)
        let paraMatch = line.match(/^(\d+)\.\s*(.*)/);
        if (paraMatch) {
            currentPara = { type: 'paragraph', heading: '', text: paraMatch[2], subParas: [] };
            spParagraphs.push(currentPara);
            continue;
        }

        // Sub-paragraph parsing ("a.", "b.", etc.)
        let subMatch = line.match(/^([a-z])\.\s*(.*)/);
        if (subMatch && currentPara) {
            currentPara.subParas.push(subMatch[2]);
            continue;
        }

        // Group heading heuristic (Title Case, short line, no number)
        if (line.length < 60 && !line.match(/^[0-9]/) && line.split(' ').every(w => w.length === 0 || w[0] === w[0].toUpperCase() || w.length < 4)) {
            // Ensure it's not a location/date line
            if (!/^(january|february|march|april|may|june|july|august|september|october|november|december|dhaka|cox|chittagong|comilla|savar|mirpur|rangpur|sylhet|rajshahi|bogra)/i.test(line)) {
                spParagraphs.push({ type: 'group_heading', text: line, heading: '', subParas: [] });
                continue;
            }
        }

        // Fallback: append to current paragraph if exists
        if (currentPara) {
            if (currentPara.subParas.length > 0) {
                currentPara.subParas[currentPara.subParas.length - 1] += ' ' + line;
            } else {
                currentPara.text += ' ' + line;
            }
        }
    }

    // === POST-PARSE CLEANUP ===
    // Find the last CONCLUSION or RECOMMENDATIONS heading, then keep only
    // paragraphs up to (and including) the last paragraph after that heading.
    let lastConcRecIdx = -1;
    for (let i = spParagraphs.length - 1; i >= 0; i--) {
        if (spParagraphs[i].type === 'main_heading') {
            const headingText = spParagraphs[i].text.toUpperCase().trim();
            if (headingText.includes('CONCLUSION') || headingText.includes('RECOMMENDATION')) {
                lastConcRecIdx = i;
                break;
            }
        }
    }

    if (lastConcRecIdx >= 0) {
        // Find the last paragraph (numbered) after that heading
        let lastParaAfterConc = -1;
        for (let i = lastConcRecIdx + 1; i < spParagraphs.length; i++) {
            if (spParagraphs[i].type === 'paragraph') {
                lastParaAfterConc = i;
            } else if (spParagraphs[i].type === 'main_heading') {
                // Another main heading after conclusion (e.g., RECOMMENDATIONS after CONCLUSION)
                // Check if it's also a conclusion/recommendation type
                const ht = spParagraphs[i].text.toUpperCase().trim();
                if (ht.includes('CONCLUSION') || ht.includes('RECOMMENDATION')) {
                    lastConcRecIdx = i; // Update to this later heading
                    lastParaAfterConc = -1; // Reset
                } else {
                    break; // Some other heading (BIBLIOGRAPHY etc.) — stop here
                }
            }
        }

        // Truncate: keep everything up to and including the last paragraph after conclusion/recommendations
        if (lastParaAfterConc >= 0) {
            spParagraphs = spParagraphs.slice(0, lastParaAfterConc + 1);
        } else {
            // No paragraphs found after the heading — just keep up to the heading
            spParagraphs = spParagraphs.slice(0, lastConcRecIdx + 1);
        }
    }
    
    renderSpParagraphsEditor();
    updateSpPreview();
    alert("AI parsing complete. Please review the extracted structure in the manual adjustments forms below.");
}

function renderSpParagraphsEditor() {
    const container = document.getElementById('sp-paras-container');
    container.innerHTML = '';
    
    spParagraphs.forEach((para, pIdx) => {
        const pCard = document.createElement('div');
        pCard.className = 'editor-card';
        pCard.style.padding = '10px';
        pCard.style.marginBottom = '10px';
        pCard.style.backgroundColor = 'var(--panel-bg)';
        
        const header = document.createElement('div');
        header.style.display = 'flex';
        header.style.justifyContent = 'space-between';
        header.style.marginBottom = '8px';
        
        const typeSelect = document.createElement('select');
        typeSelect.innerHTML = `
            <option value="paragraph" ${para.type==='paragraph'?'selected':''}>Numbered Paragraph</option>
            <option value="main_heading" ${para.type==='main_heading'?'selected':''}>Main Heading (Centered, Caps)</option>
            <option value="group_heading" ${para.type==='group_heading'?'selected':''}>Group Heading (Left, Title Case)</option>
            <option value="figure" ${para.type==='figure'?'selected':''}>Figure / Image Placeholder</option>
        `;
        typeSelect.addEventListener('change', (e) => {
            para.type = e.target.value;
            renderSpParagraphsEditor();
            updateSpPreview();
        });
        
        const btnDelete = document.createElement('button');
        btnDelete.className = 'btn btn-danger btn-sm';
        btnDelete.innerHTML = '<i class="fa-solid fa-xmark"></i>';
        btnDelete.onclick = () => {
            spParagraphs.splice(pIdx, 1);
            renderSpParagraphsEditor();
            updateSpPreview();
        };
        
        header.appendChild(typeSelect);
        header.appendChild(btnDelete);
        pCard.appendChild(header);
        
        const textInput = document.createElement('textarea');
        textInput.className = 'form-control';
        textInput.style.width = '100%';
        textInput.style.marginBottom = '10px';
        textInput.rows = para.type === 'figure' ? 1 : 3;
        textInput.placeholder = para.type === 'figure' ? 'Figure Name (e.g. Figure-1: Flowers of Bangladesh)' : 'Paragraph Text';
        textInput.value = para.text;
        textInput.addEventListener('input', (e) => {
            para.text = e.target.value;
            updateSpPreview();
        });
        
        if (para.type === 'paragraph' || para.type === 'figure') {
            const headingInput = document.createElement('input');
            headingInput.type = 'text';
            headingInput.className = 'form-control';
            headingInput.style.width = '100%';
            headingInput.style.marginBottom = '10px';
            headingInput.placeholder = para.type === 'figure' ? 'Source (Optional, e.g. Source: Bangladesh Agricultural Board 2017)' : 'Paragraph Heading (Optional, e.g. General.)';
            headingInput.value = para.heading || '';
            headingInput.addEventListener('input', (e) => {
                para.heading = e.target.value;
                updateSpPreview();
            });
            pCard.appendChild(headingInput);
        }
        
        pCard.appendChild(textInput);
        
        if (para.type === 'paragraph') {
            const subContainer = document.createElement('div');
            subContainer.style.paddingLeft = '20px';
            
            para.subParas.forEach((subText, sIdx) => {
                const sRow = document.createElement('div');
                sRow.style.display = 'flex';
                sRow.style.gap = '8px';
                sRow.style.marginBottom = '5px';
                
                const sInput = document.createElement('textarea');
                sInput.className = 'form-control';
                sInput.style.flexGrow = '1';
                sInput.rows = 2;
                sInput.value = subText;
                sInput.addEventListener('input', (e) => {
                    para.subParas[sIdx] = e.target.value;
                    updateSpPreview();
                });
                
                const sDelBtn = document.createElement('button');
                sDelBtn.className = 'btn btn-danger btn-sm';
                sDelBtn.innerHTML = '<i class="fa-solid fa-minus"></i>';
                sDelBtn.onclick = () => {
                    para.subParas.splice(sIdx, 1);
                    renderSpParagraphsEditor();
                    updateSpPreview();
                };
                
                sRow.appendChild(sInput);
                sRow.appendChild(sDelBtn);
                subContainer.appendChild(sRow);
            });
            
            const btnAddSub = document.createElement('button');
            btnAddSub.className = 'btn btn-secondary btn-sm';
            btnAddSub.innerHTML = '<i class="fa-solid fa-plus"></i> Add Sub-para';
            btnAddSub.onclick = () => {
                para.subParas.push('');
                renderSpParagraphsEditor();
                updateSpPreview();
            };
            
            subContainer.appendChild(btnAddSub);
            pCard.appendChild(subContainer);
        }
        
        container.appendChild(pCard);
    });
}

function addSpParagraph() {
    spParagraphs.push({
        type: 'paragraph',
        text: '',
        subParas: []
    });
    renderSpParagraphsEditor();
    updateSpPreview();
}

function updateSpPreview() {
    const getVal = (id) => { const el = document.getElementById(id); return el ? el.value : ''; };
    const setTxt = (id, txt) => { const el = document.getElementById(id); if (el) el.innerText = txt; };
    const setHTML = (id, html) => { const el = document.getElementById(id); if (el) el.innerHTML = html; };

    setTxt('sp-p-ref', getVal('sp-ref'));
    setTxt('sp-p-copy', getVal('sp-copy') ? `Copy No ${getVal('sp-copy')}` : '');
    setTxt('sp-p-pages', getVal('sp-pages') ? `Total Pages ${getVal('sp-pages')}` : '');
    
    setTxt('sp-p-title', getVal('sp-title').toUpperCase());
    
    // Signature left block (Location, Date, Annexes)
    const sigLeftText = getVal('sp-sig-left');
    const sigLeftLines = sigLeftText.split('\n').filter(l => l.trim());
    setTxt('sp-p-loc', sigLeftLines[0] || '');
    setTxt('sp-p-date', sigLeftLines.slice(1).join('\n') || '');
    
    // Signature right block (Name, Rank, Appointment)
    const sigRightText = getVal('sp-sig-right');
    const sigRightLines = sigRightText.split('\n').filter(l => l.trim());
    setTxt('sp-p-sig-name', (sigRightLines[0] || '').toUpperCase());
    setTxt('sp-p-sig-rank', sigRightLines[1] || '');
    setTxt('sp-p-sig-appt', sigRightLines.slice(2).join('\n') || '');
    
    const distText = getVal('sp-distribution');
    if (distText.trim()) {
        setHTML('sp-p-distribution', '<strong>Distribution:</strong><br><br>' + distText.replace(/\n/g, '<br>'));
    } else {
        setHTML('sp-p-distribution', '');
    }
    
    // Body paragraphs
    const bodyContainer = document.getElementById('sp-p-body');
    bodyContainer.innerHTML = '';
    
    let globalParaCounter = 1;
    
    spParagraphs.forEach((para) => {
        if (para.type === 'main_heading') {
            const h = document.createElement('div');
            h.className = 'sp-main-heading';
            h.innerText = para.text;
            bodyContainer.appendChild(h);
        } else if (para.type === 'group_heading') {
            const h = document.createElement('div');
            h.className = 'sp-group-heading';
            h.innerText = para.text;
            bodyContainer.appendChild(h);
        } else if (para.type === 'figure') {
            const f = document.createElement('div');
            f.style.textAlign = 'center';
            f.style.marginBottom = '16pt';
            f.innerHTML = `<p style="margin:0 0 8pt 0; font-family:Arial,sans-serif; font-size:12pt;">${para.text}</p>
                           <div style="border: 2px dashed #999; padding: 40px; margin: 0 auto 8pt auto; width: 60%; color: #999;">[ IMAGE PLACEHOLDER ]</div>
                           <p style="margin:0; text-align:left; padding-left: 20%; font-family:Arial,sans-serif; font-size:12pt;">${para.heading || ''}</p>`;
            bodyContainer.appendChild(f);
        } else if (para.type === 'paragraph') {
            const p = document.createElement('div');
            p.className = 'preview-para';
            p.style.marginBottom = '12pt';
            
            let headingHTML = para.heading ? `<u>${para.heading}</u>&nbsp;&nbsp;&nbsp;` : '';
            
            p.innerHTML = `<span style="font-weight:500;">${globalParaCounter}.</span>&nbsp;&nbsp;&nbsp;${headingHTML}${para.text}`;
            bodyContainer.appendChild(p);
            
            globalParaCounter++;
            
            if (para.subParas.length > 0) {
                para.subParas.forEach((sub, sIdx) => {
                    const sp = document.createElement('div');
                    sp.className = 'preview-sub-para';
                    sp.style.marginBottom = '12pt';
                    sp.style.marginLeft = '28pt'; // Tab indent
                    const subLetter = String.fromCharCode(97 + sIdx);
                    
                    sp.innerHTML = `${subLetter}.&nbsp;&nbsp;&nbsp;${sub}`;
                    bodyContainer.appendChild(sp);
                });
            }
        }
    });
}

function buildServicePaperWordDocHTML() {
    const ref = document.getElementById('sp-ref').value;
    const copy = document.getElementById('sp-copy').value;
    const pages = document.getElementById('sp-pages').value;
    const title = document.getElementById('sp-title').value;
    const sigLeft = document.getElementById('sp-sig-left').value;
    const sigRight = document.getElementById('sp-sig-right').value;
    const distr = document.getElementById('sp-distribution').value;

    let body = `
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
    <meta charset="utf-8">
    <title>Export Service Paper</title>
    <style>
        @page WordSection1 {
            size: 595.3pt 841.9pt;
            margin: 70.85pt 70.85pt 70.85pt 70.85pt;
            mso-header-margin: 35.4pt;
            mso-footer-margin: 35.4pt;
            mso-paper-source: 0;
        }
        div.WordSection1 { page: WordSection1; }
        p { margin: 0; font-family: Arial, sans-serif; font-size: 12pt; line-height: 1.5; }
    </style>
</head>
<body>
<div class="WordSection1">
`;

    // Security Top removed

    // Header table row
    if (pages) {
        body += `<p style="margin: 0 0 12pt 0; text-indent: 0; mso-tab-count: 1; tab-stops: 397pt; font-family:Arial,sans-serif; font-size:12pt;">`;
        body += `${ref}<span style="mso-tab-count:1">\t</span>${copy ? 'Copy No ' + copy : ''}`;
        body += `</p>\n`;
        body += `<p style="margin: 0 0 24pt 0; text-indent: 0; mso-tab-count: 1; tab-stops: 397pt; font-family:Arial,sans-serif; font-size:12pt;">`;
        body += `<span style="mso-tab-count:1">\t</span>Total Pages ${pages}`;
        body += `</p>\n`;
    } else {
        body += `<p style="margin: 0 0 24pt 0; text-indent: 0; mso-tab-count: 1; tab-stops: 397pt; font-family:Arial,sans-serif; font-size:12pt;">`;
        body += `${ref}<span style="mso-tab-count:1">\t</span>${copy ? 'Copy No ' + copy : ''}`;
        body += `</p>\n`;
    }

    // Main Title
    if (title.trim()) {
        body += `<p style="text-align:center; font-family:Arial,sans-serif; font-weight:bold; font-size:12pt; text-decoration:underline; text-transform:uppercase; margin: 0 0 24pt 0;">${title.toUpperCase()}</p>\n`;
    }

    // Paragraphs
    let globalParaCounter = 1;
    spParagraphs.forEach((para) => {
        if (para.type === 'main_heading') {
            body += `<p style="text-align:center; font-family:Arial,sans-serif; font-weight:bold; font-size:12pt; text-decoration:underline; text-transform:uppercase; margin: 0 0 12pt 0;">${para.text.toUpperCase()}</p>\n`;
        } else if (para.type === 'group_heading') {
            body += `<p style="margin: 0 0 12pt 0; font-family:Arial,sans-serif; font-size:12pt; text-decoration:underline; font-weight:bold;">${para.text}</p>\n`;
        } else if (para.type === 'figure') {
            body += `<p style="margin: 12pt 0 12pt 0; text-align:center; font-family:Arial,sans-serif; font-size:12pt;">${para.text}</p>\n`;
            body += `<table width="80%" align="center" style="margin: 0 auto; border: 1px solid #000;"><tr><td style="text-align:center; padding:50px; color:#666; font-family:Arial,sans-serif;">[ Insert Image Here ]</td></tr></table>\n`;
            if (para.heading) {
                body += `<p style="margin: 4pt 0 12pt 0; text-align:left; font-family:Arial,sans-serif; font-size:12pt;">${para.heading}</p>\n`;
            }
        } else if (para.type === 'paragraph') {
            let headingHTML = para.heading ? `<u><b>${para.heading}</b></u><span style="mso-tab-count:1">\t</span>` : '';
            body += `<p style="margin: 0 0 12pt 0; text-align: justify; font-family:Arial,sans-serif; font-size:12pt; mso-tab-count:1; tab-stops: 24pt;">`;
            body += `${globalParaCounter}.<span style="mso-tab-count:1">\t</span>${headingHTML}${para.text}`;
            body += `</p>\n`;
            
            globalParaCounter++;

            if (para.subParas.length > 0) {
                para.subParas.forEach((sub, sIdx) => {
                    const subLetter = String.fromCharCode(97 + sIdx);
                    body += `<p style="margin: 0 0 12pt 28pt; text-align: justify; font-family:Arial,sans-serif; font-size:12pt;">`;
                    body += `${subLetter}.&nbsp;&nbsp;&nbsp;${sub}`;
                    body += `</p>\n`;
                });
            }
        }
    });

    // Signature Row
    body += `<p style="margin: 24pt 0 0 0;">&nbsp;</p>`;
    
    const leftLines = sigLeft.split('\n').map(l => l.trim());
    const rightLines = sigRight.split('\n').map(l => l.trim());
    const maxSigLines = Math.max(leftLines.length, rightLines.length);

    for (let i = 0; i < maxSigLines; i++) {
        const lLine = leftLines[i] || "";
        const rLine = rightLines[i] || "";
        body += `<p style="margin: 0 0 0 0; mso-tab-count: 1; tab-stops: 397pt; font-family:Arial,sans-serif; font-size:12pt;">`;
        if (lLine || rLine) {
            body += `${lLine}<span style="mso-tab-count:1">\t</span>${rLine}`;
        } else {
            body += `&nbsp;`;
        }
        body += `</p>\n`;
    }

    body += `<p style="margin: 0 0 24pt 0;">&nbsp;</p>\n`;

    // Distribution
    if (distr.trim()) {
        body += `<p style="margin: 0 0 12pt 0; font-family:Arial,sans-serif; font-size:12pt; text-decoration:underline;">Distribution:</p>\n`;
        const dLines = distr.split('\n').filter(l => l.trim());
        dLines.forEach(l => {
            body += `<p style="margin: 0 0 0 0; font-family:Arial,sans-serif; font-size:12pt;">${l}</p>\n`;
        });
        body += `<p style="margin: 0 0 24pt 0;">&nbsp;</p>\n`;
    }

    // Security Bottom removed

    body += `</div></body></html>`;
    return body;
}

function downloadSpWordDoc() {
    const htmlContent = buildServicePaperWordDocHTML();
    downloadBlob(htmlContent, 'service_paper.doc', 'application/msword');
}

function copySpText() {
    // Generate simple text output
    let text = "";
    text += document.getElementById('sp-ref').value + "\t\t" + (document.getElementById('sp-copy').value ? "Copy No " + document.getElementById('sp-copy').value : "") + "\n";
    text += "\t\t" + (document.getElementById('sp-pages').value ? "Total Pages " + document.getElementById('sp-pages').value : "") + "\n\n";
    text += document.getElementById('sp-title').value.toUpperCase() + "\n\n";
    
    let c = 1;
    spParagraphs.forEach(p => {
        if (p.type === 'main_heading') text += p.text + "\n\n";
        else if (p.type === 'group_heading') text += p.text + "\n\n";
        else if (p.type === 'figure') {
            text += "    " + p.text + "\n";
            text += "    [ IMAGE ]\n";
            if (p.heading) text += "    " + p.heading + "\n";
            text += "\n";
        }
        else {
            let headingText = p.heading ? p.heading + "\t" : "";
            text += c + ".\t" + headingText + p.text + "\n\n";
            c++;
            p.subParas.forEach((sub, sIdx) => {
                text += "\t" + String.fromCharCode(97+sIdx) + ".\t" + sub + "\n\n";
            });
        }
    });

    const lLines = document.getElementById('sp-sig-left').value.split('\n');
    const rLines = document.getElementById('sp-sig-right').value.split('\n');
    const maxL = Math.max(lLines.length, rLines.length);
    for (let i = 0; i < maxL; i++) {
        text += (lLines[i] || "") + "\t\t" + (rLines[i] || "") + "\n";
    }
    text += "\n";

    if (document.getElementById('sp-distribution').value) {
        text += "Distribution:\n" + document.getElementById('sp-distribution').value + "\n\n";
    }
    
    navigator.clipboard.writeText(text).then(() => {
        alert("Text copied to clipboard!");
    });
}
window.loadSpTemplateSample = function() {
    document.getElementById('sp-ref').value = 'PF/419/DSCSC/2026';
    document.getElementById('sp-copy').value = '';
    document.getElementById('sp-pages').value = '6';
    document.getElementById('sp-title').value = 'ARTIFICIAL INTELLIGENCE AND DECISION SUPERIORITY: IMPLICATIONS FOR THE MILITARY DECISION-MAKING PROCESS';
    
    spParagraphs = [
        { type: 'main_heading', text: 'INTRODUCTION', heading: '', subParas: [] },
        { type: 'paragraph', heading: 'General.', text: 'Artificial intelligence is rapidly reshaping the character of modern warfare and military planning. Modern battlefields generate vast quantities of information from sensors, surveillance systems, cyber networks and unmanned platforms. This expanding information environment increasingly influences how commanders analyze situations and make operational decisions. Understanding how artificial intelligence affects the Military Decision-Making Process therefore becomes essential for modern armed forces and requires careful analytical examination in contemporary operations and strategic planning environments.', subParas: [] },
        { type: 'paragraph', heading: '', text: 'Artificial intelligence is becoming important because modern military commanders must process information faster than adversaries. Traditional analytical methods alone struggle to manage the speed, scale and complexity of contemporary multi-domain operations. AI-supported analytical tools offer the potential to enhance situational awareness, accelerate planning cycles and support more informed operational choices.', subParas: [] },
        { type: 'paragraph', heading: 'Scope.', text: 'This paper therefore explains how artificial intelligence influences the Military Decision-Making Process. Firstly, it outlines the concept of artificial intelligence and evaluates its advantages and disadvantages in military decision-making. Secondly, it analyzes the key technological, operational and ethical challenges created by AI integration.', subParas: [] },
        { type: 'main_heading', text: 'AIM', heading: '', subParas: [] },
        { type: 'paragraph', heading: '', text: 'The aim of this paper is to examine the advantages and disadvantages of integrating Artificial Intelligence into the Military Decision-Making Process, analyze the key challenges of AI adoption, and make recommendations for the responsible and effective use of AI in military decision-making.', subParas: [] },
        { type: 'main_heading', text: 'CONCEPT OF ARTIFICIAL INTELLIGENCE', heading: '', subParas: [] },
        { type: 'group_heading', text: 'Definition and Core Components', heading: '', subParas: [] },
        { type: 'paragraph', heading: '', text: 'Artificial Intelligence refers to computer systems designed to perform tasks that typically require human intelligence, including visual perception, speech recognition, decision-making and language translation. In the military context, AI encompasses machine learning algorithms, neural networks, natural language processing and computer vision systems that can process, analyze and act upon large datasets faster than human operators.', subParas: [] },
        { type: 'paragraph', heading: '', text: 'The core components of military AI systems include the following:', subParas: [
            'Machine Learning. Algorithms that improve performance through experience and data exposure without explicit programming.',
            'Natural Language Processing. Systems that interpret and generate human language for intelligence analysis and communication.',
            'Computer Vision. Technology that enables machines to interpret visual information from sensors, satellites and surveillance platforms.'
        ]},
        { type: 'main_heading', text: 'ADVANTAGES AND DISADVANTAGES OF AI IN MILITARY DECISION-MAKING', heading: '', subParas: [] },
        { type: 'group_heading', text: 'Advantages', heading: '', subParas: [] },
        { type: 'paragraph', heading: 'Speed of Processing.', text: 'AI systems can analyze vast quantities of intelligence data in real time, significantly reducing the time required for information processing during the decision-making cycle. This enhanced speed allows commanders to make faster and more informed decisions during time-critical operations.', subParas: [] },
        { type: 'paragraph', heading: 'Pattern Recognition.', text: 'Machine learning algorithms excel at identifying patterns, anomalies and trends within large datasets that human analysts might overlook. This capability proves particularly valuable in signals intelligence, imagery analysis and predictive threat assessment.', subParas: [] },
        { type: 'group_heading', text: 'Disadvantages', heading: '', subParas: [] },
        { type: 'paragraph', heading: 'Algorithmic Opacity.', text: 'Complex AI models often function as black boxes, making it difficult for commanders to understand exactly how the system reached its recommendations. This lack of transparency creates challenges for accountability and trust in military operations where the rationale for decisions must be clearly defensible.', subParas: [] },
        { type: 'paragraph', heading: 'Cybersecurity Vulnerabilities.', text: 'AI systems remain susceptible to adversarial attacks including data poisoning, model manipulation and cyber intrusion. These vulnerabilities create significant risks when AI-generated insights directly influence operational planning and tactical decisions.', subParas: [] },
        { type: 'main_heading', text: 'CONCLUSION', heading: '', subParas: [] },
        { type: 'paragraph', heading: '', text: 'This paper examined the transformation of the Military Decision-Making Process through the integration of Artificial Intelligence in modern military operations. It highlighted how AI enhances decision-making by improving data analysis, situational awareness, and the development of operational courses of action. At the same time, the analysis demonstrated that AI introduces notable disadvantages, including the risk of overreliance on automated systems, limited transparency of algorithms, and vulnerability to unreliable data inputs. These findings illustrate that while AI offers powerful analytical advantages, its integration must be carefully balanced with human judgment and professional military expertise.', subParas: [] },
        { type: 'paragraph', heading: '', text: 'The study also identified several challenges associated with incorporating Artificial Intelligence into the Military Decision-Making Process. These challenges include algorithmic opacity, cybersecurity vulnerabilities, ethical concerns regarding accountability, and institutional difficulties in adapting traditional military planning systems to emerging technologies. Such factors demonstrate that technological innovation alone cannot guarantee improved decision-making unless supported by appropriate doctrine, organizational adaptation, and governance frameworks.', subParas: [] },
        { type: 'main_heading', text: 'RECOMMENDATIONS', heading: '', subParas: [] },
        { type: 'paragraph', heading: '', text: 'Upon analyzing the influence of Artificial Intelligence on the Military Decision-Making Process and examining the associated challenges and mitigation measures, this paper recommends the following:', subParas: [
            'Military organizations should develop doctrine and operational frameworks to guide the responsible integration of Artificial Intelligence within the Military Decision-Making Process.',
            'Armed forces should strengthen cybersecurity infrastructure and data protection mechanisms to safeguard AI-enabled decision-support systems from adversarial interference.',
            'Professional military education institutions should incorporate Artificial Intelligence literacy and human-machine collaboration training to enable commanders and staff officers to effectively employ AI-supported decision systems.'
        ]}
    ];
    
    document.getElementById('sp-sig-left').value = "Dhaka\nMarch 2026";
    document.getElementById('sp-sig-right').value = "MUHAMMAD SAMRAT HOSSAIN\nWing Commander\nCourse Participant";
    document.getElementById('sp-distribution').value = "Wing Commander Arman Nur Meraz, psc, GD(P)\nAir Directing Staff\nDefence Services Command and Staff College\nMirpur Cantonment\nDhaka";
    
    renderSpParagraphsEditor();
    updateSpPreview();
};
