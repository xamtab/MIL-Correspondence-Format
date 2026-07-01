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
        refContainer.style.display = 'flex';
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
        mainRow.innerHTML = `
            <span class="preview-para-num">${idx + 1}.</span>
            <div class="preview-para-content">${para.text}</div>
        `;
        paraEl.appendChild(mainRow);
        
        // Nested sub-paragraphs
        if (para.subParas.length > 0) {
            para.subParas.forEach((subText, subIdx) => {
                const subRow = document.createElement('div');
                subRow.className = 'preview-sub-para';
                const subLetter = String.fromCharCode(97 + subIdx);
                subRow.innerHTML = `
                    <span class="preview-sub-para-num">${subLetter}.</span>
                    <div class="preview-para-content">${subText}</div>
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
        distEl.innerHTML += '<div class="dist-section-header">Distribution:</div>';
        
        // Build 2-column or list view
        const grid = document.createElement('div');
        grid.className = 'dist-col-grid';
        
        // Column 1: External Action / Info
        const col1 = document.createElement('div');
        if (act || info) {
            col1.innerHTML += '<div class="dist-sub-section">External:</div>';
            if (act) {
                const actDiv = document.createElement('div');
                actDiv.className = 'dist-list';
                actDiv.innerHTML = '<span style="text-decoration:underline;">Action:</span>';
                act.split('\n').forEach(line => {
                    actDiv.innerHTML += `<div class="dist-item">${line}</div>`;
                });
                col1.appendChild(actDiv);
            }
            if (info) {
                const infoDiv = document.createElement('div');
                infoDiv.className = 'dist-list';
                infoDiv.innerHTML = '<span style="text-decoration:underline;">Information:</span>';
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
            col2.innerHTML += '<div class="dist-sub-section">Internal:</div>';
            if (intAct) {
                const intActDiv = document.createElement('div');
                intActDiv.className = 'dist-list';
                intActDiv.innerHTML = '<span style="text-decoration:underline;">Action:</span>';
                intAct.split('\n').forEach(line => {
                    intActDiv.innerHTML += `<div class="dist-item">${line}</div>`;
                });
                col2.appendChild(intActDiv);
            }
            if (intInfo) {
                const intInfoDiv = document.createElement('div');
                intInfoDiv.className = 'dist-list';
                intInfoDiv.innerHTML = '<span style="text-decoration:underline;">Information:</span>';
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
    
    let text = `\t\t\t\t${sec}\n\n`;
    
    if (prec) text += `\t\t\t\t\t${prec}\n`;
    if (copyVal) text += `\t\t\t\t\tCopy No ${copyVal}\n`;
    if (pagesVal) text += `\t\t\t\t\tTotal Pages ${pagesVal}\n`;
    
    text += `${sender.split('\n').map(l => '\t\t\t\t\t' + l).join('\n')}\n\n`;
    text += `${ref}\t\t\t\t${date}\n\n`;
    text += `${subject.toUpperCase()}\n\n`;
    
    if (letterReferences.length > 0) {
        text += `Refs:\n`;
        letterReferences.forEach((r, i) => {
            const letter = String.fromCharCode(65 + i);
            text += `${letter}.\t${r}\n`;
        });
        text += `\n`;
    }
    
    letterParagraphs.forEach((p, idx) => {
        text += `${idx + 1}.\t${p.text}\n`;
        p.subParas.forEach((sub, subIdx) => {
            const letter = String.fromCharCode(97 + subIdx);
            text += `\t${letter}.\t${sub}\n`;
        });
        text += `\n`;
    });
    
    const sigName = document.getElementById('let-sig-name').value;
    const sigRank = document.getElementById('let-sig-rank').value;
    const sigAppt = document.getElementById('let-sig-appt').value;
    
    text += `\n\t\t\t\t\t${sigName.toUpperCase()}\n`;
    text += `\t\t\t\t\t${sigRank}\n`;
    text += `\t\t\t\t\t${sigAppt}\n\n`;
    
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

    // --- Top Security Classification (Centered) ---
    body += `<p style="text-align:center; font-family:Arial,sans-serif; font-weight:bold; font-size:11pt; letter-spacing:1px; margin-bottom:24pt;">${sec}</p>\n`;

    // --- Precedence, Copy, Pages + Sender Address (Table layout: sender left, meta right) ---
    let metaRight = '';
    if (prec) metaRight += `<p style="font-family:Arial,sans-serif; font-weight:bold; font-size:12pt; color:#cc0000; margin:0 0 4pt 0;">${prec}</p>`;
    if (copyVal) metaRight += `<p style="font-family:Arial,sans-serif; font-size:10pt; margin:0 0 2pt 0;">Copy No ${copyVal}</p>`;
    if (pagesVal) metaRight += `<p style="font-family:Arial,sans-serif; font-size:10pt; margin:0 0 2pt 0;">Total Pages ${pagesVal}</p>`;

    const senderLines = sender.split('\n').map(l => l.trim()).filter(l => l).join('<br>');

    body += `<table style="width:100%; border:none; border-collapse:collapse; margin-bottom:16pt;">
        <tr>
            <td style="vertical-align:top; width:50%; padding:0; border:none;">&nbsp;</td>
            <td style="vertical-align:top; text-align:left; padding:0; border:none;">
                ${metaRight}
                <p style="font-family:Arial,sans-serif; font-size:11pt; line-height:1.4; margin-top:8pt;">${senderLines}</p>
            </td>
        </tr>
    </table>\n`;

    // --- Reference Number and Date (Table: ref left, date right) ---
    body += `<table style="width:100%; border:none; border-collapse:collapse; margin-bottom:16pt;">
        <tr>
            <td style="text-align:left; font-family:Arial,sans-serif; font-weight:500; font-size:11pt; padding:0; border:none;">${ref}</td>
            <td style="text-align:right; font-family:Arial,sans-serif; font-size:11pt; padding:0; border:none;">${date}</td>
        </tr>
    </table>\n`;

    // --- Subject Heading (Bold, Underlined, Uppercase) ---
    if (subject.trim()) {
        body += `<p style="font-family:Arial,sans-serif; font-weight:bold; font-size:12pt; text-decoration:underline; text-transform:uppercase; margin-bottom:12pt;">${subject.toUpperCase()}</p>\n`;
    }

    // --- References ---
    if (letterReferences.length > 0) {
        body += `<table style="border:none; border-collapse:collapse; margin-bottom:16pt;">
            <tr>
                <td style="vertical-align:top; font-family:Arial,sans-serif; font-weight:bold; text-decoration:underline; font-size:11pt; padding:0 12pt 0 0; border:none; white-space:nowrap;">Refs:</td>
                <td style="vertical-align:top; padding:0; border:none;">`;
        letterReferences.forEach((refText, idx) => {
            const letter = String.fromCharCode(65 + idx);
            body += `<p style="font-family:Arial,sans-serif; font-size:11pt; margin:0 0 4pt 0;">${letter}.&nbsp;&nbsp;&nbsp;&nbsp;${refText}</p>`;
        });
        body += `</td></tr></table>\n`;
    }

    // --- Paragraphs (Body) ---
    letterParagraphs.forEach((para, idx) => {
        // Main paragraph: numbered with indentation
        body += `<table style="width:100%; border:none; border-collapse:collapse; margin-bottom:12pt;">
            <tr>
                <td style="vertical-align:top; width:30pt; font-family:Arial,sans-serif; font-size:12pt; font-weight:500; padding:0; border:none;">${idx + 1}.</td>
                <td style="vertical-align:top; font-family:Arial,sans-serif; font-size:12pt; text-align:justify; padding:0; border:none;">${para.text}</td>
            </tr>
        </table>\n`;

        // Sub-paragraphs: lettered with deeper indent
        if (para.subParas.length > 0) {
            para.subParas.forEach((subText, subIdx) => {
                const subLetter = String.fromCharCode(97 + subIdx);
                body += `<table style="width:100%; border:none; border-collapse:collapse; margin-bottom:6pt; margin-left:30pt;">
                    <tr>
                        <td style="vertical-align:top; width:24pt; font-family:Arial,sans-serif; font-size:12pt; padding:0; border:none;">${subLetter}.</td>
                        <td style="vertical-align:top; font-family:Arial,sans-serif; font-size:12pt; text-align:justify; padding:0; border:none;">${subText}</td>
                    </tr>
                </table>\n`;
            });
        }
    });

    // --- Signature Block (Right-aligned) ---
    body += `<table style="width:100%; border:none; border-collapse:collapse; margin-top:36pt; margin-bottom:24pt;">
        <tr>
            <td style="width:55%; border:none;">&nbsp;</td>
            <td style="vertical-align:top; text-align:left; font-family:Arial,sans-serif; padding:0; border:none;">
                <p style="font-weight:bold; text-transform:uppercase; margin:0 0 2pt 0; font-size:12pt;">${sigName.toUpperCase()}</p>
                <p style="margin:0 0 2pt 0; font-size:11pt;">${sigRank}</p>
                <p style="margin:0; font-size:11pt;">${sigAppt}</p>
            </td>
        </tr>
    </table>\n`;

    // --- Distribution Block ---
    if (act || info || intAct || intInfo) {
        body += `<div style="border-top:1px solid #000000; padding-top:12pt; margin-top:12pt; font-family:Arial,sans-serif; font-size:10.5pt;">`;
        body += `<p style="font-weight:bold; text-decoration:underline; margin-bottom:8pt;">Distribution:</p>`;

        body += `<table style="width:100%; border:none; border-collapse:collapse;">
            <tr>`;

        // Column 1: External
        body += `<td style="vertical-align:top; width:50%; padding:0 12pt 0 0; border:none;">`;
        if (act || info) {
            body += `<p style="font-weight:600; margin-bottom:4pt;">External:</p>`;
            if (act) {
                body += `<p style="text-decoration:underline; margin:0 0 2pt 12pt;">Action:</p>`;
                act.split('\n').forEach(line => {
                    body += `<p style="margin:0 0 1pt 12pt;">${line}</p>`;
                });
            }
            if (info) {
                body += `<p style="text-decoration:underline; margin:8pt 0 2pt 12pt;">Information:</p>`;
                info.split('\n').forEach(line => {
                    body += `<p style="margin:0 0 1pt 12pt;">${line}</p>`;
                });
            }
        }
        body += `</td>`;

        // Column 2: Internal
        body += `<td style="vertical-align:top; width:50%; padding:0; border:none;">`;
        if (intAct || intInfo) {
            body += `<p style="font-weight:600; margin-bottom:4pt;">Internal:</p>`;
            if (intAct) {
                body += `<p style="text-decoration:underline; margin:0 0 2pt 12pt;">Action:</p>`;
                intAct.split('\n').forEach(line => {
                    body += `<p style="margin:0 0 1pt 12pt;">${line}</p>`;
                });
            }
            if (intInfo) {
                body += `<p style="text-decoration:underline; margin:8pt 0 2pt 12pt;">Information:</p>`;
                intInfo.split('\n').forEach(line => {
                    body += `<p style="margin:0 0 1pt 12pt;">${line}</p>`;
                });
            }
        }
        body += `</td></tr></table>`;
        body += `</div>\n`;
    }

    // --- Bottom Security Classification (Centered) ---
    body += `<p style="text-align:center; font-family:Arial,sans-serif; font-weight:bold; font-size:11pt; letter-spacing:1px; margin-top:24pt;">${sec}</p>\n`;

    // --- Build the full Word-compatible HTML document ---
    const fullHTML = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>MIL-Correspondence-Format - JSSDM Letter</title>
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
    @page {
        size: 21cm 29.7cm;
        margin: 3cm 2cm 3cm 2cm;
    }
    body {
        font-family: Arial, sans-serif;
        font-size: 12pt;
        line-height: 1.5;
        color: #000000;
    }
    table {
        border: none;
        border-collapse: collapse;
    }
    td {
        border: none;
        padding: 0;
    }
    p {
        font-family: Arial, sans-serif;
    }
</style>
</head>
<body>
${body}
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
