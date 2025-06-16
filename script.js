document.addEventListener('DOMContentLoaded', () => {
    // --- ELEMENT SELECTION ---
    const editor = document.getElementById('text-input');
    // ... all other element selections from previous versions
    const officeButton = document.getElementById('office-button');
    const officeMenu = document.getElementById('office-button-menu');
    const ribbonTabs = document.querySelectorAll('.ribbon-tab');
    const ribbonSections = document.querySelectorAll('.ribbon-section');
    const statusWords = document.getElementById('status-words');

    // --- INITIALIZATION ---
    const init = () => {
        // Load from localStorage
        const savedText = localStorage.getItem('wordWebEditorContent');
        if (savedText) {
            editor.innerHTML = savedText;
        } else {
            editor.innerHTML = '<h1>Welcome to Your Web Word Processor!</h1><p>This is a demonstration of a Microsoft Word 2007 replica built with HTML, CSS, and JavaScript. You can open and save <b>.docx</b> files using the Office Button in the top-left corner.</p>';
        }
        updateAllStats();
        setupEventListeners();
    };

    // --- EVENT LISTENERS SETUP ---
    const setupEventListeners = () => {
        // Editor events
        editor.addEventListener('input', updateAllStats);
        editor.addEventListener('keyup', updateButtonStates);
        editor.addEventListener('mouseup', updateButtonStates);

        // Office Button
        officeButton.addEventListener('click', () => officeMenu.classList.toggle('active'));
        document.addEventListener('click', (e) => {
            if (!officeButton.contains(e.target) && !officeMenu.contains(e.target)) {
                officeMenu.classList.remove('active');
            }
        });

        // Ribbon Tabs
        ribbonTabs.forEach(tab => {
            tab.addEventListener('click', () => {
                const targetTab = tab.dataset.tab;
                ribbonTabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                ribbonSections.forEach(section => {
                    section.classList.toggle('active', section.id === `tab-${targetTab}`);
                });
            });
        });

        // All ribbon commands
        setupCommandButtons();
    };

    // --- CORE FUNCTIONS ---
    const updateAllStats = () => {
        const text = editor.innerText;
        const words = text.trim().split(/\s+/).filter(Boolean);
        const wordCount = words.length;
        const charCount = text.length;
        const sentenceCount = (text.match(/[.!?]+|\n+/g) || []).length || (text.trim() ? 1 : 0);
        const paragraphCount = editor.innerHTML.split(/<(p|div|h[1-6]|blockquote)[^>]*>/gi).length - 1;

        // Update UI
        document.getElementById('details-words').textContent = wordCount;
        document.getElementById('details-chars').textContent = charCount;
        document.getElementById('details-sentences').textContent = sentenceCount;
        document.getElementById('details-paragraphs').textContent = paragraphCount;
        document.getElementById('details-reading-time').textContent = `${Math.ceil(wordCount / 270)} min`;
        statusWords.textContent = `Words: ${wordCount}`;
        
        // Save to localStorage
        localStorage.setItem('wordWebEditorContent', editor.innerHTML);
    };

    const updateButtonStates = () => {
        const commands = ['bold', 'italic', 'underline', 'strikethrough', 'subscript', 'superscript'];
        commands.forEach(cmd => {
            const btn = document.getElementById(`btn-${cmd}`);
            if (btn) btn.classList.toggle('is-active', document.queryCommandState(cmd));
        });
    };

    const formatDoc = (command, value = null) => {
        editor.focus();
        document.execCommand(command, false, value);
    };
    
    // --- COMMAND BUTTONS SETUP ---
    const setupCommandButtons = () => {
        const commands = {
            'btn-undo': { cmd: 'undo' }, 'btn-redo': { cmd: 'redo' },
            'btn-bold': { cmd: 'bold' }, 'btn-italic': { cmd: 'italic' },
            'btn-underline': { cmd: 'underline' }, 'btn-strikethrough': { cmd: 'strikethrough' },
            'btn-subscript': { cmd: 'subscript' }, 'btn-superscript': { cmd: 'superscript' },
            'btn-align-left': { cmd: 'justifyLeft' }, 'btn-align-center': { cmd: 'justifyCenter' },
            'btn-align-right': { cmd: 'justifyRight' }, 'btn-align-justify': { cmd: 'justifyFull' },
            'btn-ordered-list': { cmd: 'insertOrderedList' }, 'btn-unordered-list': { cmd: 'insertUnorderedList' },
            'btn-clear-format': { cmd: 'removeFormat' }
        };

        for (const [id, config] of Object.entries(commands)) {
            document.getElementById(id).addEventListener('click', () => formatDoc(config.cmd, config.val));
        }

        // Select-based commands
        document.getElementById('heading-select').addEventListener('change', e => formatDoc('formatBlock', e.target.value));
        document.getElementById('font-family-select').addEventListener('change', e => formatDoc('fontName', e.target.value));
        document.getElementById('font-size-select').addEventListener('change', e => formatDoc('fontSize', e.target.value));
        
        // Color pickers
        const fontColorPicker = document.getElementById('font-color-picker');
        const highlightColorPicker = document.getElementById('highlight-color-picker');
        fontColorPicker.addEventListener('input', e => {
            formatDoc('foreColor', e.target.value);
            document.getElementById('font-color-icon').style.borderBottomColor = e.target.value;
        });
        highlightColorPicker.addEventListener('input', e => {
            formatDoc('hiliteColor', e.target.value);
            document.getElementById('highlight-color-icon').style.borderBottomColor = e.target.value;
        });

        // File operations
        document.getElementById('btn-print').addEventListener('click', () => window.print());
        document.getElementById('btn-save-html').addEventListener('click', saveAsHtml);
        document.getElementById('btn-save-docx').addEventListener('click', saveAsDocx);
        document.getElementById('btn-open-docx').addEventListener('click', () => document.getElementById('docx-file-input').click());
        document.getElementById('docx-file-input').addEventListener('change', openDocx);
        document.getElementById('btn-insert-image').addEventListener('click', () => document.getElementById('image-file-input').click());
        document.getElementById('image-file-input').addEventListener('change', insertImage);
    };

    // --- FILE OPERATIONS ---
    const saveAsHtml = () => {
        const htmlContent = new Blob([editor.innerHTML], { type: "text/html" });
        saveAs(htmlContent, "document.html");
    };

    const saveAsDocx = () => {
        const doc = new docx.Document({
            sections: [{
                children: [
                    // This is a simplified conversion. A full one is much more complex.
                    // We'll just take the plain text for now for simplicity. A more advanced
                    // parser would be needed to convert HTML tags to docx.js objects.
                    new docx.Paragraph({ text: editor.innerText })
                ]
            }]
        });

        docx.Packer.toBlob(doc).then(blob => {
            saveAs(blob, "document.docx");
        });
    };
    
    const openDocx = (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const arrayBuffer = e.target.result;
            mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
                .then(result => {
                    editor.innerHTML = result.value;
                    updateAllStats();
                })
                .catch(err => console.log(err));
        };
        reader.readAsArrayBuffer(file);
    };

    const insertImage = (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            formatDoc('insertImage', e.target.result);
        };
        reader.readAsDataURL(file);
    };

    // --- START THE APP ---
    init();
});
