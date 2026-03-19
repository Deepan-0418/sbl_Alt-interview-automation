document.addEventListener('DOMContentLoaded', () => {

    function disableAutofillAndSuggestions(elements) {
        elements.forEach(element => {
            element.setAttribute('autocomplete', 'off');
            if (element.tagName === 'INPUT' || element.tagName === 'TEXTAREA') {
                element.setAttribute('autocorrect', 'off');
                element.setAttribute('autocapitalize', 'off');
                element.setAttribute('spellcheck', 'false');
            }
        });
    }

    const userInput              = document.getElementById('user_input');
    const formTimeLimit          = document.getElementById('form_time_limit');
    const wpmInput               = document.getElementById('wpm');
    const accuracyInput          = document.getElementById('accuracy');
    const statsDiv               = document.getElementById('stats');
    const timerDiv               = document.getElementById('timer');
    const selectedParagraphInput = document.getElementById('selected_paragraph');
    const sampleTextContainer    = document.querySelector('.sample-text');
    const typingForm             = document.getElementById('typingForm');

    // Fixed time limit from server — never changes during a session
    const TIME_LIMIT_SECONDS = formTimeLimit ? parseInt(formTimeLimit.value) : 300;

    disableAutofillAndSuggestions(
        document.querySelectorAll('#typingForm input, #typingForm textarea')
    );

    let startTime        = null;
    let timerInterval    = null;
    let isTestRunning    = false;
    let isTestStarted    = false;

    let sampleText  = sampleTextContainer
        ? sampleTextContainer.dataset.text.replace(/\s+/g, ' ').trim()
        : '';
    let sampleWords = sampleText
        ? sampleText.split(/\s+/).filter(w => w.length > 0)
        : [];

    let currentWordIndex = 0;
    let correctWords     = 0;
    let totalWordsTyped  = 0;
    let wordResults      = []; // true/false per submitted word index
    let lastSpaceCount   = 0;  // tracks how many spaces have been submitted

    let lines            = [];
    let currentLineIndex = 0;
    const maxLines       = 4;

    // ── Line creation ─────────────────────────────────────────
    function createLines() {
        lines = [];
        let currentLine   = [];
        let lineWordCount = 0;
        const wordsPerLine = 10;

        for (let i = 0; i < sampleWords.length; i++) {
            currentLine.push(sampleWords[i]);
            lineWordCount++;
            if (lineWordCount >= wordsPerLine || i === sampleWords.length - 1) {
                lines.push({
                    words:          currentLine,
                    startWordIndex: i - lineWordCount + 1,
                });
                currentLine   = [];
                lineWordCount = 0;
            }
        }
    }

    // ── Render visible lines ──────────────────────────────────
    function displayCurrentLines() {
        if (currentLineIndex < lines.length) {
            const visibleLines = lines.slice(currentLineIndex, currentLineIndex + maxLines);
            sampleTextContainer.innerHTML = visibleLines
                .map(line =>
                    `<div class="line">${line.words
                        .map((word, wordIdx) => {
                            const gi = line.startWordIndex + wordIdx;
                            return `<span class="word" data-index="${gi}">${word}</span>`;
                        })
                        .join('&nbsp;')}</div>`
                )
                .join('');
            highlightCurrentWord();
        } else {
            finishTest();
        }
    }

    // ── Word highlighting ─────────────────────────────────────
    function highlightCurrentWord() {
        if (!sampleTextContainer) return;
        const spans = sampleTextContainer.querySelectorAll('span.word');

        spans.forEach(span => {
            const gi = parseInt(span.dataset.index);
            span.classList.remove('current', 'correct', 'incorrect');

            if (gi === currentWordIndex) {
                span.classList.add('current');
            } else if (gi < currentWordIndex) {
                span.classList.add(wordResults[gi] === true ? 'correct' : 'incorrect');
            }
        });
    }

    // ── Finish test ───────────────────────────────────────────
    function finishTest() {
        clearInterval(timerInterval);
        isTestRunning = false;
        if (userInput) userInput.disabled = true;

        const wpm      = calculateWPM();
        const accuracy = calculateAccuracy();

        if (wpmInput)      { wpmInput.value = wpm;           disableAutofillAndSuggestions([wpmInput]); }
        if (accuracyInput) { accuracyInput.value = accuracy; disableAutofillAndSuggestions([accuracyInput]); }
        if (statsDiv)  statsDiv.textContent = `Completed! WPM: ${wpm} | Accuracy: ${accuracy}%`;
        if (timerDiv)  timerDiv.textContent = 'Time Remaining: 0:00';

        if (typingForm) typingForm.submit();
    }

    // ── Init ──────────────────────────────────────────────────
    function initializeSampleText() {
        if (sampleTextContainer && sampleText) {
            createLines();
            displayCurrentLines();
            if (selectedParagraphInput) {
                selectedParagraphInput.value = sampleText;
                disableAutofillAndSuggestions([selectedParagraphInput]);
            }
        }
    }

    initializeSampleText();

    // ── Timer ─────────────────────────────────────────────────
    function startTimer() {
        startTime     = Date.now();
        isTestStarted = true;
        isTestRunning = true;

        if (statsDiv) statsDiv.textContent = 'WPM: 0 | Accuracy: 100%';
        if (timerDiv) timerDiv.textContent = formatTime(TIME_LIMIT_SECONDS);

        timerInterval = setInterval(() => {
            const elapsed = (Date.now() - startTime) / 1000;
            if (elapsed >= TIME_LIMIT_SECONDS) {
                finishTest();
            } else {
                updateStats();
            }
        }, 500);
    }

    // ── Get the word the user is currently typing ─────────────
    // The textarea accumulates all words separated by spaces.
    // The "current word" is everything after the last space.
    function getCurrentWordFromInput() {
        if (!userInput) return '';
        const val   = userInput.value;
        const parts = val.split(' ');
        return parts[parts.length - 1];
    }

    // ── Keyboard handling ─────────────────────────────────────
    if (userInput) {
        disableAutofillAndSuggestions([userInput]);

        userInput.addEventListener('copy',  e => e.preventDefault());
        userInput.addEventListener('paste', e => e.preventDefault());
        userInput.addEventListener('cut',   e => e.preventDefault());

        userInput.addEventListener('keydown', e => {

            // ── Space: advance word ───────────────────────────
            if (e.key === ' ') {
                const currentInput = getCurrentWordFromInput();

                // Block double-space — ignore if no word has been typed yet
                if (currentInput.length === 0) {
                    e.preventDefault();
                    return;
                }

                // Start timer on first word submission if not already started
                if (!isTestStarted) startTimer();
                if (!isTestRunning) {
                    e.preventDefault();
                    return;
                }

                if (currentWordIndex < sampleWords.length) {
                    // Record result for this word
                    const wasCorrect = currentInput === sampleWords[currentWordIndex];
                    wordResults[currentWordIndex] = wasCorrect;
                    if (wasCorrect) correctWords++;
                    totalWordsTyped++;
                    currentWordIndex++;
                    lastSpaceCount++;

                    // Allow the space to be added naturally to the textarea
                    // so the user sees their typed history — do NOT preventDefault here

                    // Scroll lines if needed
                    let lastVisibleWordIndex = -1;
                    const visibleLines = lines.slice(currentLineIndex, currentLineIndex + maxLines);
                    if (visibleLines.length > 0) {
                        const lastLine = visibleLines[visibleLines.length - 1];
                        lastVisibleWordIndex = lastLine.startWordIndex + lastLine.words.length - 1;
                    }
                    if (
                        currentWordIndex > lastVisibleWordIndex &&
                        currentLineIndex + maxLines < lines.length
                    ) {
                        currentLineIndex++;
                        displayCurrentLines();
                    } else if (currentWordIndex >= sampleWords.length) {
                        finishTest();
                        return;
                    }

                    highlightCurrentWord();
                    updateStats();
                } else {
                    e.preventDefault();
                }
                return;
            }

            // ── Start timer on first real keypress ───────────
            if (
                !isTestStarted &&
                !['Backspace', 'Delete', 'ArrowLeft', 'ArrowRight',
                  'ArrowUp', 'ArrowDown'].includes(e.key)
            ) {
                startTimer();
            }

            if (!isTestRunning) return;

            // ── Backspace: restrict to current word only ──────
            // Prevent deleting into previously submitted words
            if (e.key === 'Backspace') {
                const val       = userInput.value;
                const cursorPos = userInput.selectionStart;
                // The boundary is the position right after the last space
                const lastSpacePos = val.lastIndexOf(' ');
                const boundary     = lastSpacePos + 1; // can't go before this
                if (cursorPos <= boundary) {
                    e.preventDefault();
                }
                return;
            }

            // ── Block navigation keys ─────────────────────────
            if (['Delete', 'ArrowLeft', 'ArrowRight',
                 'ArrowUp', 'ArrowDown'].includes(e.key)) {
                e.preventDefault();
            }
        });

        userInput.addEventListener('input', () => {
            if (isTestRunning) {
                userInput.selectionStart = userInput.selectionEnd = userInput.value.length;
                highlightCurrentWord();
                disableAutofillAndSuggestions([userInput]);
                updateStats();
            }
        });

        userInput.addEventListener('click', e => {
            if (isTestRunning) {
                e.preventDefault();
                userInput.selectionStart = userInput.selectionEnd = userInput.value.length;
            }
        });

        userInput.addEventListener('select', e => {
            if (isTestRunning) {
                e.preventDefault();
                userInput.selectionStart = userInput.selectionEnd = userInput.value.length;
            }
        });
    }

    // ── Helpers ───────────────────────────────────────────────
    function formatTime(seconds) {
        const m = Math.floor(seconds / 60);
        const s = Math.floor(seconds % 60);
        return `Time Remaining: ${m}:${s.toString().padStart(2, '0')}`;
    }

    function calculateWPM() {
        if (!startTime) return 0;
        const elapsed = Math.min((Date.now() - startTime) / 1000, TIME_LIMIT_SECONDS);
        const minutes = elapsed / 60;
        return minutes > 0 ? Math.round(correctWords / minutes) : 0;
    }

    function calculateAccuracy() {
        if (totalWordsTyped === 0) return 100;
        return Math.max(0, Math.round((correctWords / totalWordsTyped) * 100));
    }

    function updateStats() {
        if (!startTime || !isTestRunning) return;
        const elapsed       = (Date.now() - startTime) / 1000;
        const wpm           = calculateWPM();
        const accuracy      = calculateAccuracy();
        const remainingTime = Math.max(0, TIME_LIMIT_SECONDS - elapsed);

        if (statsDiv)      statsDiv.textContent = `WPM: ${wpm} | Accuracy: ${accuracy}%`;
        if (wpmInput)      { wpmInput.value = wpm;           disableAutofillAndSuggestions([wpmInput]); }
        if (accuracyInput) { accuracyInput.value = accuracy; disableAutofillAndSuggestions([accuracyInput]); }
        if (timerDiv)      timerDiv.textContent = formatTime(remainingTime);
    }

});