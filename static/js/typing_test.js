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

    const submitBtn           = document.getElementById('submitBtn');
    const userInput           = document.getElementById('user_input');
    const formTimeLimit       = document.getElementById('form_time_limit');   // fixed value from server
    const wpmInput            = document.getElementById('wpm');
    const accuracyInput       = document.getElementById('accuracy');
    const statsDiv            = document.getElementById('stats');
    const timerDiv            = document.getElementById('timer');
    const selectedParagraphInput = document.getElementById('selected_paragraph');
    const sampleTextContainer = document.querySelector('.sample-text');
    const typingForm          = document.getElementById('typingForm');

    // Read fixed time limit from hidden field (set by server, never changes)
    const TIME_LIMIT_SECONDS = formTimeLimit ? parseInt(formTimeLimit.value) : 300;

    const formElements = document.querySelectorAll('#typingForm input, #typingForm textarea');
    disableAutofillAndSuggestions(formElements);

    let startTime        = null;
    let timerInterval    = null;
    let isTestRunning    = false;
    let isTestStarted    = false;
    let sampleText       = sampleTextContainer ? sampleTextContainer.dataset.text.replace(/\s+/g, ' ').trim() : '';
    let sampleWords      = sampleText ? sampleText.split(/\s+/).filter(w => w.length > 0) : [];
    let currentWordIndex = 0;
    let correctWords     = 0;
    let totalWordsTyped  = 0;
    let lines            = [];
    let currentLineIndex = 0;
    const maxLines       = 4;

    function createLines() {
        lines = [];
        let words = [...sampleWords];
        let currentLine = [];
        let lineWordCount = 0;
        const wordsPerLine = 10;
        for (let i = 0; i < words.length; i++) {
            currentLine.push(words[i]);
            lineWordCount++;
            if (lineWordCount >= wordsPerLine || i === words.length - 1) {
                lines.push({ words: currentLine, startWordIndex: i - lineWordCount + 1 });
                currentLine = [];
                lineWordCount = 0;
            }
        }
    }

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

    function displayCurrentLines() {
        if (currentLineIndex < lines.length) {
            const visibleLines = lines.slice(currentLineIndex, currentLineIndex + maxLines);
            sampleTextContainer.innerHTML = visibleLines.map(line =>
                `<div class="line">${line.words.map((word, wordIdx) => {
                    const globalWordIndex = line.startWordIndex + wordIdx;
                    return `<span class="word" data-index="${globalWordIndex}">${word}</span>`;
                }).join('&nbsp;')}</div>`
            ).join('');
            highlightCurrentWord();
        } else {
            // Paragraph fully typed — submit immediately
            clearInterval(timerInterval);
            isTestRunning = false;
            if (userInput) userInput.disabled = true;
            const wpm      = calculateWPM();
            const accuracy = calculateAccuracy();
            if (wpmInput)      { wpmInput.value = wpm;           disableAutofillAndSuggestions([wpmInput]); }
            if (accuracyInput) { accuracyInput.value = accuracy; disableAutofillAndSuggestions([accuracyInput]); }
            if (statsDiv)  statsDiv.textContent  = `Paragraph completed! WPM: ${wpm} | Accuracy: ${accuracy}%`;
            if (timerDiv)  timerDiv.textContent  = 'Time Remaining: 0:00';
            if (typingForm) typingForm.submit();
        }
    }

    function highlightCurrentWord() {
        const spans      = sampleTextContainer.querySelectorAll('span.word');
        const inputWords = userInput ? userInput.value.trim().split(/\s+/).filter(w => w.length > 0) : [];
        spans.forEach(span => {
            const globalIndex = parseInt(span.dataset.index);
            span.classList.remove('current', 'correct', 'incorrect');
            if (globalIndex === currentWordIndex) {
                span.classList.add('current');
            } else if (globalIndex < currentWordIndex && globalIndex < inputWords.length) {
                span.classList.add(inputWords[globalIndex] === sampleWords[globalIndex] ? 'correct' : 'incorrect');
            }
        });
    }

    initializeSampleText();

    function startTimer() {
        startTime = Date.now();
        isTestStarted = true;
        isTestRunning = true;
        if (statsDiv)  statsDiv.textContent = 'WPM: 0 | Accuracy: 100%';
        if (timerDiv)  timerDiv.textContent = formatTime(TIME_LIMIT_SECONDS);

        timerInterval = setInterval(() => {
            const elapsed = (Date.now() - startTime) / 1000;
            if (elapsed >= TIME_LIMIT_SECONDS) {
                clearInterval(timerInterval);
                isTestRunning = false;
                if (userInput) userInput.disabled = true;
                const wpm      = calculateWPM();
                const accuracy = calculateAccuracy();
                if (wpmInput)      { wpmInput.value = wpm;           disableAutofillAndSuggestions([wpmInput]); }
                if (accuracyInput) { accuracyInput.value = accuracy; disableAutofillAndSuggestions([accuracyInput]); }
                if (statsDiv)  statsDiv.textContent = `Time's up! WPM: ${wpm} | Accuracy: ${accuracy}%`;
                if (timerDiv)  timerDiv.textContent = 'Time Remaining: 0:00';
                highlightCurrentWord();
                if (typingForm) typingForm.submit();
            } else {
                updateStats();
            }
        }, 500);
    }

    if (userInput) {
        disableAutofillAndSuggestions([userInput]);
        userInput.addEventListener('copy',  e => e.preventDefault());
        userInput.addEventListener('paste', e => e.preventDefault());
        userInput.addEventListener('cut',   e => e.preventDefault());

        userInput.addEventListener('keydown', e => {
            // Start timer on first real keypress
            if (!isTestStarted && !['Backspace','Delete','ArrowLeft','ArrowRight','ArrowUp','ArrowDown',' '].includes(e.key)) {
                startTimer();
            }

            if (!isTestRunning) return;

            if (e.key === ' ') {
                const inputText  = userInput.value.trim();
                const inputWords = inputText.split(/\s+/).filter(w => w.length > 0);
                const currentInput = inputWords[inputWords.length - 1] || '';
                if (currentInput.length > 0 && currentWordIndex < sampleWords.length) {
                    if (currentInput === sampleWords[currentWordIndex]) correctWords++;
                    totalWordsTyped++;
                    currentWordIndex++;
                    // Scroll lines if needed
                    let lastVisibleWordIndex = -1;
                    const visibleLines = lines.slice(currentLineIndex, currentLineIndex + maxLines);
                    if (visibleLines.length > 0) {
                        const lastLine = visibleLines[visibleLines.length - 1];
                        lastVisibleWordIndex = lastLine.startWordIndex + lastLine.words.length - 1;
                    }
                    if (currentWordIndex > lastVisibleWordIndex && currentLineIndex + maxLines < lines.length) {
                        currentLineIndex++;
                        displayCurrentLines();
                    }
                    highlightCurrentWord();
                    updateStats();
                    return;
                }
                e.preventDefault();
                return;
            }

            if (e.key === 'Backspace') {
                const inputText  = userInput.value;
                const cursorPos  = userInput.selectionStart;
                const inputWords = inputText.split(/\s+/).filter(w => w.length > 0);
                const previousWords = inputWords.slice(0, Math.min(inputWords.length, currentWordIndex)).join(' ');
                const boundary = previousWords.length + (inputWords.length > currentWordIndex ? 1 : 0);
                if (cursorPos <= boundary) e.preventDefault();
                return;
            }

            if (['Delete','ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(e.key)) {
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

    function formatTime(seconds) {
        const m = Math.floor(seconds / 60);
        const s = Math.floor(seconds % 60);
        return `Time Remaining: ${m}:${s.toString().padStart(2, '0')}`;
    }

    function calculateWPM() {
        if (!startTime) return 0;
        const elapsed  = Math.min((Date.now() - startTime) / 1000, TIME_LIMIT_SECONDS);
        const minutes  = elapsed / 60;
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