const CW_URL_PATTERN = /https?:\/\/([^/]+)\/([^/]+)\/e\/eng\/([^/]+)/;
const CW_DOC_PATTERN = /#\/(?:efinancials|letter)\/([^/?\s]+)/;

const urlInput = document.getElementById('cwUrl');
const parsedInfo = document.getElementById('parsedInfo');
const generateBtn = document.getElementById('generateBtn');
const statusEl = document.getElementById('status');
const errorEl = document.getElementById('error');
const successEl = document.getElementById('success');

urlInput.addEventListener('input', function () {
    const match = this.value.trim().match(CW_URL_PATTERN);
    if (match) {
        const docMatch = this.value.trim().match(CW_DOC_PATTERN);
        let info = 'Tenant: ' + match[2] +
            '  |  Engagement: ' + match[3].slice(0, 12) + '\u2026';
        if (docMatch) {
            info += '  |  Document: ' + docMatch[1].slice(0, 12) + '\u2026';
        } else {
            info += '  \u2014  please include a #/letter/<id> fragment';
        }
        parsedInfo.textContent = info;
        parsedInfo.classList.add('parsed-success');
        urlInput.classList.remove('input-error');
    } else if (this.value.trim()) {
        parsedInfo.textContent = 'URL not recognized \u2014 expected a Caseware engagement URL';
        parsedInfo.classList.remove('parsed-success');
    } else {
        parsedInfo.textContent = 'Paste a Caseware document URL to export a letter (e.g. ...#/letter/<documentId>)';
        parsedInfo.classList.remove('parsed-success');
    }
});

generateBtn.addEventListener('click', async function () {
    clearMessages();

    const url = urlInput.value.trim();
    if (!url) {
        showError('Please paste a Caseware URL.');
        urlInput.classList.add('input-error');
        urlInput.focus();
        return;
    }

    const match = url.match(CW_URL_PATTERN);
    if (!match) {
        showError('Invalid URL format. Expected a Caseware engagement URL (https://<host>/<tenant>/e/eng/<id>/...)');
        urlInput.classList.add('input-error');
        urlInput.focus();
        return;
    }

    const docMatch = url.match(CW_DOC_PATTERN);
    if (!docMatch) {
        showError('URL must include a document fragment (e.g. #/letter/<documentId>)');
        urlInput.classList.add('input-error');
        urlInput.focus();
        return;
    }

    const templateName = document.getElementById('templateName').value.trim();
    setLoading(true);

    try {
        const response = await fetch('/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url, templateName: templateName || 'Letter' }),
        });

        if (!response.ok) {
            let errMsg = 'An error occurred while generating the letter.';
            try {
                const err = await response.json();
                errMsg = err.error || errMsg;
            } catch (_) {}
            throw new Error(errMsg);
        }

        const blob = await response.blob();
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;

        const disposition = response.headers.get('Content-Disposition');
        const filenameMatch = disposition && disposition.match(/filename="?([^"]+)"?/);
        a.download = filenameMatch ? filenameMatch[1] : 'letter.docx';

        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(blobUrl);

        showSuccess('Letter exported successfully.');
    } catch (e) {
        showError(e.message);
    } finally {
        setLoading(false);
    }
});

urlInput.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') generateBtn.click();
});

function setLoading(loading) {
    generateBtn.disabled = loading;
    statusEl.hidden = !loading;
}

function showError(msg) {
    errorEl.textContent = msg;
    errorEl.hidden = false;
}

function showSuccess(msg) {
    successEl.textContent = msg;
    successEl.hidden = false;
}

function clearMessages() {
    errorEl.hidden = true;
    successEl.hidden = true;
    urlInput.classList.remove('input-error');
}
