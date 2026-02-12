/**
 * PDFConverter - Frontend Application Logic
 * Handles file upload, conversion configuration, and result management.
 */

"use strict";

// ============================================================
// State Management
// ============================================================
const state = {
    file: null,
    storedName: null,
    pdfInfo: null,
    selectedFormat: null,
    currentPreviewPage: 0,
    isConverting: false,
};

// ============================================================
// DOM Elements
// ============================================================
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const dom = {
    // Upload
    dropZone: $("#drop-zone"),
    fileInput: $("#file-input"),
    fileInfo: $("#file-info"),
    fileName: $("#file-name"),
    fileMeta: $("#file-meta"),
    removeFile: $("#remove-file"),
    previewImage: $("#preview-image"),
    prevPage: $("#prev-page"),
    nextPage: $("#next-page"),
    pageIndicator: $("#page-indicator"),
    pageInfoGrid: $("#page-info-grid"),

    // Config
    uploadSection: $("#upload-section"),
    configSection: $("#config-section"),
    resultSection: $("#result-section"),
    formatGrid: $("#format-grid"),
    ocrToggle: $("#ocr-toggle"),
    ocrOptions: $("#ocr-options"),
    ocrLang: $("#ocr-lang"),
    imageOptions: $("#image-options"),
    imageDpi: $("#image-dpi"),
    convertBtn: $("#convert-btn"),

    // Result
    progressContainer: $("#progress-container"),
    progressFill: $("#progress-fill"),
    progressText: $("#progress-text"),
    resultCard: $("#result-card"),
    resultIcon: $("#result-icon"),
    resultTitle: $("#result-title"),
    resultDetails: $("#result-details"),
    downloadBtn: $("#download-btn"),
    convertAnother: $("#convert-another"),

    // Toast
    toastContainer: $("#toast-container"),
};

// ============================================================
// Utility Functions
// ============================================================
function formatFileSize(bytes) {
    if (bytes === 0) return "0 B";
    const k = 1024;
    const sizes = ["B", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

function showToast(message, type = "info") {
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;

    const icon = type === "error" ? "error" : type === "success" ? "check_circle" : "info";
    toast.innerHTML = `
        <span class="material-icons-round" style="font-size:20px">${icon}</span>
        <span>${message}</span>
    `;

    dom.toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.classList.add("removing");
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

async function apiRequest(url, options = {}) {
    try {
        const response = await fetch(url, options);
        const data = await response.json();
        if (!response.ok) {
            throw new Error(data.error || `HTTP ${response.status}`);
        }
        return data;
    } catch (error) {
        if (error.name === "TypeError" && error.message.includes("fetch")) {
            throw new Error("Cannot connect to server. Make sure the server is running.");
        }
        throw error;
    }
}

// ============================================================
// File Upload
// ============================================================
function initDragDrop() {
    const dz = dom.dropZone;

    ["dragenter", "dragover", "dragleave", "drop"].forEach((evt) => {
        dz.addEventListener(evt, (e) => {
            e.preventDefault();
            e.stopPropagation();
        });
    });

    ["dragenter", "dragover"].forEach((evt) => {
        dz.addEventListener(evt, () => dz.classList.add("dragover"));
    });

    ["dragleave", "drop"].forEach((evt) => {
        dz.addEventListener(evt, () => dz.classList.remove("dragover"));
    });

    dz.addEventListener("drop", (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) handleFileSelect(files[0]);
    });

    dz.addEventListener("click", () => dom.fileInput.click());
    dom.fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) handleFileSelect(e.target.files[0]);
    });
}

async function handleFileSelect(file) {
    // Validate
    if (!file.name.toLowerCase().endsWith(".pdf")) {
        showToast("Please select a PDF file", "error");
        return;
    }

    if (file.size > 500 * 1024 * 1024) {
        showToast("File size exceeds 500 MB limit", "error");
        return;
    }

    state.file = file;
    showToast(`Uploading "${file.name}"...`, "info");

    // Upload
    const formData = new FormData();
    formData.append("file", file);

    try {
        const info = await apiRequest("/api/upload", {
            method: "POST",
            body: formData,
        });

        if (info.error) {
            showToast(info.error, "error");
            return;
        }

        state.storedName = info.stored_name;
        state.pdfInfo = info;

        // Update UI
        displayFileInfo(info);
        showToast("File uploaded successfully!", "success");
    } catch (error) {
        showToast(`Upload failed: ${error.message}`, "error");
    }
}

function displayFileInfo(info) {
    dom.dropZone.style.display = "none";
    dom.fileInfo.classList.remove("hidden");

    dom.fileName.textContent = info.filename;
    dom.fileMeta.textContent = `${info.page_count} pages • ${formatFileSize(info.file_size)}`;

    // Load preview
    state.currentPreviewPage = 0;
    loadPreview(0);

    // Page info grid
    dom.pageInfoGrid.innerHTML = "";
    info.pages.forEach((page) => {
        const badge = document.createElement("div");
        let cssClass = "page-badge";
        let icon = "description";
        let label = `Page ${page.number}`;

        if (page.needs_ocr) {
            cssClass += " ocr-needed";
            icon = "document_scanner";
            label += " (OCR needed)";
        } else if (page.has_text) {
            cssClass += " has-text";
            icon = "text_snippet";
            label += " (Has text)";
        }

        badge.className = cssClass;
        badge.innerHTML = `
            <span class="material-icons-round">${icon}</span>
            <span>${label}</span>
        `;
        dom.pageInfoGrid.appendChild(badge);
    });

    // Show config section
    dom.configSection.classList.remove("hidden");
}

function loadPreview(pageNum) {
    if (!state.storedName) return;
    dom.previewImage.src = `/api/preview/${state.storedName}?page=${pageNum}`;
    dom.pageIndicator.textContent = `Page ${pageNum + 1} / ${state.pdfInfo.page_count}`;
}

function initPreviewNav() {
    dom.prevPage.addEventListener("click", () => {
        if (state.currentPreviewPage > 0) {
            state.currentPreviewPage--;
            loadPreview(state.currentPreviewPage);
        }
    });

    dom.nextPage.addEventListener("click", () => {
        if (state.pdfInfo && state.currentPreviewPage < state.pdfInfo.page_count - 1) {
            state.currentPreviewPage++;
            loadPreview(state.currentPreviewPage);
        }
    });
}

// ============================================================
// Format Selection
// ============================================================
function initFormatSelection() {
    $$(".format-card").forEach((card) => {
        card.addEventListener("click", () => {
            // Deselect all
            $$(".format-card").forEach((c) => c.classList.remove("selected"));
            // Select clicked
            card.classList.add("selected");
            state.selectedFormat = card.dataset.format;

            // Show/hide image options
            if (state.selectedFormat === "image") {
                dom.imageOptions.classList.remove("hidden");
            } else {
                dom.imageOptions.classList.add("hidden");
            }

            // Enable convert button
            dom.convertBtn.disabled = false;
        });
    });
}

function initOCRToggle() {
    dom.ocrToggle.addEventListener("change", () => {
        dom.ocrOptions.style.display = dom.ocrToggle.checked ? "block" : "none";
    });
}

// ============================================================
// Conversion
// ============================================================
function initConversion() {
    dom.convertBtn.addEventListener("click", startConversion);
}

async function startConversion() {
    if (!state.storedName || !state.selectedFormat || state.isConverting) return;

    state.isConverting = true;
    dom.convertBtn.disabled = true;
    dom.convertBtn.innerHTML = '<span class="spinner"></span><span>Converting...</span>';

    // Show result section with progress
    dom.resultSection.classList.remove("hidden");
    dom.progressContainer.style.display = "block";
    dom.resultCard.classList.add("hidden");

    // Animate progress
    animateProgress(0, 80, 3000);

    try {
        const payload = {
            stored_name: state.storedName,
            format: state.selectedFormat,
            use_ocr: dom.ocrToggle.checked,
            ocr_lang: dom.ocrLang.value,
            dpi: parseInt(dom.imageDpi.value),
            image_format: "png",
        };

        const result = await apiRequest("/api/convert", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });

        if (result.success) {
            animateProgress(80, 100, 300);
            setTimeout(() => showResult(result), 500);
        } else {
            showError(result.error || "Conversion failed");
        }
    } catch (error) {
        showError(error.message);
    } finally {
        state.isConverting = false;
        dom.convertBtn.disabled = false;
        dom.convertBtn.innerHTML = '<span class="material-icons-round">transform</span><span>Convert Now</span>';
    }
}

function animateProgress(from, to, duration) {
    const startTime = Date.now();
    const update = () => {
        const elapsed = Date.now() - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const value = from + (to - from) * easeOut(progress);
        dom.progressFill.style.width = `${value}%`;
        dom.progressText.textContent = `Converting... ${Math.round(value)}%`;
        if (progress < 1) requestAnimationFrame(update);
    };
    requestAnimationFrame(update);
}

function easeOut(t) {
    return 1 - Math.pow(1 - t, 3);
}

function showResult(result) {
    dom.progressContainer.style.display = "none";
    dom.resultCard.classList.remove("hidden");

    dom.resultIcon.className = "material-icons-round result-icon success";
    dom.resultIcon.textContent = "check_circle";
    dom.resultTitle.textContent = "Conversion Complete!";

    const time = result.time ? `${result.time.toFixed(2)}s` : "N/A";
    let details = `Converted to ${state.selectedFormat.toUpperCase()} in ${time}`;
    if (result.page_count) details += ` • ${result.page_count} pages`;
    if (result.tables_found !== undefined) {
        details += result.tables_found ? " • Tables detected" : " • No tables found (text extracted)";
    }
    if (result.fallback) details += " • (Used fallback method)";
    dom.resultDetails.textContent = details;

    // Setup download
    const outputPath = result.output;
    const fileName = outputPath.split(/[/\\]/).pop();
    dom.downloadBtn.href = `/api/download/${fileName}`;
    dom.downloadBtn.download = fileName;

    showToast("Conversion successful!", "success");
}

function showError(message) {
    dom.progressContainer.style.display = "none";
    dom.resultCard.classList.remove("hidden");

    dom.resultIcon.className = "material-icons-round result-icon error";
    dom.resultIcon.textContent = "error";
    dom.resultTitle.textContent = "Conversion Failed";
    dom.resultDetails.textContent = message;
    dom.downloadBtn.style.display = "none";

    showToast(`Error: ${message}`, "error");
}

// ============================================================
// Reset / Convert Another
// ============================================================
function initReset() {
    dom.removeFile.addEventListener("click", resetAll);
    dom.convertAnother.addEventListener("click", resetAll);
}

function resetAll() {
    // Reset state
    state.file = null;
    state.storedName = null;
    state.pdfInfo = null;
    state.selectedFormat = null;
    state.currentPreviewPage = 0;
    state.isConverting = false;

    // Reset UI
    dom.dropZone.style.display = "block";
    dom.fileInfo.classList.add("hidden");
    dom.configSection.classList.add("hidden");
    dom.resultSection.classList.add("hidden");
    dom.fileInput.value = "";

    // Reset format selection
    $$(".format-card").forEach((c) => c.classList.remove("selected"));
    dom.convertBtn.disabled = true;
    dom.convertBtn.innerHTML = '<span class="material-icons-round">transform</span><span>Convert Now</span>';

    // Reset result
    dom.progressContainer.style.display = "block";
    dom.progressFill.style.width = "0%";
    dom.resultCard.classList.add("hidden");
    dom.downloadBtn.style.display = "inline-flex";

    // Reset image options
    dom.imageOptions.classList.add("hidden");
}

// ============================================================
// Theme Toggle
// ============================================================
function initTheme() {
    const toggle = document.getElementById("theme-toggle");
    const html = document.documentElement;
    const saved = localStorage.getItem("pdf-theme");

    // Apply saved theme or default to light
    if (saved === "dark") {
        html.setAttribute("data-theme", "dark");
        toggle.querySelector(".material-icons-round").textContent = "light_mode";
    }

    toggle.addEventListener("click", () => {
        const current = html.getAttribute("data-theme");
        const next = current === "dark" ? "light" : "dark";
        html.setAttribute("data-theme", next);
        localStorage.setItem("pdf-theme", next);
        toggle.querySelector(".material-icons-round").textContent =
            next === "dark" ? "light_mode" : "dark_mode";
    });
}

// ============================================================
// Initialize
// ============================================================
document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    initDragDrop();
    initPreviewNav();
    initFormatSelection();
    initOCRToggle();
    initConversion();
    initReset();

    // Keyboard shortcuts
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") resetAll();
    });
});
