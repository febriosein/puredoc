let currentJobId = null;
let currentReportData = null;

document.addEventListener("DOMContentLoaded", () => {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const uploadBtn = document.getElementById("uploadBtn");
  const uploadSection = document.getElementById("uploadSection");
  const analysisSection = document.getElementById("analysisSection");
  const progressSection = document.getElementById("progressSection");
  const resultSection = document.getElementById("resultSection");
  const convertBtn = document.getElementById("convertBtn");
  const cancelBtn = document.getElementById("cancelBtn");
  const newDocBtn = document.getElementById("newDocBtn");
  const deleteBtn = document.getElementById("deleteBtn");
  const downloadReportBtn = document.getElementById("downloadReportBtn");

  // Drag and drop events
  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  });

  uploadBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFileSelected(e.target.files[0]);
    }
  });

  // Handle file upload and analysis
  async function handleFileSelected(file) {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      alert("Hanya file .docx yang didukung pada versi MVP ini.");
      return;
    }

    if (file.size > 100 * 1024 * 1024) {
      alert("Ukuran file melebihi batas 100 MB.");
      return;
    }

    // Switch to uploading state
    uploadSection.style.display = "none";
    progressSection.style.display = "block";
    updateProgressStep(1, "Mengunggah & Menganalisis Dokumen...");

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/upload", {
        method: "POST",
        body: formData
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.detail || "Gagal mengunggah file.");
      }

      currentJobId = data.job_id;
      displayAnalysis(data.filename, data.analysis);

    } catch (err) {
      alert("Error: " + err.message);
      resetToUpload();
    }
  }

  // Display analysis results
  function displayAnalysis(filename, analysis) {
    progressSection.style.display = "none";
    analysisSection.style.display = "block";

    document.getElementById("analysisDocName").textContent = filename;
    document.getElementById("analysisDocSize").textContent = `${analysis.file_size_mb} MB`;
    document.getElementById("analysisSummaryMsg").textContent = analysis.summary_message;

    document.getElementById("statImagesCount").textContent = analysis.image_count;
    document.getElementById("statMinRes").textContent = analysis.min_resolution;
    document.getElementById("statMaxRes").textContent = analysis.max_resolution;
    document.getElementById("statFontsCount").textContent = `${analysis.fonts.length} Font`;

    // Check font warnings
    const fontStatus = analysis.font_status;
    const warningBox = document.getElementById("fontWarningBox");
    const warningList = document.getElementById("fontWarningList");

    if (fontStatus && !fontStatus.all_available && fontStatus.missing_fonts.length > 0) {
      warningBox.style.display = "block";
      warningList.innerHTML = fontStatus.missing_fonts.map(m => 
        `<li>Font <strong>${m.font}</strong> tidak tersedia di server. Disarankan fallback: <em>${m.suggested_fallback}</em></li>`
      ).join("");
    } else {
      warningBox.style.display = "none";
    }

    // Complex elements summary
    const elems = analysis.complex_elements;
    const complexParts = [];
    if (elems.tables_count > 0) complexParts.push(`${elems.tables_count} Tabel`);
    if (elems.equations_count > 0) complexParts.push(`${elems.equations_count} Equation`);
    if (elems.textboxes_count > 0) complexParts.push(`${elems.textboxes_count} Text Box`);
    if (elems.shapes_count > 0) complexParts.push(`${elems.shapes_count} Shapes`);
    document.getElementById("statComplexElems").textContent = complexParts.length > 0 ? complexParts.join(", ") : "Tidak ada";
  }

  // Convert Button Click
  convertBtn.addEventListener("click", async () => {
    if (!currentJobId) return;

    analysisSection.style.display = "none";
    progressSection.style.display = "block";

    updateProgressStep(2, "Merender dengan Mesin Microsoft Word...");

    try {
      const response = await fetch(`/api/jobs/${currentJobId}/convert`, {
        method: "POST"
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.detail || "Gagal melakukan konversi.");
      }

      currentReportData = data;
      displayResults(data);

    } catch (err) {
      alert("Kesalahan konversi: " + err.message);
      resetToUpload();
    }
  });

  cancelBtn.addEventListener("click", () => {
    resetToUpload();
  });

  // Display Result and Quality Report
  function displayResults(data) {
    progressSection.style.display = "none";
    resultSection.style.display = "block";

    const v = data.verification;
    const badgeCard = document.getElementById("statusBadgeCard");
    const statusTitle = document.getElementById("statusTitle");
    const statusDesc = document.getElementById("statusDesc");

    badgeCard.className = `status-badge-card ${v.status_color}`;
    statusTitle.textContent = `Status: ${v.status_label}`;
    statusDesc.textContent = v.status_reasons.join(". ");

    // Setup PDF Preview
    const pdfIframe = document.getElementById("pdfIframe");
    pdfIframe.src = data.preview_url;

    // Download Button setup
    const downloadPdfBtn = document.getElementById("downloadPdfBtn");
    downloadPdfBtn.href = data.download_url;

    // Build Quality Table (Section 13 PRD)
    const tableBody = document.getElementById("qualityTableBody");
    tableBody.innerHTML = v.quality_table.map(row => {
      let statusTag = "tag-success";
      if (row.status.includes("Selisih") || row.status.includes("Peringatan")) {
        statusTag = "tag-warning";
      }
      return `
        <tr>
          <td><strong>${row.item}</strong></td>
          <td>${row.word_original}</td>
          <td>${row.pdf_result}</td>
          <td><span class="${statusTag}">${row.status}</span></td>
        </tr>
      `;
    }).join("");
  }

  // Progress update helper
  function updateProgressStep(stepNum, label) {
    document.getElementById("progressLabel").textContent = label;
    document.querySelectorAll(".step-item").forEach((el, idx) => {
      el.classList.remove("active", "done");
      if (idx + 1 < stepNum) el.classList.add("done");
      if (idx + 1 === stepNum) el.classList.add("active");
    });
  }

  // Download JSON Report
  downloadReportBtn.addEventListener("click", () => {
    if (!currentReportData) return;
    const blob = new Blob([JSON.stringify(currentReportData, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `puredoc_report_${currentJobId}.json`;
    a.click();
    URL.revokeObjectURL(url);
  });

  // Delete Job button (FR-20)
  deleteBtn.addEventListener("click", async () => {
    if (!currentJobId) return;
    if (!confirm("Hapus file dokumen Word dan PDF dari server sekarang demi privasi?")) return;

    try {
      await fetch(`/api/jobs/${currentJobId}`, { method: "DELETE" });
      alert("File dokumen dan PDF telah berhasil dihapus permanen dari server.");
      resetToUpload();
    } catch (err) {
      alert("Gagal menghapus file: " + err.message);
    }
  });

  newDocBtn.addEventListener("click", () => {
    resetToUpload();
  });

  function resetToUpload() {
    currentJobId = null;
    currentReportData = null;
    fileInput.value = "";
    document.getElementById("pdfIframe").src = "";
    uploadSection.style.display = "block";
    analysisSection.style.display = "none";
    progressSection.style.display = "none";
    resultSection.style.display = "none";
  }
});
