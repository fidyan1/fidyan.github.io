import JSZip from "jszip";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import mammoth from "mammoth";
import { renderAsync } from "docx-preview";
import html2pdf from "html2pdf.js";

// DOM Elements
const templateFileInput = document.getElementById("templateFile");
const namesFileInput = document.getElementById("namesFile");
const generateBtn = document.getElementById("generateBtn");
const fileList = document.getElementById("fileList");
const downloadAllBtn = document.getElementById("downloadAllBtn");
const countSpan = document.getElementById("count");
const btnText = generateBtn.querySelector(".btn-text");
const loader = generateBtn.querySelector(".loader");

const zipFileInput = document.getElementById("zipFile");
const convertBtn = document.getElementById("convertBtn");
const convertBtnText = convertBtn.querySelector(".btn-text");
const convertLoader = convertBtn.querySelector(".loader");

// State
let generatedFiles = [];

// Event Listeners
generateBtn.addEventListener("click", handleGenerate);
downloadAllBtn.addEventListener("click", handleDownloadAll);
convertBtn.addEventListener("click", handleConvertZip);

// File input UI handlers
[templateFileInput, namesFileInput, zipFileInput].forEach(input => {
  input.addEventListener("change", (e) => {
    const file = e.target.files[0];
    const wrapper = input.parentElement;
    const nameDisplay = wrapper.querySelector(".file-name-display");
    const uploadText = wrapper.querySelector(".upload-text");
    
    if (file) {
      nameDisplay.textContent = file.name;
      nameDisplay.classList.remove("hidden");
      uploadText.classList.add("hidden");
    } else {
      nameDisplay.classList.add("hidden");
      uploadText.classList.remove("hidden");
    }
  });
});

async function handleGenerate() {
  const templateFile = templateFileInput.files[0];
  const namesFile = namesFileInput.files[0];

  if (!templateFile || !namesFile) {
    alert("Mohon upload file template dan daftar nama (.docx).");
    return;
  }

  setLoading(true);
  generatedFiles = [];
  fileList.innerHTML = "";
  const downloadControls = document.getElementById("downloadControls");
  downloadControls.classList.add("hidden");
  downloadAllBtn.disabled = true;

  try {
    // 1. Read Files
    const templateBuffer = await readFileAsArrayBuffer(templateFile);
    const namesBuffer = await readFileAsArrayBuffer(namesFile);
    
    // 2. Extract Names
    const namesResult = await mammoth.extractRawText({ arrayBuffer: namesBuffer });
    const names = namesResult.value.split(/\r?\n/).map(n => n.trim()).filter(n => n);

    if (names.length === 0) {
      alert("Tidak ada nama yang ditemukan. Pastikan satu nama per baris.");
      setLoading(false);
      return;
    }

    // 3. Process Logic
    let processedCount = 0;
    
    for (const name of names) {
        // Update UI Progress
        processedCount++;
        btnText.textContent = `Memproses ${processedCount}/${names.length}...`;

        // A. PREPARE THE DOCX (XML Manipulation)
        let zip = new PizZip(templateBuffer);
        
        // XML Pre-processing for Robust Tag Support
        try {
            const xmlFile = "word/document.xml";
            if (zip.files[xmlFile]) {
                let xml = zip.file(xmlFile).asText();
                
                // Heal {{nama}}
                xml = xml.replace(/(\{\{)(?:<[^>]+>)*?(nama)(?:<[^>]+>)*?(\}\})/gi, "{{nama}}");
                
                // Support [nama]
                const bracketPattern = /(\[)(?:<[^>]+>)*?(nama)(?:<[^>]+>)*?(\])/gi;
                // Direct replacement if found
                xml = xml.replace(bracketPattern, name);

                zip.file(xmlFile, xml);
            }
        } catch (e) {
            console.warn("XML Preprocessing failed:", e);
        }

        // Templating (Standard)
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: () => ""
        });

        try {
            doc.render({ nama: name });
        } catch (e) {
            console.warn("Templating error (ignored due to pre-processing):", e);
        }

        // Generate Docx Blob
        const docxBlob = doc.getZip().generate({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        
        // Save DOCX format
        const filename = `${name}.docx`;
        generatedFiles.push({
            name: name, // Raw name for display
            filename: filename,
            blob: docxBlob
        });
    }

    renderFileList();
    countSpan.textContent = generatedFiles.length;
    if (generatedFiles.length > 0) {
      downloadControls.classList.remove("hidden");
      downloadAllBtn.disabled = false;
    }

  } catch (error) {
    console.error(error);
    alert("Terjadi kesalahan: " + error.message);
  } finally {
    setLoading(false);
    btnText.textContent = "Buat Surat";
  }
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
  });
}

function renderFileList() {
  if (generatedFiles.length === 0) return;

  fileList.innerHTML = generatedFiles.map((file, index) => `
      <div class="file-item">
        <span class="file-name">${file.filename}</span>
        <button class="secondary-btn small-btn" onclick="downloadFile(${index})">Unduh</button>
      </div>
  `).join("");
}

window.downloadFile = (index) => {
    const file = generatedFiles[index];
    if (!file) return;

    const url = URL.createObjectURL(file.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = file.filename;
    a.click();
    URL.revokeObjectURL(url);
};

// ZIP Download for Docx files
async function handleDownloadAll() {
    if (generatedFiles.length === 0) return;
    const zip = new JSZip();
    
    generatedFiles.forEach(file => {
        zip.file(file.filename, file.blob);
    });

    try {
        const content = await zip.generateAsync({ type: "blob" });
        const url = URL.createObjectURL(content);
        const a = document.createElement("a");
        a.href = url;
        
        const filenameInput = document.getElementById("zipFilename");
        let filename = filenameInput.value.trim();
        if (!filename) {
            filename = "Semua_Surat_Docx.zip";
        } else if (!filename.toLowerCase().endsWith(".zip")) {
            filename += ".zip";
        }
        
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
    } catch (e) {
        alert("Gagal membuat ZIP");
    }
}

function setLoading(isLoading) {
  generateBtn.disabled = isLoading;
  if (!isLoading) loader.classList.add("hidden");
  else loader.classList.remove("hidden");
}

function setConvertLoading(isLoading) {
    convertBtn.disabled = isLoading;
    if (!isLoading) convertLoader.classList.add("hidden");
    else convertLoader.classList.remove("hidden");
}

async function handleConvertZip() {
    const zipFile = zipFileInput.files[0];
    if (!zipFile) {
        alert("Mohon upload file ZIP terlebih dahulu.");
        return;
    }

    setConvertLoading(true);
    convertBtnText.textContent = "Mengekstrak ZIP...";

    // 1. Prepare Container (Hidden, high-fidelity)
    const container = document.createElement("div");
    container.style.position = "fixed";
    container.style.left = "0"; // Move on-screen
    container.style.top = "0";
    container.style.zIndex = "-1000"; // Hide behind background
    document.body.appendChild(container);

    try {
        // 2. Load ZIP
        const zipBuffer = await readFileAsArrayBuffer(zipFile);
        const zip = new JSZip();
        const loadedZip = await zip.loadAsync(zipBuffer);
        
        const docxFiles = Object.keys(loadedZip.files).filter(filename => 
            filename.endsWith(".docx") && !filename.startsWith("~") && !loadedZip.files[filename].dir
        );

        if (docxFiles.length === 0) {
            throw new Error("Tidak ada file .docx ditemukan dalam ZIP ini.");
        }

        const pdfZip = new JSZip();
        let processedCount = 0;
        let successCount = 0;
        let failCount = 0;

        for (const filename of docxFiles) {
            processedCount++;
            convertBtnText.textContent = `Mengonversi ${processedCount}/${docxFiles.length}...`;

            try {
                const docxContent = await loadedZip.file(filename).async("blob");
                
                // Clear previous render
                container.innerHTML = "";
                
                // Create Page Wrapper (Visual Box)
                const pageWrapper = document.createElement("div");
                pageWrapper.className = "docx-wrapper";
                // A4 Dimensions: 210mm x 297mm
                // At 96 DPI: ~794px x 1123px. We use slightly larger for better buffer.
                pageWrapper.style.width = "795px"; 
                pageWrapper.style.minHeight = "1123px";
                pageWrapper.style.backgroundColor = "white";
                pageWrapper.style.margin = "0";
                pageWrapper.style.padding = "0"; // Let docx-preview handle margins
                container.appendChild(pageWrapper);

                // Render DOCX -> DOM
                await renderAsync(docxContent, pageWrapper, null, {
                    className: "docx-content",
                    inWrapper: false, 
                    ignoreWidth: false, 
                    ignoreHeight: false,
                    experimental: true, 
                    useBase64URL: true,
                    breakPages: true // Important for pagination
                });

                const style = document.createElement("style");
                style.textContent = `
                    .docx-wrapper { background: white; }
                    /* Remove shadow to make it look like a flat paper for PDF */
                    .docx-wrapper section, .docx-wrapper article { 
                        box-shadow: none !important; 
                        margin-bottom: 0 !important;
                    }
                    
                    /* FIXED: Do NOT constrain images. Let them use their natural/DOCX size. */
                    .docx-wrapper img { 
                        max-width: none !important; 
                        height: auto !important; 
                        /* Blend mode for signatures to look "stamped" */
                        mix-blend-mode: multiply; 
                    }
                `;
                pageWrapper.appendChild(style);

                // Wait for images/fonts to settle - INCREASED for heavy images
                await new Promise(resolve => setTimeout(resolve, 1500));

                // DOM -> PDF
                const pdfBlob = await html2pdf().set({
                    margin: 0, // We control margins in CSS
                    filename: 'document.pdf',
                    image: { type: 'jpeg', quality: 1.0 }, // Max quality
                    html2canvas: { 
                        scale: 4, // 4x Scale for high DPI (approx 300+ DPI)
                        useCORS: true,
                        letterRendering: true,
                        scrollY: 0,
                        windowWidth: 795,
                        dpi: 300 // Explicit DPI hint
                    },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
                }).from(pageWrapper).output("blob");

                const pdfFilename = filename.replace(/\.docx$/i, ".pdf");
                pdfZip.file(pdfFilename, pdfBlob);
                successCount++;

            } catch (err) {
                console.error(`Gagal mengonversi ${filename}:`, err);
                failCount++;
            }
        }

        if (successCount === 0) {
            throw new Error("Gagal mengonversi semua file.");
        }

        // 3. Generate Final ZIP
        convertBtnText.textContent = "Mengemas ZIP...";
        const content = await pdfZip.generateAsync({ type: "blob" });
        
        // 4. Download
        const url = URL.createObjectURL(content);
        const a = document.createElement("a");
        a.href = url;
        a.download = `Hasil_Konversi_PDF_${new Date().getTime()}.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        alert(`Selesai!\nBerhasil: ${successCount}\nGagal: ${failCount}`);

    } catch (error) {
        console.error(error);
        alert("Terjadi kesalahan: " + error.message);
    } finally {
        if (document.body.contains(container)) {
            document.body.removeChild(container);
        }
        setConvertLoading(false);
        convertBtnText.textContent = "Konversi ke PDF";
    }
}
