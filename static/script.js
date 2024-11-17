// Fungsi untuk memicu pemilihan file umum
function triggerFileSelection() {
    console.log("Trigger file selection");

    // Mencari elemen input untuk PDF, Word, Excel, atau PPT
    const fileInput = document.getElementById('pdfFileInput') || document.getElementById('wordFileInput') || document.getElementById('excelFileInput') || document.getElementById('pptFileInput');
    if (fileInput) {
        fileInput.click();
        console.log("File input element found and clicked:", fileInput.id);
    } else {
        console.error("No file input element found.");
    }
}

// Event listener untuk PDF to Word
document.getElementById('pdfFileInput')?.addEventListener('change', () => {
    console.log("PDF file selected.");
    handleFileSelection('pdfFileInput', '/upload');
});

// Event listener untuk Word to PDF
document.getElementById('wordFileInput')?.addEventListener('change', () => {
    console.log("Word file selected.");
    handleFileSelection('wordFileInput', '/upload_word_to_pdf');
});

// Event listener untuk Excel to PDF
document.getElementById('excelFileInput')?.addEventListener('change', () => {
    console.log("Excel file selected.");
    handleFileSelection('excelFileInput', '/upload_excel_to_pdf');
});

// Event listener untuk PPT to PDF
document.getElementById('pptFileInput')?.addEventListener('change', () => {
    console.log("PPT file selected.");
    handleFileSelection('pptFileInput', '/upload_ppt_to_pdf'); // URL endpoint khusus PPT to PDF
});

// Fungsi untuk menangani unggahan file dan mengirimkannya ke server
function handleFileSelection(inputId, uploadUrl) {
    const fileInput = document.getElementById(inputId);
    if (fileInput && fileInput.files.length > 0) {
        const file = fileInput.files[0];
        console.log("File chosen:", file.name);

        const formData = new FormData();
        
        // Menentukan jenis file berdasarkan input ID
        if (inputId === 'pdfFileInput') {
            formData.append('pdf_file', file);
        } else if (inputId === 'wordFileInput') {
            formData.append('word_file', file);
        } else if (inputId === 'excelFileInput') {
            formData.append('excel_file', file);
        } else if (inputId === 'pptFileInput') {
            formData.append('ppt_file', file);
        }

        console.log("Sending request to URL:", uploadUrl);
        
        fetch(uploadUrl, {
            method: 'POST',
            body: formData,
        })
        .then(response => {
            console.log("Response received:", response);
            if (!response.ok) {
                throw new Error("Network response was not ok");
            }
            return response.json();
        })
        .then(data => {
            console.log("Data received:", data);
            const downloadContainer = document.getElementById("downloadLinkContainer");
            downloadContainer.innerHTML = "";  // Bersihkan tautan unduhan sebelumnya jika ada

            if (data.success) {
                alert("File converted successfully");

                // Buat tautan unduhan
                const downloadLink = document.createElement("a");
                downloadLink.href = data.download_url;
                downloadLink.innerText = "Download Converted File";
                downloadLink.style.display = "block";
                downloadLink.style.marginTop = "20px";
                downloadLink.style.padding = "10px";
                downloadLink.style.backgroundColor = "#4CAF50";
                downloadLink.style.color = "white";
                downloadLink.style.textAlign = "center";
                downloadLink.style.textDecoration = "none";
                downloadLink.style.borderRadius = "5px";

                downloadContainer.appendChild(downloadLink);
            } else {
                alert(`Error: ${data.message}`);
            }
        })
        .catch(error => {
            console.error("Error during file upload:", error);
            alert("An error occurred while uploading the file.");
        });
    } else {
        console.error("No file selected for upload.");
        alert("Please select a file to upload.");
    }
}

// Fungsi khusus untuk memicu pemilihan file untuk Excel to PDF
function triggerFileSelectionExcelToPDF() {
    const fileInput = document.getElementById('excelFileInput');
    if (fileInput) {
        fileInput.click();
        console.log("Excel file input clicked.");
    } else {
        console.error("Excel file input not found.");
    }
}

// Fungsi khusus untuk memicu pemilihan file untuk PPT to PDF
function triggerFileSelectionPPT() {
    const fileInput = document.getElementById('pptFileInput');
    if (fileInput) {
        fileInput.click();
        console.log("pptFileInput ditemukan dan click() dipanggil");
    } else {
        console.error("pptFileInput tidak ditemukan");
    }
}
