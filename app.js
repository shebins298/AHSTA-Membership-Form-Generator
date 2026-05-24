// ===============================
// AHSTA Membership Generator
// app.js
// ===============================

// ---------- GLOBAL STATE ----------
let membershipData = [];
let groupedSchools = {};

// ---------- DOM ELEMENTS ----------
const excelFileInput = document.getElementById("excelFileInput");

const uploadButton = document.getElementById("uploadButton");

const selectedFileName = document.getElementById("selectedFileName");

const totalSchools = document.getElementById("totalSchools");

const totalMembers = document.getElementById("totalMembers");

const schoolTableBody = document.getElementById("schoolTableBody");

const schoolSearchInput = document.getElementById("schoolSearchInput");

const districtSelect = document.getElementById("districtSelect");

const downloadAllPdfButton = document.getElementById("downloadAllPdfButton");

const downloadAllDocxButton = document.getElementById("downloadAllDocxButton");

const previewModal = document.getElementById("previewModal");

const previewContainer = document.getElementById("previewContainer");

const closePreviewButton = document.getElementById("closePreviewButton");

// ===============================
// EVENT LISTENERS
// ===============================
uploadButton.addEventListener("click", () => {
  excelFileInput.click();
});

excelFileInput.addEventListener("change", handleExcelUpload);

schoolSearchInput.addEventListener("input", handleSchoolSearch);

closePreviewButton.addEventListener("click", closePreviewModal);

downloadAllPdfButton.addEventListener("click", downloadAllSchoolsAsPdf);

downloadAllDocxButton.addEventListener("click", downloadAllSchoolsAsDocx);

// ===============================
// HANDLE EXCEL UPLOAD
// ===============================
function handleExcelUpload(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  selectedFileName.textContent = file.name;

  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);

    const workbook = XLSX.read(data, {
      type: "array",
    });

    const sheetName = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[sheetName];

    membershipData = XLSX.utils.sheet_to_json(worksheet, {
      defval: "",
    });

    processMembershipData();
  };

  reader.readAsArrayBuffer(file);
}

// ===============================
// PROCESS MEMBERSHIP DATA
// ===============================
function processMembershipData() {
  groupedSchools = {};

  membershipData.forEach((row) => {
    const schoolName = getValue(row, ["School", "School Name", "SCHOOL"]);

    if (!schoolName) {
      return;
    }

    if (!groupedSchools[schoolName]) {
      groupedSchools[schoolName] = [];
    }

    groupedSchools[schoolName].push({
      teacherName: getValue(row, ["Name", "Teacher Name", "NAME"]),

      subject: getValue(row, ["Subject", "SUBJECT"]),

      phone: getValue(row, ["Phone", "Mobile", "WhatsApp", "WHATSAPP NO"]),

      district: getValue(row, ["District", "DISTRICT"]),
    });
  });

  updateSummary();

  renderSchoolTable();

  enableBulkButtons();
}

// ===============================
// SAFE VALUE GETTER
// ===============================
function getValue(row, possibleKeys) {
  for (const key of possibleKeys) {
    if (row[key] !== undefined) {
      return String(row[key]).trim();
    }
  }

  return "";
}

// ===============================
// UPDATE SUMMARY
// ===============================
function updateSummary() {
  totalSchools.textContent = Object.keys(groupedSchools).length;

  totalMembers.textContent = membershipData.length;
}

// ===============================
// RENDER SCHOOL TABLE
// ===============================
function renderSchoolTable(filteredData = groupedSchools) {
  schoolTableBody.innerHTML = "";

  const schools = Object.keys(filteredData);

  if (schools.length === 0) {
    schoolTableBody.innerHTML = `
            <tr>
                <td colspan="4" class="empty-text">
                    No schools found
                </td>
            </tr>
        `;

    return;
  }

  schools.forEach((schoolName, index) => {
    const members = filteredData[schoolName];

    const row = document.createElement("tr");

    row.innerHTML = `
                <td>${index + 1}</td>

                <td class="school-name-cell">
                    ${schoolName}
                </td>

                <td>
                    ${members.length}
                </td>

                <td class="action-cell">

                    <button
                        class="table-btn preview-btn"
                        data-school="${schoolName}"
                    >
                        Preview
                    </button>

                    <button
                        class="table-btn pdf-btn"
                        data-school="${schoolName}"
                    >
                        PDF
                    </button>

                    <button
                        class="table-btn docx-btn"
                        data-school="${schoolName}"
                        style="display:none;"
                    >
                        DOCX
                    </button>

                </td>
            `;

    schoolTableBody.appendChild(row);
  });

  attachTableButtonEvents();
}

// ===============================
// ATTACH BUTTON EVENTS
// ===============================
function attachTableButtonEvents() {
  document.querySelectorAll(".preview-btn").forEach((button) => {
    button.addEventListener("click", () => {
      showPreview(button.dataset.school);
    });
  });

  document.querySelectorAll(".pdf-btn").forEach((button) => {
    button.addEventListener("click", async () => {
      await generatePdf(button.dataset.school);
    });
  });

  document.querySelectorAll(".docx-btn").forEach((button) => {
    button.addEventListener("click", async () => {
      await generateDocx(button.dataset.school);
    });
  });
}

// ===============================
// SEARCH
// ===============================
function handleSchoolSearch(event) {
  const searchText = event.target.value.toLowerCase().trim();

  const filteredSchools = {};

  Object.keys(groupedSchools).forEach((schoolName) => {
    if (schoolName.toLowerCase().includes(searchText)) {
      filteredSchools[schoolName] = groupedSchools[schoolName];
    }
  });

  renderSchoolTable(filteredSchools);
}

// ===============================
// SHOW PREVIEW
// ===============================
function showPreview(schoolName) {
  previewContainer.innerHTML = generateMembershipTemplate(schoolName);

  previewModal.classList.remove("hidden");
}

// ===============================
// CLOSE PREVIEW
// ===============================
function closePreviewModal() {
  previewModal.classList.add("hidden");
}

// ===============================
// GENERATE TEMPLATE
// ===============================
function generateMembershipTemplate(schoolName) {
  return generateMembershipPages(schoolName).join("");
}

// ===============================
// GENERATE PAGES
// ===============================
function generateMembershipPages(schoolName) {
  const members = groupedSchools[schoolName];

  // const district = members[0]?.district || "ERNAKULAM";

  //const district = "";

  const district = districtSelect.value || "";

  const rowsPerPage = 18;

  const pages = chunkArray(members, rowsPerPage);

  return pages.map((pageMembers, pageIndex) => {
    let rowsHtml = "";

    for (let i = 0; i < rowsPerPage; i++) {
      const member = pageMembers[i];

      const actualIndex = pageIndex * rowsPerPage + i + 1;

      rowsHtml += `

                    <tr>

                        <td class="cell sl-cell">
                            ${member ? actualIndex : ""}
                        </td>

                        <td class="cell name-cell">
                            ${member ? member.teacherName : ""}
                        </td>

                        <td class="cell subject-cell">
                            ${member ? member.subject : ""}
                        </td>

                        <td class="cell mobile-cell">
                            ${member ? member.phone : ""}
                        </td>

                        <td class="cell signature-cell"></td>

                    </tr>
                `;
    }

    const isLastPage = pageIndex === pages.length - 1;

    return `

                <div class="pdf-page">

                    <div class="header-image-container">

                        <img
                            src="images/ahsta-header.png"
                            class="header-image"
                        />

                        <div class="membership-title">

                        MEMBERSHIP LIST 2026-27
                
                    </div>

                    </div>

                    <div class="school-info">

                        <div class="school-row">

                            <span class="label">
                                Name of District :
                            </span>

                            <span class="value">
                            ${district}
                            </span>

                        </div>

                        <div class="school-row">

                            <span class="label">
                                Name of School :
                            </span>

                            <span class="value">
                                ${schoolName}
                            </span>

                        </div>

                    </div>

                    <table class="pdf-table">

                        <thead>

                            <tr>

                                <th class="cell sl-cell">
                                    Sl No
                                </th>

                                <th class="cell name-cell">
                                    Name
                                </th>

                                <th class="cell subject-cell">
                                    Subject
                                </th>

                                <th class="cell mobile-cell">
                                    Mobile Number
                                </th>

                                <th class="cell signature-cell">
                                    Signature
                                </th>

                            </tr>

                        </thead>

                        <tbody>

                            ${rowsHtml}

                        </tbody>

                    </table>

                    <div class="after-table-space"></div>

                    ${
                      isLastPage
                        ? `

                            <div class="signature-row">

                                <div>
                                    Sign. of Unit Secretary
                                </div>

                                <div>
                                    Sign. of Dist. Secretary
                                </div>

                                <div>
                                    Sign. of General Secretary
                                </div>

                            </div>

                        `
                        : ""
                    }

                </div>
            `;
  });
}

// ===============================
// CHUNK ARRAY
// ===============================
function chunkArray(array, size) {
  const chunks = [];

  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }

  return chunks;
}

// ===============================
// GENERATE PDF
// ===============================
async function generatePdf(schoolName) {
  const pages = generateMembershipPages(schoolName);

  const pdf = new jspdf.jsPDF({
    orientation: "portrait",

    unit: "mm",

    format: "a4",
  });

  for (let i = 0; i < pages.length; i++) {
    const container = document.createElement("div");

    container.style.width = "210mm";

    container.style.background = "white";

    container.innerHTML = pages[i];

    document.body.appendChild(container);

    await new Promise((resolve) => setTimeout(resolve, 300));

    const canvas = await html2canvas(container.firstElementChild, {
      scale: 2,

      useCORS: true,
    });

    const imgData = canvas.toDataURL("image/jpeg", 1);

    if (i > 0) {
      pdf.addPage();
    }

    pdf.addImage(
      imgData,

      "JPEG",

      0,

      0,

      210,

      297
    );

    document.body.removeChild(container);
  }

  pdf.save(`${sanitizeFileName(schoolName)}.pdf`);
}

// ===============================
// BULK PDF
// ===============================
async function downloadAllSchoolsAsPdf() {
  const schoolNames = Object.keys(groupedSchools);

  for (const schoolName of schoolNames) {
    await generatePdf(schoolName);

    await new Promise((resolve) => setTimeout(resolve, 800));
  }
}

// ===============================
// GENERATE DOCX
// ===============================
async function generateDocx(schoolName) {
  const html = generateMembershipTemplate(schoolName);

  const convertedHtml = `

        <html>

            <head>

                <meta charset="utf-8">

            </head>

            <body>

                ${html}

            </body>

        </html>
    `;

  const blob = window.htmlDocx.asBlob(convertedHtml);

  saveAs(
    blob,

    `${sanitizeFileName(schoolName)}.docx`
  );
}

// ===============================
// BULK DOCX
// ===============================
async function downloadAllSchoolsAsDocx() {
  const schoolNames = Object.keys(groupedSchools);

  for (const schoolName of schoolNames) {
    await generateDocx(schoolName);

    await new Promise((resolve) => setTimeout(resolve, 800));
  }
}

// ===============================
// ENABLE BUTTONS
// ===============================
function enableBulkButtons() {
  downloadAllPdfButton.disabled = false;

  downloadAllDocxButton.disabled = false;
}

// ===============================
// SANITIZE FILE NAME
// ===============================
function sanitizeFileName(fileName) {
  return fileName.replace(/[\\/:*?"<>|]/g, "_");
}
