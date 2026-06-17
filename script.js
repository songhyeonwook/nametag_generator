const state = {
  rows: [],
  columns: [],
  templateFile: null,
  templateSize: null,
};

const EMU_PER_CM = 360000;
const A4 = {
  landscape: { width: Math.round(29.7 * EMU_PER_CM), height: Math.round(21.0 * EMU_PER_CM) },
  portrait: { width: Math.round(21.0 * EMU_PER_CM), height: Math.round(29.7 * EMU_PER_CM) },
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  templateInput: document.querySelector("#templateInput"),
  fileName: document.querySelector("#fileName"),
  templateName: document.querySelector("#templateName"),
  companySelect: document.querySelector("#companySelect"),
  departmentSelect: document.querySelector("#departmentSelect"),
  nameSelect: document.querySelector("#nameSelect"),
  orientationSelect: document.querySelector("#orientationSelect"),
  gridSelect: document.querySelector("#gridSelect"),
  copyInput: document.querySelector("#copyInput"),
  generateButton: document.querySelector("#generateButton"),
  sampleButton: document.querySelector("#sampleButton"),
  summaryText: document.querySelector("#summaryText"),
  countBadge: document.querySelector("#countBadge"),
  pagePreview: document.querySelector("#pagePreview"),
  previewTable: document.querySelector("#previewTable"),
};

function normalize(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function hasSpreadsheetLibrary() {
  return typeof window.XLSX !== "undefined";
}

function hasZipLibrary() {
  return typeof window.JSZip !== "undefined";
}

function setLibraryWarningIfNeeded() {
  if (hasSpreadsheetLibrary() && hasZipLibrary()) return;
  els.summaryText.textContent =
    "외부 라이브러리를 불러오지 못했습니다. 인터넷 연결 또는 CDN 차단 여부를 확인하세요.";
}

function escapeHtml(value) {
  return normalize(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeXml(value) {
  return normalize(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function findColumnIndex(columns, candidates, fallback) {
  const lowered = columns.map((column) => normalize(column).toLowerCase());
  for (const candidate of candidates) {
    const index = lowered.indexOf(candidate.toLowerCase());
    if (index >= 0) return index;
  }
  return Math.min(fallback, Math.max(columns.length - 1, 0));
}

function setSelectOptions(select, columns, selectedIndex) {
  select.innerHTML = "";
  columns.forEach((column, index) => {
    const option = document.createElement("option");
    option.value = column;
    option.textContent = column;
    option.selected = index === selectedIndex;
    select.appendChild(option);
  });
}

function getAttendees() {
  if (!state.rows.length) return [];

  const companyKey = els.companySelect.value;
  const departmentKey = els.departmentSelect.value;
  const nameKey = els.nameSelect.value;
  const copies = Math.max(1, Math.min(10, Number(els.copyInput.value) || 1));
  const attendees = [];

  for (const row of state.rows) {
    const company = normalize(row[companyKey]);
    const department = normalize(row[departmentKey]);
    const name = normalize(row[nameKey]);
    if (company && department && name) {
      for (let copy = 0; copy < copies; copy += 1) {
        attendees.push({ company, department, name, dept_name: `${department} ${name}` });
      }
    }
  }

  return attendees;
}

function parseGrid() {
  const [cols, rows] = els.gridSelect.value.split("x").map((value) => Number(value));
  return { cols, rows };
}

function getOutputLayout(templateSize) {
  const page = A4[els.orientationSelect.value];
  const { cols, rows } = parseGrid();
  const cellWidth = page.width / cols;
  const cellHeight = page.height / rows;
  const scale = Math.min(cellWidth / templateSize.width, cellHeight / templateSize.height);
  const tagWidth = Math.round(templateSize.width * scale);
  const tagHeight = Math.round(templateSize.height * scale);

  return {
    page,
    cols,
    rows,
    cellWidth,
    cellHeight,
    scale,
    tagWidth,
    tagHeight,
    perPage: cols * rows,
  };
}

function updatePreview() {
  const attendees = getAttendees();
  const hasTemplate = Boolean(state.templateFile && state.templateSize);
  const { cols, rows } = parseGrid();
  const perPage = cols * rows;
  const pageCount = attendees.length ? Math.ceil(attendees.length / perPage) : 0;

  els.generateButton.disabled = attendees.length === 0 || !hasTemplate;
  els.countBadge.textContent = `${attendees.length}명`;

  if (!state.rows.length && !hasTemplate) {
    els.summaryText.textContent = "명단과 PPTX 템플릿을 업로드하면 생성 정보가 표시됩니다.";
  } else if (!state.rows.length) {
    els.summaryText.textContent = "명단 Excel/CSV 파일을 업로드하세요.";
  } else if (!hasTemplate) {
    els.summaryText.textContent = "PPTX 템플릿을 업로드하세요. 첫 번째 슬라이드를 A4에 다중 배치합니다.";
  } else {
    els.summaryText.textContent = `총 ${attendees.length}개 명찰을 ${cols}x${rows} 배열로 배치하여 ${pageCount}페이지를 생성합니다.`;
  }

  renderPagePreview(attendees);
  renderTable(attendees);
}

function renderEmptyPagePreview(title, message) {
  els.pagePreview.className = "template-preview";
  els.pagePreview.innerHTML = `
    <strong>${escapeHtml(title)}</strong>
    <span>${escapeHtml(message)}</span>
  `;
}

function toPercent(value, total) {
  return `${((value / total) * 100).toFixed(4)}%`;
}

function renderPagePreview(attendees) {
  if (!state.rows.length && !state.templateFile) {
    renderEmptyPagePreview(
      "A4 다중 명찰 배치",
      "명단과 PPTX 템플릿을 업로드하면 첫 페이지 배치가 표시됩니다.",
    );
    return;
  }

  if (!state.rows.length) {
    renderEmptyPagePreview("명단 대기", "Excel 또는 CSV 명단을 업로드하면 미리보기가 표시됩니다.");
    return;
  }

  if (!state.templateFile || !state.templateSize) {
    renderEmptyPagePreview("템플릿 대기", "PPTX 템플릿을 업로드하면 A4 배치 미리보기가 표시됩니다.");
    return;
  }

  if (!attendees.length) {
    renderEmptyPagePreview("표시할 데이터 없음", "회사명, 부서, 이름이 모두 채워진 행을 찾지 못했습니다.");
    return;
  }

  const layout = getOutputLayout(state.templateSize);
  const pageAttendees = attendees.slice(0, layout.perPage);
  els.pagePreview.className = `page-preview ${els.orientationSelect.value}`;
  els.pagePreview.innerHTML = "";

  pageAttendees.forEach((attendee, index) => {
    const row = Math.floor(index / layout.cols);
    const col = index % layout.cols;
    const left = Math.round(col * layout.cellWidth + (layout.cellWidth - layout.tagWidth) / 2);
    const top = Math.round(row * layout.cellHeight + (layout.cellHeight - layout.tagHeight) / 2);

    const tag = document.createElement("div");
    tag.className = "preview-tag";
    tag.style.left = toPercent(left, layout.page.width);
    tag.style.top = toPercent(top, layout.page.height);
    tag.style.width = toPercent(layout.tagWidth, layout.page.width);
    tag.style.height = toPercent(layout.tagHeight, layout.page.height);
    tag.innerHTML = `
      <strong>${escapeHtml(attendee.company)}</strong>
      <span>${escapeHtml(attendee.dept_name)}</span>
    `;
    els.pagePreview.appendChild(tag);
  });

  const info = document.createElement("div");
  info.className = "preview-page-info";
  info.textContent = `1페이지 / ${layout.cols}x${layout.rows}`;
  els.pagePreview.appendChild(info);
}

function renderTable(attendees) {
  els.previewTable.innerHTML = "";
  if (!attendees.length) {
    els.previewTable.innerHTML = '<tr><td colspan="3">표시할 데이터가 없습니다.</td></tr>';
    return;
  }

  attendees.slice(0, 20).forEach((attendee) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(attendee.company)}</td>
      <td>${escapeHtml(attendee.department)}</td>
      <td>${escapeHtml(attendee.name)}</td>
    `;
    els.previewTable.appendChild(row);
  });
}

async function handleFileUpload(event) {
  if (!hasSpreadsheetLibrary()) {
    alert("Excel/CSV 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결 또는 CDN 차단 여부를 확인하세요.");
    event.target.value = "";
    return;
  }

  const file = event.target.files?.[0];
  if (!file) return;

  try {
    els.fileName.textContent = file.name;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    state.rows = rows;
    state.columns = rows.length ? Object.keys(rows[0]) : [];

    const companyIndex = findColumnIndex(state.columns, ["회사명", "company"], 0);
    const departmentIndex = findColumnIndex(state.columns, ["부서", "department", "dept"], 1);
    const nameIndex = findColumnIndex(state.columns, ["이름", "name"], 2);

    setSelectOptions(els.companySelect, state.columns, companyIndex);
    setSelectOptions(els.departmentSelect, state.columns, departmentIndex);
    setSelectOptions(els.nameSelect, state.columns, nameIndex);
    updatePreview();
  } catch (error) {
    console.error(error);
    alert("파일을 읽지 못했습니다. Excel 또는 CSV 형식인지 확인하세요.");
  }
}

async function handleTemplateUpload(event) {
  if (!hasZipLibrary()) {
    alert("PPTX 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결 또는 CDN 차단 여부를 확인하세요.");
    event.target.value = "";
    return;
  }

  const file = event.target.files?.[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    alert("PPTX 템플릿 파일을 선택하세요.");
    event.target.value = "";
    return;
  }

  try {
    const zip = await window.JSZip.loadAsync(await file.arrayBuffer());
    const presentationXmlFile = zip.file("ppt/presentation.xml");
    if (!presentationXmlFile) {
      throw new Error("PPTX 구조를 읽지 못했습니다.");
    }

    const presentationXml = await presentationXmlFile.async("string");
    state.templateFile = file;
    state.templateSize = getPresentationSize(presentationXml);
    els.templateName.textContent = file.name;
    updatePreview();
  } catch (error) {
    console.error(error);
    state.templateFile = null;
    state.templateSize = null;
    els.templateName.textContent = "첫 번째 슬라이드를 명찰 템플릿으로 사용합니다.";
    event.target.value = "";
    alert(`템플릿을 읽지 못했습니다.\n\n${error.message || error}`);
    updatePreview();
  }
}

function replacePlaceholders(slideXml, attendee) {
  const replacements = {
    "{{company}}": attendee.company,
    "{{department}}": attendee.department,
    "{{name}}": attendee.name,
    "{{dept_name}}": attendee.dept_name,
  };

  let output = slideXml;
  for (const [placeholder, value] of Object.entries(replacements)) {
    output = output.replaceAll(placeholder, escapeXml(value));
  }
  return output;
}

function parseXml(xml) {
  return new DOMParser().parseFromString(xml, "application/xml");
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function getPresentationSize(presentationXml) {
  const doc = parseXml(presentationXml);
  const sizeNode = doc.getElementsByTagName("p:sldSz")[0];
  if (!sizeNode) {
    throw new Error("템플릿의 슬라이드 크기 정보를 찾지 못했습니다.");
  }
  return {
    width: Number(sizeNode.getAttribute("cx")),
    height: Number(sizeNode.getAttribute("cy")),
  };
}

function setPresentationSize(presentationXml, page) {
  const doc = parseXml(presentationXml);
  const sizeNode = doc.getElementsByTagName("p:sldSz")[0];
  if (!sizeNode) {
    throw new Error("템플릿의 슬라이드 크기 정보를 찾지 못했습니다.");
  }
  sizeNode.setAttribute("cx", String(page.width));
  sizeNode.setAttribute("cy", String(page.height));
  sizeNode.setAttribute("type", "A4");
  return serializeXml(doc);
}

function getSpTreeChildren(slideDoc) {
  const spTree = slideDoc.getElementsByTagName("p:spTree")[0];
  if (!spTree) throw new Error("템플릿 슬라이드의 도형 정보를 찾지 못했습니다.");
  return Array.from(spTree.childNodes).filter((node) => {
    if (node.nodeType !== Node.ELEMENT_NODE) return false;
    return !["p:nvGrpSpPr", "p:grpSpPr"].includes(node.nodeName);
  });
}

function getTemplateCommonXml(slideDoc) {
  const cSld = slideDoc.getElementsByTagName("p:cSld")[0];
  if (!cSld) throw new Error("템플릿 슬라이드 구조를 읽지 못했습니다.");
  const bgNode = Array.from(cSld.childNodes).find((node) => node.nodeName === "p:bg");
  const clrMapOvr = slideDoc.getElementsByTagName("p:clrMapOvr")[0];
  return {
    backgroundXml: bgNode ? serializeXml(bgNode) : "",
    clrMapXml: clrMapOvr ? serializeXml(clrMapOvr) : '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>',
  };
}

function scaledNumber(value, scale, offset = 0) {
  return String(Math.round(Number(value) * scale + offset));
}

function scaleElementGeometry(element, scale, dx, dy) {
  element.querySelectorAll("a\\:off, off").forEach((node, index) => {
    const offsetX = index === 0 ? dx : 0;
    const offsetY = index === 0 ? dy : 0;
    if (node.hasAttribute("x")) node.setAttribute("x", scaledNumber(node.getAttribute("x"), scale, offsetX));
    if (node.hasAttribute("y")) node.setAttribute("y", scaledNumber(node.getAttribute("y"), scale, offsetY));
  });

  element.querySelectorAll("a\\:ext, ext").forEach((node) => {
    if (node.hasAttribute("cx")) node.setAttribute("cx", scaledNumber(node.getAttribute("cx"), scale));
    if (node.hasAttribute("cy")) node.setAttribute("cy", scaledNumber(node.getAttribute("cy"), scale));
  });

  element.querySelectorAll("[sz]").forEach((node) => {
    const size = Number(node.getAttribute("sz"));
    if (Number.isFinite(size) && size > 0) {
      node.setAttribute("sz", String(Math.max(100, Math.round(size * scale))));
    }
  });

  element.querySelectorAll("a\\:ln, ln").forEach((node) => {
    const width = Number(node.getAttribute("w"));
    if (Number.isFinite(width) && width > 0) {
      node.setAttribute("w", String(Math.max(1270, Math.round(width * scale))));
    }
  });
}

function renumberShapeIds(element, idAllocator) {
  element.querySelectorAll("p\\:cNvPr, cNvPr").forEach((node) => {
    node.setAttribute("id", String(idAllocator.next()));
  });
}

function createIdAllocator() {
  let current = 2;
  return {
    next() {
      current += 1;
      return current;
    },
  };
}

function buildSlideXmlFromTemplate(templateSlideXml, attendees, layout) {
  const sourceDoc = parseXml(templateSlideXml);
  const outputDoc = document.implementation.createDocument(
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p:sld",
  );
  const root = outputDoc.documentElement;
  root.setAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
  root.setAttribute("xmlns:p", "http://schemas.openxmlformats.org/presentationml/2006/main");
  root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

  const common = getTemplateCommonXml(sourceDoc);
  const cSld = outputDoc.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:cSld");
  if (common.backgroundXml) {
    cSld.appendChild(outputDoc.importNode(parseXml(common.backgroundXml).documentElement, true));
  }

  const spTree = outputDoc.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:spTree");
  const spTreeBase = parseXml(
    '<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>',
  );
  Array.from(spTreeBase.documentElement.childNodes).forEach((node) => {
    spTree.appendChild(outputDoc.importNode(node, true));
  });

  const templateShapes = getSpTreeChildren(sourceDoc);
  const idAllocator = createIdAllocator();

  attendees.forEach((attendee, index) => {
    const row = Math.floor(index / layout.cols);
    const col = index % layout.cols;
    const dx = Math.round(col * layout.cellWidth + (layout.cellWidth - layout.tagWidth) / 2);
    const dy = Math.round(row * layout.cellHeight + (layout.cellHeight - layout.tagHeight) / 2);

    templateShapes.forEach((shape) => {
      const replacedXml = replacePlaceholders(serializeXml(shape), attendee);
      const shapeDoc = parseXml(replacedXml);
      const copiedShape = outputDoc.importNode(shapeDoc.documentElement, true);
      scaleElementGeometry(copiedShape, layout.scale, dx, dy);
      renumberShapeIds(copiedShape, idAllocator);
      spTree.appendChild(copiedShape);
    });
  });

  cSld.appendChild(spTree);
  root.appendChild(cSld);
  root.appendChild(outputDoc.importNode(parseXml(common.clrMapXml).documentElement, true));
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + serializeXml(outputDoc);
}

function removeExistingSlides(zip) {
  Object.keys(zip.files).forEach((path) => {
    if (/^ppt\/slides\/slide\d+\.xml$/.test(path) || /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(path)) {
      zip.remove(path);
    }
  });
}

function buildPresentationRelationshipsXml(existingXml, slideCount) {
  const relsWithoutSlides = existingXml.replace(
    /\s*<Relationship\b[^>]*Type="[^"]*\/slide"[^>]*Target="slides\/slide\d+\.xml"[^>]*\/>/g,
    "",
  );
  const slideRels = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<Relationship Id="rIdGeneratedSlide${slideNumber}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNumber}.xml"/>`;
  }).join("");

  return relsWithoutSlides.replace("</Relationships>", `${slideRels}</Relationships>`);
}

function buildPresentationXml(existingXml, slideCount) {
  const slideIds = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<p:sldId id="${256 + index}" r:id="rIdGeneratedSlide${slideNumber}"/>`;
  }).join("");

  if (/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.test(existingXml)) {
    return existingXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${slideIds}</p:sldIdLst>`);
  }

  return existingXml.replace("</p:presentation>", `<p:sldIdLst>${slideIds}</p:sldIdLst></p:presentation>`);
}

function buildContentTypesXml(existingXml, slideCount) {
  const withoutSlideOverrides = existingXml.replace(
    /\s*<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*ContentType="application\/vnd\.openxmlformats-officedocument\.presentationml\.slide\+xml"[^>]*\/>/g,
    "",
  );
  const slideOverrides = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<Override PartName="/ppt/slides/slide${slideNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }).join("");

  return withoutSlideOverrides.replace("</Types>", `${slideOverrides}</Types>`);
}

function updateAppPropertiesXml(existingXml, slideCount) {
  if (!existingXml) return existingXml;
  if (/<Slides>\d+<\/Slides>/.test(existingXml)) {
    return existingXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${slideCount}</Slides>`);
  }
  return existingXml;
}

async function generateFromTemplate() {
  if (!hasZipLibrary()) {
    alert("PPTX 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결 또는 CDN 차단 여부를 확인하세요.");
    return;
  }

  const attendees = getAttendees();
  if (!attendees.length) {
    alert("회사명, 부서, 이름이 모두 채워진 행을 찾지 못했습니다.");
    return;
  }

  if (!state.templateFile) {
    alert("PPTX 템플릿을 업로드하세요.");
    return;
  }

  try {
    els.generateButton.disabled = true;
    els.generateButton.textContent = "생성 중...";

    const zip = await window.JSZip.loadAsync(await state.templateFile.arrayBuffer());
    const slideXmlFile = zip.file("ppt/slides/slide1.xml");
    const presentationXmlFile = zip.file("ppt/presentation.xml");
    const presentationRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    const contentTypesFile = zip.file("[Content_Types].xml");

    if (!slideXmlFile || !presentationXmlFile || !presentationRelsFile || !contentTypesFile) {
      throw new Error("PPTX 구조를 읽지 못했습니다. PowerPoint에서 저장한 .pptx 파일인지 확인하세요.");
    }

    const templateSlideXml = await slideXmlFile.async("string");
    const templateSlideRelsXml = await zip.file("ppt/slides/_rels/slide1.xml.rels")?.async("string");
    const presentationXml = await presentationXmlFile.async("string");
    const presentationRelsXml = await presentationRelsFile.async("string");
    const contentTypesXml = await contentTypesFile.async("string");
    const appPropertiesXml = await zip.file("docProps/app.xml")?.async("string");
    const templateSize = getPresentationSize(presentationXml);
    const layout = getOutputLayout(templateSize);
    const pageCount = Math.ceil(attendees.length / layout.perPage);

    removeExistingSlides(zip);

    for (let start = 0; start < attendees.length; start += layout.perPage) {
      const slideNumber = Math.floor(start / layout.perPage) + 1;
      const chunk = attendees.slice(start, start + layout.perPage);
      zip.file(`ppt/slides/slide${slideNumber}.xml`, buildSlideXmlFromTemplate(templateSlideXml, chunk, layout));
      if (templateSlideRelsXml) {
        zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, templateSlideRelsXml);
      }
    }

    const resizedPresentationXml = setPresentationSize(presentationXml, layout.page);
    zip.file("ppt/presentation.xml", buildPresentationXml(resizedPresentationXml, pageCount));
    zip.file("ppt/_rels/presentation.xml.rels", buildPresentationRelationshipsXml(presentationRelsXml, pageCount));
    zip.file("[Content_Types].xml", buildContentTypesXml(contentTypesXml, pageCount));
    if (appPropertiesXml) {
      zip.file("docProps/app.xml", updateAppPropertiesXml(appPropertiesXml, pageCount));
    }

    const blob = await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    downloadBlob(blob, "nametags_from_template.pptx");
  } catch (error) {
    console.error(error);
    alert(`명찰 PPTX 생성에 실패했습니다.\n\n${error.message || error}`);
  } finally {
    els.generateButton.textContent = "PPTX 생성";
    updatePreview();
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function downloadSampleCsv() {
  const rows = [
    ["회사명", "부서", "이름"],
    ["회사 A", "마케팅팀", "김하나"],
    ["회사 A", "영업팀", "이도윤"],
    ["회사 B", "인사팀", "박서연"],
  ];
  const csv = rows.map((row) => row.map((cell) => `"${cell}"`).join(",")).join("\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  downloadBlob(blob, "sample_attendees.csv");
}

[els.companySelect, els.departmentSelect, els.nameSelect, els.orientationSelect, els.gridSelect, els.copyInput].forEach((element) => {
  element.addEventListener("input", updatePreview);
  element.addEventListener("change", updatePreview);
});

els.fileInput.addEventListener("change", handleFileUpload);
els.templateInput.addEventListener("change", handleTemplateUpload);
els.generateButton.addEventListener("click", generateFromTemplate);
els.sampleButton.addEventListener("click", downloadSampleCsv);

updatePreview();
window.addEventListener("load", () => {
  window.setTimeout(setLibraryWarningIfNeeded, 800);
});
