const searchInput = document.querySelector("#searchInput");
const exactMatch = document.querySelector("#exactMatch");
const wholeWordMatch = document.querySelector("#wholeWordMatch");
const glossaryCheckboxes = document.querySelector("#glossaryCheckboxes");
const termList = document.querySelector("#termList");
const resultArea = document.querySelector("#resultArea");
const glossaryFiles = document.querySelector("#glossaryFiles");
const uploadButton = document.querySelector("#uploadButton");
const statusMessage = document.querySelector("#statusMessage");

let searchDebounce;

function debounce(fn, wait = 250) {
  return function (...args) {
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => fn(...args), wait);
  };
}

async function fetchJSON(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    const error = await response.text();
    throw new Error(error || "Request failed");
  }
  return response.json();
}

async function refreshGlossaries() {
  try {
    const { glossaries } = await fetchJSON("/api/glossaries");
    renderGlossaries(glossaries);
    statusMessage.textContent = `Loaded glossaries: ${glossaries.length}`;
    await refreshTerms();
  } catch (err) {
    statusMessage.textContent = err.message;
  }
}

function renderGlossaries(glossaries) {
  glossaryCheckboxes.innerHTML = "";
  if (!glossaries.length) {
    glossaryCheckboxes.innerHTML = "<p class='empty'>No glossaries loaded yet.</p>";
    return;
  }

  glossaries.forEach((glossary) => {
    const wrapper = document.createElement("label");
    wrapper.className = "glossary-checkbox";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = glossary.selected;
    checkbox.dataset.id = glossary.id;
    checkbox.addEventListener("change", handleGlossaryToggle);

    const name = document.createElement("span");
    name.textContent = `${glossary.name} (${glossary.terms_count} rows)`;
    name.className = "glossary-name";
    wrapper.appendChild(checkbox);
    wrapper.appendChild(name);

    if (!glossary.preload_terms) {
      const note = document.createElement("span");
      note.className = "glossary-note";
      note.textContent = "Large glossary (search only)";
      wrapper.appendChild(note);
    }
    glossaryCheckboxes.appendChild(wrapper);
  });
}

async function handleGlossaryToggle(event) {
  const { id } = event.target.dataset;
  const selected = event.target.checked;
  try {
    await fetchJSON(`/api/glossaries/${id}/selection`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ selected }),
    });
    await refreshTerms();
  } catch (err) {
    statusMessage.textContent = err.message;
  }
}

async function refreshTerms() {
  try {
    const params = new URLSearchParams();
    params.append("search", searchInput.value);
    params.append("exact", exactMatch.checked);
    params.append("whole_word", wholeWordMatch.checked);
    const { terms } = await fetchJSON(`/api/terms?${params.toString()}`);
    renderTermList(terms);
  } catch (err) {
    statusMessage.textContent = err.message;
  }
}

function renderTermList(terms) {
  termList.innerHTML = "";
  if (!terms.length) {
    termList.innerHTML = "<li class='empty'>No terms match the current filters.</li>";
    return;
  }

  terms.forEach((term) => {
    const item = document.createElement("li");
    item.textContent = term;
    item.dataset.term = term;
    termList.appendChild(item);
  });
}

termList.addEventListener("click", (event) => {
  const term = event.target.dataset.term;
  if (term) {
    fetchTermDetails(term);
  }
});

async function fetchTermDetails(term) {
  try {
    const data = await fetchJSON(`/api/terms/details/${encodeURIComponent(term)}`);
    renderTermDetails(data);
  } catch (err) {
    statusMessage.textContent = err.message;
  }
}

function renderTermDetails(data) {
  if (!data.results.length) {
    resultArea.innerHTML = "<p class='empty'>No definition available for the selected term.</p>";
    return;
  }
  resultArea.innerHTML = "";
  data.results.forEach((entry) => {
    const group = document.createElement("section");
    group.className = "result-group";

    const title = document.createElement("h3");
    title.textContent = entry.glossary;
    group.appendChild(title);

    entry.rows.forEach((row) => {
      const list = document.createElement("dl");
      list.className = "row-data";
      Object.entries(row).forEach(([key, value]) => {
        const term = document.createElement("dt");
        term.textContent = key;
        const description = document.createElement("dd");
        description.textContent = value || "N/A";
        list.appendChild(term);
        list.appendChild(description);
      });
      group.appendChild(list);
    });

    resultArea.appendChild(group);
  });
}

uploadButton.addEventListener("click", async () => {
  const files = Array.from(glossaryFiles.files);
  if (!files.length) {
    statusMessage.textContent = "Select one or more CSV/XLSX files first.";
    return;
  }

  const formData = new FormData();
  files.forEach((file) => formData.append("files", file, file.name));

  try {
    await fetchJSON("/api/glossaries/upload", {
      method: "POST",
      body: formData,
    });
    glossaryFiles.value = "";
    statusMessage.textContent = "Glossaries uploaded successfully.";
    await refreshGlossaries();
  } catch (err) {
    statusMessage.textContent = err.message;
  }
});

[searchInput, exactMatch, wholeWordMatch].forEach((el) =>
  el.addEventListener("input", debounce(refreshTerms))
);

window.addEventListener("DOMContentLoaded", () => {
  refreshGlossaries();
  resultArea.innerHTML = "<p class='empty'>Upload a glossary and select a term to see its definition.</p>";
});

