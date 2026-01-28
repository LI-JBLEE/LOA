const form = document.getElementById("upload-form");
const salesInput = document.getElementById("sales_file");
const peopleInput = document.getElementById("people_file");
const salesName = document.getElementById("sales_filename");
const peopleName = document.getElementById("people_filename");
const statusText = document.getElementById("status-text");
const progressFill = document.getElementById("progress-fill");
const rowCount = document.getElementById("row-count");
const outputPath = document.getElementById("output-path");
const downloadButton = document.getElementById("download-button");
const runButton = document.getElementById("run-button");

let downloadUrl = null;
let progressTimer = null;

const setProgress = (value) => {
  progressFill.style.width = `${Math.min(100, Math.max(0, value))}%`;
};

const resetOutput = () => {
  rowCount.textContent = "-";
  outputPath.textContent = "-";
  downloadButton.disabled = true;
  downloadUrl = null;
};

const startProgress = () => {
  let progress = 5;
  setProgress(progress);
  statusText.textContent = "Uploading...";
  progressTimer = setInterval(() => {
    progress = Math.min(90, progress + Math.random() * 8);
    setProgress(progress);
    statusText.textContent = "Processing...";
  }, 500);
};

const stopProgress = () => {
  if (progressTimer) {
    clearInterval(progressTimer);
    progressTimer = null;
  }
};

salesInput.addEventListener("change", () => {
  salesName.textContent = salesInput.files[0]?.name || "No file selected";
});

peopleInput.addEventListener("change", () => {
  peopleName.textContent = peopleInput.files[0]?.name || "No file selected";
});

downloadButton.addEventListener("click", () => {
  if (downloadUrl) {
    window.location.href = downloadUrl;
  }
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!salesInput.files.length || !peopleInput.files.length) {
    statusText.textContent = "Please select both files.";
    return;
  }

  resetOutput();
  runButton.disabled = true;
  startProgress();

  const formData = new FormData();
  formData.append("sales_file", salesInput.files[0]);
  formData.append("people_file", peopleInput.files[0]);

  const xhr = new XMLHttpRequest();
  xhr.open("POST", "/process", true);

  xhr.upload.onprogress = (event) => {
    if (event.lengthComputable) {
      const uploadProgress = (event.loaded / event.total) * 40;
      setProgress(Math.max(5, uploadProgress));
      statusText.textContent = "Uploading...";
    }
  };

  xhr.onreadystatechange = () => {
    if (xhr.readyState !== XMLHttpRequest.DONE) {
      return;
    }
    stopProgress();
    runButton.disabled = false;

    if (xhr.status !== 200) {
      statusText.textContent = "Error.";
      try {
        const errorResponse = JSON.parse(xhr.responseText);
        outputPath.textContent = errorResponse.message || "Processing failed.";
      } catch {
        outputPath.textContent = "Processing failed.";
      }
      setProgress(0);
      return;
    }

    const response = JSON.parse(xhr.responseText);
    if (response.status !== "ok") {
      statusText.textContent = "Error.";
      outputPath.textContent = response.message || "Processing failed.";
      setProgress(0);
      return;
    }

    setProgress(100);
    statusText.textContent = "Completed.";
    rowCount.textContent = response.row_count ?? "-";
    outputPath.textContent = response.output_path ?? "-";
    downloadUrl = response.download_url;
    downloadButton.disabled = false;
  };

  xhr.send(formData);
});
