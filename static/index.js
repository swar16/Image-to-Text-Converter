document.addEventListener('DOMContentLoaded', (event) => {
  const fileInput = document.getElementById('file-input');
  const pdfDropBox = document.getElementById('pdfDropBox');
  const fileNameDisplay = document.getElementById('fileNameDisplay');

  const updateFileNameDisplay = (file) => {
    fileNameDisplay.textContent = `Selected file: ${file.name}`;
  };

  pdfDropBox.addEventListener('click', () => {
    fileInput.click();
  });

  pdfDropBox.addEventListener('dragover', (e) => {
    e.preventDefault();
  });

  pdfDropBox.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer.files.length) {
      fileInput.files = e.dataTransfer.files;
      updateFileNameDisplay(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) {
      updateFileNameDisplay(fileInput.files[0]);
      }
  });
});
