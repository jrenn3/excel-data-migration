import './style.css';

// Get references to DOM elements
const uploadButton = document.getElementById('uploadButton');
const fileInput = document.getElementById('fileInput');

// "State" variable to track selected file
let selectedFile = null;

uploadButton.addEventListener('click', () => {
  fileInput.click();
});

fileInput.addEventListener('change', (event) => {
  const file = event.target.files[0];
  if (file) {
    selectedFile = file;

    // Immediately upload after file selection (or trigger another button if you prefer)
    uploadFile(file);
  }
});

function uploadFile(file) {
  const formData = new FormData();
  formData.append('file', file);

  // Optionally show a loading indicator here

  fetch('/upload', {
    method: 'POST',
    body: formData
  })
    .then(async (response) => {
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server error: ${errorText}`);
      }
      return response.blob(); // assuming the backend returns the new Excel file
    })
    .then((blob) => {
      // Create a download link
      const downloadUrl = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = 'updated_template.xlsx'; // or a dynamic name from response
      document.body.appendChild(link);
      link.click();
      link.remove();
    })
    .catch((error) => {
      console.error('Upload failed:', error.message);
      alert('Upload failed: ' + error.message);
    });
}