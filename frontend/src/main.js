import './style.css';

// Get references to DOM elements
const uploadButton = document.getElementById('uploadButton');
const fileInput = document.getElementById('fileInput');

uploadButton.addEventListener('click', () => {
  fileInput.click();
});

fileInput.addEventListener('change', (event) => { //trigger when file is selected
  const file = event.target.files[0]; // Get the selected file, 0 refres to the first file in the list
  if (file) {
    console.log(`File selected: ${file.name}`);
    uploadFile(file); // Immediately upload after file selection 
  }
});

function uploadFile(file) {
  const formData = new FormData(); // browser API for handling form data
  formData.append('file', file); // builds request mimicing a form submission, with key 'file' and value as the file object

  fetch('https://excel-data-migration-backend.onrender.com/upload', { //defines the path to the server endpoint TODO-change to actual server
    method: 'POST', //posting data to the server
    body: formData // represents the form data
  })
    .then(async (response) => {
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server error: ${errorText}`);
      }
      return response.blob(); // reads the binary content (like an Excel file)
    })
    .then((blob) => {
      // Create a download link
      const downloadUrl = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = 'updated_template.xlsm'; // todo- customize
      document.body.appendChild(link);
      link.click();
      link.remove();
    })
    .catch((error) => {
      console.error('Upload failed:', error.message);
      alert('Upload failed: ' + error.message);
    });
}