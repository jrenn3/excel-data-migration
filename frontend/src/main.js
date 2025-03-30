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

  const serverUrl = 
    window.location.hostname === 'localhost' // check if the app is running on localhost
      ? 'http://localhost:5000/upload' // local server URL
      : 'https://excel-data-migration-backend.onrender.com/upload'; // production server URL

  fetch(serverUrl, { //defines the path to the server endpoint
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
      link.download = 'DOWNLOADED_TEMPLATE_DEV_DRAFT_FUNdsForecast_ByMoe_v6_2025_03_29_1030.xlsm'; // todo- customize
      document.body.appendChild(link);
      link.click();
      link.remove();
    })
    .catch((error) => {
      console.error('Upload failed:', error.message);
      alert('Upload failed: ' + error.message);
    });
}