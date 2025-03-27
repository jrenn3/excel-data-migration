import './style.css';

document.getElementById('uploadButton').addEventListener('click', () => {
  document.getElementById('fileInput').click();
});

document.getElementById('fileInput').addEventListener('change', (event) => {
  const file = event.target.files[0];
  if (file) {
    console.log(`File selected: ${file.name}`);
    // Add further processing logic here, e.g., uploading to a server or parsing the file
  }
});