/**
* Create a full-screen centered dialog for reading/previewing CSV or JSON files.
*/
(function(){

// ---------- Overlay (dark background behind modal) ----------
const overlay = document.createElement('div');
overlay.style.cssText = `
position:fixed;
top:0; left:0;
width:100%; height:100%;
background:rgba(0,0,0,0.45);
z-index:99998;
`;
document.body.append(overlay);

// ---------- Modal container (80% of screen) ----------
const modal = document.createElement('div');
modal.style.cssText = `
position:fixed;
top:50%; left:50%;
transform:translate(-50%, -50%);
width:80vw; height:80vh;
background:white;
padding:24px;
border-radius:14px;
box-shadow:0 12px 40px rgba(0,0,0,0.3);
font-family:sans-serif;
display:flex;
flex-direction:column;
z-index:99999;
`;
document.body.append(modal);

// ---------- Close button ----------
const closeBtn = document.createElement('button');
closeBtn.textContent = "×"; // simple "X" style close symbol
closeBtn.style.cssText = `
position:absolute;
top:14px; right:18px;
background:#ff3b3b;
color:white;
border:none;
border-radius:50%;
width:32px; height:32px;
font-size:20px;
font-weight:bold;
cursor:pointer;
display:flex;
align-items:center;
justify-content:center;
`;
modal.append(closeBtn);

// Remove modal + overlay when close is clicked
closeBtn.addEventListener('click', () => {
modal.remove();
overlay.remove();
});

// ---------- Header ----------
const header = document.createElement('h2');
header.textContent = "Contact Critical Insight Updater";
header.style.textAlign = "center";
header.style.marginBottom = "16px";
modal.append(header);

// ---------- File selection input ----------
const fileInput = document.createElement('input');
fileInput.type = "file";
fileInput.accept = ".csv,.txt";
fileInput.style.cssText = `padding:8px; font-size:15px; margin-bottom:14px; width:100%;`;
modal.append(fileInput);

// ---------- Content display area (for JSON output) ----------
const displayArea = document.createElement('pre');
displayArea.style.cssText = `
flex:1; /* take all remaining height */
background:#f7f7f7;
padding:16px;
border-radius:9px;
overflow:auto;
font-size:14px;
margin-bottom:14px;
`;
displayArea.textContent = "Processing output will appear here...";
modal.append(displayArea);

// ---------- Submit button ----------
const submitBtn = document.createElement('button');
submitBtn.textContent = "Process file";
submitBtn.style.cssText = `
padding:14px;
background:#0078ff;
color:white;
border:none;
border-radius:8px;
font-size:17px;
cursor:pointer;
`;
modal.append(submitBtn);

// ---------- Submit click handler ----------
submitBtn.addEventListener('click', async () => {
const file = fileInput.files[0];
if(!file){
alert("Choose a file first");
return;
}

// 1) Read CSV text (one contact GUID per line)
const text = await file.text();
const contactIds = text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);

// 2) Configure your environment + the field/value to update
const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
const apiRoot   = `${clientUrl}/api/data/v9.2`;

// 👉 Set the field & the value you want on every contact:
const fieldLogicalName = 'mag_hascriticalinsights';
const fieldValue       = false;
let processedCount = 0;

// 3) Build the Targets payload for UpdateMultiple
function buildTargets(ids, fieldName, value) {
  return ids.map(id => ({
    [fieldName]: value,                          // include only the column(s) you’re changing
    '@odata.type': 'Microsoft.Dynamics.CRM.contact', // REQUIRED by UpdateMultiple
    'contactid': id                              // key of the record
  }));
}

// 4) Chunk helper (tune size per your org’s behavior)
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// 5) Call UpdateMultiple for a chunk
async function updateMultipleChunk(targetsChunk) {
  const url = `${apiRoot}/contacts/Microsoft.Dynamics.CRM.UpdateMultiple`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'OData-Version': '4.0',
      'OData-MaxVersion': '4.0',
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      // Optional: include more error detail (especially helpful for elastic tables)
      'Prefer': 'odata.include-annotations=*'
    },
    body: JSON.stringify({ Targets: targetsChunk })
  });

  const responseText = await res.text();
  if (!res.ok) {
    displayArea.textContent += `\n\nUpdateMultiple failed: ${res.status}, ${res.statusText}, ${responseText}`;
    throw new Error(`UpdateMultiple failed: ${res.status} ${res.statusText}`);
  }

  processedCount += targetsChunk.length;
  // UpdateMultiple normally returns 204 No Content for standard tables
  displayArea.textContent += `\n   Updated ${processedCount} contact(s)`;
}

// 6) Build, chunk, and send
const targets = buildTargets(contactIds, fieldLogicalName, fieldValue);

// Start with ~500 per request for standard tables; adjust if you hit size/time limits
const chunks = chunk(targets, 200); // try 100–1000; elastic tables prefer ~100 [2](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bulk-operations)

displayArea.textContent += `\nSubmitting ${targets.length} updates in ${chunks.length} request(s)...`;
for (let i = 0; i < chunks.length; i++) {
  console.time(`UpdateMultiple-${i + 1}`);
  await updateMultipleChunk(chunks[i]);
  console.timeEnd(`UpdateMultiple-${i + 1}`);
}
console.log('All UpdateMultiple requests completed.');



displayArea.textContent += `\n\n${targets.length} records processed.`; // show in UI
});

})();