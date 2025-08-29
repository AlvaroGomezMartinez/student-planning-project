/* Script: remove-remote-node-files.js
 * Purpose: Use local clasp credentials (~/.clasprc.json) to call the Apps Script API
 * and remove files in the remote script project that match node_modules/**.
 *
 * Usage: node scripts/remove-remote-node-files.js <scriptId>
 * It will list files removed and update the remote project.
 *
 * Notes: The script reads ~/.clasprc.json for tokens and uses googleapis.
 */

const fs = require('fs');
const path = require('path');
const {google} = require('googleapis');

async function main(){
  const scriptId = process.argv[2];
  if(!scriptId){
    console.error('Usage: node scripts/remove-remote-node-files.js <scriptId>');
    process.exit(2);
  }

  const credPath = path.join(require('os').homedir(), '.clasprc.json');
  if(!fs.existsSync(credPath)){
    console.error('No ~/.clasprc.json found');
    process.exit(3);
  }

  const creds = JSON.parse(fs.readFileSync(credPath,'utf8'));
  // Prefer stored OAuth2 tokens
  const tokens = creds.tokens || creds.token || creds;

  const oauth2Client = new google.auth.OAuth2();
  oauth2Client.setCredentials(tokens);

  const script = google.script({version:'v1', auth: oauth2Client});
  try{
    const res = await script.projects.getContent({scriptId});
    const files = (res.data.files || []);
    if(!files.length){
      console.log('No files in remote project');
      return;
    }

    const toRemove = files.filter(f => {
      const name = f.name || f.fileName || '';
      return name.includes('node_modules') || name.startsWith('node_modules/');
    }).map(f => f.name || f.fileName);

    if(!toRemove.length){
      console.log('No remote node_files found');
      return;
    }

    console.log('Remote files matched to remove:', toRemove.length);
    toRemove.slice(0,100).forEach(n => console.log('  -', n));

    const filtered = files.filter(f => !( (f.name||f.fileName||'').includes('node_modules') ));

    // Prepare update payload - the API expects 'files' array
    const updateBody = { files: filtered };
    const updateRes = await script.projects.updateContent({ scriptId, requestBody: updateBody });
    console.log('updateContent response status:', updateRes.status);
    console.log('Removed', toRemove.length, 'files from remote project.');
  }catch(err){
    console.error('Error:', err.message || err);
    if(err.errors) console.error(err.errors);
    process.exit(4);
  }
}

main();
