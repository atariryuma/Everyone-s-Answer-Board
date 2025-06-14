const fs = require('fs');
const {google} = require('googleapis');

async function main() {
  const clasp = JSON.parse(fs.readFileSync('.clasp.json', 'utf8'));
  const auth = new google.auth.GoogleAuth({
    scopes: [
      'https://www.googleapis.com/auth/script.projects',
      'https://www.googleapis.com/auth/script.deployments'
    ]
  });
  const authClient = await auth.getClient();
  const script = google.script({version: 'v1', auth: authClient});
  const res = await script.projects.deployments.list({
    scriptId: clasp.scriptId,
  });
  const deployments = res.data.deployments || [];
  if (!deployments.length) {
    throw new Error('No deployments found');
  }
  const webAppUrl = deployments
    .flatMap(d => d.entryPoints || [])
    .find(e => e.webApp)?.webApp?.url;
  if (!webAppUrl) {
    throw new Error('Web app URL not found in deployments');
  }
  await script.scripts.run({
    scriptId: clasp.scriptId,
    requestBody: {
      function: 'saveWebAppUrl',
      parameters: [webAppUrl],
    }
  });
  console.log('Saved Web App URL:', webAppUrl);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
