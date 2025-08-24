let clientId;
function sendActivity(action){
  if (!action || typeof google === 'undefined' || !google.script) return;
  const run = id => google.script.run.logClientActivity(id, action);
  if (clientId){
    run(clientId);
  } else {
    google.script.run.withSuccessHandler(id => {
      clientId = id;
      run(id);
    }).getClientId();
  }
}
