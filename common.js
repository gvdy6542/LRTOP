let clientId;
function sendActivity(action){
  if (!action) return;
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
