(function () {
  "use strict";

  // Call the initialize API first
  microsoftTeams.app.initialize();

  microsoftTeams.app.getContext().then(function (context) {
    if (context?.app?.host?.name) {
      updateHubName(context.app.host.name);
      console.log('test here')
      const tasksEl = document.getElementById("tasks")
let tasks;

var requestOptions = {
  method: 'GET',
  redirect: 'follow'
};

fetch("http://localhost:8080/task", requestOptions)
  .then(response => response.json())
  .then(result => 
    {
      const html = result.map(row =>  

                `<ul>
                  <li><label>Task name:</label> ${row.name}</li>
                  <li><label>Creation date:</label> ${row.creationDate}</li>
                  <li><label>Task state:</label> ${row.taskState}</li>
                </ul>`
               
            ).join('')
            tasksEl.innerHTML = html; 
    }
  )
  .catch(error => console.log('error', error));

    }
  });

  function updateHubName(hubName) {
    if (hubName) {
      document.getElementById("hubName").innerHTML = hubName;
      console.log("testststststst")
      const tasksEl = document.getElementById("tasks")
let tasks;

var requestOptions = {
  method: 'GET',
  redirect: 'follow'
};

fetch("http://localhost:8080/task", requestOptions)
  .then(response => response.json())
  .then(result => 
    {
      const html = result.map(row =>  

                `<ul>
                  <li><label>Task name:</label> ${row.name}</li>
                  <li><label>Creation date:</label> ${row.creationDate}</li>
                  <li><label>Task state:</label> ${row.taskState}</li>
                </ul>`
               
            ).join('')
            tasksEl.innerHTML = html; 
    }
  )
  .catch(error => console.log('error', error));

    }
  }
})();


