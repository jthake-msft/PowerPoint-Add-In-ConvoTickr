(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();
      jQuery('#get-data-from-selection').click(getDataFromSelection);
    });
  };

  // Reads data from current document selection and displays a notification
  function getDataFromSelection(){
 
    AzureADAuth.getAccessToken()
    .then(function (token) {
        //handle token
        $("#lblToken").html(token);
        
        var groupId = 'fc669da8-bdd7-4b71-90e7-3ac5f26703b4';
        var groupConversationsUrl = 'https://graph.microsoft.com/v1.0/groups/' + groupId + '/conversations';
        
        
    })
    .error(function (err) {
        //handle error
    });

  }
})();
