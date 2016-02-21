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
        
        var groupId = '202c77b2-5f8f-4a31-a938-7b77b351c8ed';
        var groupConversationsUrl = 'https://graph.microsoft.com/v1.0/groups/' + groupId + '/conversations';
        
        $.ajax({
            url: groupConversationsUrl,
            type: "GET",
            }
         })
        .then(function(data) {
            
        })
        .error(function(err)
        {
            
        });
    })
    .error(function (err) {
        //handle error
    });

  }
})();
