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
    })
    .error(function (err) {
        //handle error
    });

  }
})();
