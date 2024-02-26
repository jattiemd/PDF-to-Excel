
// loading icon for conversion
$("#submit").click(function () {
  $("#loader-wrapper").css("display", "flex");
});

var setCookie = function (name, value, expiracy) {
  var exdate = new Date();
  exdate.setTime(exdate.getTime() + expiracy * 1000);
  var c_value =
    escape(value) +
    (expiracy == null ? "" : "; expires=" + exdate.toUTCString());
  document.cookie = name + "=" + c_value + "; path=/";
};

var getCookie = function (name) {
  var i, x, y, ARRcookies = document.cookie.split(";");
  for (i = 0; i < ARRcookies.length; i++) {
    x = ARRcookies[i].substr(0, ARRcookies[i].indexOf("="));
    y = ARRcookies[i].substr(ARRcookies[i].indexOf("=") + 1);
    x = x.replace(/^\s+|\s+$/g, "");
    if (x == name) {
      return y ? decodeURI(unescape(y.replace(/\+/g, " "))) : y; //;//unescape(decodeURI(y));
    }
  }
};

$("#submit").click(function () {
  $("#loader-wrapper").css("display", "flex");
  setCookie("downloadStarted", 0, 100); //Expiration could be anything... As long as we reset the value
  setTimeout(checkDownloadCookie, 1000); //Initiate the loop to check the cookie.
});
var downloadTimeout;
var checkDownloadCookie = function () {
  if (getCookie("downloadStarted") == 1) {
    setCookie("downloadStarted", "false", 100); //Expiration could be anything... As long as we reset the value
    $("#loader-wrapper").css("display", "none");
  } else {
    downloadTimeout = setTimeout(checkDownloadCookie, 1000); //Re-run this function in 1 second.
  }
};


// Select All tables checkbox display 
document.addEventListener('DOMContentLoaded', function() {
  var tableCheckboxes = document.querySelectorAll('.tableCheckboxes');
  var tableCheckboxAll = document.getElementById('tableCheckboxAll');
  var tableCheckboxAllLabel = document.getElementById('tableCheckboxAllLabel');
  var generateExcelBtn = document.getElementById('generateExcelBtn');

  // Check if there are visible Tablecheckbox checkboxes
  var anyVisible = Array.from(tableCheckboxes).some(function(checkbox) {
    return checkbox.offsetParent !== null;
  });
  
  // Toggle the visibility of TablecheckboxAll based on whether any Tablecheckbox checkboxes are visible
  tableCheckboxAll.style.display = anyVisible ? 'inline-flex' : 'none';
  tableCheckboxAllLabel.style.display = anyVisible ? 'inline-flex' : 'none';
  generateExcelBtn.style.display = anyVisible ? 'initial' : 'none';
});


// Check all checkboxes
function checkAll() {
  var checkboxes = document.querySelectorAll('.tableCheckboxes');
  var selectAllCheckbox = document.getElementById('tableCheckboxAll');

  checkboxes.forEach(function (checkbox) {
    checkbox.checked = selectAllCheckbox.checked;
  });
}

// Event listener for the 'tableCheckboxAll' checkbox
document.getElementById('tableCheckboxAll').addEventListener('change', checkAll);

// Hide convertContent div if generateExcelBtn is visible
$(document).ready(function(){
  if ($('#generateExcelBtn').is(':visible')) {
    $('#convertContent').hide(); 
  }
});

// Hide tables & checkboxes upon the appearance of the download button
$(document).ready(function(){
  if ($("#downloadBtn").is(":visible")) {
    $("#tables&Checkboxes").hide();
  }
});

// Remove confrim resubmission of form when reloading webpage
if ( window.history.replaceState ) {
  window.history.replaceState( null, null, window.location.href );
}

// Add event listener to the download link
document.getElementById("downloadLink").addEventListener("click", function() {
  // Redirect to the index page after a brief delay (adjust the delay as needed)
  setTimeout(function() {
    window.location.href = "/"; // Redirect to the index page
  }, 1000); // Delay in milliseconds
});