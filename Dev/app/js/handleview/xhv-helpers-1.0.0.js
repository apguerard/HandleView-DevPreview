/*   Copyright (C) 2019 Alain Guérard - All Rights Reserved
  You may use, distribute and modify this code under the
  terms of the MIT license.
 
  You should have received a copy of the MIT license with
  this file. If not, please visit : https://opensource.org/licenses/MIT */


/* 
 <summary>
  Prevent the right click to display the standard browser context menu
 </summary>
 <returns>No value. But the standard right click is disabled.</returns> 
*/  
function disableBrowserRightClick(){
    document.addEventListener('contextmenu', function(event){event.preventDefault()});
}
//END - disableBrowserRightClick()

/* 
 <summary>
  Remove actual rows in an HTML table.
  Requires JQuery
 </summary>
 <param name="tableId">The id value of the table you want to remove row from</param>
 <param name="fromRow">From which row number you want to remove. Useful to keep the headers rows.</param>
 <returns>No value. But rows are removed from the DOM in the table.</returns> 
*/  
function removeRows(tableId, fromRow){
    $("#" + tableId).find("tr:gt(" + fromRow + ")").remove();
    $("#" + tableId).find("tr:eq(" + fromRow + ")").remove();
}
//END - removeRows()

/* 
 <summary>
  Send a basic email
 </summary>
 <param name="email">The email of the recipient</param>
 <param name="subject">Subject of the email.</param>
 <param name="msg">Body of the email.</param>
 <returns>No value. Open your default mail program to send this email.</returns> 
*/ 
function sendMail(email,subject,msg)
{
    document.location.href = "mailto:"+ email + "?subject="
    + encodeURIComponent(subject)
    + "&body=" + encodeURIComponent(msg);
}
//END - sendMail()