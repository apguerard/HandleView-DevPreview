var ele = document.getElementById('container');
if(ele) {
    ele.style.visibility = "visible";
 }   

 //createSpinner() method is used to create spinner

ej.popups.createSpinner({

    // Specify the target for the spinner to show
  
    target: document.getElementById('container')
  
});
  
//showSpinner() will make the spinner visible

ej.popups.showSpinner(document.getElementById('container'));

setInterval(function () {

//hideSpinner() method used hide spinner 

ej.popups.hideSpinner(document.getElementById('container'))

}, 100000);