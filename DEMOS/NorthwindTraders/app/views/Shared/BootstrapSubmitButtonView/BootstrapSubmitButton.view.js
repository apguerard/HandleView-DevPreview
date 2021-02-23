function bindBootstrapValidation(formId, submitButtonId){
    // Fetch all the forms we want to apply custom Bootstrap validation styles to
    var forms = document.querySelectorAll('#'+ formId );
    // Loop over them and prevent submission
    var validation = Array.prototype.filter.call(forms, function(form) {
      form.addEventListener('submit', function(event) {
        if (form.checkValidity() === false) {
          event.preventDefault();
          event.stopPropagation();
          form.classList.add('was-validated');
        } else {
          event.preventDefault();
          event.stopPropagation();
          var submitButtonHidden = $('#' + submitButtonId);
          submitButtonHidden.click();         
          form.classList.remove('was-validated'); 
        }
        
      }, false);
    });
  };