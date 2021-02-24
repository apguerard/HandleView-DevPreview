ej.base.enableRipple(true);

    
        // Rendering modal Dialog by enabling 'isModal' as true
        var dialogObj = new ej.popups.Dialog({
            width: '335px',
            header: 'Login',
            content: 'Northwind Login.',
            target: document.getElementById('target'),
            isModal: true,
            animationSettings: { effect: 'None' },
            buttons: [{
                    click: dlgButtonClick,
                    buttonModel: { content: 'Login', isPrimary: true }
                }],
            open: dialogOpen,
            close: dialogClose
        });
        dialogObj.appendTo('#modalDialog');

        function dlgButtonClick() {
            dialogObj.hide();
        }
        function dialogClose() {
            document.getElementById('wrapper').style.display = 'block';
        }
        function dialogOpen() {
            document.getElementById('wrapper').style.display = 'none';
        }
        