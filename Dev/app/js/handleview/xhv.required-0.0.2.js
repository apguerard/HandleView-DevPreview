/* Copyright (C) 2019 Alain Guérard - All Rights Reserved
   You may use, distribute and modify this code under the
   terms of the MIT license.
 
   You should have received a copy of the MIT license with
   this file. If not, please visit : https://opensource.org/licenses/MIT */


//jQuery selector Extention
jQuery.extend(jQuery.expr[':'], {
    attrStartsWith: function (el, _, b) {
        for (var i = 0, atts = el.attributes, n = atts.length; i < n; i++) {
            if(atts[i].nodeName.toLowerCase().indexOf(b[3].toLowerCase()) === 0) {
                return true; 
            }
        } 
        return false;
    }
});
//END - jQuery selector Extention

/* 
 <summary>
   Attach the VBA eventDispatcher to dispatch HTML event to VBA function.
   This function is called directly from VBA code when a component load it's HTML template.
 </summary>
 <param name="guid">the guid of the compoenent</param>
 <returns>No value. But each eventlistener of the component is attached to the eventDispatcher element.</returns> 
*/
function attachEventDispatcher(guid){

    $('[xhv-eventlistener=' + guid  +']' ).each (function(){

    var ev = [];
    var node = $(this);

    /* add it in array */
    ev.push ([this.attributes["xhv-event"].value,this.attributes["xhv-eventhandler"].value, this.attributes["xhv-params"].value]);	

    /* treat array */
    ev.forEach(function(value){
        
        var  eventDispatcherId = "#eventdispatcher" + guid

        node.unbind(value[0]); 

        node.bind(value[0]
            , function(e){
            e.preventDefault();
            $(eventDispatcherId).attr("xhv-eventhandler",value[1]);
            $(eventDispatcherId).attr("xhv-params",value[2]);
            $(eventDispatcherId).click();
        });
    })	
});
};
// END - attachEventDispatcher() 

/* 
 <summary>
   Detach the VBA eventDispatcher that dispatch HTML event to VBA function.
   This function should be call directly from VBA code when a component dispose function is called.
 </summary>
 <param name="from">The HTML node from which the search of command should begin.</param>
 <returns>No value. But each command of the component is attached to the eventDispatcher element.</returns> 
*/
function detachEventDispatcher(from){
    $('#' + from + ' :attrStartsWith("xhv-eventhandler:")').each(function(){
        $(this).off();
    });
}
//END - detachEventDispatcher()




