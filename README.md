# HandleView
A VBA framework that lets you build your UI in HTML by using **only one form**. I like to call it a Single Form Application (SFA).  
This framework was originally developed by <a href="https://github.com/apguerard/HandleView-DevPreview">Alain Guerard</a> but has undergone some significant revamping here.  Credit goes to him for putting all the peices together to get VBA talking to JS and vice versa!  

## What is HandleView?
Bringing Microsoft Access user interface development to the 21st century, HandleView is a library that leverages the Web Browser control to render Javascript, HTML, CSS (Bootstrap anyone?) along with JQuery and now, with <a href="https://github.com/syncfusion/ej2-javascript-ui-controls">Syncfusion's amazing e2j-javascript ui control library.</a>

Use your standard web development skills to create beautiful, responsive user interfaces!

## How does it work?
Upon startup, the only form in the Access Database launches fully maximized with only one control - a Web Browser control.  Using the on_load event, the VBA controller renders the startup 'route' to load a view template file from disk, interpolates the tokens with relevant data and controls, and outputs a fully built page that is then loaded into the WB DOM.  When the WB loads the DOM, it immediately executes any Javascript code at least once during the lifetime of the view and renders the page accordingly.

Using your favorite Javascript UI Library (JQuery UI, or the newly supoorted Syncfusion controls), you build your components into an in-memory DOM until you have finished building your view.  Applying the in memory DOM to the WB causes it to render in the WB control.

Data can be passed from the VBA controller to the view during the interpolation process and by passing in values using a special 'props' attribute.  Any data and control structures are fully assembled as HTML by the VBA controller.  Once the controller is finished constructing the entire page, it renders it in the WB control on the form.

Rendering causes the WB to reload it's new DOM and executes any javascript at least once when the document is fully loaded.  Any javascript that you want executed the next time the WB refreshes can be accomplished with special attributes on the custom script tag.

## What about form submissions?
Through some javascript intervention, POSTS are intercepted, forms can be validated, and then a simple click event on a hidden button is fired instead of the POST.  This click event (or any DOM event for that matter) can be assigned to a VBA controller method by assigning the method as the event handler.  The WB passes data back to VBA by way of an event and form data.  VBA dutifully grabs the form data and converts it to parameters that are required by the event handler.

## What about navigation?
At the most fundamental level, navigation is nothing more than raised events with parameters.  Those events could be caused by anyting as simple as a menu button on the Nav bar or as a result of a complex set of instructions called by javascript.  Parameters are passed to the controller during nav in the form of JSON query strings appended to the end of the requested route.

When a controller's navigation method is called, the url is parsed for a requested route and for JSON data.  The requested route is then looked up in the routes module to determine which controller to call and which view it should build.  From there, the whole process described above starts again.

## Does xHV use Dependency Injection?
YES!  While reflection is not available in VBA, classes are regestered against interfaces and can be instantiated either as a Singleton or as Transient.

Please check out the Wiki for details.

