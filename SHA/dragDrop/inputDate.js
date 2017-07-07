

// iMouseDown represents the current mouse button state: up or down
/*
lMouseState represents the previous mouse button state so that we can
check for button clicks and button releases:

if(iMouseDown && !lMouseState) // button just clicked!
if(!iMouseDown && lMouseState) // button just released!
*/
var mouseOffset = null;

//link and push
var keycode0 = 68;
var dKeyDown = false;
var frstKeySlctn = null;
var scndKeySlctn = null;

//group select
var keycode1 = 71;
var gKeyDown = false;
var lMouseState = false;
var frstSlctnGrup = null;
var scndSlctnGrup = null;
var subSlctnOffsetY = 0; // last selected sub item offset distance to dragHelper
var subSlctnOffsetX = 0; // last selected sub item offset distance to dragHelper

//s key for marking screw pulling
var keycode2 = 83;
var sKeyDown = false;
var slctntarget = null;

//context menu
var _replaceContext = false; // replace the system context menu?
var _mouseOverContext = false; // is the mouse over the context menu?
var _submouseOverContext = false;//is the mouse over the sub context menu?
var _divContext = null; 

// Demo 0 variables
var blockMarginTop = 0;
var blockMarginLeft = 0;
var dragHelperparentNodeoffsetTop = 0;
var dragHelperparentNodeoffsetLeft = 0;

var dragHelper = null;

var ganttZoneleft = 0;
//for production line menu event
var targetClone = null;
//scalar div
var scalar1 = null;
var bckgrndImage1 = null;
var pixelsPerDay = 90;



Number.prototype.NaN0 = function () { return isNaN(this) ? 0 : this; };


function mouseCoords(ev) {

    if (ev.pageX || ev.pageY) {
        return { x: ev.pageX, y: ev.pageY };
    } else {
        return { x: ev.clientX + document.body.scrollLeft - document.body.clientLeft, y: ev.clientY + document.body.scrollTop - document.body.clientTop };
    }
}

function getPosition(e) {
    var left = 0;
    var top = 0;

    while (e.offsetParent) {
        
        left += e.offsetLeft;
        top += e.offsetTop;
        e = e.offsetParent;
    }

    left += e.offsetLeft;
    top += e.offsetTop;

    return { x: left, y: top };
}

function getMouseOffset(target, ev) {
    ev = ev || window.event;

    var docPos = getPosition(target);
    var mousePos = mouseCoords(ev);
    return { x: mousePos.x - docPos.x, y: mousePos.y - docPos.y };
}


function mouseMove(ev) {
    ev = ev || window.event;   
    //Firefox uses event.target here, MSIE uses event.srcElement
   
    var target = ev.target || ev.srcElement;
    var mousePos = mouseCoords(ev);

    

    if (targetClone) {
        // move our helper div to wherever the mouse is (adjusted by mouseOffset)
        dragHelper.style.top = (mousePos.y - mouseOffset.y - dragHelperparentNodeoffsetTop - blockMarginTop) + 'px';
        dragHelper.style.left = (mousePos.x - mouseOffset.x - dragHelperparentNodeoffsetLeft - blockMarginLeft) + 'px';
        //debugger;

    }

    if (lMouseState && frstSlctnGrup && scndSlctnGrup) {
        // move our helper div to wherever the mouse is (adjusted by mouseOffset)
        dragHelper.style.top = (mousePos.y - mouseOffset.y - dragHelperparentNodeoffsetTop - blockMarginTop - subSlctnOffsetY) + 'px';
        dragHelper.style.left = (mousePos.x - mouseOffset.x - dragHelperparentNodeoffsetLeft - blockMarginLeft - subSlctnOffsetX) + 'px';
    }

    //if mouse not over customized context menu and its sub context menu and contextMenu is on
    if (!_mouseOverContext && !_submouseOverContext && _replaceContext) {
        CloseContext();
    }

    //window.status = "event.clientY:" + ev.clientY + " mousePos.y:" + mousePos.y;
    // this helps prevent items on the page from being highlighted while dragging
    return false;
}

function mouseUp(ev) {


    ev = ev || window.event;

    lMouseState = false; // mouse is up



    if ((ev.button != 2) && targetClone) { //if left button or middle button is up instead of right button

        InteractionWithDatabase1('ganttZone', 'action=updateorder&timepixels=' + parseInt(dragHelper.offsetLeft) + '&lineno=' + $('ganttZone').getAttribute('crntLine') + '&orderkey=' + targetClone.id)
        window.status = "move order '" + targetClone.id  + "' to a new start time";
        for (var i = 0; i < dragHelper.childNodes.length; i++) dragHelper.removeChild(dragHelper.childNodes[i]);
        dragHelper.style.display = 'none';
        targetClone = null;
    }

    if (scndSlctnGrup && frstSlctnGrup) {

        for (var i = 0; i < dragHelper.childNodes.length; i++) {
            //dragHelper.parentNode.getElementById(dragHelper.childNodes[i].id).style.visibility = "visible";
            //$(dragHelper.childNodes[i].id).style.visibility = "visible";
            //dragHelper.removeChild(dragHelper.childNodes[i]);
            //dragHelper.style.display = 'none';
        }

    }

 
}


function mouseDown(ev) {

    lMouseState = true; // mouse is down

    ev = ev || window.event;

    var target = ev.target || ev.srcElement;

    //if left button or middle button is clicked instead of right button===============================
    if (ev.button != 2) { 
       

        //initiate below values in order to speed up processing in the step of mouse moving
        blockMarginTop = parseInt(target.style.marginTop);
        blockMarginLeft = parseInt(target.style.marginLeft);
        dragHelperparentNodeoffsetTop = dragHelper.parentNode.offsetTop;
        dragHelperparentNodeoffsetLeft = dragHelper.parentNode.offsetLeft;


        // if the target is one of the order blocks 
        if (target.className.indexOf("g-") >= 0) {

                // Only ctrl key  is still down ============================
                if (dKeyDown) { // if ctrl key is down, first priority to handle it and just ignore g key
                    if (!frstKeySlctn) {
                        frstKeySlctn = target.id;
                        target.style.border = "1px dashed red";
                    }
                    else {
                        if (scndKeySlctn && (scndKeySlctn != frstKeySlctn)) {
                            $(scndKeySlctn).style.border = "1px solid black";
                        }
                        scndKeySlctn = target.id;
                        target.style.border = "1px dashed red";
                    }

                    return;
                }

                // if g key is down
                if ( gKeyDown ){  

                    if (!frstSlctnGrup) {
                        frstSlctnGrup = target.id;
                        target.style.border = "1px dashed yellow";
                    }
                    else {
                        if (scndSlctnGrup) { //if we already have selected second element, then we stop further selection
                            return;
                                $(scndSlctnGrup).style.border = "1px solid black";
                                for (var i = 0; i < dragHelper.childNodes.length; i++) {
                                    dragHelper.parentNode.getElementById(dragHelper.childNodes[i].id).style.visibility = "visible";
                                    dragHelper.removeChild(dragHelper.childNodes[i]);
                                }
                            }

                            scndSlctnGrup = target.id;
                            target.style.border = "1px dashed yellow";
                            
                            if (scndSlctnGrup != frstSlctnGrup) {

                                    var temp = null;

                                    subSlctnOffsetY = $(scndSlctnGrup).offsetTop  - $(frstSlctnGrup).offsetTop;
                                    subSlctnOffsetX = $(scndSlctnGrup).offsetLeft - $(frstSlctnGrup).offsetLeft;

                                    if ($(frstSlctnGrup).offsetLeft > $(scndSlctnGrup).offsetLeft) {
                                        temp = frstSlctnGrup;
                                        frstSlctnGrup = scndSlctnGrup;
                                        scndSlctnGrup = temp;
                                    }

                                    var sibling1 = $(frstSlctnGrup);
                                    var offsetLeft1 = sibling1.offsetLeft;
                                    for (var i = 0; i < dragHelper.childNodes.length; i++) dragHelper.removeChild(dragHelper.childNodes[i]);
                                    var Clone1 = sibling1.cloneNode(true);
                                    dragHelper.appendChild(Clone1);
                                    Clone1.style.left = '0px';

                                    if (temp) {
                                        subSlctnOffsetY = 0;
                                        subSlctnOffsetX = 0;
                                    }

                                    sibling1.style.visibility = "hidden";

                                    do {
                                        
                                        sibling1 = sibling1.nextSibling;

                                        Clone1 = sibling1.cloneNode(true);
                                        dragHelper.appendChild(Clone1);
                                        Clone1.style.left = (sibling1.offsetLeft - offsetLeft1) + 'px';
                                        sibling1.style.visibility = "hidden";

                                    }
                                    while (sibling1.id != scndSlctnGrup);

                                   
                                    dragHelper.style.left = $(frstSlctnGrup).style.left;
                                    dragHelper.style.top = ($(frstSlctnGrup).offsetTop - blockMarginTop) + 'px';
                                    dragHelper.style.width = (parseInt(Clone1.style.left) + parseInt(Clone1.style.width) + parseInt(Clone1.style.borderWidth) * 2) + "px";
                                    dragHelper.style.display = 'block';

                                    mouseOffset = getMouseOffset(target, ev);
                        
                            }
                    }

                    return;
                }


                // Only s key  is still down ============================
                if (sKeyDown) { // if ctrl key is down, first priority to handle it and just ignore g key
               
                    slctntarget = target;
                    return;
                }
            
                var targetPosition = getPosition(target);
                mouseOffset = getMouseOffset(target, ev);
                
                targetClone = target.cloneNode(true);
                dragHelper.style.left = target.style.left;
                dragHelper.style.top = (target.offsetTop - blockMarginTop) + 'px';
                              
                target.style.visibility = "hidden";
                for (var i = 0; i < dragHelper.childNodes.length; i++) dragHelper.removeChild(dragHelper.childNodes[i]);
                targetClone.style.left = '0px';
                dragHelper.appendChild(targetClone);
                dragHelper.style.display = 'block';

        }
    }



    //if right button  is clicked on top of one of order details block=============================================

    if ((ev.button == 2) && (target.className.indexOf("g-") >= 0)) {
        _replaceContext = true;
        UsingASPdotNet(_divContext.id, 'action=contextMenu&lineno=' + $('ganttZone').getAttribute('crntLine') + '&orderkey=' + target.id);
        window.status = "order context menu pop up"
    }
    else {
        _replaceContext = false;
    }

}


function ContextShow(event) { 
         
        // IE is evil and doesn't pass the event object 
        if (event == null) event = window.event; // we assume we have a standards compliant browser, but check if we have IE 
        var target = event.target != null ? event.target : event.srcElement;
        var mousePos = mouseCoords(event);

        if (_replaceContext) { 
            
            // document.body.scrollTop does not work in IE 
            var scrollTop = document.body.scrollTop ? document.body.scrollTop : document.documentElement.scrollTop;
            var scrollLeft = document.body.scrollLeft ? document.body.scrollLeft : document.documentElement.scrollLeft;
            // hide the menu first to avoid an "up-then-over" visual effect
            _divContext.style.visibility = 'hidden';
            _divContext.style.display = 'block';
            _divContext.style.left = (event.clientX + scrollLeft - 2) + 'px';
            //if ( mousePos.y <= 31*16 ) {
            if (mousePos.y <= (5 * 16 + parseInt(_divContext.offsetHeight))) {
                if (mousePos.y <= 5 * 16) {
                    _divContext.style.top = mousePos.y + scrollTop - 16 + 'px'; 
                }
                else {
                    _divContext.style.top = 5 * 16 + scrollTop + 'px';
                }
            }
            else {
                //_divContext.style.top = event.clientY + scrollTop - parseInt(_divContext.offsetHeight) + 2 + 'px';
                _divContext.style.top = (mousePos.y - parseInt(_divContext.offsetHeight) + scrollTop + 16) + 'px';
            }
            _divContext.style.visibility = 'visible';

            return false;
        }
 }


 function CloseContext() {
     _mouseOverContext = false;
     _submouseOverContext = false;
        _replaceContext = false;
        _divContext.style.display = 'none';
        
 }


 //change production line for an order
 function changeLine(theobject){
     InteractionWithDatabase1('ganttZone', 'action=changeline&lineno=' + $('ganttZone').getAttribute('crntLine') + '&orderkey=' + theobject.parentNode.getAttribute('id1') + '&newlineno=' + theobject.getAttribute('id1'));
     window.status = "order '" + theobject.parentNode.getAttribute('id1') + "' is changed to new line " + theobject.getAttribute('id1');
 }


 function InitContext() {
        _divContext = $('divContext');
        _divContext.onmousemove = function () { _mouseOverContext = true; };
        _divContext.onmouseout = function () {_mouseOverContext = false;};

}

//to display contextmenu list
function dsplyCntxtMnuLst(theobject) {
        _submouseOverContext = true;
        theobject.parentNode.style.display = "block";
        theobject.style.backgroundColor = "Silver";

    var cells = theobject.parentNode.getElementsByTagName("li");

    for (var i = 0; i < cells.length; i++) {
        if (cells[i] !== theobject) {
            cells[i].style.backgroundColor = "white"; //back to original color if this is not the current object
        };
    }
}



function keyDown(ev) {
    ev = ev || window.event;
 // if we are using custom context menu, disable these short cut functions in case potential conflicts
    if (_divContext.style.display != 'block') {
        if (ev.keyCode == keycode0) { //D key is down for linking and delay in order
            dKeyDown = true;
            return true;
        }

        if (ev.keyCode == keycode1) { //G key is down for group moving
            gKeyDown = true;
            return true;
        }

        if (ev.keyCode == keycode2) { //S key is down for screw mark
            sKeyDown = true;
            return true;
        }
	}
}

function keyUp(ev) {
    ev = ev || window.event;

    // if we are using custom context menu, disable these short cut functions in case potential conflicts
    if (_divContext.style.display != 'block') {
        // when s key is up
        if (ev.keyCode == keycode2) {

            sKeyDown = false;

            InteractionWithDatabase1('ganttZone', 'action=pullscrew&orderkey=' + slctntarget.id + '&lineno=' + $('ganttZone').getAttribute('crntLine') + ''); // &screwstatus=' + target.childNodes[1].innerHTML);
            window.status = "screw pulling is changed for order '" + slctntarget.id  + "'";

            slctntarget = null;
            return true;
        }



        if (ev.keyCode == keycode1) { // when G key is up

            gKeyDown = false;

            if (frstSlctnGrup) {
                $(frstSlctnGrup).style.border = "1px solid black";
            }

            if (scndSlctnGrup) {
                $(scndSlctnGrup).style.border = "1px solid black";
            }

            if (frstSlctnGrup && scndSlctnGrup && (frstSlctnGrup != scndSlctnGrup)) { //need select two different elements

                InteractionWithDatabase1('ganttZone', 'action=moveingroup&frstslctnID=' + frstSlctnGrup + '&scndslctnID=' + scndSlctnGrup + '&areaoffsetleft=' + parseInt(dragHelper.offsetLeft) + '&lineno=' + $('ganttZone').getAttribute('crntLine') + '');
                window.status = "Start times were changed for a group of orders";
            }

            frstSlctnGrup = null;
            scndSlctnGrup = null;

            return true;
        }



        if (ev.keyCode == keycode0) {  // when D key is up
            dKeyDown = false;

            if (frstKeySlctn) {
                $(frstKeySlctn).style.border = "1px solid black";
            }

            if (scndKeySlctn) {
                $(scndKeySlctn).style.border = "1px solid black";
            }

            if (frstKeySlctn && scndKeySlctn && (frstKeySlctn != scndKeySlctn)) { //need select two different elements

                InteractionWithDatabase1('ganttZone', 'action=updateorderinbatch&frstslctnID=' + frstKeySlctn + '&scndslctnID=' + scndKeySlctn + '&lineno=' + $('ganttZone').getAttribute('crntLine') + '');
                window.status = "A group of order were changed (latter order's start time equal to the earlier one's finish time)"
            }


            frstKeySlctn = null;
            scndKeySlctn = null;

            return true;

        }

    }
}


document.onmousemove = mouseMove;
document.onmousedown = mouseDown;
document.onmouseup = mouseUp;
document.onkeydown = keyDown;
document.onkeyup = keyUp;
document.oncontextmenu = ContextShow;


window.onscroll = function () {

    var substitute = getScrollXY().x + ganttZoneleft;
    for (var i = 0; i < scalar1.childNodes.length; i++) {
        scalar1.childNodes[i].style.left = (i * pixelsPerDay - substitute) + 'px';
        bckgrndImage1.childNodes[i].style.left = (i * pixelsPerDay - substitute) + 'px';

    }
    
}

//ignore vertical scroll operation
function getScrollXY() {
    var scrOfX = 0;
    var scrOfY = 0;
    if (typeof (window.pageYOffset) == 'number') {
        //Netscape compliant
        //scrOfY = window.pageYOffset;
        scrOfX = window.pageXOffset;
    } else if (document.documentElement.scrollLeft) { // document.documentElement && || document.documentElement.scrollTop)) {
        //IE6 standards compliant mode
        //scrOfY = document.documentElement.scrollTop;
        scrOfX = document.documentElement.scrollLeft;
    } else if (document.body.scrollLeft) {  //document.body &&  || document.body.scrollTop)) {
       //DOM compliant
        //scrOfY = document.body.scrollTop;
        scrOfX = document.body.scrollLeft;
    }

    return {x : scrOfX, y : scrOfY};
}



window.onload = function () {

    //var nvgtrursrgnt = navigator.userAgent;
   // if ((nvgtrursrgnt.indexOf("MSIE 7.0") >= 0) || (nvgtrursrgnt.indexOf("MSIE 8.0") >= 0)) {
        //do nothing
   // }
   // else {
        //alert("It only can function well in internet explorer 7.0 or above while other explorer probably can not function well for some features. ");
   // }



    // Create our helper object that will show the item while dragging
    dragHelper = document.createElement('div');
    dragHelper.style.cssText = "position:absolute;display:none;";

    //document.body.appendChild(dragHelper);


    InteractionWithDatabase1('lines', 'action=productionlines'); // fill up div with id 'lines'

    $('scalar').style.cssText = "border: 0px solid black;z-index:1;position:fixed;background-color:white;top:3.5em;"; //width:" + (90 * 90 + 2) + "px;"
    UsingASPdotNet('scalar', 'action=scalar'); // fill up div with id 'scalar'
    $('bckgrndImg').style.cssText = "border: 0px solid black;z-index:-1;position:fixed;background-color:transparent;top:4em;"; //width:" + (90 * 90 + 2) + "px;"
    UsingASPdotNet('bckgrndImg', 'action=bckgrndImg');


    var ganttChartArea = document.createElement("div");
    ganttChartArea.className = "DragContainer";
    ganttChartArea.style.cssText = "border: 0px solid black;margin:0px 0px;left:0px;top:4.9em;position:absolute;width:auto;";  // + (90 * 90) + "px;"
    ganttChartArea.id = "ganttZone";
    ganttZoneleft = parseInt(ganttChartArea.style.left);

    document.body.appendChild(ganttChartArea);

    //Fill container area with first available production line
    document.getElementsByTagName('td')[0].onmousemove();
    //ganttChartArea.getElementsByTagName('td')[0].onmousemove();


    //make both scalar and fixedscalar's left equal to element ganttChartArea's left
    $('scalar').style.left = ganttChartArea.style.left;

    scalar1 = $('scalar');
    bckgrndImage1 = $('bckgrndImg');
    //get the number of pixels per day
    if (scalar1.childNodes.length >1 ){
        pixelsPerDay = parseInt(scalar1.childNodes[1].style.left) - parseInt(scalar1.childNodes[0].style.left)
    }



    InitContext();
}



//AJAX get order line details from ASPX database service
function InteractionWithDatabase1(elementID1, actioncode) {

    var xmlhttp;
    if (elementID1 == "") {
       return;

    }
    if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp = new XMLHttpRequest();
    }
    else {// code for IE6, IE5
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    xmlhttp.onreadystatechange = function () {
        if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {

            $(elementID1).innerHTML = xmlhttp.responseText;
            //create the box to contain the object to be moved around
            dragHelper = document.createElement('div');
            dragHelper.style.cssText = "position:absolute;display:none;";
            $(elementID1).appendChild(dragHelper);
        }
    }
    xmlhttp.open("POST", "DataExchange.aspx", false);
    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    xmlhttp.send(actioncode);
    //xmlhttp.send(); 

}


function UsingASPdotNet(elementID1, actioncode) {

    var xmlhttp;
    if (elementID1 == "") {
        return;

    }
    if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp = new XMLHttpRequest();
    }
    else {// code for IE6, IE5
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    xmlhttp.onreadystatechange = function () {
        if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {

            $(elementID1).innerHTML = xmlhttp.responseText;
        }
    }
    xmlhttp.open("POST", "DataExchange.aspx", false);
    //xmlhttp.open("GET", "DataExchange.aspx?", true);
    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    xmlhttp.send(actioncode);
    //xmlhttp.send(); 

}



//show all the Gantt chart for orders per production line
function dsplyOrderByLine(theobject) {

    if ($('ganttZone').getAttribute('crntLine') != theobject.getAttribute('id').toString()) {

        $('ganttZone').setAttribute('crntLine', theobject.getAttribute("id").toString());
        InteractionWithDatabase1('ganttZone', 'action=orderlines&lineno=' + theobject.getAttribute("id").toString() + '');
        window.status = "show orders by production line " + theobject.getAttribute("id").toString();
        theobject.style.backgroundColor = "Silver";

    }

    var cells = theobject.parentNode.getElementsByTagName("td");

    for (var i = 0; i < cells.length; i++) {
        if (cells[i] !== theobject) {
            cells[i].style.backgroundColor = "yellow"; //back to original color if this is not the current object
        };
    }


};

// comes from prototype.js; this is simply easier on the eyes and fingers
function $(id) { return document.getElementById(id); };