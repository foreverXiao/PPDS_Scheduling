


var PX = new (function () {
    var mouseOffset = null;



    // Demo 0 variables
    var blockMarginTop = 0;
    var blockMarginLeft = 0;
    var dragHelperparentNodeoffsetTop = 0;
    var dragHelperparentNodeoffsetLeft = 0;

    var dragHelper = null;

    var ganttZoneleft = 0;
    var ganttZonetop = 0;
    //for production line menu event
    var targetClone = null;
    //scalar div
    var scalar1 = null;
    var bckgrndImage1 = null;
    var pixelsPerDay = 90;

    var currentHorizontalScrollPos = 0;
    var currentVerticalScrollPos = 0;
    //Number.prototype.NaN0 = function () { return isNaN(this) ? 0 : this; };


    

    //AJAX get order line details from ASPX database service
    function InteractionWithDatabase1(elementID1, actioncode) {

        if (navigator.userAgent.toLowerCase().indexOf('chrome') > -1) { //if this is a chrome explorer
            actioncode += "&navigator=chrome";
        };

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
        xmlhttp.open("POST", "DtExchgChecker.aspx", false);
        xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        xmlhttp.send(actioncode);
        //xmlhttp.send(); 

    };


    function UsingASPdotNet(elementID1, actioncode) {

        //if (navigator.userAgent.toLowerCase().indexOf('chrome') > -1) { //if this is a chrome explorer
        //    actioncode += "&navigator=chrome";
        //};



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
        xmlhttp.open("POST", "DtExchgChecker.aspx", false);
        //xmlhttp.open("GET", "DataExchange.aspx", true);
        xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        xmlhttp.send(actioncode);
        //$(elementID1).innerHTML = xmlhttp.responseText;

    };



    //show all the Gantt chart for orders per production line
    this.dsplyOrderByLine = function dsplyOrderByLine(theobject) {

        if ($('ganttZone').getAttribute('crntLine') != theobject.getAttribute('id').toString()) {

            $('ganttZone').setAttribute('crntLine', theobject.getAttribute("id").toString());
            InteractionWithDatabase1('ganttZone', 'action=orderlines&lineno=' + theobject.getAttribute("id").toString() + '');
            window.status = "show orders by production line " + theobject.getAttribute("id").toString();
            theobject.style.backgroundColor = "white";

            //force to scroll to the left 
            window.scrollBy(0, -1);

            window.scrollTo(currentHorizontalScrollPos, currentVerticalScrollPos);
        }

        var cells = theobject.parentNode.getElementsByTagName("td");

        for (var i = cells.length - 1; i > -1; i--) {
            if (cells[i] !== theobject) {
                cells[i].style.backgroundColor = "#9999CC"; //back to original color if this is not the current object
            };
        }

    };

    // comes from prototype.js; this is simply easier on the eyes and fingers
    function $(id) { return document.getElementById(id); };

    


    //ignore vertical scroll operation
    function getScrollXY() {
        var scrOfX = 0;
        var scrOfY = 0;
        if (typeof (window.pageYOffset) == 'number') {
            //Netscape compliant
            scrOfY = window.pageYOffset;
            scrOfX = window.pageXOffset;
        } else if (document.documentElement.scrollLeft) { // document.documentElement && || document.documentElement.scrollTop)) {
            //IE6 standards compliant mode
            scrOfY = document.documentElement.scrollTop;
            scrOfX = document.documentElement.scrollLeft;
        } else if (document.body.scrollLeft) {  //document.body &&  || document.body.scrollTop)) {
            //DOM compliant
            scrOfY = document.body.scrollTop;
            scrOfX = document.body.scrollLeft;
        }

        return { x: scrOfX, y: scrOfY };
    };

    this.onload = function () {


        // Create our helper object that will show the item while dragging
        dragHelper = document.createElement('div');
        dragHelper.style.cssText = "position:absolute;display:none;";


        InteractionWithDatabase1('lines', 'action=productionlines'); // fill up div with id 'lines'

        UsingASPdotNet('scalar', 'action=scalar'); // fill up div with id 'scalar'


        var ganttChartArea = $('ganttZone');
        ganttChartArea.className = "DragContainer";


        UsingASPdotNet('bckgrndImg', 'action=bckgrndImg');

        //Fill container area with first available production line
        document.getElementsByTagName('td')[0].onclick();
        //ganttChartArea.getElementsByTagName('td')[0].onmousemove();


        //make both scalar and fixedscalar's left equal to element ganttChartArea's left
        $('scalar').style.left = ganttChartArea.style.left;
        $('scalar').style.width = $('scalar').childNodes.length * pixelsPerDay + "px";
        $('bckgrndImg').style.width = $('scalar').style.width;

        scalar1 = $('scalar');
        bckgrndImage1 = $('bckgrndImg');
        //get the number of pixels per day
        if (scalar1.childNodes.length > 1) {
            pixelsPerDay = parseInt(scalar1.childNodes[1].style.left) - parseInt(scalar1.childNodes[0].style.left)
        }



    };

    this.onscroll = function () {

        currentHorizontalScrollPos = getScrollXY().x + ganttZoneleft;
        currentVerticalScrollPos = getScrollXY().y + ganttZonetop;


        if (scalar1 == null) {
            scalar1 = $('scalar');
        };

        if (bckgrndImage1 == null) {
            bckgrndImage1 = $('bckgrndImg');
        };

        scalar1.style.left = -currentHorizontalScrollPos + 'px';
        bckgrndImage1.style.left = -currentHorizontalScrollPos + 'px';

    };


});




window.onscroll = PX.onscroll;

window.onload = PX.onload