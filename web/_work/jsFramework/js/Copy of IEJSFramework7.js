// JavaScript Document	
// JS Framework to be used for newzapp V5.
// Created by MH.


function object(original) {
	function F() {}
	F.prototype = original;
	return new F();
};

var nzObjects = {
	nzSoapArr: [], // Array for soap data nodes
	nzCount: 0, // Count for nzSoapArr array
	nzWebServiceMethodName: "", // The webservice method we are calling
	NoIDOnClick: function(){
		document.body.onmousedown = function(ele) {    
			var uiItem = eval("ele ? ele." + nzConstants.nzConstTarget + " : " + nzConstants.nzConstWindow + "." + nzConstants.nzConstEvent + "." + nzConstants.nzConstSrcElement);		
			doSomething(uiItem);
		};		
		document.body.onkeydown = function(ele) {
			var ek = event.keyCode; 
			if(ek == 13){// 13: return key
				var uiItem = eval("ele ? ele." + nzConstants.nzConstTarget + " : " + nzConstants.nzConstWindow + "." + nzConstants.nzConstEvent + "." + nzConstants.nzConstSrcElement);			
				doSomething(uiItem);
			}
		};		
	},
	nzAddSoapData: function(name,value,webServiceMethodName,cDataTag){
		this.nzWebServiceMethodName = webServiceMethodName;
		//alert("cDataTag: " + cDataTag);
		if(cDataTag != undefined){
			this.nzSoapArr[this.nzCount] = "<" + name + ">![CDATA[" + value + "]]</" + name + ">";
			//alert("cDataTag != undefined");
		}
		else{
			//alert("cDataTag == undefined");
			this.nzSoapArr[this.nzCount] = "<" + name + ">" + value + "</" + name + ">";
		}
		this.nzCount += 1;		
	},
	nzSoapEnvelope: function(SoapData){
		if(this.nzWebServiceMethodName != undefined){
			var ns = "http://www.newzapp.co.uk/";
			var SoapEnvelope = 
				/*"<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
				"<soap:Envelope " +
				"xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" " +
				"xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" " +
				"xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
				"<soap:Body><Request>" +
				//"<" + this.nzWebServiceMethodName + " xmlns=\"" + ns + "\">" +
				"<" + this.nzWebServiceMethodName + ">" +
				SoapData +
				"</" + this.nzWebServiceMethodName + "></Request></soap:Body></soap:Envelope>";*/
				//"<SOAP:Envelope xmlns:SOAP='urn:schemas-xmlsoap-org:soap.v1'>" + 
				"<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
				"<soap:Envelope " +
				"xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" " +
				"xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" " +
				"xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
				"<soap:Body>" +
				//"<" + this.nzWebServiceMethodName + ">" +
				"<" + this.nzWebServiceMethodName + " xmlns=\"" + ns + "\">" +
				SoapData +
				"</" + this.nzWebServiceMethodName + "></soap:Body></soap:Envelope>";
				this.nzWebServiceMethodName = ""; //Reset the webservice method name
		}
		else{
			var SoapEnvelope = 				
				"<SOAP:Envelope xmlns:SOAP='urn:schemas-xmlsoap-org:soap.v1'>" + 
				"<SOAP:Body><Request>" +
				SoapData +
				"</Request></SOAP:Body></SOAP:Envelope>";
		}
		return SoapEnvelope;
	},
	nzGetSoapData: function(){
		var SoapData = "";
		for(var i = 0;i<this.nzSoapArr.length;i++){
			SoapData += this.nzSoapArr[i];
		}
		this.nzCount = 0; // Reset the count when we are finished with nzSoapArr
		return this.nzSoapEnvelope(SoapData);
	}
};

//******************************************************************************
// We need to create the new nzObj. Then we can call the object by nzObj.MethodName.

var nzObj = object(nzObjects);

// *****************************************************************************
// Pop up box object //

nzPopUpBox = {
	nzcontainerID: "", // create empty ID
	nzcontainerContentID: "", // create empty ID
	nzdialogContentID: "", // create empty ID
	nzdialogTitleID: "", // create empty ID
	nziframeID: "", // create empty ID
	nzWidth: 400, // default width value
	nzHeight: 400, // default height value
	nzCreatePopUpHTML: function(){
		var strHTML = "";
		strHTML = strHTML+'<div id="'+this.nzcontainerID+'" class="overlayBox" style="display:none;">';
		strHTML = strHTML+'<div id="'+this.nzcontainerContentID+'" class="overlayContent">';
		strHTML = strHTML+'<div class="dialog_TitleBox">';
		strHTML = strHTML+'<img align="right" src="/images/icons/icon_close.gif" onmouseover="this.src=\'/images/icons/icon_closeovr.gif\'" onmouseout="this.src=\'/images/icons/icon_close.gif\'" onclick="document.getElementById(\''+this.nzcontainerID+'\').style.display=\'none\'; document.getElementById(\''+this.nziframeID+'\').style.display=\'none\'; document.documentElement.style.overflow  = \'\';">';
		strHTML = strHTML+'<div id="'+this.nzdialogTitleID+'" class="dialog_Title">NewZapp';
		strHTML = strHTML+'</div>';
		strHTML = strHTML+'</div>';		
		strHTML = strHTML+'<div id="'+this.nzdialogContentID+'" class="dialog_ContentBox"></div>';
		
		if(this.nzcontainerID.toLowerCase() == "loadingbox"){
			strHTML = strHTML+'<div class="dialog_ContentBox"><img src="/images/icons/loadingbaranim.gif"></div>';	
		}
		
		strHTML = strHTML+'</div>';
		strHTML = strHTML+'<iframe id="'+this.nziframeID+'" class="overlayFrame" height="100%" width="100%" src="/editsite/blank.asp"></iframe>';
		strHTML = strHTML+'</div>';
		
		document.write(strHTML);
		this.nzSetSize(this.nzWidth,this.nzHeight);// we then set the height and width of the container to the default height and width values.		
	},
	nzSetTitle: function(t){
		$(""+this.nzdialogTitleID+"").nzHTML(t);
		return this;
	},
	nzSetSize: function(w,h){
		this.nzWidth = w; // set the width
		this.nzHeight = h; // set the height
		return this; // return this in order to chain.
	},
	nzAssignID: function(){
		this.nzcontainerContentID = "content_" + this.nzcontainerID;
		this.nzdialogContentID = "dialogContent_" + this.nzcontainerID; 
		this.nzdialogTitleID = "dialogTitle_" + this.nzcontainerID; 
		this.nziframeID = "iframe_" + this.nzcontainerID;
	},
	nzCenterDialogBox: function(){
		var s = document.getElementById(""+this.nzcontainerContentID+"").style;
		s.left = (document.documentElement.offsetWidth/2) - (parseFloat(this.nzWidth)/2);
		s.top = (document.documentElement.clientHeight /2)  - (parseFloat(this.nzHeight)/2);	
	},
	nzSetContent: function(strMessage){
		//document.getElementById(""+this.nzdialogContentID+"").innerHTML = strMessage;
		$(""+this.nzdialogContentID+"").nzHTML(strMessage);
		return this; // return this in order to chain.
	},
	nzShowDialog: function(){
		this.nzCenterDialogBox();	
		eval(nzConstants.nzConstDocument + "." + nzConstants.nzConstDocumentElement + "." + nzConstants.nzConstStyle + "." + nzConstants.nzConstOverflow + " = " + "'hidden'");		
		$(""+this.nziframeID+"",""+this.nzcontainerID+"").toggleDisplay();
		return this; // return this in order to chain.
	},
	nzHideDialog: function(){	
		eval(nzConstants.nzConstDocument + "." + nzConstants.nzConstDocumentElement + "." + nzConstants.nzConstStyle + "." + nzConstants.nzConstOverflow + " = " + "''");
		$(""+this.nziframeID+"",""+this.nzcontainerID+"").toggleDisplay();
		return this; // return this in order to chain.
	},
	nzPopUpType: function(name){
		this.nzcontainerID = name; 
		this.nzAssignID(); // we will populate the empty ID's here when we set up the object.
		this.nzCreatePopUpHTML(); // we will create the pop up.
		return this; // return this in order to chain.
	}
};

// End of pop up box object
// *****************************************************************************

// ***** These are external properties that will be used within the nzElementContent object *****//
// These will be used as the global variables that handles the element count, and position. //

nzElementContentProps = {
	nzEleConPropsCount: 0, // This will count the number of elements created.
	nzEleConZIndex: 0, // this will count the zIndex
	/*nzEleConStartX: 0, // this will position the x element
	nzEleConStartY: 0, // this will position the y element
	nzEleConHeight: 0, // this is the height of the current element
	nzEleConWidth: 0, // this is the width of the current element*/
	nzEleTitleSelected: null, // Will display which element has been selected.
	nzEleContentsSelected: null, // Will display which element has been selected.
	nzEleContainerMode: "Normal" // if the container mode is normal, cascade or tiling. Default is normal
};

// ***** Main content window object ***** //

var nzElementContainer = {
	
	//***** Start of property values to be used by title and content element ***** //
	nzEleConCount: 0, // default element count
	nzEleConMainConID: "", // create emmpty Main Container ID
	nzEleConContentID: "", // create empty content ID
	nzEleConTitleID: "", // create empty title ID
	nzEleConTitleHeight: 0, // default Title height value
	nzEleConTitleWidth: 0, // default Title width value
	nzEleConContentHeight: 0, // default content height value
	nzEleConContentWidth: 0, // default content width value
	nzEleConStyleLeft: 0, // default style left value
	nzEleConStyleTop: 0, // default style top value
	nzEleConMaxHeight: 0, // default max height value. This will be set when the object is created
	nzEleConMaxWidth: 0, // default max width value. This will be set when the object is created
	nzEleConMin: false, // default container minimized option
	nzEleConMax: false, // default container maximized option
	nzEleConPin: false, // default container pin option
	nzEleConClose: false, // default container close option
	nzEleConResize: false, // default container resize option
	nzEleConTitleClassName: "divDragTitle", // default title class name
	nzEleConContentClassName: "divDragContent", // default content class name
	//nzEleConZIndex: 0, // default zIndex value
	nzEleConStartX: 0, // default x pos
	nzEleConStartY: 0, // default y pos
	nzEleConTitleCaption: "", // default title caption
	nzEleConSelected: false,
	nzEleConContentSelected: false,
	nzEleConTitleMode: "Normal",
	nzEleConTitleBarVisibility: true, // default title bar visibility value.
	nzEleCondragMinBtn: "",	
	nzEleCondragMaxBtn: "",	
	nzEleCondragPinBtn: "",	
	nzEleCondragCloseBtn: "",
	nzEleConDragObjTitle: null,
	nzEleConDragOffSetX: 0,
	nzEleConDragOffSetY: 0,
	nzEleConContents: "",
	// What other properties will we need here?
	
	//***** End of property values *****//
	
	//***** Start of methods *****//
	nzAddPx: function(num){
		return String(num) + "px";
	},
	nzFindParentTagById: function(obj,parentname){
		while (obj) {
			if (obj.id.match(parentname)) {
				return obj;
			}
			
			if (obj.parentElement) {
				obj = obj.parentElement;
			}
			else {
				return null;
			}
		}
		return null;
	},
	nzfindParentDiv: function(obj){
		while (obj) {
			if (obj.tagName.toUpperCase() == "DIV") {
				return obj;
			}
			
			if (obj.parentElement) {
				obj = obj.parentElement;
			}
			else {
				return null;
			}
		}
		return null;
	},
	nzEleConMouseUp: function(e){
		if (this.nzEleConDragObjTitle) {
			//finishDrag(e);
			this.nzEleConFinishDrag(e);
		}
	},
	nzEleConMouseMove: function(e){
		// If not null, then we're in a drag
		if (this.nzEleConDragObjTitle) {
		
			if (!e.preventDefault) {
				// This is the IE version for handling a strange
				// problem when you quickly move the mouse
				// out of the window and let go of the button.
				if (e.button == 0) {
					//finishDrag(e);
					this.nzEleConFinishDrag(e);
					return;
				}
			}
		
			this.nzEleConDragObjTitle.style.left = this.nzAddPx(e.clientX - this.nzEleConDragOffSetX);
			this.nzEleConDragObjTitle.style.top = this.nzAddPx(e.clientY - this.nzEleConDragOffSetY);
			this.nzEleConDragObjTitle.content.style.left = this.nzAddPx(e.clientX - this.nzEleConDragOffSetX);
			this.nzEleConDragObjTitle.content.style.top = this.nzAddPx(e.clientY - this.nzEleConDragOffSetY + 20);
			if (e.preventDefault) {
				e.preventDefault();
			}
			else {
				e.cancelBubble = true;
				return false;
			}
		}
	},
	nzEleConMouseDown: function(e){
		//alert("Count: " + nzElementContentProps.nzEleConZIndex);
		// These first two lines are written to handle both FF and IE
		var curElem = e.srcElement || e.target;
		var dragTitle = e.currentTarget || this.nzfindParentDiv(curElem);
		if (dragTitle) {
			if (dragTitle.className != 'divDragTitle') {
				return;
			}
		}
		//document.getElementById("Count").innerHTML = "curElem: " + curElem + "<br>dragTitle: " + dragTitle;
	
		// Start the drag, but first make sure neither is null
		if (curElem && dragTitle) {
		
			// Attach the document handlers. We don't want these running all the time.
			//addDocumentHandlers(true);
			this.nzEleConAddDocumentHandlers(true);
			
			// Move this window to the front.
			nzElementContentProps.nzEleConZIndex ++;
			//document.getElementById("div_innerHTML").innerHTML = "<br>topZ: " + nzElementContentProps.nzEleConZIndex;
			dragTitle.style.zIndex = nzElementContentProps.nzEleConZIndex;
			dragTitle.content.style.zIndex = nzElementContentProps.nzEleConZIndex;
			// store the selected element
			nzElementContentProps.nzEleTitleSelected = dragTitle.id;
			nzElementContentProps.nzEleContentsSelected = dragTitle.content.id;
			
			//nzElementContainer.nzEleConZIndex = nzElementContentProps.nzEleConZIndex;		
					
			 // We then need to change the class of the element as we know it has been selected.		
			//dragTitle.className =  "divElementSelected";
		
			// Check if it's the button. If so, don't drag.
			if (curElem.className != "divTitleButton") {
			   
				// Save away the two objects
				this.nzEleConDragObjTitle = dragTitle;
				
				// Calculate the offset
				this.nzEleConDragOffSetX = e.clientX - 
					dragTitle.offsetLeft;
				this.nzEleConDragOffSetY = e.clientY - 
					dragTitle.offsetTop;
					
				// Don't let the default actions take place
				if (e.preventDefault) {
					e.preventDefault();
				}
				else {
					document.onselectstart = function () { return false; };
					e.cancelBubble = true;
					return false;
				}
			}
		}
	},
	nzEleConFinishDrag: function(e){
		var finalX = e.clientX - this.nzEleConDragOffSetX;
		var finalY = e.clientY - this.nzEleConDragOffSetY;
		if (finalX < 0) { finalX = 0 }; // If the element is dragged outside the window to the left, it will revert the X position back to 0
		if (finalY < 0) { finalY = 0 }; // If the element is dragged outside the window to the top, it will revert the Y position back to 0	 	
		
		// Add finalX to offsetWidth to get right side of element.
		// If right side of element is greater then maxWidth, then restore finalX to (maxwidth - offsetWidth)
		var nEleRightPos = (finalX + this.nzEleConDragObjTitle.offsetWidth);
		
		if(nEleRightPos > this.nzEleConDragObjTitle.MaxWidth){
			finalX = (this.nzEleConDragObjTitle.MaxWidth - this.nzEleConDragObjTitle.offsetWidth);
		}
		
		var nEleBottomPos = (finalY + (this.nzEleConDragObjTitle.content.offsetHeight + this.nzEleConDragObjTitle.offsetHeight) - 2);
		
		var nEleTotalHeight = (this.nzEleConDragObjTitle.MaxHeight + this.nzEleConDragObjTitle.content.MaxHeight);
		var nEleTotalOffSetHeight = (this.nzEleConDragObjTitle.offsetHeight + this.nzEleConDragObjTitle.content.offsetHeight) - 2;
		
		if(nEleBottomPos > nEleTotalHeight){
			finalY = (nEleTotalHeight - nEleTotalOffSetHeight);
		}
		
		//document.getElementById("div_innerHTML").innerHTML = "nEleTotalOffSetHeight: " + nEleTotalOffSetHeight + "<br>nEleTotalHeight: " + nEleTotalHeight + "<br>finalY: " + finalY + "<br>nEleBottomPos: " + nEleBottomPos;
		
		this.nzEleConDragObjTitle.style.left = this.nzAddPx(finalX);
		this.nzEleConDragObjTitle.style.top = this.nzAddPx(finalY);
		this.nzEleConDragObjTitle.content.style.left = this.nzAddPx(finalX);
		this.nzEleConDragObjTitle.content.style.top = this.nzAddPx(finalY + 20);
		
		// Done, so reset to null
		this.nzEleConDragObjTitle = null;
		//addDocumentHandlers(false);
		this.nzEleConAddDocumentHandlers(false);
		
		if (e.preventDefault) {
			e.preventDefault();
		}
		else {
			document.onselectstart = null;
			e.cancelBubble = true;
			return false;
		}
	},
	nzEleConContentMouseDown: function(e){
		//alert("inside contentMouseDown");
		// Move the window to the front
		// Use a handy trick for IE vs FF
		var dragContent = e.srcElement || e.currentTarget;
		if ( ! dragContent.id.match("dragContent")) {
			dragContent = this.nzFindParentTagById(dragContent, "dragContent");
			//dragContent = findParentTagById(dragContent, "dragContent");

		}
		if (dragContent) {
			nzElementContentProps.nzEleConZIndex ++;
			dragContent.style.zIndex = nzElementContentProps.nzEleConZIndex;
			dragContent.titlediv.style.zIndex = nzElementContentProps.nzEleConZIndex;
			
			// store the selected element
			nzElementContentProps.nzEleTitleSelected = dragContent.titlediv.id;
			nzElementContentProps.nzEleContentsSelected = dragContent.id;
			
		}
	},
	nzEleConAddDocumentHandlers: function(addOrRemove){
		 if (addOrRemove) {        
			// IE
			//document.onmousedown = function() { mouseDown(window.event) } ;
			//document.onmousemove = function() { mouseMove(window.event) } ;
			// document.onmouseup = function() { mouseUp(window.event) } ;
			document.onmousedown = function() { nzElementContainer.nzEleConMouseDown(window.event) } ;		
			document.onmousemove = function() { nzElementContainer.nzEleConMouseMove(window.event) } ;			   
			document.onmouseup = function() { nzElementContainer.nzEleConMouseUp(window.event) } ;
		}
		else {
			// IE
			// Be careful here. If you have other code that sets these events,
			// you'll want this code here to restore the values to your other handlers,
			// rather than just clear them out.
			document.onmousedown = null;
			document.onmousemove = null;
			document.onmouseup = null;        
		}
	},	
	// Set content properties //
	nzEleConSetContentProps: function(obj){
		obj.id = this.nzEleConContentID;
		obj.className = this.nzEleConContentClassName;
		obj.style.height = this.nzEleConContentHeight;
		obj.style.width = this.nzEleConContentWidth;
		obj.style.left = this.nzEleAddPx(this.nzEleConStartX);
		obj.style.top = this.nzEleAddPx(this.nzEleConStartY + this.nzEleConTitleHeight);
		obj.style.zIndex = nzElementContentProps.nzEleConZIndex;
		//obj.style.overflow = "auto";
		obj.MaxHeight = (this.nzEleConMaxHeight - this.nzEleConTitleHeight); // We need to take the title height away from the content height.
		obj.MaxWidth = this.nzEleConMaxWidth;
		obj.Selected = this.nzEleConSelected; // If the container has been selected or not
		
		// Maximize Props
		obj.MaximizeCurrentHeight = this.nzEleConContentHeight;
		obj.MaximizeCurrentWidth = this.nzEleConContentWidth;
		obj.MaximizeCurrentTop = this.nzEleConStartY;
		obj.MaximizeCurrentLeft = this.nzEleConStartX;
		
		// Minimize Props
		obj.MinimizeCurrentHeight = this.nzEleConTitleHeight;
		obj.MinimizeCurrentWidth = this.nzEleConTitleWidth;
		obj.MinimizeCurrentTop = this.nzEleConStartY;
		obj.MinimizeCurrentLeft = this.nzEleConStartX;
	},
	// Set title properties //
	nzEleConSetTitleProps: function(obj){
		obj.id = this.nzEleConTitleID;
		obj.className = this.nzEleConTitleClassName;
		obj.style.height = this.nzEleConTitleHeight;
		obj.style.width = this.nzEleConTitleWidth;
		obj.style.left = this.nzEleAddPx(this.nzEleConStartX);
		obj.style.top = this.nzEleAddPx(this.nzEleConStartY);
		obj.style.zIndex = nzElementContentProps.nzEleConZIndex;		
		obj.Min = this.nzEleConMin;
		obj.Max = this.nzEleConMax;
		obj.Pin = this.nzEleConPin;
		obj.Close = this.nzEleConClose;
		obj.Resize = this.nzEleConResize;
		obj.MaxHeight = this.nzEleConTitleHeight; // Max title height should always be 20px.
		obj.MaxWidth = this.nzEleConMaxWidth;
		obj.Selected = this.nzEleConSelected; // If the container has been selected or not
		
		// Maximize Props
		obj.MaximizeCurrentHeight = this.nzEleConTitleHeight;
		obj.MaximizeCurrentWidth = this.nzEleConTitleWidth;
		obj.MaximizeCurrentTop = this.nzEleConStartY;
		obj.MaximizeCurrentLeft = this.nzEleConStartX;
		
		// Minimize Props
		obj.MinimizeCurrentHeight = this.nzEleConTitleHeight;
		obj.MinimizeCurrentWidth = this.nzEleConTitleWidth;
		obj.MinimizeCurrentTop = this.nzEleConStartY;
		obj.MinimizeCurrentLeft = this.nzEleConStartX;
		
	},
	nzEleConCreateElement: function(){
		// This is where we build the content window.	
		// We create outer layer div to wrap the title and content in
		//var nzEleOuterContainer = document.createElement("div");                     
		//nzEleOuterContainer.id = "divContainer_" + this.nzEleConCount;
		//nzEleOuterContainer.id = this.nzEleConMainConID;                                           
		//document.getElementById(this.nzEleConMainConID).appendChild(nzEleOuterContainer);
		//topZ++
		nzElementContentProps.nzEleConZIndex ++;
		//alert("ZIndex Count: " + nzElementContentProps.nzEleConZIndex);
		//this.nzEleConZIndex = nzElementContentProps.nzEleConZIndex;
		//document.getElementById("div_innerHTML").innerHTML = this.nzEleConZIndex;
		//document.getElementById("div_innerHTML").innerHTML = "<br>Props Count: " + nzElementContentProps.nzEleConZIndex;
		var nzEleContent = document.createElement("div");		
		this.nzEleConSetContentProps(nzEleContent);
		nzEleContent.innerHTML = this.nzEleConContents;
		//nzEleContent.innerHTML = '<table cellspacing="0" cellpadding="0"><tr><td colspan="2">This is the content of the element.</td></tr></table>';
		
		//if (canMove) {
			//nzEleContent.attachEvent("onmousedown", function(e) { return contentMouseDown(e) });			
			nzEleContent.attachEvent("onmousedown", function(e) { return nzElementContainer.nzEleConContentMouseDown(e) });
			
		//}

		this.nzEleConMainConID.appendChild(nzEleContent);
		
		//Then we create the title div
		var nzEleTitle = document.createElement("div"); 
		this.nzEleConSetTitleProps(nzEleTitle);
		
		// When we first create the element, we will set the default title bar options.
		this.nzEleConDefaultTitleBarBtn(nzEleTitle);
		
		// Then we need to add the action listeners to the title bar.	
		// If canMove is false, don't register event handlers
		//if (canMove) {
			// IE
			//nzEleTitle.attachEvent("onmousemove", function(e) { return mouseMove(e) });
			nzEleTitle.attachEvent("onmousemove", function(e) { return nzElementContainer.nzEleConMouseMove(e) });			
			//nzEleTitle.attachEvent("onmousedown", function(e) { return mouseDown(e) });
			nzEleTitle.attachEvent("onmousedown", function(e) { return nzElementContainer.nzEleConMouseDown(e) });
			
			//nzEleTitle.attachEvent("onmouseup", function(e) { return mouseUp(e) });
			nzEleTitle.attachEvent("onmouseup", function(e) { return nzElementContainer.nzEleConMouseUp(e) });
			//nzEleTitle.attachEvent("onmouseup", nzConObj.nzOnMouseUp(e));
			
			//nzEleTitle.attachEvent("onmousemove", function(e) { return nzConObj.nzOnMouseMove(e) });
			//nzEleTitle.attachEvent("onmousedown", function(e) { return nzConObj.nzOnMouseDown(e) });
			//nzEleTitle.attachEvent("onmouseup", function(e) { return nzConObj.nzOnMouseUp(e) });
						
			this.nzEleConMainConID.appendChild(nzEleTitle);
		//alert("nzEleTitle1: " + nzEleTitle);
		//}
		
		// Need to add the event handlers below to resize the container.
		//nzEleContent.attachEvent("onmousemove", function(e) { return ResizeMouseMove(e) });
		//nzEleContent.attachEvent("onmousedown", function(e) { return ResizeMouseDown(e) });
		//nzEleContent.attachEvent("onmouseup", function(e) { return ResizeMouseUp(e) });
		//this.nzEleConResizeContainer(nzEleContent);
		// Save away the content DIV into the title DIV for 
		// later access, and vice versa
		nzEleTitle.content = nzEleContent;
		nzEleContent.titlediv = nzEleTitle;
		//this.nzEleSetSize();

		// store the selected element
		nzElementContentProps.nzEleTitleSelected = nzEleTitle.id;
		nzElementContentProps.nzEleContentsSelected = nzEleContent.id;
	},
	nzEleConSetTitleBarBtn: function(){
		// We need to find out which buttons the user wants enabled in the title bar.
		// Because the user has specified this function, we will reset all of the title bar values first to get a clean slate.
		this.nzEleConMin = false;
		this.nzEleConMax = false;
		this.nzEleConPin = false;
		this.nzEleConClose = false;
		
		var nTitleBarBtnCount = 0;
		var EleTitle = document.getElementById(this.nzEleConTitleID);
	
		for(var i=0;i<arguments.length;i++){
			switch(arguments[i].toLowerCase()){
				case "min":
					this.nzEleConMin = true;
					EleTitle.Min = this.nzEleConMin;
					nTitleBarBtnCount ++;
					break;
				case "max":
					this.nzEleConMax = true;
					EleTitle.Max = this.nzEleConMax;
					nTitleBarBtnCount ++;
					break;
				case "close":
					this.nzEleConClose = true;
					EleTitle.Close = this.nzEleConClose;
					nTitleBarBtnCount ++;
					break;
				case "pin":
					this.nzEleConPin = true;
					EleTitle.Pin = this.nzEleConPin;
					nTitleBarBtnCount ++;
					break;	
			}
		}
		
		if(nTitleBarBtnCount == 0){
			// No buttons were selected so we will default the user to have 3 buttons enabled - min, max and close.
			this.nzEleConDefaultTitleBarBtn(EleTitle);
		}
		else{
			//They have specified at least 1 button. In here we could also have the close button as default as well.
			this.nzEleConCreateTitleBar(EleTitle);	
		}		
		return this; // return this in order to chain.
	},
	nzEleConDefaultTitleBarBtn: function(ele){
		// Below are the 3 buttons enabled at default if the user does not specify which buttons they want in the title bar.
		this.nzEleConMin = true;
		this.nzEleConMax = true;
		this.nzEleConClose = true;
		
		this.nzEleConCreateTitleBar(ele);
	},
	nzEleConCreateTitleBar: function(ele){
		// This function will create the title bar for each container. The buttons will be in a specific order.
		// The buttons will always be in the following order.
		// 1: Min, 2: Max, 3: Pin, 4: Close
		var strHTML = '';
		strHTML += "<table><tr><td ondblclick='nzElementContainer.nzEleConMaximizeBtn(" + this.nzEleConTitleID + "," + this.nzEleConContentID + "," + this.nzEleCondragMaxBtn + ");'>" + this.nzEleConTitleCaption + "</td><td style='text-align:right;'>"; // start of title bar.
		if(this.nzEleConMin == true){
			strHTML += '<img src="buttontop.gif" class="divTitleButton" title="Minimize" id="' + this.nzEleCondragMinBtn + '" onmousedown="javascript:nzElementContainer.nzEleConMinimizeBtn(' + this.nzEleConTitleID + "," + this.nzEleConContentID + "," + this.nzEleCondragMaxBtn + ');" />'
		}
		if(this.nzEleConMax == true){
			strHTML += '<img src="/images/iconCross.gif" class="divTitleButton" title="Maximize" id="' + this.nzEleCondragMaxBtn + '" onmousedown="javascript:nzElementContainer.nzEleConMaximizeBtn(' + this.nzEleConTitleID + "," + this.nzEleConContentID + "," + this.nzEleCondragMaxBtn + ');" />'
		}
		if(this.nzEleConPin == true){
			strHTML += '<img src="/images/iconCross.gif" class="divTitleButton" title="Pin" id="' + this.nzEleCondragPinBtn + '" onmousedown="javascript:nzElementContainer.nzEleConPinBtn(' + this.nzEleConCount +');" />'
		}
		if(this.nzEleConClose == true){
			// For some reason passing through this.nzEleConMainConID.id instead of this.nzEleConMainConID only seems to work otherwise a JS error comes up saying "Expected ']'".
			strHTML += '<img src="/images/iconCross.gif" class="divTitleButton" title="Close" id="' + this.nzEleCondragCloseBtn + '" onmousedown="javascript:nzElementContainer.nzEleConCloseBtn(' + this.nzEleConTitleID + "," + this.nzEleConContentID + "," + this.nzEleConMainConID.id + ');" />'
		}
				
		strHTML += '</td></tr></table>';// end of title bar	
		ele.innerHTML = strHTML;		
	},
	nzEleConResizeContainer: function(ele){
		//var ele = document.getElementById("dragContent_Cont" + String(id));
		var eleHeight = ele.style.height;
		var eleWidth = ele.style.width;
		//ele.onmousedown = window.event.clientY;
		document.getElementById("div_innerHTML").innerHTML = window.event.clientY;
		if(window.event.clientY == eleHeight){
			alert("height: " + ele.style.height);
		}
	},
	nzEleConPinBtn: function(id){
		//if(document.getElementById("dragTitle_Cont" + id).nzEleConTitleMode == "Pin"){
			// This will add the draggable handler for the element.
			// In this function we will allow the user to resize the content container if the element is not pinned.
			this.nzEleConResize = true; // Allow the element to be resizable
			//addDocumentHandlers(true);
			
			
		//}
		/*else{*/
			
			// Otherwise we pin the element.
			//document.getElementById("dragTitle_Cont" + id).nzEleConTitleMode = "Pin";
			//var nzEleTitle = document.getElementById("dragTitle_Cont" + id);
			//this.nzEleConResize = false; // Disallow the element to be resizable

		//}
	},
	nzEleConCloseBtn: function(objTitle,objContent,objMainCont){		
		// This function will delete the content window container
		var ChildContentEle = objContent;
		var ChildTitleEle = objTitle;
		var ParentEle = objMainCont;
		ParentEle.removeChild(ChildContentEle);
		ParentEle.removeChild(ChildTitleEle);
		
		nzElementContentProps.nzEleConPropsCount --; // we decrement the external element count property
		this.nzEleConCount = nzElementContentProps.nzEleConPropsCount; // we then assign that property to the main element content count property			
	},
	nzEleConMaximizeBtn: function(objTitle,objContent,objBtnID){
		//alert(objTitle);
		//alert(objContent);
		//alert(objBtnID);
		if(objTitle.nzEleConTitleMode == "Max"){
			// We need to restore the container back to its previous size.
			objTitle.nzEleConTitleMode = "MaxRestore";
			var TitleEle = objTitle;
			var ContentEle = objContent;
			var img = objBtnID;
			
			// Now we want to restore the title and content areas original size and positioning before it was minimized.
			// title
			TitleEle.style.left = TitleEle.MaximizeCurrentLeft;
			TitleEle.style.top = TitleEle.MaximizeCurrentTop;
			TitleEle.style.width = TitleEle.MaximizeCurrentWidth;
			TitleEle.style.height = TitleEle.MaximizeCurrentHeight;
			// content
			ContentEle.style.height = ContentEle.MaximizeCurrentHeight;
			ContentEle.style.width = ContentEle.MaximizeCurrentWidth;	
			ContentEle.style.left = ContentEle.MaximizeCurrentLeft;		
			ContentEle.style.top = ContentEle.MaximizeCurrentTop;
			
			if(img != null){
				img.title = "Maximize"
				img.src = "/images/iconCross.gif";
			}
		}
		else{
			// Otherwise we maximize the container.
			objTitle.nzEleConTitleMode = "Max";
			
			var TitleEle = objTitle;
			var ContentEle = objContent;
			var img = objBtnID;
			//alert("img: " + img);
			// First we need to store the current left, top, width and height for the title and content area.			
			// title
			var TitleEleLeft = TitleEle.style.left;
			var TitleEleTop = TitleEle.style.top;
			var TitleEleWidth = TitleEle.style.width;
			var TitleEleHeight = TitleEle.style.height;
			// content
			var ContentEleLeft = ContentEle.style.left;
			var ContentEleTop = ContentEle.style.top;
			var ContentEleWidth = ContentEle.style.width;
			var ContentEleHeight = ContentEle.style.height;
			
			// Then we need to save those settings into the title and content area current properties.
			// title
			TitleEle.MaximizeCurrentHeight = TitleEleHeight;
			TitleEle.MaximizeCurrentWidth = TitleEleWidth;	 	
			TitleEle.MaximizeCurrentLeft = TitleEleLeft;			
			TitleEle.MaximizeCurrentTop = TitleEleTop;
			// content
			ContentEle.MaximizeCurrentHeight = ContentEleHeight;
			ContentEle.MaximizeCurrentWidth = ContentEleWidth;	 	
			ContentEle.MaximizeCurrentLeft	 = ContentEleLeft;			
			ContentEle.MaximizeCurrentTop = ContentEleTop;
			
			// Now we need to set the new height and width.	
			//alert("Content Width: " + ContentEle.MaxWidth);	
			//alert("Content Height: " + ContentEle.MaxHeight);	
			TitleEle.style.width = (TitleEle.MaxWidth - 2);
			ContentEle.style.width = (ContentEle.MaxWidth - 2);
			ContentEle.style.height = (ContentEle.MaxHeight - 2);
			
			TitleEle.style.left = 0;
			ContentEle.style.left = 0;
			TitleEle.style.top = 0;
			ContentEle.style.top = TitleEle.style.height;
			
			if(img != null){
				// Change the button's property settings
				img.title = "Restore Container"
				img.src = "buttonbottom.gif";
			}
		}	
	},
	nzEleConMinimizeBtn: function(objTitle,objContent,objBtnID){
		// This function will minimize the content window and just show the title bar.
		if(objTitle.nzEleConTitleMode == "Min"){
			// Then we need to restore the container back to its previous size.
			objTitle.nzEleConTitleMode = "MinRestore";
			var TitleEle = objTitle;
			var ContentEle = objContent;
			var img = objBtnID;
			
			// Now we want to restore the title and content areas original size and positioning before it was minimized.
			// title
			TitleEle.style.left = TitleEle.MinimizeCurrentLeft;
			TitleEle.style.top = TitleEle.MinimizeCurrentTop;
			TitleEle.style.width = TitleEle.MinimizeCurrentWidth;
			TitleEle.style.height = TitleEle.MinimizeCurrentHeight;
			// content
			ContentEle.style.height = ContentEle.MinimizeCurrentHeight;
			ContentEle.style.width = ContentEle.MinimizeCurrentWidth;	
			ContentEle.style.left = ContentEle.MinimizeCurrentLeft;			
			ContentEle.style.top = ContentEle.MinimizeCurrentTop;					
		
			img.title = "Minimize"
			img.src = "buttontop.gif";
			ContentEle.style.display = "block";
		}
		else{
			
			// Otherwise we minimize the container.
			objTitle.nzEleConTitleMode = "Min";
			
			var TitleEle = objTitle;
			var ContentEle = objContent;
			var img = objBtnID;
			
			// First we need to store the current left, top, width and height for the title and content area.			
			// title
			var TitleEleLeft = TitleEle.style.left;
			var TitleEleTop = TitleEle.style.top;
			var TitleEleWidth = TitleEle.style.width;
			var TitleEleHeight = TitleEle.style.height;
			// content
			var ContentEleLeft = ContentEle.style.left;
			var ContentEleTop = ContentEle.style.top;
			var ContentEleWidth = ContentEle.style.width;
			var ContentEleHeight = ContentEle.style.height;
			
			// Then we need to save those settings into the title and content area current properties.
			// title
			TitleEle.MinimizeCurrentHeight = TitleEleHeight;
			TitleEle.MinimizeCurrentWidth = TitleEleWidth;	 	
			TitleEle.MinimizeCurrentLeft = TitleEleLeft;			
			TitleEle.MinimizeCurrentTop = TitleEleTop;
			// content
			ContentEle.MinimizeCurrentHeight = ContentEleHeight;
			ContentEle.MinimizeCurrentWidth = ContentEleWidth;	 	
			ContentEle.MinimizeCurrentLeft = ContentEleLeft;			
			ContentEle.MinimizeCurrentTop = ContentEleTop;			
			
			// Then we need to set the new settings for the minimized title bar.
			//if(id == 1){
				// nzEleConStartX should be 0. If its not something has gone wrong at this stage.
				//alert("start X: " + this.nzEleConStartX);
				//TitleEle.style.left = this.nzEleConStartX;
				//TitleEle.style.left = 20
			//}
			//else{
				// still need to sort this bit out.
				//TitleEle.style.left = (TitleEle.style.width.replace("px","") + TitleEle.style.width.replace("px",""));
			//}
			
			// we get the current element count.
			var nCount = nzElementContentProps.nzEleConPropsCount;
			
			// Then we get the width of the main container.
			//var nContWidth = document.getElementById("div_MainCont").style.height;
			var nContWidth = objTitle.MaxWidth;
			// then we divide the main container by the element count to see how much space we can occupy for each container.			
			var nElementWidth = (nContWidth / nCount);
			
			// We then check the containers width against the nElementWidth to see if it will fit along the bottom row.
			//if(nElementWidth
			
			// If nElementWidth is greater then 200px then we will move the minimized container onto the next level of minimized containers.
			//if(nElementWidth
			
			TitleEle.style.left = 0;
			TitleEle.style.top = (ContentEle.MaxHeight - 2);
			TitleEle.style.width = 200;
			
			// Now we need to increment the left position.
			//this.nzEleConStartX += 200;
			
			// Now we need to turn the content display off.
			ContentEle.style.display = "none";
			
			// Change the button's property settings
			img.title = "RestoreUp"
        	img.src = "buttonbottom.gif";
		}
	},
	nzEleSetSize: function(){
		
		// no need to change title height because that is always set at 20px.
		//this.nzEleConTitleID.MaxHeight;
		
		//this.nzEleConContentHeight = h;
		//this.nzEleConContentWidth = w;
		//this.nzEleConTitleWidth = w;
		
		var nzEleTitle = document.getElementById(this.nzEleConTitleID);
		var nzEleContent = document.getElementById(this.nzEleConContentID);
		/*alert("nzEleConTitleHeight: " + this.nzEleConTitleHeight);
		alert("nzEleConContentHeight: " + this.nzEleConContentHeight);
		alert("nzEleConTitleWidth: " + this.nzEleConTitleWidth);
		alert("nzEleConContentWidth: " + this.nzEleConContentWidth);*/
		// If the height or width is greater then the main container height or width, then shrink to fit main container.
		if(nzEleContent.MaxHeight < this.nzEleConContentHeight){
			//alert("nzEleTitle1: " + nzEleTitle);
			nzEleContent.style.height = (nzEleContent.MaxHeight - 2);
		}
		else
		{
			nzEleContent.style.height = this.nzEleConContentHeight;	
		}
		
		if(nzEleContent.MaxWidth < this.nzEleConContentWidth){
			//alert("nzEleTitle2: " + nzEleTitle);
			nzEleTitle.style.width = nzEleContent.MaxWidth;		
			nzEleContent.style.width = nzEleContent.MaxWidth;
		}
		else{
			//alert("nzEleTitle3: " + nzEleTitle);
			nzEleTitle.style.width = (this.nzEleConTitleWidth - 2);		
			nzEleContent.style.width = (this.nzEleConContentWidth - 2);	
		}

		return this; // return this in order to chain.
	},
	nzEleAddPx: function(num){
		return String(num) + "px";
	},
	nzEleConAddClass: function(cssName){
		switch(name)
		{
			case "SleakBlue":
				this.nzEleConTitleClassName = "divDragTitle";
				this.nzEleConContentClassName = "divDragContent";
				break;
		}
		return this; // return this in order to chain.
	},
	nzEleConAssignID: function(eleID,conID,title,html){
		this.nzEleConID = eleID;
		this.nzEleConMainConID = conID;
		this.nzEleConContentID = "dragContent_" + this.nzEleConID + this.nzEleConCount;
		this.nzEleConTitleID = "dragTitle_" + this.nzEleConID + this.nzEleConCount;
		this.nzEleConTitleCaption = title;
		this.nzEleCondragMinBtn = "dragMinBtn" + this.nzEleConCount;	
		this.nzEleCondragMaxBtn = "dragMaxBtn" + this.nzEleConCount;	
		this.nzEleCondragPinBtn = "dragPinBtn" + this.nzEleConCount;
		this.nzEleCondragCloseBtn = "dragCloseBtn" + this.nzEleConCount;
		this.nzEleConContents = html;
	},
	nzEleConGetMainConWidth:function(){
		if(this.nzEleConMainConID.style.width != null){
			//alert("width1: " + document.getElementById(this.nzEleConMainConID).style.width);
			return this.nzEleConMainConID.style.width.replace("px","");
		}
		else{
			//alert("width2: " + document.getElementById(this.nzEleConMainConID).style.width);
			return this.nzEleConMainConID.offsetWidth.replace("px","");
		}
	},
	nzEleConGetMainConHeight:function(){
		if(this.nzEleConMainConID.style.height != null){
			return this.nzEleConMainConID.style.height.replace("px","")
		}
		else{
			return dthis.nzEleConMainConID.offsetHeight.replace("px","");
		}		
	}/*,
	nzEleConInIt: function(eleID,conID,title){		
		nzElementContentProps.nzEleConPropsCount ++; // we increment the external element count property
		this.nzEleConCount = nzElementContentProps.nzEleConPropsCount; // we then assign that property to the main element content count property		
		this.nzEleConAssignID(eleID,conID,title); // we will populate the empty ID's here when we set up the object.
		
		//this.nzEleConMaxWidth = document.getElementById(conID).style.width.replace("px",""); // This will set the container height
		//this.nzEleConMaxHeight = document.getElementById(conID).style.height.replace("px",""); // This will set the container width
		this.nzEleConMaxWidth = this.nzEleConGetMainConWidth();
		this.nzEleConMaxHeight = this.nzEleConGetMainConHeight();
		//alert("Width1: " + this.nzEleConMaxWidth);
		//alert("Height1: " + this.nzEleConMaxHeight);
				
		this.nzEleConCreateElement(); // we will create the element.

		return this; // return this in order to chain.
	}*/
	
	// What other methods will we need?
	
	//***** End of methods *****//
};



//******************************************************************************
// Main newzapp framework //

(function() {
  // private constructor

	function _$(els) {
		_$.prototype.elements = []; // This is used to add the elements to the array.  It can handle more then 1 at a time.
		for (var i=0; i<els.length; i++) {
			var element = els[i];
			if (typeof element == 'string') {
				//element = document.getElementById(element);
				element = eval(nzConstants.nzConstDocument + "." + nzConstants.nzConstGetElementByID + "('" + element + "')");
				//element = _$.prototype.nzEval(nzConstants.nzConstDocument + "." + nzConstants.nzConstGetElementByID + "('" + element + "')");
			}
			this.elements.push(element);
		}
		return this;
	}
	
	_$.prototype = {		
		/*nzEval: function(){
			//alert("element: " + element);
			for(var i = 0;i<arguments.length;i++){
			//	alert("arguments: " + arguments[i]);
				return eval(arguments[i]);
			}
		},*/
		nzElementExists: function(){
			for(var i=0;i<this.elements.length;i++){
				if(this.elements[i]){
					return true;
				}
				else{
					return false;	
				}
			}
		},
		nzSetPropValueByPropName: function(prop,value){
			for(var i=0;i<this.elements.length;i++){
				//alert(this.elements[i]);
				//eval("this.elements[i]." + prop + " = '" + value + "'");
				eval("this.elements[i]." + prop + " = '" + value + "'");
				//return document.getElementById("dragTitle_Cont1").MaxHeight;
			}
		},
		nzGetPropValueByPropName: function(prop){
			for(var i=0;i<this.elements.length;i++){
				//alert(this.elements[i]);
				return eval("this.elements[i]." + prop);
				//return document.getElementById("dragTitle_Cont1").MaxHeight;
			}
			//return this;
		},
		nzGetHeight: function(){
			for(var i=0;i<this.elements.length;i++){
				return eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstHeight);
			}
		},
		nzGetWidth: function(){
			for(var i=0;i<this.elements.length;i++){
				return eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstWidth);
			}
		},
		nzEleSetSize: function(h,w){
			for(var i=0;i<this.elements.length;i++){
				//this.elements[i].style.height = h;
				//this.elements[i].style.width = w;
				//alert(this.elements[i]);
				eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstHeight + " = '" + h + "'");
				
				eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstWidth + " = '" + w + "'");
			}
			return this;
		},
		nzEleConInIt: function(eleID,title,h,w,html){
			var conID = "";
			for(var i=0;i<this.elements.length;i++){
				conID = this.elements[i];	
			}
			
			var ObjElementContent = object(nzElementContainer); // creates new nzElementContainer object
			
			nzElementContentProps.nzEleConPropsCount ++; // we increment the external element count property
			ObjElementContent.nzEleConCount = nzElementContentProps.nzEleConPropsCount; // we then assign that property to the main element content count property		
			ObjElementContent.nzEleConAssignID(eleID,conID,title,html); // we will populate the empty ID's here when we set up the object.
			
			if(h == "" || h == undefined){ // set default height values
				ObjElementContent.nzEleConTitleHeight = 20;			
				ObjElementContent.nzEleConContentHeight = 200;
			}
			else
			{
				ObjElementContent.nzEleConTitleHeight = 20;			
				ObjElementContent.nzEleConContentHeight = h;
			}
			
			if(w == "" || w == undefined){// set default width values
				ObjElementContent.nzEleConTitleWidth = 300;						
				ObjElementContent.nzEleConContentWidth = 300;	
			}
			else
			{
				ObjElementContent.nzEleConTitleWidth = w;
				ObjElementContent.nzEleConContentWidth = w;
				
			}
			ObjElementContent.nzEleConMaxWidth = ObjElementContent.nzEleConGetMainConWidth();
			ObjElementContent.nzEleConMaxHeight = ObjElementContent.nzEleConGetMainConHeight();					
			ObjElementContent.nzEleConCreateElement(); // we will create the element.
	
			return this; // return this in order to chain.
		}/*,
		nzEleConCreateElement: function(ObjElementContent){
			ObjElementContent.nzEleConCreateElement();
		}*/,
		nzGetXMLHTTP: function(req){
			try{
				if(req == "XMLHTTP"){
					var req = new XMLHttpRequest();
					if(req.readyState == null) 
					{
						req.readyState = 1;
						req.addEventListener("load", 
											function() 
											{
												req.readyState = 4;
												if(typeof req.onreadystatechange == "function")
													req.onreadystatechange();
											},
											false);
					}
					return req;
					
				}
				if(req == "ActiveX"){
					return new ActiveXObject(nzConstants.nzConstXMLDoc);
				}				
			}
			catch(ex){}			
		},
		nzAjax: function(func,strSoap,webServiceName){
			//alert("strSoap: " + strSoap);
						
			var strXMLURL = "";
			var strXSLURL = "";
			if(func != ""){
				strXMLURL = nzXMLDocument.getXMLDoc(func);
				strXSLURL = nzXSLDocument.getXSLDoc(func);
			}
			
			if(webServiceName == undefined){	
				if(strXMLURL != "" && strXSLURL != ""){
					for(var i = 0;i<this.elements.length;i++){
						var strID = this.elements[i];
					}

					var objXML = _$.prototype.nzGetXMLHTTP("XMLHTTP");
	
					var objXSL = _$.prototype.nzGetXMLHTTP("ActiveX");
					objXSL.async = true;	
		
					var objXMLDATA;
					objXMLDATA = _$.prototype.nzGetXMLHTTP("ActiveX");
					objXMLDATA.async = true;
						
					_$.prototype.nzSendSoap(objXML,objXMLDATA,strXMLURL,strSoap,"POST",true);					
					objXSL.load(strXSLURL);		
					//alert("got here");
					_$.prototype.nzOnReadyStateChangeXMLDATA(objXMLDATA,objXSL,strID);
					_$.prototype.nzOnReadyStateChangeXSL(objXMLDATA,objXSL,strID,objXML);

				}
				else{
					alert("xmlurl or xslurl is incorrect");
				}
			}
			else{
				var xmlHttp = _$.prototype.nzGetXMLHTTP("XMLHTTP");
					
				//xmlHttp.open("GET", strXMLURL + "?wsdl", true);
				//xmlHttp.onreadystatechange = function() 
				//{
					//if(xmlHttp.readyState == 4){
						_$.prototype.nzOnReadyStateChangeWebService(xmlHttp,strSoap,webServiceName,strXMLURL,func,"POST",true);
					//}
				//}
				//xmlHttp.send(null);
			}
			return this; //Return this in order to chain
		},
		nzGetChildNodeValue: function(xmlNode,value){
			var treeTxt=""; //var to content temp
			
			for(var i=0;i<xmlNode.childNodes.length;i++){//each child node
				if(xmlNode.childNodes[i].nodeType == 1){//no white spaces
					//node name
					treeTxt = xmlNode.childNodes[i].nodeName + ": "
					if(xmlNode.childNodes[i].childNodes.length==0){
						//no children. Get nodeValue
						treeTxt = xmlNode.childNodes[i].nodeValue 
						for(var z=0;z<xmlNode.childNodes[i].attributes.length;z++){
							var atrib = xmlNode.childNodes[i].attributes[z];
							treeTxt = " (" + atrib.nodeName + " = " + atrib.nodeValue + ")";
						}						
					}else if(xmlNode.childNodes[i].childNodes.length>0){
						//children. get first child
						treeTxt = xmlNode.childNodes[i].firstChild.nodeValue;
						if(treeTxt == null){
							for(var z=0;z<xmlNode.childNodes[i].attributes.length;z++){
								var atrib = xmlNode.childNodes[i].attributes[z];
								treeTxt = " (" + atrib.nodeName + " = " + atrib.nodeValue + ")";
							}
							//recursive to child of children							
							treeTxt = _$.prototype.nzGetChildNodeValue(xmlNode.childNodes[i],treeTxt);							
						}
					}
				}
			}
			return treeTxt;
		},
		bXMLLoaded: false,
		bXSLLoaded: false,
		nzOnReadyStateChangeXSL: function(objXMLDATA,objXSL,strID,objXML){
			objXSL.onreadystatechange = function() {
				if (objXSL.readyState==4)
				{
					_$.prototype.bXSLLoaded = true;	
					//alert("_$.prototype.bXMLLoaded1: " + _$.prototype.bXMLLoaded);
					//alert("_$.prototype.bXSLLoaded1: " + _$.prototype.bXSLLoaded);
					if (_$.prototype.bXMLLoaded == true && _$.prototype.bXSLLoaded == true)
					{			
						//alert("XML1: " + objXML.responseXML);
						objXMLDATA.load(objXML.responseXML);
						
						//strID.innerHTML = objXMLDATA.transformNode(objXSL);
						_$.prototype.bXMLLoaded = false; // reset the values once we're finished
						_$.prototype.bXSLLoaded = false; // reset the values once we're finished
						AjaxComplete(strID,objXMLDATA.transformNode(objXSL));
						
					}
				}
			}
		},
		nzOnReadyStateChangeXMLDATA: function(objXMLDATA,objXSL,strID){			
			objXMLDATA.onreadystatechange = function() {
				if (objXMLDATA.readyState==4)
				{
					_$.prototype.bXMLLoaded = true;
					//alert("_$.prototype.bXMLLoaded2: " + _$.prototype.bXMLLoaded);
					//alert("_$.prototype.bXSLLoaded2: " + _$.prototype.bXSLLoaded);
					if (_$.prototype.bXMLLoaded == true && _$.prototype.bXSLLoaded == true)
					{			
						//alert("XML2: " + objXMLDATA.transformNode(objXSL));
						//strID.innerHTML = objXMLDATA.transformNode(objXSL);
						_$.prototype.bXMLLoaded = false; // reset the values once we're finished
						_$.prototype.bXSLLoaded = false; // reset the values once we're finished
						AjaxComplete(strID,objXMLDATA.transformNode(objXSL));
						
					}
				}
			}	
		},
		nzOnReadyStateChangeWebService: function(xmlHttp,strSoap,webServiceName,strXMLURL,func,SendType,async){
			xmlHttp.open(SendType, strXMLURL, async);
			//var wsdl = xmlHttp.responseXML;
			var ns = "http://www.newzapp.co.uk/";
			//alert("wsdl.documentElement.attributes['targetNamespace']: " + wsdl.childNodes[0].attributes["targetNamespace"]);
			//var ns = (wsdl.documentElement.attributes["targetNamespace"] + "" == "undefined") ? wsdl.documentElement.attributes.getNamedItem("targetNamespace").nodeValue : wsdl.documentElement.attributes["targetNamespace"].value;
			//alert("ns: " + ns);
			var soapaction = ((ns.lastIndexOf("/") != ns.length - 1) ? ns + "/" : ns) + func;
			xmlHttp.setRequestHeader("SOAPAction", soapaction);
			xmlHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
			
			xmlHttp.onreadystatechange = function() 
			{
				if(xmlHttp.readyState == 4){
					var NodeValue = _$.prototype.nzGetChildNodeValue(xmlHttp.responseXML,"");
					eval(webServiceName + "('" + NodeValue + "')");
				}
			}
			
			xmlHttp.send(strSoap);
		},
		nzSendSoap: function(objXML,objXMLDATA,strXMLURL,strSoap,SendType,async){
			objXML.open(SendType, ""+strXMLURL+"", async);
			objXML.onreadystatechange = function() {
				if (objXML.readyState==4)
				{			
					objXMLDATA.load(objXML.responseXML);			
				}
			}
					
			objXML.send(strSoap);	
		},
		//Adds a new class to the element
		addClass:function(){
			if(this.elements.length == 1){
				for(var i = 0;i<arguments.length;i++){
					eval("this.elements[0]." + nzConstants.nzConstClassName + " = '" + arguments[i] + "'");
				}
				//_$.prototype.nzEval("this.elements[0]." + nzConstants.nzConstClassName + " = '" + name + "'");
			}
			else{
				for(var i = 0;i<this.elements.length;i++){
					for(var j = 0;j<arguments.length;j++){
						eval("this.elements[i]." + nzConstants.nzConstClassName + " += ' ' + '" + arguments[j] + "'");
					}
					//_$.prototype.nzEval("this.elements[i]." + nzConstants.nzConstClassName + " += ' ' + '" + name + "'");
				}
			}
			return this; //Return this in order to chain
		},
		nzOverflow: function(){
			//document.documentElement.style.overflow  = "hidden";
			//alert(this.elements.length);
			/*for(var i = 0;i<=arguments.length;i++){	
				for(var i = 0;i<this.elements.length;i++){
					eval("this.elements[i]." + nzConstants.nzConstDocumentElement + "." + nzConstants.nzConstStyle + "." + nzConstants.nzConstDisplay + " = 'block'");
					
				}
			}*/
		},
		//Show and Hide the elements found
		toggleDisplay: function(){
			for(var i = 0;i<this.elements.length;i++){
				//Check the status of the element to know if it could be displayed or closed			
				if (eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstDisplay + " == 'none'")){
					eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstDisplay + " = 'block'");
				}
				else{
					eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstDisplay + " = 'none'");
				}
			}
			return this;//Return this in order to chain
		},
		nzShowDisplay: function(){
			for(var i = 0;i<this.elements.length;i++){
				//Check the status of the element to know if it could be displayed or closed			
					eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstDisplay + " = 'block'");
			}
			return this;//Return this in order to chain
		}
		,
		nzHideDisplay: function(){
			for(var i = 0;i<this.elements.length;i++){
				//Check the status of the element to know if it could be displayed or closed			
					eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstDisplay + " = 'none'");
			}
			return this;//Return this in order to chain
		},				
		nzHTML: function(){	
			for(var i = 0;i<=arguments.length;i++){							
				if(arguments[i] != undefined){ // Will set the innerHTML
					for(var j=0;j<this.elements.length;j++){
						eval("this.elements[j]." + nzConstants.nzConstInnerHTML + " = arguments[i]");	
						//_$.prototype.nzEval("this.elements[i]." + nzConstants.nzConstInnerHTML + " = html");
					}					
					return this;
				}
				else // Will get the innerHTML. You will not be able to chain this.
				{
					for(var j=0;j<this.elements.length;j++){
						return eval("this.elements[j]." + nzConstants.nzConstInnerHTML);
						//_$.prototype.nzEval("this.elements[i]." + nzConstants.nzConstInnerHTML);
					}
				}
			}
		},				
		nzText: function(text){						
			if(text != undefined){ // Will set the innerText
				for(var i=0;i<this.elements.length;i++){
					//this.elements[i].innerText= text;	
					eval("this.elements[i]." + nzConstants.nzConstInnerText + " = text");
					
				}					
				return this;
			}
			else // Will get the innerText. You will not be able to chain this.
			{
				for(var i=0;i<this.elements.length;i++){
					//return this.elements[i].innerText;
					return eval("this.elements[i]." + nzConstants.nzConstInnerText);
				}
			}
		},				
		nzValue: function(value){						
			if(value != undefined){ // Will set the innerText
				for(var i=0;i<this.elements.length;i++){
					eval("this.elements[i]." + nzConstants.nzConstValue + " = value");
					
				}					
				return this;
			}
			else // Will get the value. You will not be able to chain this.
			{
				for(var i=0;i<this.elements.length;i++){
					//return this.elements[i].innerText;
					return eval("this.elements[i]." + nzConstants.nzConstValue);
				}
			}
		},
		nzAppendText:function(text){ // This will append text onto an existing text.
			text = document.createTextNode(text); //Create a new Text Node with the string supplied
			for(var i=0;i<this.elements.length;i++){
				//this.elements[i].appendChild(text);
				eval("this.elements[i].appendChild(text)");
			}
			return this;//Return this in order to chain
		},
		nzBGColor: function(value){
			if(value != undefined){ 
				for(var i=0;i<this.elements.length;i++){// Will set the background color.
					//alert(this.elements[i] + "." + _$.prototype.nzConstStyle + "." + _$.prototype.nzConstBGCol + " = " + value);
					eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstBGCol + " = value");
				}
				return this;//Return this in order to chain
			}
			else{
				for(var i=0;i<this.elements.length;i++){
					return 	eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstBGCol);
				}
			}
		},
		nzFGColor: function(value){						
			if(value != undefined){ 
				for(var i=0;i<this.elements.length;i++){// Will set the foreground color.					
					eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstFGCol + " = value"); 
				}
				return this;//Return this in order to chain
			}
			else{
				for(var i=0;i<this.elements.length;i++){
					return 	eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstFGCol);
				}
			}
		},
		nzMargin: function(){		
			
			for(var i = 0;i<this.elements.length;i++){
				switch(arguments.length) //can take multiple arguments.
				{		
					case 1:
						//_$.prototype.nzPadding1Arg(this.elements[i],nzPaddingStyle,nzPadValue,nzUnits);
						//alert(arguments[0]);

						eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstMarginTop + " = '10px'"); 
						//document.getElementById("content").style.marginTop = "12px";
						
						break;					
				}
			}
			return this;//Return this in order to chain
		},
		nzPadding: function(){	
		
			var nzPaddingStyle = "";
			var nzUnits = "";
			var nzPadValue = 0;
			var nzPaddingScript = "";
		
			for(var i = 0;i<arguments.length;i++){
				if(typeof arguments[i] == 'string'){ //Verify if the parameter is a string
					switch(arguments[i].toUpperCase())
					{
						case "TOP":
							nzPaddingStyle = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingTop;
							break;
						case "LEFT":
							nzPaddingStyle = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingLeft;
							break;
						case "RIGHT":
							nzPaddingStyle = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingRight;
							break;
						case "BOTTOM":
							nzPaddingStyle = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingBottom;
							break;
						case "ALL":				
							nzPaddingStyle = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingAll;
							break;
						case "PX":
							nzUnits = "px";
							break;
						case "CM":
							nzUnits = "cm";
							break;
					}
				}
				else if(typeof arguments[i] == 'number'){ //Verify if the parameter is a number					
					nzPadValue = arguments[i];						
				}
			}
			
			for(var i = 0;i<this.elements.length;i++){
				switch(arguments.length) //can take multiple arguments.
				{		
					case 0:
						// No arguments was passed in so we will return the padding All.
						return eval("this.elements[i]." + nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingAll); 
					case 1:
						_$.prototype.nzPadding1Arg(this.elements[i],nzPaddingStyle,nzPadValue,nzUnits);
						break;
					case 2:
						_$.prototype.nzPadding2Arg(this.elements[i],nzPaddingStyle,nzPadValue,nzUnits);
						break;
					case 3:
						_$.prototype.nzPadding3Arg(this.elements[i],nzPaddingStyle,nzPadValue,nzUnits);
						break;
				}
			}
			return this; //Return this in order to chain
		},
		// Start of method overloading
		
		nzPadding1Arg: function(ele,nzPaddingStyle,nzPadValue,nzUnits){
			if(nzPaddingStyle != ""){
				// if only the padding style is inserted, we then use 10px as default.
				nzPaddingScript = nzPaddingStyle + " = '10px'"; 
			}
			else if(nzUnits != ""){
				// if only the padding units is inserted, we then use style as All and 10 as default.
				nzPaddingScript = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingAll + " = '10" + nzUnits + "'";
			}
			else if(nzPadValue != ""){
				// if only the padding value is inserted, we then use style as All and px as default.
				nzPaddingScript = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingAll + " = '" + nzPadValue + "px'";
			}
			eval("ele." + nzPaddingScript);	
		},
		nzPadding2Arg: function(ele,nzPaddingStyle,nzPadValue,nzUnits){
			if(nzPaddingStyle != "" && nzPadValue != ""){
				nzPaddingScript = nzPaddingStyle + " = '" + nzPadValue + "px'";
			}
			else if(nzPaddingStyle != "" && nzUnits != ""){
				nzPaddingScript = nzPaddingStyle + " = '10" + nzUnits + "'";
			}
			else if(nzPadValue != "" && nzUnits != ""){
				nzPaddingScript = nzConstants.nzConstStyle + "." + nzConstants.nzConstPaddingAll + " = '" + nzPadValue + nzUnits +"'";
			}
			eval("ele." + nzPaddingScript);
		},
		nzPadding3Arg: function(ele,nzPaddingStyle,nzPadValue,nzUnits){			
			if(nzPaddingStyle != "" && nzPadValue != "" && nzUnits != ""){
				nzPaddingScript = nzPaddingStyle + " = '" + nzPadValue + nzUnits + "'";
			}
			eval("ele." + nzPaddingScript);
		},
		
		// End of method overloading functions.
		
		on: function(action, callback){
		   if(this.elements[0].addEventListener){
				for(var i = 0;i<this.elements.length;i++){
					this.elements[i].addEventListener(action,callback,false); //Adding the event by the W3C for Firefox, Safari, Opera...   
				}
			}
		   else if(this.elements[0].attachEvent){
				for(var i = 0;i<this.elements.length;i++){
					this.elements[i].attachEvent('on'+action,callback); //Adding the event to Internet Explorer :(
				}
			}
			return this;//Return this in order to chain
		}		
	}
	
	window.$ = function() {
		return new _$(arguments);
	}
})();
	
// End of main newzapp framework
// *****************************************************************************