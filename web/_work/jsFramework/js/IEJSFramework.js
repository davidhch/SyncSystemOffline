/*
JavaScript Document	
JS Framework for newzapp V5.
Created by MH.
Modified by DC 05 August 2009
*/

(
// DO NOT DELETE THE OPEN BRACKET ABOVE
//--------------------------------------------------------------------
 
 function() {
  	// private constructor
	window.$ = function() {
		return new _$(arguments);
	}
	
	function _$(els) {
		_$.prototype.elements = []; // This is used to add the elements to the array.  It can handle more then 1 at a time.
		for (var i=0; i<els.length; i++) {
			var element = els[i];
			if (typeof element == 'string') {
				element = eval(_C.document + "." + _C.getElementById + "('" + element + "')");
			}
			this.elements.push(element);
		}
		
		return this;
	}
	
	_$.prototype = {		
				
		innerHTML: function(){
			for(var i = 0;i<=arguments.length;i++){							
				if(arguments[i] != undefined){
					for(var j=0;j<this.elements.length;j++){
						eval("this.elements[j]." + _C.innerHTML + " = arguments[i]");	
					}					
					return this;
				}
				else // Will get the innerHTML. You will not be able to chain this.
				{
					
					for(var j=0;j<this.elements.length;j++){	
						return eval("this.elements[j]." + _C.innerHTML);
					}
				}
			}
		},				
		innerText: function(text){						
			if(text != undefined){ 
				for(var i=0;i<this.elements.length;i++){
					eval("this.elements[i]." + _C.innerHTML + " = text");
				}					
				return this;
			}
			else // Will get the innerText. You will not be able to chain this.
			{
				for(var i=0;i<this.elements.length;i++){
					return eval("this.elements[i]." + _C.InnerText);
				}
			}
		},				
		value: function(value){						
			if(value != undefined){ // Will set the innerText
				for(var i=0;i<this.elements.length;i++){
					eval("this.elements[i]." + _C.value + " = value");
					
				}					
				return this;
			}
			else // Will get the value. You will not be able to chain this.
			{
				for(var i=0;i<this.elements.length;i++){
					//return this.elements[i].innerText;
					return eval("this.elements[i]." + _C.value);
				}
			}
		},
		
		forLoop: function(val,bGet){
			for(var i=0;i<this.elements.length;i++){
				if (bGet==1){
					return eval("_$.prototype.elements["+i+"]."+ val);
				}
				else{
					eval("_$.prototype.elements["+i+"]."+ val);
				}
			}
		}
		
		,
		addClass: function(val){return this;}
		,
		style: {
				// get the sub functions return
				//---------- START OBJECT style [named to match w3c dot notation].
				backgroundColor: function(val){
					if(val != undefined){ 
						_$.prototype.forLoop(_C.style + "." + _C.backgroundColor + " = '" + val + "'",0);
						return this; //Return this in order to chain
					}
					else{
						return _$.prototype.forLoop(_C.style + "." + _C.backgroundColor,1);
					}
				}
				,
				color: function(val){
					if(val != undefined){ 
						_$.prototype.forLoop(_C.style + "." + _C.color + " = '" + val + "'",0);
						return this; //Return this in order to chain
					}
					else{
						return _$.prototype.forLoop(_C.style + "." + _C.color,1);
					}
				},
				
				margin: function(val){	//  replace with an overloading method	
					if(val != undefined){ 
						_$.prototype.forLoop(_C.style + "." + _C.margin + " = '" + val + "'",0);
						return this; //Return this in order to chain
					}
					else{
						return _$.prototype.forLoop(_C.style + "." + _C.margin,1);
					}
				},
				//---------- END OBJECT style [named to match w3c dot notation].
				
			}
		,
		// example of 2 levels deep
		dc: {
			  style:{
				  backgroundColor: function(value){}
			  }
		}
		,
		

		on: function(action, callback){
			
			for(var i = 0;i<this.elements.length;i++){
				switch(action){
					case "click":						
						this.elements[i].onclick = callback; //Adding the onclick
						break;
					case "mouseover":
						this.elements[i].onmouseover = callback; //Adding the onmouseover
						break;
					default:
						this.elements[i].attachEvent('on'+action,callback); //Adding the event to Internet Explorer
				}
			}
			
			return this;//Return this in order to chain
		},	
		
		
//=================================================================================================
//	METHODS - FOR EXTERNTAL OBJECTS - we'll be using external objects when building components
// 	check js file is loaded in head, if NOT loaded THEN load
	
		init_tabs: function(){
			for(var i = 0;i<this.elements.length;i++){tabs.init_tabs( this.elements[i] );}
		},
		
		init_anotherThing:function(){
			for(var i = 0;i<this.elements.length;i++){anotherThing.init_anotherThing( this.elements[i] );}	
		},
		
		init_subscriberList:function(){
			for(var i = 0;i<this.elements.length;i++){subscriberList.init_subscriberList( this.elements[i] );}	
		}
	}
	
	
}
//--------------------------------------------------------------------
// DO NOT DELETE THE CLOSE BRACKET ABOVE AND FUNCTIONS BRACKETS BELOW
)();
