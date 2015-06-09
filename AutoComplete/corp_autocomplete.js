    //##############################################################################
    //Inicia a Super class
    function CORP(){};
    //##############################################################################
    CORP.prototype.AttEvent = function (ev, fun){
	    if (document.attachEvent)
		    document.attachEvent(ev, fun);
	    else
		    document.addEventListener(ev, fun, true);	
    }
    //####################################################
    CORP.prototype.AttEventElement = function (el, ev, fun){
	    if (el.attachEvent)
		    el.attachEvent(ev, fun);
	    else
		    el.addEventListener(ev, fun, true);	
    }	
    //####################################################    
    //captura o elemento
    CORP.prototype.GetElement = function(id){
    	if (document.getElementById)return document.getElementById(id);
    	else if (document.all) return document.all[id];
    }
    //captura o Evento do elemento
    CORP.prototype.GetEvent = function(eEvent){
		return eEvent ? eEvent : window.event; 
    }
    //captura o elemento que disparou e evento e retorna a Property desejada
    CORP.prototype.GetEventDispatched = function(eEvent,ReturnProperty){
		switch(ReturnProperty){
		case "target" :								
		case "srcElement" :		
			return eEvent.target ? eEvent.target : eEvent.srcElement; 
		case "fromElement" :
		case "relatedTarget" :
			return eEvent.fromElement ? eEvent.fromElement : eEvent.relatedTarget; 
		}	
    }
    CORP.prototype.conv2hex = function(palavra){
	     var hexChars = "0123456789ABCDEF";
         var hexStr   = "";
         var a        = 0 ;
         var b        = 0 ;
         var chr      = 0 ;
    	 //------------------------------------------------------
	     for (var i = 0; i < palavra.length; i++) { 
	        //---------------------------------------------------
            chr = palavra.charCodeAt(i);
	        a   = chr % 16;
	        b   = (chr - a) / 16;
    	    //---------------------------------------------------
	        hexStr += "" + hexChars.charAt(b) + hexChars.charAt(a);
	     }
         //------------------------------------------------------    	 	 
	     return hexStr; 
	     //------------------------------------------------------
    }
    //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@    
    //——————————————————————————————————————————————————————————————————————————————
    CORP.prototype.Msg = new CORP_MSG(); function CORP_MSG(){};       
    // Show Menssage
    CORP_MSG.prototype.ShowMessage = function(msg,it_pos_relative,timeShow){        
        var div =  document.createElement("div");
        var el  = corp.GetElement(it_pos_relative);
        /*-----------------------------------------*/            
        div.id="CORP_MSG_MsgBox";
        with (div.style) { 
    	    top= el.offsetTop + "px"; 
    	        	    
    	    if (aut.activeTypeEvent == 1 || aut.activeTypeEvent == 3)    	        
    	        left=( 20 + el.offsetWidth + el.offsetLeft ) + "px";     	        	    
    	    else{
	            var btn = corp.GetElement(aut.items[aut.activeElement].button);
	            left=( 20 + btn.offsetWidth + el.offsetWidth + el.offsetLeft ) + "px";                
    	    }
    	    
    	    position="absolute";	    	
    	    zIndex=1;    	    
    	    visibility="visible";    	    
    	}    	
    	div.innerText = msg;
    	document.body.appendChild(div);    	
        if (timeShow != undefined) {
            window.clearTimeout()
            window.setTimeout("c_msg.destroy()", timeShow);
        }
    }    
    CORP_MSG.prototype.destroy  = function () {
        var div  =  corp.GetElement("CORP_MSG_MsgBox");
         if ( div != null ) 
            document.body.removeChild(div);                 
    }       
    //——————————————————————————————————————————————————————————————————————————————
    //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@     
    //Cria uma copia de Ajax para CORP
    CORP.prototype.Ajax = new CORP_AJAX();
    //#############################################################################
    //Inicia a Sub class
    function CORP_AJAX(){
    	this.url;
    	this.ajax;
    }
    //##############################################################################
    //captura o HttpRequest
    CORP_AJAX.prototype.GetAjax = function GetHttpRequest()
     {if (window.XMLHttpRequest) {return new XMLHttpRequest();} else if (window.ActiveXObject) {	return new ActiveXObject("Microsoft.XMLHTTP");}}
    //——————————————————————————————————————————————————————————————————————————————
    //##############################################################################
    CORP_AJAX.prototype.GetAjaxValue = function(eventAttach){	
    	this.GetAjaxHttpRequest(eventAttach);
    }
    //##############################################################################
    // captura o retorno em texto
    CORP_AJAX.prototype.GetAjaxHttpRequest = function GetResponseText(eventAttach){	
    	var dt = new Date();//para atualizar sempre a url		
    	var cp =  new CORP()    	
    	cp.Ajax.ajax  = this.GetAjax();
    	cp.Ajax.ajax.open("GET", cp.Ajax.url + "&dt=" + dt, true);
    	cp.Ajax.ajax.send(null);
    	cp.Ajax.ajax.onreadystatechange = function(){
    		if (cp.Ajax.ajax.readyState==4){
    		    c_msg.destroy();
    			eventAttach(cp.Ajax.ajax.responseText);
    		}
    	}		
    }             
    //##############################################################################  
    //Cria uma copia de AutoComplete para CORP                  
    CORP.prototype.autoComplete = new CORP_AUTOCOMPLETE(); 
    //##############################################################################        
    function CORP_AUTOCOMPLETE() { 
        this.items = new Object();
        this.TmpDig = 0; 
        this.activeTime;
        this.activeElement;        
        this.activeURL;
        this.activeTypeEvent; //1 = autoComplete || 2 = button click  || 3 = ENTER
        //this.site = "http://localhost/autoComplete/Default.aspx?autoCompleteURL="                
        this.site = "http://10.0.0.31/geral/corp/autocomplete/getautocomplete.aspx?autoCompleteURL=";
        //
        this.eventAttach = function (texto)
        {            
            var txt = corp.GetElement(aut.activeElement);            
            /*-----------------------------------------*/            
            var div =  document.createElement("div");                
            div.id = "divautoCompleteListList"    	    
            /*-----------------------------------------*/            
            with (div.style) { 
    	        top=(txt.offsetTop + txt.offsetHeight) + "px"; 
    	        left=txt.offsetLeft + "px";     	        
    	        width=txt.offsetWidth + "px";
    	        position="absolute";	    	
    	        zIndex=0;
    	        visibility="visible";
    	        display="inline";        	        
    	    }
    	    /*-----------------------------------------*/
    	    if (texto.substring(0, 3) == "&&&")
    	       alert("Nada foi encontrado!");
    	    else
    	        div.innerHTML = texto;    	        	    
    	        
    	    document.body.appendChild(div);
    	    /*-----------------------------------------*/
    	    var lst = corp.GetElement("autoCompleteList");    	   
    	    /*-----------------------------------------*/    	    
    	    if (lst != null) {
    	        /*-----------------------------------------*/    	        	        
    	        lst.onkeypress = function(eEvent) {
					var ev =  corp.GetEvent(eEvent);		    	          
    	            if (ev.keyCode == 13 /*Enter*/) 
    	            lst.ondblclick();
    	        }
    	        //--------------------------------------------    	            	            	        
    	        lst.ondblclick = function (eEvent) {
  	            	var it    = aut.items[aut.activeElement];
    	            var hdn   = corp.GetElement(it.value);
    	            var txt   = corp.GetElement(it.text);
    	            //--------------------------------------------
    	            with (lst.options[lst.selectedIndex]) {
    	                hdn.value      = value;
    	                txt.value      = text;
    	                aut.activeText = text;
    	            }
    	            //--------------------------------------------
    	            aut.submit()
    	            //--------------------------------------------
        	        aut.destroy();
    	        }
    	        //--------------------------------------------
    	        lst.style.width = txt.offsetWidth;
    	        lst.focus();      	                        
    	    }
    	    
    	    /*-----------------------------------------*/
        }
        //
    }     
    /*==================================================
       t => campo texto (text)
       v => campo valor (hidden)
       u => param. url
       k => keysize
       p => 1 = autoComplete || 2 = button click
       b => botão
     ==================================================*/
    CORP_AUTOCOMPLETE.prototype.add = function (t, v, u, k, p, b) {       
        var elKeyEvent = corp.GetElement(t);        
        it = new aut.item.AddItem(t, v, u, k, p, b);                                    
        elKeyEvent.onfocus = aut.destroy;
        //1 = autoComplete || 2 = button click|| 3 = ENTER        
        switch(it.typeEvent){
            case 1 :
                //elKeyEvent.onkeydown  = aut.item.onkeydown;
                //elKeyEvent.onkeypress = aut.item.onkeypress;
                corp.AttEventElement(elKeyEvent, "onkeydown", aut.item.onkeydown);
                corp.AttEventElement(elKeyEvent, "onkeypress", aut.item.onkeypress);
                break;
            case 2 :
                //elKeyEvent.onkeyup = aut.item.onkeyup;
                corp.AttEventElement(elKeyEvent, "onkeyup", aut.item.onkeyup);
                var elButtonEvent = corp.GetElement(it.button)
                elButtonEvent.idText = it.text;
                //elButtonEvent.onclick =  aut.item.buttonClick;
                corp.AttEventElement(elButtonEvent, "onclick", aut.item.buttonClick);
                break;
            case 3 :
                //elKeyEvent.onkeypress = aut.item.onkeypress;
                corp.AttEventElement(elKeyEvent, "onkeypress", aut.item.onkeypress);
        }        
        this.items[it.text]=it;
    }
    //==================================================
    CORP_AUTOCOMPLETE.prototype.progress = function() {
        //---------------------------------------------
        aut.TmpDig++;        
        var txt = corp.GetElement(aut.activeElement);        
        //---------------------------------------------
        if (aut.TmpDig >= 2) {
            //---------------------------------------------
            window.clearInterval(aut.activeTime);
            aut.TmpDig=0;
            //---------------------------------------------
            c_msg.ShowMessage("carregando...", aut.activeElement)
            //---------------------------------------------
            var cp  =  new CORP();
            var prm =  txt.value;            
	        cp.Ajax.url= aut.site + aut.activeURL + "&prm=" + cp.conv2hex(prm);
	        cp.Ajax.GetAjaxValue(aut.eventAttach);
	        //---------------------------------------------
        }
    }
    //---------------------------------------------
    CORP_AUTOCOMPLETE.prototype.destroy =  function (eEvent) {
        var div =  corp.GetElement("divautoCompleteListList")
        if ( div != null )         
            document.body.removeChild(div);        
        aut.TmpDig=0;
        window.clearInterval(aut.activeTime);        
		var ev =  corp.GetEvent(eEvent);		
		//alert(ev);
        if (ev != null)
            if (ev.keyCode == 27) {
                var el  = corp.GetElement(aut.activeElement);        
                if (el != null) el.focus();
            }
    }      
    CORP_AUTOCOMPLETE.prototype.submit =  function (eEvent) {
    }          
    //###############################################
    //Cria uma copia de AutoCompleteItem para CORP
    //---------------------------------------------      
    CORP_AUTOCOMPLETE.prototype.item = new CORP_AUTOCOMPLETE_ITEM(); 
    //---------------------------------------------
    //Construtora
    function CORP_AUTOCOMPLETE_ITEM(){};
    //---------------------------------------------
    CORP_AUTOCOMPLETE_ITEM.prototype.AddItem  = function (t, v, u, k, p, b) { 
        this.text   = t;
        this.value  = v;        
        this.url    = u;
        this.typeEvent = ( p == undefined ) ? 1 : p; //1 = autoComplete || 2 = button click || 3 Enter
        if (this.typeEvent == 1 )
            this.keySize = (k - 1)
        else
            if (this.typeEvent == 2)
                this.keySize = k 
        this.button  = b;
    }
    //---------------------------------------------
    //CORP_AUTOCOMPLETE_ITEM.prototype.onfocus = CORP_AUTOCOMPLETE.prototype.destroy;
    //---------------------------------------------
    CORP_AUTOCOMPLETE_ITEM.prototype.onkeydown = function(eEvent) {        
		var ev =  corp.GetEvent(eEvent);
		var el = corp.GetEventDispatched(ev,"target");		
		//??????????????????
		window.status = el;
		//??????????????????
        var it  = aut.items[el.id];
        var hdn = corp.GetElement(it.value);
        var txt = corp.GetElement(it.text);
        
        if (hdn.value != "" && ev.keyCode == 8 && txt.value.length >= it.keySize)  txt.onkeypress();                 
        
    }
    //---------------------------------------------
    CORP_AUTOCOMPLETE_ITEM.prototype.onkeypress = function(eEvent) { 
		var ev =  corp.GetEvent(eEvent);
		var el = corp.GetEventDispatched(ev,"target");

        aut.TmpDig=0;
        /* ------------------*/
        it  = aut.items[el.id]; 
        hdn = corp.GetElement(it.value);
        hdn.value = "";
        /* ------------------*/        
        with(it) {                     
            aut.activeElement = text;
            aut.activeURL = url;
            aut.activeTypeEvent = typeEvent;
        }
        /* ------------------------------------------------------*/    
        if (aut.activeTypeEvent == 3 ){   
            if ( ev.keyCode == 13){
                window.clearInterval(aut.activeTime);
                aut.activeTime =  window.setInterval(aut.progress, 400);        
            }               
        }else{
            var el = corp.GetElement(it.text);        
            if ( el.value.length >= it.keySize && ev.keyCode != 27 ){
                window.clearInterval(aut.activeTime);
                aut.activeTime =  window.setInterval(aut.progress, 400);
                /*------------------------------------------------------*/
                //c_msg.ShowMessage("carregando...", aut.activeElement)
                /*------------------------------------------------------*/
            }                
        }
    }    
   //---------------------------------------------      
    CORP_AUTOCOMPLETE_ITEM.prototype.onkeyup = function(eEvent) {
		var ev =  corp.GetEvent(eEvent);
		var el = corp.GetEventDispatched(ev,"target");		    
        var it  = aut.items[el.id];
        var hdn = corp.GetElement(it.value);
        var txt = corp.GetElement(it.text);
        
        aut.activeElement = it.text;
        
        if (hdn.value != "") 
            if (aut.activeText != txt.value) { 
                c_msg.ShowMessage("Texto modificado. Clique no bot" + String.fromCharCode(227) + "o para atualizar valores.", aut.activeElement,  3000);
                hdn.value = "";                        
            }        
    }   
    //---------------------------------------------
    CORP_AUTOCOMPLETE_ITEM.prototype.buttonClick = function() {
        it = aut.items[this.idText];
       //---------------------------------------------
        aut.destroy();
       //---------------------------------------------
        with(it) {                     
            aut.activeElement = text;
            aut.activeURL = url;
            aut.activeTypeEvent = typeEvent;
        }       
        //---------------------------------------------
        var txt = corp.GetElement(it.text);
        var hdn = corp.GetElement(it.value);  
        hdn.value = "";  // limpando o value
        //---------------------------------------------
        if ( txt.value.length >= it.keySize ){
            //---------------------------------------------
            var cp  =  new CORP();
            var prm =  txt.value;
	        cp.Ajax.url= aut.site + aut.activeURL + "&prm=" + cp.conv2hex(prm);
	        cp.Ajax.GetAjaxValue(aut.eventAttach);
	        c_msg.ShowMessage("carregando...", aut.activeElement)
	        //---------------------------------------------
	    }else{	        
	        c_msg.ShowMessage("Qtd. min. de caracteres " + it.keySize + ".", it.button, 1200)
        }    
        return false;
    }           
    //---------------------------------------------
    //###############################################
    var aut = new CORP().autoComplete;
    var c_msg = new CORP().Msg;
    var corp = new CORP();    
    
    //---------------------------------------------