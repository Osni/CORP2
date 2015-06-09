    //##############################################################################
    //Inicia a Super class
    function CORP(){};
    //##############################################################################
    //captura o elemento
    CORP.prototype.GetElement = function(id){
    	if (document.getElementById)return document.getElementById(id);
    	else if (document.all) return document.all[id];	
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
        var el  = corp.GetElement(it_pos_relative);
        var d   = corp.GetElement("divMsg" + aut.activeElement);
        /*-----------------------------------------*/            
        with (d.style) { 
    	    position="absolute";
    	    zIndex=1;    	    
    	    display = ""
    	}
    	d.innerText = msg;    
        if (timeShow != undefined) {
            window.clearTimeout()
            window.setTimeout("c_msg.destroy()", timeShow);
        }
    }    
    CORP_MSG.prototype.destroy  = function () {      
        var dv  =  corp.GetElement("divMsg" + aut.activeElement);
        if (dv.style.display != "none") dv.style.display = "none" 
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
    		if (cp.Ajax.ajax.readyState==4) {    		    
    		    c_msg.destroy("divMsg" + aut.activeElement);
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
        this.activeText;
        this.activeTime;
        this.activeElement;        
        this.activeURL;
        this.activeTypeEvent; //1 = autoComplete || 2 = button click
        this.site = "http://localhost/autoComplete/getautocomplete.aspx?autoCompleteURL="        
        //
        this.eventAttach = function (texto)
        {            
            var txt = corp.GetElement(aut.activeElement);            
            /*-----------------------------------------*/            
            div = corp.GetElement("div" + txt.id);
            /*-----------------------------------------*/            
            with (div.style) { 
    	        position="absolute";
    	        zIndex=0;
    	        display = ""
    	    }
    	    /*-----------------------------------------*/
    	    div.innerHTML = texto;
    	    /*-----------------------------------------*/ 
    	    var lst = corp.GetElement("autoCompleteList");    	   
    	    /*-----------------------------------------*/    	    
    	    if (lst != null) {	            	    
    	        /*-----------------------------------------*/    	        	        
    	        lst.onkeypress = function() {     	                	            
    	            if (event.keyCode == 13 /*Enter*/ || event.keyCode == 0 /*dblclick*/) {
    	                var it    = aut.items[aut.activeElement];
    	                var hdn   = corp.GetElement(it.value);
    	                var txt   = corp.GetElement(it.text);
    	                //--------------------------------------------
    	                with (lst.options[lst.selectedIndex]) {
    	                    hdn.value      = value;
    	                    txt.value      = text;
    	                    aut.activeText = text;
    	                }    	        	                            
    	                txt.focus();
    	                //--------------------------------------------
    	            }     	
    	            aut.destroy();
    	        }
    	        //--------------------------------------------    	            	        
    	        lst.ondblclick = lst.onkeypress; 
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
        with (elKeyEvent) {             
            onfocus =  aut.item.onfocus;            
            //1 = autoComplete || 2 = button click
            if ( it.typeEvent == 1){
                onkeydown  = aut.item.onkeydown;
                onkeypress = aut.item.onkeypress;
            }else{               
                onkeyup = aut.item.onkeyup;
                var elButtonEvent = corp.GetElement(it.button)                
                elButtonEvent.idText = it.text;
                elButtonEvent.onclick =  aut.item.buttonClick;
            }
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
            var cp  =  new CORP();
            var prm =  txt.value;
	        cp.Ajax.url= aut.site + aut.activeURL + "&prm=" + cp.conv2hex(prm);
	        cp.Ajax.GetAjaxValue(aut.eventAttach);
	        //---------------------------------------------
        }
    }
    //---------------------------------------------
    CORP_AUTOCOMPLETE.prototype.destroy =  function () {
        
        if (aut.activeElement != null) { 
            var div =  corp.GetElement("div" + aut.activeElement)
            if ( div != null ) {
                div.innerHTML= '';
                div.style.display = "none"
            }
        }                    
        //-----------------------------------------
        aut.TmpDig=0;
        window.clearInterval(aut.activeTime);        
        if (event != null)
            if (event.keyCode == 27) {
                var el  = corp.GetElement(aut.activeElement);
                if (el != null) el.focus();
            }
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
        this.typeEvent = ( p == undefined ) ? 1 : 2; //1 = autoComplete || 2 = button click
        if (this.typeEvent == 1 )
            this.keySize = (k - 1)
        else
            this.keySize = k 
        this.button  = b;
    }
    //---------------------------------------------
    CORP_AUTOCOMPLETE_ITEM.prototype.onfocus = CORP_AUTOCOMPLETE.prototype.destroy;
    //---------------------------------------------
    CORP_AUTOCOMPLETE_ITEM.prototype.onkeydown = function() {     
        var it  = aut.items[this.id];
        var hdn = corp.GetElement(it.value);
        var txt = corp.GetElement(it.text);
        
        if (hdn.value != "" && event.keyCode == 8 && txt.value.length >= it.keySize)  txt.onkeypress(); 
                
    }
    //---------------------------------------------
    CORP_AUTOCOMPLETE_ITEM.prototype.onkeypress = function() { 
        aut.TmpDig=0;
        /* ------------------*/
        it  = aut.items[this.id];                
        hdn = corp.GetElement(it.value);
        hdn.value = "";
        /* ------------------*/        
        with(it) {                     
            aut.activeElement = text;
            aut.activeURL = url;
            aut.activeTypeEvent = typeEvent;
        }        
        /* ------------------------------------------------------*/        
        var el = corp.GetElement(it.text);        
        if ( el.value.length >= it.keySize && event.keyCode != 27 ){            
            window.clearInterval(aut.activeTime);
            aut.activeTime =  window.setInterval(aut.progress, 300);
            /*------------------------------------------------------*/
            c_msg.ShowMessage("carregando...", aut.activeElement)
            /*------------------------------------------------------*/
        }                
    }  
    //---------------------------------------------      
    CORP_AUTOCOMPLETE_ITEM.prototype.onkeyup = function() {
        var it  = aut.items[this.id];
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