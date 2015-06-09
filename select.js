select_search = {

	validChars:  function (key){			
			
		var valid = /[0-9]|[a-z]|[ ]/ig.test(String.fromCharCode(key));						
		
		if (valid == false){
			
			var keyLow = String.fromCharCode(key).toLowerCase().charCodeAt(0);
			
			switch(parseInt(keyLow)) {
				case 224: case 225: case 226: case 227:  // à á â ã
				case 228: case 231:	case 232: case 233:  // ä ç è é
				case 234: case 235: case 236: case 237:  // ê ë ì í
				case 238: case 239: case 241: case 242:  // î ï ñ ò
				case 243: case 244: case 245: case 249:  // ó ô õ ù
				case 250: case 251: case 252:			 // ú û ü
					valid = true;
					break;
			}
		}
		
		return valid;
		//return /[0-9]|[a-z]|[äàáâãëèéêïìíîöòóôõüùúûñç ]/ig.test(String.fromCharCode(key));
    },	    	    
    
    isIE: false,
    
    testIfIE : function() { 
			var w = window.navigator;
			select_search.isIE = (/MSIE/ig.test(w.appVersion) || /MSIE/ig.test(w.userAgent));
			return select_search.isIE;
    },
    
    pesquisaSelect: function (evt) {
		
			evt = evt ? evt : window.event;
			var cbo = evt.target || evt.srcElement;
			var tecla = evt.which || evt.keyCode ;			
			var textoaux = "";

			if (tecla == 9 || tecla == 13) return true; //FireFox					
			
			if (select_search.validChars(tecla)) {
														
			    textoaux = cbo.texto;
		   		textoaux += String.fromCharCode(tecla);
		   				   				   		
				select_search.selectItems(cbo, textoaux);
			
			}else if (tecla == 27){
			    cbo.texto = "";
			    cbo.options[0].selected = true;
			}
			return false;
	},
	
	deleteKey: function(evt) {
			evt = evt ? evt : window.event;
			var cbo = evt.target || evt.srcElement;
			var tecla = evt.which || evt.keyCode ;
			
			if (tecla == 46){
				if(cbo.texto.length != 0){
				    cbo.texto = cbo.texto.substring(0, cbo.texto.length - 1);
				    select_search.selectItems(cbo, cbo.texto);
				}else
				    cbo.options[0].selected = true;
			}else if(tecla == 9 || tecla == 13) 
				return true;
	},
	
	selectItems: function (cbo, textoaux){
			
			var items = cbo.options;
			var aproximada = cbo.getAttribute("aproximada") ? new Boolean(cbo.getAttribute("aproximada").toLowerCase().replace("false", "")) : false;
			var ignoreCase = cbo.getAttribute("ignorecase") ? new Boolean(cbo.getAttribute("ignorecase").toLowerCase().replace("false", "")) : true;
			var found;
			var text;

			for (var i = 0; i < items.length; i++) {
			
				//Busca Conteúdo
				if (ignoreCase != false){
    				text = aproximada ? items[i].text.substr(0, textoaux.length).toUpperCase() :
    					        items[i].text.toUpperCase();
    				found = text.match(textoaux.toUpperCase());
    			}else{
    				text = aproximada ? items[i].text.substr(0, textoaux.length) : items[i].text;
    				found = text.match(textoaux);
    			}
    			    			
				if(found) { // Achou						
					items[i].selected = true;					
					cbo.texto = textoaux;
					break;
				}
			}				
	},					
					
    Initialize: function () { 
		   var cbos = document.getElementsByTagName("select");
		   for (var i = 0; i < cbos.length; i++) {
		        var noapply = cbos[i].getAttribute("noapply") ? new Boolean(cbos[i].getAttribute("noapply").toLowerCase().replace("false", "")) : false;
		        if (noapply == false) {
				    cbos[i].onkeypress = select_search.pesquisaSelect;
				    cbos[i].onkeydown = select_search.deleteKey;
				    cbos[i].idx = i;
				    cbos[i].texto = "";
				}
		   }        
		   select_search.testIfIE();
    }
}