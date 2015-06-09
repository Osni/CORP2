
rptView = {
    lHeight: 0,
    lProcHeight: 0,
    it: null,
    hs: true,
    x: null,
    y: null,
    w: null,
    speed: 30,
    bFlag: true,
    cssText: "",
    cssText1: "",
    Exportando: false,
    colunasDetalhe: null,
    colsParam: [],
    
    ShowCaixa: function(){
        if (rptView.hs == true) {
            rptView.lProcHeight += rptView.speed;
            if (rptView.lProcHeight >= rptView.lHeight) {
               rptView.y.src   = "http://10.0.0.238:9090/CORPNET/WizardRpt/imagens/fechar.gif";
               rptView.y.title = "Ocultar";
               rptView.w.style.visibility = "visible";
	           window.clearInterval(rptView.it);
            }
        }else{
            rptView.lProcHeight -= rptView.speed;
            if (rptView.lProcHeight <= 30) {
	           rptView.y.src = "http://10.0.0.238:9090/CORPNET/WizardRpt/imagens/abrir.gif";
               rptView.y.title = "Exibir";
               rptView.w.style.visibility = "hidden";
               rptView.lProcHeight = 0;
	           window.clearInterval(rptView.it);
            }
        }
        rptView.x.style.height = rptView.lProcHeight + "px";
    },

    ocultaexibeparam: function (id, h, t) {
        rptView.y = t;
        rptView.x = document.getElementById(id);
        rptView.w = document.getElementById("rptParamTitulo");
        rptView.lProcHeight = rptView.x.offsetHeight;
        rptView.lHeight = h;
        rptView.hs = !rptView.hs;
        rptView.it= window.setInterval("rptView.ShowCaixa()",  40);
        return false;
    },

    movepage: function (direcao) {    
        var page_atual = document.getElementById("hdnPage");
        var max_page   = document.getElementById("hdnMaxPage");
        
        //
        if (page_atual == null) return false;
        //
               
        var p = document.getElementById("rptPagina" + page_atual.value) 
        var text = document.getElementById("txtNumPag");
        
        var lblMaxPage = document.getElementById("lblMaxPage");
        lblMaxPage.innerText = max_page.value;
    	
        //
        switch (direcao) {
            case -1 : //MoveFirst
		        // Oculta página atual;
		        if (page_atual.value != 1) {
		            p.style.display = "none";
			        page_atual.value = 1;	
		        }			
		        break;
        
            case  0 : //MovePrevious
                if (parseInt(page_atual.value) > 1) {                
			        p.style.display = "none";
			        page_atual.value--;
		        }
                break;
                            
            case  1 : //MoveNext
                if (parseInt(page_atual.value) < parseInt(max_page.value)) {
			        p.style.display = "none";
			        page_atual.value++;
		        }
                break;
            
            case  2 : //MoveLast 
		        if (page_atual.value != max_page.value) {
			        p.style.display = "none";
			        page_atual.value = max_page.value;
		        }
		        break;
        
        }
        //        		
        p = document.getElementById("rptPagina" + page_atual.value);
        p.style.display = "";
        text.value = page_atual.value;
        //
        return false
        //				
    },

    gotoPage: function (e) {
        var page_atual = document.getElementById("hdnPage");
        var max_page   = document.getElementById("hdnMaxPage");
    	
        //
        if (page_atual == null) return false;
        //
    	
        //Verifica se página está no intervalo válido
        if (parseInt(e.value) < 1 || parseInt(e.value) > parseInt(max_page.value)) return false;
    			
        //Oculta página atual
        p = document.getElementById("rptPagina" + page_atual.value);
        p.style.display = "none";
    	
        //Exibe página selecionada
        p = document.getElementById("rptPagina" + e.value);
        p.style.display = "";
        //Atualiza página atual
        page_atual.value = e.value;
        //	
        return false
    },

    imprimir: function (paginado) { 
        ///
        page_atual = document.getElementById("hdnPage");
        if (page_atual != null) {
            if (paginado) {
	            var page_atual = document.getElementById("hdnPage");
	            var max_page   = document.getElementById("hdnMaxPage");
    		
	            for (var i=1; i <=  max_page.value; i++) { 
		            //Exibe todas as páginas
		            var p = document.getElementById("rptPagina" + i);
		            p.style.display = ""			
		            //			
	            }
            }	
            //		    
            window.print();
            //	
        }		
        return false;
    },

    previouspage: function () { 
        window.history.go(-1);
        return false;    
    },

    Esperando: function (){    
      setTimeout(function() { 
          var id_el = rptView.Exportando == true ? "imgExportar" : "imgShowRel";
          var el = crossBrowser.elem(id_el);
          el.title = "Processando..."
          el.src = "http://10.0.0.238:9090/CORPNET2/imagens/esperando.gif";
          el.disabled = true;
      },100);  
    },         
        
    Retornando: function() {
        if (rptView.Exportando) {
            setTimeout(function() { 
                var el = document.getElementById("imgExportar");
	            el.src="http://10.0.0.238:9090/CORPNET2/imagens/excel.gif";
	            el.disabled=false;	  
	            manageform.deActivateButtons(false);                      
                rptView.Exportando = false;                
            }, 300);
        }
    }, 
        
    ChecaOperador: function() {
        var cpos = rptView.colsParam;
        var blnErro = false;
        for(var i = 0; i < cpos.length; i++) {
            //-------------------------------------------
            if(manageform.campos_checar == undefined)
                manageform.campos_checar = new Array(); 
            //-------------------------------------------
            var cpoParam = crossBrowser.elem(cpos[i][0]);
            var valida = cpoParam.getAttribute("valida");
            //-------------------------------------------
            valida = ((valida == null) ? "" : valida);
            //-------------------------------------------
            if(valida == "false"){
               var cpoOper = crossBrowser.elem(cpos[i][1]);
               if(manageform.trim(cpoParam.value) != '' && cpoOper.value == ''){
                   blnErro = true;
                   manageform.campos_checar[manageform.campos_checar.length] = new Array(cpoParam.id, 'Informe qual o crit' + String.fromCharCode(233) + 'rio.');
                }
            }
        }
        //-----------------------------------
        if(blnErro) {
            manageform.Adverte();
            return false;
        }else{
            return true;
        }
        //-----------------------------------
    },
   
    Initialize: function(){             

        crossBrowser.testIfIE();
        
        if (crossBrowser.isIE){
            crossBrowser.AttEvent("focusin", rptView.Retornando);
        }else{         
            document.addEventListener('DOMFocusIn', rptView.Retornando, false);            
            document.addEventListener('focus', rptView.Retornando, true);
        }
        //-------------------------------------
        rptView.SelecionaCampos("INPUT");        
        rptView.SelecionaCampos("SELECT");        
        //-------------------------------------    
    },
    
    SelecionaCampos: function(tagName) {
        var cpos = document.getElementsByTagName(tagName);
        for(var i = 0; i < cpos.length; i++) {                       
            var id_operador = cpos[i].getAttribute('rpt_param');
            var id_param = cpos[i].id;
            
            if (id_operador){
                rptView.colsParam[rptView.colsParam.length] = new Array(id_param, id_operador);
            }                            
        }
    }
    
}

crossBrowser.AttEventElement(window, "load", function(evt){
    try {
	     var tb = document.getElementsByTagName("table");        
	     for (i=0;i<tb.length;i++){                  
		    if(tb[i].id.toLowerCase() == "rpttabdetalhe"){                
			    var rws =  tb[i].rows;
			    for (r=0;r<rws.length;r++){                  
				    rws[r].onclick =  function(evt){   
					    try{					
						    var evt = evt || window.event;
						    var trg = evt.target || evt.srcElement;
						    var rws = trg.parentNode;
						    for (r=0;r<rws.cells.length;r++){        
							    rws.cells[r].style.backgroundColor="#EEF5F8";
						    } 
					    }catch(e){}
				    }
				    rws[r].ondblclick =  function(evt){      
					    try{					
						    var evt = evt || window.event;
						    var trg = evt.target || evt.srcElement;
						    var rws = trg.parentNode;
						    for (r=0;r<rws.cells.length;r++){        
							    rws.cells[r].style.backgroundColor="";
						    } 
					    }catch(e){}
				    }                
			    }
		    }
	     }        		
    }catch(e){}
});

