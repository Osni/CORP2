//<![CDATA[

/*

	RECURSOS APLICADOS PARA CORP VERSÃO 2

*/


mapkeys = {
    UP_ARROW: 38,
    DOWN_ARROW: 40,
    PG_DOWN: 33,
    PG_UP: 34,
    END: 35,
    HOME: 36,
    ESC: 27,
    ENTER: 13
};

jsFilter = {  

    fields: [],
    btnApl: null,
    btnNext: null,
    btnPrev: null,
    btnLast: null,
    btnFirst: null,
    tab: null,
    
    idxRowInit: 2,
    idxActualRow: 2,    
    blnSendForm: false,
    activeCell: null,    
    
    postForm: function (evt){        
        //---------------------------------------
        var el = crossBrowser.srcElem(evt);
        var key = crossBrowser.key(evt);
        //---------------------------------------
        if(key == mapkeys.ENTER && !jsFilter.blnSendForm) 
            jsFilter.btnApl.onclick();                  
        //---------------------------------------
    }, 
    
    clearFields: function() {
        for (e in jsFilter.fields)  
            jsFilter.fields[e].value = "";
    },
        
    elementsAreNotEmpty: function(){         
        var els = jsFilter.fields;
        for (var i = 0; i < els.length; i++){            
            if(manageform.trim(els[i].value) != '')
                return true;
        }
        return false;
    },
        
    changeFocus: function(evt){  
        var el = crossBrowser.srcElem(evt);        
        if (evt.type == 'focus')            
            el.className = 'filterFocus';
        else            
            el.className = '';        
    },    
    
    mouseOverOut: function (evt){
        evt = crossBrowser.events(evt);
        var cell = crossBrowser.srcElem(evt);
        var row = cell.parentNode;
        if (evt.type == "mouseover"){
            //---------------------------------------------
            if (jsFilter.tab.rows[jsFilter.idxActualRow]) // remove seleção anterior
                jsFilter.selectRow(false, jsFilter.tab.rows[jsFilter.idxActualRow]);
            //---------------------------------------------    
			jsFilter.idxActualRow = row.rowIndex;
			jsFilter.activeCell = cell;
			jsFilter.blnSendForm = true;
			jsFilter.selectRow(true, row);
		}else{
			jsFilter.idxActualRow = jsFilter.idxRowInit - 1;
			jsFilter.activeCell = null;
			jsFilter.selectRow(false, row);
		};
    },
    
    selectRow: function(paint, row){        
        var css = '';
        if(paint) 
            css = "selTD";
        else
            if(row.rowIndex % 2 != 0)
                css = "lightTD";
            else
                css = "darkTD";
        row.className = css;		
    }, 
        
    setCellsFocus: function() {        
        for (var i = jsFilter.idxRowInit; i < jsFilter.tab.rows.length; i++) {            
            cell = jsFilter.tab.rows[i].cells[0];
            cell.onmouseover = jsFilter.mouseOverOut;
            cell.onmouseout  = jsFilter.mouseOverOut;
        }
    },            
            
    moveRows: function(evt){
		//-------------------------------------------------
		var key = crossBrowser.key(evt);
		var i = 0;
		//-------------------------------------------------		
		if (key == mapkeys.UP_ARROW) {	// move up
			//=============================================
			jsFilter.idxActualRow--;
			jsFilter.blnSendForm = true;
			
			if (jsFilter.idxActualRow <= jsFilter.idxRowInit)
				jsFilter.idxActualRow = jsFilter.idxRowInit;
			//=============================================	
		} else if (key == mapkeys.DOWN_ARROW) {	// move down
			//=============================================
			jsFilter.idxActualRow++;
			jsFilter.blnSendForm = true;
			
			if (jsFilter.idxActualRow >= jsFilter.tab.rows.length)
				jsFilter.idxActualRow = jsFilter.tab.rows.length - 1;
		    //=============================================
        } else if (key == mapkeys.PG_UP && evt.ctrlKey) {// CTRL + PAGE-UP => Next Page
            //=============================================
            jsFilter.btnNext.onclick();
            //=============================================
        } else if (key == mapkeys.PG_DOWN && evt.ctrlKey) {// CTRL + PAGE-DOWN => Previous Page
            //=============================================
            jsFilter.btnPrev.onclick();  
            //=============================================
        } else if (key == mapkeys.END && evt.ctrlKey) { // CTRL + END => Last Page
            //=============================================
            jsFilter.btnLast.onclick();  
            //=============================================
        } else if (key == mapkeys.HOME && evt.ctrlKey) { // CTRL + HOME => First Page
            //=============================================
            jsFilter.btnFirst.onclick();  
		    	//=============================================
		} else if (key == mapkeys.ESC) {               // Reset filter
		    //=============================================
            setTimeout(function() {            
                jsFilter.clearFields(); 
                jsFilter.btnApl.onclick();
            }, 100);
            //=============================================
		} else if (key == mapkeys.ENTER && jsFilter.blnSendForm == true){ // Select item and post to a Basic Form
				//=============================================			
				jsFilter.activeCell.parentNode.onclick();				
				//=============================================
		}else jsFilter.blnSendForm = false;
		//-------------------------------------------------
		for (var i = jsFilter.idxRowInit; i < jsFilter.tab.rows.length; i++){
			if (i == jsFilter.idxActualRow){
				jsFilter.activeCell = jsFilter.tab.rows[i].cells[0];
				jsFilter.selectRow(true, jsFilter.tab.rows[i]);
			}else
				jsFilter.selectRow(false, jsFilter.tab.rows[i]);			
				
		};						
		//-------------------------------------------------
    },
    
    init: function(fields, idxRowInit) {      
		//---------------------------------------------------------
        jsFilter.idxRowInit = idxRowInit == undefined ? jsFilter.idxRowInit : idxRowInit;
        jsFilter.idxActualRow = jsFilter.idxRowInit - 1;
        //---------------------------------------------------------
        crossBrowser.AttEvent("keydown", function(evt){ jsFilter.moveRows(evt);	})
        //---------------------------------------------------------        
        for (var el in fields){
            //-----------------------------------------------------
            var field = crossBrowser.elem(fields[el]);
            //-----------------------------------------------------
            if(field != null){
                jsFilter.fields[el] = field;
                field.idx = parseInt(el, 10);
                //-----------------------------------------------------
                crossBrowser.AttEventElement(field, "keydown",
                                            function(evt){
                                               jsFilter.postForm(evt);
                                            });
                //-----------------------------------------------------
                crossBrowser.AttEventElement(field, "focus",
                                            function(evt){                                            
                                               jsFilter.changeFocus(evt, '');
                                            });           
                //-----------------------------------------------------
                crossBrowser.AttEventElement(field, "blur",
                                            function(evt){                                            
                                               jsFilter.changeFocus(evt, '');
                                            });           
                //-----------------------------------------------------
            }
        }        
        //---------------------------------
        jsFilter.tab = crossBrowser.elem('filter_table');        
        jsFilter.btnApl  =  crossBrowser.elem("filter_button_aplicar");
        jsFilter.btnPrev = crossBrowser.elem("filter_button_previous");
        jsFilter.btnNext = crossBrowser.elem("filter_button_next");
        jsFilter.btnLast = crossBrowser.elem("filter_button_last");
        jsFilter.btnFirst = crossBrowser.elem("filter_button_first");
        //---------------------------------
        jsFilter.setCellsFocus();
        //---------------------------------
        setTimeout(function(){
            jsFilter.fields[0].focus();    
        }, 100);
        //---------------------------------
    }
    
};
//]]>
