function k(t, prox)
{	
	if (event.keyCode == 13 )						
	{
		var f  =  document.forms[0];
		if (f!=null)
			f.submit();
	}
}	
function SetFoco(){
	document.getElementById("txtAcessoUsuario").focus();
}
function Enviar(obj,enviar)
{
	eval(obj + ".submit()")
}

function MenuPrincipal(obj, Flag) 
{    	
	switch (Flag) 
	{ 
	case 1: 
			obj.style.cursor = 'hand';
			obj.style.background = "#F7F6F3" ;       
			obj.style.borderColor ="#524EA5";    			
			if ( obj.alt=="")
				obj.style.borderWidth = "0px";	
			obj.alt="Ativo";		        			
		break; 
	case 2 : 	
	        obj.alt="";        
			obj.style.cursor      = 'hand';            
			obj.style.background  = "#DFE6EF";      
			obj.style.borderColor = "#524EA5";
			obj.style.borderWidth = "1px";			 	
		break; 
//	default : 
//			obj.style.cursor = 'hand';            
//			obj.style.background = "D7D4D4";                  
	} 	
}
function ClickMenu(Obj, td)
{			
	//----------------------------------------------------------------
	var len =  td.parentNode.parentNode.parentNode.rows.length;
	for ( i=0; i < len ; i++ ) {
		var t =  td.parentNode.parentNode.parentNode.rows[i].childNodes[0];
		t.style.borderWidth = "0px";	
		t.title = "";	
	}	
	td.style.borderWidth = "1px";			 	
	td.title = "ativo";
	//----------------------------------------------------------------
	menu.hMenu.value= Obj	 ;		
	menu.hForm.value= Obj ;			
	menu.action=Obj;
	menu.method="post";
	menu.target="CORPO";
	menu.submit();	
	
}
