function Schema ( ) { }; 

// =============================================================================
//                         Cross browser XMLHttpRequest
// =============================================================================

if ( typeof XMLHttpRequest == "undefined" ) 
{
    XMLHttpRequest = function ( ) {
        try { return new ActiveXObject("Msxml2.XMLHTTP.6.0") } catch (e) { }
        try { return new ActiveXObject("Msxml2.XMLHTTP.3.0") } catch (e) { }
        try { return new ActiveXObject("Msxml2.XMLHTTP") } catch (e) { }
        try { return new ActiveXObject("Microsoft.XMLHTTP") } catch (e) { }
        return ( undefined );
    }
};


// =============================================================================
//                                 EVENT SYSTEM
// =============================================================================

Schema.Event = function ( ) { };

Schema.Event.AddListener = function ( obj, functionName, functionCode, ctx, tag )
{
    if ( !ctx )
        ctx = obj;
        
    var fn = null;
    if ( obj.addEventListener )
    { 
        fn = function ( e ) { functionCode.call(ctx, e, tag); e.cancelBubble=true; if(e.stopPropagation){e.stopPropagation();} };
        obj.addEventListener ( functionName, fn, false );
    }
    else if ( obj.attachEvent )
    {
        fn = function ( e ) { var event = window.event; functionCode.call(ctx, event, tag); event.cancelBubble = true; if(event.stopPropagation){event.stopPropagation();} };
        obj.attachEvent ( "on" + functionName, fn );
    }
    return ( fn );
};


Schema.Event.OnDomContentLoaded = function ( fn )
{
	if ( !fn )
		return;

	if ( document.addEventListener )
		document.addEventListener ( "DOMContentLoaded", fn, false );
	else
	{
		document.onreadystatechange = function() 
		{
			if ( document.readyState == "interactive" || document.readyState == "complete" )
				fn ( );
		}
	}
}


Schema.Event.RemoveListener = function ( obj, functionName, functionCode )
{
    if ( obj.removeEventListener )
        obj.removeEventListener ( functionName, functionCode, true );
    else if ( obj.detachEvent )
        obj.detachEvent ( "on" + functionName, functionCode );
}



// =============================================================================
//                               LANGUAGE FEATURES
// =============================================================================

Schema.I18N = function ( ) { };
Schema.I18N.currentLang = 'en';
Schema.I18N.currentGuiLang = 'en';
Schema.I18N.data = { 'all': {}, 'de': {}, 'en': {}, 'fr': {}, 'ja': {} };

Schema.I18N.GetTranslation = function ( index, gui )
{
    var dict = gui ? Schema.I18N.data[Schema.I18N.currentGuiLang] : Schema.I18N.data[Schema.I18N.currentLang];
    var result;
    
    if ( dict && dict[index] )
        result = dict[index];
        
    if ( !result && ((!gui && Schema.I18N.currentLang != 'all') || (gui && Schema.I18N.currentGuiLang != 'sys')) )
    {
        dict = Schema.I18N.data['all'];
        if ( dict && dict[index] )
            result = dict[index];
    }

    if ( result )
    {
        if ( arguments.length > 1 )
        {
            var call = [];
            call.push ( result );
            for ( var i = 2; i < arguments.length; i++ )
                call.push ( arguments[i] );
                
            return ( Schema.String.Format.apply(this,call) );
        }
        else
            return ( result );
    }
    
    return ( "" );
}


function _ ( index )
{
	var args = [];
	args.push ( index );
	args.push ( false );
	for ( var i = 1; i < arguments.length; i++ )
		args.push ( arguments[i] );
		
    return ( Schema.I18N.GetTranslation.apply(this,args) );
}


function __ ( index )
{
	var args = [];
	args.push ( index );
	args.push ( true );
	for ( var i = 1; i < arguments.length; i++ )
		args.push ( arguments[i] );
		
    return ( Schema.I18N.GetTranslation.apply(this,args) );
}
 


// =============================================================================
//                                REGEXP EXTENSION
// =============================================================================

Schema.Regex = function ( ) { };
Schema.Regex.Create = function ( value, options )
{
	var regexOpt = "";
	if ( (options & Schema.Regex.Options.AllMatches) != 0 )
		regexOpt += "g";
	if ( (options & Schema.Regex.Options.IgnoreCase) != 0 )
		regexOpt += "i";
	if ( (options & Schema.Regex.Options.Multiline) != 0 )
		regexOpt += "m";

	var sp = Schema.Regex.Unicode.Separator + Schema.Regex.Unicode.Punctuation;
			
	value = Schema.Regex.Encode ( value );
	
	if ( (options & Schema.Regex.Options.Exact) != 0 )
	{
		string = "^" + value + "$";
	}
	else if ( (options & Schema.Regex.Options.WholeWord) != 0 )
	{
		string = "^" + value + "(?=[" + sp + "]+)|[" + sp + "]" + value + "(?=[" + sp + "])|[" + sp + "]" + value + "$|^" + value + "$";  
	}
	else if ( (options & Schema.Regex.Options.StartWord) != 0 )
	{
		string = "^" + value + "|[" + sp + "]" + value;  
	}
	else
		string = value;
		 
	return ( new RegExp(string, regexOpt) );
}


Schema.Regex.Encode = function ( value )
{
	return ( value.replace(/([.*+?^${}()|[\]\/\\])/g, '\\$1').replace(/\s+/g, "\\s") );
}



Schema.Regex.Options = function ( ) { };
Schema.Regex.Options.None = 0;
Schema.Regex.Options.IgnoreCase = 1;
Schema.Regex.Options.WholeWord = 2;
Schema.Regex.Options.StartWord = 4;
Schema.Regex.Options.Exact = 8;
Schema.Regex.Options.AllMatches = 16;
Schema.Regex.Options.Multiline = 32;

Schema.Regex.Unicode = function ( ) { };
Schema.Regex.Unicode.Punctuation = "\\u0021-\\u0023\\u0025-\\u002A\\u002C-\\u002F\\u003A\\u003B\\u003F\\u0040\\u005B-\\u005D\\u005F\\u007B\\u007D\\u00A1\\u00AB\\u00B7\\u00BB\\u00BF\\u037E\\u0387\\u055A-\\u055F\\u0589\\u058A\\u05BE\\u05C0\\u05C3\\u05C6\\u05F3\\u05F4\\u0609\\u060A\\u060C\\u060D\\u061B\\u061E\\u061F\\u066A-\\u066D\\u06D4\\u0700-\\u070D\\u07F7-\\u07F9\\u0964\\u0965\\u0970\\u0DF4\\u0E4F\\u0E5A\\u0E5B\\u0F04-\\u0F12\\u0F3A-\\u0F3D\\u0F85\\u0FD0-\\u0FD4\\u104A-\\u104F\\u10FB\\u1361-\\u1368\\u166D\\u166E\\u169B\\u169C\\u16EB-\\u16ED\\u1735\\u1736\\u17D4-\\u17D6\\u17D8-\\u17DA\\u1800-\\u180A\\u1944\\u1945\\u19DE\\u19DF\\u1A1E\\u1A1F\\u1B5A-\\u1B60\\u1C3B-\\u1C3F\\u1C7E\\u1C7F\\u2010-\\u2027\\u2030-\\u2043\\u2045-\\u2051\\u2053-\\u205E\\u207D\\u207E\\u208D\\u208E\\u2329\\u232A\\u2768-\\u2775\\u27C5\\u27C6\\u27E6-\\u27EF\\u2983-\\u2998\\u29D8-\\u29DB\\u29FC\\u29FD\\u2CF9-\\u2CFC\\u2CFE\\u2CFF\\u2E00-\\u2E2E\\u2E30\\u3001-\\u3003\\u3008-\\u3011\\u3014-\\u301F\\u3030\\u303D\\u30A0\\u30FB\\uA60D-\\uA60F\\uA673\\uA67E\\uA874-\\uA877\\uA8CE\\uA8CF\\uA92E\\uA92F\\uA95F\\uAA5C-\\uAA5F\\uFD3E\\uFD3F\\uFE10-\\uFE19\\uFE30-\\uFE52\\uFE54-\\uFE61\\uFE63\\uFE68\\uFE6A\\uFE6B\\uFF01-\\uFF03\\uFF05-\\uFF0A\\uFF0C-\\uFF0F\\uFF1A\\uFF1B\\uFF1F\\uFF20\\uFF3B-\\uFF3D\\uFF3F\\uFF5B\\uFF5D\\uFF5F-\\uFF65";
Schema.Regex.Unicode.Separator   = "\\u0020\\u00A0\\u1680\\u180E\\u2000-\\u200A\\u2028\\u2029\\u202F\\u205F\\u3000";




// =============================================================================
//                               REQUEST HANDLER
// =============================================================================

Schema.Web = function ( ) { };

Schema.Web.Navigate = function ( url, init )
{
	var getstring = "";
	if ( init.get )
	{
		for ( var name in init.get )
		{
			if ( getstring != "" )
				getstring += "&";
			getstring += encodeURI(name) + "=" + encodeURI(init.get[name]);
		}
	}

	if ( init.post )
	{
		var form = document.createElement ( "form" );
		form.method = 'POST';
		form.action = url + ((getstring != "") ? "?" + getstring : "");
		for ( var name in init.post )
		{
			var input = Schema.DOM.CreateInput("hidden");
			input.name = name;
			input.value = init.post[name];
			form.appendChild ( input );
		}
		document.getElementsByTagName("body")[0].appendChild ( form );
		form.submit();
	}
	else
	{
		var loc = url;
		if ( getstring != "" )
			loc += "?" + getstring;
		location.href = loc;
	}
}
 

Schema.Web.Request = function ( url, init )
{
    this.callback = (init) ? init.callback : null;
    this.request = null;
    this.params = (init && init.params) ? init.params : [];
    this.async = (init) ? (init.async === true) : true;
    this.url = url;
	this.method = (init && init.method) ? init.method : 'POST';

    this.Execute = function ( )
    {
        var body = '';
        if ( this.params )
        {
            for ( key in this.params )
            {
                if ( body != "" )
                    body += "&";
                
                value = this.params[key];
                if ( value && typeof(value) == "string" )
                {
	                value = value.replace ( /%/g, "%25" );
	                value = value.replace ( /&/g, "%26" );
	                value = value.replace ( /\\r/g, "%0D" );
	                value = value.replace ( /\\n/g, "%0A" );
					value = value.replace ( /\+/g, "%2B" );
	                value = value.replace ( / /g, "%20" );
	            }
                body += key + "=" + value;
            }
        }

        this.request = new XMLHttpRequest ( );
        this.request.open ( this.method, this.url, this.async );
		if ( this.method.toLowerCase() == 'post' )
        	this.request.setRequestHeader ( "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8" );
        this.request.setRequestHeader ( "Content-Length", body.length );
        
        var req = this.request;
        var self = this;
        if ( this.async )
        {
            this.request.onreadystatechange = function ( ) 
			{ 
				if ( req.readyState == 4 ) 
				{ 
					if ( self.callback ) 
						self.callback(self); 
				}
			};
            this.request.send ( body );
        }
        else
        {
            this.request.send ( body );
            return ( this.request.responseText );                
        }
    };
    
    
    
    this.GetJsonObjects = function ( )
    {
        var text = this.GetResponseText ( );
        if ( text == "" )
            return ( null );
        else            
            return ( Schema.Json.Evaluate(text) ); 
    };
    
    
    this.GetResponseText = function ( )
    {
        return ( this.request.responseText );
    };
};


// =============================================================================
//                                STRING FUNCTIONS
// =============================================================================


Schema.String = function ( ) { };
Schema.String.Format = function ( org )
{
    for ( var i = 1; i < arguments.length; i++ )
    {
        org = org.replace ( new RegExp("\\{" + (i-1) + "\\}", "g"), arguments[i] );
    }
        
    return ( org );
}

Schema.String.IsNumber = function ( obj )
{
	return ( obj.replace( /[\+-]?\s*\d+/, "") = "" );
}



Schema.String.Tokenize = function ( tokenStream )
{
	tokenStream = tokenStream.replace(/\s+/g, " ");
	var tokens = [];
	var pos = 0;
	
	while ( pos < tokenStream.length )
	{
		var token = Schema.String.TokenizeNext ( tokenStream, pos );
		if ( !token )
			break;

		tokens.push ( token.value );
		pos = token.end+1;
	}
	
	return ( tokens );
};


Schema.String.TokenizeNext = function ( tokenStream, pos )
{
    while ( pos < tokenStream.length && tokenStream.substr(pos, 1) == " " )
        pos++;
    
    if ( pos >= tokenStream.length )
    	return ( null );
    
	var st = pos;
	var inner = false;
	var lastIsEscape = false;
		
    while ( pos < tokenStream.length )
	{
		var c = tokenStream.substr ( pos, 1 );

		if ( c == ' ' && !inner )
		{
			pos--;
			break;
		}
		else if ( c == '\\' && inner )
			lastIsEscape = !lastIsEscape;
		else if ( c == '"' && !lastIsEscape )
			inner = !inner;
		else
			lastIsEscape = false;
		
		pos++;				  
	}
		
    return ( {"value": tokenStream.substr(st,pos-st+1), "end": pos+1} );
};

Schema.String.Trim = function(str)
{
	if(!str || str == "")
		return str;
	
	var res = "";
	var i;
	
	// look from beginning until a non-whitespace char
	for(i = 0; i < str.length; ++i)
	{
		if(str.charAt(i) == ' ')
			continue;
		else
			break;
	}
	res = str.substring(i);
	
	// look from end until a non-whitespace char
	for(i = res.length-1; i >= 0; --i)
	{
		if(res.charAt(i) == ' ')
			continue;
		else
			break;
	}
	return res.substring(0, i+1);
}



// =============================================================================
//                               UTILITY FUNCTIONS
// =============================================================================

Schema.Utils = function ( ) { };

Schema.Utils.Apply = function ( source, target )
{
	if ( source && target )
	{
		for ( var name in source )
			target[name] = source[name];
	}
	return ( target );
};


Schema.Utils.ClearSelection = function ( )
{
	var sel;
	if ( document.selection && document.selection.empty ) 
		document.selection.empty();
	else if ( window.getSelection ) 
	{
		sel = window.getSelection();
		if ( sel && sel.removeAllRanges )
			sel.removeAllRanges();
	}
}


Schema.Utils.Contains = function ( obj, name )
{
	if ( obj.length && obj.length > 0 )
	{
		for ( var i = 0; i < obj.length; i++ )
		{
			if ( obj[i] == name )
				return ( true );
		}
	}
	
	return ( false );
}


Schema.Utils.Clone = function ( obj )
{
	var result = {};
	if ( !obj )
		return ( result );
	for ( var name in obj )
		result[name] = obj[name];
		
	return ( result );
}

Schema.Utils.GetClientSize = function ( )
{
    var clientX, clientY;
    if (self.innerHeight)
    {
    	clientX = self.innerWidth;
    	clientY = self.innerHeight;
    }
    else if (document.documentElement && document.documentElement.clientHeight)
    {
    	clientX = document.documentElement.clientWidth;
    	clientY = document.documentElement.clientHeight;
    }
    else if (document.body)
    {
    	clientX = document.body.clientWidth;
    	clientY = document.body.clientHeight;
    }
    
    return ( {"Width":clientX, "Height":clientY} );
};


Schema.Utils.GetPosition = function ( elem )
{
	var left = 0;
	var top = 0;
	do
	{
		if ( elem.offsetParent )
		{
			left += (elem.offsetLeft || 0);
			top += (elem.offsetTop || 0);
		}
	}
	while ( elem = elem.offsetParent )
	return ( {'x': left, 'y':top} );
}


Schema.Utils.SplitObject = function ( obj, names )
{
	var result = {};
	for ( var key in obj )
	{
		if ( Schema.Utils.Contains(names, key) )
			result[key] = obj[key];
	}

	for ( var i = 0; i < names.length; i++ )
	{
		if ( obj[names[i]] )
			delete obj[names[i]];
	}
	return ( result );
}




// =============================================================================
//                              JSON HANDLING
// =============================================================================


Schema.Json = function ( ) { };

Schema.Json.Evaluate = function ( text )
{
    if ( /^[\],:{}\s]*$/.test(text.replace(/\\["\\\/bfnrtu]/g, '@').replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g, ']').replace(/(?:^|:|,)(?:\s*\[)+/g, ''))) 
        return ( eval('(' + text + ')') );
    else
        return ( null );
}



// =============================================================================
//                                DOM HELPERS
// =============================================================================

Schema.DOM = function ( ) { };

Schema.DOM.AppendChild = function ( node, parent )
{
    parent.appendChild ( node );
}


Schema.DOM.Create = function ( init )
{
    var result = new Array ( );
    for ( var i = 0; i < init.length; i++ )
    {
        if ( init[i] && init[i].length > 0 )
        {
            var elem = document.createElement ( init[i][0] );
            result.push ( elem );
            
            if ( init[i].length > 1 && init[i][1] )
                result.push ( Schema.DOM.Create(init[i][1]) );
        }
    }
    
    return ( result );
};


Schema.DOM.CreateEx = function ( init )
{
    var result = new Array ( );
    for ( var i = 0; i < init.length; i++ )
    {
        var elem = document.createElement ( init[i].name );
        result.push ( elem );
        
        if ( init[i].attributes )
        {
            for ( var key in init[i].attributes )
                elem.setAttribute ( key, init[i].attributes[key] );
        }
        if ( init[i].style )
        {
            for ( var key in init[i].attributes )
                elem.style.setAttribute ( key, init[i].style[key] );
        }
        if ( init[i].text )
            elem.innerText = init[i].text;
        if ( init[i].html )
            elem.innerHTML = init[i].html;
        if ( init[i].children )
            result.push ( Schema.DOM.CreateEx(init[i].children) );
    }
    
    return ( result );
};


Schema.DOM.CreateInput = function ( type )
{
    var elem = document.createElement ( "input" );
    elem.setAttribute ( "type", type );
    return ( elem );
   
};


Schema.DOM.InsertAfter = function ( node, org )
{
    if ( !node || !org )
        return;
        
    if ( org.nextSibling )
        org.parentNode.insertBefore ( node, org.nextSibling );
    else
        org.parentNode.appendChild ( node );
};


Schema.DOM.InsertBefore = function ( node, org )
{
    if ( node && org )
        org.parentNode.insertBefore ( node, org ); 
};


Schema.DOM.PrependChild = function ( node, parent )
{
    if ( parent.firstChild )
        Schema.DOM.InsertBefore ( node, parent.firstChild ); 
    else
        Schema.DOM.AppendChild ( node, parent );
};


Schema.DOM.RemoveChildren = function ( node )
{
    while ( node.firstChild )
        Schema.DOM.RemoveElement ( node.firstChild );
};


Schema.DOM.RemoveElement = function ( node )
{
	if ( typeof(node) == 'string' )
		node = document.getElementById(node);
		
    if ( node && node.parentNode )
        node.parentNode.removeChild ( node );
};


Schema.DOM.SetStyle = function ( node, style )
{
    if ( node && style )
    {
        for ( var entry in style )
            node.style[String(entry)] = style[entry];
    }
};

Schema.DOM.InnerHTML = function (node)
{
	if(node.nodeType == 3)   // text node
		return node.nodeValue;
	else
	{
		var res = "";
		var n = node.firstChild;
		while(n)
		{
			if(n.nodeType == 3)
				res += n.nodeValue;
			else if(n.nodeType == 1)  // element node
				res += Schema.DOM.InnerHTML(n);
				
			n = n.nextSibling;
		}
		return res;
	}
};

// =============================================================================
//                                DOM HELPERS
// =============================================================================

Schema.Cookies = function ( ) { };
Schema.Cookies.Get = function ( key )
{
	Schema.Cookies.Parse ( );
	return ( Schema.Cookies._obj[key] );
}

Schema.Cookies.Parse = function ( )
{
	if ( Schema.Cookies._obj )
		return;
		
	Schema.Cookies._obj = {};
	if ( document.cookie )
	{
		var elements = document.cookie.split ( ";" );
		for ( var i = 0; i < elements.length; i++ )
		{
			var pair = elements[i].split("=");
			Schema.Cookies._obj[Schema.String.Trim(pair[0])] = unescape(Schema.String.Trim(pair[1]));
		}
	}
}

Schema.Cookies.Set = function ( key, value, expiredays )
{
	if( !key )
		return;
		
	Schema.Cookies.Parse ( );
	Schema.Cookies._obj[key]= value;
	
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=key+"=" +escape(value)+ ((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}