if ( !this["Schema"] )
	Schema = function ( ) {};

Schema.Search = function ( init ) 
{
    var self = this;
    var fulltextindex = null;
    var parseTree = [];
	var metaOptions;
	var metaRegex;
    var options;
    
	var _ID;
	var _TITLE;
	var _FILENAME;
	var _FILESIZE;
	var _META;
	var _CONTENT;
	var defaultLocations;
	
	
	this.CreateExpressionTree = function ( obj, options )
	{
		for ( var i = 0; i < obj.length; i++ )
		{
			if ( obj[i].type == 'AND' )
			{
				if ( obj[i].what && obj[i].what.length > 0 )
				{
					obj[i].regex = Schema.Regex.Create ( obj[i].what, options );
					if ( obj[i].metaOptions || obj[i].metaOptions == 0 )
						obj[i].metaRegex = Schema.Regex.Create ( obj[i].what, obj[i].metaOptions );
					else
						obj[i].metaRegex = Schema.Regex.Create ( obj[i].what, metaOptions );
				}
			}
			else
			{
				for ( var j = 0; j < obj[i].expr.length; j++ )
				{
					if ( obj[i].expr[j].what && obj[i].expr[j].what.length > 0 )
					{
						obj[i].expr[j].regex = Schema.Regex.Create ( obj[i].expr[j].what, options );
						if ( obj[i].expr[j].metaOptions || obj[i].expr[j].metaOptions == 0 )
							obj[i].expr[j].metaRegex = Schema.Regex.Create ( obj[i].expr[j].what, obj[i].expr[j].metaOptions );
						else
							obj[i].expr[j].metaRegex = Schema.Regex.Create ( obj[i].expr[j].what, metaOptions );
					}
				}
			}
		}
		
		return ( obj );
	};



    this.Execute = function ( tokenStream )
    {    
        var result = [];

        parseTree = Parse ( tokenStream );
        
        var expressionTree = self.CreateExpressionTree ( parseTree, options );
        if ( expressionTree.length == 0 )
        	return ( result );
    
		for ( var i = 0; i < fulltextindex.length; i++ )
		{
			var item = fulltextindex[i];
			
			var itemIsResult = true;
			for ( var j = 0; j < expressionTree.length; j++ )
			{
				if ( !self.Search(expressionTree[j], item) )
				{
					itemIsResult = false;
					break;
				}
			}
			
			if ( itemIsResult )
				result.push ( i );
		}
        
        return ( result );
    };



	this.GetExpressions = function ( )
	{
		return ( parseTree );
	};

 	
	this.GetFulltextIndex = function ( )
 	{
 		return ( fulltextindex );
 	};
 	

	this.GetMetaOptions = function ( )
	{
		return ( metaOptions );
	};

 	
 	this.GetOptions = function ( )
 	{
 		return ( options );
 	};
 	

    var Initialize = function ( init )
    {
        fulltextindex = init.fulltextindex;
        if ( init.options )
        	options = init.options;
        else
        	options = Schema.Regex.Options.IgnoreCase;
        	
        if ( init.metaOptions )
        	metaOptions = init.metaOptions;
        else
        	metaOptions = Schema.Regex.Options.IgnoreCase;

        _ID = Schema.Search.Location.Id;
		_TITLE = Schema.Search.Location.Title;
		_FILENAME = Schema.Search.Location.Filename;
		_FILESIZE = Schema.Search.Location.Filesize;
		_META = Schema.Search.Location.Meta;
		_CONTENT = Schema.Search.Location.Content;
        
        if ( init.defaultLocations )
        	defaultLocations = init.defaultLocations;
        else
        	defaultLocations = [_TITLE, _CONTENT];
    };


    this.Search = function ( info, item )
	{
		if ( info.type == 'AND' )
		{
			return ( self.SearchAnd(info, item) );
		}
		else
		{
			for ( var i = 0; i < info.expr.length; i++ )
			{
				if ( self.SearchAnd(info.expr[i], item) )
					return ( true );
			}
			return ( false );
		}
	}; 
	
	
	this.SearchAnd = function ( info, item )
	{
		for ( var i = 0; i < info.where.length; i++ )
		{
			if ( info.where[i] == _ID )
			{
				if ( item[_ID] == info.what )
					return ( !info.invert );
			}
			else if ( info.where[i] == _META )
			{
				if ( self.SearchMeta(info, item) )
					return ( !info.invert );
			}
			else
			{
				if ( info.regex.exec(item[info.where[i]]) )
					return ( !info.invert );
			}
		}
		return ( info.invert );
	};
	
	
	this.SearchMeta = function ( info, item )
	{
		var metaexp = info.metaRegex ? info.metaRegex : metaRegex; 
	
		if ( !item[_META] || item[_META].length <= 0 )
			return ( false );
		
		if ( !info.constraints || info.constraints.length <= 0 )
		{
			if ( !info.what || info.what == "" )
			{
				for ( var name in item[_META] )
					return ( true );
				return ( false );
			}
			else
			{
				for ( var name in item[_META] )
				{
					if ( metaexp.exec(item[_META][name]) )
						return ( true ); 
				}
				return ( false );
			}
		}
		else
		{
			for ( var i = 0; i < info.constraints.length; i++ )
			{
				if ( !info.what || info.what == "" )
				{
					if ( item[_META][info.constraints[i]] )
						return ( true );
				}
				else
				{	
					var entry = item[_META][info.constraints[i]];
					if ( entry && metaexp.exec(entry) )
						return ( true );
				}
			} 
			return ( false );
		}
	};
 

    var Parse = function ( tokenStream )
    {   
        var tree = [];
        
        var tokens = Tokenize ( tokenStream );

        for ( var i = 0; i < tokens.length; i++ )
        {
        	if ( tokens[i].type == 'OR' )
        	{
        		var orToken = {'type': 'OR', 'expr': []};
        		for ( var j = 0; j < tokens[i].value.length; j++ )
        		{
        			var info = self.ParseToken(tokens[i].value[j]);
        			if ( info )
						orToken.expr.push ( info );
        		}
        			
        		tree.push ( orToken );
        	}
        	else
        	{
        		var info = self.ParseToken(tokens[i].value[0]);
        		if ( info )
        			tree.push ( info );
        	}
        }
        
        return ( tree );
	};
	


	var ParseMeta = function ( token, info )
	{
		info.where = [_META];
		token = token.substr(4);
				
		if ( token.substr(0,1) == '[' )
		{
			var inner = false;
			var lastIsEscape = false;
			
			var pos = 1;
			while ( pos < token.length )
			{
				var c = token.substr(pos,1);
				if ( c == ']' && !inner )
					break;
				else if ( c == '\\' && inner )
					lastIsEscape = !lastIsEscape;
				else if ( c == '"' && !lastIsEscape )
					inner = false;
				else
					lastIsEscape = false;
					
				pos++;
			}
			
			if ( pos >= token.length )
				throw _("Search.Error.ResultWriter.InvalidConstraints", token);


			var temp = token.substr(1, pos-1).split(",");
			var constraints = [];
			for ( var i = 0; i < temp.length; i++ )
			{
				var value = Prepare(temp[i]);
				if ( value != "" )
					constraints.push ( value ); 
			}
			
			if ( constraints.length > 0 )
				info.constraints = constraints;

			if ( token.substr(pos+1,1) == '(' )
				pos = ParseMetaOption ( token, pos+1, info );

			if ( !pos )
			{
				throw _("Search.Error.Parser.InvalidOptions", token);
				return ( null );
			}
				
			
			if ( token.substr(pos+1,1) == ':' )
				info.what = Prepare( token.substr(pos+2) );
		}
		else if ( token.substr(0,1) == '(' )
		{
			var pos = ParseMetaOption ( token, 0, info );
			
			if ( token.substr(pos+1,1) == ':' )
				info.what = Prepare( token.substr(pos+2) );
		}
		else
		{
			info.what = Prepare ( token.substr(1) );
		}
		
		return ( info );
	};
	

	var ParseMetaOption = function ( token, pos, info )
	{
		var len = token.length;
		var end;
		var st = pos;

		
		while ( pos < len )
		{
			if ( token.substr(pos,1) == ')' )
			{
				end = pos;
				break; 	
			}
			pos++;
		}
		
		if ( end )
		{
			var item = token.substr(st+1, end-st-1);
			if ( !item || item.length == 0 || parseInt(item) == NaN )
				throw _("Search.Error.Parser.InvalidOptions", token);
				
			info.metaOptions = parseInt(item); 
		}
		else
			throw _("Search.Error.Parser.InvalidOptions", token);
			
		return ( end );
	};
	
	
	
	this.ParseToken = function ( token )
	{
		var info = {'type': 'AND', 'invert': false};
		if ( token.substr(0,1) == '-' )
		{
			info.invert = true;
			token = token.substr(1);
		}
		
		if ( token.substr(0,5) == 'meta:' || token.substr(0,5) == 'meta[' || token.substr(0,5) == 'meta(' )
			ParseMeta ( token, info );
		else if ( token.substr(0,6) == 'title:' )
		{
			info.where = [_TITLE];
			info.what = Prepare ( token.substr(6) );
		}
		else if ( token.substr(0,3) == 'id:' )
		{
			info.where = [_ID];
			info.what = Prepare ( token.substr(3) );
		}
		else if ( token.substr(0,8) == 'content:' )
		{
			info.where = [_CONTENT];
			info.what = Prepare ( token.substr(8) );
		}
		else if ( token.substr(0,14) == 'x-set-options:' )
		{
			options = parseInt(token.substr(14));
			info = null;
		}
		else if ( token.substr(0,18) == 'x-set-metaoptions:' )
		{
			metaOptions = parseInt(token.substr(18));
			info = null;
		}
		else 
		{
			info.where = defaultLocations;
			info.what = Prepare ( token );
		}
		
		return ( info );
	};
	
	
	var Prepare = function ( input )
	{
		if ( !input )
			return ( "" );
			
		if ( input.substr(0,1) == '"' )
			input = input.substr(1);
		if ( input.substr(input.length-1, 1) == '"' )
			input = input.substr(0, input.length-1);
		return ( input );
	};


	this.SetMetaOptions = function ( opts )
	{
		metaOptions = opts;
	};


 	this.SetOptions = function ( opts ) 
	{
 		options = opts;
 	};


    var Tokenize = function ( tokenStream )
    {
    	tokenStream = tokenStream.replace(/\s+/g, " ");
    	var tokens = [];
    	var isOrActive = false;
    	var pos = 0;
    	
    	while ( pos < tokenStream.length )
    	{
    		var token = Schema.String.TokenizeNext ( tokenStream, pos );
    		if ( !token )
    			break;

			if ( token.value.toUpperCase() == "OR" )
			{
				if ( tokens.length > 0 )
				{
					tokens[tokens.length-1].type = 'OR';
					isOrActive = true;
				}
			}    			
			else
			{
				if ( tokens.length == 0 || !isOrActive )
					tokens.push ( {'type': 'AND', 'value': [token.value]} );
				else
					tokens[tokens.length-1].value.push(token.value);
					
				isOrActive = false;
			}

			pos = token.end+1;
    	}
    	
    	return ( tokens );
    };


    Initialize ( init );
};



Schema.Search.Location = function ( ) { };
Schema.Search.Location.Id = 0;
Schema.Search.Location.Title = 1;
Schema.Search.Location.Filename = 2;
Schema.Search.Location.Filesize = 3;
Schema.Search.Location.Meta = 4;
Schema.Search.Location.Content = 5;



Schema.Search.ResultWriter = function ( init )
{
	var self = this;
	var element = null;
	var elementName = null;
	this.expressionTree = [];
	var environment = 3;
	var pageSize = 10;
	var boundaryRegex = new RegExp ( "[" + Schema.Regex.Unicode.Separator + Schema.Regex.Unicode.Punctuation + "]" );
	var propertyName = "writer";
	var result;

	var CheckHighlight = function ( expr, section )
	{
    	var doHighlight = false;
		for ( var j = 0; j < section.length; j++ )
		{
			if ( ContainsValue(expr.where, section[j]) )
			{
				doHighlight = true;
				break;
			}
		}
    	
    	return ( doHighlight );
	};



    this.Clear = function ( )
    {
        if ( element )
            element.innerHTML = "";
    };



	var ContainsValue = function ( ary, value )
	{
		for ( var i = 0; i < ary.length; i++ )
		{
			if ( ary[i] == value )
				return ( true );
		}
		
		return ( false );
	};



    this.Execute = function ( resultSet )
    {
		if ( !element )
		{
			if ( !document.getElementById(elementName) )
				throw _("Search.Error.ResultWriter.NoElement");
				
			element = document.getElementById ( elementName );
		}

		result = resultSet;
    	self.expressionTree = search.CreateExpressionTree ( search.GetExpressions(), search.GetOptions() | Schema.Regex.Options.AllMatches );
    	
    	self.ShowPage ( 1 );
    };

    
    
	var GetEnvironment = function ( page )
	{
		var maxPage = GetMaxPage ( );

		if ( page <= environment+1 )
			return ( {'start': 1, 'end': Math.min(environment+environment+1, maxPage)} );
			

		if ( maxPage <= environment+environment+1 )
			return ( {'start': 1, 'end': maxPage} );
			
		if ( maxPage - page < environment )
			return ( {'start': Math.max(1, maxPage-environment-environment), 'end': maxPage} );
		
		return ( {'start': page-environment, 'end': Math.min(page+environment, maxPage)} );
	};
	    
	    
	var GetEnvironmentSize = function ( )
	{
		return ( environment );
	};


	var GetMaxPage = function ( )
	{
		if ( !result || result.length <= 0 )
			return ( 0 );
		if ( pageSize <= 0 )
			return ( 1 );
			
		var len = result.length;
		var count = 0;
		
		while ( len > 0 )
		{
			count++;
			len -= pageSize;
		}
		
		return ( count );
	};
	
	

    this.GetPageLink = function ( page, text )
    {
    	return ( '<a href="javascript:' + propertyName + ".ShowPage(" + page + ');" class="searchResultPagingLink">' + text + '</a>' ); 
    };

    
    
	var GetPageRegion = function ( page )
	{
		if ( pageSize <= 0 )
			return ( {'start': 0, 'end': result.length - 1} );
			
		var st = (page-1) * pageSize;
		var end = Math.min ( st+pageSize-1, result.length-1 );
		if ( st > result.length )
			return ( null );
			
		return ( {'start': st, 'end': end} );
	};
	
	
	
	this.GetPageSize = function ( )
	{
		return ( pageSize );
	};


    
	this.Highlight = function ( text, section )
	{
		var include = (section ? section : [Schema.Search.Location.Content]);
		
        for ( var i = 0; i < self.expressionTree.length; i++ )
        {
        	if ( self.expressionTree[i].type == 'AND' )
        	{
				if ( CheckHighlight(self.expressionTree[i], include) )
        			text = text.replace(self.expressionTree[i].regex, self.HighlightFunction);
        	}
        	else
        	{
        		for ( var j = 0; j < self.expressionTree[i].expr.length; j++ )
        		{
        			if ( CheckHighlight(self.expressionTree[i].expr[j], include) )
            			text = text.replace(self.expressionTree[i].expr[j].regex, self.HighlightFunction);
            	}
        	}
    	}
    	
    	return ( text );
	};
		
	
	
	this.HighlightFunction = function ( match )
	{
		var st = "";
		var end = "";
		if ( boundaryRegex.test(match.substr(0,1)) )
		{
			st = match.substr(0,1);
			match = match.substr(1);
		}
		if ( boundaryRegex.test(match.substr(match.length-1,1)) )
		{
			end = match.substr(match.length-1,1); 
			match = match.substr(0, match.length-1);
		}
			
		return ( st + '<span class="searchEntryHighlight">'+match+'</span>' + end );
	};



	var Initialize = function ( init )
	{
		elementName = init.element;
		
        if ( init.pageSize && parseInt(init.pageSize) > 0 )
            pageSize = parseInt(init.pageSize);
        if ( init.environment && parseInt(init.environment) > 0 )
			environment = parseInt(init.environment);
		if ( init.propertyName )
			propertyName = init.propertyName;
	};

	
	this.SetEnvironmentSize = function ( size )
	{
		environment = size;
	};
	
	
	this.SetPageSize = function ( size )
	{
		pageSize = size;
	};
	
	
    this.ShowPage = function ( page )
    {
    	self.Clear ( );
    
		self.WritePagingAndCount ( page );
    	if ( result && result.length > 0 )
    	{
    		var region = GetPageRegion ( page );
    		if ( region )
    		{
	    		var fulltextindex = search.GetFulltextIndex();
	    		for ( var i = region.start; i <= region.end; i++ )
	    			self.WriteResult ( fulltextindex[result[i]] );
	    			
	    		self.WritePagingAndCount ( page );
	    		window.scrollTo ( 0, 0 );
	    	}
    	}
    };


    this.WriteContext = function ( item )
    {
    	var _CONTENT = Schema.Search.Location.Content;
    	
    	var ix = item[_CONTENT].length;
    	
        for ( var i = 0; i < self.expressionTree.length; i++ )
        {
        	var temp;
        	if ( self.expressionTree[i].type == 'AND' )
        	{
        		if ( !ContainsValue(self.expressionTree[i].where, _CONTENT) )
        			continue;

            	var temp = item[_CONTENT].search ( self.expressionTree[i].regex );
            	if ( temp >= 0 )
            		ix = Math.min ( temp, ix );
            }
            else
            {
            	for ( var j = 0; j < self.expressionTree[i].expr.length; j++ )
            	{
            		if ( !ContainsValue(self.expressionTree[i].expr[j].where, _CONTENT) )
            			continue;
            			
            		var temp = item[_CONTENT].search ( self.expressionTree[i].expr[j].regex );
            		if ( temp >= 0 )
            			ix = Math.min ( temp, ix );
            	}
            }
        }
        
        if ( ix < 0 || ix >= item[_CONTENT].length )
        	ix = 0;
        
        if ( ix >= 0 )
        {
            var pos = ( ix < 100 ? 0 : ix - 100 );
            var st = pos;
            var ismiddle = true;
            
            for ( var j = ix-30; j >= pos; j-- )
            {
            	var c = item[_CONTENT].substr(j,1); 
                if ( c == '.' || c == '!' )
                {
                    st = j+1;
                    ismiddle = false;
                    break;
                }
                if ( item[_CONTENT].substr(j,1) == ' ' )
                    st = j;
            }

            var end = st + 200;
            if ( end >= item[_CONTENT].length )
            	end = item[_CONTENT].length;
            	
            for ( var j = end; j > ix+30; j-- )
            {
                if ( item[_CONTENT].substr(j,1) == " " )
                {
                    end = j;
                    break;
                }
            }
            
            var text = (ismiddle && st > 0 ? "... " : "") + item[_CONTENT].substr(st, end-st) + (end < item[_CONTENT].length ? "..." : "");

            var out = document.createElement ( "div" );
            out.className = "searchEntryContent";
            text = self.Highlight ( text ) 
            out.innerHTML = text;
			return ( out );
		}
    };


    this.WriteResult = function ( item )
    {
        var content = document.createElement ( "div" );
        content.className = "searchEntry";
		
		var temp;
		
		temp = self.WriteResultTitle ( item );
		if ( temp )
        	content.appendChild ( temp );
        
    	temp = self.WriteContext ( item );
    	if ( temp )
    		content.appendChild ( temp );
        
        temp = self.WriteResultFileInfo ( item );
        if ( temp )
        	content.appendChild ( temp );
        	
        
        element.appendChild ( content );
    };


	this.WriteResultFileInfo = function ( item )
	{    
        var info = document.createElement ( "div" );
        info.className = "searchEntryInfo";
        var html = item[Schema.Search.Location.Filename];
        
        var size = parseInt(item[Schema.Search.Location.Filesize]);
        if ( size != NaN )
        	html += " - " + Math.round(size/1024) + "k";
        	
        info.innerHTML = html;
        return ( info );
    };

    
    
    this.WriteResultTitle = function ( item )
    {
    	var loc = Schema.Search.Location;
    	
   		var title = document.createElement ( "div" );
        title.className = "searchEntryTitle";
        if ( self.NavigateTo )
			title.innerHTML = self.NavigateTo ( item );
        else
        	title.innerHTML = '<a href="' + item[loc.Filename] + '">' + self.Highlight((item[loc.Title] && item[loc.Title].length > 0) ? item[loc.Title] : item[loc.Filename]) + '</a>';
    	return ( title );
     };
    
    
    
    this.WritePagingAndCount = function ( page )
    {
    	var content = document.createElement ( "div" );
    	content.className = "searchResult";
    	
    	var table = document.createElement ( "table" );
    	table.border = 0;
    	table.cellPadding = 0;
    	table.cellSpacing = 0;
    	content.appendChild ( table );
    	
    	var tbody = document.createElement ( "tbody" );
    	table.appendChild ( tbody );
    	
    	var row = document.createElement ( "tr" );
    	tbody.appendChild ( row );
    	
    	var cell1 = document.createElement ( "td" );
    	cell1.className = "searchResultPaging";
    	row.appendChild ( cell1 );
    	
    	if ( result && result.length > 0 )
    	{
	    	var html = "";
	    	var range = GetEnvironment ( page );
	    	if ( page != range.start )
	    		html += self.GetPageLink ( page-1, _("Search.Result.Previous") ) + " " + _("Search.Result.Separator");
	    	
	    	if ( range.start != range.end )
	    	{
		    	for ( var i = range.start; i <= range.end; i++ )
		    	{
		    		if ( i == page )
		    			html += " " + '<span class="searchResultPagingActive">' + i + '</span>';
		    		else
		    			html += " " + self.GetPageLink ( i, i );
		    	}
		    }
	    		
	    	if ( page != range.end )
	    		html += " " + _("Search.Result.Separator") + " " + self.GetPageLink ( page+1, _("Search.Result.Next") );
	
			cell1.innerHTML = html;    	 
    	}
    	
    	var cell2 = document.createElement ( "td" );
    	cell2.className = "searchResultInfo";

		var region = GetPageRegion ( page );

    	if ( !result || result.length <= 0 || !region )
    		cell2.innerHTML = _("Search.Result.None");
    	else
    		cell2.innerHTML = _("Search.Result.Position", (region.start+1) + " - " + (region.end+1), result.length);

    	row.appendChild ( cell2 );
    	
    	element.appendChild ( content );
    };
    
	
	Initialize ( init );
};