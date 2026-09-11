Schema.Tree = function ( init )
{
    var self = this;

	var contextMenuHandler = null;
	var element = null;
    var focusedNode = null;
	var hasIcons;
    var isRendered = false;
	var loader = null;
    var nodes = [];
	var visualization = {};
    
    
    this.AppendNode = function ( init )
    {
		var node;
		if ( init instanceof Schema.TreeNode )
			node = init;
		else
			node = new Schema.TreeNode ( init );
			
		node.InternalInit ( this, null, 1 );
		node.Render ( );
		nodes.push ( node );
			
		Schema.DOM.AppendChild ( node.GetElement(), this.GetElement() );
		return ( node );
    };
    
    
    this.ClearNodes = function ( )
    {
        for ( var i = 0; i < nodes.length; i++ )
            nodes[i].Destroy ( );
            
        nodes.splice ( 0, nodes.length );
		focusedNode = null;
    };
    
    
	this.ExpandToNode = function ( node )
	{
		if ( node && node instanceof Schema.TreeNode )
		{
			var store = [];
			var temp = node;
			while ( temp )
			{
				store.push ( temp );
				temp = temp.GetParent();
			}
			
			for ( var i = store.length-1; i >= 0; i-- )
				store[i].Expand ( );
		}
	};


	this.GetContextMenuHandler = function ( )
	{
		return ( contextMenuHandler );
	};
	
	
	this.GetElement = function ( )
	{
		return ( element );
	};
	
	
	this.GetFocusedNode = function ( )
	{
		return ( focusedNode );
	};
	
	
    this.GetNodeById = function ( id )
    {
        for ( var i = 0; i < nodes.length; i++ )
        {
            var search = nodes[i].GetNodeById ( id );
            if ( search )
                return ( search );
        }
        
        return ( null );
    };


	this.GetLoader = function ( )
	{
		return ( loader );
	};


	this.GetNodes = function ( )
	{
		return ( nodes );
	};


	this.GetVisualization = function ( )
	{
		return ( visualization ); 
	};
	
	
	this.HasIcons = function ( )
	{
		return ( hasIcons );
	};
	

	var Initialize = function ( init )
	{
		// Initialize render element...
		if ( typeof(init.element) == "string" )
			element = document.getElementById ( init.element );
		else
			element = init.element;
		element.innerHTML = "";
			
		if ( !element )
			throw "No element given for tree!";
			
		// Other data...
		contextMenuHandler = init.contextMenuHandler ? init.contextMenuHandler : null;
		loader             = init.loader ? init.loader : null;
		visualization      = init.visualization ? init.visualization : Schema.TreeVisualization;
		hasIcons           = init.hasIcons ? true : false;
	};


	this.InsertNode = function ( pos, init )
    {
        if ( pos < 0 || pos >= nodes.length )
            return;
            
		var node;
		if ( init instanceof Schema.TreeNode )
			node = init;
		else
			node = new Schema.TreeNode ( init );
			
		node.InternalInit ( this, null, 1 );
		node.Render ( );
		
        if ( pos == 0 )
            Schema.DOM.PrependChild ( node.GetElement(), this.GetElement() );
        else
            Schema.DOM.InsertBefore ( node.GetElement(), nodes[pos].GetElement() );

        nodes.splice ( pos, 0, node );

        return ( node );
    };


	this.LoadStatic = function ( data )
	{
		this.ClearNodes ( );
		if ( data.length > 0 )
		{
			for ( var i = 0; i < data.length; i++ )
			{
				var node = this.AppendNode ( data[i] );
				node.LoadStatic ( data[i].nodes );
			}
		}
		else
		{
			var node = this.AppendNode ( data );
			node.LoadStatic ( data.nodes );
		}
	};
	

    this.OnFocusedNodeChanged = function ( treeNode )
    {
    };


	this.OnHighlightNode = function ( treeNode, select )
	{
		if ( select )
			treeNode.GetSelectionElement().style.backgroundColor = "#B0B0FF";
		else
		    treeNode.GetSelectionElement().style.backgroundColor = "Transparent";
	};        

		
    this.RemoveNode = function ( treeNode )
    {
        if ( treeNode && treeNode.GetParent() == this )
        {
            treeNode.Destroy ( );
            var index = nodes.indexOf ( treeNode );
            if ( index >= 0 )
                nodes.splice ( index, 1 );
        }
        else if ( treeNode )
            treeNode.GetParent().RemoveNode ( treeNode );
            
        if ( treeNode == focusedNode )
        {
            focusedNode = null;
            this.OnFocusedNodeChanged ( null );
        }
    };
    

	this.Reload = function ( )
	{
		if ( loader )
		{
			this.ClearNodes ( );
			loader.Request ( null );
		}
	};


    this.Render = function ( )
    {
        if ( isRendered )
            return;

        isRendered = true;
                                
        if ( loader )
        {
            loader.Request ( null );
        }
        else
        {
            for ( var i = 0; i < nodes.length; i++ )
                nodes[i].Render ( );
        }            
    };
    
    Initialize ( init );
};

Schema.Tree.IconPath = "resources/tree/";



Schema.TreeNode = function ( init )
{
    var self = this;
    
	var attributes;
    var childContainer = null;
    var collapsed = true;
    var element = null;
	var href = "";
    var icon = null;
	var id;
    var imgElement;
    var isRendered = false;
    var leaf = false;
    var level = 0;
    var nodes = [];
    var parent = null;
    var selectionElement;
    var text = "";
    var tree = null;
    
    
    this.AppendNode = function ( init )
    {
		var node;
		if ( init instanceof Schema.TreeNode )
			node = init;
		else
			node = new Schema.TreeNode ( init );
			
		nodes.push ( node );
		node.InternalInit ( tree, this, level+1 );
		node.Render ( );
		
		Schema.DOM.AppendChild ( node.GetElement(), childContainer );
			
		leaf = false;
		UpdateToggleImage ( );

		return ( node );
    };


    this.ClearNodes = function ( )
    {
        for ( var i = 0; i < nodes.length; i++ )
            nodes[i].Destroy ( );
            
        nodes.splice ( 0, nodes.length );
    };


    this.ClearSelect = function ( )
    {
    	tree.OnHighlightNode ( this, false );
    };
	

    this.Collapse = function ( )
    {
        collapsed = true;
        childContainer.style.display = 'none';

		UpdateToggleImage ( );        
    };


    this.Destroy = function ( )
    {
        // Destroy own panel...
        var elem = this.GetElement ( );
        if ( elem )
        {
            Schema.DOM.RemoveElement ( elem );
            elem = null;
        }
    };


    this.Expand = function ( )
    {
		for ( var i = 0; i < nodes.length; i++ )
			nodes[i].Collapse ( );

		var loader = tree.GetLoader();
        if ( loader )
            loader.Request ( this );

		collapsed = false;
        childContainer.style.display = 'block';

		UpdateToggleImage ( );
    };
    
    
    this.GetElement = function ( )
    {
        return ( element );
    };
	
	
	this.GetSelectionElement = function ( )
	{
		return ( selectionElement );
	};
	
	this.GetId = function ( )
	{
		return ( id );
	}
	
	
	this.GetHRef = function ( )
	{
		return ( href );
	}
    
    
    this.GetNodeById = function ( _id )
    {
        if ( id == _id )
            return ( this );
            
        for ( var i = 0; i < nodes.length; i++ )
        {
            var search = nodes[i].GetNodeById ( _id );
            if ( search )
                return ( search );
        }
        
        return ( null );
    };
    
	
	this.GetParent = function ( )
	{
		return ( parent );
	};
    
	
    var Initialize = function ( init )
    {
		if ( !init ) 
			init = {};
			
        text = init.text ? init.text : "&#160;";
        id = init.id ? init.id : null;        
        leaf = init.leaf ? true : false;
        icon = init.icon ? init.icon : null;
        tree = init.tree;
        level = init.level ? init.level : 0;
        attributes = init.attributes ? init.attributes : [];
		href = init.href ? init.href : "";

        element = document.createElement ( "div" );
    };


    this.InsertNode = function ( pos, init )
    {
        if ( pos < 0 || pos >= nodes.length )
            return;
            
		var node;
		if ( init instanceof Schema.TreeNode )
			node = init;
		else
			node = new Schema.TreeNode ( init );
			
		node.InternalInit ( tree, this, level+1 );
		node.Render ( );

        if ( pos == null )
            Schema.DOM.PrependChild ( node.GetElement(), childContainer );
        else
        {
            var point = nodes[pos];
            Schema.DOM.InsertBefore ( node.GetElement(), point.GetElement() );
        }

        nodes.splice ( pos, 0, node );
        
        // imgElement.style.display = 'block';
        leaf = false;
        UpdateToggleImage ( );
        
        return ( node );
    };
	
	
	
	this.InternalInit = function ( _tree, _parent, _level )
	{
		if ( isRendered )
			throw "Tree node already bound!";
			
		tree = _tree;
		parent = _parent;
		level = _level;
	};
	

	this.IsLeaf = function ( )
	{
		if ( tree.GetLoader() )
			return ( leaf );
		else
			return ( nodes.length == 0 );
	};
	
	this.LoadStatic = function ( data )
	{
		this.ClearNodes ( );
		if ( data )
		{
			for ( var i = 0; i < data.length; i++ )
			{
				var node = this.AppendNode ( data[i] );
				node.LoadStatic ( data[i].nodes );
			}
		}
	};
    
    
    this.RemoveNode = function ( treeNode )
    {
        if ( treeNode && treeNode.GetParent() == this )
        {
            treeNode.Destroy ( );
            var index = nodes.indexOf ( treeNode );
            if ( index >= 0 )
                nodes.splice ( index, 1 );
        }
        else if ( treeNode )
            treeNode.parent.RemoveNode ( treeNode );
    };

    
	this.Render = function ( )
	{
		if ( isRendered )
			return;
		isRendered = true;			
			
        var table = element.appendChild ( document.createElement("table") );
        table.cellPadding = 1;
        table.cellSpacing = 0;
        // table.style.whiteSpace = "nowrap";
                
        var tbody = table.appendChild ( document.createElement("tbody") );
        var row = tbody.appendChild ( document.createElement("tr") );
        
        var cell;
        for ( var i = 0; i < level-1; i++ )
        {
            cell = row.appendChild ( document.createElement("td") );
			cell.className = 'treeNodeIconEmpty';

            var img = cell.appendChild ( new Image() );
			img.src = tree.GetVisualization().Empty;
            img.style.width = "16px";
            img.style.height = "16px";
        }

        // Render icon cell...        
        cell = row.appendChild ( document.createElement("td") );
        cell.className = 'treeNodeIcon';
        
        imgElement = new Image ( );
        if ( self.IsLeaf() )
        {   
            imgElement.src = tree.GetVisualization().LeafIcon;
            imgElement.style.height = "16px";
            imgElement.style.width = "16px";
        }
        else
            imgElement.src = tree.GetVisualization().CollapsedIcon;
            
        var toggleFn = self.Toggle;
        Schema.Event.AddListener ( imgElement, 'click', function(e) { toggleFn.call(self); }, self ); 
            
        cell.appendChild ( imgElement );


        // Render image cell...
		if ( tree.HasIcons() )
		{
			cell = row.appendChild ( document.createElement("td") );
			cell.className = "treeNodeIcon";
			 
			var img = new Image ( );
			img.src = icon ? icon : tree.GetVisualization().Empty;
			img.style.height = "16px";
			img.style.width  = "16px";
			cell.appendChild ( img );
		}
        
        
        // Render text cell...
        cell = row.appendChild ( document.createElement("td") );
		cell.className = "treeNodeText";
        cell.innerHTML = text;
        selectionElement = cell;
    
        var selectFn = self.Select;
        var contextFn = self.ShowContextMenu; 
        Schema.Event.AddListener ( cell, 'click', function(e) { selectFn.call(self); }, self );
        Schema.Event.AddListener ( cell, 'contextmenu', function(e) { contextFn.call(self, e.clientX, e.clientY); return ( false ); }, self );
        
        childContainer = document.createElement ( "div" );
        childContainer.style.display = 'none'; 
        element.appendChild ( childContainer );
        
        for ( var i = 0; i < nodes.length; i++ )
        	nodes[i].Render ( );
	};

      
    this.Select = function ( internal )
    {
        if ( tree.focusedNode == this )
            return;
            
        if ( tree.focusedNode )
            tree.focusedNode.ClearSelect ( );

        tree.focusedNode = this;
        tree.OnHighlightNode ( this, true ); 
        if ( !internal )
        	tree.OnFocusedNodeChanged ( this );
    };


    this.ShowContextMenu = function ( x, y )
    {
		var handler = tree.GetContextMenuHandler ( );
		if ( handler )
            handler.Execute ( this, x, y );
    };
    
    
    this.Toggle = function ( )
    {
        if ( collapsed )
            this.Expand ( );
        else
            this.Collapse ( );
    };


	var UpdateToggleImage = function ( )
	{
		if ( imgElement )
		{
			var img = new Image ( );
	        if ( self.IsLeaf() )
	            img.src = tree.GetVisualization().LeafIcon;
	        else if ( collapsed )
	            img.src = tree.GetVisualization().CollapsedIcon;
	        else
				img.src = tree.GetVisualization().ExpandedIcon;
			
			imgElement.src = img.src;
		}
	};    
    
    Initialize ( init );
};



Schema.TreeLoader = function ( url, tree )
{
    this.url = url;
    this.tree = tree;
    this.autoExpandFirstLevel = false;
    
    
    this.GetNodeInfo = function ( id )
    {
        var request = new Schema.Request ( this.url, {'async':false} );
        request.params["id"] = id;
        request.Execute ( );
        
        return ( request.GetJsonObjects() );
    };
    
    
    this.Request = function ( treeNode )
    {
        var request = new Schema.Request ( this.url, {'async':true, 'callback': this.OnLoad} );
        request.treeNode = treeNode;
        request.tree = tree;
        
        if ( treeNode )
        {
            request.params["id"] = treeNode.id;
            request.autoExpand = false;
        }
        else
            request.autoExpand = this.autoExpandFirstLevel;
        
        request.Execute ( );
    };
    
    
    this.OnLoad = function ( request )
    {
        var parent = (request.treeNode) ? request.treeNode : request.tree;
		if ( parent )
		{
			parent.ClearNodes ( );
			
			var objs = request.GetJsonObjects ( );
			if ( objs )
			{
				for ( var i = 0; i < objs.length; i++ )
				{
					var node = parent.AppendNode ( objs[i] );
					if ( request.autoExpand )
						node.Expand ( );
				}
			}
		}
        
        
        if ( this.OnLoadComplete )
            this.OnLoadComplete ( request );    
    };
}


Schema.TreeVisualization = 
{
    'CollapsedIcon'           : Schema.Tree.IconPath + "node_collapsed.gif",
	'Empty'                   : Schema.Tree.IconPath + "empty.gif",
	'ExpandedIcon'            : Schema.Tree.IconPath + "node_expanded.gif",
	'LeafIcon'                : Schema.Tree.IconPath + "empty.gif"
};