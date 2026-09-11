// ----------------------------------------------------------------------------
// LEXICON SCRIPT
// ----------------------------------------------------------------------------

Schema.Tooltips = function ( ) { }
Schema.Tooltips.last = null;

Schema.Tooltips.Hide = function ( toolTipId, elemId )
{
	var toolTip = document.getElementById ( toolTipId );
    toolTip.style.display = "none";
    toolTip = null;
}

Schema.Tooltips.Show = function ( toolTipId, elemId )
{
	// Clear tooltip, if one is already active...
	if ( Schema.Tooltips.last )
	{
		Schema.Tooltips.last.style.display = "none";
		Schema.Tooltips.last = null;
	}
	
    var toolTip = document.getElementById ( toolTipId );
	var elem = document.getElementById ( elemId );
    var page = document.getElementsByTagName("body")[0];
    if ( toolTip.parentNode != page )
    {
    	Schema.DOM.RemoveElement ( toolTip );
    	Schema.DOM.PrependChild ( toolTip, page );
    }
    toolTip.style.position = "absolute";
    toolTip.style.zIndex = 100;
    toolTip.style.display = "block";
    
    // Set tooltip position...
	var size = Schema.Utils.GetClientSize();

	var pos = Schema.Utils.GetPosition(elem);
	var pageY = Schema.Utils.GetPosition(page).y;
	var top = pos.y - pageY + 10 + elem.offsetHeight;
	
		
	var left = pos.x;
	if ( (pos.x + toolTip.offsetWidth) > (size.Width - 25) )
		left = size.Width - 25 - toolTip.offsetWidth;
	if ( left < 0 )
	    left = 0;

	toolTip.style.top = top + "px";
	toolTip.style.left = left + "px";
}


function selectOption ( x ) 
{
	if(x == "nothing")
		return;
	else
		document.location.href = x;
}

function VideoSpanHover(id, span_id)
{
	var img = document.getElementById(id);
	var play = document.getElementById(span_id);
	if (play!= null)
	{
		play.style.width = img.width + "px";
		play.style.height = img.height + "px";
		play.style.position = "absolute";		
		var titlePage = document.getElementsByClassName("titlePageContent");
		if (titlePage.length!= 0)
		{
			play.style.left = img.offsetLeft  + "px";
		}
	}
}

function VideoOnClick (src_mp4, src_ogg, src_webm, id, autoplay, loop, preload, mute, span_id, fallback_text)
{
		var videoElement = document.createElement("video");
		videoElement.setAttribute("controls", "controls");
		videoElement.setAttribute("style", "max-width: 100%");
		//Set VideoAttributes
		SetVideoAttributes(videoElement, autoplay, loop, preload, mute);
		//source mp4, ogg, webm
		CreateVideoSource(src_mp4, videoElement, "video/mp4" );
		CreateVideoSource(src_ogg, videoElement, "video/ogg" );
		CreateVideoSource(src_webm, videoElement, "video/webm" );
		//fallback
		var fallback = document.createElement ("div");
		fallback.setAttribute("class", "fallback");
		fallback.innerHTML = fallback_text;
		videoElement.appendChild (fallback);
		var img = document.getElementById(id);
		VideoRemoveSpan(span_id, img);
		img.parentNode.replaceChild(videoElement, img);
}

function VideoRemoveSpan (id, element)
{
	var span = document.getElementById(id);
	if (span!= null)
	{
		element.parentNode.removeChild(span);
	}	
}

function CreateVideoSource (src, element, mimetype)
{
	if (src!="")
	{
		var srcElement = document.createElement("source");
		srcElement.setAttribute("src", src);
		srcElement.setAttribute("type", mimetype);	
		element.appendChild(srcElement);
	}
}

function SetVideoAttributes (video, autoplay, loop, preload, mute)
{
	if (autoplay != "")
	{
		video.setAttribute("autoplay", autoplay);
	}
	if (loop != "")
	{
		video.setAttribute("loop", loop);
	}
	if (mute != "")
	{
		video.setAttribute("muted", mute);
	}
	if (preload != "")
	{	
		video.setAttribute("preload", preload);
	}	
}
