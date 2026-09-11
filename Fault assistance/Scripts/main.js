'use strict';

document.addEventListener('DOMContentLoaded', function () {
    var nav = document.querySelector('.contentLayoutOne')
      , toggleButton = document.querySelector('.togglemenu')

    if (toggleButton) {
        toggleButton.addEventListener('click', function (event) {
            event.preventDefault()
			
          	if(nav.classList) {
				nav.classList.toggle('closed')
			} else {
				if(hasClass(nav,'closed')) {
					removeClass(nav,'closed')
				} else {
					addClass(nav,'closed')
				}
			}

        })
    }
})

function hasClass(el, className)
{
    if (el.classList)
        return el.classList.contains(className);
    return !!el.className.match(new RegExp('(\\s|^)' + className + '(\\s|$)'));
}

function addClass(el, className)
{
    if (el.classList)
        el.classList.add(className)
    else if (!hasClass(el, className))
        el.className += " " + className;
}

function removeClass(el, className)
{
    if (el.classList)
        el.classList.remove(className)
    else if (hasClass(el, className))
    {
        var reg = new RegExp('(\\s|^)' + className + '(\\s|$)');
        el.className = el.className.replace(reg, ' ');
    }
}