jQuery.fn.center = function() {
    var container = $(window);
    var top = -this.height() / 2;
    var left = -this.width() / 2;
    return this.css('position', 'fixed').css({ 'margin-left': left + 'px', 'margin-top': top + 'px', 'left': '50%', 'top': '300px' });
}

