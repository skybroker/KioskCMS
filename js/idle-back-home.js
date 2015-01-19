var COUNT_DOWN = 10;
var bContinue = false;
var nCountDown = COUNT_DOWN;
jQuery.fn.center = function() {
    var container = $(window);
    var top = -this.height() / 2;
    var left = -this.width() / 2;
    return this.css('position', 'fixed').css({ 'margin-left': left + 'px', 'margin-top': top + 'px', 'left': '50%', 'top': '300px' });
}
function countDown() {
	if(!bContinue)
		return;
	$(".count-down-seconds").text(nCountDown);
	if(nCountDown < 1) {
		top.location.href = "default.asp";
	}
	--nCountDown;
	setTimeout(function(){
		countDown();
	}, 1000);
}
function startCountDown() {
	nCountDown = COUNT_DOWN;
	bContinue = true;
	$(".idle-page-closing-panel").center().fadeIn(1000);
	countDown();
}
function stopCountDown() {
	$(".idle-page-closing-panel").hide();
	bContinue = false;
}
$(function(){
	$(document.body).append("<div class='idle-page-closing-panel'>用户无操作，<span class='count-down-seconds'>10</span>秒后返回主页!</div>");
	var idleTimeout = 30000;
	$(document).bind("idle.idleTimer",function(){
		if(!bContinue) {
			startCountDown();
		}
	});
	$(document).bind("active.idleTimer",function(){
		stopCountDown();
	});
	$.idleTimer(idleTimeout);
});
