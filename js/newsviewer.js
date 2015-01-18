$(function(){
	$(".scroll-up").on("click", function(){
		window.scrollBy(0, -$(window).height() * 0.9);
	});
	$(".scroll-down").on("click", function(){
		window.scrollBy(0, $(window).height() * 0.9);
	});
	$('a').removeAttr('href');
	$("[onclick]:not(.kiosk-element)").removeAttr("onclick");
});
