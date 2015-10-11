$(function(){
	$(".scroll-up").on("click", function(){
		window.scrollBy(0, -$(window).height() * 0.9);
	});
	$(".scroll-down").on("click", function(){
		window.scrollBy(0, $(window).height() * 0.9);
	});
	document.all.contentFrame.height = 9999;
});
