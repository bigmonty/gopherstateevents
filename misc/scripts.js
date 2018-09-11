function pop(url, x, y){
	mywindow=window.open(url,'','width='+x+',height='+y+',scrollbars=yes');
	mywindow.focus();
}

function openThis(url, x, y){
		features='top=100,left=100,width='+x+',height='+y+',scrollbars=yes,toolbar=yes,menubar=yes,scrollbars=yes,resizable=yes,location=yes,directories=yes,status=yes';
		mywindow=window.open(url,'profile',features);
		mywindow.focus();	
}

function openThis2(url, x, y){
		features='top=100,left=100,width='+x+',height='+y+',scrollbars=yes,toolbar=yes,menubar=yes,scrollbars=yes,resizable=yes,location=yes,directories=yes,status=yes';
		mywindow=window.open(url,'profile2',features);
		mywindow.focus();	
}

