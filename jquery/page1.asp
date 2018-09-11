<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <title>Demo</title>
<style>
a.test {
    font-weight: bold;
}
</style>
</head>
<body>
    <a href="http://jquery.com/">jQuery</a>
    <script src="jquery-1.10.2.js"></script>
    <script>
 
    $( document ).ready(function() {
        $( "a" ).click(function( event ) {
            alert( "The link will no longer take you to jquery.com" );
            event.preventDefault();
        });
    });
 
 $( "a" ).click(function( event ) {
 
    event.preventDefault();
 
    $( this ).hide( "slow" );
 
});

    </script>
</body>
</html>