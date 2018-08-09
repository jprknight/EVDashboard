<script language="javascript" type="text/javascript" src="./js/tablefilter.js"></script>
<link rel="stylesheet" type="text/css" href="./css/filtergrid.css" />
<script src="./js/sortabletable.js" language="javascript" type="text/javascript"></script>
<script src="./js/tfAdapter.sortabletable.js" language="javascript" type="text/javascript"></script>
<script type="text/javascript" src="./js/mootools.js"></script>
<script type="text/javascript" src="./js/calendar.rc4.js"></script>
<link rel="stylesheet" type="text/css" href="./css/calendar.css" />
<link rel="stylesheet" type="text/css" href="./css/TFExt_FiltersRowVisibility.css" />
<link rel="stylesheet" type="text/css" href="./css/TFExt_ColsVisibility.css" />
<link rel="stylesheet" type="text/css" href="./css/TFExt_ColsResizer.css" />
<style type="text/css" media="screen">

	@import "filtergrid.css";
	/*====================================================
		- html elements
	=====================================================*/
	
	table{ border:1px solid #ccc; font-size:85%; }
	th{
		background: url(./images/nav_light.png) center center repeat-x;
		border-left:1px solid #C7C7C7;
		padding:5px; color:#fff; height:20px;
	}
	td{ padding:5px; border-bottom:1px solid #ccc; border-right:1px solid #ccc; }
	pre{ margin:5px; padding:5px; background-color:#f4f4f4; border:1px solid #ccc; }
	
	
	/*====================================================
		- TF grid classes
	=====================================================*/
	.myLoader{ 
		position:absolute; padding: 15px 0 15px 0;
		margin:100px 0 0 15px; width:200px; 
		z-index:1000; font-size:12px; font-weight:bold;
		border:1px solid #666; background:#ffffcc; 
		vertical-align:middle;
		
	/* custom class for Columns Visibility Manager Extension */
	.colsMngContainer{
		position:absolute;
		display:none;
		border:2px outset #999;
		height:auto; width:250px;
		background:#E4FAE4;
		margin:18px 0 0 0; z-index:10000;
		padding:10px 10px 10px 10px;
		text-align:left; font-size:12px;
	}
</style>