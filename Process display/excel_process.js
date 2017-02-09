//So we first load the alasaql script to read from the excel. And when it is loaded, we run the rest of the code. It is a hack, but it works.
var script = document.createElement('script');
script.src = 'https://cdn.jsdelivr.net/alasql/0.3/alasql.min.js'
script.onload = function () {

// Add HTML
var myHTML='<div class="jumbotron container-fluid"> <h2>HR Process map</h2> <div class="row separator"></div> <div class="row"> <div class="col-md-2 box2 box-height ">Core Processes</div> <div class="col-md-10 core_here"></div> </div> <div class="row separator"></div> <div class="row "> <div class="col-md-2 box2 box-height ">Supporting Processes</div> <div class="col-md-10 supporting_p_here "></div> </div> <div class="row separator"></div> <div class="row"> <div class="col-md-2 box2 box-height ">Supporting Functions</div> <div class="col-md-10 supporting_f_here"></div> </div> </div>'
$('.entry').prepend(myHTML)

color_main='red'
color_sub='blue'
core=[],
i=0, 
supsf=[],
sf=[],
sp=[],
num_supportingp=0,
pf_index=[],
links=[],
inner_array=[],
inner_array2=[],
id=0,
one_toggled=0,
clean_flag=0,
links2=[],
dict={}

dataSource =alasql('SELECT * FROM XLSX('+'\"'+excelLocation+'\"'+',{headers:false,sheetid:"Relations"})',[],

function(data){

//Task 1: Acquire names of Core Processes
//data holds the spreadsheet. This pushes all the rows from the A column into a list called core
for( i=1;i<data.length;i++){core.push(data[i]['A'])}
//Task 2: Acquire names of Supporting Processes and Functions
//This takes elements from the first row (where names are), and pushes them into supsf. Then, the first element is discarded, as there is no name there
for (i in data[0]){supsf.push(data[0][i])}
supsf.splice(0,1);
//Task 3: Divide supsf into processes and functions
//To keep track of both for future use
//Separate the names using the _ and push them into sp and sf. Count the number of supporting processes
//Pf index holds whether the item in question is a supporting process or function
for(i in supsf){
	s=supsf[i].split('_')
	if(s[0]=='SP'){
		sp.push(s[1])
		pf_index.push('P')
		num_supportingp+=1
	}
	else if(s[0]=='SF'){
		sf.push(s[1])
		pf_index.push('F')
	}
}
//Create the array for viz
//The ith element of links will contain the position in the pf_index of the i+1th core process.
//Core process 1 is in links[0]. The values within are the n-th childs of supporting_p_here that needs to have its color changed.
//But for supporting functions, they are actually shifted by the number of supporting processes. So if something is a supporting function, it will occupy the n-n_supporting_p position
// This creates the links array, which has subarrays. The first array of links has the elements that have to be coloured among SP and SF, in the order set in excel
$.each(data,function(index,value){
	//Don't read the first index, that contains the names of the supporting processes and functions
	if(index!=0){
		//Reset the variables so that we don't carry over between loops
		i=0;inner_array=[];inner_array2=[]
		$.each(value,function(index2,value2){
			//Don't push the first column, that has the name of the processes. Push only if there is something in there (if value2)
			if(index2!='A'&&value2){inner_array.push(i)}
			//And increase the position counter		
			i+=1	
		})
		links.push(inner_array)
	}
})
// Create the visualisation

//Use the div elements to append them a number of div blue boxes with the text dictated by the arrays core,sp, and sf

for(i in core){
	$( ".core_here" ).append( "<div class='col-md-2 box box-height core'>"+core[i]+"</div>" )}
for(i in sp){$( ".supporting_p_here" ).append( "<div class='col-md-2 box box-height'>"+sp[i]+"</div>" )}
for(i in sf){$( ".supporting_f_here" ).append( "<div class='col-md-2 box box-height'>"+sf[i]+"</div>" )}
//Add effects
$('.core').mouseover(
//When should it trigger illumination? 
// If one is toggled, do not illuminate, unless that is the toggled element
// If none is toggled, illuminate
function(){
	if(one_toggled==0||$(this).hasClass("glow") && one_toggled==1){
		id=$(this).index()
		//Add colour to the item over which we are hovering the mouse
		$(this).css('background',color_main)
		//First, get the array that holds which elements must be iluminated. That's links[id]

		$.each(links[id],function(index,value){
			//Is item 'index' a process or function? Colour them accordingly
			if(pf_index[value-1]=='P'){
				$('.supporting_p_here .box:nth-child('+value+')').css('background',color_sub)}
				else if (pf_index[value-1]=='F') {			
					$('.supporting_f_here .box:nth-child('+(value-num_supportingp)+')').css('background',color_sub)}
				})

	}
	//If it is already toggled, or other is, just colour that button
	else{
		$(this).css('background',color_main)
		
	}
})
//On mouse leave
// If that one is toggled, do not disable on leave
//Otherwise, disable

$('.core').mouseout(
function(){
	if($(this).hasClass("glow")==false&&one_toggled==0||clean_flag==1){
		id=$(this).index()
		$(this).css('background','')
		
			$.each(links[id],function(index,value){
			if(pf_index[value-1]=='P'){
				$('.supporting_p_here .box:nth-child('+value+')').css('background','')}
			else if (pf_index[value-1]=='F') {
				$('.supporting_f_here .box:nth-child('+(value-num_supportingp)+')').css('background','')}
			})	
		
	}
//If not, just colour it
	else{
		$(this).css('background','')
	}
})
//This ensures that colours change when clicks happen
$('.core').click(
	function(){			
		//If the clicked element was selected, de-select it, and sat a global variable saying that no element is now selected
			if($(this).hasClass("glow")){
				$(this).removeClass("glow")
				one_toggled=0	
				$(this).find(".customlink").on('click', DoPrevent);	
								
			}
			else{
				// If it was not selected, first, enter in 'special mode' by setting clean_flag, then make it glow, and remove other boxes's
				//  selected Functions. Then, trigger a mousover, so that we don't have to move the mouse out and in, and we get the highlighting directly
				clean_flag=1
				$(this).siblings().trigger('mouseout');	
				$(this).siblings().find(".customlink").off('click', DoPrevent);
				$(this).siblings().find(".customlink").on('click', DoPrevent);	

				$(this).find(".customlink").off('click', DoPrevent);	
				$(this).addClass("glow")				
				$(this).trigger('mouseover');
				$(this).siblings().removeClass("glow")	
				clean_flag=0			
				one_toggled=1;	


					
			}			
		})
//End of hovercode

//Now, add the links. We read this separately because apparently you can't read it
})
dataSource =alasql('SELECT * FROM XLSX('+'\"'+excelLocation+'\"'+',{headers:false,sheetid:"Links"})',[],
//Wrap with <a href="LINK" class=customlink> and </a>
	function(data){		
		//Strategy: Build a dictionary that contains the names from the first column
		// Then, create an array with the references to all elements we want to add links to
		//Walk through the elements, adding the links, matching using their existing value
		data.forEach(function(value,i){
			if(value['A']!=undefined){
		    label=value['A'].split('_')
			if(label[0]=='SP' || label[0]=='SF'){label=label[1]}
			else{label=label[0]}				
			dict[label]=value['B']}
		})	
			
		l=$.merge($( ".core_here" ).children(),$( ".supporting_p_here" ).children())
		$.merge(l,$( ".supporting_f_here" ).children())		
		l.each(function(i,value){			
			value.innerHTML=linkize(dict[value.innerText],value.innerText)
		})
		//$(this).siblings().trigger('mouseout');	
		$('.core .customlink').on('click', DoPrevent);
$('.box-height').matchHeight();		
	})

// Then, deactivate the links for the core processes, and trigger a click of the div when the href

function DoPrevent(e) {
  e.preventDefault();
  e.stopPropagation();
  $(this).parent().trigger("click")
  
}
//Helper function to add links to items. It accepts both www.google.com and http://www.google.com and google.com types of links. Just in case.
function linkize(link,text){
			if(link!=undefined){
			link=link.split('http://')
			if(link[0]==""){link.splice(0,1)}	
			link=link[0]
			link=link.split('www.')
			if(link[0]==""){link.splice(0,1)}		
			linked_text="<a href='http://"+link+"\'' "+"class='customlink'>"+text+"</a>"}
			else{
				linked_text=text
			}	
			return linked_text
}



};
document.head.appendChild(script);