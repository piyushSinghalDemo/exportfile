/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
// var Worker = require('tiny-worker');
var Excel = require('exceljs');   
var fs = require('fs');
var Hapi = require('hapi'), 
server = new Hapi.Server();
// var mongoose = require('mongoose');
// mongoose.connect('mongodb://127.0.0.1:27017/exportToExcel')
// var logit = require('./_logit');
/* var details = mongoose.model('details' ,{
	userName : String,
	mobile : String,
	link : String,
});
var personaldatas = mongoose.model('personaldatas',{
	Name : String,
	mobile : String,
	city : String,
});
var portfolio = mongoose.model('portfolio',{
	Name : String,
	url : String,
	Disclaimer : String
}); */

var port = 7262;
server.connection({ host:'localhost', port: port}); 

/**
 * For Hello World Api
 */
server.route({ method: 'GET', path: '/',
handler: function(req, res) {

	helloWorld();
}});
/**
 * Single Tab
 *  
 */
server.route({ method: 'GET', path: '/sheet',
handler: function(req, res) {

	singleTab();
}});
/**
 * Multi Tab Export
 */
server.route({ method: 'GET', path: '/multi-sheet-excel',
handler: function(req, res) {

	MultipleSheet(res);   
}});
/**
 * Formatted Data (Huper Link, Date Format)
 */
server.route({ method: 'GET', path: '/formatted-sheet',
handler: function(req, res) {

	formatData(res);
}});

server.route({ method: 'POST', path: '/',
config:{payload:{ output: 'file', parse: true, allow: 'multipart/form-data'}},
handler: function(req, res) {
	// logit(req.raw.req, req.raw.res);
	if(req.query.f) return post_file(req, res, req.query.f);
	return post_data(req, res);
}});
server.route({ method: 'POST', path: '/file',
handler: function(req, res) {
	// logit(req.raw.req, req.raw.res);
	if(req.query.f) return post_file(req, res, req.query.f);
	return post_data(req, res);
}});
server.start(function(err) {
	if(err) throw err;
	console.log('Serving HTTP on port ' + port);
});

/**
 * Basic Hello World example
 */
function helloWorld(){
	
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook(); 
	var worksheet = workbook.addWorksheet('Discography'); 
	worksheet.getCell('A1').value = "Hello World";	 
	// save workbook to disk 
	 workbook.xlsx.writeFile('hello_world.xlsx').then(function() { 
	  console.log("saved");
	});
	
}

/**
 * Basic Hello World example
 */
function singleTab(){
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook(); 
	var worksheet = workbook.addWorksheet('Discography'); 
	// add column headers 
	worksheet.columns = [ 
		{ header: 'Name', key: 'userName'}, 
		{ header: 'Mobile', key: 'mobile'},
		{ header: 'City', key: 'city'}  
	]; 
	
	// add row using keys 
	worksheet.addRow({album: "Taylor Swift", year: 2006}); 
	
	// add rows the dumb way 
	worksheet.addRow(["Fearless", 2008]);  
	
	// add an array of rows 
	 var rows = [ 
	  ["Speak Now", 2010], 
	  {album: "Red", year: 2012} 
	];  
	worksheet.addRows(rows); 
	// save workbook to disk 
	 workbook.xlsx.writeFile('taylor_swift.xlsx').then(function() { 
	  console.log("saved");
	});
	
}
/**
 * Basic Multiple Sheet example
 */
function MultipleSheet(res){	
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook(); 
	var worksheet = workbook.addWorksheet('Discography'); 
	var secondSheet = workbook.addWorksheet('Dummy'); 
	// add column headers 
	worksheet.columns = [ 
		{ header: 'Album', key: 'album'}, 
		{ header: 'Year', key: 'year'} 
	];
	secondSheet.columns = [ 
		{ header: 'Demo1', key: 'demo1'}, 
		{ header: 'Demo2', key: 'demo2'} 
	];
	
	// add row using keys 
	worksheet.addRow({demo1: "Taylor Swift", demo2: 2006}); 
	secondSheet.addRow({demo1: "Taylor Swift", demo2: 2006}); 
	// add rows the dumb way 
	worksheet.addRow(["Fearless", 2008]);  
	secondSheet.addRow(["Fearless", 2008]);  
	// add an array of rows 
	var rows = [ 
	  ["Speak Now", 2010], 
	  {album: "Red", year: 2012} 
	];
	var rows2 = [ 
		["Speak Now", 2010], 
		{demo1: "Red", demo2: 2012} 
	  ]; 
	worksheet.addRows(rows); 
	secondSheet.addRows(rows2);
	// edit cells directly 
	worksheet.getCell('A6').value = "1989"; 
	secondSheet.getCell('B6').value = 2014; 
	
	// save workbook to disk 
	workbook.xlsx.writeFile('multitab.xlsx').then(function() { 
	  console.log("saved");
	//   res.download('taylor_swift.xlsx');
	}); 
}

/**
 * Add Image
 */
function addImage(){
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook(); 
	var worksheet = workbook.addWorksheet('Discography'); 
		
	// add column headers 
	worksheet.columns = [ 
		{ header: 'Album', key: 'album'}, 
		{ header: 'Year', key: 'year'} 
	]; 
	
	// add row using keys 
	worksheet.addRow({album: "Taylor Swift", year: 2006}); 
	
	// add rows the dumb way 
	worksheet.addRow(["Fearless", 2008]);  
	
	// add an array of rows 
	var rows = [ 
		["Speak Now", 2010], 
		{album: "Red", year: 2012} 
	]; 
	worksheet.addRows(rows); 
	// edit cells directly 
	worksheet.getCell('A6').value = "1989"; 
	worksheet.getCell('B6').value = 2014; 
	// save workbook to disk 
	workbook.xlsx.writeFile('taylor_swift.xlsx').then(function() { 
		console.log("saved");
		// res.download('test.xlsx'); 
	});
}
/**
 * Format Data 
 * 
 */
function formatData(){
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook(); 
	var worksheet = workbook.addWorksheet('Discography'); 
	// add column headers 
	worksheet.columns = [ 
		{ header: 'Album', key: 'album'}, 
		{ header: 'Year', key: 'year'},
		{ header: 'url', key: 'url', width: 32},
		{ header: "D.O.B.", key: "DOB", width: 10 } 
	]; 
	worksheet.getCell('A1').font = {
		family: 4,
		size: 16,
		underline: true,
		bold: true
	};
	worksheet.getCell('B1').font = {
		family: 4,
		size: 16,
		underline: true,
		bold: true
	};
	worksheet.getCell('C1').font = {
		family: 4,
		size: 16,
		underline: true,
		bold: true
	};
	worksheet.getCell('D1').font = {
		family: 4,
		size: 16,
		underline: true,
		bold: true
	};
	// add row using keys 
	worksheet.addRow({album: "Taylor Swift", year: 2006, url:"www.google.com", DOB: new Date(1970,1,1)}); 
	
	// add rows the dumb way 
	worksheet.addRow(["Fearless", 2008,"www.google.com" , new Date(1970,1,1)]);  
	
	// add an array of rows 
	var rows = [ 
		["Speak Now", 2010,"www.google.com", new Date(1970,1,1)], 
		{album: "Red", year: 2012,url:"www.google.com", DOB: new Date()} 
	]; 
	worksheet.addRows(rows); 
	worksheet.getCell('C2').value = { text: 'www.mylink.com', hyperlink: 'http://www.mylink.com' };
	worksheet.getCell('C3').value = { text: 'www.mylink.com', hyperlink: 'http://www.mylink.com' };
	worksheet.getCell('C3').value = { text: 'www.google.com', hyperlink: 'http://www.google.com' };
	workbook.xlsx.writeFile('formatted.xlsx').then(function() { 
	 console.log("saved");
		// res.download('test.xlsx'); 
	});
}

