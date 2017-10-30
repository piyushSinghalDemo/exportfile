/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
// var Worker = require('tiny-worker');
var Excel = require('exceljs');
var fs = require('fs');
var Hapi = require('hapi');
const Boom = require('boom');
const server = new Hapi.Server();
	var port = 3000;
	server.connection({
		host: 'localhost',
		port: port
	});
const dbOpts = {
    url: 'mongodb://localhost:27017/exportToExcel',
    settings: {
        poolSize: 10
    },
    decorate: true
};
/* server.pack.require('hapi-mongodb', dbOpts, function (err) {
	if (err) {
		console.err(err);
		throw err;
	}

}); */
server.register({
	register: require('hapi-mongodb'),
	options: dbOpts
	}, function (err) {
	if (err) {
	console.error(err);
	throw err;
	}
	server.start(function () {
		console.log(`Server started at ${server.info.uri}`);
	});	
});


/**
 * For Hello World Api
 */
server.route({
	method: 'GET',
	path: '/',
	handler: function (request, reply) {
		helloWorld();
		// const db = request.mongo.db;
		/* db.collection('personal').find().toArray(function (err, doc){
			reply(doc);
			}); */
	}
});
/**
 * Single Tab
 *  
 */
server.route({
	method: 'GET',
	path: '/sheet',
	handler: function (request, reply) {

		singleTab(request);
	}
});
/**
 * Multi Tab Export
 */
server.route({
	method: 'GET',
	path: '/multi-sheet-excel',
	handler: function (req, res) {

		MultipleSheet(res);
	}
});
/**
 * Formatted Data (Huper Link, Date Format)
 */
server.route({
	method: 'GET',
	path: '/formatted-sheet',
	handler: function (req, res) {

		formatData(res);
	}
});

server.route({
	method: 'POST',
	path: '/',
	config: {
		payload: {
			output: 'file',
			parse: true,
			allow: 'multipart/form-data'
		}
	},
	handler: function (req, res) {
		// logit(req.raw.req, req.raw.res);
		if (req.query.f) return post_file(req, res, req.query.f);
		return post_data(req, res);
	}
});
server.route({
	method: 'POST',
	path: '/file',
	handler: function (req, res) {
		// logit(req.raw.req, req.raw.res);
		if (req.query.f) return post_file(req, res, req.query.f);
		return post_data(req, res);
	}
});
/* server.start(function (err) {
	if (err) throw err;
	console.log('Serving HTTP on port ' + port);
}); */

/**
 * Basic Hello World example
 */
function helloWorld() {

	// create workbook & add worksheet 
	var workbook = new Excel.Workbook();
	var worksheet = workbook.addWorksheet('Discography');
	worksheet.getCell('A1').value = "Hello World";
	// save workbook to disk 
	workbook.xlsx.writeFile('hello_world.xlsx').then(function () {
		console.log("saved");
	});

}

/**
 * Basic Hello World example
 */
function singleTab(request) {
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook();
	var worksheet = workbook.addWorksheet('Discography');
	// add column headers 
	worksheet.columns = [{
			header: 'Name',
			key: 'Name'
		},
		{
			header: 'Mobile',
			key: 'Mobile'
		},
		{
			header: 'City',
			key: 'city'
		}
	];
	const db = request.mongo.db;
	db.collection('personal').find().toArray(function (err, doc){
		worksheet.addRows(doc);
		console.log("Doc length"+doc.length);
		workbook.xlsx.writeFile('taylor_swift.xlsx').then(function () {
			console.log("saved");
		});
		});
	// add row using keys 
	/* worksheet.addRow({
		album: "Taylor Swift",
		year: 2006
	}); */

	// add rows the dumb way 
	// worksheet.addRow(["Fearless", 2008]);

	// add an array of rows 
	/* var rows = [
		["Speak Now", 2010],
		{
			album: "Red",
			year: 2012
		}
	]; */
	// save workbook to disk 
	

}
/**
 * Basic Multiple Sheet example
 */
function MultipleSheet(res) {
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook();
	var worksheet = workbook.addWorksheet('Discography');
	var secondSheet = workbook.addWorksheet('Dummy');
	// add column headers 
	worksheet.columns = [{
			header: 'Album',
			key: 'album'
		},
		{
			header: 'Year',
			key: 'year'
		}
	];
	secondSheet.columns = [{
			header: 'Demo1',
			key: 'demo1'
		},
		{
			header: 'Demo2',
			key: 'demo2'
		}
	];

	// add row using keys 
	worksheet.addRow({
		demo1: "Taylor Swift",
		demo2: 2006
	});
	secondSheet.addRow({
		demo1: "Taylor Swift",
		demo2: 2006
	});
	// add rows the dumb way 
	worksheet.addRow(["Fearless", 2008]);
	secondSheet.addRow(["Fearless", 2008]);
	// add an array of rows 
	var rows = [
		["Speak Now", 2010],
		{
			album: "Red",
			year: 2012
		}
	];
	var rows2 = [
		["Speak Now", 2010],
		{
			demo1: "Red",
			demo2: 2012
		}
	];
	worksheet.addRows(rows);
	secondSheet.addRows(rows2);
	// edit cells directly 
	worksheet.getCell('A6').value = "1989";
	secondSheet.getCell('B6').value = 2014;

	// save workbook to disk 
	workbook.xlsx.writeFile('multitab.xlsx').then(function () {
		console.log("saved");
		//   res.download('taylor_swift.xlsx');
	});
}

/**
 * Add Image
 */
function addImage() {
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook();
	var worksheet = workbook.addWorksheet('Discography');

	// add column headers 
	worksheet.columns = [{
			header: 'Album',
			key: 'album'
		},
		{
			header: 'Year',
			key: 'year'
		}
	];

	// add row using keys 
	worksheet.addRow({
		album: "Taylor Swift",
		year: 2006
	});

	// add rows the dumb way 
	worksheet.addRow(["Fearless", 2008]);

	// add an array of rows 
	var rows = [
		["Speak Now", 2010],
		{
			album: "Red",
			year: 2012
		}
	];
	worksheet.addRows(rows);
	// edit cells directly 
	worksheet.getCell('A6').value = "1989";
	worksheet.getCell('B6').value = 2014;
	// save workbook to disk 
	workbook.xlsx.writeFile('taylor_swift.xlsx').then(function () {
		console.log("saved");
		// res.download('test.xlsx'); 
	});
}
/**
 * Format Data 
 * 
 */
function formatData() {
	// create workbook & add worksheet 
	var workbook = new Excel.Workbook();
	var worksheet = workbook.addWorksheet('Discography');
	// add column headers 
	worksheet.columns = [{
			header: 'Album',
			key: 'album'
		},
		{
			header: 'Year',
			key: 'year'
		},
		{
			header: 'url',
			key: 'url',
			width: 32
		},
		{
			header: "D.O.B.",
			key: "DOB",
			width: 10
		}
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
	worksheet.addRow({
		album: "Taylor Swift",
		year: 2006,
		url: "www.google.com",
		DOB: new Date(1970, 1, 1)
	});

	// add rows the dumb way 
	worksheet.addRow(["Fearless", 2008, "www.google.com", new Date(1970, 1, 1)]);

	// add an array of rows 
	var rows = [
		["Speak Now", 2010, "www.google.com", new Date(1970, 1, 1)],
		{
			album: "Red",
			year: 2012,
			url: "www.google.com",
			DOB: new Date()
		}
	];
	worksheet.addRows(rows);
	worksheet.getCell('C2').value = {
		text: 'www.mylink.com',
		hyperlink: 'http://www.mylink.com'
	};
	worksheet.getCell('C3').value = {
		text: 'www.mylink.com',
		hyperlink: 'http://www.mylink.com'
	};
	worksheet.getCell('C3').value = {
		text: 'www.google.com',
		hyperlink: 'http://www.google.com'
	};
	workbook.xlsx.writeFile('formatted.xlsx').then(function () {
		console.log("saved");
		// res.download('test.xlsx'); 
	});
}
