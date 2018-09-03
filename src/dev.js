const path = require('path');
const fs = require('fs');

const CommandLine = require('./deploy/commandline').CommandLine;
const Exceltk = require('./deploy/exceltk').Exceltk;

/**
 * main entry
 * @returns
 */
function main() {
    let commandLine = new CommandLine({
    	'h':false,       // help
    	'pub': false,    // pub command
    	'platform':null  // platform: macos/windows
    });

    let config = commandLine.parse();
    if(!config.pub||config.h){
    	console.log('----------------------');
    	console.log('exceltk dev tool 0.0.1');
    	console.log('----------------------');
    	console.log('useage:');
    	console.log('------');
    	console.log('node dev.js pub -platform macos');
    	console.log('node dev.js pub -platform windows');
    	return;
    }


    let exceltk = new Exceltk(config);
    exceltk.run();
}

main();


