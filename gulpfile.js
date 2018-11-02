var gulp        = require('gulp');
var browserSync = require('browser-sync').create();


gulp.task('default', function() { 
       var files = [   
            '**/*.asp',   
             '**/*.css',   
              '**/*.js'    
            ];
        browserSync.init(files,{
            proxy: "http://192.168.2.242:82" 
             });
            });
