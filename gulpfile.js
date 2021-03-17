'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');
const replace = require('gulp-replace');
const regexp = /\d{4}-\d{2}-\d{2} \d{2}:\d{2}/;

const dateTime = new Date();
const isoDateTime = dateTime.toISOString();
const isoDate = isoDateTime.substring(0,10);
const isoTime = isoDateTime.substring(11,16);
const buildTimeStamp = isoDate + ' ' + isoTime;

 
const replaceBuildTimeStamp = build.subTask ('replaceBuildTimeStamp', function () {
  return gulp.src(['src/webparts/permissionCenter/PermissionCenterWebPart.ts'])
    .pipe(replace(regexp, buildTimeStamp))
    .pipe(gulp.dest('src/webparts/permissionCenter/'))
    ;
});

build.rig.addPreBuildTask(replaceBuildTimeStamp);

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.initialize(require('gulp'));
