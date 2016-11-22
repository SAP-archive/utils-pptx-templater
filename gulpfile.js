'use strict';

let gulp = require('gulp');
let eslint = require('gulp-eslint');
let mocha = require('gulp-mocha');
let istanbul = require('gulp-istanbul');

gulp.task('eslint', () => {
    return gulp.src(['src/**/*.js', '!src/**/*.min.js'])
        .pipe(eslint())
        .pipe(eslint.format())
        .pipe(eslint.failAfterError());
});

gulp.task('pre-test', () => {
    return gulp.src([
        './src/*.js'
    ])
    .pipe(istanbul({includeUntested: true}))
    .pipe(istanbul.hookRequire());
});

gulp.task('test', () => {
    return gulp.src(['./test/*.spec.js'], {
        read: false
    })
    .pipe(mocha({
        reporter: 'spec',
        timeout: 2000
    }))
    .pipe(istanbul.writeReports())
    .pipe(istanbul.enforceThresholds({ thresholds: { global: 90 } }));
});

gulp.task('test-no-timeout', () => {
    gulp.src(['./test/*.spec.js'], {
        read: false
    })
        .pipe(mocha({
            reporter: 'spec',
            timeout: 400000
        }));
});

gulp.task('default', ['test']);

