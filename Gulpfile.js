const { src, dest, series } = require('gulp')
const expose = require('gulp-expose')
const rename = require('gulp-rename')
const jsdoc = require('gulp-jsdoc3')
const del = require("del")

const IDENTIFIER = "DocPro"

const build = function () {
    return src("src/*.js")
        .pipe(expose("this", IDENTIFIER))
        .pipe(rename({ extname: '.gs' }))
        .pipe(dest("dist"))
}

const clean = function () {
    return del(['dist/*'])
}

const doc = function () {
    // const config = require('./jsdoc.json')
    return src(["README.md", "./src/**/*.js"], { read: false })
        .pipe(jsdoc())
}

exports.build = series(clean, build)
exports.doc = doc