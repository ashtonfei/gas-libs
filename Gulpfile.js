const { src, dest, series } = require('gulp')
const expose = require('gulp-expose')
const rename = require('gulp-rename')
const del = require("del")

const IDENTIFIER = "GoogleFormConcatenator"

const build = function () {
    return src("src/*.js")
        .pipe(expose("this", IDENTIFIER))
        .pipe(rename({ extname: '.gs' }))
        .pipe(dest("dist"))
}

const clean = function () {
    return del(['dist/*'])
}

exports.default = series(clean, build)