const { src, dest } = require('gulp')
const expose = require('gulp-expose')
const rename = require('gulp-rename')

const IDENTIFIER = "LIB"

exports.build = function () {
    return src("src/*.js")
        .pipe(expose("this", IDENTIFIER))
        .pipe(rename({ extname: '.gs' }))
        .pipe(dest("dist"))
}