const express = require("express")
const multer = require("multer")
const fs = require("fs")
const mammoth = require("mammoth")
const stream = require("stream")
const HtmlToDocx = require("html-to-docx")

require("dotenv").config()

const app = express()
const upload = multer({dest: `${__dirname}/uploads`})

app.post("/word-to-html", upload.single("file"), function (req, res) {
    mammoth.convertToHtml({path: req.file.path})
        .then(result => {
            res.send(result.value)
        })
        .catch(() => {
            res.status(400)
            res.send("Error: failed to process file.")
        })
})

app.post("/html-to-word", upload.single("file"), async function (req, res) {
    let converted = await HtmlToDocx(
        fs.readFileSync(req.file.path).toString()
    )

    let readStream = new stream.PassThrough()
    readStream.end(converted)

    res.attachment(`${req.file.originalname}.docx`)
    readStream.pipe(res)
})

app.listen(process.env.PORT, process.env.HOSTNAME, () => {
    console.log(`Server started at port ${process.env.HOSTNAME}:${process.env.PORT}.`)
})