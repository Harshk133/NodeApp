const express = require("express");
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
var app = express();
var bodyParser = require("body-parser");
const PORT = 7000;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.use(express.static("./public"));
app.set("view engine", "ejs");

app.get("/", (req, res) => {
    res.render("page");
});

// This routing is for Certificate!
app.get("/certificate", (req, res)=>{
    res.render("page");
});

app.post("/certificate", (req, res) => {
    console.log(req.body);
    try {
        const templateFile = fs.readFileSync(path.resolve(__dirname, './public/templateDocx/JPR.docx'), 'binary');
        const zip = new PizZip(templateFile);
        let outputDocument = new Docxtemplater(zip);

        const dataToAdd_ceritficate = {
            STUDENT_NAME: req.body.name,
            STUDENT_ENR: req.body.enrollmentno,
            MICROPROJECT_SUBJECT: req.body.micro,
            COI: req.body.coi,
            SUBJECT: req.body.sub,
            TEACHER_NAME: req.body.teachername
        };
        outputDocument.setData(dataToAdd_ceritficate);

        try {
            outputDocument.render()
            let outputDocumentBuffer = outputDocument.getZip().generate({ type: 'nodebuffer' });
            fs.writeFileSync(path.resolve(__dirname, `./public/files/computer/certificate/${req.body.name}-computer-certificate.docx`), outputDocumentBuffer);
        }
        catch (error) {
            console.error(`ERROR Filling out Template:`);
            console.error(error)
        }
    } catch (error) {
        console.error(`ERROR Loading Template:`);
        console.error(error);
    }
    res.download(`./public/files/computer/certificate/${req.body.name}-computer-certificate.docx`);
});


app.get("/microproject", (req, res) => {
    res.render("page");
});

app.post("/microproject", (req, res) => {
    console.log(req.body);
    try {
        const templateFile = fs.readFileSync(path.resolve(__dirname, './public/templateDocx/EST.docx'), 'binary');
        const zip = new PizZip(templateFile);
        let outputDocument = new Docxtemplater(zip);

        const dataToAdd = {
            STUDENT_NAME: req.body.Mname,
            STUDENT_ENR: req.body.Menrollmentno,
            ROLLNO: req.body.Mrollno,
            TEACHER_NAME: req.body.Mteacher
        };
        outputDocument.setData(dataToAdd);

        try {
            outputDocument.render()
            let outputDocumentBuffer = outputDocument.getZip().generate({ type: 'nodebuffer' });
            fs.writeFileSync(path.resolve(__dirname, `./public/files/computer/microproject/${req.body.Mname}-computer-microproject.docx`), outputDocumentBuffer);
        }
        catch (error) {
            console.error(`ERROR Filling out Template:`);
            console.error(error)
        }
    } catch (error) {
        console.error(`ERROR Loading Template:`);
        console.error(error);
    }
    res.download(`./public/files/computer/microproject/${req.body.Mname}-computer-microproject.docx`);
});

app.get("/another", (req, res) => {
    res.render("page");
});

app.post("/another", (req, res) => {
    console.log(req.body);
        try {
            const templateFile = fs.readFileSync(path.resolve(__dirname, './public/templateDocx/DTE.docx'), 'binary');
            const zip = new PizZip(templateFile);
            let outputDocument = new Docxtemplater(zip);
    
            const dataToAdd = {
                STUDENT_NAME: req.body.Mnameo,
                STUDENT_ENR: req.body.Mrollnoo,
                ROLLNO: req.body.Mteachero,
                TEACHER_NAME: req.body.Mteachero
            };
            outputDocument.setData(dataToAdd);
    
            try {
                outputDocument.render()
                let outputDocumentBuffer = outputDocument.getZip().generate({ type: 'nodebuffer' });
                fs.writeFileSync(path.resolve(__dirname, `./public/files/computer/microproject/${req.body.Mnameo}-computer-microproject.docx`), outputDocumentBuffer);
            }
            catch (error) {
                console.error(`ERROR Filling out Template:`);
                console.error(error)
            }
        } catch (error) {
            console.error(`ERROR Loading Template:`);
            console.error(error);
        }
        res.download(`./public/files/computer/microproject/${req.body.Mnameo}-computer-microproject.docx`);
    }    
);

// app.post("/mcss", (req, res) => {
//     console.log(req.body);
//     try {
//         const templateFile = fs.readFileSync(path.resolve(__dirname, './public/templateDocx/CSS MICROPROJECT.docx'), 'binary');
//         const zip = new PizZip(templateFile);
//         let outputDocument = new Docxtemplater(zip);

//         const dataToAdd = {
//             STUDENT_NAME_CSS: req.body.mcssname,
//             STUDENT_ENR_CSS: req.body.mcssenrollment,
//             ROLLNO_CSS: req.body.mcssenrollno,
//             TEACHER_NAME_CSS: req.body.mcssteacher
//         };
//         outputDocument.setData(dataToAdd);

//         try {
//             outputDocument.render()
//             let outputDocumentBuffer = outputDocument.getZip().generate({ type: 'nodebuffer' });
//             fs.writeFileSync(path.resolve(__dirname, `./public/files/computer/microproject/${req.body.mcssname}-computer-microproject.docx`), outputDocumentBuffer);
//         }
//         catch (error) {
//             console.error(`ERROR Filling out Template:`);
//             console.error(error)
//         }
//     } catch (error) {
//         console.error(`ERROR Loading Template:`);
//         console.error(error);
//     }
//     res.download(`./public/files/computer/microproject/${req.body.mcssname}-computer-microproject.docx`);
// });



app.listen(PORT, () => {
    console.log(`Server is running on ${PORT} port!`)
});
