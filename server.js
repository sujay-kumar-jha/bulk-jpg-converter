const express = require("express");
const cors = require("cors");
const axios = require("axios");
const sharp = require("sharp");
const fs = require("fs-extra");
const multer = require("multer");
const ExcelJS = require("exceljs");
const { v4: uuidv4 } = require("uuid");
const path = require("path");

const app = express();

app.use(cors());
app.use(express.json());

app.use(express.static("public"));
app.use("/converted", express.static("converted"));
app.use("/reports", express.static("reports"));

fs.ensureDirSync("temp");
fs.ensureDirSync("converted");
fs.ensureDirSync("reports");

const upload = multer({
    dest: "temp/"
});

app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.post("/upload", upload.single("file"), async (req, res) => {

    try {

        const filePath = req.file.path;

        const fileContent = fs.readFileSync(filePath, "utf-8");

        const links = fileContent
            .split(/\r?\n/)
            .map(link => link.trim())
            .filter(link => link !== "");

        let results = [];

        for (let link of links) {

            try {

                const id = uuidv4();

                const tempImage = `temp/${id}`;

                const outputImage = `converted/${id}.jpg`;

                const response = await axios({
                    url: link,
                    responseType: "arraybuffer",
                    timeout: 15000
                });

                fs.writeFileSync(tempImage, response.data);

                await sharp(tempImage)
                    .jpeg({ quality: 90 })
                    .toFile(outputImage);

                fs.removeSync(tempImage);

                const newLink =
                    `${req.protocol}://${req.get("host")}/${outputImage}`;

                results.push({
                    oldLink: link,
                    newLink: newLink
                });

            } catch (err) {

                results.push({
                    oldLink: link,
                    newLink: "Conversion Failed"
                });
            }
        }

        // CREATE EXCEL FILE

        const workbook = new ExcelJS.Workbook();

        const worksheet = workbook.addWorksheet("Converted Links");

        worksheet.columns = [
            {
                header: "Old Link",
                key: "oldLink",
                width: 70
            },
            {
                header: "New JPG Link",
                key: "newLink",
                width: 70
            }
        ];

        results.forEach(item => {
            worksheet.addRow(item);
        });

        const reportName = `report-${Date.now()}.xlsx`;

        const reportPath = `reports/${reportName}`;

        await workbook.xlsx.writeFile(reportPath);

        fs.removeSync(filePath);

        res.json({
            success: true,
            download:
                `${req.protocol}://${req.get("host")}/${reportPath}`,
            results
        });

    } catch (error) {

        res.status(500).json({
            success: false,
            message: "Server Error"
        });
    }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});