import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import multer from 'multer';
import ExcelJS from 'exceljs';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
// Direct path to avoid buggy index.js in pdf-parse
const pdf = require('pdf-parse/lib/pdf-parse.js');

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);


// Ensure directories exist
if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');
if (!fs.existsSync('outputs')) fs.mkdirSync('outputs');



const app = express();
const PORT = process.env.PORT || 3000;

// Set EJS as the view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Serve static files (for CSS, JS, Images)
app.use(express.static(path.join(__dirname, 'public')));

// Root Route
app.get('/', (req, res) => {
    res.render('index', {
        title: 'PDF to Excel Converter',
        message: 'Welcome to the premium converter!',
        error: null
    });
});


const upload = multer({ dest: "uploads/" });
app.post("/upload", upload.single("pdf"), async (req, res) => {
    try {
        const filePath = req.file.path;

        // 📄 PDF read
        const dataBuffer = fs.readFileSync(filePath);
        const data = await pdf(dataBuffer);

        // ✅ Delete uploaded PDF after processing
        fs.unlinkSync(filePath);

        const text = data.text;
        
        // 🎯 Extract data (example)
        const contactNumber = text.match(/Contract\s*No\.?\s*:\s*([A-Z0-9-]+)/i)?.[1];
        const contractGeneratedDate = text.match(/Contract\s*Generated\s*Date\s*:\s*([0-9A-Za-z-]+)/i)?.[1];
        const bidRaPbpNo = text.match(/Bid\s*\/\s*RA\s*\/\s*PBP\s*No\.?\s*:\s*([A-Z0-9\/.]+)/i)?.[1];
        const duration = text.match(/Duration[\s\S]*?(\d{1,3})/)?.[1];
        const amountOfContract = text.match(/Total\s*Contract\s*Value[\s\S]*?(\d+)/i)?.[1];

        // ✅ Validation: Agar zaruri fields nahi mile toh error throw karein
        if (!contactNumber && !bidRaPbpNo) {
            throw new Error("Invalid PDF: Required keys (Contract No or Bid No) not found in the uploaded file.");
        }

        const extractedData = {
            contactNumber: contactNumber || "Not Found",
            contractGeneratedDate: contractGeneratedDate || "Not Found",
            bidRaPbpNo: bidRaPbpNo || "Not Found",
            duration: duration || "Not Found",
            amountOfContract: amountOfContract || "Not Found",
        };

        // 📊 Excel create
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Data");

        worksheet.columns = [
            { header: "Contact Number", key: "contactNumber" },
            { header: "Contract Generated Date", key: "contractGeneratedDate" },
            { header: "Bid / RA / PBP No", key: "bidRaPbpNo" },
            { header: "Duration", key: "duration" },
            { header: "Amount of Contract (Including All Duties and Taxes INR)", key: "amountOfContract" },
        ];

        worksheet.addRow(extractedData);

        const outputPath = `outputs/output-${Date.now()}.xlsx`;

        await workbook.xlsx.writeFile(outputPath);

        // 📥 send file and cleanup
        res.download(outputPath, (err) => {
            if (err) {
                console.error("Download Error:", err);
            }
            // ✅ Delete Excel file after download completes
            if (fs.existsSync(outputPath)) {
                fs.unlinkSync(outputPath);
            }
        });

    } catch (error) {
        console.error("Processing Error:", error.message);
        
        // cleanup on error if file still exists
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }

        // Render index with error message instead of simple res.send
        res.render('index', {
            title: 'PDF to Excel Converter',
            message: 'Welcome to the premium converter!',
            error: error.message
        });
    }


});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

