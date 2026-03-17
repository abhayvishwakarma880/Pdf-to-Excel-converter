// const multer = require("multer");
import multer from "multer";

export const upload = multer({ dest: "uploads/" });

// app.post("/upload", upload.single("pdf"), async (req, res) => {
//   const filePath = req.file.path;
//   // next step
// });