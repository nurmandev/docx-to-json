const express = require("express");
const { uploadReport } = require("../controllers/reportController");
const multer = require("multer");

const router = express.Router();
const upload = multer({ dest: "uploads/" });

router.post("/upload", upload.single("file"), uploadReport);

module.exports = router;
