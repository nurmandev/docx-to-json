const { extractDocxContent } = require("../utils/fileProcessor");
const fs = require("fs");
const mongoose = require("mongoose");

const Report = require("../models/report");

const uploadReport = async (req, res) => {
  try {
    const filePath = req.file.path;
    const jsonData = await extractDocxContent(filePath);

    // Save the JSON data to MongoDB
    const report = new Report({ content: jsonData }); // Adjust based on your model's structure
    await report.save();

    // Remove the file after processing
    fs.unlinkSync(filePath);

    // Respond with the saved document
    res.json(report);
  } catch (error) {
    console.error("Error processing the file:", error);
    res
      .status(500)
      .json({ message: "An error occurred while processing the file." });
  }
};

module.exports = { uploadReport };
