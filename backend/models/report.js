const mongoose = require("mongoose");

const ReportSchema = new mongoose.Schema(
  {
    content: [mongoose.Schema.Types.Mixed],
  },
  { timestamps: true }
);

// Model based on the schema
const Report = mongoose.model("Report", ReportSchema);

module.exports = Report;
