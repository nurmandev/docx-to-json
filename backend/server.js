const express = require("express");
const connectDB = require("./config/db");
const cors = require("cors");
const reportRoutes = require("./routes/reportRoutes");
const dotenv = require("dotenv");

dotenv.config();
connectDB();

const app = express();
app.use(cors());
app.use(express.json());
app.use("/api/reports", reportRoutes);

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
