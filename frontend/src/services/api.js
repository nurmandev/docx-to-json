import axios from "axios";

const API_URL = "https://docx-to-json.onrender.com/api/reports";
// const API_URL = "http://localhost:5000/api/reports";

export const uploadDocx = async (file) => {
  const formData = new FormData();
  formData.append("file", file);

  const response = await axios.post(`${API_URL}/upload`, formData);
  return response.data;
};
