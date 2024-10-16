import axios from "axios";

const API_URL = "http://127.0.0.1:5000/api";

export const uploadImage = async (file) => {
  const formData = new FormData();
  formData.append("file", file);

  const response = await axios.post(`${API_URL}/upload-image`, formData);
  return response.data;
  // return { url: "image" + Date.now().toString() };
};

export const handleStoreJson = async (data) => {
  try {
    const res = await fetch("http://localhost:5000/api/store-json", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data),
    });

    const result = await res.json();
    return result;
  } catch (error) {
    console.error("Error storing JSON:", error);
  }
};
