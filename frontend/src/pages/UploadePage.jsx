import React, { useState } from "react";
import { uploadDocx } from "../services/api";

const UploadPage = () => {
  const [file, setFile] = useState(null);

  const handleUpload = async () => {
    if (file) {
      const response = await uploadDocx(file);
      console.log("Uploaded:", response);
    }
  };

  return (
    <div>
      <input type="file" onChange={(e) => setFile(e.target.files[0])} />
      <button onClick={handleUpload}>Upload</button>
    </div>
  );
};

export default UploadPage;
