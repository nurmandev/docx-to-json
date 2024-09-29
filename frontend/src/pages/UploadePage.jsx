import React, { useState } from "react";
import { extractDocxContent } from "../utils";
import { handleStoreJson } from "../services/api";

const UploadPage = () => {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleUpload = async () => {
    if (file) {
      try {
        setLoading(true);
        const content = await extractDocxContent(file);
        console.log("Extracted content:", content);
        const res = await handleStoreJson({ content });
        console.log(res);
      } catch (error) {
        console.error("Failed to extract content:", error);
      } finally {
        setLoading(false);
      }
    }
  };

  return (
    <div>
      <input type="file" onChange={(e) => setFile(e.target.files[0])} />
      {loading ? (
        <span>Extracting...</span>
      ) : (
        <button onClick={handleUpload}>Upload</button>
      )}
    </div>
  );
};

export default UploadPage;
