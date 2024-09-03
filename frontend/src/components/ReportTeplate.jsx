import React from "react";

const ReportTemplate = ({ reportData }) => {
  return (
    <div>
      {reportData.headers.map((header, index) => (
        <h1 key={index} style={{ color: header.color }}>
          {header.text}
        </h1>
      ))}
      {/* Render other elements similarly */}
    </div>
  );
};

export default ReportTemplate;
