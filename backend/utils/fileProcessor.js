const cheerio = require("cheerio");
const path = require("path");
const unzipper = require("unzipper");
const fs = require("fs");
const rimraf = require("rimraf");
const cloudinary = require("cloudinary").v2;

const dotenv = require("dotenv");

dotenv.config();

// Configure Cloudinary
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

async function uploadToCloudinary(imagePath) {
  return cloudinary.uploader.upload(imagePath, {
    folder: "docx", // Optional: specify a folder in Cloudinary
    use_filename: true, // Use the original filename for the uploaded file
    unique_filename: true, // Prevents Cloudinary from adding random characters to the filename
  });
}

async function extractDocxContent(filePath) {
  try {
    const outputDir = path.join(__dirname, "output");
    await fs.promises.mkdir(outputDir, { recursive: true });

    await fs
      .createReadStream(filePath)
      .pipe(unzipper.Extract({ path: outputDir }))
      .promise();

    const documentXmlPath = path.join(outputDir, "word", "document.xml");
    const stylesXmlPath = path.join(outputDir, "word", "styles.xml");
    const numberingXmlPath = path.join(outputDir, "word", "numbering.xml");

    const documentXml = await fs.promises.readFile(documentXmlPath, "utf8");
    const stylesXml = await fs.promises.readFile(stylesXmlPath, "utf8");
    const numberingXml = await fs.promises.readFile(numberingXmlPath, "utf8");

    const $doc = cheerio.load(documentXml, { xmlMode: true });
    const $styles = cheerio.load(stylesXml, { xmlMode: true });
    const $numbering = cheerio.load(numberingXml, { xmlMode: true });

    // Build numbering map from numbering.xml
    const numberingMap = buildNumberingMap($numbering);

    // console.log(JSON.stringify(numberingMap));

    const styleMap = {};
    $styles("w\\:style").each((_, style) => {
      const styleId = $styles(style).attr("w:styleId");
      const styleType = $styles(style).attr("w:type");

      if (styleId && styleType) {
        styleMap[styleId] = {
          type: styleType,
          name: $styles(style).find("w\\:name").attr("w:val"),
          basedOn: $styles(style).find("w\\:basedOn").attr("w:val"), // Capture the basedOn attribute
          runProperties: extractRunStyles($styles(style).find("w\\:rPr")),
          paragraphProperties: extractParagraphStyles(
            $styles(style).find("w\\:pPr")
          ),
        };
      }
    });

    // Resolve inherited styles based on the 'basedOn' attribute
    Object.keys(styleMap).forEach((styleId) => {
      const style = styleMap[styleId];
      if (style.basedOn) {
        const inheritedStyle = styleMap[style.basedOn];
        if (inheritedStyle) {
          style.runProperties = {
            ...inheritedStyle.runProperties,
            ...style.runProperties,
          };
          style.paragraphProperties = {
            ...inheritedStyle.paragraphProperties,
            ...style.paragraphProperties,
          };
        }
      }
    });

    function extractRunStyles(rPr) {
      const styles = {};

      if (rPr.find("w\\:b").length > 0) styles.bold = true;
      if (rPr.find("w\\:i").length > 0) styles.italic = true;
      if (rPr.find("w\\:u").length > 0) styles.underline = true;
      if (rPr.find("w\\:strike").length > 0) styles.strikeThrough = true;

      const color = rPr.find("w\\:color").attr("w:val");
      if (color) styles.color = color;

      const fontSize = rPr.find("w\\:sz").attr("w:val");
      if (fontSize) styles.fontSize = fontSize;

      const font = rPr.find("w\\:rFonts").attr("w:ascii");
      if (font) styles.font = font;

      const backgroundColor = rPr.find("w\\:shd").attr("w:fill");
      if (backgroundColor) styles.backgroundColor = backgroundColor;

      const highlight = rPr.find("w\\:highlight").attr("w:val");
      if (highlight) styles.highlight = highlight;

      return styles;
    }

    function extractParagraphStyles(pPr) {
      const styles = {};

      const alignment = pPr.find("w\\:jc").attr("w:val");
      if (alignment) styles.alignment = alignment;

      const spacingBefore = pPr.find("w\\:spacing").attr("w:before");
      if (spacingBefore) styles.spacingBefore = spacingBefore;

      const spacingAfter = pPr.find("w\\:spacing").attr("w:after");
      if (spacingAfter) styles.spacingAfter = spacingAfter;

      const indentLeft = pPr.find("w\\:ind").attr("w:left");
      if (indentLeft) styles.indentLeft = indentLeft;

      const indentRight = pPr.find("w\\:ind").attr("w:right");
      if (indentRight) styles.indentRight = indentRight;

      return styles;
    }

    async function parseElement(element) {
      const children = [];

      for (const child of element.children()) {
        const tag = $doc(child)[0].tagName;

        if (tag === "w:p") {
          const paragraphData = {
            type: "paragraph",
            text: "",
            styles: {},
          };

          const pPr = $doc(child).find("w\\:pPr");
          const pStyleId = pPr.find("w\\:pStyle").attr("w:val");

          if (pStyleId && styleMap[pStyleId]) {
            paragraphData.styles = {
              ...styleMap[pStyleId].paragraphProperties,
              ...styleMap[pStyleId].runProperties,
            };
          }

          if (pPr.length) {
            paragraphData.styles = {
              ...paragraphData.styles,
              ...extractParagraphStyles(pPr),
            };

            const numPr = pPr.find("w\\:numPr");
            if (numPr.length > 0) {
              const numId = numPr.find("w\\:numId").attr("w:val");
              const ilvl = numPr.find("w\\:ilvl").attr("w:val");

              const listData = extractListInfo(numId, ilvl, numberingMap);
              paragraphData.listData = listData;
            }
          }

          const nextChild = [];
          for (const run of $doc(child).find("w\\:r, w\\:drawing, w\\:pict")) {
            const runTag = $doc(run)[0].tagName;

            if (runTag === "w:r") {
              const runText = $doc(run).find("w\\:t").text();
              const rPr = $doc(run).find("w\\:rPr");
              const rStyleId = rPr.find("w\\:rStyle").attr("w:val");
              let runStyles = {};

              if (rStyleId && styleMap[rStyleId]) {
                runStyles = {
                  ...styleMap[rStyleId].runProperties,
                };
              }
              runStyles = {
                ...runStyles,
                ...extractRunStyles(rPr),
              };
              nextChild.push({
                text: runText,
                styles: runStyles,
              });
            } else if (runTag === "w:drawing") {
              const imageData = await parseDrawing(run);
              console.log(imageData);
              children.push(imageData);
            }
          }

          paragraphData.text = nextChild
            .filter((child) => child.text)
            .map((child) => child.text)
            .join("");
          paragraphData.styles = {
            ...paragraphData.styles,
            ...nextChild.styles,
          };
          paragraphData.styleName = pStyleId;
          children.push(paragraphData);
        } else if (tag === "w:tbl") {
          const tableData = {
            type: "table",
            rows: [],
          };

          for (const row of $doc(child).find("w\\:tr")) {
            const rowData = [];

            for (const cell of $doc(row).find("w\\:tc")) {
              const cellData = {
                type: "cell",
                content: await parseElement($doc(cell)),
              };
              rowData.push(cellData);
            }

            tableData.rows.push(rowData);
          }

          children.push(tableData);
        } else if (tag === "w:sectPr") {
          const sectionData = {
            type: "section",
            styles: {
              pageSize:
                $doc(child).find("w\\:pgSz").attr("w:w") +
                "x" +
                $doc(child).find("w\\:pgSz").attr("w:h"),
              margins: {
                top: $doc(child).find("w\\:pgMar").attr("w:top"),
                bottom: $doc(child).find("w\\:pgMar").attr("w:bottom"),
                left: $doc(child).find("w\\:pgMar").attr("w:left"),
                right: $doc(child).find("w\\:pgMar").attr("w:right"),
              },
            },
          };
          children.push(sectionData);
        }
      }

      return children;
    }

    function buildNumberingMap($numbering) {
      const abstractNumMap = {}; // Stores the abstract numbering definitions
      const numberingMap = {}; // Stores the final mapping of numId to its levels

      // Recursive function to resolve abstract numbering, including any references to other abstractNumIds
      function resolveAbstractNum(abstractNumId) {
        // If this abstractNumId is already resolved, return it
        if (abstractNumMap[abstractNumId]?.resolved) {
          return abstractNumMap[abstractNumId].levels;
        }

        const levels = abstractNumMap[abstractNumId]?.levels || [];

        // Check if this abstractNumId refers to another abstractNumId
        const referencedAbstractNumId =
          abstractNumMap[abstractNumId]?.referencedAbstractNumId;
        if (referencedAbstractNumId) {
          // Recursively resolve the referenced abstractNumId and merge the levels
          const referencedLevels = resolveAbstractNum(referencedAbstractNumId);
          Object.assign(levels, referencedLevels);
        }

        // Mark this abstractNumId as resolved to avoid circular references
        abstractNumMap[abstractNumId] = {
          levels,
          resolved: true,
        };

        return levels;
      }

      // Loop through each abstract numbering definition and build the abstractNumMap
      $numbering("w\\:abstractNum").each((_, abstractNum) => {
        const abstractNumId = $numbering(abstractNum).attr("w:abstractNumId");
        const levels = [];

        $numbering(abstractNum)
          .find("w\\:lvl")
          .each((_, level) => {
            const levelIndex = $numbering(level).attr("w:ilvl");
            const numFmt = $numbering(level).find("w\\:numFmt").attr("w:val");
            const lvlText = $numbering(level).find("w\\:lvlText").attr("w:val");

            // Get indentation spacing
            const indent = {};
            const indElement = $numbering(level).find("w\\:ind");
            if (indElement.length > 0) {
              indent.left = indElement.attr("w:left") || null;
              indent.hanging = indElement.attr("w:hanging") || null;
              indent.firstLine = indElement.attr("w:firstLine") || null;
            }

            levels[levelIndex] = {
              numFmt,
              lvlText,
              indent,
            };
          });

        // Check if this abstractNum refers to another abstractNum
        const referencedAbstractNumId = $numbering(abstractNum)
          .find("w\\:nsid")
          .attr("w:val");
        abstractNumMap[abstractNumId] = {
          levels,
          referencedAbstractNumId: referencedAbstractNumId || null,
          resolved: false, // Initially set as unresolved
        };
      });

      // Loop through each num element to build the numberingMap
      $numbering("w\\:num").each((_, num) => {
        const numId = $numbering(num).attr("w:numId");
        const abstractNumId = $numbering(num)
          .find("w\\:abstractNumId")
          .attr("w:val");

        if (abstractNumId) {
          numberingMap[numId] = resolveAbstractNum(abstractNumId);
        }
      });

      return numberingMap;
    }

    async function parseDrawing(drawingElement) {
      const anchor = $doc(drawingElement).find("wp\\:anchor");
      const graphicData = $doc(drawingElement).find("a\\:graphicData");

      const imageData = {
        type: "image",
        src: null,
        position: {
          horizontal: null,
          vertical: null,
        },
        size: {
          width: null,
          height: null,
        },
        properties: {},
      };

      const positionH = anchor.find("wp\\:positionH wp\\:align").text();
      const positionV = anchor.find("wp\\:positionV wp\\:align").text();
      const cx = anchor.find("wp\\:extent").attr("cx");
      const cy = anchor.find("wp\\:extent").attr("cy");

      if (positionH && positionV) {
        imageData.position.horizontal = positionH;
        imageData.position.vertical = positionV;
      }

      if (cx && cy) {
        imageData.size.width = parseInt(cx, 10);
        imageData.size.height = parseInt(cy, 10);
      }

      const blip = graphicData.find("a\\:blip");
      const embedId = blip.attr("r:embed");

      if (embedId) {
        const relsPath = path.join(
          outputDir,
          "word",
          "_rels",
          "document.xml.rels"
        );
        const relsXml = fs.readFileSync(relsPath, "utf8");
        const $rels = cheerio.load(relsXml, { xmlMode: true });

        const target = $rels(`Relationship[Id="${embedId}"]`).attr("Target");
        if (target) {
          const imagePath = path.join(outputDir, "word", target);

          try {
            const uploadResult = await uploadToCloudinary(imagePath);
            imageData.src = uploadResult.secure_url; // Cloudinary secure URL of the uploaded image
          } catch (error) {
            console.error("Error uploading image to Cloudinary:", error);
          }
        }
      }
      console.log(imageData);
      return imageData;
    }

    function extractListInfo(numId, ilvl, numberingMap) {
      const levelInfo = numberingMap[numId] && numberingMap[numId][ilvl];

      return levelInfo
        ? {
            isBullet: levelInfo.numFmt === "bullet",
            bulletText: levelInfo.lvlText,
            level: ilvl,
            indent: levelInfo.indent,
          }
        : null;
    }

    const parsedContent = await parseElement($doc("w\\:document > w\\:body"));
    rimraf.sync(outputDir);

    return mapSections(parsedContent);
  } catch (err) {
    console.error("Error extracting DOCX content:", err);
    throw err;
  }
}

function mapSections(paragraphs) {
  const sections = [];
  let currentSection = null;
  let currentSubsection = null;
  let preamble = { body: [] };

  paragraphs.forEach((paragraph) => {
    const { styleName, text } = paragraph;
    if (paragraph.type === "image") {
      console.log("image", paragraph);
    }

    if (["Heading1", "TGTHEADING1"].includes(styleName)) {
      if (currentSubsection) {
        currentSection.body.push(currentSubsection);
      }

      // If there's an active section, push it to the sections array
      if (currentSection) {
        sections.push(currentSection);
      }

      // Start a new main section
      currentSection = {
        title: text,
        body: [],
        // subsections: [],
      };

      // Reset current subsection
      currentSubsection = null;
    } else if (styleName === "TGTHEADING2" && currentSection) {
      // If there's an active subsection, push it to the subsections array
      if (currentSubsection) {
        currentSection.body.push(currentSubsection);
      }

      // Start a new subsection within the current section
      currentSubsection = {
        title: text,
        body: [],
      };
    } else if (currentSubsection) {
      // If there's an active subsection, add body to it
      currentSubsection.body.push(paragraph);
    } else if (currentSection) {
      // If there's no active subsection, add body to the current section
      currentSection.body.push(paragraph);
    } else {
      // If no section has started, add body to the preamble
      preamble.body.push(paragraph);
    }
  });

  // Push the last subsection and section if they exist
  if (currentSubsection) {
    currentSection.body.push(currentSubsection);
  }

  if (currentSection) {
    sections.push(currentSection);
  }

  // Include preamble if it has any body
  if (preamble.body.length > 0) {
    sections.unshift({
      title: "Cover Page",
      body: preamble.body,
    });
  }

  return sections;
}

module.exports = { extractDocxContent };
